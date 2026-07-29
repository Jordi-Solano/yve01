"""El asignador pega el estado del cruce a la factura CORRECTA.

EL BUG QUE ESTO PROTEGE, y por que era el mas grave abierto:
`asignador_cuentas` juntaba los informes de cruce y los deduplicaba por
`archivo` A SECAS. Con dos hoteles que suban un fichero con el mismo nombre
sobrevivia UN estado y se le pegaba a las dos facturas, asi que una
**DIFERENCIA_IMPORTE real llegaba al panel en verde**: un cobro inflado listo
para aprobar y para que Oracle lo contabilizara.

No fue un error de razonamiento. Cuando se escribio, `archivo` SI era unico por
factura. La separacion por hotel cambio la clave del escritor a
`(archivo, hotel_id)` y este lado se quedo atras.

QUE SE PROTEGE:
  1. la clave incluye el hotel: dos facturas con el mismo archivo en hoteles
     distintos reciben CADA UNA su estado
  2. nadie PIERDE el estado que ya tenia — incluida la factura cuyo informe no
     trae columna de hotel (la red por `archivo` a secas)
  3. la red NO se aplica cuando el archivo es ambiguo: ahi esta el bug
  4. la prioridad PO > F&B > albaran se conserva
  5. `oracle_status` sobrevive al regenerado (candado de la Fase 0)
  6. la clave temporal no se cuela como columna del informe

`--sabotaje` devuelve la clave a `archivo` a secas y comprueba que los asserts
GRITAN. Escribe una copia saboteada EN DISCO porque la prueba lanza el
asignador como SUBPROCESO y un parche en memoria no cruza esa frontera.
"""
import json
import os
import shutil
import subprocess
import sys
from datetime import date

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

import pandas as pd  # noqa: E402

HOY = date.today().strftime("%Y%m%d")
TENANT = "test-asig-clave"
ARBOL = os.path.join(BASE, "tenants", TENANT)
H_A, H_B = "HTESTAA", "HTESTBB"
NF = "NO_ENCONTRADO"
SABOTAJE = "--sabotaje" in sys.argv

MODULO = "asignador_cuentas.py"
_SABOTEADO = os.path.join(BASE, "asignador_cuentas_SABOTAJE.py")


def _fac(archivo, numero, prov, base, hotel, oracle=""):
    cuota = round(base * 0.21, 2)
    return {"archivo": archivo, "numero_factura": numero, "fecha": "22/07/2026",
            "nombre_proveedor": prov, "base_imponible": base, "porcentaje_iva": 21.0,
            "cuota_iva": cuota, "total_factura": round(base + cuota, 2),
            "descripcion_concepto": "", "hotel_id": hotel, "oracle_status": oracle}


FACTURAS = [
    # una YA CONTABILIZADA: su marcador no puede perderse al regenerar
    _fac("fa1.pdf", "F-A1", "Pescados Rias SL", 1000.0, H_A, "CONTABILIZADA"),
    _fac("fa2.pdf", "F-A2", "Congelados Ebre SL", 800.0, H_A),
    _fac("fx1.pdf", "F-X1", "Coca-Cola Iberia SL", 300.0, ""),      # sin hotel
    _fac("fa3.pdf", "F-A3", "Textil Nord SL", 450.0, H_A),          # sin cruce
    _fac("hoja.csv#s-n-1", "", "Aves Gali SL", 620.0, H_A),         # sin numero
    _fac("fa5.pdf", "F-A5", "Makro Espana SA", 900.0, H_A),         # en dos informes
    _fac("fa6.pdf", "F-A6", "Lacteos Segre SA", 700.0, H_A),        # la red
    # EL BUG: mismo archivo, hoteles distintos, estados OPUESTOS
    _fac("factura_julio.pdf", "F-A9", "Carnes Vic SL", 1000.0, H_A),
    _fac("factura_julio.pdf", "F-B9", "Aves Gali SL", 1200.0, H_B),
]


def _alb(archivo, numero, hotel, estado, detalle):
    # los vacios como la cadena NO_ENCONTRADO, que es lo que escribe de verdad
    # `matching_ap_albaran`: es la trampa de normalizacion que la clave debe
    # tratar igual que un NaN
    return {"archivo": archivo, "numero_factura": numero or NF,
            "nombre_proveedor": NF, "hotel_id": hotel,
            "estado_matching": estado, "detalle_matching": detalle}


INFORME_ALBARAN = [
    _alb("fa1.pdf", "F-A1", H_A, "MATCH_ALBARAN_OK", "1 albaran cuadra"),
    _alb("fa2.pdf", "F-A2", H_A, "DIFERENCIA_IMPORTE", "cobra 450 EUR mas"),
    _alb("fx1.pdf", "F-X1", "", "MATCH_ALBARAN_OK", "1 albaran cuadra"),
    _alb("hoja.csv#s-n-1", "", H_A, "MATCH_ALBARAN_OK", "1 albaran cuadra"),
    _alb("fa5.pdf", "F-A5", H_A, "MATCH_ALBARAN_OK", "1 albaran cuadra"),
    _alb("factura_julio.pdf", "F-A9", H_A, "MATCH_ALBARAN_OK", "1 albaran cuadra"),
    _alb("factura_julio.pdf", "F-B9", H_B, "DIFERENCIA_IMPORTE", "cobra 600 EUR mas"),
]
# En SU PROPIO fichero para que la columna hotel_id NO EXISTA de verdad: metido
# en la misma hoja, pandas le pondria NaN y la columna si estaria.
INFORME_SIN_HOTEL = [{"archivo": "fa6.pdf", "numero_factura": "F-A6",
                      "estado_matching": "MATCH_ALBARAN_OK",
                      "detalle_matching": "1 albaran cuadra"}]
# Va PRIMERO en la lista del asignador: tiene que ganar sobre el de albaranes.
INFORME_OTRAS = [{"archivo": "fa5.pdf", "numero_factura": "F-A5", "hotel_id": H_A,
                  "estado_matching": "MATCH_CORRECTO", "detalle": "PO PO-77 cuadra"}]


def _montar():
    if os.path.isdir(ARBOL):
        shutil.rmtree(ARBOL)
    proc = os.path.join(ARBOL, "facturas-procesadas")
    rep = os.path.join(ARBOL, "reportes")
    datos = os.path.join(ARBOL, "datos-referencia")
    for d in (proc, rep, datos, os.path.join(ARBOL, "aprobaciones")):
        os.makedirs(d, exist_ok=True)
    for f in ("proveedores.xlsx", "plan_cuentas.xlsx"):
        src = os.path.join(BASE, "datos-referencia", f)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(datos, f))
    json.dump([{"id": H_A, "nombre": "Hotel Test A", "activo": True},
               {"id": H_B, "nombre": "Hotel Test B", "activo": True}],
              open(os.path.join(datos, "hoteles.json"), "w"))
    pd.DataFrame(FACTURAS).to_excel(os.path.join(proc, f"facturas_ap_{HOY}.xlsx"),
                                    sheet_name="Facturas", index=False)
    with pd.ExcelWriter(os.path.join(rep, f"matching_albaran_{HOY}.xlsx")) as w:
        pd.DataFrame(INFORME_ALBARAN).to_excel(w, sheet_name="Facturas", index=False)
    pd.DataFrame(INFORME_OTRAS).to_excel(
        os.path.join(rep, f"matching_otras_{HOY}.xlsx"), index=False)
    pd.DataFrame(INFORME_SIN_HOTEL).to_excel(
        os.path.join(rep, f"matching_fb_{HOY}.xlsx"), index=False)


def _correr():
    """Ejecuta el asignador y devuelve (estados, df, salida)."""
    _montar()
    env = dict(os.environ, YVE_TENANT=TENANT, YVE_HOTEL="")
    r = subprocess.run([sys.executable, MODULO], cwd=BASE, capture_output=True,
                       text=True, timeout=300, env=env)
    assert r.returncode == 0, r.stderr[-1200:]
    proc = os.path.join(ARBOL, "facturas-procesadas")
    ruta = sorted(f for f in os.listdir(proc) if f.startswith("facturas_contabilizadas_"))
    assert ruta, "el asignador no ha generado informe"
    df = pd.read_excel(os.path.join(proc, ruta[-1]))
    estados = {}
    for fila in df.to_dict("records"):
        num = str(fila.get("numero_factura") or "").strip()
        clave = num if num and num.lower() != "nan" else str(fila.get("archivo") or "")
        est = str(fila.get("estado_matching") or "").strip()
        estados[clave] = est if est and est.lower() != "nan" else "SIN_ESTADO"
    return estados, df, r.stdout


_CACHE = {}


def estados():
    if "v" not in _CACHE:
        _CACHE["v"] = _correr()
    return _CACHE["v"]


# ── 1 y 3 · la clave lleva el hotel ───────────────────────────────────────

def test_cada_hotel_recibe_su_estado():
    e, _df, _o = estados()
    assert e["F-A9"] == "MATCH_ALBARAN_OK", e["F-A9"]
    assert e["F-B9"] == "DIFERENCIA_IMPORTE", (
        "F-B9 tiene una diferencia de importe REAL en su informe y llega al panel "
        f"como {e['F-B9']}. Un cobro inflado en verde, listo para aprobar y para "
        "que Oracle lo contabilice. La clave del cruce ha vuelto a ser `archivo` "
        "a secas y se esta pegando el estado del otro hotel.")
    print("  ✔ dos facturas con el mismo archivo en hoteles distintos: cada una su estado")


# ── 2 · nadie pierde estado ───────────────────────────────────────────────

def test_nadie_pierde_su_estado():
    e, _df, _o = estados()
    esperado = {"F-A1": "MATCH_ALBARAN_OK", "F-A2": "DIFERENCIA_IMPORTE",
                "F-X1": "MATCH_ALBARAN_OK", "hoja.csv#s-n-1": "MATCH_ALBARAN_OK"}
    for k, v in esperado.items():
        assert e.get(k) == v, (
            f"{k} tenia estado y ahora sale como {e.get(k)}. Este cambio solo "
            "puede conservar o dar estado, nunca quitarlo.")
    assert e["F-A3"] == "SIN_ESTADO", (
        f"F-A3 no sale en ningun informe y le ha aparecido un estado: {e['F-A3']}")
    print("  ✔ ninguna factura pierde el estado que tenia (ni gana uno inventado)")


def test_la_red_rescata_el_informe_sin_hotel():
    """Un informe SIN columna de hotel (de antes de la separacion) frente a una
    factura CON hotel: la clave fuerte no casa, y la red por `archivo` a secas
    tiene que salvarlo. Es el unico caso en que la clave fuerte falla."""
    e, _df, salida = estados()
    assert e["F-A6"] == "MATCH_ALBARAN_OK", (
        f"F-A6 ha perdido su cruce ({e['F-A6']}): su informe no trae columna de "
        "hotel y la red no lo ha rescatado.")
    assert "nombre de fichero" in salida, (
        "la red no ha entrado en accion, asi que este caso no esta probado de "
        "verdad — F-A6 tiene que cruzar POR LA RED, no por la clave fuerte")
    print("  ✔ la red rescata el informe sin columna de hotel, y se usa de verdad")


# ── 4 · la prioridad ──────────────────────────────────────────────────────

def test_la_prioridad_se_conserva():
    e, _df, _o = estados()
    assert e["F-A5"] == "MATCH_CORRECTO", (
        f"F-A5 sale en el informe de PO y en el de albaranes, y ha ganado "
        f"{e['F-A5']}. La lista de informes esta ordenada a proposito "
        "(PO > F&B > albaran) y `keep=\"first\"` la sostiene.")
    print("  ✔ con dos informes gana el primero de la lista (PO > F&B > albaran)")


# ── 5 y 6 · el informe que se escribe ─────────────────────────────────────

def test_oracle_status_sobrevive():
    _e, df, _o = estados()
    assert "oracle_status" in df.columns, "se ha perdido la columna oracle_status"
    fila = df[df["numero_factura"].astype(str) == "F-A1"]
    assert not fila.empty and str(fila.iloc[0]["oracle_status"]).upper() == "CONTABILIZADA", (
        "el marcador de Oracle de una factura ya contabilizada no ha sobrevivido "
        "al regenerado. Es el UNICO candado que impide contabilizar dos veces.")
    print("  ✔ oracle_status sobrevive al regenerado del informe")


def test_la_clave_no_se_cuela_en_el_informe():
    _e, df, _o = estados()
    colados = [c for c in df.columns if str(c).startswith("_")]
    assert not colados, (
        f"columnas internas en facturas_contabilizadas: {colados}. Ese fichero "
        "lo leen el panel, las aprobaciones y Oracle.")
    print("  ✔ la clave temporal no se cuela como columna del informe")


PRUEBAS = [test_cada_hotel_recibe_su_estado, test_nadie_pierde_su_estado,
           test_la_red_rescata_el_informe_sin_hotel, test_la_prioridad_se_conserva,
           test_oracle_status_sobrevive, test_la_clave_no_se_cuela_en_el_informe]

# Con la clave rota, estas TIENEN que caer. Las demas siguen pasando: el
# sabotaje solo quita el hotel de la clave, no rompe el modulo entero.
DEBEN_CAER = {"test_cada_hotel_recibe_su_estado"}


def _sabotear():
    """Copia del modulo con la clave devuelta a `archivo` a secas."""
    src = open(os.path.join(BASE, MODULO), encoding="utf-8").read()
    viejo = '''    return _txt_cruce(fila.get("archivo")) + "|" + _txt_cruce(fila.get("hotel_id"))'''
    assert src.count(viejo) == 1, (
        "el sabotaje ya no encuentra la linea de la clave: el test ha dejado de "
        "saber que romper y hay que ponerlo al dia")
    src = src.replace(viejo, '    return _txt_cruce(fila.get("archivo"))', 1)
    open(_SABOTEADO, "w", encoding="utf-8").write(src)
    return os.path.basename(_SABOTEADO)


def main():
    global MODULO
    if SABOTAJE:
        MODULO = _sabotear()
    print(f"\n{'SABOTAJE — ' if SABOTAJE else ''}asignador: el estado va a la factura correcta")
    print("=" * 64)
    fallos, caidas = [], set()
    try:
        for p in PRUEBAS:
            try:
                p()
            except AssertionError as e:
                caidas.add(p.__name__)
                if not SABOTAJE:
                    fallos.append(p.__name__)
                    print(f"  ✗ {p.__name__}\n      {e}")
    finally:
        if os.path.isdir(ARBOL):
            shutil.rmtree(ARBOL)
        if os.path.exists(_SABOTEADO):
            os.remove(_SABOTEADO)

    print("=" * 64)
    if SABOTAJE:
        no_gritaron = DEBEN_CAER - caidas
        if no_gritaron:
            print("  ✗ con la clave rota NO han fallado: " + ", ".join(sorted(no_gritaron)))
            print("    Un test que no puede fallar no protege de nada.")
            return 1
        print(f"  ✔ con la clave rota falla la que tiene que fallar")
        return 0
    if fallos:
        print(f"  {len(fallos)} FALLO(S)")
        return 1
    print(f"  {len(PRUEBAS)} pruebas OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
