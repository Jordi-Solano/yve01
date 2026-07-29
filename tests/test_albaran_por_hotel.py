"""El cruce factura <-> albaran respeta el hotel.

Antes de esto, con dos hoteles en el mismo tenant, la factura de uno cuadraba
con el albaran del otro y salia como MATCH_ALBARAN_OK: un numero equivocado
presentado como correcto.

QUE SE PROTEGE, y por que cada cosa:

  1. dos hoteles distintos NO cruzan                     (el bug)
  2. igualdad ESTRICTA: el vacio solo con el vacio       (decision del usuario)
  3. con 0 hoteles y con 1 hotel sale EXACTAMENTE lo mismo que sin la regla
     (son el tenant recien creado y el hotel suelto, o sea la mayoria: si esto
      se rompe, se rompe para casi todos los clientes y no para el caso raro)
  4. el hueco se EXPLICA: una factura que pierde su albaran por el hotel lo
     dice. Sin esto el arreglo cambia un falso positivo por un falso negativo
     mudo, y alguien reclama al proveedor una entrega que si llego
  5. las lineas no se mezclan entre hoteles que suban un fichero con el MISMO
     nombre (el nivel 3 busca por `archivo`)
  6. `hotel_id` viaja hasta la hoja Facturas del informe
  7. el resultado NO depende de la sesion: el modulo no lee YVE_HOTEL ni
     pregunta cual es el hotel activo. Si algun dia alguien filtra la carga, el
     informe —que se reescribe entero— perderia las filas del otro hotel.
     Esto se comprueba con el AST, no con grep: el propio modulo EXPLICA en su
     docstring que no lee YVE_HOTEL, asi que un grep se encontraria a si mismo.

Se puede ejecutar con `--sabotaje`: revienta a proposito la regla del hotel y
comprueba que los asserts GRITAN. Un test que no puede fallar no protege nada.
"""
import ast
import json
import os
import shutil
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

import pandas as pd  # noqa: E402

TENANT = "test-alb-hotel"
ARBOL = os.path.join(BASE, "tenants", TENANT)
H_A, H_B = "HTESTAAA", "HTESTBBB"
H_C = "HTESTCCC"   # entra nuevo: NO tiene ni un albaran
SIN = ""

SABOTAJE = "--sabotaje" in sys.argv


# ── fixtures ──────────────────────────────────────────────────────────────

def _fac(archivo, numero, prov, base, hotel, fecha="20/07/2026"):
    return {"archivo": archivo, "numero_factura": numero, "fecha": fecha,
            "nombre_proveedor": prov, "base_imponible": base,
            "porcentaje_iva": 21.0, "cuota_iva": round(base * 0.21, 2),
            "total_factura": round(base * 1.21, 2),
            "descripcion_concepto": "", "hotel_id": hotel}


def _alb(numero, prov, total, hotel, entrega="15/07/2026"):
    return {"clave": f"{numero}|{prov}", "archivo": f"{numero}.pdf",
            "numero_albaran": numero, "nombre_proveedor": prov,
            "fecha_entrega": entrega, "referencia_pedido": "",
            "referencia_factura": "", "total_albaran": total, "hotel_id": hotel}


def _lin(archivo, numero, prov, hotel, desc, cant, prec):
    return {"archivo": archivo, "numero_factura": numero,
            "nombre_proveedor": prov, "n_linea": 1, "descripcion": desc,
            "cantidad": cant, "unidad": "ud", "precio_unitario": prec,
            "importe": round(cant * prec, 2), "hotel_id": hotel}


def _lin_a(numero, prov, hotel, desc, cant, prec):
    return {"clave": f"{numero}|{prov}", "numero_albaran": numero,
            "nombre_proveedor": prov, "n_linea": 1, "descripcion": desc,
            "cantidad": cant, "unidad": "ud", "precio_unitario": prec,
            "importe": round(cant * prec, 2), "hotel_id": hotel}


# Mismo proveedor en los dos hoteles: es el caso que hacia cruzar mal.
# Y `factura_julio.pdf` repetido a proposito, con proveedores DISTINTOS para que
# no compitan por albaranes: asi lo unico que se mezclaria son las lineas.
FACTURAS = [
    _fac("fa1.pdf", "F-A1", "Pescados Rias SL", 1000.0, H_A),
    _fac("fa2.pdf", "F-A2", "Congelados Ebre SL", 800.0, H_A),
    _fac("fx1.pdf", "F-X1", "Coca-Cola Iberia SL", 300.0, SIN),
    _fac("fx2.pdf", "F-X2", "Verduras del Prat SL", 600.0, SIN),
    _fac("factura_julio.pdf", "F-A4", "Carnes Vic SL", 1000.0, H_A),
    _fac("factura_julio.pdf", "F-B4", "Aves Gali SL", 1200.0, H_B),
    # EL CORTE. Sin albaran por ningun lado; lo que se prueba es como se llaman.
    # Hotel A registra desde 15/07 pero el corte GLOBAL es 10/01 (de ALB-B0):
    _fac("fa5.pdf", "F-A5", "Aceites Baix SL", 400.0, H_A, "15/03/2026"),
    # Hotel C no ha registrado NI UNO: es el hotel que entra nuevo
    _fac("fc5.pdf", "F-C5", "Limpiezas Ter SL", 350.0, H_C, "15/03/2026"),
]
ALBARANES = [
    _alb("ALB-A1", "Pescados Rias SL", 1000.0, H_A),
    _alb("ALB-B2", "Congelados Ebre SL", 800.0, H_B),      # <- otro hotel
    _alb("ALB-X1", "Coca-Cola Iberia SL", 300.0, SIN),
    _alb("ALB-A3", "Verduras del Prat SL", 600.0, H_A),    # <- tiene hotel
    _alb("ALB-A4", "Carnes Vic SL", 1000.0, H_A),
    _alb("ALB-B4", "Aves Gali SL", 1200.0, H_B),
    # el mas antiguo de todos, y de OTRO hotel: es el que marcaba el corte global
    _alb("ALB-B0", "Frutas del Sur SL", 200.0, H_B, "10/01/2026"),
]
LIN_F = [_lin("factura_julio.pdf", "F-A4", "Carnes Vic SL", H_A, "Solomillo", 10, 100.0),
         _lin("factura_julio.pdf", "F-B4", "Aves Gali SL", H_B, "Solomillo", 10, 120.0)]
LIN_A = [_lin_a("ALB-A4", "Carnes Vic SL", H_A, "Solomillo", 10, 100.0),
         _lin_a("ALB-B4", "Aves Gali SL", H_B, "Solomillo", 10, 120.0)]


def _montar(facturas=None, albaranes=None, lin_f=None, lin_a=None, censo=True):
    """Deja el arbol del tenant listo. Se limpia entero cada vez (regla: una
    bateria que se deja ficheros de la prueba anterior falsea la siguiente)."""
    if os.path.isdir(ARBOL):
        shutil.rmtree(ARBOL)
    for sub in ("datos-referencia", "reportes", "facturas-procesadas",
                "facturas-entrada", "aprobaciones"):
        os.makedirs(os.path.join(ARBOL, sub), exist_ok=True)
    json.dump([{"id": H_A, "nombre": "Hotel Test A", "activo": True},
               {"id": H_B, "nombre": "Hotel Test B", "activo": True},
               {"id": H_C, "nombre": "Hotel Test C", "activo": True}] if censo else [],
              open(os.path.join(ARBOL, "datos-referencia", "hoteles.json"), "w"))
    proc = os.path.join(ARBOL, "facturas-procesadas")
    with pd.ExcelWriter(os.path.join(proc, "facturas_ap_20260720.xlsx")) as w:
        pd.DataFrame(facturas if facturas is not None else FACTURAS
                     ).to_excel(w, sheet_name="Facturas", index=False)
        pd.DataFrame(lin_f if lin_f is not None else LIN_F
                     ).to_excel(w, sheet_name="Lineas", index=False)
    with pd.ExcelWriter(os.path.join(proc, "albaranes_20260720.xlsx")) as w:
        pd.DataFrame(albaranes if albaranes is not None else ALBARANES
                     ).to_excel(w, sheet_name="Albaranes", index=False)
        pd.DataFrame(lin_a if lin_a is not None else LIN_A
                     ).to_excel(w, sheet_name="Lineas", index=False)


def _limpiar():
    if os.path.isdir(ARBOL):
        shutil.rmtree(ARBOL)


# ── el modulo, con el sabotaje opcional ───────────────────────────────────

import matching_ap_albaran as M  # noqa: E402

# El modulo que ejecuta la prueba de punta a punta. En sabotaje NO puede ser el
# mismo: `test_informe_completo` lo lanza como SUBPROCESO, y un parche en
# memoria no cruza esa frontera. La primera version de este test parcheaba solo
# el objeto importado y por eso la prueba del informe seguia pasando con la
# regla rota — o sea, no protegia nada. Se escribe una copia saboteada EN DISCO.
MODULO = "matching_ap_albaran.py"
_SABOTEADO = os.path.join(BASE, "matching_ap_albaran_SABOTAJE.py")

if SABOTAJE:
    # Revienta EXACTAMENTE lo que el test protege: la regla del hotel deja de
    # mirar el hotel y las lineas dejan de acotarse. Es el codigo de antes.
    M._mismo_hotel = lambda fila_f, alb: True
    _de = M._lineas_de
    M._lineas_de = lambda df, col, vals, hotel=None: _de(df, col, vals, None)

    _src = open(os.path.join(BASE, MODULO), encoding="utf-8").read()
    for _viejo, _nuevo in (
            ("    return _hotel(fila_f) == _hotel(alb)", "    return True"),
            ("            and (hotel is None or _hotel(r) == hotel)]", "            ]"),
            # y el corte vuelve a ser global, sin la excepcion del hotel nuevo:
            ("        corte = (cortes or {}).get(hot)",
             "        corte = min((cortes or {}).values(), default=None)"),
            ("        if con_albaran is not None and hot not in con_albaran:",
             "        if False:")):
        assert _src.count(_viejo) == 1, (
            f"el sabotaje ya no encuentra {_viejo!r} en {MODULO}: el test ha "
            "dejado de saber que romper y hay que ponerlo al dia")
        _src = _src.replace(_viejo, _nuevo, 1)
    open(_SABOTEADO, "w", encoding="utf-8").write(_src)
    MODULO = os.path.basename(_SABOTEADO)


def _emparejar(facturas, albaranes):
    df_f = pd.DataFrame(facturas)
    df_a = pd.DataFrame(albaranes)
    empare, _porque, bloq = M.emparejar(df_f, df_a)
    # {numero_factura: [numeros de albaran]}
    return ({str(df_f.loc[i]["numero_factura"]):
             sorted(str(df_a.loc[j]["numero_albaran"]) for j in idxs)
             for i, idxs in empare.items()}, df_f, df_a, bloq)


# ── 1 y 2 · la regla ──────────────────────────────────────────────────────

def test_dos_hoteles_no_cruzan():
    cruces, _f, _a, _b = _emparejar(FACTURAS, ALBARANES)
    assert cruces["F-A1"] == ["ALB-A1"], f"el cruce bueno del hotel A: {cruces['F-A1']}"
    assert cruces["F-A2"] == [], (
        "la factura del hotel A se ha llevado un albaran del hotel B: "
        f"{cruces['F-A2']}")
    print("  ✔ una factura no cruza con el albaran de otro hotel")


def test_el_vacio_no_es_comodin():
    cruces, _f, _a, _b = _emparejar(FACTURAS, ALBARANES)
    assert cruces["F-X1"] == ["ALB-X1"], (
        f"sin hotel con sin hotel tiene que seguir cruzando: {cruces['F-X1']}")
    assert cruces["F-X2"] == [], (
        "una factura SIN hotel se ha llevado un albaran CON hotel: "
        f"{cruces['F-X2']} (la regla acordada es igualdad estricta)")
    print("  ✔ lo que no lleva hotel solo cruza con lo que no lleva hotel")


# ── 3 · los dos casos que NO se pueden romper ─────────────────────────────

def test_cero_hoteles_igual_que_siempre():
    """Tenant sin censo: TODO esta sin asignar, asi que todo cruza con todo."""
    fac = [dict(f, hotel_id="") for f in FACTURAS]
    alb = [dict(a, hotel_id="") for a in ALBARANES]
    cruces, _f, _a, _b = _emparejar(fac, alb)
    for num, esperado in (("F-A1", ["ALB-A1"]), ("F-A2", ["ALB-B2"]),
                          ("F-X1", ["ALB-X1"]), ("F-X2", ["ALB-A3"])):
        assert cruces[num] == esperado, (
            f"con 0 hoteles {num} tiene que cruzar como siempre: "
            f"{cruces[num]} en vez de {esperado}")
    print("  ✔ con 0 hoteles el resultado es el de siempre")


def test_un_hotel_igual_que_siempre():
    fac = [dict(f, hotel_id=H_A) for f in FACTURAS]
    alb = [dict(a, hotel_id=H_A) for a in ALBARANES]
    cruces, _f, _a, _b = _emparejar(fac, alb)
    assert cruces["F-A2"] == ["ALB-B2"], (
        f"con 1 hotel todo lleva la misma etiqueta y tiene que cruzar: {cruces['F-A2']}")
    print("  ✔ con 1 hotel el resultado es el de siempre")


def test_la_columna_puede_no_existir():
    """Ficheros de antes de la separacion: ni siquiera hay columna `hotel_id`.

    Tienen que cruzar entre ellos, no desaparecer. `_hotel` los ve como '' y esa
    es la caja 'sin asignar'.
    """
    fac = [{k: v for k, v in f.items() if k != "hotel_id"} for f in FACTURAS]
    alb = [{k: v for k, v in a.items() if k != "hotel_id"} for a in ALBARANES]
    cruces, _f, _a, _b = _emparejar(fac, alb)
    assert cruces["F-A2"] == ["ALB-B2"], (
        f"sin columna de hotel todo es 'sin asignar' y cruza: {cruces['F-A2']}")
    print("  ✔ sin columna `hotel_id` (ficheros viejos) siguen cruzando entre ellos")


def test_el_nan_es_sin_asignar():
    """Excel devuelve NaN donde se guardo '', y `str(nan)` es 'nan', que NO es
    vacio para Python. Sin `_txt` dos documentos igual de huerfanos no cruzarian.
    Paso de verdad: el informe imprimia `hotel=nan`."""
    fac = [dict(f, hotel_id=float("nan")) if f["hotel_id"] == SIN else f
           for f in FACTURAS]
    alb = [dict(a, hotel_id=float("nan")) if a["hotel_id"] == SIN else a
           for a in ALBARANES]
    cruces, _f, _a, _b = _emparejar(fac, alb)
    assert cruces["F-X1"] == ["ALB-X1"], (
        f"NaN y '' son el mismo 'sin asignar': {cruces['F-X1']}")
    print("  ✔ NaN y cadena vacia son el mismo 'sin asignar'")


# ── 5 · las lineas no se mezclan ──────────────────────────────────────────

def test_lineas_no_se_mezclan_entre_hoteles():
    df = pd.DataFrame(LIN_F)
    del_a = M._lineas_de(df, "archivo", ["factura_julio.pdf"], H_A)
    assert len(del_a) == 1, (
        f"dos hoteles con el mismo nombre de fichero: se han cogido {len(del_a)} "
        "lineas en vez de 1 — las del otro hotel se estan colando")
    assert M._txt(del_a[0].get("numero_factura")) == "F-A4"
    print("  ✔ las lineas de un `factura_julio.pdf` no se mezclan con las del otro")


# ── 7 · no depende de la sesion (AST, no grep) ────────────────────────────

def test_no_pregunta_por_el_hotel_activo():
    arbol = ast.parse(open(os.path.join(BASE, "matching_ap_albaran.py"),
                           encoding="utf-8").read())
    malos = []
    for nodo in ast.walk(arbol):
        # os.environ.get("YVE_HOTEL") / os.environ["YVE_HOTEL"]
        if isinstance(nodo, ast.Constant) and nodo.value == "YVE_HOTEL":
            malos.append("YVE_HOTEL")
        # censo_hoteles.activo() / .para_guardar() -> serian la sesion decidiendo
        if isinstance(nodo, ast.Attribute) and nodo.attr in ("activo", "para_guardar"):
            malos.append(f".{nodo.attr}()")
        # almacen_datos.solo_del_hotel_activo(...)
        if isinstance(nodo, ast.Name) and nodo.id == "solo_del_hotel_activo":
            malos.append("solo_del_hotel_activo")
    assert not malos, (
        "el cruce ha empezado a depender de la sesion (" + ", ".join(sorted(set(malos)))
        + "). El informe se reescribe ENTERO en cada pasada: filtrando la carga, "
        "la pasada de un hotel borraria del informe de hoy las filas del otro, y "
        "esa hoja es la etapa que gana en almacen_datos.albaranes().")
    print("  ✔ el modulo no consulta el hotel de la sesion (AST, no grep)")


# ── 4 y 6 · el informe de verdad, de punta a punta ────────────────────────

def test_informe_completo():
    _montar()
    env = dict(os.environ, YVE_TENANT=TENANT, YVE_HOTEL="")
    r = subprocess.run([sys.executable, MODULO], cwd=BASE,
                       capture_output=True, text=True, timeout=300, env=env)
    assert r.returncode == 0, r.stderr[-1200:]
    reps = [f for f in os.listdir(os.path.join(ARBOL, "reportes"))
            if f.startswith("matching_albaran_")]
    assert reps, "no se ha generado el informe"
    ruta = os.path.join(ARBOL, "reportes", sorted(reps)[-1])

    df = pd.read_excel(ruta, sheet_name="Facturas")
    assert "hotel_id" in df.columns, (
        "la hoja Facturas del informe no lleva `hotel_id`. La etiqueta tiene que "
        "sobrevivir la cadena ENTERA: es la leccion que costo cara en AR, donde "
        "se estampo en la primera etapa y se dio por hecho el resto.")
    filas = {str(f["numero_factura"]): f for f in df.to_dict("records")}
    assert M._txt(filas["F-A1"]["hotel_id"]) == H_A, filas["F-A1"]["hotel_id"]

    # el hueco se EXPLICA
    det = str(filas["F-A2"]["detalle_matching"])
    assert "otro hotel" in det, (
        "F-A2 se queda sin albaran por el hotel y no lo dice. Un hueco mudo se "
        f"lee como 'no llego la mercancia' y acaba en una reclamacion que no "
        f"toca. Detalle: {det}")
    assert "Hotel Test B" in det, (
        f"el aviso tiene que NOMBRAR el hotel, no decir 'otro': {det}")
    det_x = str(filas["F-X2"]["detalle_matching"])
    assert "sin hotel asignado" in det_x, (
        f"una factura sin hotel tiene que decir que esta en su propia isla: {det_x}")

    # y las dos facturas del `factura_julio.pdf` cuadran, cada una con lo suyo
    for num in ("F-A4", "F-B4"):
        assert str(filas[num]["estado_matching"]) == "MATCH_ALBARAN_OK", (
            f"{num} es correcta y sale como {filas[num]['estado_matching']}: las "
            "lineas del otro hotel se estan colando por el nombre del fichero")
    # EL CORTE, por hotel. Con el corte global las dos serian FACTURA_SIN_ALBARAN:
    # una alerta que nadie puede accionar, en el hotel equivocado.
    assert str(filas["F-A5"]["estado_matching"]) == "ANTERIOR_AL_REGISTRO", (
        "F-A5 es de marzo y su hotel no registraba albaranes hasta julio; sale "
        f"como {filas['F-A5']['estado_matching']}. El corte se ha vuelto global: "
        "un hotel estaria recibiendo alertas por la fecha del primer albaran de OTRO.")
    assert "Hotel Test A" in str(filas["F-A5"]["detalle_matching"]), (
        f"el corte tiene que decir de QUE hotel habla: {filas['F-A5']['detalle_matching']}")
    assert str(filas["F-C5"]["estado_matching"]) == "ANTERIOR_AL_REGISTRO", (
        "el hotel C no ha registrado ni un albaran y su factura sale como "
        f"{filas['F-C5']['estado_matching']}. Un hotel que entra nuevo se llenaria "
        "de alertas el dia que OTRO hotel sube su primer albaran.")
    print("  ✔ el informe lleva hotel_id, explica cada hueco y no mezcla lineas")
    print("  ✔ el corte es por hotel, y un hotel sin albaranes no genera alertas")


def test_sin_fecha_no_silencia_nada():
    """La trampa: "este hotel no registra albaranes" NO se puede deducir de "no
    tiene corte". Si se dedujera, un tenant cuyos albaranes no traen fecha
    legible pasaria de alertar de todo a no alertar de NADA, en silencio. Y ese
    tenant es el caso de 0 hoteles, que es el que no se puede mover."""
    fac = [dict(f, hotel_id="") for f in FACTURAS]
    alb = [dict(a, hotel_id="", fecha_entrega="") for a in ALBARANES]
    _montar(facturas=fac, albaranes=alb, lin_f=[], lin_a=[])
    env = dict(os.environ, YVE_TENANT=TENANT, YVE_HOTEL="")
    subprocess.run([sys.executable, MODULO], cwd=BASE, capture_output=True,
                   text=True, timeout=300, env=env)
    reps = sorted(f for f in os.listdir(os.path.join(ARBOL, "reportes"))
                  if f.startswith("matching_albaran_"))
    df = pd.read_excel(os.path.join(ARBOL, "reportes", reps[-1]), sheet_name="Facturas")
    estados = {str(f["numero_factura"]): str(f["estado_matching"])
               for f in df.to_dict("records")}
    assert estados["F-A5"] == "FACTURA_SIN_ALBARAN", (
        "hay albaranes registrados, solo que sin fecha legible: eso NO es "
        f"'este hotel no registra albaranes'. F-A5 sale como {estados['F-A5']} y "
        "se estaria silenciando una alerta buena.")
    print("  ✔ albaranes sin fecha legible no silencian las alertas")


def test_sin_censo_los_estados_no_cambian():
    """El censo se BORRA en cada despliegue de Render. Si el cruce dependiera de
    el, un despliegue cambiaria los numeros. Solo puede empeorar el texto."""
    def _estados(censo):
        _montar(censo=censo)
        env = dict(os.environ, YVE_TENANT=TENANT, YVE_HOTEL="")
        subprocess.run([sys.executable, MODULO], cwd=BASE,
                       capture_output=True, text=True, timeout=300, env=env)
        reps = sorted(f for f in os.listdir(os.path.join(ARBOL, "reportes"))
                      if f.startswith("matching_albaran_"))
        df = pd.read_excel(os.path.join(ARBOL, "reportes", reps[-1]),
                           sheet_name="Facturas")
        return {str(f["numero_factura"]): str(f["estado_matching"])
                for f in df.to_dict("records")}
    con, sin = _estados(True), _estados(False)
    assert con == sin, f"el censo cambia los estados: {con} vs {sin}"
    print("  ✔ sin censo salen los mismos estados (solo cambia el texto)")


PRUEBAS = [test_dos_hoteles_no_cruzan, test_el_vacio_no_es_comodin,
           test_cero_hoteles_igual_que_siempre, test_un_hotel_igual_que_siempre,
           test_la_columna_puede_no_existir, test_el_nan_es_sin_asignar,
           test_lineas_no_se_mezclan_entre_hoteles,
           test_no_pregunta_por_el_hotel_activo,
           test_informe_completo, test_sin_fecha_no_silencia_nada,
           test_sin_censo_los_estados_no_cambian]

# Las que el sabotaje TIENE que tumbar. Las demas siguen pasando a proposito:
# el sabotaje solo quita la regla del hotel, no rompe el modulo entero.
DEBEN_CAER = {"test_dos_hoteles_no_cruzan", "test_el_vacio_no_es_comodin",
              "test_lineas_no_se_mezclan_entre_hoteles", "test_informe_completo"}


def main():
    print(f"\n{'SABOTAJE — ' if SABOTAJE else ''}Cruce factura↔albarán por hotel")
    print("=" * 62)
    fallos, caidas = [], set()
    try:
        for p in PRUEBAS:
            try:
                p()
            except AssertionError as e:
                caidas.add(p.__name__)
                if not SABOTAJE:
                    fallos.append(f"{p.__name__}: {e}")
                    print(f"  ✗ {p.__name__}\n      {e}")
    finally:
        _limpiar()
        # la copia saboteada NO se puede quedar en el arbol: acabaria commiteada
        if os.path.exists(_SABOTEADO):
            os.remove(_SABOTEADO)

    if SABOTAJE:
        no_gritaron = DEBEN_CAER - caidas
        print("=" * 62)
        if no_gritaron:
            print("  ✗ con la regla del hotel rota, estas NO han fallado: "
                  + ", ".join(sorted(no_gritaron)))
            print("    Un test que no puede fallar no protege de nada.")
            return 1
        print(f"  ✔ con la regla rota fallan las {len(DEBEN_CAER)} que tienen que fallar")
        return 0

    print("=" * 62)
    if fallos:
        print(f"  {len(fallos)} FALLO(S)")
        return 1
    print(f"  {len(PRUEBAS)} pruebas OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
