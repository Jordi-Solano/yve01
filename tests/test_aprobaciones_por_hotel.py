"""/aprobaciones-ap/ solo muestra las facturas del hotel elegido.

EL BUG: este panel globeaba `facturas_contabilizadas_*.xlsx` por su cuenta y NO
filtraba por hotel. En produccion salia "POR APROBAR 3" con facturas de tres
hoteles distintos mezcladas, y sin decir de cual era cada una.

Por que importa mas que en cualquier otra pantalla: aqui una persona APRUEBA, y
lo que aprueba es lo que Oracle contabiliza despues. El error no se queda en una
pantalla — acaba en el libro mayor del hotel equivocado.

Medido con cinco facturas de tres hoteles y una sin asignar:

  hotel elegido        ANTES               DESPUES
  Costa Azul           las 5               CA-001, CA-002
  Plaza Mayor          las 5               PM-001
  Ribera               las 5               RB-001
  (vista de grupo)     las 5               las 5

Se filtra con `solo_del_hotel_activo`, la misma pieza que usa el panel de AR, y
falla en CERRADO.

Y se cierra el agujero mudo que abre el filtro estricto: una factura SIN hotel
—residuo de antes de la separacion— no sale en ningun hotel y no hay forma de
aprobarla. Antes eso habria sido silencio; ahora el panel lo dice.

LO QUE NO SE TOCA, y esta comprobado aqui: las rutas. Este modulo lee
`facturas-procesadas/` y escribe `aprobaciones/aprobaciones_ap.xlsx` en la RAIZ,
que son EXACTAMENTE las que usa `oracle_lector_facturas`. Si el panel leyera el
arbol del tenant y Oracle la raiz, se aprobaria una cosa y se contabilizaria
otra. Arreglar eso es tocar Oracle y va en su propio paso.

`--sabotaje` quita el filtro y quita el aviso, y comprueba que las dos se notan.
"""
import json
import os
import shutil
import sys
from datetime import date

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

SABOTAJE = "--sabotaje" in sys.argv
HOY = date.today().strftime("%Y%m%d")
CA, PM, RB = "HTAPCA", "HTAPPM", "HTAPRB"

# El modulo resuelve sus rutas en la RAIZ, igual que Oracle, asi que la prueba
# trabaja sobre el tenant `default` — o estaria probando un camino que en
# produccion no existe. Se hace copia y se restaura.
os.environ["YVE_TENANT"] = "default"

ESPERADO = {CA: ["CA-001", "CA-002"], PM: ["PM-001"], RB: ["RB-001"]}
TODAS = ["CA-001", "CA-002", "PM-001", "RB-001", "SIN-001"]


def _f(num, prov, base, hotel, estado="MATCH_ALBARAN_OK"):
    cuota = round(base * 0.21, 2)
    return {"archivo": f"{num}.pdf", "numero_factura": num, "fecha": "28/07/2026",
            "nombre_proveedor": prov, "NIF_proveedor": "B1234",
            "descripcion_concepto": "Compras", "base_imponible": base,
            "porcentaje_iva": 21.0, "cuota_iva": cuota,
            "total_factura": round(base + cuota, 2), "hotel_id": hotel,
            "cuenta_debe_gasto": 6001, "estado_asignacion": "REGLA_PROVEEDOR",
            "estado_matching": estado, "detalle_matching": "",
            "departamento_po": "F&B", "tipo_proveedor": "FB"}


FACTURAS = [_f("CA-001", "Pescados Rias SL", 1000.0, CA),
            _f("CA-002", "Congelados Ebre SL", 800.0, CA, "DIFERENCIA_IMPORTE"),
            _f("PM-001", "Bodegas Priorat SL", 1500.0, PM),
            _f("RB-001", "Cafes Nord SA", 300.0, RB),
            _f("SIN-001", "Proveedor Antiguo SL", 250.0, "")]


class Entorno:
    """Monta el arbol de la RAIZ para la prueba y lo devuelve como estaba."""

    def __enter__(self):
        import pandas as pd
        from tenant_dirs import datos_dir, procesadas_dir, aprobaciones_dir
        self.dirs = [procesadas_dir(), aprobaciones_dir()]
        for d in self.dirs + [datos_dir()]:
            os.makedirs(d, exist_ok=True)
        self.copias = {}
        for d in self.dirs:
            c = d + ".test-backup"
            if os.path.isdir(c):
                shutil.rmtree(c)
            shutil.copytree(d, c)
            self.copias[d] = c
            for f in os.listdir(d):
                p = os.path.join(d, f)
                if os.path.isfile(p):
                    os.remove(p)
        self.hjson = os.path.join(datos_dir(), "hoteles.json")
        self.hprev = open(self.hjson).read() if os.path.exists(self.hjson) else None
        json.dump([{"id": CA, "nombre": "Hotel Costa Azul", "activo": True},
                   {"id": PM, "nombre": "Hotel Plaza Mayor", "activo": True},
                   {"id": RB, "nombre": "Hotel Ribera", "activo": True}],
                  open(self.hjson, "w"))
        pd.DataFrame(FACTURAS).to_excel(
            os.path.join(procesadas_dir(), f"facturas_contabilizadas_{HOY}.xlsx"), index=False)

        import dashboard as D
        D.app.config["TESTING"] = True
        D.app.config["WTF_CSRF_ENABLED"] = False
        self.c = D.app.test_client()
        self.c.post("/api/login", json={"username": "admin", "password": "admin123"})
        return self

    def __exit__(self, *a):
        for d, c in self.copias.items():
            if os.path.isdir(d):
                shutil.rmtree(d)
            shutil.copytree(c, d)
            shutil.rmtree(c)
        if self.hprev is not None:
            open(self.hjson, "w").write(self.hprev)
        return False

    def con_hotel(self, hotel):
        with self.c.session_transaction() as s:
            s["tenant_id"] = "default"
            if hotel:
                s["hotel_activo"] = hotel
            else:
                s.pop("hotel_activo", None)
        os.environ["YVE_HOTEL"] = hotel or ""
        return self.c


def test_cada_hotel_ve_solo_sus_facturas():
    fallos = []
    with Entorno() as e:
        for hotel, esperado in ESPERADO.items():
            c = e.con_hotel(hotel)
            fac = c.get("/aprobaciones-ap/api/facturas?estado=todas").get_json() or []
            nums = sorted(f.get("numero_factura", "?") for f in fac)
            if nums != esperado:
                fallos.append(f"con {hotel} elegido el panel ve {nums} y tendria "
                              f"que ver {esperado}")
            st = c.get("/aprobaciones-ap/api/stats").get_json() or {}
            if st.get("pendientes") != len(esperado):
                fallos.append(f"con {hotel} el tile 'Por aprobar' dice "
                              f"{st.get('pendientes')} y tendria que decir {len(esperado)}")
        # sin hotel elegido = vista de grupo: se ve todo, incluida la sin asignar
        c = e.con_hotel("")
        nums = sorted(f.get("numero_factura", "?") for f in
                      (c.get("/aprobaciones-ap/api/facturas?estado=todas").get_json() or []))
        if nums != TODAS:
            fallos.append(f"en vista de grupo el panel ve {nums} y tendria que ver {TODAS}")
    assert not fallos, (
        "el panel no respeta el hotel:\n      " + "\n      ".join(fallos) +
        "\n      Aqui una persona APRUEBA y lo que aprueba es lo que Oracle "
        "contabiliza: ensenar las facturas de otro hotel no se queda en la "
        "pantalla, acaba en el libro mayor equivocado.")
    print(f"  ✔ los {len(ESPERADO)} hoteles ven solo lo suyo, y la vista de grupo lo ve todo")


def test_la_factura_sin_hotel_no_desaparece_en_silencio():
    with Entorno() as e:
        c = e.con_hotel(CA)
        st = c.get("/aprobaciones-ap/api/stats").get_json() or {}
        assert st.get("sin_hotel") == 1, (
            f"el panel dice que hay {st.get('sin_hotel')!r} facturas sin hotel y "
            "hay 1. El filtro es de igualdad estricta: una factura sin hotel no "
            "sale en NINGUN hotel y no hay forma de aprobarla. Si el panel no lo "
            "dice, es un agujero mudo — la factura existe, Oracle no la va a "
            "contabilizar porque nadie la puede aprobar, y nadie se entera.")
        fac = c.get("/aprobaciones-ap/api/facturas?estado=todas").get_json() or []
        assert all(f.get("hotel") for f in fac), (
            "hay filas sin nombre de hotel en el panel: quien aprueba tiene que "
            f"ver de que hotel es cada factura. {[f.get('numero_factura') for f in fac if not f.get('hotel')]}")
    print("  ✔ la factura sin hotel se cuenta y se avisa, y cada fila dice su hotel")


def test_la_aprobacion_deja_constancia_del_hotel():
    import pandas as pd
    from tenant_dirs import aprobaciones_dir
    with Entorno() as e:
        c = e.con_hotel(CA)
        r = c.post("/aprobaciones-ap/api/accion", json={
            "numero_factura": "CA-001", "clave": "CA-001", "accion": "APROBADA",
            "comentario": "ok banco de pruebas", "departamento": "F&B"})
        assert (r.get_json() or {}).get("ok"), f"la aprobacion ha fallado: {r.get_json()}"
        ruta = os.path.join(aprobaciones_dir(), "aprobaciones_ap.xlsx")
        assert os.path.exists(ruta), "no se ha escrito aprobaciones_ap.xlsx"
        df = pd.read_excel(ruta)
        # ORACLE cruza por esta columna: no se puede haber movido
        assert "numero_factura" in df.columns and "CA-001" in set(df["numero_factura"].astype(str)), (
            "la columna `numero_factura` es la que lee el gate de Oracle y tiene "
            f"que traer la factura aprobada. Columnas: {list(df.columns)}")
        assert "accion" in df.columns and "APROBADA" in set(df["accion"].astype(str)), \
            "la columna `accion` es el gate de Oracle: sin ella no se contabiliza nada"
        assert "hotel_id" in df.columns and str(df["hotel_id"].iloc[0]) == CA, (
            f"la aprobacion no deja constancia del hotel: hotel_id="
            f"{df['hotel_id'].iloc[0] if 'hotel_id' in df.columns else '(no existe)'!r}. "
            "Una aprobacion sin saber de que hotel es no es auditable.")
    print("  ✔ la aprobacion guarda el hotel y no mueve lo que lee Oracle")


def test_las_rutas_siguen_siendo_las_de_oracle(ruta=None):
    """El panel y Oracle tienen que mirar el MISMO fichero.

    Si el panel se pasara al arbol del tenant y Oracle siguiera en la raiz, se
    aprobaria una cosa y se contabilizaria otra — un fallo mucho peor que el que
    se esta arreglando. Este invariante existe para que nadie lo "mejore" sin
    tocar Oracle a la vez.
    """
    import ast
    ruta = ruta or os.path.join(BASE, "app_aprobacion_ap.py")
    for fichero, nombres in ((ruta, ("PROCESADAS_DIR", "APROBACIONES_DIR")),
                             (os.path.join(BASE, "oracle_lector_facturas.py"),
                              ("PROCESADAS_DIR", "APROBACIONES_DIR"))):
        arbol = ast.parse(open(fichero, encoding="utf-8").read())
        vistos = {}
        for n in arbol.body:
            if isinstance(n, ast.Assign) and len(n.targets) == 1 \
                    and isinstance(n.targets[0], ast.Name) and n.targets[0].id in nombres:
                vistos[n.targets[0].id] = ast.unparse(n.value)
        for nom in nombres:
            assert nom in vistos, f"{os.path.basename(fichero)}: no encuentro {nom}"
            assert "BASE_DIR" in vistos[nom], (
                f"{os.path.basename(fichero)}: {nom} = {vistos[nom]} ya no sale de "
                "BASE_DIR. Este panel y Oracle tienen que leer el MISMO fichero: "
                "si uno se va al arbol del tenant y el otro se queda en la raiz, "
                "se aprueba una factura y se contabiliza otra. Cambiarlo es "
                "tocar Oracle, y eso va en su propio paso.")
    print("  ✔ el panel y Oracle siguen leyendo las mismas rutas (AST)")


PRUEBAS = [test_cada_hotel_ve_solo_sus_facturas,
           test_la_factura_sin_hotel_no_desaparece_en_silencio,
           test_la_aprobacion_deja_constancia_del_hotel,
           test_las_rutas_siguen_siendo_las_de_oracle]


def main():
    print(f"\n{'SABOTAJE — ' if SABOTAJE else ''}/aprobaciones-ap/ respeta el hotel elegido")
    print("=" * 70)

    if SABOTAJE:
        import app_aprobacion_ap as A
        malos = 0

        # (a) se quita el filtro por hotel: el panel vuelve a ensenar el grupo
        bueno = A.cargar_facturas_ap
        A.cargar_facturas_ap = A._facturas_crudas
        try:
            test_cada_hotel_ve_solo_sus_facturas()
            print("  ✗ sin el filtro por hotel, el test NO ha fallado.")
            malos += 1
        except AssertionError as e:
            print(f"  ✔ se quita el filtro por hotel:\n      {str(e)[:150]}")
        finally:
            A.cargar_facturas_ap = bueno

        # (b) se deja de contar lo que el filtro tira: agujero mudo otra vez
        bueno2 = A.facturas_sin_hotel
        A.facturas_sin_hotel = lambda: 0
        try:
            test_la_factura_sin_hotel_no_desaparece_en_silencio()
            print("  ✗ sin el aviso de facturas sin hotel, el test NO ha fallado.")
            malos += 1
        except AssertionError as e:
            print(f"  ✔ se deja de avisar de las facturas sin hotel:\n      {str(e)[:150]}")
        finally:
            A.facturas_sin_hotel = bueno2

        # (c) alguien "mejora" las rutas y las separa de las de Oracle
        src = open(os.path.join(BASE, "app_aprobacion_ap.py"), encoding="utf-8").read()
        viejo = 'PROCESADAS_DIR   = os.path.join(BASE_DIR, "facturas-procesadas")'
        assert src.count(viejo) == 1, "el sabotaje no encuentra la ruta que romper"
        copia = os.path.join(BASE, "app_aprobacion_ap_SABOTAJE.py")
        open(copia, "w", encoding="utf-8").write(src.replace(
            viejo, "from tenant_dirs import procesadas_dir\nPROCESADAS_DIR = procesadas_dir()", 1))
        try:
            try:
                test_las_rutas_siguen_siendo_las_de_oracle(copia)
                print("  ✗ con las rutas separadas de Oracle, el invariante NO ha fallado.")
                malos += 1
            except AssertionError as e:
                print(f"  ✔ las rutas se separan de las de Oracle:\n      {str(e)[:150]}")
        finally:
            if os.path.exists(copia):
                os.remove(copia)

        print("=" * 70)
        return 1 if malos else 0

    fallos = []
    for p in PRUEBAS:
        try:
            p()
        except AssertionError as e:
            fallos.append(p.__name__)
            print(f"  ✗ {p.__name__}\n      {e}")
    print("=" * 70)
    if fallos:
        print(f"  {len(fallos)} FALLO(S)")
        return 1
    print(f"  {len(PRUEBAS)} pruebas OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
