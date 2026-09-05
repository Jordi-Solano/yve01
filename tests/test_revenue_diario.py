"""El Revenue Diario del DRR: que salga, y que cuando no salga diga por que.

EL BUG, y la lección de por qué duró tanto:

`dashboard` filtraba las filas de ingreso del Trial Balance por la cadena
`"INCOME"` escrita a mano. Los DRR reales traen `"REVENUE"`. El filtro casaba
CERO filas, así que el gráfico se construía con cero barras — y encima se veía,
porque la tarjeta tenía el atributo `style` DUPLICADO (en HTML gana el primero,
así que el `display:none` se descartaba) y el guard era `if (!d.dias) return`,
que en JavaScript no aborta con un array vacío porque `[]` es truthy.

Resultado: un panel entero vacío, visible, sin un solo aviso. Desde el primer
commit del lector. Y sobrevivió porque NO HABÍA NINGUNA PRUEBA que recorriera
este camino: `tests/fixtures_drr_fase_e.py` importa `tests/crear.py`, que nunca
se commiteó. Esta prueba y `tests/crear_drr.py` son eso que faltaba.

CUATRO COMPROBACIONES:

  1. DE PUNTA A PUNTA — se construye un DRR, se procesa con el lector de verdad
     y se conduce `/api/drr_daily_chart`. Con las DOS grafías, INCOME y REVENUE,
     en días distintos del mismo fichero: el lector acepta las dos y el panel
     tiene que tratarlas igual. Aquí es donde el bug se cae.

  2. LA AUSENCIA DICE POR QUÉ — un DRR con días pero sin ninguna fila de ingreso
     tiene que devolver `motivo` explicando qué secciones sí venían. Sin esto,
     "no hay datos" y "el filtro está roto" se ven exactamente igual, que es
     justo lo que pasó.

  3. UNA SOLA LISTA DE SECCIONES (AST) — el nombre de las secciones lo pone
     `lector_drr`. Si el panel vuelve a escribir "INCOME" a mano, el invariante
     grita.

  4. LO QUE SE VE (JS sin comentarios) — que el guard mire la longitud y no la
     existencia, que el `style` no esté duplicado, y que las tarjetas enseñen
     `Rooms Occupied` y `Rooms Revenue`, que son el numerador y el denominador
     con los que se comprueban el ADR y el RevPAR a mano.

Y de paso los decimales: un RevPAR de 83,70 no es 84, y ese redondeo se
propagaba al RevPAR ponderado del grupo.

`--sabotaje` devuelve cada uno de los cuatro fallos por separado.
"""
import ast
import glob
import json
import os
import re
import shutil
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

SABOTAJE = "--sabotaje" in sys.argv
TENANT = "test-rev-diario"
H = "HTRD01"

# 400 habitaciones, 30 días de mes -> 12.000 noches disponibles en el MTD.
# Los números están elegidos para que la división se pueda comprobar a mano:
#   ADR    = 950.400 / 7.200 = 132,00
#   RevPAR = 950.400 / 12.000 = 79,20   <- el que salía como 79
METRICAS = {
    "Occupancy %":   (0.62, 0.60, 0.61, 0.65),
    "Rooms Occupied": (248, 7200, 7320, 7800),
    "ADR":           (135.00, 132.00, 133.00, 138.00),
    "Revenue PAR":   (83.70, 79.20, 81.13, 89.70),
    "Rooms Revenue": (33480, 950400, 973560, 1076400),
    "Total Revenue": (40130, 1143250, 1170160, 1279900),
    "GOP":           (12440, 354400, 360000, 400000),
    "GOP %":         (0.31, 0.31, 0.31, 0.31),
    "Spend PAR":     (55.20, 53.80, 54.00, 55.00),
}

# Tres días. El 1 con la sección escrita REVENUE, el 2 con INCOME, el 3 con
# REVENUE y descuadrado. Las dos grafías en el MISMO fichero es lo que hace que
# esta prueba cace el bug: con solo una podría pasar por casualidad.
# Las fechas van como `date`, no como texto: es lo que trae un .xlsm de verdad,
# y el lector las escribe distinto segun el tipo de la celda. Con texto salian
# etiquetas como 'ul-01' en el eje — un fixture que no se parece al fichero real
# prueba otra cosa.
from datetime import date as _d

DIAS = [
    (1, _d(2026, 7, 1), [("REVENUE", "Room Revenue Transient", 0, 20000),
                       ("REVENUE", "Food Revenue", 0, 5000),
                       ("EXPENSES", "Salaries", 9000, 0)], 0),
    (2, _d(2026, 7, 2), [("INCOME", "Room Revenue Transient", 0, 18000),
                       ("INCOME", "Food Revenue", 0, 4000),
                       ("EXPENSES", "Salaries", 8500, 0)], 0),
    (3, _d(2026, 7, 3), [("REVENUE", "Room Revenue Transient", 0, 22000),
                       ("REVENUE", "Food Revenue", 0, 6000),
                       ("EXPENSES", "Salaries", 9200, 0)], 1250),
]
# El endpoint devuelve el valor ABSOLUTO de la suma de los totales de cada
# seccion de ingreso, o sea debe - haber en valor absoluto.
ESPERADO_REVENUE = {1: 25000.0, 2: 22000.0, 3: 28000.0}
ESPERADO_GASTOS = {1: 9000.0, 2: 8500.0, 3: 9200.0}


def _entorno(dias=DIAS, metricas=None):
    """Deja un DRR procesado en el tenant y devuelve un cliente listo."""
    os.environ["YVE_TENANT"] = TENANT
    os.environ["YVE_HOTEL"] = H
    import dashboard as D
    from tenant_dirs import datos_dir, entrada_dir, reportes_dir
    from crear_drr import construir

    for d in (datos_dir(), entrada_dir(), reportes_dir()):
        if os.path.isdir(d):
            shutil.rmtree(d)
        os.makedirs(d, exist_ok=True)
    json.dump([{"id": H, "nombre": "Hotel Revenue Diario", "activo": True}],
              open(os.path.join(datos_dir(), "hoteles.json"), "w"))

    xlsm = construir(os.path.join(entrada_dir(), "drr_prueba.xlsm"),
                     METRICAS if metricas is None else metricas, dias)
    r = subprocess.run([sys.executable, "lector_drr.py", xlsm], cwd=BASE,
                       capture_output=True, text=True, env={**os.environ})
    informes = sorted(glob.glob(os.path.join(reportes_dir(), "drr_*.xlsx")))
    assert informes, (f"el lector no ha generado informe (rc={r.returncode}). "
                      f"stdout: {(r.stdout or '')[-400:]} stderr: {(r.stderr or '')[-300:]}")

    D.app.config["TESTING"] = True
    c = D.app.test_client()
    c.post("/api/login", json={"username": "admin", "password": "admin123"})
    with c.session_transaction() as s:
        s["tenant_id"] = TENANT
        s["hotel_activo"] = H
    return D, c


# ── 1 · de punta a punta ──────────────────────────────────────────────────

def test_el_revenue_diario_sale_con_las_dos_grafias():
    _D, c = _entorno()
    g = c.get("/api/drr_daily_chart").get_json()
    assert g and not g.get("error"), f"el endpoint no ha contestado bien: {g}"

    assert g.get("dias") == [1, 2, 3], (
        f"el grafico trae los dias {g.get('dias')!r} y tendria que traer [1, 2, 3]. "
        "Si viene vacio, el filtro de secciones no casa con lo que escribe el "
        "lector — es EL bug: el panel buscaba 'INCOME' y los ficheros reales "
        "dicen 'REVENUE'.")

    for i, dia in enumerate(g["dias"]):
        assert g["revenue"][i] == ESPERADO_REVENUE[dia], (
            f"el dia {dia} trae {g['revenue'][i]} de ingresos y tendria que "
            f"traer {ESPERADO_REVENUE[dia]}. El dia 2 usa la grafia INCOME y "
            "los otros REVENUE: si falla solo uno, el panel esta tratando las "
            "dos grafias de forma distinta.")
        assert g["expenses"][i] == ESPERADO_GASTOS[dia], \
            f"el dia {dia} trae {g['expenses'][i]} de gastos, esperaba {ESPERADO_GASTOS[dia]}"

    assert g.get("oob") == [False, False, True], (
        f"los descuadres salen {g.get('oob')!r} y el descuadrado es el dia 3")
    assert g.get("oob_count") == 1, (
        f"oob_count dice {g.get('oob_count')!r}. Tiene que contar los dias que "
        "se ESTAN MOSTRANDO: antes se calculaba sobre el fichero entero y decia "
        "a la vez 'cero dias' y 'un dia descuadrado'.")
    assert g.get("labels") == ["07-01", "07-02", "07-03"], \
        f"las etiquetas salen {g.get('labels')!r}"
    print(f"  ✔ los 3 dias con las 2 grafias: {g['revenue']} (endpoints y lector de verdad)")


# ── 2 · la ausencia dice por que ──────────────────────────────────────────

def test_cuando_no_hay_revenue_se_dice_por_que():
    # un DRR con dias, pero sin una sola fila de ingreso
    solo_gastos = [(1, _d(2026, 7, 1), [("EXPENSES", "Salaries", 9000, 0)], 0)]
    _D, c = _entorno(dias=solo_gastos)
    g = c.get("/api/drr_daily_chart").get_json()
    assert g.get("dias") == [], f"esperaba ningun dia y hay {g.get('dias')!r}"
    motivo = g.get("motivo") or ""
    assert motivo, (
        "el endpoint devuelve cero dias y NO dice por que. Es el fallo de fondo "
        "de todo esto: 'no hay datos' y 'el filtro esta roto' se veian igual, y "
        "por eso el bug duro desde el primer commit. Una ausencia tambien tiene "
        "que decir de donde viene.")
    assert "EXPENSES" in motivo.upper(), (
        f"el motivo es «{motivo}» y tendria que nombrar las secciones que SI "
        "venian en el fichero: es el dato con el que se diagnostica.")
    print(f"  ✔ sin ingresos, el endpoint explica: «{motivo[:70]}»")


# ── 3 · una sola lista de secciones ───────────────────────────────────────

def test_las_secciones_las_nombra_el_lector(ruta=None):
    ruta = ruta or os.path.join(BASE, "dashboard.py")
    src = open(ruta, encoding="utf-8").read()

    import lector_drr
    assert hasattr(lector_drr, "SECCIONES_INGRESO"), (
        "`lector_drr.SECCIONES_INGRESO` no existe. El vocabulario del Trial "
        "Balance vive donde se lee el fichero, no repartido por el panel.")
    assert {"INCOME", "REVENUE"} <= set(lector_drr.SECCIONES_INGRESO), (
        f"SECCIONES_INGRESO es {set(lector_drr.SECCIONES_INGRESO)} y tiene que "
        "aceptar las DOS grafias: los ficheros reales usan REVENUE y el lector "
        "documenta INCOME como valida.")

    # y el panel no puede tener la cadena escrita a mano en una comparacion
    arbol = ast.parse(src)
    a_mano = []
    for n in ast.walk(arbol):
        if isinstance(n, ast.Compare) and isinstance(n.comparators[0], ast.Constant) \
                and n.comparators[0].value in ("INCOME", "REVENUE"):
            a_mano.append(n.lineno)
        if isinstance(n, ast.Constant) and n.value in ("INCOME", "REVENUE") \
                and isinstance(getattr(n, "parent", None), ast.List):
            a_mano.append(n.lineno)
    assert not a_mano, (
        f"el panel compara con la cadena 'INCOME'/'REVENUE' a mano en la(s) "
        f"linea(s) {a_mano}. Ahi estaba el bug: dos listas de secciones que se "
        "desalinearon. Usa `lector_drr.SECCIONES_INGRESO`.")
    print("  ✔ las secciones salen de `lector_drr`, no de una cadena del panel (AST)")


# ── 4 · lo que se ve ──────────────────────────────────────────────────────

def _js(ruta=None):
    src = open(ruta or os.path.join(BASE, "dashboard.py"), encoding="utf-8").read()
    js = "\n".join(re.findall(r"<script[^>]*>(.*?)</script>", src, re.S))
    js = re.sub(r"/\*.*?\*/", "", js, flags=re.S)
    js = re.sub(r"^\s*//.*$", "", js, flags=re.M)
    return js


def test_la_pantalla_no_pinta_un_grafico_vacio(ruta=None):
    src = open(ruta or os.path.join(BASE, "dashboard.py"), encoding="utf-8").read()

    # (a) el `style` duplicado en la tarjeta del grafico
    tarjeta = re.search(r"<div[^>]*id=\"drr-chart-card\"[^>]*>", src)
    assert tarjeta, "no encuentro la tarjeta `drr-chart-card`"
    assert tarjeta.group(0).count("style=") == 1, (
        f"la tarjeta del grafico tiene {tarjeta.group(0).count('style=')} atributos "
        "`style`. En HTML gana el PRIMERO, asi que el segundo —el que llevaba "
        "`display:none`— se descartaba y la tarjeta salia visible siempre, con "
        "datos o sin ellos.")

    # (b) el guard tiene que mirar la LONGITUD, no la existencia
    js = _js(ruta)
    fn = re.search(r"async function renderDRRChart\(\)\s*\{(.*?)\n\}", js, re.S)
    assert fn, "no encuentro `renderDRRChart` en el JS"
    cuerpo = fn.group(1)
    assert re.search(r"d\.dias\.length\s*===?\s*0|!d\.dias\.length|"
                     r"length\s*<\s*1|Array\.isArray\(d\.dias\)", cuerpo), (
        "`renderDRRChart` no comprueba la LONGITUD de `dias`. Con `!d.dias` no "
        "aborta nunca, porque en JavaScript un array vacio es truthy: enseñaba "
        "la tarjeta y dibujaba un grafico de cero barras.")
    assert "motivo" in cuerpo, (
        "cuando no hay dias, la pantalla no dice por que. El endpoint manda "
        "`motivo` justamente para eso: un hueco silencioso es lo que costo la "
        "prueba de integracion.")
    print("  ✔ un solo `style`, el guard mira la longitud, y el hueco se explica (HTML + JS)")


def test_las_tarjetas_dejan_comprobar_el_adr_a_mano(ruta=None):
    # Tras el rediseño ya no hay lista `SHOW`: las tarjetas se construyen con
    # `tile(...)` por grupo. El invariante es el mismo — el panel tiene que
    # seguir enseñando el numerador (Rooms Revenue) y el denominador (Rooms
    # Occupied) con los que se comprueba a mano el ADR/RevPAR.
    js = _js(ruta)
    fn = re.search(r"async function renderDRR\(s\)\s*\{(.*?)\n\}", js, re.S)
    assert fn, "no encuentro `renderDRR`"
    cuerpo = fn.group(1)
    faltan = [k for k in ("Rooms Occupied", "Rooms Revenue") if k not in cuerpo]
    assert not faltan, (
        f"renderDRR no enseña {faltan}. El ADR y el RevPAR se comprueban a mano "
        "con Rooms Occupied (denominador) y Rooms Revenue (numerador): tienen que "
        "salir en el panel.")
    print("  ✔ el panel enseña Rooms Occupied y Rooms Revenue")


def test_el_revpar_conserva_sus_decimales():
    """Un RevPAR de 83,70 no es 84. Y el redondeo no se queda en la pantalla:
    `agregador_grupo` reparsea ESTA cadena para ponderar el RevPAR del grupo."""
    import dashboard as D
    _D, c = _entorno()
    st = c.get("/api/stats_drr").get_json() or {}
    m = st.get("metricas") or {}

    revpar = (m.get("Revenue PAR") or {})
    # Sep 2026 (Jordi): el DRR en formato español, como el resto del panel
    assert revpar.get("mtd") == "79,20 €", (
        f"el RevPAR MTD sale «{revpar.get('mtd')}» y tiene que ser «79,20 €» "
        "(950.400 / 12.000). Con «€79» se pierden 20 centimos por habitacion "
        "disponible, y el agregador del grupo hereda el error.")
    assert revpar.get("today") == "83,70 €", \
        f"el RevPAR de hoy sale «{revpar.get('today')}», esperaba «83,70 €»"

    # y el recuento de habitaciones no lleva decimales ni moneda
    ocup = (m.get("Rooms Occupied") or {})
    assert ocup.get("mtd") == "7,200", (
        f"las habitaciones ocupadas salen «{ocup.get('mtd')}» y son «7,200»: un "
        "recuento no lleva decimales.")

    # el agregador tiene que poder volver a leer lo que el panel escribe
    assert D.num_drr(revpar["mtd"]) == 79.20, (
        f"`num_drr` no sabe releer «{revpar['mtd']}» ({D.num_drr(revpar['mtd'])}). "
        "Si no, el RevPAR ponderado del grupo se calcula con otro numero.")
    print("  ✔ RevPAR 79,20 € y 83,70 €, habitaciones 7,200, y el agregador los relee")


PRUEBAS = [test_el_revenue_diario_sale_con_las_dos_grafias,
           test_cuando_no_hay_revenue_se_dice_por_que,
           test_las_secciones_las_nombra_el_lector,
           test_la_pantalla_no_pinta_un_grafico_vacio,
           test_las_tarjetas_dejan_comprobar_el_adr_a_mano,
           test_el_revpar_conserva_sus_decimales]


# ── sabotaje ──────────────────────────────────────────────────────────────

def _copia(cambios, que):
    src = open(os.path.join(BASE, "dashboard.py"), encoding="utf-8").read()
    for viejo, nuevo in cambios:
        assert src.count(viejo) == 1, (
            f"el sabotaje de «{que}» ya no encuentra que romper "
            f"({src.count(viejo)} apariciones): hay que ponerlo al dia")
        src = src.replace(viejo, nuevo, 1)
    dst = os.path.join(BASE, "dashboard_SABOTAJE.py")
    open(dst, "w", encoding="utf-8").write(src)
    return dst


SABOTAJES = [
    ("el panel vuelve a escribir 'INCOME' a mano",
     test_las_secciones_las_nombra_el_lector,
     [("        income = df[_secc.isin(SECCIONES_INGRESO)].copy()",
       '        income = df[df["Sección"] == "INCOME"].copy()')]),
    ("vuelve el `style` duplicado en la tarjeta",
     test_la_pantalla_no_pinta_un_grafico_vacio,
     [('<div class="card" id="drr-chart-card" style="margin-bottom:20px">',
       '<div class="card" style="x" id="drr-chart-card" style="margin-bottom:20px">')]),
    ("las tarjetas pierden Rooms Occupied",
     test_las_tarjetas_dejan_comprobar_el_adr_a_mano,
     [("tile('Hab. ocupadas', 'Rooms Occupied', '', 'El denominador del ADR y el numerador de la ocupación')",
       "tile('Hab. ocupadas', 'Occupancy %', '', '')")]),
]


def main():
    print(f"\n{'SABOTAJE — ' if SABOTAJE else ''}el Revenue Diario del DRR")
    print("=" * 70)

    if SABOTAJE:
        malos = 0
        for nombre, prueba, cambios in SABOTAJES:
            copia = _copia(cambios, nombre)
            try:
                try:
                    prueba(copia)
                except AssertionError as e:
                    print(f"  ✔ {nombre}:\n      {str(e)[:140]}")
                    continue
                print(f"  ✗ {nombre}: el invariante NO ha fallado.")
                malos += 1
            finally:
                if os.path.exists(copia):
                    os.remove(copia)

        # el de punta a punta: se sabotea la lista de secciones en memoria, que
        # es donde vive la regla ahora
        import lector_drr
        bueno = lector_drr.SECCIONES_INGRESO
        lector_drr.SECCIONES_INGRESO = frozenset({"INCOME"})   # como estaba antes
        try:
            test_el_revenue_diario_sale_con_las_dos_grafias()
            print("  ✗ con las secciones reducidas a INCOME, el test NO ha fallado.")
            malos += 1
        except AssertionError as e:
            print(f"  ✔ las secciones vuelven a ser solo INCOME:\n      {str(e)[:140]}")
        finally:
            lector_drr.SECCIONES_INGRESO = bueno

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
