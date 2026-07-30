"""La FUGA del DRR entre hoteles: un hotel SIN DRR (Ribera/Faro) que ensenaba
el DRR de Costa Azul. Y la leccion de por que se escapo.

EL BUG no estaba en el backend. `drr_del_hotel` filtra bien: para un hotel sin
DRR devuelve None y `/api/stats_drr` responde `null`. El agujero era de
REPINTADO, en el navegador:

    async function loadDRR() {
      const data = await (await fetch('/api/stats_drr')).json();
      if (data) renderDRR(data);   // <- con data null, renderDRR NO se llamaba
    }

`renderDRR` YA tenia una rama de "sin datos" (el `if (!s || s.error)`), pero el
`if (data)` de `loadDRR` impedia que se llamara nunca con null. Resultado: al
cambiar a un hotel sin DRR el panel se quedaba con el del hotel anterior. Y como
`_invalidarPaneles` fuerza recargar el panel a la vista, se "recargaba"... y
seguia sin borrarse, en silencio.

Y ademas la rama de "sin datos" solo vaciaba `drr-metrics` (las tarjetas). El
panel DRR pinta SIETE sitios —tarjetas, calendario de dias, alertas, barra de
estado (con el NOMBRE del fichero), fecha de subida, barra de budget y el
grafico—: vaciar solo uno dejaba los otros seis con los numeros del otro hotel.

CUATRO COMPROBACIONES:

  1. CONTRATO DEL BACKEND — dos hoteles, uno con DRR y otro sin. El que tiene lo
     ensena; el que no, `/api/stats_drr` y `/api/drr_daily_chart` devuelven
     null. Es la respuesta de la que depende el frontend para vaciar.

  2. NO SE CONTAMINAN — despues de mirar el hotel vacio, el que tiene DRR
     sigue devolviendo EL SUYO. La separacion es por nombre de fichero, no por
     "el ultimo".

  3. loadDRR REPINTA SIEMPRE (JS) — no puede volver el `if (data)`: `renderDRR`
     tiene que llamarse tambien con null, que es quien vacia. Aqui vivia el bug.

  4. EL VACIO LIMPIA EL PANEL ENTERO (JS) — la rama de "sin datos" tiene que
     tocar el calendario, las alertas, la barra de estado y el grafico, no solo
     las tarjetas. Si no, el hotel vacio ensena media pantalla del otro.

`--sabotaje` devuelve cada fallo de JS por separado.
"""
import glob
import os
import re
import shutil
import subprocess
import sys
from datetime import date as _d

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

SABOTAJE = "--sabotaje" in sys.argv
TENANT = "test-drr-fuga"
CON, SIN = "HCONDRR", "HSINDRR"     # uno con DRR, otro sin nada

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
DIAS = [(1, _d(2026, 7, 1), [("REVENUE", "Room Revenue Transient", 0, 20000),
                             ("REVENUE", "Food Revenue", 0, 5000),
                             ("EXPENSES", "Salaries", 9000, 0)], 0)]


# ── entorno: DOS hoteles, uno con DRR ──────────────────────────────────────

def _entorno():
    """Tenant con dos hoteles; procesa el DRR SOLO del hotel `CON`."""
    import json
    from datetime import date
    os.environ["YVE_TENANT"] = TENANT
    os.environ["YVE_HOTEL"] = CON
    import dashboard as D
    from tenant_dirs import datos_dir, entrada_dir, reportes_dir
    from crear_drr import construir

    for dd in (datos_dir(), entrada_dir(), reportes_dir()):
        if os.path.isdir(dd):
            shutil.rmtree(dd)
        os.makedirs(dd, exist_ok=True)
    json.dump([{"id": CON, "nombre": "Hotel Con DRR", "activo": True},
               {"id": SIN, "nombre": "Hotel Sin DRR", "activo": True}],
              open(os.path.join(datos_dir(), "hoteles.json"), "w"))

    xlsm = construir(os.path.join(entrada_dir(), "drr.xlsm"), METRICAS, DIAS)
    subprocess.run([sys.executable, "lector_drr.py", xlsm], cwd=BASE,
                   capture_output=True, text=True,
                   env={**os.environ, "YVE_HOTEL": CON})
    hoy = date.today().strftime("%Y%m%d")
    informes = glob.glob(os.path.join(reportes_dir(), "drr_procesado_*.xlsx"))
    assert informes, "el lector no genero el DRR del hotel CON"
    # el fichero tiene que llevar el sufijo del hotel CON, no quedar suelto
    assert any(f"_{CON}_" in os.path.basename(p) for p in informes), (
        f"el DRR no lleva el sufijo del hotel: {[os.path.basename(p) for p in informes]}")

    D.app.config["TESTING"] = True
    c = D.app.test_client()
    c.post("/api/login", json={"username": "admin", "password": "admin123"})
    return D, c


def _poner_hotel(c, hid):
    with c.session_transaction() as s:
        s["tenant_id"] = TENANT
        if hid:
            s["hotel_activo"] = hid
        else:
            s.pop("hotel_activo", None)


# ── 1 · el contrato del backend ────────────────────────────────────────────

def test_el_hotel_sin_drr_responde_vacio():
    _D, c = _entorno()

    _poner_hotel(c, CON)
    st = c.get("/api/stats_drr").get_json()
    ch = c.get("/api/drr_daily_chart").get_json()
    assert st, "el hotel CON DRR tendria que devolver stats y devuelve vacio"
    assert ch and ch.get("dias") == [1], (
        f"el hotel CON DRR tendria que traer el dia [1] y trae {ch!r}")

    _poner_hotel(c, SIN)
    st_sin = c.get("/api/stats_drr").get_json()
    ch_sin = c.get("/api/drr_daily_chart").get_json()
    assert st_sin is None, (
        f"/api/stats_drr del hotel SIN DRR tendria que ser null y es {st_sin!r}. "
        "Si trae datos, el backend estaria filtrando mal — y no es el caso: la "
        "fuga era de repintado en el frontend, no de aqui.")
    assert ch_sin is None, (
        f"/api/drr_daily_chart del hotel SIN DRR tendria que ser null y es {ch_sin!r}")
    print("  ✔ hotel con DRR lo devuelve; hotel sin DRR devuelve null (backend limpio)")


# ── 2 · no se contaminan ───────────────────────────────────────────────────

def test_mirar_el_vacio_no_contamina_al_que_tiene():
    _D, c = _entorno()
    _poner_hotel(c, SIN)
    assert c.get("/api/stats_drr").get_json() is None
    # y ahora, de vuelta al que tiene: sigue siendo EL SUYO
    _poner_hotel(c, CON)
    st = c.get("/api/stats_drr").get_json()
    assert st, "tras mirar el hotel vacio, el que tiene DRR ha dejado de devolverlo"
    rev = (st.get("metricas", {}).get("Revenue PAR", {}) or {}).get("mtd")
    assert rev == "€79.20", (
        f"el RevPAR del hotel CON DRR sale «{rev}», esperaba «€79.20»: la vuelta "
        "no debe cambiar sus numeros")
    print("  ✔ tras el hotel vacio, Costa Azul sigue mostrando el suyo")


# ── utilidades JS ──────────────────────────────────────────────────────────

def _js(ruta=None):
    src = open(ruta or os.path.join(BASE, "dashboard.py"), encoding="utf-8").read()
    js = "\n".join(re.findall(r"<script[^>]*>(.*?)</script>", src, re.S))
    js = re.sub(r"/\*.*?\*/", "", js, flags=re.S)
    js = re.sub(r"^\s*//.*$", "", js, flags=re.M)
    return js


def _cuerpo_loadDRR(js):
    m = re.search(r"async function loadDRR\(\)\s*\{(.*?)\n\}", js, re.S)
    assert m, "no encuentro `loadDRR` en el JS"
    return m.group(1)


def _rama_vacio_renderDRR(js):
    m = re.search(r"function renderDRR\(s\)\s*\{\s*if \(!s \|\| s\.error\)\s*\{(.*?)\n\s*return;",
                  js, re.S)
    assert m, "no encuentro la rama de «sin datos» de `renderDRR`"
    return m.group(1)


# ── 3 · loadDRR repinta SIEMPRE ────────────────────────────────────────────

def test_loadDRR_repinta_tambien_el_vacio(ruta=None):
    js = _js(ruta)
    cuerpo = _cuerpo_loadDRR(js)
    plano = re.sub(r"\s+", "", cuerpo)
    assert "renderDRR(data)" in plano, (
        "`loadDRR` no llama a `renderDRR(data)`. Es quien pinta el panel.")
    assert "if(data)renderDRR" not in plano, (
        "ha vuelto el `if (data) renderDRR(data)` de `loadDRR`. Ese guard es EL "
        "bug: con un hotel sin DRR el backend devuelve null, el guard salta la "
        "llamada y el panel se queda con el DRR del hotel anterior. `renderDRR` "
        "tiene que llamarse tambien con null — es su rama de «sin datos» la que "
        "vacia.")
    print("  ✔ loadDRR llama a renderDRR siempre, tambien con null (no vuelve el `if (data)`)")


# ── 4 · el vacio limpia el panel entero ────────────────────────────────────

def test_el_vacio_limpia_todo_el_panel(ruta=None):
    js = _js(ruta)
    rama = _rama_vacio_renderDRR(js)
    faltan = [x for x in ("drr-days", "drr-alerts", "drr-status", "renderDRRChart")
              if x not in rama]
    assert not faltan, (
        f"la rama de «sin datos» de renderDRR no vacia {faltan}. El panel DRR "
        "pinta siete sitios (tarjetas, calendario, alertas, barra de estado con "
        "el NOMBRE del fichero, fecha, budget y grafico). Si el vacio solo toca "
        "las tarjetas, el hotel sin DRR ensena el calendario y el grafico del "
        "otro hotel: media pantalla equivocada.")
    print("  ✔ el vacio limpia calendario, alertas, barra de estado y grafico (no solo las tarjetas)")


PRUEBAS = [test_el_hotel_sin_drr_responde_vacio,
           test_mirar_el_vacio_no_contamina_al_que_tiene,
           test_loadDRR_repinta_tambien_el_vacio,
           test_el_vacio_limpia_todo_el_panel]


# ── sabotaje ────────────────────────────────────────────────────────────────

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
    ("vuelve el `if (data)` de loadDRR",
     test_loadDRR_repinta_tambien_el_vacio,
     [("    // Siempre, tambien con `data` null: renderDRR vacia el panel cuando no hay\n"
       "    // DRR. Con el viejo `if (data)` el vacio no se repintaba y quedaba a la\n"
       "    // vista el DRR del hotel anterior. Esa era la fuga de Ribera/Faro.\n"
       "    renderDRR(data);",
       "    if (data) renderDRR(data);")]),
    ("el vacio vuelve a tocar solo las tarjetas",
     test_el_vacio_limpia_todo_el_panel,
     [("    _vaciar('drr-metrics', '<div class=\"empty\"><p>' + _msg + '</p></div>');\n"
       "    _vaciar('drr-days', '');\n"
       "    _vaciar('drr-alerts', '');\n"
       "    var _st = document.getElementById('drr-status'); if (_st) _st.textContent = '';\n"
       "    var _up = document.getElementById('drr-last-upload'); if (_up) _up.textContent = '';\n"
       "    var _bb = document.getElementById('drr-budget-bar'); if (_bb) _bb.style.display = 'none';\n"
       "    renderDRRChart();",
       "    _vaciar('drr-metrics', '<div class=\"empty\"><p>' + _msg + '</p></div>');")]),
]


def main():
    print(f"\n{'SABOTAJE — ' if SABOTAJE else ''}la fuga del DRR entre hoteles")
    print("=" * 70)

    if SABOTAJE:
        malos = 0
        for nombre, prueba, cambios in SABOTAJES:
            copia = _copia(cambios, nombre)
            try:
                try:
                    prueba(copia)
                except AssertionError as e:
                    print(f"  ✔ {nombre}:\n      {str(e)[:150]}")
                    continue
                print(f"  ✗ {nombre}: el invariante NO ha fallado.")
                malos += 1
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
