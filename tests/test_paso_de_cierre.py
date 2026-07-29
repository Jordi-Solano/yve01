"""El pipeline llega al final por CUALQUIER puerta, tambien por la foto.

EL BUG, en dos mitades con la misma causa —"el paso siguiente no se dispara":

  · El LOTE cerraba dentro de su propio EventSource, que el frontend corta a los
    60 s. El cruce y el asignador son dos subprocesos con 180 s de margen cada
    uno, asi que en el ultimo lote no llegaban a arrancar. Medido: el lote 2
    empezo a las 22:25:50 con el corte en 22:26:50 y el bloque anterior acabo en
    22:26:35 — quince segundos para dos subprocesos.

  · La FOTO no cerraba NUNCA. `/api/scan_documento` es un POST por imagen y no
    tenia paso de cierre ninguno. Una foto de un albaran se guardaba y no
    relanzaba el cruce; y peor, una foto de una FACTURA se guardaba, salia en
    verde, y no llegaba jamas a Aprobaciones AP porque nadie le asignaba cuenta
    ni asiento. El documento existia para el usuario y no para quien lo aprueba.

El arreglo saca el cierre a `/api/cerrar_pipeline_stream`, que el frontend llama
UNA vez cuando ha terminado todo. Eso reparte la garantia en dos: el servidor
tiene que ofrecer el paso, y el frontend tiene que llamarlo. Aqui se comprueban
las dos, porque con una sola el bug sobrevive:

  1. DE PUNTA A PUNTA — una foto de un albaran, con la vision mockeada, y
     despues el cierre: el cruce corre y la factura llega a
     `facturas_contabilizadas`, que es el fichero que lee Aprobaciones AP.
  2. INVARIANTE (AST) — ni el lote ni el scan lanzan los subprocesos del cierre
     por su cuenta, y el lote anuncia lo que deja pendiente.
  3. INVARIANTE (JS, sin comentarios) — el frontend llama al cierre en los DOS
     caminos: al acabar los lotes y al acabar las fotos. Un cierre que el
     servidor ofrece y el frontend no llama es exactamente el bug de antes.

`--sabotaje` rompe cada invariante en una COPIA y comprueba que grita. El
dashboard son 730 KB y registra rutas de Flask: la copia se analiza con AST y
con texto, nunca se importa.
"""
import ast
import glob
import io
import json
import os
import re
import shutil
import sys
from datetime import date

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

HOY = date.today().strftime("%Y%m%d")
TENANT = "test-cierre"
H_A = "HCIERRET"
SABOTAJE = "--sabotaje" in sys.argv

# Los subprocesos del cierre. Lanzarlos desde dentro del stream del lote es
# justo el bug: mueren con la conexion.
SUBPROCESOS = ("matching_ap_albaran.py", "asignador_cuentas.py")
# Las funciones que NO pueden lanzarlos: son las que atienden una conexion que
# el navegador corta por su cuenta.
PROHIBIDO_EN = ("api_procesar_batch_stream", "api_scan_documento")


# ── 2 · invariante del servidor (AST) ─────────────────────────────────────

def _funciones_que_lanzan(ruta):
    """{nombre de funcion: [subprocesos que lanza]}, por anidamiento real."""
    arbol = ast.parse(open(ruta, encoding="utf-8").read())
    fuera = {}

    def rec(nodo, pila):
        for h in ast.iter_child_nodes(nodo):
            nueva = pila + [h.name] if isinstance(h, (ast.FunctionDef, ast.AsyncFunctionDef)) else pila
            if isinstance(h, ast.Constant) and h.value in SUBPROCESOS:
                for fn in nueva:
                    fuera.setdefault(fn, []).append(h.value)
            rec(h, nueva)

    rec(arbol, [])
    return fuera


def test_el_cierre_no_vive_dentro_del_stream(ruta=None):
    ruta = ruta or os.path.join(BASE, "dashboard.py")
    lanzan = _funciones_que_lanzan(ruta)

    for fn in PROHIBIDO_EN:
        dentro = lanzan.get(fn, [])
        assert not dentro, (
            f"`{fn}` lanza {dentro} por su cuenta. Ahi es donde estaba el bug: "
            "esa funcion atiende una conexion que el navegador corta —el lote a "
            "los 60 s, la foto al acabar el POST— y el subproceso se queda sin "
            "arrancar. El cierre va en `_generar_cierre`, que tiene su propia "
            "conexion y su propio reloj.")

    falta = [s for s in SUBPROCESOS if s not in lanzan.get("_generar_cierre", [])]
    assert not falta, (
        f"`_generar_cierre` no lanza {falta}. Si el paso de cierre no los lanza "
        "y el lote tampoco, no los lanza NADIE: los documentos se guardan y no "
        "se cruzan ni se contabilizan nunca.")
    print(f"  ✔ los {len(SUBPROCESOS)} subprocesos del cierre solo se lanzan desde `_generar_cierre` (AST)")


def test_el_lote_anuncia_lo_que_deja_pendiente(ruta=None):
    """El lote ya no cierra: tiene que DECIR lo que queda, o el frontend no
    sabe que pedir y el cierre no corre para nada."""
    ruta = ruta or os.path.join(BASE, "dashboard.py")
    arbol = ast.parse(open(ruta, encoding="utf-8").read())
    visto = set()

    def rec(nodo, pila):
        for h in ast.iter_child_nodes(nodo):
            nueva = pila + [h.name] if isinstance(h, (ast.FunctionDef, ast.AsyncFunctionDef)) else pila
            # el `yield f'data: CIERRE_PENDIENTE:...'` del lote
            if isinstance(h, ast.JoinedStr):
                crudo = "".join(v.value for v in h.values if isinstance(v, ast.Constant))
                if "CIERRE_PENDIENTE" in crudo:
                    visto.update(nueva)
            # el `'cierre': [...]` que devuelve el scan
            if isinstance(h, ast.Dict):
                for k in h.keys:
                    if isinstance(k, ast.Constant) and k.value == "cierre":
                        visto.update(nueva)
            rec(h, nueva)

    rec(arbol, [])
    for fn in PROHIBIDO_EN:
        assert fn in visto, (
            f"`{fn}` no dice lo que deja pendiente de cerrar. Guarda documentos "
            "y no cierra: si tampoco lo anuncia, el frontend no tiene forma de "
            "saber que hay que cruzar ni que hay que contabilizar, y el "
            "documento se queda a medias sin que nadie se entere.")
    print(f"  ✔ las {len(PROHIBIDO_EN)} puertas de entrada anuncian lo que dejan pendiente (AST)")


# ── 3 · invariante del frontend (JS sin comentarios) ──────────────────────

def _js_sin_comentarios(ruta=None):
    """El JS del dashboard con los comentarios fuera.

    Sin quitarlos, el invariante lo cumpliria un comentario que dijera
    `_correrCierre(...)`, que es exactamente no comprobar nada.
    """
    ruta = ruta or os.path.join(BASE, "dashboard.py")
    src = open(ruta, encoding="utf-8").read()
    js = "\n".join(re.findall(r"<script[^>]*>(.*?)</script>", src, re.S))
    js = re.sub(r"/\*.*?\*/", "", js, flags=re.S)
    js = re.sub(r"^\s*//.*$", "", js, flags=re.M)
    return js


def test_el_frontend_llama_al_cierre(ruta=None):
    js = _js_sin_comentarios(ruta)
    assert "function _correrCierre" in js, (
        "no existe `_correrCierre` en el JS: el servidor ofrece el paso de "
        "cierre y nadie lo llama, o sea que no corre nunca.")

    # los DOS caminos, cada uno con su llamada
    caminos = {
        "al acabar los lotes (_finish)":
            re.search(r"function\s+_finish\s*\([^)]*\)\s*\{(.*?)\n  \}", js, re.S),
        "al acabar las fotos (solo fotos, sin documentos)":
            re.search(r"\}\s*else\s*\{[^{}]*?_correrCierre", js, re.S),
    }
    faltan = [k for k, m in caminos.items()
              if not m or "_correrCierre" not in m.group(0)]
    assert not faltan, (
        f"el frontend NO llama al cierre {faltan}. Es la mitad del bug que no "
        "se ve en el servidor: el endpoint existe y funciona, pero si el "
        "navegador no lo llama al terminar, los documentos se quedan guardados "
        "y sin cruzar igual que antes. El camino de las fotos es el que estaba "
        "roto — no lo dejes suelto otra vez.")

    # y el reloj del cierre NO puede ser el del lote: son dos subprocesos de
    # 180 s, 60 s no llegan
    relojes = [int(n) for n in re.findall(r"_acabar\(true\);\s*\n\s*\},\s*(\d+)\)", js)]
    assert relojes and min(relojes) >= 300000, (
        f"el reloj del cierre es {relojes} ms. El cierre lanza dos subprocesos "
        "con 180 s de margen cada uno: con menos de 300 s se corta a media "
        "faena, que es el bug de siempre con otro nombre.")
    print("  ✔ el frontend llama al cierre en los 2 caminos, con reloj de "
          f"{min(relojes) // 1000} s (JS sin comentarios)")


# ── 1 · de punta a punta ──────────────────────────────────────────────────

def test_una_foto_de_albaran_cierra_el_pipeline():
    import pandas as pd
    import anthropic
    os.environ["YVE_TENANT"] = TENANT
    os.environ["YVE_HOTEL"] = H_A
    os.environ.setdefault("ANTHROPIC_API_KEY", "test-no-se-usa")
    import dashboard as D
    from tenant_dirs import datos_dir, entrada_dir, procesadas_dir, reportes_dir

    for d in (entrada_dir(), procesadas_dir(), reportes_dir()):
        if os.path.isdir(d):
            shutil.rmtree(d)
    for d in (datos_dir(), entrada_dir(), procesadas_dir(), reportes_dir()):
        os.makedirs(d, exist_ok=True)
    json.dump([{"id": H_A, "nombre": "Hotel Cierre", "activo": True}],
              open(os.path.join(datos_dir(), "hoteles.json"), "w"))
    for f in ("proveedores.xlsx", "plan_cuentas.xlsx"):
        src = os.path.join(BASE, "datos-referencia", f)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(datos_dir(), f))

    # la factura de ayer, la que la entrega nueva tiene que re-evaluar
    pd.DataFrame([{
        "archivo": "previa.pdf", "numero_factura": "FP-9001", "fecha": "22/07/2026",
        "nombre_proveedor": "PESCADOS RIAS S.L.", "base_imponible": 1000.0,
        "porcentaje_iva": 21.0, "cuota_iva": 210.0, "total_factura": 1210.0,
        "descripcion_concepto": "", "hotel_id": H_A,
    }]).to_excel(os.path.join(procesadas_dir(), f"facturas_ap_{HOY}.xlsx"),
                 sheet_name="Facturas", index=False)

    REG = {"tipo_documento": "ALBARAN", "numero_albaran": "ALB-9001",
           "nombre_proveedor": "PESCADOS RIAS S.L.", "fecha_entrega": "18/07/2026",
           "total_albaran": 1000.0,
           "lineas": [{"descripcion": "Merluza", "cantidad": 50,
                       "precio_unitario": 20.0, "importe": 1000.0}]}

    class _Msgs:
        def create(self, *a, **k):
            class _T: text = json.dumps(REG)
            class _R: content = [_T()]
            return _R()

    class _Cli:
        def __init__(self, *a, **k): self.messages = _Msgs()

    anthropic.Anthropic = _Cli   # la vision, mockeada: no hay clave en el sandbox

    D.app.config["TESTING"] = True
    D.app.config["WTF_CSRF_ENABLED"] = False
    c = D.app.test_client()
    c.post("/api/login", json={"username": "admin", "password": "admin123"})
    with c.session_transaction() as s:
        s["tenant_id"] = TENANT
        s["hotel_activo"] = H_A

    r = c.post("/api/scan_documento", content_type="multipart/form-data",
               data={"image": (io.BytesIO(b"\xff\xd8\xff foto"), "albaran.jpg")})
    j = r.get_json() or {}
    assert j.get("ok") and j.get("tipo") == "ALBARAN", f"la foto no se ha leido: {j}"
    assert j.get("cierre") == ["albaran"], (
        f"la foto no pide cerrar: cierre={j.get('cierre')!r}. Es el bug: el "
        "albaran se guarda y el cruce no se relanza nunca.")

    rr = c.get("/api/cerrar_pipeline_stream?pasos=" + ",".join(j["cierre"]))
    log = [l[5:].strip() for l in rr.get_data(as_text=True).splitlines()
           if l.startswith("data:")]
    assert any("Cruzando facturas con albaranes" in l for l in log), \
        f"el cierre no ha cruzado: {log}"
    assert any("Asignando cuentas contables" in l for l in log), \
        f"el cierre no ha contabilizado: {log}"

    # la prueba de verdad: el fichero que lee Aprobaciones AP
    cont = glob.glob(os.path.join(procesadas_dir(), "facturas_contabilizadas_*.xlsx"))
    assert cont, ("no se ha generado `facturas_contabilizadas`. Es lo que lee "
                  "Aprobaciones AP: sin ese fichero, el documento no existe "
                  "para quien tiene que aprobarlo.")
    df = pd.read_excel(cont[0])
    assert "FP-9001" in set(df.get("numero_factura", pd.Series()).astype(str)), \
        f"la factura no ha llegado a contabilizadas: {df.to_dict('records')}"
    print("  ✔ una foto de un albaran cruza y contabiliza (endpoints de verdad)")

    for d in (entrada_dir(), procesadas_dir(), reportes_dir()):
        if os.path.isdir(d):
            shutil.rmtree(d)


PRUEBAS = [test_el_cierre_no_vive_dentro_del_stream,
           test_el_lote_anuncia_lo_que_deja_pendiente,
           test_el_frontend_llama_al_cierre,
           test_una_foto_de_albaran_cierra_el_pipeline]


# ── sabotaje ──────────────────────────────────────────────────────────────

def _copia(cambios, que):
    """Copia de dashboard.py con `cambios` aplicados. [(viejo, nuevo)]"""
    src = open(os.path.join(BASE, "dashboard.py"), encoding="utf-8").read()
    for viejo, nuevo in cambios:
        assert src.count(viejo) == 1, (
            f"el sabotaje de «{que}» ya no encuentra lo que tiene que romper "
            f"({src.count(viejo)} apariciones): el test ha dejado de saber que "
            "sabotear y hay que ponerlo al dia")
        src = src.replace(viejo, nuevo, 1)
    dst = os.path.join(BASE, "dashboard_SABOTAJE.py")
    open(dst, "w", encoding="utf-8").write(src)
    return dst


SABOTAJES = [
    ("el cierre vuelve dentro del stream del lote",
     test_el_cierre_no_vive_dentro_del_stream,
     [("            _pend = []",
       "            import subprocess as _x\n"
       "            _x.run(['python3', 'asignador_cuentas.py'])\n"
       "            _pend = []")]),
    ("el lote deja de anunciar lo que queda",
     test_el_lote_anuncia_lo_que_deja_pendiente,
     [("""                yield f'data: CIERRE_PENDIENTE:{",".join(_pend)}\\n\\n'""",
       "                pass")]),
    ("el frontend deja de llamar al cierre tras las fotos",
     test_el_frontend_llama_al_cierre,
     [("      var _avCierre = await _correrCierre(_cierreFotos, addLine);",
       "      var _avCierre = false;")]),
]


def main():
    print(f"\n{'SABOTAJE — ' if SABOTAJE else ''}el paso de cierre del pipeline")
    print("=" * 66)
    if SABOTAJE:
        malos = 0
        for nombre, prueba, cambios in SABOTAJES:
            copia = _copia(cambios, nombre)
            try:
                try:
                    prueba(copia)
                except AssertionError as e:
                    print(f"  ✔ {nombre}:\n      {str(e)[:130]}")
                    continue
                print(f"  ✗ {nombre}: el invariante NO ha fallado.")
                print("    Un test que no puede fallar no protege de nada.")
                malos += 1
            finally:
                if os.path.exists(copia):
                    os.remove(copia)
        print("=" * 66)
        return 1 if malos else 0

    fallos = []
    for p in PRUEBAS:
        try:
            p()
        except AssertionError as e:
            fallos.append(p.__name__)
            print(f"  ✗ {p.__name__}\n      {e}")
    print("=" * 66)
    if fallos:
        print(f"  {len(fallos)} FALLO(S)")
        return 1
    print(f"  {len(PRUEBAS)} pruebas OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
