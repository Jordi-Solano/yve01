"""Los dos caminos del lote leen TODAS las banderas del enrutador.

EL BUG: `_enrutar_tipo_doc` devuelve cinco banderas (`has_ar`, `has_ap`, `ap_n`,
`albaran`, `orden_compra`). El camino de HOJAS DE CALCULO leia cuatro; el camino
del CLASIFICADOR IA —por donde entran TODOS los PDF y TODAS las fotos, o sea el
mas usado— leia **solo `has_ar`**.

Consecuencia medida: un albaran en PDF se guardaba perfectamente y despues el
cruce NO se relanzaba, porque `has_albaran` seguia en False. Una entrega nueva no
volvia a evaluar la factura que ayer no cuadraba, que es justo lo que la fase
3b·2 existe para conseguir. Y lo desconcertante: el MISMO albaran funcionaba
subido en Excel y no en PDF.

DOS COMPROBACIONES, y hacen falta las dos:

  1. de punta a punta: se conduce `/api/procesar_batch_stream` de verdad con el
     clasificador mockeado y se comprueba que un albaran relanza el cruce. Es el
     comportamiento; el mock existe porque en el sandbox no hay clave de la IA y
     porque lo que se prueba NO es que la IA acierte el tipo (eso solo se
     comprueba en produccion) sino que el lote reaccione cuando lo dice.

  2. el INVARIANTE, con AST: cada llamada a `_enrutar_tipo_doc` dentro del lote
     tiene que leer las cuatro banderas que se usan. Esta es la que habria
     cazado el bug el dia que se escribio, y la que lo caza si mañana se añade
     un tercer camino de entrada y se vuelve a olvidar una.

`--sabotaje` quita una lectura de bandera de una COPIA del fichero y comprueba
que el invariante grita. La copia se analiza con AST, sin importarla: dashboard
son 730 KB y registra rutas de Flask, asi que importar un duplicado seria peor
que el bug.
"""
import ast
import json
import os
import shutil
import sys
from datetime import date

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

HOY = date.today().strftime("%Y%m%d")
TENANT = "test-banderas"
H_A = "HBANDAA"
SABOTAJE = "--sabotaje" in sys.argv

# Las banderas que el enrutador PUEDE poner y que el lote USA. `orden_compra`
# se queda fuera a proposito: hoy no la lee nadie en ningun camino, asi que
# exigirla seria pedir codigo muerto.
BANDERAS = ("has_ar", "has_ap", "ap_n", "albaran")


# ── 2 · el invariante (AST) ───────────────────────────────────────────────

def _bloques_con_enrutador(ruta):
    """Cada BLOQUE de sentencias que llama a `_enrutar_tipo_doc`, con las
    banderas que lee ese mismo bloque.

    A nivel de BLOQUE y no de funcion, y esto lo cazo el sabotaje: las dos
    llamadas del lote —la de hojas de calculo y la del clasificador IA— viven
    en la MISMA funcion, asi que comprobando por funcion la que si leia
    `albaran` tapaba a la que no. Un invariante al nivel equivocado deja pasar
    justo el bug que existe.

    Se acota a las funciones del lote a proposito: `api_scan_documento` tambien
    llama al enrutador, pero ese endpoint no tiene bloque de cierre —no lanza el
    cruce ni el asignador— asi que leer las banderas alli no serviria de nada.
    Esa mitad se arregla con el paso de cierre, no con banderas.
    """
    arbol = ast.parse(open(ruta, encoding="utf-8").read())
    fuera = []

    def _lee(nodos):
        leidas = set()
        for n in nodos:
            for sub in ast.walk(n):
                if (isinstance(sub, ast.Call) and isinstance(sub.func, ast.Attribute)
                        and sub.func.attr == "get"
                        and isinstance(sub.func.value, ast.Name)
                        and sub.func.value.id == "_flags"
                        and sub.args and isinstance(sub.args[0], ast.Constant)):
                    leidas.add(sub.args[0].value)
        return leidas

    def _llama(n):
        return any(isinstance(s, ast.Call) and isinstance(s.func, ast.Name)
                   and s.func.id == "_enrutar_tipo_doc" for s in ast.walk(n))

    for fn in ast.walk(arbol):
        if not isinstance(fn, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        if "batch" not in fn.name.lower():
            continue
        for nodo in ast.walk(fn):
            for campo in ("body", "orelse", "finalbody"):
                cuerpo = getattr(nodo, campo, None)
                if not isinstance(cuerpo, list):
                    continue
                # el bloque es el que contiene la ASIGNACION
                # `_msg, _marca, _flags = _enrutar_tipo_doc(...)` como sentencia
                # DIRECTA, no uno de mas arriba que la lleve dentro de un if.
                directos = [s for s in cuerpo
                            if isinstance(s, ast.Assign) and _llama(s.value)]
                if directos:
                    fuera.append((fn.name, cuerpo[0].lineno, _lee(cuerpo)))
    return fuera


def test_el_lote_lee_todas_las_banderas(ruta=None):
    ruta = ruta or os.path.join(BASE, "dashboard.py")
    bloques = _bloques_con_enrutador(ruta)
    assert len(bloques) >= 2, (
        f"esperaba al menos 2 bloques del lote llamando al enrutador (el de "
        f"hojas de calculo y el del clasificador IA) y he encontrado "
        f"{len(bloques)}: el test ha dejado de saber donde mirar")
    for fn, linea, leidas in bloques:
        faltan = [b for b in BANDERAS if b not in leidas]
        assert not faltan, (
            f"el bloque de `{fn}` que empieza en la linea {linea} llama a "
            f"`_enrutar_tipo_doc` y NO lee {faltan}. El enrutador puede devolver "
            "esas banderas y el lote las necesita: sin `albaran` un albaran en "
            "PDF no relanza el cruce, y sin `has_ap` no corre el asignador. Es "
            "el bug de siempre — un camino de entrada que se deja una bandera "
            "por el suelo.")
    print(f"  ✔ los {len(bloques)} bloques del lote leen las {len(BANDERAS)} banderas (AST)")


# ── 1 · de punta a punta ──────────────────────────────────────────────────

def test_un_albaran_en_pdf_relanza_el_cruce():
    import pandas as pd
    import lector_facturas_ap as LFA
    os.environ["YVE_TENANT"] = TENANT
    os.environ["YVE_HOTEL"] = H_A
    import dashboard as D
    from tenant_dirs import datos_dir, entrada_dir, procesadas_dir, reportes_dir

    for d in (entrada_dir(), procesadas_dir(), reportes_dir()):
        if os.path.isdir(d):
            shutil.rmtree(d)
    for d in (datos_dir(), entrada_dir(), procesadas_dir(), reportes_dir()):
        os.makedirs(d, exist_ok=True)
    json.dump([{"id": H_A, "nombre": "Hotel Banderas", "activo": True}],
              open(os.path.join(datos_dir(), "hoteles.json"), "w"))
    for f in ("proveedores.xlsx", "plan_cuentas.xlsx"):
        src = os.path.join(BASE, "datos-referencia", f)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(datos_dir(), f))

    # una factura de antes, para que el albaran nuevo tenga con que cruzar
    pd.DataFrame([{
        "archivo": "previa.pdf", "numero_factura": "FP-9001", "fecha": "22/07/2026",
        "nombre_proveedor": "PESCADOS RIAS S.L.", "base_imponible": 1000.0,
        "porcentaje_iva": 21.0, "cuota_iva": 210.0, "total_factura": 1210.0,
        "descripcion_concepto": "", "hotel_id": H_A,
    }]).to_excel(os.path.join(procesadas_dir(), f"facturas_ap_{HOY}.xlsx"),
                 sheet_name="Facturas", index=False)

    open(os.path.join(entrada_dir(), "alb.pdf"), "wb").write(b"%PDF-1.4\n% mock\n")

    def _mock(fpath, proveedores=None, *a, **k):
        return {"tipo_documento": "ALBARAN", "numero_albaran": "ALB-9001",
                "nombre_proveedor": "PESCADOS RIAS S.L.",
                "fecha_entrega": "18/07/2026", "total_albaran": 1000.0,
                "lineas": [{"descripcion": "Merluza", "cantidad": 50,
                            "precio_unitario": 20.0, "importe": 1000.0}]}
    LFA.procesar_factura_ap = _mock

    D.app.config["TESTING"] = True
    c = D.app.test_client()
    c.post("/api/login", json={"username": "admin", "password": "admin123"})
    with c.session_transaction() as s:
        s["tenant_id"] = TENANT
        s["hotel_activo"] = H_A
    r = c.get('/api/procesar_batch_stream?archivos=' + json.dumps(["alb.pdf"]))
    log = [l[5:].strip() for l in r.get_data(as_text=True).splitlines() if l.startswith("data:")]

    assert any("Albar" in l for l in log), f"el albaran no se ha guardado: {log}"
    assert any("Cruzando facturas con albaranes" in l for l in log), (
        "el albaran se ha guardado pero el cruce NO se ha relanzado. Es el bug: "
        f"la bandera `albaran` se pierde en el camino del PDF. Log: {log}")
    print("  ✔ un albaran en PDF relanza el cruce (endpoint SSE de verdad)")

    for d in (entrada_dir(), procesadas_dir(), reportes_dir()):
        if os.path.isdir(d):
            shutil.rmtree(d)


PRUEBAS = [test_el_lote_lee_todas_las_banderas, test_un_albaran_en_pdf_relanza_el_cruce]


def _sabotear():
    """Copia de dashboard.py sin la lectura de `albaran`, para el AST."""
    src = open(os.path.join(BASE, "dashboard.py"), encoding="utf-8").read()
    viejo = """                                if _flags.get('albaran'):
                                    has_albaran = True"""
    assert src.count(viejo) == 1, (
        "el sabotaje ya no encuentra la lectura de `albaran` en el camino del "
        "PDF: el test ha dejado de saber que romper y hay que ponerlo al dia")
    dst = os.path.join(BASE, "dashboard_SABOTAJE.py")
    open(dst, "w", encoding="utf-8").write(src.replace(viejo, "", 1))
    return dst


def main():
    print(f"\n{'SABOTAJE — ' if SABOTAJE else ''}el lote lee las banderas del enrutador")
    print("=" * 62)
    if SABOTAJE:
        copia = _sabotear()
        try:
            try:
                test_el_lote_lee_todas_las_banderas(copia)
            except AssertionError as e:
                print(f"  ✔ con la bandera quitada, el invariante grita:\n      {str(e)[:150]}")
                return 0
            print("  ✗ con la bandera quitada el invariante NO ha fallado.")
            print("    Un test que no puede fallar no protege de nada.")
            return 1
        finally:
            if os.path.exists(copia):
                os.remove(copia)

    fallos = []
    for p in PRUEBAS:
        try:
            p()
        except AssertionError as e:
            fallos.append(p.__name__)
            print(f"  ✗ {p.__name__}\n      {e}")
    print("=" * 62)
    if fallos:
        print(f"  {len(fallos)} FALLO(S)")
        return 1
    print(f"  {len(PRUEBAS)} pruebas OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
