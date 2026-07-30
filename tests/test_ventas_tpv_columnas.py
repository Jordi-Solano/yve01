"""Las cantidades del TPV llegan al food cost, no a `id_receta`.

EL BUG: `/fb/api/upload_ventas` tenia su PROPIA cadena de `elif` para reconocer
columnas, y comparaba por SUBCADENA:

    elif 'receta' in cl or 'recipe' in cl or 'id' in cl: col_map[col] = 'id_receta'

"id" esta dentro de `cantidad`, dentro de `unidad`, dentro de `unidades` y hasta
dentro de `unidades_vendidas` —el nombre canonico del propio esquema—. Y esa
rama va ANTES que la de cantidades, asi que gana. La columna de cantidades
acababa renombrada a `id_receta`, `unidades_vendidas` se rellenaba con 1 unos
renglones mas abajo, y el food cost salia por los suelos.

Medido con el endpoint de verdad, tres ficheros de 1.385 unidades:

  cabeceras            ANTES              DESPUES
  fecha·plato·qty      1.385 unidades     1.385 unidades   (este era el unico que iba)
  Fecha·Plato·Cantidad rechazado          1.385 unidades   ("faltan columnas: total_venta")
  unidades_vendidas    3 unidades         1.385 unidades   (el propio esquema, 1 por fila)
  CSV con ;            rechazado          1.385 unidades

Se arregla quitando la tercera cadena de `elif` y usando el MISMO mapa que los
otros dos caminos de ventas (el lote y la foto), que compara nombres COMPLETOS.

`--sabotaje` devuelve la comparacion por subcadena y comprueba que se nota.
"""
import io
import json
import os
import shutil
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

SABOTAJE = "--sabotaje" in sys.argv
TENANT = "test-ventas-col"
H = "HVEN01"

# tres platos, 1.385 unidades en total: el numero que tiene que sobrevivir a
# cualquier forma de escribir las cabeceras
FILAS = [("PAELLA DE MARISCO", "Arroces", 310, 7595.0),
         ("Merluza a la plancha", "Pescados", 265, 5830.0),
         ("Cafe con leche", "Bebidas", 810, 2106.0)]
UNIDADES = 1385


def _entorno():
    os.environ["YVE_TENANT"] = TENANT
    os.environ["YVE_HOTEL"] = H
    import dashboard as D
    from tenant_dirs import datos_dir
    if os.path.isdir(datos_dir()):
        shutil.rmtree(datos_dir())
    os.makedirs(datos_dir(), exist_ok=True)
    json.dump([{"id": H, "nombre": "Hotel Ventas", "activo": True}],
              open(os.path.join(datos_dir(), "hoteles.json"), "w"))
    D.app.config["TESTING"] = True
    D.app.config["WTF_CSRF_ENABLED"] = False
    c = D.app.test_client()
    c.post("/api/login", json={"username": "admin", "password": "admin123"})
    with c.session_transaction() as s:
        s["tenant_id"] = TENANT
        s["hotel_activo"] = H
    return D, c, datos_dir()


def _xlsx(cols):
    import pandas as pd
    buf = io.BytesIO()
    pd.DataFrame([("31/07/2026",) + f for f in FILAS], columns=cols).to_excel(buf, index=False)
    buf.seek(0)
    return "v.xlsx", buf


def _csv_es():
    """CSV como lo exporta un TPV de aqui: punto y coma y coma decimal."""
    lineas = ["Fecha;Plato;Categoria;Cantidad;Importe"]
    for n, cat, u, t in FILAS:
        imp = f"{t:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        lineas.append(f"31/07/2026;{n};{cat};{u};{imp}")
    return "v.csv", io.BytesIO(("\n".join(lineas) + "\n").encode("utf-8"))


CASOS = {
    "fecha·plato·categoria·qty·total_venta":
        lambda: _xlsx(["fecha", "plato", "categoria", "qty", "total_venta"]),
    "Fecha·Plato·Categoria·Cantidad·Importe":
        lambda: _xlsx(["Fecha", "Plato", "Categoria", "Cantidad", "Importe"]),
    "el nombre canonico del esquema":
        lambda: _xlsx(["fecha", "nombre_plato", "categoria", "unidades_vendidas", "total_venta"]),
    "CSV con punto y coma y coma decimal": _csv_es,
}


def test_las_cantidades_sobreviven_a_las_cabeceras():
    import pandas as pd
    D, c, ddir = _entorno()
    ruta = os.path.join(ddir, "ventas_fb_diarias.xlsx")
    fallos = []
    for etiqueta, hacer in CASOS.items():
        if os.path.exists(ruta):
            os.remove(ruta)
        nombre, buf = hacer()
        r = c.post("/fb/api/upload_ventas", content_type="multipart/form-data",
                   data={"file": (buf, nombre)})
        j = r.get_json() or {}
        if not j.get("ok"):
            fallos.append(f"«{etiqueta}»: rechazado — {str(j.get('error'))[:70]}")
            continue
        df = pd.read_excel(ruta)
        suma = float(pd.to_numeric(df["unidades_vendidas"], errors="coerce").fillna(0).sum())
        if suma != UNIDADES:
            fallos.append(f"«{etiqueta}»: {suma:g} unidades en vez de {UNIDADES} "
                          f"(id_receta={list(df.get('id_receta', []))[:3]})")
    assert not fallos, (
        "las cantidades no llegan enteras:\n      " + "\n      ".join(fallos) +
        "\n      Es el bug de la subcadena: 'id' esta dentro de `cantidad`, de "
        "`unidad` y de `unidades_vendidas`, asi que la columna de cantidades se "
        "iba a `id_receta` y `unidades_vendidas` se rellenaba con 1. El food "
        "cost sale por los suelos y parece una buena noticia.")
    print(f"  ✔ las {UNIDADES} unidades llegan con las {len(CASOS)} formas de escribir las cabeceras")


def test_el_endpoint_no_tiene_su_propio_mapa(ruta=None):
    """La causa de fondo era una TERCERA forma de reconocer las mismas columnas.
    Mientras el endpoint use el mapa comun, el bug no puede volver por aqui."""
    import ast
    ruta = ruta or os.path.join(BASE, "dashboard.py")
    arbol = ast.parse(open(ruta, encoding="utf-8").read())
    fn = next((n for n in ast.walk(arbol)
               if isinstance(n, ast.FunctionDef) and n.name == "api_upload_ventas_pos"), None)
    assert fn, "no encuentro `api_upload_ventas_pos`"

    usa_mapa = any(isinstance(n, ast.Call) and isinstance(n.func, ast.Name)
                   and n.func.id == "_normalize_cols" for n in ast.walk(fn))
    assert usa_mapa, (
        "`api_upload_ventas_pos` no llama a `_normalize_cols`: se ha vuelto a "
        "escribir su propia forma de reconocer columnas, que es de donde salio "
        "el bug. El mapa de columnas de ventas tiene que ser UNO.")

    # Y ninguna comparacion por subcadena DENTRO de un bucle sobre las columnas,
    # que es la forma exacta del bug. Se acota al bucle a proposito: un
    # `if 'x.xlsx' in _EXCEL_CACHE` tambien es un `in` con una cadena a la
    # izquierda y no tiene nada que ver — un invariante que grita por eso se
    # desactiva a la semana.
    bucles = [n for n in ast.walk(fn) if isinstance(n, ast.For)
              and any(isinstance(s, ast.Attribute) and s.attr == "columns"
                      for s in ast.walk(n.iter))]
    subcadenas = [c for b in bucles for c in ast.walk(b)
                  if isinstance(c, ast.Compare)
                  and any(isinstance(o, ast.In) for o in c.ops)
                  and isinstance(c.left, ast.Constant) and isinstance(c.left.value, str)]
    assert not subcadenas, (
        f"hay {len(subcadenas)} comparacion(es) por subcadena dentro de un bucle "
        f"sobre las columnas (linea(s) {[n.lineno for n in subcadenas]}). Es el "
        "bug tal cual: 'id' esta dentro de `unidades_vendidas`, asi que la "
        "columna de cantidades se iba a `id_receta`.")
    print("  ✔ el endpoint usa el mapa comun y no compara nombres por subcadena (AST)")


def test_dos_columnas_no_caen_en_el_mismo_nombre():
    """Un renombrado que deja nombres duplicados es peor que uno que no ocurre:
    a partir de ahi `df['col']` devuelve un DataFrame y todo lo de despues miente."""
    import pandas as pd
    import dashboard as D
    df = pd.DataFrame(columns=["cantidad", "unidades", "uds", "plato", "producto", "total", "importe"])
    salida = list(D._normalize_cols(df, D._VEN_COL_MAP).columns)
    dup = [c for c in set(salida) if salida.count(c) > 1]
    assert not dup, (
        f"el normalizador ha dejado columnas duplicadas: {dup} en {salida}. Tres "
        "columnas que suenan a cantidad no pueden acabar todas llamandose "
        "`unidades_vendidas`.")
    print(f"  ✔ con 3 columnas que suenan a cantidad no hay duplicados: {salida}")


# Los otros consumidores del normalizador: no se pide arreglarlos, estan aqui
# para que se vea si se mueven. `banco` con cabeceras en mayuscula SI cambia, y
# a mejor: antes se quedaba sin mapear y quien lo lee espera minusculas.
ESPERADO_OTROS = {
    "banco_mayus": (["Fecha", "Concepto", "Importe", "Saldo"],
                    ["fecha", "concepto", "importe", "saldo"]),
    "banco_alias": (["date", "descripcion", "cantidad", "balance"],
                    ["fecha", "concepto", "importe", "saldo"]),
    "inventario": (["producto", "familia", "precio", "stock_inicial", "stock_actual",
                    "unidad", "proveedor"],
                   ["ingrediente", "categoria", "coste_unitario", "stock_inicial_kg_l",
                    "stock_actual_kg_l", "unidad", "proveedor"]),
    "mermas": (["fecha", "producto", "cantidad", "motivo"],
               ["fecha", "ingrediente", "cantidad_merma", "causa"]),
    "ventas_ia": (["plato", "categoria", "cantidad", "total"],
                  ["nombre_plato", "categoria", "unidades_vendidas", "total_venta"]),
    "ota": (["numero_factura", "ota", "bruto", "porcentaje", "comision"],
            ["numero_factura", "nombre_ota", "importe_bruto", "porcentaje_comision",
             "importe_comision"]),
    "pactadas": (["ota", "nombre_hotel", "porcentaje_pactado", "mercado"],
                 ["OTA", "Hotel", "Porcentaje_Comision", "Mercado"]),
}


def test_los_otros_consumidores_no_se_mueven():
    import pandas as pd
    import dashboard as D
    mapas = {"banco_mayus": D._BANK_COL_MAP, "banco_alias": D._BANK_COL_MAP,
             "inventario": D._INV_COL_MAP, "mermas": D._MER_COL_MAP,
             "ventas_ia": D._VEN_COL_MAP, "ota": D._OTA_COL_MAP,
             "pactadas": D._PACT_COL_MAP}
    fallos = []
    for etiqueta, (entrada, esperado) in ESPERADO_OTROS.items():
        salida = list(D._normalize_cols(pd.DataFrame(columns=entrada), mapas[etiqueta]).columns)
        if salida != esperado:
            fallos.append(f"«{etiqueta}»: {salida} en vez de {esperado}")
    assert not fallos, (
        "el normalizador se ha movido en un camino que NO se pidio tocar:\n      "
        + "\n      ".join(fallos) +
        "\n      Es una pieza compartida por banco, inventario, mermas, ventas, "
        "comisiones y tarifas: aqui se para y se mira antes de seguir.")
    print(f"  ✔ los {len(ESPERADO_OTROS)} caminos que comparten el normalizador dan lo mismo")


PRUEBAS = [test_las_cantidades_sobreviven_a_las_cabeceras,
           test_el_endpoint_no_tiene_su_propio_mapa,
           test_dos_columnas_no_caen_en_el_mismo_nombre,
           test_los_otros_consumidores_no_se_mueven]


def _normalize_cols_ANTIGUO(df, expected_map):
    """El normalizador de antes, con la comparacion por subcadena que causaba el
    bug. Solo lo usa el sabotaje."""
    col_map = {}
    for col in df.columns:
        cl = str(col).lower().replace(' ', '_')
        if 'fecha' in cl or 'date' in cl:
            col_map[col] = 'fecha'
        elif 'receta' in cl or 'recipe' in cl or 'id' in cl:
            col_map[col] = 'id_receta'
        elif 'plato' in cl or 'nombre' in cl or 'name' in cl:
            col_map[col] = 'nombre_plato'
        elif 'categ' in cl:
            col_map[col] = 'categoria'
        elif 'unidad' in cl or 'qty' in cl or 'cantidad' in cl:
            col_map[col] = 'unidades_vendidas'
        elif 'precio' in cl or 'price' in cl or 'unit' in cl:
            col_map[col] = 'precio_unitario'
        elif 'total' in cl or 'venta' in cl or 'revenue' in cl:
            col_map[col] = 'total_venta'
    return df.rename(columns=col_map)


def main():
    print(f"\n{'SABOTAJE — ' if SABOTAJE else ''}las cantidades del TPV llegan al food cost")
    print("=" * 72)

    if SABOTAJE:
        malos = 0
        os.environ["YVE_TENANT"] = TENANT
        os.environ["YVE_HOTEL"] = H
        import dashboard as D
        bueno = D._normalize_cols
        D._normalize_cols = _normalize_cols_ANTIGUO   # la subcadena, de vuelta
        try:
            test_las_cantidades_sobreviven_a_las_cabeceras()
            print("  ✗ con la comparacion por subcadena de vuelta, el test NO ha fallado.")
            malos += 1
        except AssertionError as e:
            print(f"  ✔ vuelve la comparacion por subcadena:\n      {str(e)[:200]}")
        finally:
            D._normalize_cols = bueno

        # y el invariante de codigo, contra una copia con el mapa propio
        src = open(os.path.join(BASE, "dashboard.py"), encoding="utf-8").read()
        viejo = "        df_new = _normalize_cols(df_new, _VEN_COL_MAP)"
        assert src.count(viejo) == 1, "el sabotaje no encuentra la llamada al mapa comun"
        copia = os.path.join(BASE, "dashboard_SABOTAJE.py")
        open(copia, "w", encoding="utf-8").write(
            src.replace(viejo, "        df_new = df_new.rename(columns={'x': 'y'})", 1))
        try:
            try:
                test_el_endpoint_no_tiene_su_propio_mapa(copia)
                print("  ✗ sin la llamada al mapa comun, el invariante NO ha fallado.")
                malos += 1
            except AssertionError as e:
                print(f"  ✔ el endpoint se queda sin el mapa comun:\n      {str(e)[:150]}")
        finally:
            if os.path.exists(copia):
                os.remove(copia)

        print("=" * 72)
        return 1 if malos else 0

    fallos = []
    for p in PRUEBAS:
        try:
            p()
        except AssertionError as e:
            fallos.append(p.__name__)
            print(f"  ✗ {p.__name__}\n      {e}")
    print("=" * 72)
    if fallos:
        print(f"  {len(fallos)} FALLO(S)")
        return 1
    print(f"  {len(PRUEBAS)} pruebas OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
