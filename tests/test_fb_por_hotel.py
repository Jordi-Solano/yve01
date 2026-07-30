"""El stock, las mermas y las ventas de F&B son de CADA hotel, por las TRES puertas.

DOS BUGS que compartian una linea de codigo:

  · La CAPA 1 (la que reconoce el fichero por su nombre) y la FOTO guardaban
    inventario, mermas y ventas SIN `hotel_id`. Los paneles de F&B leen con
    `_xlsx_hotel`, que filtra por el hotel de la sesion y FALLA CERRADO: la
    fila sin hotel no es de nadie y no sale en ningun panel. Medido: el lote
    cantaba "14 items integrados" y el panel de inventario mostraba 0. Un cero
    en silencio es peor que un error, porque nadie lo va a buscar.

  · Y el inventario se deduplicaba por `ingrediente` A SECAS. El "Tomate" del
    hotel B borraba el del hotel A, con su stock y su precio de compra. Subir
    10 + 10 ingredientes con 4 en comun dejaba 16 filas en vez de 20.

La puerta del clasificador de IA era la unica que lo hacia bien, asi que el
mismo albaran de inventario funcionaba por un camino y no por el otro — el mismo
sintoma desconcertante que el albaran en Excel y en PDF.

Se arregla con UN sitio, `_guardar_fb_del_hotel`, por el que pasan las tres
puertas. Esta prueba mide las nueve combinaciones (3 tablas x 3 puertas) contando
lo que VE el panel de cada hotel, que es lo unico que importa.

`--sabotaje` le quita el sello del hotel y le devuelve la clave sin hotel, en
una copia, y comprueba que las dos se notan.
"""
import io
import json
import os
import shutil
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

SABOTAJE = "--sabotaje" in sys.argv
TENANT = "test-fb-hotel"
CA, PM = "HTFBCA", "HTFBPM"
os.environ["YVE_TENANT"] = TENANT
os.environ["YVE_HOTEL"] = CA
os.environ.setdefault("ANTHROPIC_API_KEY", "test-no-se-usa")

# 4 ingredientes en comun y 6 propios de cada hotel. Los comunes llevan PRECIO
# DISTINTO en cada hotel a proposito: ese precio es lo que la clave sin hotel
# destruia, y es lo que hace que el mismo plato tenga food cost distinto en dos
# hoteles — que no es un fallo, es el dato interesante.
COMUNES = [("Tomate", 2.10), ("Aceite de oliva", 8.90), ("Harina", 1.20), ("Sal", 0.60)]
PROPIOS = {CA: ["Gambas", "Merluza", "Arroz bomba", "Azafran", "Pimiento", "Cebolla"],
           PM: ["Cafe molido", "Pan de payes", "Vino tinto", "Jamon", "Queso", "Aceitunas"]}
FACTOR = {CA: 1.0, PM: 1.15}
ESPERADO = {"inventario.xlsx": 10, "mermas.xlsx": 4, "ventas_fb_diarias.xlsx": 2}


def _tabla(tipo, hotel):
    import pandas as pd
    if tipo == "inventario.xlsx":
        filas = [(n, round(p * FACTOR[hotel], 2)) for n, p in COMUNES]
        filas += [(n, 5.0) for n in PROPIOS[hotel]]
        return pd.DataFrame([{"producto": n, "familia": "Varios", "precio": p,
                              "stock_inicial": 40, "stock_actual": 35,
                              "unidad": "kg", "proveedor": "Prov"} for n, p in filas])
    if tipo == "mermas.xlsx":
        nombres = [n for n, _ in COMUNES[:2]] + PROPIOS[hotel][:2]
        return pd.DataFrame([{"fecha": "31/07/2026", "producto": n,
                              "cantidad": 1.5, "motivo": "Caducado"} for n in nombres])
    platos = ([("PAELLA", 310, 7595.0), ("Merluza plancha", 265, 5830.0)] if hotel == CA
              else [("Tostada", 540, 2268.0), ("Cafe con leche", 810, 2106.0)])
    return pd.DataFrame([{"fecha": "31/07/2026", "plato": n, "categoria": "X",
                          "cantidad": u, "total": t} for n, u, t in platos])


def _limpiar():
    from tenant_dirs import datos_dir, entrada_dir
    for d in (datos_dir(), entrada_dir()):
        if os.path.isdir(d):
            shutil.rmtree(d)
        os.makedirs(d, exist_ok=True)
    json.dump([{"id": CA, "nombre": "Hotel Costa Azul", "activo": True},
               {"id": PM, "nombre": "Hotel Plaza Mayor", "activo": True}],
              open(os.path.join(datos_dir(), "hoteles.json"), "w"))


def _cliente(hotel):
    import dashboard as D
    D.app.config["TESTING"] = True
    D.app.config["WTF_CSRF_ENABLED"] = False
    c = D.app.test_client()
    c.post("/api/login", json={"username": "admin", "password": "admin123"})
    with c.session_transaction() as s:
        s["tenant_id"] = TENANT
        s["hotel_activo"] = hotel
    os.environ["YVE_HOTEL"] = hotel
    return c


def _ve_el_panel(hotel, fichero):
    """Lo que sale por `_xlsx_hotel`, que es lo que pinta el panel."""
    import dashboard as D
    import tab_fb_dashboard as FB
    os.environ["YVE_HOTEL"] = hotel
    with D.app.test_request_context("/"):
        from flask import session
        session["tenant_id"] = TENANT
        session["hotel_activo"] = hotel
        FB._invalidate()
        try:
            return len(FB._xlsx_hotel(fichero))
        except Exception as e:
            return f"ERR {str(e)[:40]}"


# ── las tres puertas ──────────────────────────────────────────────────────

def _puerta_capa1(tipo, hotel):
    nombre = {"inventario.xlsx": "inventario_x.xlsx", "mermas.xlsx": "mermas_x.xlsx",
              "ventas_fb_diarias.xlsx": "ventas_tpv_x.xlsx"}[tipo]
    from tenant_dirs import entrada_dir
    _tabla(tipo, hotel).to_excel(os.path.join(entrada_dir(), nombre), index=False)
    c = _cliente(hotel)
    # OJO: hay que CONSUMIR el cuerpo. Es un SSE con `stream_with_context`, asi
    # que el generador no corre hasta que alguien lo lee — sin esta linea el
    # endpoint no hace nada y el test cree que el arreglo no funciona.
    c.get("/api/procesar_batch_stream?archivos=" + json.dumps([nombre])).get_data()


def _reg_ia(tipo, hotel):
    import dashboard as D
    mapa = {"inventario.xlsx": D._INV_COL_MAP, "mermas.xlsx": D._MER_COL_MAP,
            "ventas_fb_diarias.xlsx": D._VEN_COL_MAP}[tipo]
    items = D._normalize_cols(_tabla(tipo, hotel), mapa).to_dict("records")
    tdoc = {"inventario.xlsx": "INVENTARIO", "mermas.xlsx": "MERMAS",
            "ventas_fb_diarias.xlsx": "VENTAS_POS"}[tipo]
    reg = {"tipo_documento": tdoc}
    reg["platos" if tdoc == "VENTAS_POS" else "items"] = items
    if tdoc == "VENTAS_POS":
        reg["total_ventas"] = sum(i.get("total_venta", 0) for i in items)
        reg["fecha"] = "2026-07-31"
    return reg


def _puerta_ia(tipo, hotel):
    import lector_facturas_ap as LFA
    from tenant_dirs import entrada_dir
    reg = _reg_ia(tipo, hotel)
    f = "neutro.pdf"
    open(os.path.join(entrada_dir(), f), "wb").write(b"%PDF-1.4\n% mock\n")
    LFA.procesar_factura_ap = lambda fpath, proveedores=None, *a, **k: reg
    c = _cliente(hotel)
    c.get("/api/procesar_batch_stream?archivos=" + json.dumps([f])).get_data()


def _puerta_foto(tipo, hotel):
    import anthropic
    reg = _reg_ia(tipo, hotel)

    class _M:
        def create(self, *a, **k):
            class _T: text = json.dumps(reg)
            class _R: content = [_T()]
            return _R()

    class _C:
        def __init__(self, *a, **k): self.messages = _M()

    anthropic.Anthropic = _C
    c = _cliente(hotel)
    c.post("/api/scan_documento", content_type="multipart/form-data",
           data={"image": (io.BytesIO(b"\xff\xd8\xff foto"), "doc.jpg")})


PUERTAS = {"capa 1": _puerta_capa1, "IA": _puerta_ia, "foto": _puerta_foto}


def test_las_tres_puertas_guardan_el_hotel():
    fallos = []
    for tipo, esperado in ESPERADO.items():
        for etiqueta, puerta in PUERTAS.items():
            _limpiar()
            for hotel in (CA, PM):
                puerta(tipo, hotel)
            ve = {h: _ve_el_panel(h, tipo) for h in (CA, PM)}
            if ve[CA] != esperado or ve[PM] != esperado:
                fallos.append(f"{tipo} por «{etiqueta}»: el panel ve {ve[CA]} en "
                              f"Costa Azul y {ve[PM]} en Plaza Mayor, y tendria que "
                              f"ver {esperado} en cada uno")
    assert not fallos, (
        "hay puertas que no guardan el hotel:\n      " + "\n      ".join(fallos) +
        "\n      Un 0 significa que la fila se guardo sin `hotel_id`: los paneles "
        "leen con `_xlsx_hotel`, que falla cerrado, asi que el documento se "
        "guarda, el pipeline canta ✓ y la pantalla se queda vacia sin decir nada.")
    print(f"  ✔ las {len(PUERTAS)} puertas x {len(ESPERADO)} tablas: cada panel ve lo suyo")


def test_el_inventario_no_pierde_los_ingredientes_compartidos():
    """La clave es (ingrediente, hotel). Con `ingrediente` a secas, el stock y el
    precio de un hotel borran los del otro."""
    import pandas as pd
    from tenant_dirs import datos_dir
    _limpiar()
    for hotel in (CA, PM):
        _puerta_capa1("inventario.xlsx", hotel)
    df = pd.read_excel(os.path.join(datos_dir(), "inventario.xlsx"))
    assert len(df) == 20, (
        f"el inventario tiene {len(df)} filas y tendria que tener 20 (10 de cada "
        "hotel, con 4 ingredientes en comun). Si hay 16, la clave del "
        "deduplicado se ha quedado sin el hotel y el stock de uno ha borrado el "
        "del otro.")
    for nombre, precio_ca in COMUNES:
        filas = df[df["ingrediente"] == nombre]
        assert len(filas) == 2, (
            f"«{nombre}» aparece {len(filas)} vez/veces y tendria que aparecer 2, "
            "una por hotel: es un ingrediente que los dos hoteles compran.")
        precios = dict(zip(filas["hotel_id"], filas["coste_unitario"]))
        assert round(precios[CA], 2) == precio_ca, (
            f"«{nombre}» en Costa Azul cuesta {precios[CA]} y deberia costar "
            f"{precio_ca}: le ha llegado el precio del otro hotel.")
        assert round(precios[PM], 2) == round(precio_ca * FACTOR[PM], 2), (
            f"«{nombre}» en Plaza Mayor cuesta {precios[PM]} y deberia costar "
            f"{round(precio_ca * FACTOR[PM], 2)}.")
    print("  ✔ los 4 ingredientes compartidos existen 2 veces, cada uno con su precio")


def test_solo_hay_un_sitio_que_guarda_f_and_b(ruta=None):
    """El invariante de fondo. El bug existia porque la misma regla —estampar el
    hotel y deduplicar— vivia en tres sitios y solo uno la cumplia. Mientras
    haya un solo sitio, no puede volver a desalinearse."""
    import ast
    ruta = ruta or os.path.join(BASE, "dashboard.py")
    arbol = ast.parse(open(ruta, encoding="utf-8").read())

    escriben = {}

    def rec(nodo, pila):
        for h in ast.iter_child_nodes(nodo):
            nueva = pila + [h.name] if isinstance(h, (ast.FunctionDef, ast.AsyncFunctionDef)) else pila
            # `<df>.to_excel(<ruta>)` donde la ruta se compone con uno de los
            # tres ficheros de F&B
            if (isinstance(h, ast.Call) and isinstance(h.func, ast.Attribute)
                    and h.func.attr == "to_excel"):
                for s in ast.walk(h):
                    if isinstance(s, ast.Name) and "path" in s.id.lower():
                        escriben.setdefault(nueva[-1] if nueva else "(modulo)", set()).add(h.lineno)
            rec(h, nueva)

    rec(arbol, [])
    # los ficheros de F&B por nombre, para saber quien los menciona
    src = open(ruta, encoding="utf-8").read()
    lineas = src.splitlines()
    sospechosos = []
    for i, l in enumerate(lineas, 1):
        if ".to_excel(" not in l:
            continue
        # ¿hay un fichero de F&B mencionado en las 12 lineas anteriores?
        ventana = "\n".join(lineas[max(0, i - 13):i])
        for fich in ESPERADO:
            if f"'{fich}'" in ventana or f'"{fich}"' in ventana:
                sospechosos.append((i, fich, l.strip()[:60]))

    fuera = [s for s in sospechosos if "def _guardar_fb_del_hotel" not in
             "\n".join(lineas[max(0, s[0] - 40):s[0]])]
    assert not fuera, (
        f"hay {len(fuera)} sitio(s) que escriben un fichero de F&B sin pasar por "
        f"`_guardar_fb_del_hotel`: {[(l, f) for l, f, _ in fuera]}. Ahi es donde "
        "estaba el bug: tres puertas escribiendo los mismos ficheros y solo una "
        "estampando el hotel. La regla vive en un sitio o se vuelve a "
        "desalinear.")
    print("  ✔ solo `_guardar_fb_del_hotel` escribe los 3 ficheros de F&B (AST + texto)")


PRUEBAS = [test_las_tres_puertas_guardan_el_hotel,
           test_el_inventario_no_pierde_los_ingredientes_compartidos,
           test_solo_hay_un_sitio_que_guarda_f_and_b]


def main():
    print(f"\n{'SABOTAJE — ' if SABOTAJE else ''}F&B: el stock y las ventas son de cada hotel")
    print("=" * 70)

    if SABOTAJE:
        import dashboard as D
        malos = 0
        bueno = D._guardar_fb_del_hotel
        clave_buena = dict(D._CLAVE_FB)

        # (a) sin sellar el hotel: es la mitad de la capa 1 y de la foto
        def _sin_sello(df, fichero):
            import pandas as _pd
            ruta = os.path.join(D._ddir(), fichero)
            if os.path.exists(ruta):
                viejo = _pd.read_excel(ruta)
                if not viejo.empty:
                    df = _pd.concat([viejo, df], ignore_index=True)
            df.to_excel(ruta, index=False)
            try:
                from tab_fb_dashboard import _invalidate
                _invalidate()
            except Exception:
                pass
            return df, len(df)

        D._guardar_fb_del_hotel = _sin_sello
        try:
            test_las_tres_puertas_guardan_el_hotel()
            print("  ✗ sin estampar el hotel, el test NO ha fallado.")
            malos += 1
        except AssertionError as e:
            print(f"  ✔ se deja de estampar el hotel:\n      {str(e)[:150]}")
        finally:
            D._guardar_fb_del_hotel = bueno

        # (b) la clave del inventario, sin hotel
        def _clave_sin_hotel(df, fichero):
            import pandas as _pd
            df = df.copy()
            df['hotel_id'] = D.censo_hoteles.para_guardar()
            ruta = os.path.join(D._ddir(), fichero)
            if os.path.exists(ruta):
                viejo = _pd.read_excel(ruta)
                if not viejo.empty:
                    df = _pd.concat([viejo, df], ignore_index=True)
            if fichero == 'inventario.xlsx' and 'ingrediente' in df.columns:
                df = df.drop_duplicates(subset=['ingrediente'], keep='last')
            df.to_excel(ruta, index=False)
            try:
                from tab_fb_dashboard import _invalidate
                _invalidate()
            except Exception:
                pass
            return df, len(df)

        D._guardar_fb_del_hotel = _clave_sin_hotel
        try:
            test_el_inventario_no_pierde_los_ingredientes_compartidos()
            print("  ✗ con la clave sin hotel, el test NO ha fallado.")
            malos += 1
        except AssertionError as e:
            print(f"  ✔ la clave del inventario pierde el hotel:\n      {str(e)[:150]}")
        finally:
            D._guardar_fb_del_hotel = bueno
            D._CLAVE_FB.clear()
            D._CLAVE_FB.update(clave_buena)

        # (c) una cuarta puerta que escribe por su cuenta
        src = open(os.path.join(BASE, "dashboard.py"), encoding="utf-8").read()
        ancla = "@app.route('/demo')"
        assert src.count(ancla) == 1, "el sabotaje no encuentra el ancla"
        copia = os.path.join(BASE, "dashboard_SABOTAJE.py")
        open(copia, "w", encoding="utf-8").write(src.replace(ancla, (
            "def _atajo_nuevo(df):\n"
            "    inv_path = os.path.join(_ddir(), 'inventario.xlsx')\n"
            "    df.to_excel(inv_path, index=False)\n\n\n" + ancla), 1))
        try:
            try:
                test_solo_hay_un_sitio_que_guarda_f_and_b(copia)
                print("  ✗ con una cuarta puerta suelta, el invariante NO ha fallado.")
                malos += 1
            except AssertionError as e:
                print(f"  ✔ aparece una puerta que escribe por su cuenta:\n      {str(e)[:150]}")
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
