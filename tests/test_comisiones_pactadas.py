"""Las comisiones OTA no se comparan contra porcentajes que nadie ha pactado.

EL BUG tenia dos caras y las dos llegaban al panel con el signo cambiado:

  · `datos-referencia/comisiones_pactadas.xlsx` venia del repo RELLENO, con
    porcentajes inventados y sin columna Hotel: Booking.com 15%, Expedia 18%,
    Hotels.com 18%, Despegar 20%. Al no llevar hotel, cada fila funcionaba como
    tarifa GENERICA para todos los hoteles de todos los tenants. Medido con el
    juego de integracion: una factura de Expedia que cobraba el 18% con un 12%
    pactado por contrato salia CORRECTO —porque la semilla decia 18— y los
    1.920 EUR cobrados de mas desaparecian. Eso es un cobro inflado aprobado
    como correcto, que es el peor fallo que puede tener este producto.

  · Y al revés: Plaza Mayor tenia pactado el 14% con Booking, la semilla decia
    15%, y la factura correcta salia como DISCREPANCIA de -182 EUR. El panel
    sumaba ese importe en VALOR ABSOLUTO al total, asi que 182 EUR que la OTA
    nos habia cobrado de MENOS aparecian como 182 EUR "a devolver".

Se arregla por los dos lados: la tabla de tarifas deja de traer porcentajes
inventados (sin tarifa, el veredicto es "no puedo comparar", no "correcto"), y
cobrar por debajo de lo pactado tiene su propio estado y no se suma a lo
reclamable.

`--sabotaje` rompe cada garantia por separado y comprueba que grita.
"""
import os
import re
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

SABOTAJE = "--sabotaje" in sys.argv
SEMILLA = os.path.join(BASE, "datos-referencia", "comisiones_pactadas.xlsx")


# ── 1 · la semilla no inventa condiciones comerciales ─────────────────────

def test_la_semilla_no_trae_porcentajes_inventados(ruta=None):
    import pandas as pd
    df = pd.read_excel(ruta or SEMILLA)

    con_pct = df[df.get("Porcentaje_Comision", pd.Series(dtype=float)).notna()] \
        if "Porcentaje_Comision" in df.columns else df.iloc[0:0]
    assert con_pct.empty, (
        f"la tabla de tarifas viene del repo con {len(con_pct)} porcentaje(s) "
        f"puestos: {con_pct.to_dict('records')[:3]}. Un porcentaje que nadie ha "
        "pactado NO es un dato de referencia, es una condicion comercial "
        "inventada — y como estas filas no llevan hotel, valen como tarifa "
        "generica para todos los hoteles de todos los tenants. Asi es como una "
        "factura de Expedia al 18% con un 12% pactado salia CORRECTO y los "
        "1.920 EUR de mas se perdian. Sin tarifa, el veredicto tiene que ser "
        "'no puedo comparar'.")

    assert "Hotel" in df.columns, (
        "la tabla de tarifas no trae columna Hotel. El verificador cruza por "
        "(OTA, hotel) porque un grupo puede tener condiciones distintas por "
        "establecimiento: sin esa columna, la primera tarifa de la OTA se le "
        "aplica a todos y se reclama de mas o de menos.")
    print(f"  ✔ la tabla de tarifas viene vacia, con columna Hotel ({list(df.columns)})")


# ── 2 y 3 · el signo decide el estado, y el total solo suma lo reclamable ──

def _tarifas():
    import pandas as pd
    df = pd.DataFrame([
        {"OTA": "Booking.com", "Hotel": "Hotel Plaza Mayor",
         "Porcentaje_Comision": 14, "Mercado": "Internacional"},
        {"OTA": "Expedia", "Hotel": "Hotel Costa Azul",
         "Porcentaje_Comision": 12, "Mercado": "Internacional"},
    ])
    df["OTA_norm"] = df["OTA"].str.lower()
    df["Hotel_norm"] = df["Hotel"].str.lower()
    return df


def _factura(ota, hotel, pct, bruto=18200.0):
    return {"archivo": "f.pdf", "numero_factura": "X-1", "fecha": "31/07/2026",
            "nombre_ota": ota, "nombre_hotel": hotel, "importe_bruto": bruto,
            "porcentaje_comision": pct, "hotel_id": "H1"}


def test_cobrar_por_debajo_no_es_una_reclamacion():
    import pandas as pd
    import verificador_comisiones as VC
    import dashboard as D

    tar = _tarifas()
    # 13% facturado sobre un 14% pactado: la OTA nos cobra 182 EUR de MENOS
    debajo = VC.verificar_factura(_factura("Booking.com", "Hotel Plaza Mayor", 13.0), tar)
    assert debajo["estado"] == "COBRO_POR_DEBAJO", (
        f"una comision por DEBAJO de lo pactado sale como {debajo['estado']!r}. "
        "Con DISCREPANCIA y un importe negativo, el panel lo presentaba como "
        "dinero a devolver: 182 EUR que nos han cobrado de menos aparecian como "
        "182 EUR que reclamar. No es una reclamacion, es lo contrario.")
    assert debajo["discrepancia_euros"] == -182.0, (
        f"el importe deberia ser -182,00 EUR (signo incluido, que es la "
        f"informacion) y es {debajo['discrepancia_euros']!r}")

    # 18% facturado sobre un 12% pactado: 1.920 EUR de MAS, reclamables
    arriba = VC.verificar_factura(
        _factura("Expedia", "Hotel Costa Azul", 18.0, bruto=32000.0), tar)
    assert arriba["estado"] == "DISCREPANCIA", (
        f"una comision por ENCIMA de lo pactado sale como {arriba['estado']!r}: "
        "eso es justo lo que hay que reclamar y tiene que llegar en rojo.")
    assert arriba["discrepancia_euros"] == 1920.0, \
        f"esperaba 1.920,00 EUR reclamables y hay {arriba['discrepancia_euros']!r}"

    # y el total que ve el panel: solo lo cobrado de MAS
    s = D.calcular_stats(pd.DataFrame([debajo, arriba]))
    assert s["importe_reclamable"] == 1920.0, (
        f"el total reclamable es {s['importe_reclamable']} y tendria que ser "
        "1920.0. Antes se sumaba en valor absoluto, asi que los -182 EUR "
        "entraban como +182 y el panel pedia reclamar dinero que nos habian "
        "cobrado de menos.")
    assert s["discrepancias"] == 1 and s["cobro_debajo"] == 1, (
        f"el resumen cuenta {s['discrepancias']} discrepancia(s) y "
        f"{s['cobro_debajo']} por debajo: tendria que ser 1 y 1, cada cosa en "
        "su linea. Una factura que no se cuenta en ninguna linea desaparece "
        "del resumen sin que nadie se entere.")
    print("  ✔ por debajo no se reclama, por encima si, y el total suma solo lo cobrado de mas")


def test_sin_tarifa_no_se_da_por_bueno():
    import verificador_comisiones as VC
    tar = _tarifas()
    # Booking SI esta en la tabla, pero solo con tarifa de Plaza Mayor. A Costa
    # Azul no se le puede aplicar el porcentaje pactado para OTRO hotel.
    otro = VC.verificar_factura(_factura("Booking.com", "Hotel Costa Azul", 15.0), tar)
    assert otro["estado"] == "SIN_TARIFA_HOTEL", (
        f"una factura de un hotel sin tarifa propia sale como {otro['estado']!r}. "
        "Aplicarle el porcentaje pactado para otro establecimiento es reclamar "
        "de mas o de menos con cara de dato.")
    # Y una OTA que no esta en la tabla no puede salir correcta jamas. Desde
    # sep 2026 (decision de Jordi) se distingue: OTA reconocida sin contrato =
    # SIN_TARIFA_PACTADA; sin nombre de OTA = OTA_DESCONOCIDA. Ninguna es buena.
    desc = VC.verificar_factura(_factura("Agoda", "Hotel Plaza Mayor", 22.0), tar)
    assert desc["estado"] == "SIN_TARIFA_PACTADA", \
        f"una OTA reconocida sin tarifa sale como {desc['estado']!r}, no puede darse por buena"
    sinn = VC.verificar_factura(_factura("NO_ENCONTRADO", "Hotel Plaza Mayor", 22.0), tar)
    assert sinn["estado"] == "OTA_DESCONOCIDA", \
        f"una factura sin nombre de OTA sale como {sinn['estado']!r}"
    print("  ✔ sin tarifa con la que comparar, el veredicto no es 'correcto'")


# ── 4 · los estados nuevos se VEN en la pantalla ──────────────────────────

def test_los_estados_nuevos_se_ven(ruta=None):
    """Un estado que el frontend no conoce se pinta con su nombre en crudo y no
    se cuenta en el resumen: la factura desaparece de la vista aunque el
    servidor la haya clasificado bien."""
    src = open(ruta or os.path.join(BASE, "dashboard.py"), encoding="utf-8").read()
    js = "\n".join(re.findall(r"<script[^>]*>(.*?)</script>", src, re.S))
    js = re.sub(r"/\*.*?\*/", "", js, flags=re.S)
    js = re.sub(r"^\s*//.*$", "", js, flags=re.M)

    badges = re.search(r"function bEstado\(e\)\s*\{(.*?)\n\}", js, re.S)
    assert badges, "no encuentro `bEstado` en el JS"
    faltan = [e for e in ("COBRO_POR_DEBAJO", "SIN_TARIFA_HOTEL")
              if e not in badges.group(1)]
    assert not faltan, (
        f"`bEstado` no conoce {faltan}: esas filas se pintan con el nombre del "
        "estado en crudo, tipo 'SIN_TARIFA_HOTEL', que es lo que ve el usuario.")

    faltan_res = [e for e in ("COBRO_POR_DEBAJO", "SIN_TARIFA_HOTEL")
                  if f"c.{e}" not in js]
    assert not faltan_res, (
        f"el resumen de la seccion no cuenta {faltan_res}. Una factura que no "
        "entra en ninguna linea del resumen desaparece: ni correcta, ni "
        "discrepancia, ni nada.")
    print("  ✔ los 2 estados nuevos tienen etiqueta y linea en el resumen (JS sin comentarios)")


PRUEBAS = [test_la_semilla_no_trae_porcentajes_inventados,
           test_cobrar_por_debajo_no_es_una_reclamacion,
           test_sin_tarifa_no_se_da_por_bueno,
           test_los_estados_nuevos_se_ven]


# ── sabotaje ──────────────────────────────────────────────────────────────

def _semilla_saboteada():
    import pandas as pd
    dst = os.path.join(BASE, "comisiones_SABOTAJE.xlsx")
    pd.DataFrame([{"OTA": "Expedia", "Porcentaje_Comision": 18,
                   "Mercado": "Internacional"}]).to_excel(dst, index=False)
    return dst


def _dashboard_saboteado(cambios, que):
    src = open(os.path.join(BASE, "dashboard.py"), encoding="utf-8").read()
    for viejo, nuevo in cambios:
        assert src.count(viejo) == 1, (
            f"el sabotaje de «{que}» ya no encuentra que romper "
            f"({src.count(viejo)} apariciones): hay que ponerlo al dia")
        src = src.replace(viejo, nuevo, 1)
    dst = os.path.join(BASE, "dashboard_SABOTAJE.py")
    open(dst, "w", encoding="utf-8").write(src)
    return dst


def main():
    print(f"\n{'SABOTAJE — ' if SABOTAJE else ''}las comisiones OTA se comparan "
          f"contra lo pactado, no contra una semilla")
    print("=" * 72)

    if SABOTAJE:
        malos = 0
        casos = [
            ("la semilla vuelve a traer porcentajes inventados",
             test_la_semilla_no_trae_porcentajes_inventados, _semilla_saboteada),
            ("el frontend deja de conocer COBRO_POR_DEBAJO",
             test_los_estados_nuevos_se_ven,
             lambda: _dashboard_saboteado(
                 [("    COBRO_POR_DEBAJO: ['b-unk', '↓ Cobrado por debajo'],\n", "")],
                 "badge de COBRO_POR_DEBAJO")),
        ]
        for nombre, prueba, hacer_copia in casos:
            copia = hacer_copia()
            try:
                try:
                    prueba(copia)
                except AssertionError as e:
                    print(f"  ✔ {nombre}:\n      {str(e)[:130]}")
                    continue
                print(f"  ✗ {nombre}: el invariante NO ha fallado.")
                malos += 1
            finally:
                if os.path.exists(copia):
                    os.remove(copia)

        # el signo: se sabotea el modulo en memoria, que es donde vive la regla
        import verificador_comisiones as VC
        original = VC.COBRO_DEBAJO
        VC.COBRO_DEBAJO = "DISCREPANCIA"      # como estaba antes del arreglo
        try:
            test_cobrar_por_debajo_no_es_una_reclamacion()
            print("  ✗ con el estado colapsado a DISCREPANCIA el invariante NO ha fallado.")
            malos += 1
        except AssertionError as e:
            print(f"  ✔ el estado vuelve a colapsar en DISCREPANCIA:\n      {str(e)[:130]}")
        finally:
            VC.COBRO_DEBAJO = original

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
