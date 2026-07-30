"""Un contrato de OTA que cubre VARIAS OTAs no se colapsa en silencio.

EL BUG (nº 7): el schema del clasificador tiene la OTA arriba, en singular
—`{"ota":"nombre","tarifas":[{"nombre_hotel","porcentaje_pactado",...}]}`—. Un
acuerdo de distribución típico cubre Booking Y Expedia, así que la IA no tiene
dónde poner la OTA de cada tarifa y las junta en una cadena:
"Booking.com / Expedia".

El writer estampaba esa cadena en las 4 filas y deduplicaba por (OTA, hotel):
como la OTA era la misma para todas, 4 tarifas colapsaban a 2, se perdían las de
Booking, y ninguna factura cruzaba (todas OTA_DESCONOCIDA). Y encima cantaba un
✓ "Contrato OK". Un cobro que debería reclamarse (Expedia al 18% con 12%
pactado, 1.920 €) desaparecía.

DÓNDE ESTÁ EL ARREGLO Y DÓNDE NO:

El writer YA hace lo correcto SI cada tarifa trae su OTA (`_PACT_COL_MAP` la
recoge). La prueba B lo demuestra: con OTA por fila, las 4 tarifas sobreviven y
Expedia da sus 1.920 €. Lo único que falta para que eso pase de verdad es que la
IA ponga la OTA en cada fila — y eso es un cambio en el schema del clasificador,
que es intocable sin permiso. Este arreglo es solo la RED DE SEGURIDAD del
writer: cuando la OTA de arriba es una lista y las filas no traen la suya, NO se
inventa — se para y se avisa, en vez de guardar basura con un ✓.

  A · combinada, sin OTA por fila  -> SKIP con aviso, no guarda nada   (el bug)
  B · con OTA por fila             -> 4 tarifas, Expedia 1.920 €        (el arreglo real)
  C · una sola OTA                 -> como siempre                      (control)

`--sabotaje` devuelve el `fillna(ota)` de siempre y comprueba que A vuelve a
colapsar.
"""
import glob
import json
import os
import shutil
import subprocess
import sys
from datetime import date

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

SABOTAJE = "--sabotaje" in sys.argv
TENANT = "test-contrato-multi"
HOY = date.today().strftime("%Y%m%d")
CA, PM = "HTCMCA", "HTCMPM"

FACTURAS = [
    {"archivo": "ca_booking.pdf", "numero_factura": "BK-778001", "fecha": "31/07/2026",
     "nombre_ota": "Booking.com", "nombre_hotel": "Hotel Costa Azul",
     "importe_bruto": 86400.0, "porcentaje_comision": 15.0,
     "importe_comision": 12960.0, "hotel_id": CA},
    {"archivo": "ca_expedia.pdf", "numero_factura": "EXP-550412", "fecha": "31/07/2026",
     "nombre_ota": "Expedia", "nombre_hotel": "Hotel Costa Azul",
     "importe_bruto": 32000.0, "porcentaje_comision": 18.0,
     "importe_comision": 5760.0, "hotel_id": CA},
]

REG = {
    "A_combinada": {"tipo_documento": "CONTRATO_OTA", "ota": "Booking.com / Expedia",
                    "tarifas": [
                        {"nombre_hotel": "Hotel Costa Azul", "porcentaje_pactado": 15.0},
                        {"nombre_hotel": "Hotel Costa Azul", "porcentaje_pactado": 12.0},
                        {"nombre_hotel": "Hotel Plaza Mayor", "porcentaje_pactado": 14.0},
                        {"nombre_hotel": "Hotel Plaza Mayor", "porcentaje_pactado": 11.0}]},
    "B_por_fila": {"tipo_documento": "CONTRATO_OTA", "ota": "Booking.com / Expedia",
                   "tarifas": [
                       {"ota": "Booking.com", "nombre_hotel": "Hotel Costa Azul", "porcentaje_pactado": 15.0},
                       {"ota": "Expedia", "nombre_hotel": "Hotel Costa Azul", "porcentaje_pactado": 12.0},
                       {"ota": "Booking.com", "nombre_hotel": "Hotel Plaza Mayor", "porcentaje_pactado": 14.0},
                       {"ota": "Expedia", "nombre_hotel": "Hotel Plaza Mayor", "porcentaje_pactado": 11.0}]},
    "C_una_ota": {"tipo_documento": "CONTRATO_OTA", "ota": "Booking.com",
                  "tarifas": [
                      {"nombre_hotel": "Hotel Costa Azul", "porcentaje_pactado": 15.0},
                      {"nombre_hotel": "Hotel Plaza Mayor", "porcentaje_pactado": 14.0}]},
}


def _limpiar(D):
    from tenant_dirs import datos_dir, procesadas_dir, reportes_dir
    for d in (datos_dir(), procesadas_dir(), reportes_dir()):
        if os.path.isdir(d):
            shutil.rmtree(d)
        os.makedirs(d, exist_ok=True)
    json.dump([{"id": CA, "nombre": "Hotel Costa Azul", "activo": True},
               {"id": PM, "nombre": "Hotel Plaza Mayor", "activo": True}],
              open(os.path.join(datos_dir(), "hoteles.json"), "w"))
    import pandas as pd
    pd.DataFrame(columns=["OTA", "Hotel", "Porcentaje_Comision", "Mercado"]).to_excel(
        os.path.join(datos_dir(), "comisiones_pactadas.xlsx"), index=False)
    pd.DataFrame(FACTURAS).to_excel(
        os.path.join(procesadas_dir(), f"facturas_procesadas_{HOY}.xlsx"), index=False)


def _correr(forma):
    os.environ["YVE_TENANT"] = TENANT
    os.environ["YVE_HOTEL"] = CA
    import pandas as pd
    import dashboard as D
    _limpiar(D)
    from tenant_dirs import datos_dir, reportes_dir
    with D.app.test_request_context("/"):
        from flask import session
        session["tenant_id"] = TENANT
        session["hotel_activo"] = CA
        msg, marca, _flags = D._enrutar_tipo_doc(REG[forma], "contrato.pdf")
    tp = os.path.join(datos_dir(), "comisiones_pactadas.xlsx")
    tar = pd.read_excel(tp) if os.path.exists(tp) else pd.DataFrame()
    subprocess.run([sys.executable, "verificador_comisiones.py"], cwd=BASE,
                   capture_output=True, text=True,
                   env={**os.environ, "YVE_TENANT": TENANT, "YVE_HOTEL": CA})
    rep = os.path.join(reportes_dir(), f"verificacion_{HOY}.xlsx")
    ver = {}
    if os.path.exists(rep):
        vr = pd.read_excel(rep)
        for _, r in vr.iterrows():
            ver[str(r.get("nombre_ota"))] = {
                "estado": str(r.get("estado")),
                "discrep": r.get("discrepancia_euros")}
    return {"marca": marca, "msg": msg, "n_tarifas": len(tar), "ver": ver}


def test_multi_ota_sin_ota_por_fila_no_se_guarda_a_ciegas():
    r = _correr("A_combinada")
    assert r["marca"] == "SKIP", (
        f"un contrato multi-OTA sin OTA por fila sale con marca {r['marca']!r} y "
        "tendria que ser SKIP. Antes se guardaba: la cadena 'Booking.com / "
        "Expedia' se estampaba en las 4 filas, el dedup por (OTA, hotel) las "
        "colapsaba a 2, y cantaba un ✓ sobre tarifas que ninguna factura cruza.")
    assert r["n_tarifas"] == 0, (
        f"se han guardado {r['n_tarifas']} tarifas de un contrato que no se "
        "puede separar: es basura con OTA inventada. No guardar nada es lo "
        "honesto.")
    assert "varias OTAs" in r["msg"], (
        f"el aviso no explica el problema: «{r['msg']}». Tiene que decir que el "
        "contrato cubre varias OTAs y no se pueden separar las tarifas.")
    print("  ✔ contrato multi-OTA sin OTA por fila: SKIP con aviso, no guarda basura")


def test_con_ota_por_fila_las_cuatro_tarifas_cruzan():
    r = _correr("B_por_fila")
    assert r["marca"] == "CONTRATO_OTA_OK", f"esperaba OK y es {r['marca']!r}: {r['msg']}"
    assert r["n_tarifas"] == 4, (
        f"con OTA por fila tendria que guardar las 4 tarifas y ha guardado "
        f"{r['n_tarifas']}. Si son 2, el dedup ha vuelto a colapsar por OTA.")
    exp = r["ver"].get("Expedia", {})
    assert exp.get("estado") == "DISCREPANCIA" and float(exp.get("discrep") or 0) == 1920.0, (
        f"Expedia tendria que salir DISCREPANCIA de 1.920 € (32.000 × (18−12)/100) "
        f"y sale {exp}. Es la reclamacion que el bug hacia desaparecer.")
    book = r["ver"].get("Booking.com", {})
    assert book.get("estado") == "CORRECTO", \
        f"Booking (15% pactado, 15% facturado) tendria que ser CORRECTO y es {book}"
    print("  ✔ con OTA por fila: 4 tarifas, Booking CORRECTO y Expedia 1.920 € (verificador real)")


def test_una_sola_ota_sigue_yendo():
    r = _correr("C_una_ota")
    assert r["marca"] == "CONTRATO_OTA_OK", f"un contrato de una OTA tiene que ir: {r['msg']}"
    assert r["n_tarifas"] == 2, f"esperaba 2 tarifas y hay {r['n_tarifas']}"
    assert r["ver"].get("Booking.com", {}).get("estado") == "CORRECTO", \
        "Booking al 15/15 tiene que salir CORRECTO"
    print("  ✔ el contrato de una sola OTA sigue yendo igual")


def test_el_detector_de_lista_de_otas():
    import dashboard as D
    listas = ["Booking.com / Expedia", "Booking.com, Expedia", "Booking y Expedia",
              "Expedia + Hotels.com", "booking.com & expedia"]
    solas = ["Booking.com", "Expedia", "Hotels.com", "", None, "Agoda"]
    for v in listas:
        assert D._ota_es_lista(v), f"«{v}» es una lista de OTAs y no se ha detectado"
    for v in solas:
        assert not D._ota_es_lista(v), f"«{v}» es una sola OTA (o vacia) y se ha tomado por lista"
    print("  ✔ el detector separa 'lista de OTAs' de 'una OTA'")


PRUEBAS = [test_multi_ota_sin_ota_por_fila_no_se_guarda_a_ciegas,
           test_con_ota_por_fila_las_cuatro_tarifas_cruzan,
           test_una_sola_ota_sigue_yendo,
           test_el_detector_de_lista_de_otas]


def main():
    print(f"\n{'SABOTAJE — ' if SABOTAJE else ''}contrato de OTA que cubre varias OTAs")
    print("=" * 68)

    if SABOTAJE:
        # se devuelve el fillna(ota) de siempre: la OTA combinada vuelve a
        # estamparse en todas las filas y A colapsa
        import dashboard as D
        bueno = D._ota_es_lista
        D._ota_es_lista = lambda v: False   # como si nunca fuera una lista
        try:
            test_multi_ota_sin_ota_por_fila_no_se_guarda_a_ciegas()
            print("  ✗ con el detector desactivado, el contrato multi-OTA NO ha fallado.")
            return 1
        except AssertionError as e:
            print(f"  ✔ sin detectar la lista de OTAs, vuelve a guardar a ciegas:\n      {str(e)[:150]}")
            return 0
        finally:
            D._ota_es_lista = bueno

    fallos = []
    for p in PRUEBAS:
        try:
            p()
        except AssertionError as e:
            fallos.append(p.__name__)
            print(f"  ✗ {p.__name__}\n      {e}")
    print("=" * 68)
    if fallos:
        print(f"  {len(fallos)} FALLO(S)")
        return 1
    print(f"  {len(PRUEBAS)} pruebas OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
