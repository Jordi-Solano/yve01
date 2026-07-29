"""Los paneles siguen contestando despues de las extracciones de la fase A.

La fase A saco el cuerpo de cinco funciones que usan TODOS los paneles:

    dashboard.cargar_datos            -> cargar_datos_ar_sin_filtrar + filtro
    dashboard.cargar_datos_ap         -> cargar_datos_ap_sin_filtrar + filtro
    dashboard.api_stats_banco         -> stats_banco (pura)
    tab_ar_real.api_facturas          -> facturas_y_stats (pura)
    tab_fb_dashboard.api_resultados   -> resumen_fb (pura)

Son extracciones que conservan el comportamiento por construccion —el sitio de
siempre llama al cuerpo de siempre—, pero "por construccion" es exactamente lo
que uno se dice antes de romper algo. Esto lo comprueba de verdad, pidiendo los
endpoints y mirando que contestan y con que.

No sustituye a verificarlo en el navegador. Sirve para no llegar al navegador
con algo roto de la forma tonta.
"""
import json
import os
import shutil
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

TENANT = "fase-a-smoke"
os.environ["YVE_TENANT"] = TENANT
os.environ.pop("YVE_HOTEL", None)

import pandas as pd                                    # noqa: E402
import tenant_dirs                                     # noqa: E402

H1 = "HAAA111"


def _montar():
    base = os.path.join(BASE, "tenants", TENANT)
    shutil.rmtree(base, ignore_errors=True)
    tenant_dirs.tenant_base()
    datos = tenant_dirs.datos_dir()
    proc  = tenant_dirs.procesadas_dir()

    json.dump([{"id": H1, "nombre": "Hotel Uno", "activo": True}],
              open(os.path.join(datos, "hoteles.json"), "w", encoding="utf-8"))

    pd.DataFrame([{"numero_factura": "AP-001", "nombre_proveedor": "Prov A",
                   "total_factura": 100.0, "estado_matching": "MATCH_CORRECTO",
                   "hotel_id": H1}]).to_excel(
        os.path.join(proc, "facturas_ap_20260729.xlsx"), index=False)

    pd.DataFrame([{"numero_factura": "OTA-1", "nombre_ota": "Booking",
                   "periodo_inicio": "2026-06", "importe_bruto": 1000.0,
                   "estado": "DISCREPANCIA", "discrepancia_euros": 120.0,
                   "hotel_id": H1}]).to_excel(
        os.path.join(proc, "facturas_procesadas_20260729.xlsx"), index=False)

    pd.DataFrame([{"numero_reserva": "R-1", "cliente": "Cliente A", "total": 800.0,
                   "estado": "FACTURADO", "fecha_emision": "2026-07-20",
                   "hotel_id": H1}]).to_excel(
        os.path.join(datos, "reservas_credito.xlsx"), index=False)

    pd.DataFrame([{"id_receta": "R01", "nombre": "Ensalada", "categoria": "Entrantes",
                   "precio_venta": 12.0,
                   "ingredientes_json": json.dumps(
                       [{"ingrediente": "Tomate", "cantidad": 0.2}])}]).to_excel(
        os.path.join(datos, "recetas.xlsx"), index=False)
    pd.DataFrame([{"ingrediente": "Tomate", "stock_actual_kg_l": 10,
                   "stock_inicial_kg_l": 12, "coste_unitario": 2.0,
                   "hotel_id": H1}]).to_excel(
        os.path.join(datos, "inventario.xlsx"), index=False)
    pd.DataFrame([{"fecha": "2026-07-20", "id_receta": "R01", "nombre_plato": "Ensalada",
                   "categoria": "Entrantes", "unidades_vendidas": 10,
                   "total_venta": 120.0, "hotel_id": H1}]).to_excel(
        os.path.join(datos, "ventas_fb_diarias.xlsx"), index=False)
    pd.DataFrame([{"fecha": "2026-07-20", "ingrediente": "Tomate", "cantidad": 1,
                   "coste_merma": 2.0, "hotel_id": H1}]).to_excel(
        os.path.join(datos, "mermas.xlsx"), index=False)

    pd.DataFrame([{"fecha": "2026-07-20", "concepto": "Pago Prov A",
                   "importe": -100.0}]).to_excel(
        os.path.join(datos, "extracto_banco.xlsx"), index=False)
    return base


# Cada caso: (ruta, comprobacion sobre el json). La comprobacion mira el DATO,
# no solo el 200: un endpoint que devuelve 200 con la tabla vacia es justo el
# fallo que no queremos que pase desapercibido.
CASOS = [
    ("/api/stats",            lambda d: d.get("total") == 1),
    ("/api/stats_ap",         lambda d: d.get("total") == 1 and d.get("importe") == 100.0),
    ("/api/stats_banco",      lambda d: d is not None and d.get("total") == 1),
    ("/api/ar_real/facturas", lambda d: d.get("ok") and d["stats"]["total_facturas"] == 1),
    ("/fb/api/resultados",    lambda d: d.get("resumen", {}).get("total_ventas") == 120.0),
    ("/api/multi_hotel/agregado",
     lambda d: d.get("ok") and d["cuadra"] and d["n_hoteles"] == 1),
]

# El Excel de Multi-Hotel ya no devuelve seis hoteles escritos a mano: sale del
# agregador. Se comprueba aparte porque no es JSON.
def _comprobar_excel(c):
    import io
    import pandas as _pd
    r = c.get("/api/exportar/multihotel")
    if r.status_code != 200:
        return f"status {r.status_code}"
    df = _pd.read_excel(io.BytesIO(r.data), sheet_name="Multi-Hotel")
    nombres = df["Hotel"].astype(str).tolist()
    if "Premier London Mayfair" in nombres:
        return "sigue devolviendo los hoteles inventados"
    if "Hotel Uno" not in nombres:
        return f"no sale el hotel real; sale {nombres}"
    if "Grupo" not in nombres:
        return "falta la fila del grupo"
    fila = df[df["Hotel"] == "Hotel Uno"].iloc[0]
    if float(fila["Importe AP (EUR)"]) != 100.0:
        return f"importe AP {fila['Importe AP (EUR)']} != 100.0"
    return None


def _sesion_de_prueba():
    """Un usuario de mentira para flask-login, SIN tocar el fichero de usuarios.

    Crear un usuario de verdad ensuciaria el almacen compartido, y borrarlo
    despues es justo el tipo de limpieza que un dia falla y deja una cuenta
    viva. Se sustituye el cargador y ya.
    """
    import auth
    real = auth.get_usuario
    auth.get_usuario = lambda username: auth.Usuario(
        {"username": username, "nombre": "Test", "rol": "admin",
         "tenant": TENANT, "activo": True})
    return real


def main():
    base = _montar()
    import dashboard

    _sesion_de_prueba()
    dashboard.app.config["TESTING"] = True
    fallos = []
    print("\n── Fase A · los paneles siguen contestando ────────────────────")
    with dashboard.app.test_client() as c:
        with c.session_transaction() as s:
            s["tenant_id"] = TENANT
            s["_user_id"]  = "test"
            s["_fresh"]    = True
        for ruta, ok in CASOS:
            try:
                r = c.get(ruta)
                cuerpo = r.get_json()
                bien = r.status_code == 200 and ok(cuerpo)
                print(f"  {'OK ' if bien else 'MAL'} {ruta:<32} {r.status_code}")
                if not bien:
                    fallos.append(f"{ruta} -> {r.status_code} {str(cuerpo)[:220]}")
            except Exception as e:
                print(f"  MAL {ruta:<32} EXCEPCION")
                fallos.append(f"{ruta} -> {type(e).__name__}: {e}")

        err = _comprobar_excel(c)
        print(f"  {'OK ' if not err else 'MAL'} {'/api/exportar/multihotel':<32} "
              f"{'excel del agregador' if not err else err}")
        if err:
            fallos.append(f"/api/exportar/multihotel -> {err}")

    shutil.rmtree(base, ignore_errors=True)

    if fallos:
        print("\n FALLOS:")
        for f in fallos:
            print("   ·", f)
        return 1
    print("\n OK · los paneles contestan, y el Excel sale del agregador\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
