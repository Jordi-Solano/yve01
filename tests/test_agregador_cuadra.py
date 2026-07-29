"""FASE A · El agregador tiene que dar lo MISMO que los paneles de cada hotel.

Esta es la verificacion de la fase A y la razon de que la fase A exista antes
que cualquier pantalla: los paneles por hotel estan comprobados uno a uno en
produccion (fases 3, 4b y 5). Si el agregador llega a los mismos numeros por un
camino distinto —partiendo en vez de filtrando— es que esta bien. Si
empezaramos por la pantalla estariamos estrenando dos cosas a la vez y, al
discrepar, no sabriamos cual de las dos miente.

El montaje reparte documentos entre tres hoteles y ademas mete dos casos que en
produccion pasan de verdad y que son justo donde se pierde el cuadre:

  - un documento SIN etiqueta de hotel            -> caja `sin_asignar`
  - un documento con una etiqueta que no esta en el censo -> caja `desconocido`
    (pasa en cada despliegue: `hoteles.json` va commiteado como [] y Render no
    tiene disco, asi que los documentos sobreviven a los hoteles)

El tercer hotel se queda a proposito SIN documentos: un hotel vacio tiene que
salir con ceros, no desaparecer del panel ni —mucho peor— heredar el total del
grupo, que es el fallo de fase 0 multiplicado por N.
"""
import json
import os
import shutil
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

TENANT = "fase-a-test"
os.environ["YVE_TENANT"] = TENANT
os.environ.pop("YVE_HOTEL", None)

import pandas as pd                                    # noqa: E402
import tenant_dirs                                     # noqa: E402

H1, H1N = "HAAA111", "Hotel Uno"
H2, H2N = "HBBB222", "Hotel Dos"
H3, H3N = "HCCC333", "Hotel Tres"
FANTASMA = "HZZZ999"          # etiquetado, pero no esta en el censo


def _montar():
    base = os.path.join(BASE, "tenants", TENANT)
    shutil.rmtree(base, ignore_errors=True)
    tenant_dirs.tenant_base()                          # crea el arbol vacio
    datos = tenant_dirs.datos_dir()
    proc  = tenant_dirs.procesadas_dir()

    json.dump([{"id": H1, "nombre": H1N, "activo": True},
               {"id": H2, "nombre": H2N, "activo": True},
               {"id": H3, "nombre": H3N, "activo": True}],
              open(os.path.join(datos, "hoteles.json"), "w", encoding="utf-8"))

    # ── AP ────────────────────────────────────────────────────────────────
    pd.DataFrame([
        {"numero_factura": "AP-001", "nombre_proveedor": "Prov A", "total_factura": 100.0,
         "estado_matching": "MATCH_CORRECTO",  "hotel_id": H1},
        {"numero_factura": "AP-002", "nombre_proveedor": "Prov A", "total_factura": 250.5,
         "estado_matching": "DISCREPANCIA_PO", "hotel_id": H1},
        {"numero_factura": "AP-003", "nombre_proveedor": "Prov B", "total_factura": 75.25,
         "estado_matching": "MATCH_CORRECTO",  "hotel_id": H2},
        {"numero_factura": "AP-004", "nombre_proveedor": "Prov B", "total_factura": 40.0,
         "estado_matching": "SIN_PO",          "hotel_id": H2},
        {"numero_factura": "AP-005", "nombre_proveedor": "Prov C", "total_factura": 33.0,
         "estado_matching": "MATCH_CORRECTO",  "hotel_id": ""},          # sin asignar
        {"numero_factura": "AP-006", "nombre_proveedor": "Prov C", "total_factura": 12.0,
         "estado_matching": "MATCH_CORRECTO",  "hotel_id": FANTASMA},    # desconocido
    ]).to_excel(os.path.join(proc, "facturas_ap_20260729.xlsx"), index=False)

    # ── AR / OTA ──────────────────────────────────────────────────────────
    pd.DataFrame([
        {"numero_factura": "OTA-1", "nombre_ota": "Booking", "periodo_inicio": "2026-06",
         "importe_bruto": 1000.0, "estado": "DISCREPANCIA", "discrepancia_euros": 120.0,
         "hotel_id": H1},
        {"numero_factura": "OTA-2", "nombre_ota": "Expedia", "periodo_inicio": "2026-06",
         "importe_bruto": 500.0,  "estado": "CORRECTO", "discrepancia_euros": 0.0,
         "hotel_id": H2},
        {"numero_factura": "OTA-3", "nombre_ota": "Booking", "periodo_inicio": "2026-06",
         "importe_bruto": 300.0,  "estado": "DISCREPANCIA", "discrepancia_euros": 45.5,
         "hotel_id": ""},
    ]).to_excel(os.path.join(proc, "facturas_procesadas_20260729.xlsx"), index=False)

    # ── AR Real ───────────────────────────────────────────────────────────
    pd.DataFrame([
        {"numero_reserva": "R-1", "cliente": "Cliente A", "total": 800.0,
         "estado": "FACTURADO", "fecha_emision": "2026-07-20", "hotel_id": H1},
        {"numero_reserva": "R-2", "cliente": "Cliente B", "total": 400.0,
         "estado": "COBRADO",   "fecha_emision": "2026-06-01", "hotel_id": H2},
        {"numero_reserva": "R-3", "cliente": "Cliente C", "total": 150.0,
         "estado": "FACTURADO", "fecha_emision": "2026-07-25", "hotel_id": FANTASMA},
    ]).to_excel(os.path.join(datos, "reservas_credito.xlsx"), index=False)

    # ── F&B ───────────────────────────────────────────────────────────────
    # El recetario es del GRUPO (no lleva hotel_id): una cadena comparte carta.
    pd.DataFrame([
        {"id_receta": "R01", "nombre": "Ensalada", "categoria": "Entrantes",
         "precio_venta": 12.0,
         "ingredientes_json": json.dumps([{"ingrediente": "Tomate", "cantidad": 0.2}])},
    ]).to_excel(os.path.join(datos, "recetas.xlsx"), index=False)
    # El inventario SI es del hotel, y por eso el mismo plato cuesta distinto.
    pd.DataFrame([
        {"ingrediente": "Tomate", "stock_actual_kg_l": 10, "stock_inicial_kg_l": 12,
         "coste_unitario": 2.0, "hotel_id": H1},
        {"ingrediente": "Tomate", "stock_actual_kg_l": 8,  "stock_inicial_kg_l": 10,
         "coste_unitario": 9.0, "hotel_id": H2},
    ]).to_excel(os.path.join(datos, "inventario.xlsx"), index=False)
    pd.DataFrame([
        {"fecha": "2026-07-20", "id_receta": "R01", "nombre_plato": "Ensalada",
         "categoria": "Entrantes", "unidades_vendidas": 10, "total_venta": 120.0,
         "hotel_id": H1},
        {"fecha": "2026-07-21", "id_receta": "R01", "nombre_plato": "Ensalada",
         "categoria": "Entrantes", "unidades_vendidas": 5,  "total_venta": 60.0,
         "hotel_id": H2},
        {"fecha": "2026-07-22", "id_receta": "R01", "nombre_plato": "Ensalada",
         "categoria": "Entrantes", "unidades_vendidas": 2,  "total_venta": 24.0,
         "hotel_id": ""},
    ]).to_excel(os.path.join(datos, "ventas_fb_diarias.xlsx"), index=False)
    pd.DataFrame([
        {"fecha": "2026-07-20", "ingrediente": "Tomate", "cantidad": 1,
         "coste_merma": 2.0, "hotel_id": H1},
    ]).to_excel(os.path.join(datos, "mermas.xlsx"), index=False)

    return base


def _con_hotel(hid, fn):
    """Ejecuta `fn` como si ese hotel estuviera elegido en la sesion.

    Sin peticion Flask, `censo_hoteles.activo()` lee YVE_HOTEL — el mismo
    camino que usan los subprocesos. Es el estado del panel con ese hotel.
    """
    previo = os.environ.get("YVE_HOTEL")
    os.environ["YVE_HOTEL"] = hid
    try:
        return fn()
    finally:
        if previo is None:
            os.environ.pop("YVE_HOTEL", None)
        else:
            os.environ["YVE_HOTEL"] = previo


def _limpiar_cache_fb():
    import tab_fb_dashboard as fb
    for attr in ("_CACHE", "_cache", "_XLSX_CACHE"):
        c = getattr(fb, attr, None)
        if isinstance(c, dict):
            c.clear()


def main():
    base = _montar()
    fallos = []

    import agregador_grupo
    import dashboard
    import tab_ar_real
    import tab_fb_dashboard

    ag = agregador_grupo.agregado()
    cajas = {f["hotel_id"]: f for f in ag["hoteles"]}
    cajas["sin_asignar"] = ag["sin_asignar"]
    cajas["desconocido"] = ag["desconocido"]

    def comprobar(etiqueta, esperado, obtenido):
        if esperado != obtenido:
            fallos.append(f"{etiqueta}: panel={esperado!r} agregador={obtenido!r}")

    # ── 1. Cada hotel, contra su panel ────────────────────────────────────
    for hid, nombre in ((H1, H1N), (H2, H2N), (H3, H3N)):
        caja = cajas[hid]

        # AP — el panel es cargar_datos_ap() + calcular_stats_ap()
        s_ap = _con_hotel(hid, lambda: dashboard.calcular_stats_ap(dashboard.cargar_datos_ap()))
        comprobar(f"{nombre} · AP facturas", s_ap["total"],   caja["ap"]["facturas"])
        comprobar(f"{nombre} · AP importe",  s_ap["importe"], caja["ap"]["importe"])
        comprobar(f"{nombre} · AP cuadran",  s_ap["matches"], caja["ap"]["cuadran"])
        comprobar(f"{nombre} · AP discrepancias",
                  s_ap["discrepancias"], caja["ap"]["discrepancias"])

        # AR/OTA — el panel es cargar_datos() + calcular_stats()
        s_ar = _con_hotel(hid, lambda: dashboard.calcular_stats(dashboard.cargar_datos()[0]))
        comprobar(f"{nombre} · AR facturas", s_ar["total"], caja["ar_ota"]["facturas"])
        comprobar(f"{nombre} · AR reclamable",
                  s_ar["importe_reclamable"], caja["ar_ota"]["importe_reclamable"])

        # AR Real — el panel es _get_reservas() + facturas_y_stats()
        _l, s_rr = _con_hotel(hid, lambda: tab_ar_real.facturas_y_stats(tab_ar_real._get_reservas()))
        comprobar(f"{nombre} · AR Real facturas",
                  s_rr["total_facturas"], caja["ar_real"]["facturas"])
        comprobar(f"{nombre} · AR Real pendiente",
                  s_rr["pendiente"], caja["ar_real"]["pendiente"])

        # F&B — el panel es _xlsx_hotel() + resumen_fb()
        def _fb():
            _limpiar_cache_fb()
            return tab_fb_dashboard.resumen_fb(
                tab_fb_dashboard._xlsx("recetas.xlsx"),
                tab_fb_dashboard._xlsx_hotel("inventario.xlsx"),
                tab_fb_dashboard._xlsx_hotel("ventas_fb_diarias.xlsx"),
                tab_fb_dashboard._xlsx_hotel("mermas.xlsx"))[2]
        r_fb = _con_hotel(hid, _fb)
        comprobar(f"{nombre} · F&B ventas",
                  r_fb["total_ventas"], caja["fb"]["ventas"])
        comprobar(f"{nombre} · F&B food cost %",
                  r_fb["fc_teorico_pct"], caja["fb"]["food_cost_pct"])

    # ── 2. El hotel vacio sale a cero, no hereda el grupo ─────────────────
    vacio = cajas[H3]
    if vacio["ap"]["facturas"] != 0 or vacio["ar_ota"]["facturas"] != 0:
        fallos.append(f"{H3N} no tiene documentos y no sale a cero: {vacio}")
    if vacio["ap"]["importe"] == ag["grupo"]["ap"]["importe"] and ag["grupo"]["ap"]["importe"]:
        fallos.append(f"{H3N} esta heredando el total del grupo (fallo de fase 0)")

    # ── 3. Las dos cajas especiales, separadas ───────────────────────────
    if cajas["sin_asignar"]["ap"]["facturas"] != 1:
        fallos.append(f"sin_asignar deberia tener 1 factura AP, tiene "
                      f"{cajas['sin_asignar']['ap']['facturas']}")
    if cajas["desconocido"]["ap"]["facturas"] != 1:
        fallos.append(f"desconocido deberia tener 1 factura AP, tiene "
                      f"{cajas['desconocido']['ap']['facturas']}")

    # ── 4. El food cost del grupo esta PONDERADO, no aplanado ────────────
    #
    # El mismo tomate cuesta 2 EUR en un hotel y 9 EUR en el otro. Costeando
    # con el inventario del grupo entero, el segundo pisa al primero y sale el
    # 15,0% del hotel caro como si fuera el del grupo. El ponderado real es
    # (4,0 + 9,0) / (120 + 60 + 24) = 6,37%.
    fc_grupo = ag["grupo"]["fb"]["food_cost_pct"]
    fc_h2    = cajas[H2]["fb"]["food_cost_pct"]
    if abs(fc_grupo - 6.37) > 0.05:
        fallos.append(f"food cost del grupo deberia ser 6.37% ponderado, es {fc_grupo}%")
    if abs(fc_grupo - fc_h2) < 0.01:
        fallos.append(f"el food cost del grupo ({fc_grupo}%) es identico al del "
                      f"hotel caro: el inventario se esta aplanando")

    # ── 5. El cuadre ─────────────────────────────────────────────────────
    for fila in ag["cuadre"]:
        if not fila["cuadra"]:
            fallos.append(f"DESCUADRE en {fila['metrica']}: cajas={fila['suma_cajas']} "
                          f"grupo={fila['total_grupo']} (dif {fila['diferencia']})")

    # ── Informe ──────────────────────────────────────────────────────────
    print("\n── Fase A · agregador contra los paneles ──────────────────────")
    for f in ag["hoteles"] + [ag["sin_asignar"], ag["desconocido"], ag["grupo"]]:
        print(f"  {f['nombre']:<20} AP {f['ap']['facturas']:>2} / "
              f"{f['ap']['importe']:>8.2f} €   "
              f"OTA {f['ar_ota']['facturas']:>2} recl {f['ar_ota']['importe_reclamable']:>7.2f} €   "
              f"ARR {f['ar_real']['facturas']:>2}   "
              f"F&B {f['fb']['ventas']:>7.2f} € fc {f['fb']['food_cost_pct']:>5.1f}%")
    print(f"\n  cuadre: {'SI' if ag['cuadra'] else 'NO'}  "
          f"({len(ag['cuadre'])} metricas aditivas comprobadas)")

    shutil.rmtree(base, ignore_errors=True)

    if fallos:
        print("\n FALLOS:")
        for f in fallos:
            print("   ·", f)
        return 1
    print("\n OK · el agregador da lo mismo que los paneles, y la suma cuadra\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
