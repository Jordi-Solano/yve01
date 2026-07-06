"""
demo_generator.py — Yve.01
Genera datos ficticios realistas con los nombres que da el usuario
(hotel suelto, cadena o varias cadenas) para demos a clientes y gestorías.
"""
import os, json, random, hashlib
from datetime import datetime, timedelta, date
from pathlib import Path
import pandas as pd

BASE_DIR   = Path(__file__).parent
DATOS      = BASE_DIR / "datos-referencia"
REPORTES   = BASE_DIR / "reportes"
PROCESADAS = BASE_DIR / "facturas-procesadas"

CIUDADES   = ["Barcelona", "Madrid", "Sitges", "Valencia", "Sevilla", "Málaga", "Bilbao", "Palma", "Girona", "Tarragona"]
OTAS       = ["Booking.com", "Expedia", "Hotelbeds", "Agoda"]
PROVEEDORES = [
    ("Makro Cash & Carry", "A-28647451", "Alimentación y bebidas"),
    ("Pescados del Puerto SL", "B-63219845", "Pescado fresco"),
    ("Carnes Selectas SL", "B-61234578", "Carnes frescas"),
    ("Lavandería Industrial Norte", "B-59887123", "Servicio lavandería"),
    ("Energía Ibérica SA", "A-81948077", "Suministro eléctrico"),
    ("Frutas y Verduras Mercado SL", "B-64111098", "Fruta y verdura"),
    ("Limpiezas Profesionales BCN", "B-62009911", "Limpieza zonas comunes"),
    ("Telecomunicaciones Sur SA", "A-80229944", "Telefonía e internet"),
]
PLATOS = [
    (1, "Paella de marisco", "Arroces", 27.0, [("arroz bomba", 0.12), ("gambas", 0.10), ("mejillones", 0.15)]),
    (2, "Entrecot a la brasa", "Carnes", 32.0, [("entrecot vacuno", 0.30), ("patatas", 0.20)]),
    (3, "Lubina al horno", "Pescados", 27.0, [("lubina fresca", 0.35), ("patatas", 0.15)]),
    (4, "Ensalada mediterránea", "Entrantes", 14.0, [("tomate", 0.20), ("queso fresco", 0.10)]),
    (5, "Crema catalana", "Postres", 8.5, [("huevos", 0.10), ("leche", 0.20)]),
    (6, "Risotto de setas", "Arroces", 19.0, [("arroz bomba", 0.11), ("setas", 0.12)]),
    (7, "Pulpo a la gallega", "Entrantes", 22.0, [("pulpo", 0.25), ("patatas", 0.15)]),
    (8, "Tarta de queso", "Postres", 9.0, [("queso fresco", 0.15), ("huevos", 0.08)]),
]
INGREDIENTES = [
    ("arroz bomba", "Secos", 3.2), ("gambas", "Mariscos", 22.0), ("mejillones", "Mariscos", 4.5),
    ("entrecot vacuno", "Carnes", 18.5), ("patatas", "Verduras", 1.1), ("lubina fresca", "Pescados", 14.0),
    ("tomate", "Verduras", 2.4), ("queso fresco", "Lácteos", 7.8), ("huevos", "Lácteos", 3.4),
    ("leche", "Lácteos", 1.05), ("setas", "Verduras", 9.5), ("pulpo", "Mariscos", 16.5),
]
EMPRESAS = [
    ("Viajes Corporativos Iberia SL", "B-84512367"), ("Congresos y Eventos BCN SA", "A-58221199"),
    ("TechGlobal España SL", "B-87654321"), ("Agencia Nórdica Travel", "N-0034567-B"),
    ("Grupo Farmacéutico Levante SA", "A-46019283"),
]


def _seed_from(nombres):
    h = hashlib.md5("|".join(sorted(nombres)).encode()).hexdigest()
    return int(h[:8], 16)


def generar_demo(cadenas):
    """cadenas: [{'nombre': 'Cadena Sol', 'hoteles': ['Hotel Sol Mar', ...]}, ...]
    Devuelve resumen con contadores."""
    todos = [h for c in cadenas for h in c["hoteles"]]
    if not todos:
        raise ValueError("Sin hoteles")
    rng = random.Random(_seed_from(todos))
    hoy = date.today()
    stamp = hoy.strftime("%Y%m%d")

    # ── 1. hoteles.json + kpis_hoteles.xlsx (6 meses por hotel) ──────────
    registro, kpis = [], []
    hid = 0
    for cad in cadenas:
        for hotel in cad["hoteles"]:
            hid += 1
            hotel_id = f"DEMO{hid:02d}"
            cat = rng.choice(["3 estrellas", "4 estrellas", "4 estrellas", "5 estrellas"])
            estrellas = int(cat[0])
            hab = rng.randint(60, 120) if estrellas == 3 else rng.randint(90, 180) if estrellas == 4 else rng.randint(120, 260)
            ciudad = rng.choice(CIUDADES)
            registro.append({
                "id": hotel_id, "nombre": hotel, "ciudad": ciudad, "categoria": cat,
                "habitaciones": hab, "activo": True,
                "contacto": f"finanzas@{hotel.lower().replace(' ', '')[:14]}.es",
                "modulos": ["AR", "AP", "DRR", "FB", "Banco"],
            })
            adr_base = {3: rng.uniform(75, 110), 4: rng.uniform(110, 175), 5: rng.uniform(180, 290)}[estrellas]
            occ_base = rng.uniform(0.64, 0.87)
            for m in range(6):
                mes_d = (hoy.replace(day=1) - timedelta(days=30 * (5 - m)))
                mes = mes_d.strftime("%Y-%m")
                estac = 1 + 0.13 * (1 if mes_d.month in (6, 7, 8) else -1 if mes_d.month in (1, 2, 11) else 0)
                occ = min(0.96, max(0.42, occ_base * estac * rng.uniform(0.94, 1.06)))
                adr = adr_base * estac * rng.uniform(0.96, 1.05)
                revpar = adr * occ
                ing_rooms = revpar * hab * 30
                ing_fb = ing_rooms * rng.uniform(0.18, 0.30)
                ing_otros = ing_rooms * rng.uniform(0.03, 0.07)
                total = ing_rooms + ing_fb + ing_otros
                gop_pct = rng.uniform(0.26, 0.42)
                kpis.append({
                    "hotel_id": hotel_id, "hotel_nombre": hotel, "grupo": cad["nombre"],
                    "ciudad": ciudad, "categoria": cat, "mes": mes, "habitaciones": hab,
                    "ocupacion_pct": round(occ * 100, 1), "adr_eur": round(adr, 2),
                    "revpar_eur": round(revpar, 2), "ingresos_rooms": round(ing_rooms, 2),
                    "ingresos_fb": round(ing_fb, 2), "ingresos_otros": round(ing_otros, 2),
                    "total_ingresos": round(total, 2), "coste_ap_eur": round(total * rng.uniform(0.16, 0.24), 2),
                    "coste_ar_comisiones": round(ing_rooms * rng.uniform(0.09, 0.16), 2),
                    "food_cost_pct": round(rng.uniform(24, 34), 1), "gop_eur": round(total * gop_pct, 2),
                    "gop_pct": round(gop_pct * 100, 1),
                    "facturas_ap_pendientes": rng.randint(0, 12), "facturas_ar_pendientes": rng.randint(0, 6),
                    "alertas_activas": rng.choice([0, 0, 0, 1, 1, 2]),
                    "estado_oracle": rng.choice(["OK", "OK", "OK", "PENDIENTE"]),
                    "out_of_balance_dias": rng.choice([0, 0, 0, 0, 1]),
                })
    json.dump(registro, open(DATOS / "hoteles.json", "w", encoding="utf-8"), indent=2, ensure_ascii=False)
    pd.DataFrame(kpis).to_excel(DATOS / "kpis_hoteles.xlsx", index=False)

    # ── 2. AR: facturas OTA verificadas ───────────────────────────────────
    ar_rows = []
    for i in range(min(14, 4 + 3 * len(todos))):
        hotel = rng.choice(todos)
        ota = rng.choice(OTAS)
        bruto = round(rng.uniform(6000, 42000), 2)
        pct_pactado = rng.choice([15.0, 15.0, 17.0, 18.0])
        desvio = rng.choice([0, 0, 0, 0, 0.8, 1.5, -0.7])
        pct_fact = round(pct_pactado + desvio, 1)
        discrepancia = round(bruto * desvio / 100, 2) if desvio else 0.0
        extranjera = ota in ("Booking.com", "Expedia", "Agoda")
        ar_rows.append({
            "archivo": f"{ota.lower().replace('.', '').replace(' ', '_')}_{stamp}_{i+1:02d}.pdf",
            "numero_factura": f"{ota[:3].upper()}-{2026000 + rng.randint(100, 999)}-{i+1:03d}",
            "fecha": (hoy - timedelta(days=rng.randint(2, 55))).strftime("%d/%m/%Y"),
            "nombre_ota": ota.upper().replace(".COM", ".COM"),
            "nombre_hotel": hotel,
            "importe_bruto": bruto,
            "porcentaje_pactado": pct_pactado,
            "porcentaje_factura": pct_fact,
            "comision_calculada": round(bruto * pct_fact / 100, 2),
            "estado": "DISCREPANCIA" if discrepancia else "CORRECTO",
            "discrepancia_euros": discrepancia if discrepancia else "",
            "estado_di": "FALTA_CERTIFICADO_DI" if extranjera and rng.random() < 0.35 else "OK",
            "accion": "",
        })
    pd.DataFrame(ar_rows).to_excel(REPORTES / f"verificacion_{stamp}.xlsx", index=False)

    # ── 3. AP: facturas proveedor ────────────────────────────────────────
    ap_rows = []
    for i, (prov, nif, desc) in enumerate(PROVEEDORES):
        base = round(rng.uniform(600, 9000), 2)
        iva_pct = 21 if "Energ" in prov or "Telecom" in prov or "Limpiezas" in prov or "Lavand" in prov else 10
        cuota = round(base * iva_pct / 100, 2)
        estado_m = rng.choice(["MATCH_3WAY_OK", "MATCH_3WAY_OK", "MATCH_3WAY_OK", "MATCH_CORRECTO", "DISCREPANCIA_PO", "SIN_PO"])
        ap_rows.append({
            "archivo": f"factura_{prov.split()[0].lower()}_{stamp}.pdf",
            "numero_factura": f"{prov.split()[0][:3].upper()}-2026-{100 + i:03d}",
            "fecha": (hoy - timedelta(days=rng.randint(3, 40))).strftime("%d/%m/%Y"),
            "nombre_proveedor": prov, "NIF_proveedor": nif,
            "descripcion_concepto": desc,
            "base_imponible": base, "porcentaje_iva": iva_pct, "cuota_iva": cuota,
            "total_factura": round(base + cuota, 2),
            "estado_matching": estado_m,
            "hotel": rng.choice(todos),
        })
    pd.DataFrame(ap_rows).to_excel(PROCESADAS / f"facturas_ap_{stamp}.xlsx", index=False)

    # ── 4. Banco: la mitad de los cargos casan con facturas AP ──────────
    movs, saldo = [], 120000.0
    fecha_mov = hoy - timedelta(days=28)
    for i, ap in enumerate(ap_rows):
        if i % 2 == 0:  # pago que casa con la factura
            imp = -ap["total_factura"]
            concepto = f"Pago {ap['nombre_proveedor']} - Fra. {ap['numero_factura']}"
            ref = ap["numero_factura"][-7:]
        else:
            imp = round(rng.uniform(900, 14000), 2)
            concepto = f"Cobro {rng.choice(EMPRESAS)[0]} - RES-2026-{rng.randint(300, 499):04d}"
            ref = f"RES-{rng.randint(300, 499):04d}"
        saldo = round(saldo + imp, 2)
        movs.append({"fecha": fecha_mov.strftime("%Y-%m-%d"), "concepto": concepto,
                     "importe": imp, "saldo": saldo, "referencia": ref, "conciliado": ""})
        fecha_mov += timedelta(days=rng.randint(1, 4))
    pd.DataFrame(movs).to_excel(DATOS / "extracto_banco.xlsx", index=False)

    # ── 5. F&B: recetas, inventario, ventas 30 días, mermas ──────────────
    pd.DataFrame([
        {"id_receta": pid, "nombre": nom, "categoria": cat, "precio_venta": precio,
         "ingredientes_json": json.dumps([{"ingrediente": ing, "cantidad": cant} for ing, cant in ings])}
        for pid, nom, cat, precio, ings in PLATOS
    ]).to_excel(DATOS / "recetas.xlsx", index=False)

    pd.DataFrame([
        {"ingrediente": ing, "categoria": cat, "coste_unitario": coste,
         "stock_actual_kg_l": round(rng.uniform(3, 60), 1),
         "stock_inicial_kg_l": round(rng.uniform(40, 90), 1),
         "unidad": "kg" if cat != "Lácteos" else "l", "proveedor": rng.choice(PROVEEDORES)[0]}
        for ing, cat, coste in INGREDIENTES
    ]).to_excel(DATOS / "inventario.xlsx", index=False)

    ventas = []
    for d in range(30):
        fecha_v = (hoy - timedelta(days=29 - d)).strftime("%Y-%m-%d")
        for pid, nom, cat, precio, _ in rng.sample(PLATOS, rng.randint(4, 7)):
            uds = rng.randint(3, 26)
            ventas.append({"fecha": fecha_v, "id_receta": pid, "nombre_plato": nom, "categoria": cat,
                           "unidades_vendidas": uds, "total_venta": round(uds * precio, 2)})
    pd.DataFrame(ventas).to_excel(DATOS / "ventas_fb_diarias.xlsx", index=False)

    pd.DataFrame([
        {"fecha": (hoy - timedelta(days=rng.randint(1, 28))).strftime("%Y-%m-%d"),
         "ingrediente": ing, "categoria": cat,
         "cantidad_merma": round(rng.uniform(0.5, 4.0), 1), "unidad": "kg",
         "causa": rng.choice(["Caducidad", "Rotura de frío", "Cocción fallida", "Exceso de producción"]),
         "coste_merma": round(rng.uniform(4, 60), 2)}
        for ing, cat, _ in rng.sample(INGREDIENTES, 8)
    ]).to_excel(DATOS / "mermas.xlsx", index=False)

    # ── 6. AR Real: clientes corporativos + facturas con aging ──────────
    pd.DataFrame([
        {"nombre_cliente": emp, "nif": nif, "credito_limite": rng.choice([15000, 25000, 40000, 60000]),
         "credito_usado": round(rng.uniform(2000, 22000), 2), "dias_pago": rng.choice([30, 45, 60]),
         "email": f"cuentas@{emp.lower().split()[0]}.com"}
        for emp, nif in EMPRESAS
    ]).to_excel(DATOS / "clientes_credito.xlsx", index=False)

    reservas = []
    for i in range(9):
        emp = rng.choice(EMPRESAS)[0]
        dias_atras = rng.choice([5, 12, 20, 35, 41, 55, 70, 95, 110])
        emision = hoy - timedelta(days=dias_atras)
        reservas.append({
            "numero": f"AR-2026-{400 + i:04d}", "cliente": emp,
            "fecha_emision": emision.strftime("%Y-%m-%d"),
            "fecha_entrada": (emision - timedelta(days=rng.randint(5, 20))).strftime("%Y-%m-%d"),
            "fecha_salida": emision.strftime("%Y-%m-%d"),
            "habitaciones": rng.randint(2, 18), "noches": rng.randint(1, 5),
            "importe": round(rng.uniform(1200, 18000), 2),
            "estado": "COBRADA" if dias_atras < 15 and rng.random() < 0.5 else "PENDIENTE",
            "hotel": rng.choice(todos),
        })
    pd.DataFrame(reservas).to_excel(DATOS / "reservas_credito.xlsx", index=False)

    return {"hoteles": len(todos), "cadenas": len(cadenas), "facturas_ar": len(ar_rows),
            "facturas_ap": len(ap_rows), "movimientos_banco": len(movs),
            "ventas_fb": len(ventas), "clientes": len(EMPRESAS)}


def limpiar_demo():
    """Vacía los datos generados (cabeceras, 0 filas) y borra reportes demo."""
    import glob
    vacios = {
        "hoteles.json": [],
    }
    json.dump([], open(DATOS / "hoteles.json", "w"), indent=2)
    for f in ["kpis_hoteles.xlsx", "extracto_banco.xlsx", "ventas_fb_diarias.xlsx",
              "inventario.xlsx", "mermas.xlsx", "recetas.xlsx",
              "clientes_credito.xlsx", "reservas_credito.xlsx"]:
        p = DATOS / f
        if p.exists():
            try:
                df = pd.read_excel(p)
                df.iloc[0:0].to_excel(p, index=False)
            except Exception:
                pass
    for pat in ["verificacion_*.xlsx", "conciliacion_*.xlsx", "doble_imposicion_*.xlsx"]:
        for f in glob.glob(str(REPORTES / pat)):
            try: os.remove(f)
            except Exception: pass
    for f in glob.glob(str(PROCESADAS / "facturas_ap_*.xlsx")) + glob.glob(str(PROCESADAS / "facturas_contabilizadas_*.xlsx")):
        try: os.remove(f)
        except Exception: pass
