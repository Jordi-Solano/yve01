# -*- coding: utf-8 -*-
"""Lo que devolveria Claude para cada PDF/foto del plan (leido bien)."""
def fac(num, fecha, prov, nif, concepto, base, pct, cuota, total, lineas=None):
    d = {"es_factura": True, "numero_factura": num, "fecha": fecha, "nombre_proveedor": prov, "NIF_proveedor": nif,
         "descripcion_concepto": concepto, "base_imponible": base, "porcentaje_iva": pct, "cuota_iva": cuota, "total_factura": total, "moneda": "EUR"}
    if lineas:
        d["lineas"] = [{"descripcion": a, "cantidad": q, "unidad": u, "precio_unitario": p, "importe": round(q * p, 2)} for a, q, u, p in lineas]
    return d

def alb(num, fecha, prov, nif, lineas, total, ref):
    return {"tipo_documento": "ALBARAN", "numero_albaran": num, "nombre_proveedor": prov, "NIF_proveedor": nif, "fecha_entrega": fecha,
            "referencia_pedido": ref, "referencia_factura": None,
            "lineas": [{"descripcion": a, "cantidad": q, "unidad": u, "precio_unitario": p, "importe": round(q * p, 2)} for a, q, u, p in lineas], "total_albaran": total}

L_DG1 = [("Pollo entero fresco", 40, "kg", 4.50), ("Aceite oliva virgen 5L", 5, "ud", 32.00), ("Patata agria", 100, "kg", 1.00), ("Huevos L docena", 60, "ud", 3.00), ("Sal gruesa 25kg", 2, "ud", 10.00)]
L_DG2 = [("Merluza fresca", 30, "kg", 9.00), ("Tomate pera", 50, "kg", 1.50), ("Cebolla", 50, "kg", 0.90), ("Limones", 30, "kg", 1.00)]
L_DG2_ALB = [("Merluza fresca", 30, "kg", 9.00), ("Tomate pera", 50, "kg", 1.50), ("Cebolla", 45, "kg", 0.90)]
L_BU = [("Roba de llit (kg)", 120, "kg", 1.80), ("Tovalloles", 60, "kg", 1.50), ("Estovalles restaurant", 20, "ud", 0.20)]

IA = {
    "factura_distribucions_garraf_DG-2026-0812.pdf": fac("DG-2026-0812", "12/08/2026", "Distribucions Garraf SL", "B-61234567", "Suministro alimentación (albarán 5531)", 640.0, 10, 64.0, 704.0, L_DG1),
    "albaran_garraf_5531.pdf": alb("5531", "11/08/2026", "Distribucions Garraf SL", "B-61234567", L_DG1, 640.0, "PED-EP-311"),
    "factura_energia_llevant_EL-88213.pdf": fac("EL-88213", "14/08/2026", "Energia Llevant SA", "A-08765432", "Consumo eléctrico julio 2026 (24.480 kWh)", 980.0, 21, 195.80, 1175.80),
    "foto_factura_bugaderia_sitges.jpg": fac("BS-2026-0417", "18/08/2026", "Bugaderia Sitges SL", "B-63456789", "Servei de bugaderia", 310.0, 21, 65.10, 375.10, L_BU),
    "factura_distribucions_garraf_DG-2026-0819.pdf": fac("DG-2026-0819", "19/08/2026", "Distribucions Garraf SL", "B-61234567", "Suministro alimentación (albarán 5540)", 420.0, 10, 42.0, 462.0, L_DG2),
    "albaran_garraf_5540.pdf": alb("5540", "18/08/2026", "Distribucions Garraf SL", "B-61234567", L_DG2_ALB, 385.50, "PED-EP-318"),
    "factura_neteges_costa_NC-2026-118.pdf": fac("NC-2026-118", "05/08/2026", "Neteges Costa SL", "B-62345678", "Servicio de limpieza zonas comunes, agosto", 240.0, 21, 50.40, 290.40),
    "factura_installacions_vila_IV-0733.pdf": fac("IV-0733", "20/08/2026", "Instal·lacions Vila SL", "B-64567890", "Reparación climatización planta 3", 1250.0, 21, 262.50, 1512.50),
    "factura_neteges_costa_NC-2026-118_copia.pdf": fac("NC-2026-118", "05/08/2026", "Neteges Costa SL", "B-62345678", "Servicio de limpieza zonas comunes, agosto (reemitida)", 260.0, 21, 54.60, 314.60),
    "bono_viatges_mediterrani_VM-7781.pdf": {"tipo_documento": "BONO", "numero_bono": "VM-7781", "agencia": "Viatges Mediterrani SL", "NIF_agencia": "B-67890123", "huesped": "Família Puig Roca", "nombre_hotel": "Hotel Els Pins", "fecha_entrada": "10/08/2026", "fecha_salida": "14/08/2026", "noches": 4, "habitaciones": 2, "regimen": "AD", "precio_noche": 150.0, "importe_total": 1320.0, "moneda": "EUR", "referencia_reserva": None},
    "bono_viatges_mediterrani_VM-7790.pdf": {"tipo_documento": "BONO", "numero_bono": "VM-7790", "agencia": "Viatges Mediterrani SL", "NIF_agencia": "B-67890123", "huesped": "Sr. Marc Ferrer", "nombre_hotel": "Hotel Els Pins", "fecha_entrada": "21/08/2026", "fecha_salida": "24/08/2026", "noches": 3, "habitaciones": 1, "regimen": "AD", "precio_noche": 140.0, "importe_total": 480.0, "moneda": "EUR", "referencia_reserva": None},
    "contrato_booking_tarifas_pactadas.xlsx": {"tipo_documento": "CONTRATO_OTA", "ota": "", "tarifas": [
        {"ota": "Booking.com", "nombre_hotel": "Hotel Els Pins", "porcentaje_pactado": 15.0, "mercado": "España", "vigencia_inicio": None, "vigencia_fin": None},
        {"ota": "Expedia", "nombre_hotel": "Hotel Els Pins", "porcentaje_pactado": 18.0, "mercado": "España", "vigencia_inicio": None, "vigencia_fin": None}]},
    "factura_seguretat_garraf_SG-2026-0455.pdf": fac("SG-2026-0455", "08/08/2026", "Seguretat Garraf SL", "B-65678901", "Vigilancia nocturna agosto", 350.0, 21, 73.50, 423.50),
    "factura_assegurances_mar_AM-2026-9107.pdf": fac("AM-2026-9107", "01/08/2026", "Assegurances Mar SA", "A-58123456", "Prima seguro multirriesgo hotel, 3T 2026", 890.0, 0, 0.0, 890.0),
    "factura_vins_penedes_VP-2291.pdf": fac("VP-2291", "07/08/2026", "Vins Penedès SCCL", "F-43210987", "Vi negre Els Pins 75cl", 500.0, 21, 105.0, 605.0, [("Vi negre Els Pins 75cl", 100, "ud", 5.0)]),
    "factura_fruites_llobregat_FL-1180.pdf": fac("FL-1180", "11/08/2026", "Fruites Llobregat SL", "B-60987654", "Fruita variada", 300.0, 10, 30.0, 330.0, [("Fruita variada", 150, "kg", 2.0)]),
}


def instalar():
    import lector_facturas_ap as L
    import copy
    orig = L.extraer_con_claude

    def fake(texto, nombre_archivo, es_ota=False):
        n = nombre_archivo.split('/')[-1].split('\\')[-1]
        if n in IA:
            print(f'    [IA simulada] {n}')
            return copy.deepcopy(IA[n])
        return orig(texto, nombre_archivo, es_ota=es_ota)
    L.extraer_con_claude = fake
    return fake
