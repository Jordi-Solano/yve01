# -*- coding: utf-8 -*-
"""b95 — lo que el lector de OTA no reconoce pasa al lector general, y la factura de
la agencia de un contrato de grupo no se compara con el TOTAL del contrato.

Visto en produccion el 25 sep 2026:
  1. La factura de comision de la agencia (b88) subida como
     "VM-2026-0412_comision_CG-2026-0917.pdf" iba al lector de OTA por el NOMBRE
     ('comision'); el lector de OTA no encontraba una liquidacion de OTA (codigo 3)
     y el lote la TIRABA ("no se pudo extraer ningun dato de factura OTA — revisar
     manualmente"). Nunca llegaba a AP ni a su contrato. El nombre manda al lector
     de OTA con 'comision', 'commission' y 'ota' (que tambien esta en "nota" y en
     "cuota").
  2. Renombrada, entraba en AP… y el cruce de eventos del lote la comparaba con el
     total del contrato: "⚠ MATCHING … BEO dice 14.410,00 € vs factura 1.210,00 €
     (92% diff)". Falso: su cotejo es contra la comision ESPERADA (AR › Contratos).

Se conduce `/api/procesar_batch_stream` de verdad con los dos lectores mockeados
(en el sandbox no hay clave de la IA; lo que se prueba es lo que hace el LOTE con
lo que dicen los lectores):
  A. codigo 3 del lector de OTA → pasa al lector general → ✓ AP, marcada AP_OK, y se
     une a SU contrato (CUADRA, por referencia) en /api/ar/contratos.
  B. la factura de la agencia NO sale en el cruce de eventos ("MATCHING … diff") y
     SI sale la linea que dice donde se coteja.
  C. una liquidacion de OTA de verdad (codigo 0) sigue por el lector de OTA y el
     lector general NO se llama.
  D. la factura de otro proveedor del mismo evento sigue pasando por el cruce de
     eventos (b95 no lo apaga para las demas).

  python3.12 tests/test_ota_vacio_al_general.py
  python3.12 tests/test_ota_vacio_al_general.py --sabotaje   (cada pieza quitada POR SEPARADO:
                                                             las dos tienen que hacer fallar algo)
"""
import json
import os
import shutil
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
os.environ['YVE_BACKUP_HORA'] = 'off'

TENANT = "test-ota-vacio"
H_A = "HOTAVAC"
SABOTAJE = "--sabotaje" in sys.argv

FICHERO_COMISION = "VM-2026-0412_comision_CG-2026-0917.pdf"
FICHERO_OTA = "booking_liquidacion_agosto.pdf"
FICHERO_FLORES = "floristeria_garbi_0915.pdf"


def _contrato():
    return {"es_contrato_grupo": True, "evento": {"id": "EV-77", "nombre": "Congreso Cardio 2026"},
            "contrato_numero": "CG-2026-0917",
            "cliente": {"nombre": "Laboratorios Norte S.A.", "cif": "A08000001"},
            "agencia": {"nombre": "Viajes Meridiano S.L.", "cif": "B28004556"},
            "alojamiento": {"fecha_entrada": "2026-10-05", "fecha_salida": "2026-10-08", "noches": 3,
                            "habitaciones": 20, "total_habitaciones": 11000, "iva_pct": 10},
            "fb": {"total": 2200}, "salas": {"total": 1210},
            "comisiones": {"modo": "porcentaje", "alojamiento_pct": 10},
            "facturacion": {"pagador": "agencia"}}


def _reg(num, prov, nif, concepto, base, archivo):
    return {"archivo": archivo, "numero_factura": num, "fecha_factura": "24/09/2026", "fecha": "24/09/2026",
            "nombre_proveedor": prov, "NIF_proveedor": nif, "descripcion_concepto": concepto,
            "base_imponible": base, "porcentaje_iva": 21.0, "cuota_iva": round(base * .21, 2),
            "total_factura": round(base * 1.21, 2), "tipo_proveedor": "OTRAS"}


def main(sabotaje=None):
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    os.environ["YVE_TENANT"] = TENANT
    os.environ["YVE_HOTEL"] = H_A
    import lector_facturas_ap as LFA
    import dashboard as D
    import contratos_grupo as CG
    import lector_contratos_grupo as L
    from tenant_dirs import datos_dir, entrada_dir, procesadas_dir, reportes_dir

    originales = (D._ota_vacio_al_lector_general, D._contrato_de_agencia)
    if sabotaje == "ota":
        D._ota_vacio_al_lector_general = lambda rc: False     # el codigo 3 vuelve a la papelera
    elif sabotaje == "cruce":
        D._contrato_de_agencia = lambda fila: ''              # la comision vuelve al cruce de eventos

    for d in (datos_dir(), entrada_dir(), procesadas_dir(), reportes_dir()):
        if os.path.isdir(d):
            shutil.rmtree(d)
        os.makedirs(d, exist_ok=True)
    try:
        json.dump([{"id": H_A, "nombre": "Hotel Ota Vacio", "activo": True}],
                  open(os.path.join(datos_dir(), "hoteles.json"), "w"))
        for f in ("proveedores.xlsx", "plan_cuentas.xlsx"):
            src = os.path.join(BASE, "datos-referencia", f)
            if os.path.exists(src):
                shutil.copy2(src, os.path.join(datos_dir(), f))
        dt = _contrato()
        CG.registrar(dt, L.transformar(dt, hotel_id=H_A), hotel_id=H_A, datos_dir=datos_dir())
        # la referencia del evento que deja el lector de contratos (el cruce de eventos la usa)
        json.dump([{"evento": "Congreso Cardio 2026", "cliente": "Laboratorios Norte S.A.",
                    "documentos": {"BEO": {"total": 14410.0}}}],
                  open(os.path.join(datos_dir(), "eventos_referencia.json"), "w"))
        for f in (FICHERO_COMISION, FICHERO_OTA, FICHERO_FLORES):
            open(os.path.join(entrada_dir(), f), "wb").write(b"%PDF-1.4\n% mock\n")

        llamadas_general = []
        regs = {
            FICHERO_COMISION: _reg("VM-2026-0412", "Viajes Meridiano S.L.", "B28004556",
                                   "Comisión 10 % alojamiento — contrato CG-2026-0917 · Congreso Cardio 2026",
                                   1000.0, FICHERO_COMISION),
            FICHERO_FLORES: _reg("FG-0915", "Floristeria Garbi S.L.", "B61234567",
                                 "Decoración floral Congreso Cardio 2026", 800.0, FICHERO_FLORES),
        }

        def _general(fpath, proveedores=None, *a, **k):
            n = os.path.basename(fpath)
            llamadas_general.append(n)
            return dict(regs[n]) if n in regs else {"_skip": True, "_motivo": "mock: no es factura"}
        LFA.procesar_factura_ap = _general

        _run_real = subprocess.run
        llamadas_ota = []

        def _run(cmd, *a, **k):
            if isinstance(cmd, (list, tuple)) and any(str(x).endswith("lector_ota.py") for x in cmd):
                f = os.path.basename(str(cmd[-1]))
                llamadas_ota.append(f)
                if f == FICHERO_OTA:           # una liquidacion de OTA de verdad
                    return subprocess.CompletedProcess(cmd, 0, stdout="OK: procesado\n", stderr="")
                return subprocess.CompletedProcess(cmd, 3, stdout=f"FALTAN: ota, importe_bruto\nVACIO: {f} sin datos OTA extraibles\n", stderr="")
            return _run_real(cmd, *a, **k)
        subprocess.run = _run

        D.app.config["TESTING"] = True
        c = D.app.test_client()
        c.post("/api/login", json={"username": "admin", "password": "admin123"})
        with c.session_transaction() as s:
            s["tenant_id"] = TENANT
            s["hotel_activo"] = H_A

        def lote(ficheros):
            r = c.get('/api/procesar_batch_stream?archivos=' + json.dumps(ficheros))
            return [l[5:].strip() for l in r.get_data(as_text=True).splitlines() if l.startswith("data:")]

        # A + B · la factura de comision de la agencia llamada "...comision..."
        log = lote([FICHERO_COMISION])
        txt = "\n".join(log)
        ok(FICHERO_COMISION in llamadas_ota, "el nombre ('comision') la sigue mandando primero al lector de OTA")
        ok("no se pudo extraer ningún dato de factura OTA" not in txt,
           "el codigo 3 del lector de OTA ya NO la tira a la papelera")
        ok(FICHERO_COMISION in llamadas_general and any(l.startswith("✓ AP") for l in log),
           f"pasa al lector general y entra en AP ({[l for l in log if FICHERO_COMISION in l][:3]})")
        ok(any("no es una liquidación de OTA" in l for l in log), "el log dice por que va al lector general")
        est = c.get("/api/archivos_estado").get_json() or {}
        marca = next((f.get("resultado") for f in est.get("files", []) if f.get("nombre") == FICHERO_COMISION), None)
        ok(marca == "AP_OK", f"marcada AP_OK en el registro del lote ({marca})")
        ctr = c.get("/api/ar/contratos").get_json() or {}
        v = next((x for x in ctr.get("contratos", []) if "0917" in str(x.get("contrato") or x.get("id"))), {})
        fcv = v.get("factura_comision_vista") or {}
        ok(fcv.get("numero") == "VM-2026-0412" and fcv.get("estado") == "CUADRA" and fcv.get("origen") == "referencia",
           f"y se une a SU contrato: CUADRA por referencia ({fcv.get('numero')}, {fcv.get('estado')}, {fcv.get('origen')})")
        ok(not any("MATCHING" in l and "diff" in l for l in log),
           "la comision de la agencia NO se compara con el total del contrato (el '92% diff' falso)")
        ok(any("factura de la agencia del contrato de grupo CG-2026-0917" in l and "AR › Contratos" in l for l in log),
           "y el log dice donde se coteja: con la comision esperada en AR › Contratos")

        # C · una liquidacion de OTA de verdad sigue su camino
        antes = len(llamadas_general)
        log2 = lote([FICHERO_OTA])
        ok(any(l.startswith("✓ AR") and FICHERO_OTA in l for l in log2) and len(llamadas_general) == antes,
           f"una liquidacion de OTA (codigo 0) sigue en el lector de OTA y el general no se llama ({log2[2:3]})")

        # D · otro proveedor del mismo evento: el cruce de eventos sigue para el
        log3 = lote([FICHERO_FLORES])
        ok(any("MATCHING" in l for l in log3) and not any("factura de la agencia" in l for l in log3),
           "la factura de otro proveedor del evento sigue pasando por el cruce de eventos")
    finally:
        D._ota_vacio_al_lector_general, D._contrato_de_agencia = originales
        try:
            subprocess.run = _run_real
        except Exception:
            pass
        for d in (datos_dir(), entrada_dir(), procesadas_dir(), reportes_dir()):
            if os.path.isdir(d):
                shutil.rmtree(d)

    print("=" * 62)
    return fallos


if __name__ == "__main__":
    if SABOTAJE:
        malos = 0
        for pieza in ("ota", "cruce"):
            print(f"\nSABOTAJE '{pieza}'")
            n = main(pieza)
            if n:
                print(f"  SABOTAJE '{pieza}': {n} comprobacion(es) fallan, como debe ser")
            else:
                print(f"  SABOTAJE '{pieza}': quitada la pieza NO falla nada — el test no la protege")
                malos += 1
        raise SystemExit(1 if malos else 0)
    n = main()
    print("  todo OK" if not n else f"  {n} FALLO(S)")
    raise SystemExit(1 if n else 0)
