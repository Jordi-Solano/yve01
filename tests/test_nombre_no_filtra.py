# -*- coding: utf-8 -*-
"""b100 — el NOMBRE del fichero ya no descarta nada: todo se clasifica por el contenido.

Decision de Jordi (25 sep 2026). Antes `lector_facturas_ap.es_no_factura_por_nombre`
saltaba SIN LEER lo que se llamara 'programa', 'presupuesto', 'agenda', 'quote',
'menu ', 'minuta', 'planning'... Visto en produccion: el contrato de grupo
"contrato_grupo_CG-2026-0940_programa.pdf" salio "⚠ nombre: contiene 'programa'" y
no llego nunca a AR › Contratos. Solo siguen fuera los FORMATOS que no se pueden leer
(Word, PowerPoint, video, comprimidos): ahi no hay contenido que clasificar.

Se conduce `/api/procesar_batch_stream` de verdad; se mockean la extraccion de texto y
las dos llamadas a la IA (el clasificador y el lector de contratos), porque lo que se
prueba es que el LOTE los llame, no lo que responda la IA:
  A. "contrato_grupo_CG-2026-0940_programa.pdf" se lee, la IA dice CONTRATO y acaba en
     AR › Contratos con su programa (2 BEO, una por dia).
  B. "Presupuesto aceptado Floristeria Garbi.pdf" se lee y, como es una factura, entra en AP.
  C. un nombre "de los de antes" que NO es factura ("Menu de gala.pdf") tambien se lee: lo
     descarta el CLASIFICADOR por lo que dice, no el nombre.
  D. "Programa del evento.docx" sigue fuera por el FORMATO (no hay texto que leer).
  E. la lista de palabras ya no existe y el prompt del clasificador no se ha tocado.

  python3.12 tests/test_nombre_no_filtra.py
  python3.12 tests/test_nombre_no_filtra.py --sabotaje   (vuelve el filtro por nombre: tiene que fallar)
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

SABOTAJE = "--sabotaje" in sys.argv
TENANT = "test-nombre-no-filtra"
H_A = "HNOMBRE"

F_CONTRATO = "contrato_grupo_CG-2026-0940_programa.pdf"
F_PRESUPUESTO = "Presupuesto aceptado Floristeria Garbi.pdf"
F_MENU = "Menu de gala.pdf"
F_DOCX = "Programa del evento.docx"


def contrato_0940():
    return {"es_contrato_grupo": True, "evento": {"id": "EV-88", "nombre": "Jornadas de Oncologia 2026"},
            "contrato_numero": "CG-2026-0940",
            "cliente": {"nombre": "Farmaceutica Levante S.A.", "cif": "A46000003"},
            "agencia": {"nombre": "Eventos Costa DMC S.L.", "cif": "B66123456"},
            "alojamiento": {"fecha_entrada": "2026-11-12", "fecha_salida": "2026-11-14", "noches": 2,
                            "habitaciones": 30, "total_habitaciones": 9900, "iva_pct": 10},
            "fb": {"total": 3240, "iva_pct": 10}, "salas": {"total": 2400, "nombre": "Mediterrània", "montaje": "teatro", "dias": 2},
            "comisiones": {"modo": "neta", "texto": "Tarifas netas, no comisionables."},
            "facturacion": {"pagador": "agencia", "texto": "La factura del grupo se emite a nombre de la agencia."},
            "beo": {"funciones": [
                {"fecha": "2026-11-12", "hora_inicio": "09:00", "hora_fin": "14:00", "funcion": "Sesión plenaria",
                 "sala": "Mediterrània", "montaje": "teatro", "pax": 60, "alquiler": 1200},
                {"fecha": "2026-11-12", "hora_inicio": "14:00", "hora_fin": "15:30", "funcion": "Almuerzo buffet",
                 "sala": "Restaurante", "montaje": "buffet", "pax": 60, "precio_pp": 30},
                {"fecha": "2026-11-13", "hora_inicio": "09:00", "hora_fin": "14:00", "funcion": "Talleres",
                 "sala": "Mediterrània", "montaje": "teatro", "pax": 60, "alquiler": 1200}]}}


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    os.environ["YVE_TENANT"] = TENANT
    os.environ["YVE_HOTEL"] = H_A
    import lector_facturas_ap as LFA
    import lector_contratos_grupo as LCG
    import dashboard as D
    from tenant_dirs import datos_dir, entrada_dir, procesadas_dir, reportes_dir

    if SABOTAJE:
        viejas = ('programa', 'presupuesto', 'menu ', 'agenda', 'quote', 'minuta')

        def _filtro_de_antes(nombre):
            n = nombre.lower()
            if os.path.splitext(n)[1] in LFA.NO_FACTURA_EXTENSIONS:
                return True, "extensión"
            for kw in viejas:
                if kw in n:
                    return True, f"contiene '{kw}'"
            return False, ""
        LFA.es_no_factura_por_nombre = _filtro_de_antes

    leidos, contratos_leidos = [], []
    LFA.extraer_texto = lambda ruta: "Documento de prueba con texto suficiente para clasificarlo por su contenido."

    def _clasificador(texto, nombre, es_ota=False):
        leidos.append(nombre)
        if nombre == F_CONTRATO:
            return {"tipo_documento": "CONTRATO"}
        if nombre == F_PRESUPUESTO:
            return {"nombre_proveedor": "Floristeria Garbi S.L.", "NIF_proveedor": "B61234567", "numero_factura": "FG-0925",
                    "fecha_factura": "25/09/2026", "base_imponible": 800.0, "porcentaje_iva": 21.0, "cuota_iva": 168.0,
                    "total_factura": 968.0, "descripcion_concepto": "Decoración floral según presupuesto aceptado"}
        return {"_skip": True, "_motivo": "no es una factura: es una carta de menú"}
    LFA.extraer_con_claude = _clasificador

    def _lector_contratos(rutas):
        contratos_leidos.extend(os.path.basename(str(r)) for r in rutas)
        return contrato_0940()
    LCG.extraer_contrato_grupo = _lector_contratos

    for d in (datos_dir(), entrada_dir(), procesadas_dir(), reportes_dir()):
        if os.path.isdir(d):
            shutil.rmtree(d)
        os.makedirs(d, exist_ok=True)
    try:
        json.dump([{"id": H_A, "nombre": "Hotel Nombre", "activo": True}], open(os.path.join(datos_dir(), "hoteles.json"), "w"))
        for f in ("proveedores.xlsx", "plan_cuentas.xlsx"):
            src = os.path.join(BASE, "datos-referencia", f)
            if os.path.exists(src):
                shutil.copy2(src, os.path.join(datos_dir(), f))
        for f in (F_CONTRATO, F_PRESUPUESTO, F_MENU):
            open(os.path.join(entrada_dir(), f), "wb").write(b"%PDF-1.4\n% mock\n")
        open(os.path.join(entrada_dir(), F_DOCX), "wb").write(b"PK\x03\x04 mock docx")

        D.app.config["TESTING"] = True
        c = D.app.test_client()
        c.post("/api/login", json={"username": "admin", "password": "admin123"})
        with c.session_transaction() as s:
            s["tenant_id"] = TENANT
            s["hotel_activo"] = H_A
        r = c.get('/api/procesar_batch_stream?archivos=' + json.dumps([F_CONTRATO, F_PRESUPUESTO, F_MENU, F_DOCX]))
        log = [l[5:].strip() for l in r.get_data(as_text=True).splitlines() if l.startswith("data:")]
        txt = "\n".join(log)

        # A · el contrato llamado "..._programa.pdf"
        ctr = (c.get("/api/ar/contratos").get_json() or {}).get("contratos") or []
        v = next((x for x in ctr if x.get("contrato") == "CG-2026-0940"), {})
        ok(F_CONTRATO in leidos and F_CONTRATO in contratos_leidos and "contiene 'programa'" not in txt,
           f"A · '{F_CONTRATO}' se lee (clasificador y lector de contratos), no se salta por el nombre")
        ok(bool(v) and (v.get("beo") or {}).get("n") == 2 and any("Contrato CG-2026-0940" in l for l in log),
           f"A · y acaba en AR › Contratos con su programa: 2 BEO ({(v.get('beo') or {}).get('n')})")

        # B · un "presupuesto" que es una factura
        ok(F_PRESUPUESTO in leidos and any(l.startswith("✓ AP") and F_PRESUPUESTO in l for l in log),
           f"B · '{F_PRESUPUESTO}' se lee y entra en AP ({[l for l in log if F_PRESUPUESTO in l][:2]})")

        # C · un nombre "de los de antes" que no es factura: lo descarta el CONTENIDO
        ok(F_MENU in leidos and any(F_MENU in l and "carta de menú" in l for l in log),
           f"C · '{F_MENU}' se lee y lo descarta el clasificador por lo que dice ({[l for l in log if F_MENU in l][:2]})")

        # D · el formato que no se puede leer sigue fuera
        ok(F_DOCX not in leidos and any(F_DOCX in l and "extensión .docx" in l for l in log),
           f"D · '{F_DOCX}' sigue fuera por el formato ({[l for l in log if F_DOCX in l][:1]})")
    finally:
        for d in (datos_dir(), entrada_dir(), procesadas_dir(), reportes_dir()):
            if os.path.isdir(d):
                shutil.rmtree(d)

    # E · la lista ya no existe y el prompt del clasificador no se toca
    fuente = open(os.path.join(BASE, "lector_facturas_ap.py"), encoding="utf-8").read()
    ok("NO_FACTURA_KEYWORDS" not in fuente
       and all(LFA.es_no_factura_por_nombre(n) == (False, "") for n in
               ("Presupuesto grupo X.pdf", "Programa evento.pdf", "agenda.jpg", "quote 12.pdf", "Menu cena.pdf")),
       "E · la lista de palabras que descartaba por el nombre ya no esta (y ningun nombre descarta)")
    import hashlib
    i = fuente.find('PROMPT_CLASIFICACION = """')
    huella = hashlib.md5(fuente[i:fuente.find('"""', i + 30)].encode("utf-8")).hexdigest() if i >= 0 else ""
    # la huella del prompt tal como estaba antes de b100 (zona intocable: cambiarlo pide el OK de Jordi)
    ok(huella == "2c74e25d7944e774bc69bced11f447bf", f"E · el prompt del clasificador (PROMPT_CLASIFICACION) no se ha tocado ({huella[:8]})")
    diff = subprocess.run(['git', 'diff', '--name-only', 'HEAD'], capture_output=True, text=True, cwd=BASE).stdout.split()
    ok(not [f for f in diff if f.startswith('oracle_')], 'la zona oracle_* no se toca')

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
