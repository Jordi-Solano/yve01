"""
lector_contratos_grupo.py — Yve.01
Procesa CONTRATOS DE GRUPO / EVENTOS (Hilton BEO + contrato de grupos) a partir
de fotos/escaneos de varias páginas, y los convierte en:
  · Cliente corporativo (AR Real)
  · Factura/receivable (habitaciones + F&B + salas)  -> reservas_credito.xlsx
  · Registro de COMISIÓN que el hotel paga a la agencia (base sin IVA × %)
  · Alerta de certificado de doble imposición (cliente extranjero)

Extracción: API de visión de Claude (multipágina). Si no hay API/red,
degrada de forma elegante marcando el contrato como "pendiente de revisión"
(nunca lanza excepción hacia arriba).
"""
import os, json, base64, glob, mimetypes
from datetime import datetime

MODEL = "claude-sonnet-4-6"

# Esquema que pedimos a la visión (una sola pasada con todas las páginas)
_PROMPT = """Eres un experto en contratación hotelera. Te paso las fotos de TODAS las páginas
de un contrato de grupo/eventos de hotel (contrato + BEO + anexos). Devuelve SOLO un JSON válido
(sin texto alrededor) con esta estructura exacta (usa null si un dato no aparece):
{
 "es_contrato_grupo": true,
 "evento": {"id": "", "nombre": ""},
 "contrato_numero": "",
 "fecha_contrato": "YYYY-MM-DD",
 "hotel": {"nombre": "", "cif": ""},
 "cliente": {"nombre": "", "cif": "", "email": "", "contacto": "", "pais": ""},
 "agencia": {"nombre": "", "contacto": "", "email": "", "telefono": ""},
 "alojamiento": {"fecha_entrada":"YYYY-MM-DD","fecha_salida":"YYYY-MM-DD","noches":0,
                 "habitaciones":0,"tarifa_dui":0,"tarifa_doble":0,
                 "total_habitaciones":0,"iva_pct":10,"total_pernoctaciones":0},
 "tasa_turistica": {"por_persona_noche":0,"max_noches":0},
 "fb": {"total":0,"por_persona_dia":0,"pax":0,"dias":0,"iva_pct":10,"detalle":""},
 "salas": {"total":0},
 "comisiones": {"alojamiento_pct":0,"salas_pct":0,"fb_pct":0,"ddr_pct":0,"misc_pct":0},
 "deposito": {"pct":0,"cuando":"","iban":"","beneficiario":"","referencia":""},
 "doble_imposicion": false,
 "beos": []
}
Importante: los importes son numéricos (sin símbolo €, punto decimal). Las fechas en formato ISO.
"doble_imposicion" = true si el cliente es extranjero o el contrato menciona doble imposición / withholding / certificado de residencia fiscal.
"es_contrato_grupo" = true SOLO si es un contrato de grupo/eventos de hotel o un BEO (orden de servicio); false si son facturas sueltas, extractos u otro documento."""


def _api_key():
    return os.environ.get("ANTHROPIC_API_KEY", "") or _env_file_key()


def _env_file_key():
    try:
        p = os.path.join(os.path.dirname(__file__), ".env")
        if os.path.exists(p):
            for ln in open(p, encoding="utf-8"):
                if ln.strip().startswith("ANTHROPIC_API_KEY"):
                    return ln.split("=", 1)[1].strip().strip('"').strip("'")
    except Exception:
        pass
    return ""


def _img_block(path):
    mt = mimetypes.guess_type(path)[0] or "image/jpeg"
    with open(path, "rb") as f:
        data = base64.standard_b64encode(f.read()).decode()
    return {"type": "image", "source": {"type": "base64", "media_type": mt, "data": data}}


def extraer_contrato_grupo(image_paths):
    """Extrae los datos del contrato con la visión de Claude. Nunca lanza; si falla
    devuelve {'_needs_review': True, '_error': ...}."""
    key = _api_key()
    if not key:
        return {"_needs_review": True, "_error": "sin ANTHROPIC_API_KEY (extracción no disponible en este entorno)"}
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=key)
        content = [{"type": "text", "text": _PROMPT}]
        for p in image_paths[:30]:
            content.append(_img_block(p))
        resp = client.messages.create(model=MODEL, max_tokens=1500,
                                      messages=[{"role": "user", "content": content}])
        txt = resp.content[0].text.strip()
        # aislar el JSON
        i, j = txt.find("{"), txt.rfind("}")
        datos = json.loads(txt[i:j + 1])
        datos["_needs_review"] = False
        return datos
    except Exception as e:
        return {"_needs_review": True, "_error": str(e)[:200]}


def _f(v, d=0.0):
    try:
        if v in (None, "", "null"):
            return d
        return float(str(v).replace("€", "").replace(",", "").strip())
    except Exception:
        return d


def calcular_comisiones(datos):
    """Comisión que el hotel paga a la agencia = base SIN IVA × % por concepto."""
    def base(total, iva):
        total = _f(total); iva = _f(iva, 10)
        return total / (1 + iva / 100) if total else 0.0
    aloj = datos.get("alojamiento", {}) or {}
    fb = datos.get("fb", {}) or {}
    salas = datos.get("salas", {}) or {}
    com = datos.get("comisiones", {}) or {}
    base_aloj = base(aloj.get("total_habitaciones"), aloj.get("iva_pct", 10))
    base_fb = base(fb.get("total"), fb.get("iva_pct", 10))
    base_salas = base(salas.get("total"), 21)
    c_aloj = round(base_aloj * _f(com.get("alojamiento_pct")) / 100, 2)
    c_fb = round(base_fb * _f(com.get("fb_pct")) / 100, 2)
    c_salas = round(base_salas * _f(com.get("salas_pct")) / 100, 2)
    return {
        "alojamiento": c_aloj, "fb": c_fb, "salas": c_salas,
        "total": round(c_aloj + c_fb + c_salas, 2),
        "base_alojamiento": round(base_aloj, 2), "base_fb": round(base_fb, 2),
    }


def transformar(datos):
    """Convierte los datos extraídos en filas para AR Real + comisión + flag DI.
    Testeable sin API."""
    ev = datos.get("evento", {}) or {}
    cli = datos.get("cliente", {}) or {}
    aloj = datos.get("alojamiento", {}) or {}
    fb = datos.get("fb", {}) or {}
    salas = datos.get("salas", {}) or {}
    contrato = str(datos.get("contrato_numero") or "").strip()

    imp_hab = _f(aloj.get("total_habitaciones"))
    imp_fb = _f(fb.get("total"))
    imp_salas = _f(salas.get("total"))
    total = round(imp_hab + imp_fb + imp_salas, 2)
    comis = calcular_comisiones(datos)
    di = bool(datos.get("doble_imposicion")) or (str(cli.get("pais", "")).lower() not in ("", "españa", "espana", "spain"))

    nombre_cli = (cli.get("nombre") or "Cliente grupo").strip()
    cliente_row = {
        "nombre_cliente": nombre_cli, "NIF": cli.get("cif") or "", "nif": cli.get("cif") or "",
        "email": cli.get("email") or "", "telefono": (datos.get("agencia", {}) or {}).get("telefono", "") or "",
        "dias_pago": 30, "limite_credito": 100000, "credito_limite": 100000, "credito_usado": 0,
    }
    numero = f"GRP-{contrato}" if contrato else f"GRP-{datetime.now().strftime('%Y%m%d%H%M')}"
    noches = int(_f(aloj.get("noches"))) or 0
    reserva_row = {
        "numero_reserva": numero, "numero": numero, "cliente": nombre_cli,
        "fecha_entrada": aloj.get("fecha_entrada") or "", "fecha_salida": aloj.get("fecha_salida") or "",
        "fecha_emision": "", "habitaciones": int(_f(aloj.get("habitaciones"))), "noches": noches,
        "importe_habitaciones": imp_hab, "importe_fb": imp_fb, "importe_extras": imp_salas,
        "importe": total, "total": total, "estado": "PENDIENTE_FACTURA",
        "evento": (str(ev.get("id") or "") + " " + str(ev.get("nombre") or "")).strip(),
        "contrato": contrato, "comision_total": comis["total"],
        "requiere_certificado_di": di, "tipo": "CONTRATO_GRUPO",
    }
    return {
        "cliente": cliente_row, "reserva": reserva_row, "comisiones": comis,
        "doble_imposicion": di,
        "resumen": {
            "evento": reserva_row["evento"], "contrato": contrato, "cliente": nombre_cli,
            "total_receivable": total, "habitaciones": imp_hab, "fb": imp_fb, "salas": imp_salas,
            "comision_total": comis["total"], "requiere_certificado_di": di,
            "numero": numero,
        },
    }


def generar_beo(datos, transformado=None):
    """Genera un BEO (orden de servicio con partidas e importes) a partir de los
    datos del contrato. Determinista y testeable sin API. El total coincide con el
    receivable del contrato, para poder cotejar la factura contra el BEO."""
    ev = datos.get("evento", {}) or {}
    cli = datos.get("cliente", {}) or {}
    aloj = datos.get("alojamiento", {}) or {}
    fb = datos.get("fb", {}) or {}
    salas = datos.get("salas", {}) or {}
    tasa = datos.get("tasa_turistica", {}) or {}
    contrato = str(datos.get("contrato_numero") or "").strip()

    lineas = []
    imp_hab = _f(aloj.get("total_habitaciones"))
    if imp_hab:
        habs = int(_f(aloj.get("habitaciones"))); noches = int(_f(aloj.get("noches")))
        det = (f"{habs} hab × {noches} noches" if habs and noches else "Alojamiento del grupo")
        lineas.append({"concepto": "Alojamiento", "detalle": det,
                       "iva_pct": _f(aloj.get("iva_pct"), 10), "importe": round(imp_hab, 2)})
    imp_fb = _f(fb.get("total"))
    if imp_fb:
        pax = int(_f(fb.get("pax"))); dias = int(_f(fb.get("dias"))); ppd = _f(fb.get("por_persona_dia"))
        det = fb.get("detalle") or (f"{pax} pax × {dias} días × €{ppd:.2f}" if pax and dias else "Comidas y bebidas")
        lineas.append({"concepto": "F&B (comidas y bebidas)", "detalle": det,
                       "iva_pct": _f(fb.get("iva_pct"), 10), "importe": round(imp_fb, 2)})
    imp_salas = _f(salas.get("total"))
    if imp_salas:
        lineas.append({"concepto": "Salas / Meeting", "detalle": "Alquiler de salas y montaje",
                       "iva_pct": 21, "importe": round(imp_salas, 2)})
    tpn = _f(tasa.get("por_persona_noche"))
    if tpn:
        lineas.append({"concepto": "Tasa turística", "detalle": f"€{tpn:.2f}/persona/noche (según ocupación real)",
                       "iva_pct": 0, "importe": 0.0})

    total = round(sum(_f(l.get("importe")) for l in lineas), 2)
    evento_nombre = (str(ev.get("id") or "") + " " + str(ev.get("nombre") or "")).strip() or "Evento de grupo"
    numero = (transformado or {}).get("reserva", {}).get("numero_reserva", "")
    return {
        "tipo_documento": "BEO",
        "generado_desde_contrato": True,
        "evento": evento_nombre,
        "cliente": (cli.get("nombre") or "Cliente grupo").strip(),
        "contrato": contrato,
        "numero_reserva": numero,
        "fecha_entrada": aloj.get("fecha_entrada") or "",
        "fecha_salida": aloj.get("fecha_salida") or "",
        "pax": int(_f(fb.get("pax"))) or None,
        "fecha_generado": datetime.now().strftime("%Y-%m-%d"),
        "lineas": lineas,
        "items": [{"concepto": l["concepto"], "total": l["importe"]} for l in lineas],
        "total": total,
        "total_estimado": total,
    }


def guardar_beo(beo, datos_dir=None):
    """Guarda el BEO generado: (1) como referencia del evento para el 3-way matching
    (eventos_referencia.json) y (2) en beos_generados.json para verlo en AR Real."""
    dd = datos_dir or _datos_dir()
    os.makedirs(dd, exist_ok=True)
    evento = beo.get("evento", "") or ""
    evento_key = evento.lower().strip()[:50]
    doc = {"archivo": "BEO (generado del contrato)", "total": beo.get("total", 0),
           "items": beo.get("items", []), "fecha": beo.get("fecha_generado", ""),
           "generado": True, "lineas": beo.get("lineas", [])}
    # (1) referencia para matching
    ref_path = os.path.join(dd, "eventos_referencia.json")
    try:
        refs = json.load(open(ref_path, encoding="utf-8")) if os.path.exists(ref_path) else []
    except Exception:
        refs = []
    found = False
    for ref in refs:
        if ref.get("evento_key") == evento_key:
            ref.setdefault("documentos", {})["BEO"] = doc
            found = True
            break
    if not found:
        refs.append({"evento": evento, "evento_key": evento_key,
                     "cliente": beo.get("cliente", ""), "documentos": {"BEO": doc}})
    json.dump(refs, open(ref_path, "w", encoding="utf-8"), indent=2, ensure_ascii=False)
    # (2) store para la vista
    beos_path = os.path.join(dd, "beos_generados.json")
    try:
        beos = json.load(open(beos_path, encoding="utf-8")) if os.path.exists(beos_path) else []
    except Exception:
        beos = []
    # El hotel del evento (fase 5). Se estampa aqui, despues de leer el
    # contrato: lo decide la sesion, no el papel.
    try:
        import censo_hoteles as _censo
        beo["hotel_id"] = _censo.para_guardar()
    except Exception:
        beo["hotel_id"] = os.environ.get("YVE_HOTEL", "")
    # La identidad del BEO incluye el hotel: dos hoteles del grupo pueden tener
    # el mismo evento con el mismo numero de contrato y no son el mismo BEO.
    beos = [b for b in beos if not (b.get("evento", "").lower().strip()[:50] == evento_key
                                    and str(b.get("contrato")) == str(beo.get("contrato"))
                                    and str(b.get("hotel_id") or "") == str(beo.get("hotel_id") or ""))]
    beos.append(beo)
    json.dump(beos, open(beos_path, "w", encoding="utf-8"), indent=2, ensure_ascii=False)
    return {"referencia": ref_path, "beos": beos_path}


def _datos_dir():
    try:
        from tenant_dirs import datos_dir
        return datos_dir()
    except Exception:
        return os.path.join(os.path.dirname(__file__), "datos-referencia")


def guardar(transformado, datos_dir=None):
    """Añade/actualiza cliente y reserva en los xlsx de AR Real. Dedup por nombre/numero."""
    import pandas as pd
    dd = datos_dir or _datos_dir()
    os.makedirs(dd, exist_ok=True)
    # Clientes
    pc = os.path.join(dd, "clientes_credito.xlsx")
    dfc = pd.read_excel(pc) if os.path.exists(pc) else pd.DataFrame()
    nom = transformado["cliente"]["nombre_cliente"]
    if not (len(dfc) and (dfc.get("nombre_cliente", pd.Series(dtype=str)).astype(str) == nom).any()):
        dfc = pd.concat([dfc, pd.DataFrame([transformado["cliente"]])], ignore_index=True)
        dfc.to_excel(pc, index=False)
    # Reservas / facturas
    pr = os.path.join(dd, "reservas_credito.xlsx")
    dfr = pd.read_excel(pr) if os.path.exists(pr) else pd.DataFrame()
    num = transformado["reserva"]["numero_reserva"]
    col = "numero_reserva" if "numero_reserva" in dfr.columns else "numero"
    if len(dfr) and col in dfr.columns:
        dfr = dfr[dfr[col].astype(str) != num]  # reemplaza si ya existía
    dfr = pd.concat([dfr, pd.DataFrame([transformado["reserva"]])], ignore_index=True)
    dfr.to_excel(pr, index=False)
    return {"clientes": pc, "reservas": pr}


def _append_xlsx(path, row, dedup_col=None):
    """Añade una fila a un xlsx (creándolo si no existe). Si dedup_col, reemplaza
    la fila con el mismo valor en esa columna (idempotente al reprocesar)."""
    import pandas as pd
    df = pd.read_excel(path) if os.path.exists(path) else pd.DataFrame()
    if dedup_col and len(df) and dedup_col in df.columns and row.get(dedup_col) is not None:
        df = df[df[dedup_col].astype(str) != str(row.get(dedup_col))]
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df.to_excel(path, index=False)


def distribuir_contrato(datos, transformado, datos_dir=None):
    """Reparte los importes REALES del contrato a los módulos donde tienen sentido:
      · AP    -> comisión que el hotel paga a la agencia (pago pendiente)
      · Banco -> depósito/anticipo que el cliente adelanta (cobro previsto)
      · F&B   -> catering/banquete del evento (ingreso del evento)
    Devuelve {'ap':importe|None,'banco':importe|None,'fb':importe|None} para el log/badges."""
    dd = datos_dir or _datos_dir()
    os.makedirs(dd, exist_ok=True)
    res = {"ap": None, "banco": None, "fb": None}
    ev = datos.get("evento", {}) or {}
    evento_nombre = (str(ev.get("id") or "") + " " + str(ev.get("nombre") or "")).strip() or "Evento de grupo"
    contrato = str(datos.get("contrato_numero") or "").strip()
    hoy = datetime.now().strftime("%Y-%m-%d")

    # ── AP: comisión a la agencia ──
    comis = (transformado or {}).get("comisiones", {}) or {}
    com_total = _f(comis.get("total"))
    if com_total > 0:
        try:
            from tenant_dirs import procesadas_dir
            pdir = procesadas_dir()
        except Exception:
            pdir = dd
        os.makedirs(pdir, exist_ok=True)
        ap_file = os.path.join(pdir, "facturas_ap_" + datetime.now().strftime("%Y%m%d") + ".xlsx")
        ag = datos.get("agencia", {}) or {}
        agencia = ag.get("nombre") or (datos.get("cliente", {}) or {}).get("nombre") or "Agencia"
        base = round(com_total / 1.21, 2)
        num = ("COM-" + contrato) if contrato else ("COM-" + datetime.now().strftime("%Y%m%d%H%M"))
        # BUG 8: esta fila se escribe SIN pasar por `_guardar_factura_ap`, que
        # es quien estampa el hotel. Sin `hotel_id` la comision cae en "sin
        # asignar": no sale en ningun hotel y no se puede aprobar. El bloque de
        # F&B de mas abajo ya lo hacia bien (fase 4b) — aqui faltaba.
        try:
            import censo_hoteles as _censo
            _hid_ap = _censo.para_guardar()
        except Exception:
            _hid_ap = os.environ.get("YVE_HOTEL", "")
        _append_xlsx(ap_file, {
            "archivo": "comision_" + (contrato or evento_nombre),
            "numero_factura": num, "fecha": hoy, "nombre_proveedor": agencia,
            "NIF_proveedor": ag.get("cif", "") or "",
            "descripcion_concepto": "Comisión agencia · " + evento_nombre,
            "base_imponible": base, "porcentaje_iva": 21,
            "cuota_iva": round(com_total - base, 2), "total_factura": round(com_total, 2),
            "moneda": "EUR", "tipo": "COMISION_AGENCIA", "estado_matching": "SIN_PO",
            "hotel_id": _hid_ap,
        }, dedup_col="numero_factura")
        res["ap"] = round(com_total, 2)

    # ── Banco: depósito previsto ──
    dep = datos.get("deposito", {}) or {}
    total_recv = _f((transformado or {}).get("resumen", {}).get("total_receivable"))
    dep_pct = _f(dep.get("pct"))
    dep_imp = round(total_recv * dep_pct / 100, 2) if dep_pct else 0.0
    if dep_imp > 0:
        concepto = "Depósito previsto " + str(int(dep_pct)) + "% · " + evento_nombre + (" (contrato " + contrato + ")" if contrato else "")
        _append_xlsx(os.path.join(dd, "extracto_banco.xlsx"),
                     {"fecha": hoy, "concepto": concepto, "importe": dep_imp, "saldo": ""},
                     dedup_col="concepto")
        res["banco"] = dep_imp

    # ── F&B: catering del evento ──
    fb = datos.get("fb", {}) or {}
    fb_total = _f(fb.get("total"))
    if fb_total > 0:
        pax = int(_f(fb.get("pax"))) or 1
        plato = "Banquete evento · " + evento_nombre
        # El catering de un evento es venta de F&B del hotel donde se celebra
        # (fase 4b). Sin esto, la unica entrada a ventas que NO pasa por el lote
        # se quedaba sin etiqueta y caia en "sin asignar".
        try:
            import censo_hoteles as _censo
            _hid_fb = _censo.para_guardar()
        except Exception:
            _hid_fb = os.environ.get("YVE_HOTEL", "")
        _append_xlsx(os.path.join(dd, "ventas_fb_diarias.xlsx"), {
            "fecha": (datos.get("alojamiento", {}) or {}).get("fecha_entrada") or hoy,
            "nombre_plato": plato, "categoria": "Eventos",
            "unidades_vendidas": pax, "precio_unitario": round(fb_total / pax, 2),
            "total_venta": round(fb_total, 2),
            "hotel_id": _hid_fb,
        }, dedup_col="nombre_plato")
        res["fb"] = round(fb_total, 2)

    return res


def procesar_contrato_grupo(image_paths, datos_dir=None, guardar_datos=True):
    """Pipeline completo: extraer -> transformar -> guardar. Devuelve resumen."""
    if isinstance(image_paths, str):
        if os.path.isdir(image_paths):
            image_paths = sorted(glob.glob(os.path.join(image_paths, "*")))
        else:
            image_paths = [image_paths]
    image_paths = [p for p in image_paths if str(p).lower().endswith((".jpg", ".jpeg", ".png", ".webp", ".heic"))]
    datos = extraer_contrato_grupo(image_paths)
    if datos.get("_needs_review"):
        return {"ok": False, "needs_review": True, "error": datos.get("_error", ""),
                "message": "No se pudo extraer automáticamente; guardado para revisión manual."}
    if datos.get("es_contrato_grupo") is False:
        return {"ok": False, "needs_review": True, "reprocesar": True,
                "error": "las fotos no parecen un contrato de grupo",
                "message": "Las imágenes no parecen un contrato de grupo/BEO."}
    # `is False` solo caza el NO explicito: si la IA no devuelve el campo, `None
    # is False` es falso y pasaba de largo. Y aunque diga que si, un contrato de
    # grupo SON sus habitaciones, sus servicios y sus importes: sin nada de eso
    # se cantaba "✓ Contrato · 0,00 €", que es un exito con las manos vacias.
    # Misma regla de producto que `albaran_tiene_datos` y `po_tiene_datos`.
    t = transformar(datos)
    _r = t.get("resumen", {}) or {}
    _hay_dinero = any(_f(_r.get(k)) for k in ("total_receivable", "habitaciones", "fb", "salas"))
    _aloj = datos.get("alojamiento", {}) or {}
    _hay_estancia = bool(_f(_aloj.get("habitaciones")) or _f(_aloj.get("noches")))
    _hay_nombre = bool(str(_r.get("contrato") or "").strip()
                       or str(_r.get("evento") or "").strip()
                       or (str(_r.get("cliente") or "").strip() not in ("", "Cliente grupo")))
    if not (_hay_dinero or _hay_estancia) and not _hay_nombre:
        # `reprocesar`: si no es un contrato aprovechable, que las fotos tengan una
        # segunda oportunidad como documentos sueltos — pueden ser un albaran o
        # una factura. Es lo que ya hacia el caso "no parecen un contrato".
        return {"ok": False, "needs_review": True, "reprocesar": True,
                "error": "no se ha podido leer ni importe, ni habitaciones, ni el nombre del contrato",
                "message": "Parece un contrato de grupo, pero no se ha extraído ningún dato "
                           "aprovechable — revisar manualmente."}
    beo = generar_beo(datos, t)
    r_dist = {}
    if guardar_datos:
        t["_paths"] = guardar(t, datos_dir)
        try:
            guardar_beo(beo, datos_dir)
        except Exception:
            pass
        try:
            r_dist = distribuir_contrato(datos, t, datos_dir)
        except Exception:
            r_dist = {}
    r = t["resumen"]; r["ok"] = True
    r["beo"] = beo
    r["beo_lineas"] = len(beo.get("lineas", []))
    r["beo_total"] = beo.get("total", 0)
    r["distribucion"] = r_dist
    # BUG 8: este camino no pedia el paso de cierre, asi que la comision se
    # guardaba y ahi se quedaba — sin cuenta contable y sin asiento. Se
    # devuelve la misma lista `cierre` que `/api/scan_documento`, y el
    # frontend la junta con la de las fotos para llamar UNA vez al cierre.
    r["cierre"] = (["ap"] if r_dist.get("ap") else []) + ["ar"]
    return r


if __name__ == "__main__":
    import sys
    print(json.dumps(procesar_contrato_grupo(sys.argv[1] if len(sys.argv) > 1 else ".", guardar_datos=False),
                     ensure_ascii=False, indent=2))
