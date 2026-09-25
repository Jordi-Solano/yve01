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
_PROMPT = """Eres un experto en contratación hotelera. Te paso TODAS las páginas (fotos o PDF)
de un contrato de grupo/eventos de hotel (contrato + BEO + anexos). Devuelve SOLO un JSON válido
(sin texto alrededor) con esta estructura exacta (usa null si un dato no aparece):
{
 "es_contrato_grupo": true,
 "evento": {"id": "", "nombre": ""},
 "contrato_numero": "",
 "fecha_contrato": "YYYY-MM-DD",
 "hotel": {"nombre": "", "cif": "", "direccion": "", "telefono": ""},
 "cliente": {"nombre": "", "cif": "", "email": "", "contacto": "", "telefono": "", "direccion": "", "pais": ""},
 "agencia": {"nombre": "", "cif": "", "contacto": "", "email": "", "telefono": "", "direccion": ""},
 "alojamiento": {"fecha_entrada":"YYYY-MM-DD","fecha_salida":"YYYY-MM-DD","noches":0,
                 "habitaciones":0,"tarifa_dui":0,"tarifa_doble":0,
                 "total_habitaciones":0,"iva_pct":10,"total_pernoctaciones":0},
 "tasa_turistica": {"por_persona_noche":0,"max_noches":0},
 "fb": {"total":0,"por_persona_dia":0,"pax":0,"dias":0,"iva_pct":10,"detalle":""},
 "salas": {"total":0,"nombre":"","montaje":"","dias":0,"detalle":""},
 "comisiones": {"modo":null,"texto":"","alojamiento_pct":0,"salas_pct":0,"fb_pct":0,"ddr_pct":0,"misc_pct":0},
 "facturacion": {"pagador":null,"texto":""},
 "deposito": {"pct":0,"cuando":"","iban":"","beneficiario":"","referencia":""},
 "doble_imposicion": false,
 "beo": {"contacto": {"nombre":"","telefono":"","email":""}, "contacto_sitio": {"nombre":"","telefono":""},
         "coordinador": "", "anuncio": "", "instrucciones_facturacion": "", "alergias": [],
         "funciones": [{"fecha":"YYYY-MM-DD","hora_inicio":"HH:MM","hora_fin":"HH:MM","funcion":"","sala":"","montaje":"",
                        "pax":0,"garantizados":0,"alquiler":0,"precio_pp":0,"menu":[],"notas_montaje":"",
                        "av":[{"concepto":"","importe":0}]}]}
}
Importante: los importes son numéricos (sin símbolo €, punto decimal). Las fechas en formato ISO.
"doble_imposicion" = true si el cliente es extranjero o el contrato menciona doble imposición / withholding / certificado de residencia fiscal.
"es_contrato_grupo" = true SOLO si es un contrato de grupo/eventos de hotel o un BEO (orden de servicio); false si son facturas sueltas, extractos u otro documento.
"agencia" = la intermediaria (agencia de viajes, DMC, OPC) que contrata para el cliente; déjala vacía si el cliente contrata directamente con el hotel.
"comisiones.modo": lee SOLO el campo/cláusula de comisiones del contrato. "neta" si dice tarifa neta / tarifas netas / net rate / no comisionable (el hotel factura el neto y no hay factura de comisión); "porcentaje" si da un % de comisión para la agencia (el hotel factura el bruto y la agencia factura su comisión aparte); null si el contrato no dice nada de comisiones. "comisiones.texto" = lo que pone ese campo, literal y breve. No lo deduzcas de otra parte del contrato.
"facturacion.pagador": "agencia" si la factura del grupo se emite a la agencia o la paga la agencia; "cliente" si la paga directamente el cliente final; null si el contrato no lo dice. "facturacion.texto" = la frase del contrato que lo dice.
"beo": lo necesario para la orden de servicio (BEO). "beo.funciones" = el programa del evento DIA A DIA tal como lo detalle el contrato o su anexo/BEO (reuniones, coffee breaks, almuerzos, cenas, montajes): una entrada por funcion con su fecha, horas, sala, montaje (escuela, teatro, cabaret, imperial, cóctel...), pax, precio por persona, alquiler de sala, menú (platos y bebidas, uno por elemento), notas de montaje y audiovisuales. Si el contrato NO trae programa día a día, "funciones" = [] (no te inventes horas, salas ni fechas). "beo.alergias" = alergias o dietas especiales que diga el contrato (una por elemento). "beo.anuncio" = el texto del cartel/anuncio del evento si lo hay. "salas.nombre/montaje/dias" = la sala, el montaje y los días de sala si el contrato los da."""


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


def _bloque(path):
    """Una pagina para la IA: foto (image) o PDF entero (document). b87: los
    contratos en PDF pasan por este mismo lector (decision de Jordi, 24 sep)."""
    if str(path).lower().endswith(".pdf"):
        with open(path, "rb") as f:
            data = base64.standard_b64encode(f.read()).decode()
        return {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": data}}
    return _img_block(path)


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
            content.append(_bloque(p))
        resp = client.messages.create(model=MODEL, max_tokens=8000,   # b91: el programa del BEO (menus) alarga el JSON
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


def transformar(datos, hotel_id=None):
    """Convierte los datos extraídos en filas para AR Real + comisión + flag DI.
    Testeable sin API.

    b87 (respuestas de finanzas, 24 sep 2026):
      - la comision solo existe si el contrato la da en %; con tarifa neta (o sin
        agencia) es 0, y si el contrato no lo dice queda PENDIENTE (0 aqui; la
        persona lo decide en AR > Contratos). Yve no crea factura de comision:
        la manda la agencia (b88 la une al contrato).
      - la factura del grupo va a nombre de QUIEN PAGA (agencia o cliente). Si el
        contrato no lo dice, a nombre del cliente hasta que alguien lo decida, y
        sin ficha de cliente AR (no se sabe de quien es la deuda).
      - el cliente AR nace SIN credito: el limite sale de una peticion firmada.
      - la factura lleva el hotel (antes se guardaba sin hotel_id y con un hotel
        elegido no se veia)."""
    import contratos_grupo as CG
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
    modo = CG.modo_comision(datos)["modo"]
    comis = calcular_comisiones(datos)
    if modo != "porcentaje":
        comis = dict(comis, alojamiento=0.0, fb=0.0, salas=0.0, total=0.0)
    pagador = CG.pagador_de(datos)["quien"]
    ag = datos.get("agencia", {}) or {}
    if hotel_id is None:
        try:
            import censo_hoteles as _censo
            hotel_id = _censo.para_guardar()
        except Exception:
            hotel_id = os.environ.get("YVE_HOTEL", "")
    # DI solo si el cliente es de un pais fuera de la UE con convenio (config_di, regla de finanzas 7 sep 2026)
    from config_di import requiere_di
    di = bool(datos.get("doble_imposicion")) or requiere_di(cli.get("pais", ""))

    nombre_cli = (cli.get("nombre") or "Cliente grupo").strip()
    numero = f"GRP-{contrato}" if contrato else f"GRP-{datetime.now().strftime('%Y%m%d%H%M')}"
    # la ficha AR es la de quien paga; sin saberlo no se crea ninguna
    if pagador == "agencia":
        deudor, nif_d, mail_d, tel_d = (ag.get("nombre") or "").strip(), ag.get("cif") or "", ag.get("email") or "", ag.get("telefono") or ""
    elif pagador == "cliente":
        deudor, nif_d, mail_d, tel_d = nombre_cli, cli.get("cif") or "", cli.get("email") or "", ""
    else:
        deudor, nif_d, mail_d, tel_d = "", "", "", ""
    cliente_row = None
    if deudor:
        cliente_row = {
            "nombre_cliente": deudor, "nif": nif_d, "email": mail_d, "telefono": tel_d,
            "dias_pago": 30, "credito_limite": 0.0, "credito_usado": 0,
            "estado_ficha": "PENDIENTE", "origen": f"contrato {contrato or numero}".strip(),
            "hotel_id": hotel_id or "",
        }
    noches = int(_f(aloj.get("noches"))) or 0
    reserva_row = {
        "numero_reserva": numero, "numero": numero, "cliente": deudor or nombre_cli,
        "fecha_entrada": aloj.get("fecha_entrada") or "", "fecha_salida": aloj.get("fecha_salida") or "",
        "fecha_emision": "", "habitaciones": int(_f(aloj.get("habitaciones"))), "noches": noches,
        "importe_habitaciones": imp_hab, "importe_fb": imp_fb, "importe_extras": imp_salas,
        "importe": total, "total": total, "estado": "PENDIENTE_FACTURA",
        "evento": (str(ev.get("id") or "") + " " + str(ev.get("nombre") or "")).strip(),
        "contrato": contrato, "comision_total": comis["total"],
        "requiere_certificado_di": di, "tipo": "CONTRATO_GRUPO",
        "modo_comision": modo, "pagador": pagador, "agencia": (ag.get("nombre") or "").strip(),
        "cliente_final": nombre_cli, "hotel_id": hotel_id or "",
    }
    return {
        "cliente": cliente_row, "reserva": reserva_row, "comisiones": comis,
        "doble_imposicion": di,
        "resumen": {
            "evento": reserva_row["evento"], "contrato": contrato, "cliente": nombre_cli,
            "total_receivable": total, "habitaciones": imp_hab, "fb": imp_fb, "salas": imp_salas,
            "comision_total": comis["total"], "requiere_certificado_di": di,
            "numero": numero, "modo_comision": modo, "pagador": pagador, "deudor": deudor,
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
    """Guarda el resumen del contrato como referencia del evento para el cruce de
    documentos del evento (eventos_referencia.json). b94: ya no se guarda en
    beos_generados.json: la seccion "BEOs desde contratos" se quito (era un resumen con
    el alojamiento, no una BEO); la BEO de verdad sale del contrato (beo_contrato.py)."""
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
    return {"referencia": ref_path}


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
    # Clientes: la ficha AR de quien paga la da de alta contratos_grupo.
    # sincronizar_factura (SIN credito, b87), porque ahi ya se sabe si una
    # persona ha cambiado quien paga. Antes se creaba aqui con 100.000 € de
    # limite regalado.
    pc = os.path.join(dd, "clientes_credito.xlsx")
    # Reservas / facturas
    pr = os.path.join(dd, "reservas_credito.xlsx")
    dfr = pd.read_excel(pr) if os.path.exists(pr) else pd.DataFrame()
    num = transformado["reserva"]["numero_reserva"]
    col = "numero_reserva" if "numero_reserva" in dfr.columns else "numero"
    if len(dfr) and col in dfr.columns:
        # reprocesar el contrato no "des-emite" la factura: su estado y sus
        # fechas (emision, cobro, recordatorio) se conservan (b87)
        prev = dfr[dfr[col].astype(str) == num]
        if len(prev):
            p0 = prev.iloc[0].to_dict()
            for k in ("estado", "fecha_emision", "fecha_cobro", "ultimo_recordatorio"):
                v = p0.get(k)
                if v is not None and str(v).strip() not in ("", "nan", "NaT", "None"):
                    transformado["reserva"][k] = v
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
      · Banco -> depósito/anticipo que el cliente adelanta (cobro previsto)
      · F&B   -> catering/banquete del evento (ingreso del evento)
    La comision YA NO va a AP (b87, decision de Jordi 24 sep): la factura de
    comision la manda la agencia y Yve la une al contrato cuando llega; hasta
    entonces lo devengado va a la provision del cierre (4109). `ap` sigue en el
    resultado, siempre None, para no romper a quien lo lea.
    Devuelve {'ap':None,'banco':importe|None,'fb':importe|None} para el log/badges."""
    dd = datos_dir or _datos_dir()
    os.makedirs(dd, exist_ok=True)
    res = {"ap": None, "banco": None, "fb": None}
    ev = datos.get("evento", {}) or {}
    evento_nombre = (str(ev.get("id") or "") + " " + str(ev.get("nombre") or "")).strip() or "Evento de grupo"
    contrato = str(datos.get("contrato_numero") or "").strip()
    hoy = datetime.now().strftime("%Y-%m-%d")

    # ── AP: la comision NO se escribe (b87). Antes se creaba aqui una factura
    # COM-<contrato> en facturas_ap: al llegar la de la agencia habria dos.
    # ── Banco: depósito previsto ──
    dep = datos.get("deposito", {}) or {}
    total_recv = _f((transformado or {}).get("resumen", {}).get("total_receivable"))
    dep_pct = _f(dep.get("pct"))
    dep_imp = round(total_recv * dep_pct / 100, 2) if dep_pct else 0.0
    if dep_imp > 0:
        # A SU fichero, no al extracto: un deposito PREVISTO no es un movimiento
        # que el banco haya hecho. Antes se escribia como fila de
        # extracto_banco.xlsx y la pestaña Banco enseñaba un ingreso inexistente
        # (inventario honesto #14). La pestaña Banco lo enseña ahora como
        # "deposito previsto, no en el extracto".
        concepto = "Depósito previsto " + str(int(dep_pct)) + "% · " + evento_nombre + (" (contrato " + contrato + ")" if contrato else "")
        try:
            import censo_hoteles as _censo
            _hid_dep = _censo.para_guardar()
        except Exception:
            _hid_dep = os.environ.get("YVE_HOTEL", "")
        _append_xlsx(os.path.join(dd, "depositos_previstos.xlsx"),
                     {"fecha_contrato": hoy, "concepto": concepto, "importe": dep_imp, "pct": dep_pct,
                      "evento": evento_nombre, "contrato": contrato or "", "cliente": ((datos.get("cliente", {}) or {}).get("nombre") or ""),
                      "estado": "PREVISTO", "hotel_id": _hid_dep},
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
    image_paths = [p for p in image_paths if str(p).lower().endswith((".jpg", ".jpeg", ".png", ".webp", ".heic", ".pdf"))]
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
    reg = None
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
        try:
            import contratos_grupo as CG
            reg = CG.registrar(datos, t, hotel_id=t["reserva"].get("hotel_id"),
                               archivo=", ".join(os.path.basename(str(p)) for p in image_paths)[:200],
                               datos_dir=datos_dir)
            CG.sincronizar_factura(reg, datos_dir)
        except Exception:
            reg = None
    r = t["resumen"]; r["ok"] = True
    if reg:
        import contratos_grupo as CG
        r["contrato_id"] = reg["id"]
        r["pendientes"] = CG.pendientes(reg)
        r["comision_esperada"] = CG.comision_esperada(reg)
    r["beo"] = beo
    r["beo_lineas"] = len(beo.get("lineas", []))
    r["beo_total"] = beo.get("total", 0)
    r["distribucion"] = r_dist
    # BUG 8: este camino no pedia el paso de cierre, asi que la comision se
    # guardaba y ahi se quedaba — sin cuenta contable y sin asiento. Se
    # devuelve la misma lista `cierre` que `/api/scan_documento`, y el
    # frontend la junta con la de las fotos para llamar UNA vez al cierre.
    # b87: el contrato ya no escribe en AP (la comision la factura la agencia).
    r["cierre"] = (["ap"] if r_dist.get("ap") else []) + ["ar"]
    return r


if __name__ == "__main__":
    import sys
    print(json.dumps(procesar_contrato_grupo(sys.argv[1] if len(sys.argv) > 1 else ".", guardar_datos=False),
                     ensure_ascii=False, indent=2))
