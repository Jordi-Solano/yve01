# -*- coding: utf-8 -*-
"""beo_contrato.py — Yve.01 · la BEO (orden de servicio) generada desde el contrato de grupo (b91).

Formato: el de la BEO de ejemplo que paso Jordi (guardada en _PRUEBAS_Els_Pins/
BEO_ejemplo_contacto.pdf). Una BEO por DIA de evento, con:
  cabecera   hotel (nombre, direccion, telefono) · BEO # · Page x of y · Fecha · Group Catering
  titulo     Orden del Servicio (BEO) · nombre del evento · dia del evento
  datos      Postear como / Fecha del evento · Cuenta / Contacto · Direccion / Telefono ·
             Email · Contacto en el sitio · Telefono en el sitio · Master # / Coord. Evento ·
             Anuncio Evento
  funciones  Hora del evento | Funcion | Sala | Montaje | Agr | Gtd | Alquiler
             Cantidad | Package | Descripcion | Precio Por Persona
  dos columnas: Menu (por funcion) y Anuncio Evento | Montaje, Audio Visuales,
             Miscelaneos (alergias...) e Instrucciones de Facturacion
  pie        Firma Autorizada de la Organizacion · Fecha · <Hotel> Approval · Fecha ·
             BEO # · Page x of y · Fecha
Sin el logo ni los datos del hotel del ejemplo: los del hotel del contrato.

Lo que el contrato no trae NO se inventa: si no hay programa dia a dia (funciones con
fecha, hora y sala), sale UNA BEO con lo contratado (F&B, salas) marcada "programa por
confirmar". Numeracion por tenant (datos-referencia/beo_numeracion.json), estable: la
misma BEO (contrato + dia) conserva su numero al regenerarla.
"""
import json
import os
from datetime import date, datetime
from io import BytesIO

import candados as _cand

NUMERACION = "beo_numeracion.json"
BEO_INICIAL = 1001
_DIAS = ("lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo")
_MESES = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre",
          "octubre", "noviembre", "diciembre")


def _txt(v):
    if v is None or (isinstance(v, float) and v != v):
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "null", "nat") else s


def _f(v):
    try:
        x = float(v)
        return 0.0 if x != x else x
    except (TypeError, ValueError):
        s = _txt(v).replace("€", "").replace(" ", "")
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return 0.0


def _r(x):
    return round(float(x or 0), 2)


def _dd(datos_dir=None):
    if datos_dir:
        return str(datos_dir)
    try:
        from tenant_dirs import datos_dir as _d
        return str(_d())
    except Exception:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "datos-referencia")


def _iso(v):
    s = _txt(v)
    if not s:
        return ""
    try:
        import pandas as pd
        t = pd.to_datetime(s[:10], dayfirst=("/" in s[:10]), errors="coerce")
        return "" if pd.isna(t) else t.strftime("%Y-%m-%d")
    except Exception:
        return ""


def fecha_larga(iso):
    """'2025-07-03' -> 'jueves, 3 julio, 2025' (como la BEO de ejemplo)."""
    i = _iso(iso)
    if not i:
        return ""
    d = date.fromisoformat(i)
    return f"{_DIAS[d.weekday()]}, {d.day} {_MESES[d.month - 1]}, {d.year}"


def fecha_corta(iso):
    i = _iso(iso)
    return f"{i[8:10]}/{i[5:7]}/{i[0:4]}" if i else ""


def eur(x):
    """'€2.500,00' como en el ejemplo."""
    s = f"{_f(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return "€" + s


def _hora(v):
    s = _txt(v)
    return s[:5] if len(s) >= 4 else s


# ── numeracion estable ──────────────────────────────────────────────────────
def _num_leer(datos_dir=None):
    try:
        with open(os.path.join(_dd(datos_dir), NUMERACION), encoding="utf-8") as fh:
            d = json.load(fh)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


@_cand.protegido(_dd)   # b99: dos BEO a la vez no se llevan el mismo numero
def numero_beo(clave, datos_dir=None):
    """El numero de la BEO `clave` (contrato|dia): el que ya tenia, o el siguiente."""
    d = _num_leer(datos_dir)
    asig = d.get("asignados") or {}
    if clave in asig:
        return int(asig[clave])
    n = int(d.get("ultimo") or (BEO_INICIAL - 1)) + 1
    asig[clave] = n
    d.update({"ultimo": n, "asignados": asig})
    _cand.escribir_json(d, os.path.join(_dd(datos_dir), NUMERACION))
    return n


# ── los datos de la BEO ─────────────────────────────────────────────────────
def hotel_de(c, datos_dir=None):
    """Nombre, direccion y telefono del hotel: censo -> hotel_config.json -> lo que diga el contrato."""
    h = {"nombre": "", "direccion": "", "telefono": ""}
    try:
        import censo_hoteles
        r = censo_hoteles.por_id(c.get("hotel_id")) or {}
        h.update({"nombre": _txt(r.get("nombre")), "direccion": _txt(r.get("direccion")), "telefono": _txt(r.get("telefono"))})
    except Exception:
        pass
    try:
        with open(os.path.join(_dd(datos_dir), "hotel_config.json"), encoding="utf-8") as fh:
            cfg = json.load(fh) or {}
        h["nombre"] = h["nombre"] or _txt(cfg.get("hotel_nombre"))
        h["direccion"] = h["direccion"] or ", ".join(x for x in (_txt(cfg.get("hotel_direccion")), _txt(cfg.get("hotel_ciudad"))) if x)
        h["telefono"] = h["telefono"] or _txt(cfg.get("hotel_telefono"))
    except Exception:
        pass
    dh = ((c.get("datos_contrato") or {}).get("hotel") or {})
    h["nombre"] = h["nombre"] or _txt(dh.get("nombre")) or "Hotel"
    h["direccion"] = h["direccion"] or _txt(dh.get("direccion"))
    h["telefono"] = h["telefono"] or _txt(dh.get("telefono"))
    return h


def _funciones(datos):
    """Las funciones del programa (lo que el contrato detalla dia a dia), limpias."""
    out = []
    for f in (((datos.get("beo") or {}).get("funciones")) or []):
        if not isinstance(f, dict):
            continue
        fe = _iso(f.get("fecha"))
        nombre = _txt(f.get("funcion"))
        if not fe or not nombre:
            continue
        av = []
        for a in (f.get("av") or []):
            if isinstance(a, dict) and _txt(a.get("concepto")):
                av.append({"concepto": _txt(a.get("concepto")), "importe": _r(_f(a.get("importe")))})
        out.append({"fecha": fe, "hora_inicio": _hora(f.get("hora_inicio")), "hora_fin": _hora(f.get("hora_fin")),
                    "funcion": nombre, "sala": _txt(f.get("sala")), "montaje": _txt(f.get("montaje")),
                    "pax": int(_f(f.get("pax"))) or None, "garantizados": int(_f(f.get("garantizados"))) or None,
                    "alquiler": _r(_f(f.get("alquiler"))), "precio_pp": _r(_f(f.get("precio_pp"))),
                    "menu": [_txt(m) for m in (f.get("menu") or []) if _txt(m)], "notas_montaje": _txt(f.get("notas_montaje")),
                    "av": av})
    return sorted(out, key=lambda x: (x["fecha"], x["hora_inicio"] or "99:99"))


def _derivadas(datos):
    """Sin programa: lo contratado (salas y F&B), sin fecha ni hora (por confirmar)."""
    fb = datos.get("fb") or {}
    sl = datos.get("salas") or {}
    out = []
    if _f(sl.get("total")):
        dias = int(_f(sl.get("dias")))
        out.append({"fecha": "", "hora_inicio": "", "hora_fin": "", "funcion": "Sala / reunión" + (f" ({dias} días)" if dias else ""),
                    "sala": _txt(sl.get("nombre")), "montaje": _txt(sl.get("montaje")), "pax": int(_f(fb.get("pax"))) or None,
                    "garantizados": None, "alquiler": _r(_f(sl.get("total"))), "precio_pp": 0.0, "menu": [],
                    "notas_montaje": _txt(sl.get("detalle")), "av": []})
    if _f(fb.get("total")):
        pax, dias, pp = int(_f(fb.get("pax"))), int(_f(fb.get("dias"))), _r(_f(fb.get("por_persona_dia")))
        out.append({"fecha": "", "hora_inicio": "", "hora_fin": "", "funcion": "Restauración (F&B)" + (f" ({dias} días)" if dias else ""),
                    "sala": "", "montaje": "", "pax": pax or None, "garantizados": None, "alquiler": 0.0,
                    "precio_pp": pp, "menu": [x for x in (_txt(fb.get("detalle")),) if x], "notas_montaje": "", "av": [],
                    "dias": dias, "total": _r(_f(fb.get("total")))})
    return out


def _importe(f):
    """Lo que vale una funcion: alquiler + menu (precio por persona x pax [x dias]) + AV."""
    if f.get("total"):
        return _r(f["total"] + f.get("alquiler", 0) + sum(a["importe"] for a in f.get("av") or []))
    menu = _r((f.get("precio_pp") or 0) * (f.get("garantizados") or f.get("pax") or 0) * (f.get("dias") or 1))
    return _r((f.get("alquiler") or 0) + menu + sum(a["importe"] for a in f.get("av") or []))


def _linea_deposito(dep):
    """b98 · El deposito como UNA frase ("Depósito del 30 % a la firma"). Antes solo iba el
    "cuando" suelto ("a la firma"): si el lector no copiaba la frase entera del contrato en
    otro campo, la BEO decia "a la firma" sin mas (visto en produccion el 25 sep)."""
    dep = dep if isinstance(dep, dict) else {}
    cuando = _txt(dep.get("cuando"))
    pct = _f(dep.get("pct"))
    if pct:
        return f"Depósito del {pct:g} %".replace(".", ",") + (f" {cuando}" if cuando else "")
    if cuando and not cuando.lower().startswith(("dep", "pago", "anticipo")):
        return f"Depósito {cuando}"
    return cuando


def beos(c, datos_dir=None, numerar=True):
    """Las BEO del contrato `c` (registro de contratos_grupo): una por dia de evento,
    o una "por confirmar" si el contrato no trae programa. Devuelve lista de dicts."""
    datos = c.get("datos_contrato") or {}
    b = datos.get("beo") or {}
    cli = datos.get("cliente") or {}
    ag = datos.get("agencia") or {}
    hotel = hotel_de(c, datos_dir)
    funcs = _funciones(datos)
    por_confirmar = not funcs
    if por_confirmar:
        funcs = _derivadas(datos)
    if not funcs:
        return []
    dias = sorted({f["fecha"] for f in funcs}) if not por_confirmar else [""]
    q = (c.get("pagador") or {}).get("quien")
    cuenta = _txt(c.get("agencia")) if q == "agencia" or (not q and _txt(c.get("agencia"))) else _txt(c.get("cliente"))
    quien_cuenta = ag if cuenta and cuenta == _txt(c.get("agencia")) else cli
    contacto = (b.get("contacto") or {}) if isinstance(b.get("contacto"), dict) else {}
    sitio = (b.get("contacto_sitio") or {}) if isinstance(b.get("contacto_sitio"), dict) else {}
    alergias = [_txt(a) for a in (b.get("alergias") or []) if _txt(a)]
    fact = []
    if _txt(c.get("factura_grupo")):
        fact.append(f"Factura del grupo {_txt(c.get('factura_numero')) or c.get('factura_grupo')} a nombre de {cuenta or '(por decidir quién paga)'}"
                    + (" (paga la agencia)" if q == "agencia" else " (paga el cliente final)" if q == "cliente" else ""))
    for x in (_txt(b.get("instrucciones_facturacion")), _txt((datos.get("facturacion") or {}).get("texto")),
              _linea_deposito(datos.get("deposito"))):
        # b97: sin repetir lo que ya dice otra linea (el "cuando" del deposito,
        # "a la firma", salia suelto debajo de la frase que ya lo decia)
        if x and not any(x.lower() in y.lower() for y in fact):
            fact.append(x)
    out = []
    for dia in dias:
        fs = [f for f in funcs if f["fecha"] == dia]
        clave = f"{c.get('id')}|{dia or 'programa'}"
        out.append({
            "numero": numero_beo(clave, datos_dir) if numerar else None, "clave": clave,
            "fecha": dia, "fecha_texto": fecha_larga(dia) if dia else
                     (f"{fecha_corta(c.get('fecha_entrada'))} - {fecha_corta(c.get('fecha_salida'))}".strip(" -") or "por confirmar"),
            "emitida": date.today().strftime("%d/%m/%Y"), "hotel": hotel,
            "postear_como": _txt(c.get("evento")) or "Evento de grupo", "cuenta": cuenta,
            "direccion": _txt(quien_cuenta.get("direccion")),
            "contacto": _txt(contacto.get("nombre")) or _txt(quien_cuenta.get("contacto")),
            "telefono": _txt(contacto.get("telefono")) or _txt(quien_cuenta.get("telefono")),
            "email": _txt(contacto.get("email")) or _txt(quien_cuenta.get("email")),
            "contacto_sitio": _txt(sitio.get("nombre")), "telefono_sitio": _txt(sitio.get("telefono")),
            "master": _txt(c.get("contrato")), "coordinador": _txt(b.get("coordinador")),
            "anuncio": _txt(b.get("anuncio")), "funciones": fs, "alergias": alergias, "facturacion": fact,
            "por_confirmar": por_confirmar, "total": _r(sum(_importe(f) for f in fs)),
        })
    return out


def cotejo(c, lista=None):
    """La BEO contra el contrato: lo que suman las funciones frente a F&B + salas del
    contrato (sin alojamiento). Informativo: la BEO manda en lo operativo, el contrato
    en lo facturado."""
    lista = beos(c, numerar=False) if lista is None else lista
    imp = c.get("importes") or {}
    contrato = _r(_f(imp.get("fb")) + _f(imp.get("salas")))
    total = _r(sum(b["total"] for b in lista))
    dif = _r(total - contrato)
    return {"n": len(lista), "por_confirmar": any(b["por_confirmar"] for b in lista), "total_beo": total,
            "contrato_fb_salas": contrato, "diferencia": dif,
            "cuadra": abs(dif) <= 1.0 if contrato else total == 0,
            "dias": [b["fecha"] for b in lista if b["fecha"]]}


# ── el PDF, con el formato de la BEO de ejemplo ─────────────────────────────
def _esc(s):
    return (_txt(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def pdf(lista):
    """Un PDF con todas las BEO de la lista, cada una con su paginacion (Page x of y)."""
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas as _canvas
    from reportlab.platypus import Flowable, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    W, H = A4
    M = 12 * mm
    ANCHO = W - 2 * M
    N = ParagraphStyle("n", fontName="Helvetica", fontSize=8, leading=10)
    NB = ParagraphStyle("nb", parent=N, fontName="Helvetica-Bold")
    C = ParagraphStyle("c", parent=N, alignment=TA_CENTER)
    CB = ParagraphStyle("cb", parent=C, fontName="Helvetica-Bold")
    SEC = ParagraphStyle("sec", parent=N, fontName="Helvetica-Bold", fontSize=10, leading=12, alignment=TA_CENTER)
    R = ParagraphStyle("r", parent=N, alignment=TA_RIGHT)
    GRIS = colors.HexColor("#d9d9d9")
    AMARILLO = colors.HexColor("#ffff00")

    class _Marca(Flowable):
        """Marca en el lienzo que BEO se esta pintando (para su cabecera y su Page x of y)."""
        def __init__(self, n):
            super().__init__(); self.n = n; self.width = self.height = 0

        def wrap(self, *a):
            return (0, 0)

        def draw(self):
            self.canv._beo_actual = self.n

    por_num = {b["numero"]: b for b in lista}

    class _Lienzo(_canvas.Canvas):
        def __init__(self, *a, **k):
            super().__init__(*a, **k)
            self._paginas = []
            self._beo_actual = None

        def showPage(self):
            self._paginas.append((dict(self.__dict__), self._beo_actual))
            self._startPage()

        def save(self):
            total = {}
            for _e, n in self._paginas:
                total[n] = total.get(n, 0) + 1
            visto = {}
            for estado, n in self._paginas:
                self.__dict__.update(estado)
                visto[n] = visto.get(n, 0) + 1
                self._marco(por_num.get(n) or {}, visto[n], total[n])
                _canvas.Canvas.showPage(self)
            _canvas.Canvas.save(self)

        def _marco(self, b, pag, tot):
            h = b.get("hotel") or {}
            self.setFont("Helvetica-Bold", 11)
            self.drawCentredString(W / 2, H - 16 * mm, _txt(h.get("nombre"))[:70])
            self.setFont("Helvetica", 8)
            if h.get("direccion"):
                self.drawCentredString(W / 2, H - 20 * mm, _txt(h.get("direccion"))[:95])
            if h.get("telefono"):
                self.drawCentredString(W / 2, H - 24 * mm, "Teléfono:: " + _txt(h.get("telefono")))
            self.setFont("Helvetica-Bold", 8)
            self.drawRightString(W - M, H - 14 * mm, f"BEO #: {b.get('numero', '')}")
            self.setFont("Helvetica", 8)
            self.drawRightString(W - M, H - 18 * mm, f"Page {pag} of {tot}")
            self.drawRightString(W - M, H - 22 * mm, f"Fecha:: {b.get('emitida', '')}")
            self.setFont("Helvetica-Bold", 8)
            self.drawRightString(W - M, H - 28 * mm, "Group Catering")
            self.setFont("Helvetica-Bold", 10)
            self.drawCentredString(W / 2, H - 31 * mm, "Orden del Servicio (BEO)")
            self.setFont("Helvetica-Bold", 8)
            self.drawCentredString(W / 2, H - 35 * mm, _txt(b.get("postear_como"))[:90])
            self.drawCentredString(W / 2, H - 39 * mm, _txt(b.get("fecha_texto")))
            # pie: las dos firmas, como en el ejemplo
            y = 16 * mm
            self.setLineWidth(0.5)
            self.line(M, y + 4, W / 2 - 8 * mm, y + 4)
            self.line(W / 2 + 4 * mm, y + 4, W - M, y + 4)
            self.setFont("Helvetica", 7)
            self.drawString(M, y - 3, "Firma Autorizada de la Organización")
            self.drawRightString(W / 2 - 8 * mm, y - 3, "Fecha")
            self.drawString(W / 2 + 4 * mm, y - 3, (_txt(h.get("nombre")) + " Approval")[:60])
            self.drawRightString(W - M, y - 3, "Fecha")
            self.setFont("Helvetica-Bold", 6.5)
            self.drawString(M, y - 10, f"BEO #: {b.get('numero', '')}")
            self.setFont("Helvetica", 6.5)
            self.drawString(M, y - 16, f"Page {pag} of {tot}")
            self.drawString(M, y - 22, f"Fecha: {b.get('emitida', '')}")

    def P(t, st=N):
        return Paragraph(t, st)

    def rejilla(t, cab=False):
        est = [("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("VALIGN", (0, 0), (-1, -1), "TOP"),
               ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
               ("TOPPADDING", (0, 0), (-1, -1), 1.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)]
        if cab:
            est.append(("BACKGROUND", (0, 0), (-1, 0), GRIS))
        t.setStyle(TableStyle(est))
        return t

    story = []
    for i, b in enumerate(lista):
        if i:
            story.append(PageBreak())
        story.append(_Marca(b["numero"]))
        # datos del evento (dos columnas de etiqueta/valor)
        dir_ = "<br/>".join(_esc(x) for x in _txt(b.get("direccion")).split("\n")) if b.get("direccion") else ""
        filas = [
            [P("<b>Postear como:</b>"), P("<b>" + _esc(b["postear_como"]) + "</b>"), P("<b>Fecha del evento:</b>"), P("<b>" + _esc(b["fecha_texto"]) + "</b>")],
            [P("<b>Cuenta:</b>"), P(_esc(b["cuenta"])), P("<b>Contacto:</b>"), P(_esc(b["contacto"]))],
            [P("<b>Dirección:</b>"), P(dir_), P("<b>Teléfono:</b><br/><br/><b>Email:</b><br/><b>Contacto en el sitio:</b><br/><br/><b>Teléfono en el sitio:</b>"),
             P(_esc(b["telefono"]) + "<br/><br/>" + _esc(b["email"]) + "<br/>" + _esc(b["contacto_sitio"]) + "<br/><br/>" + _esc(b["telefono_sitio"]))],
            [P("<b>Master #:</b><br/><b>Anuncio Evento:</b>"), P(_esc(b["master"]) + "<br/>" + _esc(b["anuncio"])), P("<b>Coord. Evento:</b>"), P(_esc(b["coordinador"]))],
        ]
        story.append(rejilla(Table(filas, colWidths=[ANCHO * .15, ANCHO * .35, ANCHO * .17, ANCHO * .33])))
        story.append(Spacer(1, 3 * mm))
        # funciones
        cab = [P("<b>Hora del evento</b>", CB), P("<b>Función</b>", CB), P("<b>Sala</b>", CB), P("<b>Montaje</b>", CB),
               P("<b>Agr</b>", CB), P("<b>Gtd</b>", CB), P("<b>Alquiler</b>", CB)]
        rows = [cab]
        for f in b["funciones"]:
            hora = (f["hora_inicio"] + " - " + f["hora_fin"]).strip(" -") if (f["hora_inicio"] or f["hora_fin"]) else "por confirmar"
            rows.append([P(_esc(hora)), P(_esc(f["funcion"])), P(_esc(f["sala"])), P(_esc(f["montaje"])),
                         P(str(f["pax"] or ""), C), P(str(f["garantizados"] or ""), C), P(eur(f["alquiler"]), R)])
        story.append(rejilla(Table(rows, colWidths=[ANCHO * .14, ANCHO * .19, ANCHO * .18, ANCHO * .19, ANCHO * .08, ANCHO * .08, ANCHO * .14], repeatRows=1), cab=True))
        story.append(rejilla(Table([[P("<b>Cantidad</b>", CB), P("<b>Package</b>", CB), P("<b>Descripcion</b>", CB), P("<b>Precio Por<br/>Persona</b>", CB)]],
                                   colWidths=[ANCHO * .12, ANCHO * .22, ANCHO * .52, ANCHO * .14]), cab=True))
        # dos columnas independientes, como en el ejemplo: Menu + Anuncio Evento | Montaje +
        # Audio Visuales + Misceláneos + Instrucciones de Facturación. Una sola fila que se
        # puede partir entre paginas (splitInRow) si el programa es largo.
        colW = ANCHO * .5

        def cab_sec(txt):
            t = Table([[P(txt, SEC)]], colWidths=[colW])
            t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), GRIS), ("LINEABOVE", (0, 0), (-1, -1), 0.5, colors.black),
                                   ("LINEBELOW", (0, 0), (-1, -1), 0.5, colors.black), ("TOPPADDING", (0, 0), (-1, -1), 1.5),
                                   ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5)]))
            return t

        def bloque(pars):
            t = Table([[x] for x in pars], colWidths=[colW])
            t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                                   ("TOPPADDING", (0, 0), (-1, -1), 0.6), ("BOTTOMPADDING", (0, 0), (-1, -1), 0.6)]))
            return t

        def titulo_f(f):
            return "<b>" + _esc(" | ".join(x for x in (f["funcion"], f["sala"], (f["hora_inicio"] + " - " + f["hora_fin"]).strip(" -")) if x)) + "</b>"

        izq, der = [], []
        menus = [f for f in b["funciones"] if f["menu"] or f["precio_pp"]]
        if menus:
            izq.append(cab_sec("Menú"))
            for f in menus:
                t = [P(titulo_f(f), C)]
                pax = f.get("garantizados") or f.get("pax")
                if pax:
                    t.append(P(f"{pax} {_esc(f['funcion'])}", C))
                t += [P(_esc(m), C) for m in f["menu"]]
                if f["precio_pp"]:
                    t.append(P(f"<b>@ {eur(f['precio_pp'])} por persona</b>" + (f" x {f['dias']} días" if f.get("dias") else ""), C))
                izq += [bloque(t), Spacer(1, 2 * mm)]
        if b["anuncio"]:
            izq += [cab_sec("Anuncio Evento"), bloque([P('<font backColor="#ffff00">' + _esc(b["anuncio"]) + "</font>", C)])]
        mont = [f for f in b["funciones"] if f["montaje"] or f["notas_montaje"]]
        if mont:
            der.append(cab_sec("Montaje"))
            for f in mont:
                der += [bloque([P(titulo_f(f), C), P(_esc(f["montaje"]), C)] + ([P(_esc(f["notas_montaje"]), C)] if f["notas_montaje"] else [])),
                        Spacer(1, 1.5 * mm)]
        avs = [f for f in b["funciones"] if f["av"]]
        if avs:
            der.append(cab_sec("Audio Visuales"))
            for f in avs:
                der += [bloque([P(titulo_f(f), C)] + sum(([P(_esc(a["concepto"]), C), P(f"<b>@ {eur(a['importe'])} Per Event</b>", C)] for a in f["av"]), [])),
                        Spacer(1, 1.5 * mm)]
        misc = []
        if b["alergias"]:
            misc.append(P('<font backColor="#ffff00">ALERGIAS ALIMENTARIAS</font>', C))
            misc += [P("o&nbsp;&nbsp;" + _esc(a), N) for a in b["alergias"]]
        if b["por_confirmar"]:
            misc.append(P("<b>Programa por confirmar:</b> el contrato no detalla horarios ni salas día a día. "
                          "Complétalo con el cliente antes del evento.", C))
        if misc:
            der += [cab_sec("Misceláneos"), bloque(misc), Spacer(1, 1.5 * mm)]
        if b["facturacion"]:
            der += [cab_sec("Instrucciones de Facturación"), bloque([P(_esc(x), C) for x in b["facturacion"]])]
        if izq or der:
            t = Table([[izq or "", der or ""]], colWidths=[colW, colW], splitInRow=1)
            t.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.5, colors.black), ("LINEAFTER", (0, 0), (0, -1), 0.5, colors.black),
                                   ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                   ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
                                   ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
            story.append(t)
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=M, rightMargin=M, topMargin=44 * mm, bottomMargin=30 * mm,
                            title="BEO", author="Yve.01")
    doc.build(story, canvasmaker=_Lienzo)
    buf.seek(0)
    return buf
