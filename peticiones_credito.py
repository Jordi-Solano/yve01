# -*- coding: utf-8 -*-
"""peticiones_credito.py — Yve.01 · la peticion de credito de un cliente (b90).

El proceso real (finanzas, 24 sep 2026), y el formulario lo sigue paso a paso:
  (a) historial: si ya hemos trabajado con esa agencia o empresa, como pago
      (se saca de AR, de TODOS los hoteles del grupo: facturas, cobros, retrasos);
  (b) referencias: si ha trabajado con otro hotel del grupo se piden referencias
      (campo para anotarlas; obligatorio si Yve ve facturas en otro hotel o si
      quien pide marca que ha trabajado con otro hotel);
  (c) datos fiscales correctos y completos, porque con ellos se pide un informe
      de solvencia a Informa (campo para adjuntar el informe o anotar el resultado);
  (d) comercial estima el potencial de venta y con eso se fija un limite REVISABLE
      (potencial, limite y fecha de revision);
  (e) firman SIEMPRE dos personas, sin umbral: quien lo solicita y la directora.
      Solo el rol "direccion" firma como directora, y nunca quien pidio el credito.

Decisiones de Jordi (24 sep 2026):
  - el limite de credito de un cliente SOLO sale de una peticion firmada por los
    dos; la ficha lo enseña sin poder editarlo, con el boton "Pedir credito";
  - los clientes de contrato o de bono nacen sin credito (limite 0);
  - los limites escritos a mano antes de esto no se borran: quedan marcados
    "limite sin peticion" para revisar.

Estados: BORRADOR -> (firma quien pide) PENDIENTE_DIRECCION -> (firma Direccion)
APROBADA | RECHAZADA; una rechazada se puede reabrir (vuelve a BORRADOR y hay que
firmarla otra vez). Lo que la directora firma es exactamente lo que firmo quien
pidio: no lo cambia; si no esta de acuerdo, la rechaza con una nota.

Registro: datos-referencia/peticiones_credito.json. Informes de Informa adjuntos:
datos-referencia/credito_informa/. Funciones puras donde se puede; `datos_dir`
para las pruebas.
"""
import json
import os
import re
import unicodedata
from datetime import date, datetime

FICHERO = "peticiones_credito.json"
CARPETA_INFORMA = "credito_informa"
ESTADOS = ("BORRADOR", "PENDIENTE_DIRECCION", "APROBADA", "RECHAZADA")
ABIERTAS = ("BORRADOR", "PENDIENTE_DIRECCION")
ROL_DIRECCION = "direccion"
EXT_INFORMA = (".pdf", ".png", ".jpg", ".jpeg", ".webp")
MAX_INFORMA = 10 * 1024 * 1024
CAMPOS_FISCAL = ("razon_social", "nif", "direccion", "cp", "poblacion", "pais", "email", "telefono")
OBLIGATORIOS_FISCAL = ("razon_social", "nif", "direccion", "cp", "poblacion", "pais")


class ReglaError(ValueError):
    """La peticion no se puede firmar/cambiar por una regla del proceso (409)."""


class PermisoError(PermissionError):
    """Quien lo intenta no puede (403): no es Direccion, o es quien lo pidio."""


# ── utilidades ──────────────────────────────────────────────────────────────
def _txt(v):
    if v is None or (isinstance(v, float) and v != v):
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "null", "nat") else s


def _f(v):
    if v is None or v == "":
        return None
    try:
        x = float(v)
        return None if x != x else x
    except (TypeError, ValueError):
        s = _txt(v).replace("€", "").replace(" ", "")
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None


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


def _ahora():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


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


def norm_nombre(v):
    """Nombre de empresa comparable: sin acentos, sin forma juridica ni puntuacion."""
    s = unicodedata.normalize("NFKD", _txt(v).lower())
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[.,;:()\"'`´\-_/&]", " ", s)
    s = re.sub(r"\b(s\s*l\s*u?|s\s*a\s*u?|s\s*l\s*l|s\s*coop|sociedad limitada|sociedad anonima|ltd|limited|gmbh|sarl|sas|srl|spa|inc|llc|bv)\b", " ", s)
    return re.sub(r"\s+", " ", s).strip()


# ── NIF / CIF / NIE ─────────────────────────────────────────────────────────
_LETRAS_DNI = "TRWAGMYFPDXBNJZSQVHLCKE"


def norm_nif(v):
    s = re.sub(r"[\s.\-/]", "", _txt(v).upper())
    if s.startswith("ES") and len(s) == 11:
        s = s[2:]
    return s


def validar_nif(v, pais="ES"):
    """(ok, nif normalizado, tipo o motivo). España: DNI, NIE o CIF con su control.
    Fuera de España: numero de IVA/identificacion (5-20 letras y cifras), sin control."""
    nif = norm_nif(v)
    p = norm_nombre(pais) if pais else "es"
    espana = p in ("", "es", "esp", "espana", "spain", "espanya", "espagne", "spanien", "spagna", "espanha")
    if not nif:
        return False, "", "falta el NIF/CIF"
    if not espana:
        ok = bool(re.fullmatch(r"[A-Z0-9]{5,20}", nif))
        return ok, nif, "identificacion fiscal extranjera" if ok else "identificacion fiscal no valida"
    if re.fullmatch(r"\d{8}[A-Z]", nif):
        ok = _LETRAS_DNI[int(nif[:8]) % 23] == nif[8]
        return ok, nif, "DNI" if ok else "la letra del DNI no cuadra"
    if re.fullmatch(r"[XYZ]\d{7}[A-Z]", nif):
        num = int("XYZ".index(nif[0]).__str__() + nif[1:8])
        ok = _LETRAS_DNI[num % 23] == nif[8]
        return ok, nif, "NIE" if ok else "la letra del NIE no cuadra"
    if re.fullmatch(r"[ABCDEFGHJNPQRSUVW]\d{7}[0-9A-J]", nif):
        dig = [int(c) for c in nif[1:8]]
        pares = dig[1] + dig[3] + dig[5]
        impares = sum(sum(divmod(d * 2, 10)) for d in (dig[0], dig[2], dig[4], dig[6]))
        ctl = (10 - (pares + impares) % 10) % 10
        letra, fin = "JABCDEFGHI"[ctl], nif[8]
        if nif[0] in "PQRSNW":
            ok = fin == letra
        elif nif[0] in "ABEH":
            ok = fin == str(ctl)
        else:
            ok = fin in (str(ctl), letra)
        return ok, nif, "CIF" if ok else "el digito de control del CIF no cuadra"
    return False, nif, "formato de NIF/CIF no valido"


# ── registro ────────────────────────────────────────────────────────────────
def leer(datos_dir=None):
    try:
        with open(os.path.join(_dd(datos_dir), FICHERO), encoding="utf-8") as fh:
            d = json.load(fh)
        return [p for p in d if isinstance(p, dict)] if isinstance(d, list) else []
    except Exception:
        return []


def _escribir(lista, datos_dir=None):
    ruta = os.path.join(_dd(datos_dir), FICHERO)
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    tmp = ruta + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(lista, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, ruta)


def buscar(pid, datos_dir=None):
    return next((p for p in leer(datos_dir) if p.get("id") == pid), None)


def _guardar_una(p, datos_dir=None):
    lista = leer(datos_dir)
    lista = [x for x in lista if x.get("id") != p.get("id")] + [p]
    _escribir(lista, datos_dir)
    return p


def _log(p, usuario, que):
    h = p.get("historial") or []
    h.append({"cuando": _ahora(), "usuario": usuario, "que": que})
    p["historial"] = h[-60:]


def del_cliente(nombre, lista=None, datos_dir=None):
    n = norm_nombre(nombre)
    lista = leer(datos_dir) if lista is None else lista
    return sorted([p for p in lista if norm_nombre(p.get("cliente")) == n], key=lambda p: p.get("creada") or "")


def abierta_de(nombre, lista=None, datos_dir=None):
    return next((p for p in reversed(del_cliente(nombre, lista, datos_dir)) if p.get("estado") in ABIERTAS), None)


def aprobada_de(nombre, lista=None, datos_dir=None):
    """La ultima peticion APROBADA del cliente (la que manda en su limite)."""
    ap = [p for p in del_cliente(nombre, lista, datos_dir) if p.get("estado") == "APROBADA"]
    return sorted(ap, key=lambda p: ((p.get("firmas") or {}).get("direccion") or {}).get("cuando") or "")[-1] if ap else None


# ── (a) historial de pago en AR, de todo el grupo ───────────────────────────
def _clientes(datos_dir=None):
    import pandas as pd
    ruta = os.path.join(_dd(datos_dir), "clientes_credito.xlsx")
    if not os.path.exists(ruta):
        return pd.DataFrame()
    try:
        return pd.read_excel(ruta)
    except Exception:
        return pd.DataFrame()


def _reservas(datos_dir=None):
    import pandas as pd
    ruta = os.path.join(_dd(datos_dir), "reservas_credito.xlsx")
    if not os.path.exists(ruta):
        return pd.DataFrame()
    try:
        return pd.read_excel(ruta)
    except Exception:
        return pd.DataFrame()


def _nombre_hotel(hid):
    if not hid:
        return ""
    try:
        import censo_hoteles
        return censo_hoteles.nombre_de(hid) or hid
    except Exception:
        return hid


def historial_ar(nombre, nif="", hotel_actual="", datos_dir=None, hoy=None, dias_pago=None):
    """Como ha pagado este cliente en TODOS los hoteles del grupo (reservas_credito.xlsx).
    El cliente se reconoce por su nombre y por el de las fichas que comparten su NIF."""
    hoy = hoy or date.today()
    dfc = _clientes(datos_dir)
    nombres = {norm_nombre(nombre)} - {""}
    nifn = norm_nif(nif)
    terminos = None
    if not dfc.empty and "nombre_cliente" in dfc.columns:
        for f in dfc.to_dict("records"):
            misma = norm_nombre(f.get("nombre_cliente")) in nombres
            mismo_nif = bool(nifn) and norm_nif(f.get("nif")) == nifn
            if misma or mismo_nif:
                nombres.add(norm_nombre(f.get("nombre_cliente")))
                d = _f(f.get("dias_pago"))
                if d and terminos is None:
                    terminos = int(d)
    terminos = int(dias_pago or terminos or 30)
    rv = _reservas(datos_dir)
    fact = []
    if not rv.empty and "cliente" in rv.columns:
        for f in rv.to_dict("records"):
            if norm_nombre(f.get("cliente")) not in nombres:
                continue
            est = _txt(f.get("estado")).upper()
            if est not in ("FACTURADO", "COBRADO", "COBRADA"):
                continue            # lo pendiente de emitir no es historial de pago
            em = _iso(f.get("fecha_emision")) or _iso(f.get("fecha_entrada"))
            co = _iso(f.get("fecha_cobro")) if est in ("COBRADO", "COBRADA") else ""
            total = _r(_f(f.get("total")) or _f(f.get("importe")) or 0)
            dias = None
            if em and co:
                dias = (date.fromisoformat(co) - date.fromisoformat(em)).days
            elif em:
                dias = (hoy - date.fromisoformat(em)).days
            fact.append({"numero": _txt(f.get("numero_reserva")) or _txt(f.get("numero")), "hotel_id": _txt(f.get("hotel_id")),
                         "hotel": _nombre_hotel(_txt(f.get("hotel_id"))), "fecha_emision": em, "fecha_cobro": co,
                         "total": total, "estado": "COBRADO" if est in ("COBRADO", "COBRADA") else "FACTURADO",
                         "dias": dias, "retraso": max(0, dias - terminos) if dias is not None else None})
    cobr = [x for x in fact if x["estado"] == "COBRADO" and x["dias"] is not None]
    pend = [x for x in fact if x["estado"] == "FACTURADO"]
    venc = [x for x in pend if (x["dias"] or 0) > terminos]
    hoteles = {}
    for x in fact:
        h = hoteles.setdefault(x["hotel_id"], {"hotel_id": x["hotel_id"], "hotel": x["hotel"], "facturas": 0, "facturado": 0.0})
        h["facturas"] += 1
        h["facturado"] = _r(h["facturado"] + x["total"])
    # "otro hotel del grupo": los que no son el de la peticion; una peticion sin hotel
    # (vista de grupo) no es de ninguno, asi que cuentan todos los que tienen facturas
    otros = [h for k, h in hoteles.items() if k and k != _txt(hotel_actual)]
    fechas = sorted(x["fecha_emision"] for x in fact if x["fecha_emision"])
    return {
        "hay": bool(fact), "dias_pago": terminos,
        "facturas": len(fact), "facturado": _r(sum(x["total"] for x in fact)),
        "cobradas": len(cobr), "dias_medio_cobro": round(sum(x["dias"] for x in cobr) / len(cobr), 1) if cobr else None,
        "dias_max_cobro": max((x["dias"] for x in cobr), default=None),
        "cobradas_con_retraso": sum(1 for x in cobr if x["retraso"]),
        "pendientes": len(pend), "pendiente": _r(sum(x["total"] for x in pend)),
        "vencidas": len(venc), "vencido": _r(sum(x["total"] for x in venc)),
        "primera": fechas[0] if fechas else "", "ultima": fechas[-1] if fechas else "",
        "hoteles": sorted(hoteles.values(), key=lambda h: -h["facturado"]), "otros_hoteles": otros,
        "detalle": sorted(fact, key=lambda x: x["fecha_emision"] or "", reverse=True)[:12],
        "calculado": _ahora(),
    }


# ── la peticion ─────────────────────────────────────────────────────────────
def _ficha(nombre, datos_dir=None):
    dfc = _clientes(datos_dir)
    if dfc.empty or "nombre_cliente" not in dfc.columns:
        return {}
    n = norm_nombre(nombre)
    for f in dfc.to_dict("records"):
        if norm_nombre(f.get("nombre_cliente")) == n:
            return f
    return {}


def nueva(cliente, hotel_id="", usuario="", nombre_usuario="", datos_dir=None):
    """Crea la peticion de `cliente` en BORRADOR (o devuelve la que ya este abierta),
    con los datos fiscales que ya tenga su ficha."""
    cliente = _txt(cliente)
    if not cliente:
        raise ValueError("falta el cliente")
    lista = leer(datos_dir)
    ya = abierta_de(cliente, lista)
    if ya:
        return ya, False
    fi = _ficha(cliente, datos_dir)
    nombre = _txt(fi.get("nombre_cliente")) or cliente
    n = sum(1 for p in lista if (p.get("id") or "").startswith("PC-" + date.today().strftime("%Y%m%d"))) + 1
    p = {
        "id": f"PC-{date.today().strftime('%Y%m%d')}-{n:03d}", "cliente": nombre, "hotel_id": _txt(hotel_id),
        "estado": "BORRADOR", "creada": _ahora(), "creada_por": usuario,
        "referencias": {"otro_hotel": False, "texto": ""},
        "fiscal": {"razon_social": nombre, "nif": _txt(fi.get("nif")), "direccion": _txt(fi.get("direccion")),
                   "cp": _txt(fi.get("cp")), "poblacion": _txt(fi.get("poblacion")), "pais": _txt(fi.get("pais")) or "España",
                   "email": _txt(fi.get("email")), "telefono": _txt(fi.get("telefono"))},
        "informa": {"resultado": "", "fecha": "", "fichero": "", "nombre_fichero": ""},
        "comercial": {"potencial": None, "limite": None, "revision": "", "comentario": ""},
        "firmas": {"solicitante": None, "direccion": None},
        "historial": [],
    }
    _log(p, usuario, "creada")
    lista.append(p)
    _escribir(lista, datos_dir)
    return p, True


def guardar(pid, cambios, usuario="", datos_dir=None):
    """Cambia lo que se rellena a mano. Solo en BORRADOR (una rechazada se reabre antes)."""
    p = buscar(pid, datos_dir)
    if p is None:
        raise KeyError("peticion no encontrada")
    if p.get("estado") != "BORRADOR":
        raise ReglaError("La petición ya está firmada: no se puede cambiar." if p.get("estado") != "RECHAZADA"
                         else "La petición está rechazada: reábrela para cambiarla.")
    c = cambios or {}
    if "referencias" in c:
        r = c.get("referencias") or {}
        p["referencias"] = {"otro_hotel": bool(r.get("otro_hotel")), "texto": _txt(r.get("texto"))[:2000]}
    if "fiscal" in c:
        fi = dict(p.get("fiscal") or {})
        for k in CAMPOS_FISCAL:
            if k in (c.get("fiscal") or {}):
                fi[k] = _txt((c.get("fiscal") or {}).get(k))[:200]
        p["fiscal"] = fi
    if "informa" in c:
        inf = dict(p.get("informa") or {})
        for k in ("resultado", "fecha"):
            if k in (c.get("informa") or {}):
                v = (c.get("informa") or {}).get(k)
                inf[k] = _iso(v) if k == "fecha" else _txt(v)[:1000]
        p["informa"] = inf
    if "comercial" in c:
        co = dict(p.get("comercial") or {})
        cc = c.get("comercial") or {}
        for k in ("potencial", "limite"):
            if k in cc:
                x = _f(cc.get(k))
                if x is not None and x < 0:
                    raise ValueError(f"{k}: no puede ser negativo")
                co[k] = _r(x) if x is not None else None
        if "revision" in cc:
            co["revision"] = _iso(cc.get("revision"))
        if "comentario" in cc:
            co["comentario"] = _txt(cc.get("comentario"))[:1000]
        p["comercial"] = co
    _log(p, usuario, "guardada")
    return _guardar_una(p, datos_dir)


def adjuntar_informa(pid, nombre_fichero, contenido, usuario="", datos_dir=None):
    p = buscar(pid, datos_dir)
    if p is None:
        raise KeyError("peticion no encontrada")
    if p.get("estado") != "BORRADOR":
        raise ReglaError("La petición ya está firmada: no se puede cambiar el informe.")
    base = os.path.basename(_txt(nombre_fichero)) or "informe.pdf"
    ext = os.path.splitext(base)[1].lower()
    if ext not in EXT_INFORMA:
        raise ValueError("el informe tiene que ser PDF o imagen (png, jpg, webp)")
    if not contenido:
        raise ValueError("el fichero está vacío")
    if len(contenido) > MAX_INFORMA:
        raise ValueError("el fichero pasa de 10 MB")
    seguro = re.sub(r"[^A-Za-z0-9._-]", "_", base)[:80]
    carpeta = os.path.join(_dd(datos_dir), CARPETA_INFORMA)
    os.makedirs(carpeta, exist_ok=True)
    rel = f"{pid}_{seguro}"
    with open(os.path.join(carpeta, rel), "wb") as fh:
        fh.write(contenido)
    inf = dict(p.get("informa") or {})
    viejo = inf.get("fichero")
    if viejo and viejo != rel:
        try:
            os.remove(os.path.join(carpeta, os.path.basename(viejo)))
        except OSError:
            pass
    inf["fichero"], inf["nombre_fichero"] = rel, base
    p["informa"] = inf
    _log(p, usuario, f"informe de Informa adjunto ({base})")
    return _guardar_una(p, datos_dir)


def ruta_informa(p, datos_dir=None):
    rel = os.path.basename(_txt((p.get("informa") or {}).get("fichero")))
    if not rel:
        return None
    r = os.path.join(_dd(datos_dir), CARPETA_INFORMA, rel)
    return r if os.path.exists(r) else None


def faltan(p, historial=None, hoy=None):
    """Lo que falta para que quien pide pueda firmar: [(clave, texto)]."""
    hoy = hoy or date.today()
    out = []
    fi = p.get("fiscal") or {}
    nombres = {"razon_social": "la razón social", "nif": "el NIF/CIF", "direccion": "la dirección", "cp": "el código postal",
               "poblacion": "la población", "pais": "el país"}
    for k in OBLIGATORIOS_FISCAL:
        if not _txt(fi.get(k)):
            out.append(("fiscal." + k, "Datos fiscales: falta " + nombres[k] + "."))
    if _txt(fi.get("nif")):
        ok, _n, motivo = validar_nif(fi.get("nif"), fi.get("pais"))
        if not ok:
            out.append(("fiscal.nif", "Datos fiscales: " + motivo + "."))
    if _txt(fi.get("cp")) and norm_nombre(fi.get("pais")) in ("", "es", "espana", "spain") and not re.fullmatch(r"\d{5}", _txt(fi.get("cp"))):
        out.append(("fiscal.cp", "Datos fiscales: el código postal tiene que tener 5 cifras."))
    if _txt(fi.get("email")) and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", _txt(fi.get("email"))):
        out.append(("fiscal.email", "Datos fiscales: el email no es válido."))
    inf = p.get("informa") or {}
    if not _txt(inf.get("resultado")) and not _txt(inf.get("fichero")):
        out.append(("informa", "Informe de solvencia (Informa): adjúntalo o anota el resultado."))
    co = p.get("comercial") or {}
    if not co.get("potencial"):
        out.append(("comercial.potencial", "Comercial: falta el potencial de venta."))
    if not co.get("limite"):
        out.append(("comercial.limite", "Comercial: falta el límite propuesto."))
    rev = _iso(co.get("revision"))
    if not rev:
        out.append(("comercial.revision", "Comercial: falta la fecha de revisión del límite."))
    elif rev <= hoy.isoformat():
        out.append(("comercial.revision", "Comercial: la fecha de revisión tiene que ser futura."))
    ref = p.get("referencias") or {}
    otros = (historial or {}).get("otros_hoteles") or []
    if (ref.get("otro_hotel") or otros) and not _txt(ref.get("texto")):
        out.append(("referencias", "Referencias: ha trabajado con otro hotel del grupo; anota las referencias."
                                   + (" (Yve ve facturas en " + ", ".join(h.get("hotel") or h.get("hotel_id") for h in otros) + ")" if otros else "")))
    return out


def firmar_solicitante(pid, usuario, nombre_usuario="", datos_dir=None, hoy=None):
    """Firma 1: quien pide. Congela el historial de AR tal como estaba al firmar."""
    p = buscar(pid, datos_dir)
    if p is None:
        raise KeyError("peticion no encontrada")
    if p.get("estado") != "BORRADOR":
        raise ReglaError("Esta petición no está en borrador.")
    if not usuario:
        raise PermisoError("hace falta un usuario para firmar")
    fi = p.get("fiscal") or {}
    h = historial_ar(p.get("cliente"), fi.get("nif"), p.get("hotel_id"), datos_dir, hoy)
    f = faltan(p, h, hoy)
    if f:
        raise ReglaError("Falta: " + " ".join(t for _k, t in f))
    ok, nifn, _t = validar_nif(fi.get("nif"), fi.get("pais"))
    p["fiscal"] = dict(fi, nif=nifn)
    p["historial_ar"] = h
    p["firmas"] = {"solicitante": {"usuario": usuario, "nombre": nombre_usuario or usuario, "cuando": _ahora()}, "direccion": None}
    p["estado"] = "PENDIENTE_DIRECCION"
    _log(p, usuario, "firmada por quien la pide; pendiente de Dirección")
    return _guardar_una(p, datos_dir)


def puede_firmar_direccion(p, usuario, rol):
    """(True, "") o (False, motivo). Solo el rol Direccion y nunca quien pidio el credito."""
    if (p or {}).get("estado") != "PENDIENTE_DIRECCION":
        return False, "La petición no está pendiente de la firma de Dirección."
    if rol != ROL_DIRECCION:
        return False, "Solo firma el rol Dirección."
    sol = ((p.get("firmas") or {}).get("solicitante") or {}).get("usuario")
    if usuario and sol and usuario == sol:
        return False, "Quien pidió el crédito no puede firmarlo como Dirección."
    return True, ""


def firmar_direccion(pid, usuario, rol, decision, nota="", nombre_usuario="", datos_dir=None):
    """Firma 2: Direccion aprueba (el limite pasa a la ficha del cliente) o rechaza (con nota)."""
    p = buscar(pid, datos_dir)
    if p is None:
        raise KeyError("peticion no encontrada")
    ok, motivo = puede_firmar_direccion(p, usuario, rol)
    if not ok:
        if p.get("estado") != "PENDIENTE_DIRECCION":
            raise ReglaError(motivo)
        raise PermisoError(motivo)
    if decision not in ("aprobar", "rechazar"):
        raise ValueError("decision: aprobar o rechazar")
    if decision == "rechazar" and not _txt(nota):
        raise ValueError("para rechazar, di por qué (nota)")
    firmas = dict(p.get("firmas") or {})
    firmas["direccion"] = {"usuario": usuario, "nombre": nombre_usuario or usuario, "cuando": _ahora(),
                           "decision": decision, "nota": _txt(nota)[:1000]}
    p["firmas"] = firmas
    if decision == "aprobar":
        p["estado"] = "APROBADA"
        p["limite_aprobado"] = _r((p.get("comercial") or {}).get("limite"))
        p["revision"] = _iso((p.get("comercial") or {}).get("revision"))
        _log(p, usuario, f"aprobada por Dirección: límite {p['limite_aprobado']:.2f} EUR hasta revisión {p['revision']}")
        _guardar_una(p, datos_dir)
        aplicar_a_ficha(p, datos_dir)
    else:
        p["estado"] = "RECHAZADA"
        _log(p, usuario, "rechazada por Dirección: " + _txt(nota)[:200])
        _guardar_una(p, datos_dir)
    return p


def reabrir(pid, usuario="", datos_dir=None):
    """Una rechazada vuelve a BORRADOR para corregirla; hay que firmarla otra vez."""
    p = buscar(pid, datos_dir)
    if p is None:
        raise KeyError("peticion no encontrada")
    if p.get("estado") != "RECHAZADA":
        raise ReglaError("Solo se reabre una petición rechazada.")
    ya = abierta_de(p.get("cliente"), datos_dir=datos_dir)
    if ya and ya.get("id") != pid:
        raise ReglaError(f"El cliente ya tiene otra petición abierta ({ya.get('id')}).")
    viejas = p.get("firmas_anteriores") or []
    viejas.append(p.get("firmas"))
    p["firmas_anteriores"] = viejas[-10:]
    p["firmas"] = {"solicitante": None, "direccion": None}
    p["estado"] = "BORRADOR"
    _log(p, usuario, "reabierta para corregirla")
    return _guardar_una(p, datos_dir)


def aplicar_a_ficha(p, datos_dir=None):
    """El limite aprobado (y los datos fiscales firmados) a la ficha del cliente en
    clientes_credito.xlsx. Si no tenia ficha, se crea."""
    import pandas as pd
    ruta = os.path.join(_dd(datos_dir), "clientes_credito.xlsx")
    df = _clientes(datos_dir)
    fi = p.get("fiscal") or {}
    dire = (p.get("firmas") or {}).get("direccion") or {}
    valores = {"credito_limite": _r(p.get("limite_aprobado")), "credito_revision": _txt(p.get("revision")),
               "credito_peticion": _txt(p.get("id")), "credito_aprobado": _txt(dire.get("cuando"))[:10],
               "credito_aprobado_por": _txt(dire.get("nombre") or dire.get("usuario")),
               "nif": _txt(fi.get("nif")), "direccion": _txt(fi.get("direccion")), "cp": _txt(fi.get("cp")),
               "poblacion": _txt(fi.get("poblacion")), "pais": _txt(fi.get("pais")), "estado_ficha": "COMPLETA"}
    for k in ("email", "telefono"):
        if _txt(fi.get(k)):
            valores[k] = _txt(fi.get(k))
    n = norm_nombre(p.get("cliente"))
    if not df.empty and "nombre_cliente" in df.columns and (df["nombre_cliente"].map(norm_nombre) == n).any():
        m = df["nombre_cliente"].map(norm_nombre) == n
        for k, v in valores.items():
            if k not in df.columns:
                df[k] = ""
            df[k] = df[k].astype(object)
            df.loc[m, k] = v
    else:
        fila = dict(valores, nombre_cliente=_txt(p.get("cliente")), credito_usado=0, dias_pago=30,
                    hotel_id=_txt(p.get("hotel_id")), origen=f"petición {p.get('id')}")
        df = pd.concat([df, pd.DataFrame([fila])], ignore_index=True) if not df.empty else pd.DataFrame([fila])
    tmp = ruta + ".tmp.xlsx"
    df.to_excel(tmp, index=False)
    os.replace(tmp, ruta)
    return True


def estado_limite(fila_cliente, lista=None, datos_dir=None, hoy=None):
    """Que es el limite que enseña la ficha: 'peticion' (sale de una peticion aprobada),
    'sin_peticion' (escrito a mano antes de b90: se conserva, marcado para revisar) o
    'sin_credito'. Con la fecha de revision y si ya ha pasado."""
    hoy = hoy or date.today()
    lim = _f(fila_cliente.get("credito_limite")) or _f(fila_cliente.get("limite_credito")) or 0.0
    nombre = _txt(fila_cliente.get("nombre_cliente")) or _txt(fila_cliente.get("nombre"))
    lista = leer(datos_dir) if lista is None else lista
    ap = aprobada_de(nombre, lista)
    pid = _txt(fila_cliente.get("credito_peticion"))
    abierta = abierta_de(nombre, lista)
    if lim > 0 and ap and (not pid or pid == ap.get("id")):
        rev = _iso(fila_cliente.get("credito_revision")) or _iso(ap.get("revision"))
        return {"origen": "peticion", "peticion": ap.get("id"), "revision": rev,
                "revision_vencida": bool(rev and rev < hoy.isoformat()),
                "aprobado_por": ((ap.get("firmas") or {}).get("direccion") or {}).get("nombre") or "",
                "abierta": (abierta or {}).get("id", ""), "abierta_estado": (abierta or {}).get("estado", "")}
    return {"origen": "sin_peticion" if lim > 0 else "sin_credito", "peticion": "", "revision": "", "revision_vencida": False,
            "aprobado_por": "", "abierta": (abierta or {}).get("id", ""), "abierta_estado": (abierta or {}).get("estado", "")}


def vista(p, usuario="", rol="", datos_dir=None, hoy=None, historial=None):
    """Lo que la pantalla necesita de una peticion: la peticion, lo que falta, el
    historial (en borrador, el de ahora; firmada, el congelado al firmar) y quien
    puede hacer que."""
    fi = p.get("fiscal") or {}
    h = p.get("historial_ar") if p.get("estado") != "BORRADOR" and p.get("historial_ar") else (
        historial or historial_ar(p.get("cliente"), fi.get("nif"), p.get("hotel_id"), datos_dir, hoy))
    ok_nif, nifn, tipo = validar_nif(fi.get("nif"), fi.get("pais")) if _txt(fi.get("nif")) else (False, "", "")
    puede_d, motivo_d = puede_firmar_direccion(p, usuario, rol)
    return dict(p, historial_ar=h, faltan=[t for _k, t in faltan(p, h, hoy)] if p.get("estado") == "BORRADOR" else [],
                nif_ok=ok_nif, nif_tipo=tipo, hay_informa_fichero=bool(_txt((p.get("informa") or {}).get("fichero"))),
                puede_firmar=p.get("estado") == "BORRADOR", puede_firmar_direccion=puede_d, motivo_direccion=motivo_d,
                hotel=_nombre_hotel(p.get("hotel_id")))
