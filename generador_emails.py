"""
generador_emails.py — Yve.01
Lee los reportes de verificacion y doble imposicion,
genera emails profesionales en espanol con Claude para cada
factura con DISCREPANCIA o FALTA_CERTIFICADO_DI, y los
guarda como .txt en reportes/emails_pendientes/.
"""

import os, re, glob, textwrap
from datetime import date
import pandas as pd
from dotenv import load_dotenv
import anthropic

# ── Rutas ──────────────────────────────────────────────────────────────────
BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
REPORTES_DIR = os.path.join(BASE_DIR, "reportes")
EMAILS_DIR   = os.path.join(REPORTES_DIR, "emails_pendientes")
os.makedirs(EMAILS_DIR, exist_ok=True)

FECHA_HOY = date.today().strftime("%Y%m%d")
NF = "NO_ENCONTRADO"

# ── Cargar API key ──────────────────────────────────────────────────────────
load_dotenv(os.path.join(BASE_DIR, ".env"))
API_KEY = os.getenv("ANTHROPIC_API_KEY")
if not API_KEY:
    raise EnvironmentError(
        "No se encontro ANTHROPIC_API_KEY en el archivo .env\n"
        "Asegurate de que .env contiene: ANTHROPIC_API_KEY=tu_clave"
    )
cliente_claude = anthropic.Anthropic(api_key=API_KEY)

# ── Contactos conocidos de OTAs (se pueden ampliar) ────────────────────────
CONTACTOS_OTA = {
    "booking.com":  "finance@booking.com",
    "booking.es":   "finance@booking.com",
    "expedia":      "hotelbilling@expedia.com",
    "hotels.com":   "hotelbilling@hotels.com",
    "despegar":     "contratos@despegar.com",
}


# ── Carga de reportes ───────────────────────────────────────────────────────

def _solo_hotel(df):
    """El hotel elegido manda tambien en los emails de reclamacion.

    De estos reportes salen correos que se le mandan a Booking. Reclamar una
    factura que no es de este hotel no es un fallo de pantalla: es un email
    equivocado saliendo a un tercero.
    """
    try:
        from almacen_datos import solo_del_hotel_activo as _solo
        return _solo(df)
    except Exception:
        return df


def cargar_reporte_di() -> pd.DataFrame:
    """Carga el reporte de doble imposicion mas reciente (hoja Detalle_DI)."""
    excels = sorted(
        glob.glob(os.path.join(REPORTES_DIR, "doble_imposicion_*.xlsx")),
        reverse=True
    )
    if not excels:
        raise FileNotFoundError(
            "No se encontro doble_imposicion_*.xlsx en reportes/.\n"
            "Ejecuta primero detector_doble_imposicion.py"
        )
    ruta = excels[0]
    print(f"  Cargando reporte DI: {os.path.basename(ruta)}")
    return _solo_hotel(pd.read_excel(ruta, sheet_name="Detalle_DI"))


def cargar_reporte_verificacion() -> pd.DataFrame:
    """Carga el reporte de verificacion mas reciente como fallback."""
    excels = sorted(
        glob.glob(os.path.join(REPORTES_DIR, "verificacion_*.xlsx")),
        reverse=True
    )
    if not excels:
        return pd.DataFrame()
    return _solo_hotel(pd.read_excel(excels[0], sheet_name="Detalle"))


# ── Helpers ──────────────────────────────────────────────────────────────────

def limpiar_nombre_archivo(texto: str) -> str:
    """Convierte texto en nombre de archivo valido."""
    limpio = re.sub(r"[^\w\-]", "_", str(texto))
    return re.sub(r"_+", "_", limpio).strip("_").lower()


def determinar_tipo_problema(fila: pd.Series) -> list[str]:
    """Devuelve lista de tipos de problema para esta fila: 'discrepancia', 'certificado_di', o ambos."""
    tipos = []
    if str(fila.get("estado", "")) == "DISCREPANCIA":
        tipos.append("discrepancia")
    if str(fila.get("estado_di", "")) == "FALTA_CERTIFICADO_DI":
        tipos.append("certificado_di")
    return tipos


def formatear_importe(valor) -> str:
    """Formatea un valor numerico como importe en euros."""
    try:
        return f"{float(str(valor).replace(',','.')):,.2f} EUR"
    except Exception:
        return str(valor)


# ── Generacion de emails con Claude ─────────────────────────────────────────

def construir_prompt(fila: pd.Series, tipos: list[str]) -> str:
    """Construye el prompt para Claude con los datos exactos de la factura."""

    ota         = fila.get("nombre_ota", NF)
    hotel       = fila.get("nombre_hotel", NF)
    num_factura = fila.get("numero_factura", NF)
    fecha       = fila.get("fecha", NF)
    per_inicio  = fila.get("periodo_inicio", NF)
    per_fin     = fila.get("periodo_fin", NF)
    pct_pactado = fila.get("porcentaje_pactado", NF)
    pct_factura = fila.get("porcentaje_factura", NF)
    disc_eur    = fila.get("discrepancia_euros", NF)
    importe_bruto = fila.get("importe_bruto", NF)

    # Formatear importe de discrepancia
    try:
        disc_fmt = f"{float(str(disc_eur).replace(',','.')):,.2f} EUR" if disc_eur not in (NF, None) else NF
    except Exception:
        disc_fmt = str(disc_eur)

    contacto_ota = CONTACTOS_OTA.get(str(ota).lower(), f"departamento.facturacion@{str(ota).lower().replace(' ','')}.com")

    secciones = []

    if "discrepancia" in tipos:
        secciones.append(
            f"PROBLEMA 1 — DISCREPANCIA DE COMISION:\n"
            f"  - Porcentaje pactado en contrato: {pct_pactado}%\n"
            f"  - Porcentaje aplicado en factura: {pct_factura}%\n"
            f"  - Diferencia: {float(str(pct_factura).replace(',','.')) - float(str(pct_pactado).replace(',','.')): .1f} puntos porcentuales\n"
            f"  - Importe facturado en exceso: {disc_fmt}\n"
            f"  - Ventas brutas del periodo: {importe_bruto}\n"
        )

    if "certificado_di" in tipos:
        secciones.append(
            f"PROBLEMA 2 — CERTIFICADO DE DOBLE IMPOSICION:\n"
            f"  - La OTA ({ota}) es una entidad extranjera sujeta al Convenio de Doble Imposicion.\n"
            f"  - No se ha recibido el certificado de residencia fiscal / WHT certificate\n"
            f"    correspondiente al periodo {per_inicio} - {per_fin}.\n"
            f"  - Sin este certificado, el hotel no puede contabilizar la factura correctamente\n"
            f"    ni acreditar la exencion ante la Agencia Tributaria espanola.\n"
        )

    prompt = textwrap.dedent(f"""
    Eres el responsable financiero de {hotel}, un hotel en Espana.
    Debes redactar un email profesional en espanol dirigido al departamento de facturacion de {ota}
    (correo: {contacto_ota}) para comunicar los siguientes problemas con una factura recibida.

    DATOS DE LA FACTURA:
      - Numero de factura: {num_factura}
      - Fecha de emision: {fecha}
      - Periodo facturado: {per_inicio} al {per_fin}
      - Hotel afectado: {hotel}

    PROBLEMAS DETECTADOS:
    {chr(10).join(secciones)}

    INSTRUCCIONES PARA EL EMAIL:
    1. Tono: formal, profesional, no acusatorio. Tratar como posible error administrativo.
    2. Saludo formal al departamento de facturacion de {ota}.
    3. Referenciar claramente el numero de factura y el periodo en el primer parrafo.
    4. Explicar cada problema de forma clara con los datos exactos proporcionados.
    5. Para discrepancias: solicitar nota de credito o factura rectificativa por el importe exacto.
    6. Para certificado DI: solicitar el certificado de residencia fiscal o WHT certificate
       para el periodo indicado, con plazo de respuesta de 15 dias habiles.
    7. Cierre profesional con datos de contacto genericos [NOMBRE_CONTACTO] y [TELEFONO].
    8. Firma: Departamento Financiero, {hotel}.
    9. Longitud: entre 200 y 350 palabras. Sin excesos ni relleno.
    10. NO incluyas asunto en el email — solo el cuerpo del mensaje.

    Responde UNICAMENTE con el cuerpo del email, sin explicaciones adicionales.
    """).strip()

    return prompt, contacto_ota


def generar_email_claude(prompt: str) -> str:
    """Llama a Claude claude-sonnet-4-6 y devuelve el cuerpo del email."""
    mensaje = cliente_claude.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        messages=[{"role": "user", "content": prompt}]
    )
    return mensaje.content[0].text.strip()


def generar_reclamacion_ota(fila, idioma="es", firma_nombre="", firma_entidad=""):
    """Redacta el email de reclamación de comisión OTA. Devuelve {'asunto':.., 'cuerpo':..}.
    Tono firme y profesional, HUMANO (lo lee una persona de la OTA), con cifras exactas y
    petición clara de devolución. No suena a plantilla ni agresivo."""
    import json as _json
    g = lambda k, d="": fila.get(k, d)

    # ── Sin datos no se redacta ────────────────────────────────────────────
    # Antes se llamaba a la IA con los huecos vacios y esta, logicamente,
    # respondia PIDIENDO los datos... y ese texto se enviaba como si fuera una
    # reclamacion. Si falta algo, se falla aqui y se dice exactamente que falta.
    _OBLIGATORIOS = {
        "nombre_ota": "nombre de la OTA",
        "numero_factura": "numero de factura",
        "nombre_hotel": "hotel",
        "porcentaje_pactado": "comision pactada en contrato",
        "porcentaje_factura": "comision aplicada por la OTA",
        "discrepancia_euros": "importe a devolver",
    }
    _vacio = lambda v: v is None or str(v).strip() in ("", "NO_ENCONTRADO", "nan", "None")
    _faltan = [txt for k, txt in _OBLIGATORIOS.items() if _vacio(fila.get(k))]
    if _faltan:
        raise ValueError(
            "No se puede redactar la reclamacion: faltan datos (" + ", ".join(_faltan) + "). "
            "Revisa el informe de verificacion; NO se redacta un email sin cifras."
        )

    ota = g("nombre_ota")
    num = g("numero_factura")
    reserva = g("numero_reserva")
    hotel = g("nombre_hotel")
    periodo = (str(g("periodo_inicio")) + " - " + str(g("periodo_fin"))).strip(" -")
    bruto = g("importe_bruto"); pct_pactado = g("porcentaje_pactado")
    pct_cobrado = g("porcentaje_factura"); dif_pp = g("diferencia_pp")
    a_devolver = g("discrepancia_euros")
    # ── Firma ──────────────────────────────────────────────────────────────
    # Se rellena con el usuario LOGUEADO y con su hotel/grupo (los resuelve
    # reclamaciones_ota._firmante). Lo que no se sepa se queda como marcador
    # visible en el borrador para que el controller lo complete: mejor un hueco
    # que un dato inventado o el usuario de login.
    _fn = str(firma_nombre or "").strip()
    _fe = str(firma_entidad or "").strip()
    if _fn and _fe:
        _firma = ("- Cierra con esta firma EXACTA, en dos lineas y sin anadir cargos ni\n"
                  "  datos de contacto:\n    " + _fn + "\n    " + _fe)
    elif _fn:
        _firma = ("- Cierra firmando unicamente con este nombre, tal cual, sin anadir\n"
                  "  hotel ni cargo:\n    " + _fn)
    elif _fe:
        _firma = ("- Cierra con la firma en dos lineas: deja [Nombre] literal en la primera\n"
                  "  (lo rellena una persona) y debajo:\n    " + _fe)
    else:
        _firma = "- Para la firma usa los marcadores [Nombre] y [Hotel] (no inventes datos personales)."

    idiomas = {"es":"espanol","en":"ingles","ca":"catalan","fr":"frances","de":"aleman","it":"italiano","pt":"portugues"}
    idioma_nombre = idiomas.get(str(idioma).lower(), "espanol")
    _res = (" - Reserva: " + str(reserva)) if (reserva and reserva != NF) else ""
    prompt = f"""Eres el/la controller financiero de un hotel y escribes a {ota} para reclamar una comision mal aplicada.

DATOS EXACTOS (usalos tal cual, no inventes ni redondees):
- Factura de la OTA: {num}{_res}
- Hotel: {hotel}
- Periodo: {periodo}
- Importe bruto de las reservas: {bruto} EUR
- Comision pactada en contrato: {pct_pactado} %
- Comision aplicada por {ota}: {pct_cobrado} %
- Diferencia: {dif_pp} puntos porcentuales
- Importe cobrado de mas, a devolver: {a_devolver} EUR

TONO Y ESTILO (IMPORTANTE, lo lee una PERSONA del equipo de {ota}):
- Humano y profesional. Firme pero cordial. Nada agresivo, sin amenazas, sin sonar a plantilla ni a robot.
- Ve al grano: expon la discrepancia con los numeros exactos y pide de forma clara la regularizacion/abono de {a_devolver} EUR.
- Cierra ofreciendote a aclarar dudas y a facilitar la copia del contrato con la comision pactada.
- Escribe TODO en {idioma_nombre}.
{_firma}

Devuelve el email EXACTAMENTE en este formato, sin nada mas antes ni despues:
ASUNTO: <asunto en una sola linea>

<cuerpo del email, puede ocupar varias lineas>"""
    txt = generar_email_claude(prompt).strip()
    asunto = ""; cuerpo = ""
    _lines = txt.split("\n")
    for _idx, _ln in enumerate(_lines):
        if _ln.strip().lower().startswith("asunto:"):
            asunto = _ln.split(":", 1)[1].strip()
            cuerpo = "\n".join(_lines[_idx + 1:]).strip()
            break
    if not asunto:
        # La IA no ha devuelto el formato pedido -> lo que ha escrito NO es un
        # email de reclamacion. Antes se le inventaba un asunto y se guardaba su
        # texto tal cual: asi fue como salio un correo PIDIENDO datos. La senal
        # estaba aqui y se ignoraba.
        raise ValueError(
            "La IA no ha devuelto un email en el formato esperado (falta la linea "
            "ASUNTO:). No se guarda el borrador. Respuesta recibida: "
            + txt[:200].replace("\n", " ")
        )
    cuerpo = re.sub(r"^\s*(cuerpo|body)\s*:\s*", "", cuerpo, flags=re.I).strip()
    return {"asunto": asunto, "cuerpo": cuerpo}


def construir_archivo_email(fila: pd.Series, cuerpo: str, tipos: list[str], contacto_ota: str) -> str:
    """Construye el contenido completo del archivo .txt con cabecera y cuerpo."""
    ota         = fila.get("nombre_ota", NF)
    hotel       = fila.get("nombre_hotel", NF)
    num_factura = fila.get("numero_factura", NF)
    per_inicio  = fila.get("periodo_inicio", NF)
    per_fin     = fila.get("periodo_fin", NF)
    disc_eur    = fila.get("discrepancia_euros", NF)

    try:
        disc_fmt = f"{float(str(disc_eur).replace(',','.')):,.2f} EUR"
    except Exception:
        disc_fmt = NF

    tipo_str = " + ".join(t.upper().replace("_", " ") for t in tipos)

    cabecera = textwrap.dedent(f"""
    ================================================================================
    YVE.01 — EMAIL GENERADO AUTOMATICAMENTE
    Fecha generacion : {date.today().strftime("%d/%m/%Y")}
    Tipo de problema : {tipo_str}
    ================================================================================
    PARA     : {contacto_ota}
    DE       : finanzas@{limpiar_nombre_archivo(hotel)}.com  [COMPLETAR]
    ASUNTO   : {_asunto(tipos, ota, num_factura, per_inicio, per_fin)}
    ================================================================================
    DATOS DE REFERENCIA:
      Factura       : {num_factura}
      Periodo       : {per_inicio} — {per_fin}
      OTA           : {ota}
      Hotel         : {hotel}
      Discrepancia  : {disc_fmt if "discrepancia" in tipos else "N/A"}
    ================================================================================

    CUERPO DEL EMAIL:
    --------------------------------------------------------------------------------
    """).lstrip()

    pie = textwrap.dedent(f"""
    --------------------------------------------------------------------------------
    [REVISION REQUERIDA ANTES DE ENVIAR]
    - Completar nombre y telefono del contacto en la firma
    - Verificar correo del remitente
    - Adjuntar copia de la factura original si corresponde
    ================================================================================
    """)

    return cabecera + "\n" + cuerpo + "\n" + pie


def _asunto(tipos, ota, num_factura, per_inicio, per_fin) -> str:
    if "discrepancia" in tipos and "certificado_di" in tipos:
        return f"Factura {num_factura} — Discrepancia de comision y solicitud certificado doble imposicion"
    elif "discrepancia" in tipos:
        return f"Factura {num_factura} — Discrepancia en porcentaje de comision aplicado"
    else:
        return f"Factura {num_factura} — Solicitud certificado de residencia fiscal / doble imposicion"


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("=" * 65)
    print("  Yve.01 — Generador de Emails OTA")
    print("=" * 65)

    # Cargar datos
    try:
        df = cargar_reporte_di()
    except FileNotFoundError as e:
        print(f"\nERROR: {e}")
        return

    # Filtrar facturas que requieren accion
    mask = (df["estado"] == "DISCREPANCIA") | (df["estado_di"] == "FALTA_CERTIFICADO_DI")
    df_accion = df[mask].copy()

    if df_accion.empty:
        print("\nNo hay facturas con DISCREPANCIA o FALTA_CERTIFICADO_DI. Sin emails que generar.")
        return

    print(f"\n  Facturas con problemas: {len(df_accion)}")
    print(f"  Emails a generar: {len(df_accion)}\n")

    emails_generados = []
    importe_total_reclamado = 0.0

    for idx, (_, fila) in enumerate(df_accion.iterrows(), 1):
        ota         = fila.get("nombre_ota", NF)
        num_factura = fila.get("numero_factura", NF)
        tipos       = determinar_tipo_problema(fila)

        print(f"  [{idx}/{len(df_accion)}] {ota} | Factura {num_factura}")
        print(f"         Problemas: {', '.join(t.upper() for t in tipos)}")
        print(f"         Generando email con Claude claude-sonnet-4-6...", end="", flush=True)

        try:
            prompt, contacto_ota = construir_prompt(fila, tipos)
            cuerpo = generar_email_claude(prompt)
            contenido_txt = construir_archivo_email(fila, cuerpo, tipos, contacto_ota)

            # Nombre del archivo
            nombre_arch = f"{limpiar_nombre_archivo(ota)}_{limpiar_nombre_archivo(num_factura)}.txt"
            ruta_txt = os.path.join(EMAILS_DIR, nombre_arch)
            with open(ruta_txt, "w", encoding="utf-8") as f:
                f.write(contenido_txt)

            print(f" OK -> {nombre_arch}")
            emails_generados.append(ruta_txt)

            # Sumar importe reclamado si es discrepancia
            if "discrepancia" in tipos:
                try:
                    disc = fila.get("discrepancia_euros", None)
                    importe_total_reclamado += float(str(disc).replace(",", "."))
                except Exception:
                    pass

        except anthropic.APIConnectionError:
            print(f" ERROR: Sin conexion a internet. Verifica tu red.")
        except anthropic.AuthenticationError:
            print(f" ERROR: API key invalida. Verifica .env")
        except Exception as e:
            print(f" ERROR: {e}")

        print()

    # Resumen final
    print("=" * 65)
    print("  RESUMEN EMAILS GENERADOS")
    print("=" * 65)
    print(f"  Emails generados        : {len(emails_generados)}")
    print(f"  Importe total reclamado : {importe_total_reclamado:,.2f} EUR")
    print(f"  Carpeta de salida       : reportes/emails_pendientes/")
    print()
    for ruta in emails_generados:
        print(f"  -> {os.path.basename(ruta)}")
    print("=" * 65)


if __name__ == "__main__":
    main()
