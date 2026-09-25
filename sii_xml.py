# -*- coding: utf-8 -*-
"""sii_xml.py — Yve.01 · libros del SII en el formato de la AEAT (b79).

Genera los dos ficheros XML que se envian al Suministro Inmediato de
Informacion: el de facturas EMITIDAS (`SuministroLRFacturasEmitidas`) y el de
RECIBIDAS (`SuministroLRFacturasRecibidas`), version 1.1, con la estructura
SOAP que espera el servicio (Cabecera con titular y TipoComunicacion A0 =
alta). Salen de `fiscal.calcular` (los libros `sii.expedidas` / `sii.recibidas`).

NADA SE ENVIA: son ficheros descargables. El envio exige certificado digital y
lo hace la gestoria o su programa. Validado por finanzas (7 sep 2026) el
enfoque; los codigos (claves de regimen, tipos de factura, IDType) se toman
de la documentacion tecnica del SII y la gestoria los revisa al primer envio.
"""
from datetime import date
from xml.sax.saxutils import escape

NS_SOAP = "http://schemas.xmlsoap.org/soap/envelope/"
NS_LR = "https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroLR.xsd"
NS_SII = "https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroInformacion.xsd"
VERSION = "1.1"

# prefijos del NIF-IVA que el SII acepta como CodigoPais (UE + XI)
_PAISES_ID = {"AT", "BE", "BG", "CY", "CZ", "DE", "DK", "EE", "EL", "FI", "FR", "HR", "HU", "IE", "IT", "LT",
              "LU", "LV", "MT", "NL", "PL", "PT", "RO", "SE", "SI", "SK", "XI", "GB", "CH", "NO", "US", "SG"}


def _imp(v):
    return f"{float(v or 0):.2f}"


def _fecha_sii(iso):
    """'2026-08-31' -> '31-08-2026' (el SII quiere dd-mm-aaaa)."""
    s = str(iso or "")
    if len(s) >= 10 and s[4] == "-":
        return f"{s[8:10]}-{s[5:7]}-{s[0:4]}"
    return s


def _nif_es(nif):
    n = str(nif or "").upper().replace(" ", "")
    return n[2:] if n.startswith("ES") else n


def _es_espanol(nif):
    n = str(nif or "").upper().replace(" ", "")
    if not n:
        return False
    if n.startswith("ES"):
        return True
    return not (len(n) >= 2 and n[:2].isalpha() and n[:2] in _PAISES_ID)


def _contraparte(nombre, nif):
    """Contraparte del SII: NIF español, o IDOtro (02 = NIF-IVA de otro pais; 06 = otro documento)."""
    nom = escape(str(nombre or "")[:120])
    n = str(nif or "").upper().replace(" ", "")
    if not n:
        return f"<sii:Contraparte><sii:NombreRazon>{nom}</sii:NombreRazon><sii:IDOtro><sii:IDType>06</sii:IDType><sii:ID>SIN NIF</sii:ID></sii:IDOtro></sii:Contraparte>"
    if _es_espanol(n):
        return f"<sii:Contraparte><sii:NombreRazon>{nom}</sii:NombreRazon><sii:NIF>{escape(_nif_es(n))}</sii:NIF></sii:Contraparte>"
    pais = n[:2] if n[:2] in _PAISES_ID else ""
    if pais == "EL":
        pais = "GR"          # Grecia: prefijo EL en el NIF-IVA, codigo de pais GR
    return (f"<sii:Contraparte><sii:NombreRazon>{nom}</sii:NombreRazon><sii:IDOtro>"
            + (f"<sii:CodigoPais>{pais}</sii:CodigoPais>" if pais else "")
            + f"<sii:IDType>{'02' if pais else '06'}</sii:IDType><sii:ID>{escape(n)}</sii:ID></sii:IDOtro></sii:Contraparte>")


def _cabecera(cfg):
    nif = _nif_es(cfg.get("nif_propio") or "")
    razon = escape(str(cfg.get("razon_social") or "")[:120])
    return (f"<sii:Cabecera><sii:IDVersionSii>{VERSION}</sii:IDVersionSii>"
            f"<sii:Titular><sii:NombreRazon>{razon}</sii:NombreRazon><sii:NIF>{escape(nif)}</sii:NIF></sii:Titular>"
            f"<sii:TipoComunicacion>A0</sii:TipoComunicacion></sii:Cabecera>")


def _periodo(mes):
    a, m = str(mes)[:4], str(mes)[5:7]
    return f"<sii:PeriodoLiquidacion><sii:Ejercicio>{a}</sii:Ejercicio><sii:Periodo>{m}</sii:Periodo></sii:PeriodoLiquidacion>"


def _detalle_iva(tipo, base, cuota, etiqueta_cuota, desglose=None):
    """Un DetalleIVA por tipo. `desglose` (b96) = [{tipo, base, cuota}] exacto: un DetalleIVA por
    tramo con SUS importes. Sin desglose, `tipo` puede venir como '10/21' y la base se reparte a
    partes iguales (aproximacion de antes, solo para lo que no trae desglose)."""
    if desglose:
        return "".join(
            f"<sii:DetalleIVA><sii:TipoImpositivo>{_imp(d.get('tipo'))}</sii:TipoImpositivo><sii:BaseImponible>{_imp(d.get('base'))}</sii:BaseImponible>"
            f"<sii:{etiqueta_cuota}>{_imp(d.get('cuota'))}</sii:{etiqueta_cuota}></sii:DetalleIVA>"
            for d in desglose)
    tipos = str(tipo).split("/")
    if len(tipos) == 1:
        return (f"<sii:DetalleIVA><sii:TipoImpositivo>{_imp(tipos[0])}</sii:TipoImpositivo><sii:BaseImponible>{_imp(base)}</sii:BaseImponible>"
                f"<sii:{etiqueta_cuota}>{_imp(cuota)}</sii:{etiqueta_cuota}></sii:DetalleIVA>")
    # dos tipos sin desglose exacto: se reparte la base a partes iguales (lo marca la descripcion)
    out = []
    n = len(tipos)
    for t in tipos:
        b = round(float(base) / n, 2)
        out.append(f"<sii:DetalleIVA><sii:TipoImpositivo>{_imp(t)}</sii:TipoImpositivo><sii:BaseImponible>{_imp(b)}</sii:BaseImponible>"
                   f"<sii:{etiqueta_cuota}>{_imp(b * float(t) / 100)}</sii:{etiqueta_cuota}></sii:DetalleIVA>")
    return "".join(out)


def xml_emitidas(res, cfg):
    """Fichero SII de facturas emitidas (SuministroLRFacturasEmitidas)."""
    mes = res["mes"]; nif_prop = _nif_es(cfg.get("nif_propio") or "")
    regs = []
    for f in res["sii"]["expedidas"]:
        tipo = str(f.get("tipo_factura") or "F1")
        desc = escape((f"{f.get('origen', '')} · {f.get('nombre', '')}")[:500])
        contra = "" if tipo == "F4" else _contraparte(f.get("nombre"), f.get("nif"))
        regs.append(
            f"<siiLR:RegistroLRFacturasEmitidas>{_periodo(mes)}"
            f"<siiLR:IDFactura><sii:IDEmisorFactura><sii:NIF>{escape(nif_prop)}</sii:NIF></sii:IDEmisorFactura>"
            f"<sii:NumSerieFacturaEmisor>{escape(str(f.get('numero') or 's/n')[:60])}</sii:NumSerieFacturaEmisor>"
            f"<sii:FechaExpedicionFacturaEmisor>{_fecha_sii(f.get('fecha'))}</sii:FechaExpedicionFacturaEmisor></siiLR:IDFactura>"
            f"<siiLR:FacturaExpedida><sii:TipoFactura>{tipo}</sii:TipoFactura>"
            f"<sii:ClaveRegimenEspecialOTrascendencia>{f.get('clave_regimen') or '01'}</sii:ClaveRegimenEspecialOTrascendencia>"
            f"<sii:ImporteTotal>{_imp(f.get('total'))}</sii:ImporteTotal><sii:DescripcionOperacion>{desc}</sii:DescripcionOperacion>"
            f"{contra}"
            f"<sii:TipoDesglose><sii:DesgloseFactura><sii:Sujeta><sii:NoExenta><sii:TipoNoExenta>S1</sii:TipoNoExenta><sii:DesgloseIVA>"
            f"{_detalle_iva(f.get('tipo'), f.get('base'), f.get('cuota'), 'CuotaRepercutida', f.get('desglose'))}"
            f"</sii:DesgloseIVA></sii:NoExenta></sii:Sujeta></sii:DesgloseFactura></sii:TipoDesglose>"
            f"</siiLR:FacturaExpedida></siiLR:RegistroLRFacturasEmitidas>")
    return _envuelve("SuministroLRFacturasEmitidas", _cabecera(cfg) + "".join(regs))


def xml_recibidas(res, cfg):
    """Fichero SII de facturas recibidas (SuministroLRFacturasRecibidas)."""
    mes = res["mes"]
    hoy = date.today().isoformat()
    regs = []
    for f in res["sii"]["recibidas"]:
        nif = str(f.get("nif") or "").upper().replace(" ", "")
        if nif and _es_espanol(nif):
            emisor = f"<sii:NIF>{escape(_nif_es(nif))}</sii:NIF>"
        elif nif:
            pais = nif[:2] if nif[:2] in _PAISES_ID else ""
            pais = "GR" if pais == "EL" else pais
            emisor = ("<sii:IDOtro>" + (f"<sii:CodigoPais>{pais}</sii:CodigoPais>" if pais else "")
                      + f"<sii:IDType>{'02' if pais else '06'}</sii:IDType><sii:ID>{escape(nif)}</sii:ID></sii:IDOtro>")
        else:
            emisor = "<sii:IDOtro><sii:IDType>06</sii:IDType><sii:ID>SIN NIF</sii:ID></sii:IDOtro>"
        isp = str(f.get("inversion_sujeto_pasivo") or "N").upper() == "S"
        det = _detalle_iva(f.get("tipo"), f.get("base"), f.get("cuota"), "CuotaSoportada")
        desglose = (f"<sii:DesgloseFactura><sii:InversionSujetoPasivo>{det}</sii:InversionSujetoPasivo></sii:DesgloseFactura>" if isp
                    else f"<sii:DesgloseFactura><sii:DesgloseIVA>{det}</sii:DesgloseIVA></sii:DesgloseFactura>")
        desc = escape((f"{f.get('origen', '')} · {f.get('nombre', '')}")[:500])
        regs.append(
            f"<siiLR:RegistroLRFacturasRecibidas>{_periodo(mes)}"
            f"<siiLR:IDFactura><sii:IDEmisorFactura>{emisor}</sii:IDEmisorFactura>"
            f"<sii:NumSerieFacturaEmisor>{escape(str(f.get('numero') or 's/n')[:60])}</sii:NumSerieFacturaEmisor>"
            f"<sii:FechaExpedicionFacturaEmisor>{_fecha_sii(f.get('fecha'))}</sii:FechaExpedicionFacturaEmisor></siiLR:IDFactura>"
            f"<siiLR:FacturaRecibida><sii:TipoFactura>{f.get('tipo_factura') or 'F1'}</sii:TipoFactura>"
            f"<sii:ClaveRegimenEspecialOTrascendencia>{f.get('clave_regimen') or '01'}</sii:ClaveRegimenEspecialOTrascendencia>"
            f"<sii:DescripcionOperacion>{desc}</sii:DescripcionOperacion>"
            f"{desglose}"
            f"{_contraparte(f.get('nombre'), f.get('nif'))}"
            f"<sii:FechaRegContable>{_fecha_sii(hoy)}</sii:FechaRegContable>"
            f"<sii:CuotaDeducible>{_imp(f.get('cuota'))}</sii:CuotaDeducible>"
            f"</siiLR:FacturaRecibida></siiLR:RegistroLRFacturasRecibidas>")
    return _envuelve("SuministroLRFacturasRecibidas", _cabecera(cfg) + "".join(regs))


def _envuelve(raiz, cuerpo):
    return ('<?xml version="1.0" encoding="UTF-8"?>\n'
            f'<soapenv:Envelope xmlns:soapenv="{NS_SOAP}" xmlns:siiLR="{NS_LR}" xmlns:sii="{NS_SII}">'
            f'<soapenv:Header/><soapenv:Body><siiLR:{raiz}>{cuerpo}</siiLR:{raiz}></soapenv:Body></soapenv:Envelope>\n')


def generar(res, cfg, libro):
    """(xml, nombre_fichero) para 'emitidas' o 'recibidas'."""
    if libro == "recibidas":
        return xml_recibidas(res, cfg), f"SII_recibidas_{res['mes']}.xml"
    return xml_emitidas(res, cfg), f"SII_emitidas_{res['mes']}.xml"
