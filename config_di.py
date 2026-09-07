# -*- coding: utf-8 -*-
"""config_di.py — Yve.01 · Certificado de doble imposicion (DI): a quien se le pide.

Regla validada por finanzas (7 sep 2026): el certificado de residencia fiscal /
doble imposicion SOLO se pide a OTAs de paises FUERA de la UE que tengan
convenio de doble imposicion con España. Dentro de la UE no procede (Booking.com
B.V. es holandesa: deja de pedirse); sin convenio tampoco (no hay convenio que
aplicar).

Fuente de la lista de convenios: AEAT, "Convenios de doble imposicion firmados
por España" — https://sede.agenciatributaria.gob.es/Sede/normativa-criterios-
interpretativos/fiscalidad-internacional/convenios-doble-imposicion-firmados-espana.html
(leida el 7 sep 2026; la AEAT la actualiza: si entra un convenio nuevo, se añade
aqui). Los nombres van en español tal como los da la AEAT.
"""

# Los 27 de la UE (nombres en español, en minusculas y sin acentos al comparar)
PAISES_UE = {
    "alemania", "austria", "belgica", "bulgaria", "chipre", "croacia", "dinamarca",
    "eslovaquia", "eslovenia", "espana", "estonia", "finlandia", "francia", "grecia",
    "hungria", "irlanda", "italia", "letonia", "lituania", "luxemburgo", "malta",
    "paises bajos", "polonia", "portugal", "chequia", "rumania", "suecia",
}
# (los nombres se comparan sin acentos ni ñ: "españa" -> "espana")

# Convenios de doble imposicion firmados por España (AEAT, 7 sep 2026)
CONVENIOS_ESPANA = {
    "albania", "alemania", "andorra", "arabia saudi", "argelia", "argentina", "armenia",
    "australia", "austria", "azerbaiyan", "barbados", "belgica", "bielorrusia", "bolivia",
    "bosnia y herzegovina", "brasil", "bulgaria", "cabo verde", "canada", "catar", "chequia",
    "chile", "china", "chipre", "colombia", "corea del sur", "costa rica", "croacia", "cuba",
    "dinamarca", "ecuador", "egipto", "el salvador", "emiratos arabes unidos", "eslovaquia",
    "eslovenia", "estados unidos", "estonia", "federacion rusa", "filipinas", "finlandia",
    "francia", "georgia", "grecia", "hong kong", "hungria", "india", "indonesia", "iran",
    "irlanda", "islandia", "israel", "italia", "jamaica", "japon", "kazajstan", "kirguistan",
    "kuwait", "letonia", "lituania", "luxemburgo", "macedonia", "malasia", "malta",
    "marruecos", "mexico", "moldavia", "nigeria", "noruega", "nueva zelanda", "oman",
    "paises bajos", "pakistan", "panama", "paraguay", "polonia", "portugal", "reino unido",
    "republica dominicana", "rumania", "senegal", "serbia", "singapur", "sudafrica",
    "suecia", "suiza", "tailandia", "trinidad y tobago", "tunez", "turquia", "uruguay",
    "uzbekistan", "venezuela", "vietnam",
}

# Cada OTA con el pais desde el que FACTURA (la entidad que emite la liquidacion,
# no el mercado del cliente). Clave: el nombre normalizado que usa el verificador.
OTA_PAIS = {
    "booking.com":  ("Booking.com B.V.", "paises bajos"),
    "booking.es":   ("Booking.com España", "españa"),
    "expedia":      ("Expedia Lodging Partner Services Sàrl", "suiza"),
    "hotels.com":   ("Expedia Lodging Partner Services Sàrl", "suiza"),
    "despegar":     ("Despegar.com", "argentina"),
    "airbnb":       ("Airbnb Ireland UC", "irlanda"),
    "agoda":        ("Agoda Company Pte. Ltd.", "singapur"),
    "trip.com":     ("Trip.com Travel Singapore Pte. Ltd.", "singapur"),
    "trivago":      ("trivago N.V.", "alemania"),
    "hrs":          ("HRS GmbH", "alemania"),
    "hotelbeds":    ("Hotelbeds Spain S.L.U.", "españa"),
}

import unicodedata


def _norm(s):
    return unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().lower().strip()


def pais_de_ota(nombre_ota):
    """(entidad, pais) de la OTA, o (None, None) si no esta en la tabla."""
    k = _norm(nombre_ota)
    if k in OTA_PAIS:
        return OTA_PAIS[k]
    for ota, v in OTA_PAIS.items():
        if ota in k:
            return v
    return (None, None)


def regimen_di(pais):
    """Que toca con ese pais: 'nacional' | 'ue' | 'convenio' | 'sin_convenio' | 'desconocido'.
    Solo 'convenio' pide certificado de doble imposicion."""
    p = _norm(pais)
    if not p:
        return "desconocido"
    if p in ("espana", "spain"):
        return "nacional"
    if p in PAISES_UE:
        return "ue"
    if p in CONVENIOS_ESPANA:
        return "convenio"
    return "sin_convenio"


def requiere_di(pais):
    return regimen_di(pais) == "convenio"
