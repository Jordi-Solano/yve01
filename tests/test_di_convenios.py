# -*- coding: utf-8 -*-
"""Certificado de doble imposicion (b77, regla de finanzas 7 sep 2026): solo se
pide a OTAs de paises FUERA de la UE CON convenio con España.
  - config_di: los 27 de la UE, la lista de convenios de la AEAT, cada OTA con su pais
  - regimen: España -> nacional, UE -> ue, convenio -> convenio, resto -> sin_convenio
  - detector: Booking.com (Paises Bajos) NO_APLICA; Expedia (Suiza) y Agoda (Singapur)
    si piden certificado; OTA no reconocida -> OTA_DESCONOCIDA
  - contratos de grupo: misma regla para el cliente

  python3.12 tests/test_di_convenios.py
  python3.12 tests/test_di_convenios.py --sabotaje
"""
import os
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
os.chdir(BASE)
SABOTAJE = '--sabotaje' in sys.argv


def main():
    fallos = 0

    def ok(cond, msg):
        nonlocal fallos
        print(f"  {'OK ' if cond else 'FALLA'}  {msg}")
        if not cond:
            fallos += 1

    import config_di as C
    if SABOTAJE:
        C.OTA_PAIS['booking.com'] = ('Booking.com B.V.', 'suiza')
    ok(len(C.PAISES_UE) == 27 and 'espana' in C.PAISES_UE and 'paises bajos' in C.PAISES_UE and 'suiza' not in C.PAISES_UE and 'reino unido' not in C.PAISES_UE, f"los 27 de la UE, sin Suiza ni Reino Unido ({len(C.PAISES_UE)})")
    ok(len(C.CONVENIOS_ESPANA) >= 90 and all(p in C.CONVENIOS_ESPANA for p in ('suiza', 'estados unidos', 'reino unido', 'singapur', 'argentina', 'mexico', 'paises bajos')), f"lista de convenios de la AEAT ({len(C.CONVENIOS_ESPANA)} paises)")
    ok('bahamas' not in C.CONVENIOS_ESPANA and 'monaco' not in C.CONVENIOS_ESPANA, 'sin convenio: Bahamas, Monaco')
    casos = [('España', 'nacional'), ('Espana', 'nacional'), ('Países Bajos', 'ue'), ('Irlanda', 'ue'), ('Alemania', 'ue'),
             ('Suiza', 'convenio'), ('Estados Unidos', 'convenio'), ('Singapur', 'convenio'), ('Reino Unido', 'convenio'),
             ('Bahamas', 'sin_convenio'), ('', 'desconocido')]
    malos = [(p, C.regimen_di(p)) for p, e in casos if C.regimen_di(p) != e]
    ok(not malos, f"regimen por pais: nacional / ue / convenio / sin_convenio ({malos})")
    ok(C.requiere_di('Suiza') and not C.requiere_di('Países Bajos') and not C.requiere_di('España') and not C.requiere_di('Bahamas'), 'requiere_di: solo fuera de la UE con convenio')
    # cada OTA de la tabla tiene pais y regimen conocido
    sin = [o for o, (ent, pais) in C.OTA_PAIS.items() if C.regimen_di(pais) == 'desconocido' or not ent]
    ok(not sin, f"cada OTA de la tabla tiene entidad y pais con regimen conocido ({sin})")
    ok(C.pais_de_ota('Booking.com')[1] == 'paises bajos' and C.pais_de_ota('BOOKING.COM')[1] == 'paises bajos' and C.pais_de_ota('booking.es')[1] == 'españa', 'Booking.com -> Paises Bajos (mayusculas da igual), booking.es -> España')
    ok(C.pais_de_ota('OTA Rara')[1] is None, 'una OTA fuera de la tabla no tiene pais')

    # el detector
    import pandas as pd
    import detector_doble_imposicion as DI
    if SABOTAJE:
        import importlib; importlib.reload(DI)
    def an(ota, estado='CORRECTO'):
        return DI.analizar_factura(pd.Series({'archivo': 'no-existe.pdf', 'nombre_ota': ota, 'mercado': 'internacional', 'estado': estado}))
    b = an('booking.com'); e = an('expedia'); a = an('agoda'); es = an('booking.es'); x = an('otaquenadieconoce'); ab = an('airbnb')
    ok(b['estado_di'] == 'NO_APLICA' and b['tipo_mercado'] == 'ue' and 'paises bajos' in b['motivo_di'], f"Booking.com B.V.: de la UE, NO se pide certificado ({b['estado_di']}, {b.get('motivo_di')})")
    ok(ab['estado_di'] == 'NO_APLICA' and ab['tipo_mercado'] == 'ue', f"Airbnb Ireland: de la UE, no se pide ({ab['estado_di']})")
    ok(e['estado_di'] == 'FALTA_CERTIFICADO_DI' and e['tipo_mercado'] == 'extranjera' and e['pais_ota'] == 'suiza', f"Expedia (Suiza, convenio): se pide y falta en el PDF ({e['estado_di']})")
    ok(a['estado_di'] == 'FALTA_CERTIFICADO_DI' and a['pais_ota'] == 'singapur', f"Agoda (Singapur, convenio): se pide ({a['estado_di']})")
    ok(es['estado_di'] == 'NO_APLICA' and es['tipo_mercado'] == 'nacional', f"booking.es: nacional ({es['estado_di']})")
    ok(x['estado_di'] == 'OTA_DESCONOCIDA', f"OTA no reconocida: OTA_DESCONOCIDA ({x['estado_di']})")
    ok(DI.clasificar_mercado('rara', 'nacional') == 'nacional' and DI.clasificar_mercado('rara', 'internacional') == 'desconocida', 'columna Mercado: "nacional" vale, "internacional" sin pais no decide')
    # contratos de grupo
    src = open('lector_contratos_grupo.py', encoding='utf-8').read()
    ok('requiere_di(cli.get("pais", ""))' in src and 'not in ("", "españa", "espana", "spain")' not in src, 'contratos de grupo: la DI del cliente sigue la misma regla (config_di), no "todo lo que no es España"')
    ok('extranjera = ota in ("Expedia", "Agoda")' in open('demo_generator.py', encoding='utf-8').read(), 'el demo ya no marca a Booking con certificado pendiente')
    ok('OTAS_EXTRANJERAS' not in open('detector_doble_imposicion.py', encoding='utf-8').read(), 'la lista vieja de "extranjeras" ha desaparecido del detector')

    print()
    if SABOTAJE:
        print('SABOTAJE: se esperaban fallos' if fallos else '*** SABOTAJE SIN EFECTO ***')
        sys.exit(0 if fallos else 1)
    print('TODO OK' if not fallos else f'{fallos} FALLOS')
    sys.exit(1 if fallos else 0)


if __name__ == '__main__':
    main()
