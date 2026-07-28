#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
comprobar_i18n.py — cobertura de traducción y estado de la migración.

Por qué existe: la deuda de traducción de este proyecto es INVISIBLE. Cuando
falta una entrada en `_i18nStrMap`, la app no falla — enseña español y ya está.
Así se llegó a que portugués tuviera 74 entradas donde inglés tiene 444, sin que
nadie lo decidiera. Este script convierte esa deuda silenciosa en un número.

También lleva la cuenta de la migración al sistema nuevo, leyendo
`_PANELES_MIGRADOS` del propio código: la cuenta sale del código y no de un
documento que se queda viejo.

Uso:
    python3 comprobar_i18n.py          # informe
    python3 comprobar_i18n.py --check  # sale con código 1 si algo empeora
"""
import json
import os
import re
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
DASH = os.path.join(BASE, 'dashboard.py')
IDIOMAS = ('en', 'ca', 'fr', 'de', 'it', 'pt')
LINEA = '─' * 68


def _fuente():
    return open(DASH, encoding='utf-8').read()


def _bloque_llaves(src, cabecera):
    """El cuerpo de `cabecera: { ... }` contando llaves, no por regex.

    Una regex perezosa se para en la primera `}` anidada y cuenta de menos, que
    es justo el tipo de medida que engaña.
    """
    i = src.find(cabecera)
    if i < 0:
        return ''
    j = src.index('{', i)
    prof, k = 0, j
    while k < len(src):
        if src[k] == '{':
            prof += 1
        elif src[k] == '}':
            prof -= 1
            if prof == 0:
                return src[j + 1:k]
        k += 1
    return ''


def strmap(src):
    """{idioma: {español: traducción}} de `_i18nStrMap`."""
    cuerpo = _bloque_llaves(src, 'var _i18nStrMap')
    out = {}
    for lang in IDIOMAS:
        trozo = _bloque_llaves(cuerpo, '\n  %s:' % lang)
        pares = re.findall(r'"((?:[^"\\]|\\.)*)"\s*:\s*"((?:[^"\\]|\\.)*)"', trozo)
        out[lang] = dict(pares)
    return out


def sse(src):
    """Longitud de los siete arrays paralelos de mensajes SSE."""
    L = src.split('\n')
    def largo(patron, desde=0):
        i = next((k for k in range(desde, len(L)) if re.match(patron, L[k])), None)
        if i is None:
            return None, 0
        j = next(k for k in range(i + 1, len(L)) if L[k].strip() in ('];', ']', '],'))
        return i, len(re.findall(r'"(?:[^"\\]|\\.)*"', '\n'.join(L[i + 1:j])))
    _, n = largo(r'var _sseFrags = \[')
    out = {'_sseFrags': n}
    k = next((x for x, l in enumerate(L) if 'var _sseTrans = {' in l), 0)
    for lang in IDIOMAS:
        _, out[lang] = largo(r'\s*%s\s*:\s*\[' % lang, k)
    return out


def migrados(src):
    m = re.search(r'var _PANELES_MIGRADOS = \[(.*?)\]', src, re.S)
    return re.findall(r"'([^']+)'", m.group(1)) if m else []


PANELES = ['ar', 'ap', 'drr', 'banco', 'notif', 'fb', 'ar_real', 'multi_hotel']
SUBPANELES_FB = ['fb:resumen', 'fb:inventario', 'fb:mermas', 'fb:recetas']


def main():
    src = _fuente()
    mapa = strmap(src)
    base = len(mapa.get('en', {}))
    hecho = migrados(src)
    fallos = []

    print(LINEA)
    print('  COBERTURA DE TRADUCCIÓN  (_i18nStrMap)')
    print(LINEA)
    print('  %-4s %8s %8s   %s' % ('', 'entradas', 'vs en', 'sin traducir (valor = español)'))
    for lang in IDIOMAS:
        d = mapa.get(lang, {})
        iguales = [k for k, v in d.items() if k == v]
        pct = (len(d) / base * 100) if base else 0
        aviso = '  <<< muy por debajo' if pct < 80 else ''
        print('  %-4s %8d %7.0f%%   %d%s' % (lang, len(d), pct, len(iguales), aviso))
        if pct < 80:
            fallos.append('%s solo cubre el %.0f%% de lo que cubre en' % (lang, pct))

    print()
    print(LINEA)
    print('  MENSAJES SSE  (los siete arrays paralelos, casados por índice)')
    print(LINEA)
    n = sse(src)
    ref = n.get('_sseFrags', 0)
    for k, v in n.items():
        estado = 'OK' if v == ref else '<<< DESALINEADO'
        print('  %-12s %4d  %s' % (k, v, estado))
        if v != ref:
            fallos.append('%s tiene %d entradas y _sseFrags %d' % (k, v, ref))

    print()
    print(LINEA)
    print('  MIGRACIÓN AL SISTEMA NUEVO')
    print(LINEA)
    todos = PANELES + SUBPANELES_FB
    for p in todos:
        print('  [%s] %s' % ('x' if p in hecho else ' ', p))
    print()
    print('  %d de %d migrados. El resto usa la red de seguridad (_pintarYa).' %
          (len(hecho), len(todos)))

    print()
    print(LINEA)
    if fallos:
        print('  %d AVISO(S):' % len(fallos))
        for f in fallos:
            print('   · %s' % f)
    else:
        print('  Sin avisos.')
    print(LINEA)

    if '--check' in sys.argv and fallos:
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
