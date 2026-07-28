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


# Casa "clave": "valor" Y 'clave': 'valor'. La version anterior solo miraba la
# comilla doble y por eso daba 74 donde hay 218: la mayoria de la tabla esta
# escrita con comilla simple. Una medida que enseña un tercio de la realidad
# es peor que no medir.
_PAR = re.compile(r"""(?:"((?:[^"\\]|\\.)*)"|'((?:[^'\\]|\\.)*)')\s*:\s*"""
                  r"""(?:"((?:[^"\\]|\\.)*)"|'((?:[^'\\]|\\.)*)')""")


def strmap(src):
    """{idioma: {español: traducción}} de `_i18nStrMap`."""
    cuerpo = _bloque_llaves(src, 'var _i18nStrMap')
    out = {}
    for lang in IDIOMAS:
        trozo = _bloque_llaves(cuerpo, '\n  %s:' % lang)
        pares = [((a or b), (c or d)) for a, b, c, d in _PAR.findall(trozo)]
        out[lang] = dict(pares)
    return out


def emoji_frontend(src):
    """{emoji: veces} en el HTML que sirve dashboard.py."""
    a = src.index('HTML = r"""')
    b = src.index('"""', a + 11)
    fuera = {0xFE0F}
    cuenta = {}
    for ch in src[a:b]:
        o = ord(ch)
        if o in fuera or 0x2500 <= o <= 0x257F:
            continue
        if (0x2190 <= o <= 0x2BFF or 0x1F000 <= o <= 0x1FAFF):
            cuenta[ch] = cuenta.get(ch, 0) + 1
    return cuenta


def tabla_iconos():
    """Los emoji que la tabla convierte en icono SVG."""
    p = os.path.join(BASE, 'static', 'yve-icons.js')
    if not os.path.exists(p):
        return set()
    t = open(p, encoding='utf-8').read()
    i = t.index('var MAP={')
    j = t.index('};', i)
    out = set()
    for esc in re.findall(r"'((?:\\u[0-9a-fA-F]{4})+)'\s*:", t[i:j]):
        out.add(esc.encode().decode('unicode_escape')
                   .encode('utf-16', 'surrogatepass').decode('utf-16'))
    return out


# Se quedan como caracteres a proposito: las banderas del selector de idioma,
# los puntos de estado (el color ES el dato) y los simbolos tipograficos, que
# ya salen monocromos con la fuente de la pagina.
CRUDOS_A_PROPOSITO = set('\u2190\u2191\u2192\u2193\u2194\u27a4\u25cf\u25cb\u25a0\u25a1'
                         '\u25b2\u25bc\u25b6\u25c0\u2605\u2606\u2b1b\u2b1c\u2022\u2026'
                         '\U0001f7e2\U0001f7e1\U0001f534\U0001f535\U0001f7e0\U0001f7e3')
CRUDOS_A_PROPOSITO |= {chr(c) for c in range(0x1F1E6, 0x1F200)}   # banderas


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


_RX_PREF = re.compile(r'^([^\w\s¿¡«"\'(\[.–—-]+)\s')


def _pref(s):
    m = _RX_PREF.match(s or '')
    return m.group(1) if m else ''


def _emoji_traducciones():
    """Claves cuyo emoji no coincide entre el español del HTML y los .json.

    Es el fallo que hacia que "👥 Administración" se quedara en "Administració"
    sin icono: la traduccion sustituye el textContent entero, asi que si no
    lleva el emoji, el icono se va con el.
    """
    src = _fuente()
    a = src.index('HTML = r"""')
    b = src.index('"""', a + 11)
    es = {}
    for m in re.finditer(r'data-i18n="([^"]+)"[^>]*>([^<]{1,80})<', src[a:b]):
        es.setdefault(m.group(1), m.group(2).strip())
    D = os.path.join(BASE, 'static', 'i18n')
    if not os.path.isdir(D):
        return {}
    malas = {}
    for f in sorted(x for x in os.listdir(D) if x.endswith('.json')):
        datos = json.load(open(os.path.join(D, f), encoding='utf-8'))
        for k, t in es.items():
            v = datos.get(k)
            if not isinstance(v, str) or not v.strip():
                continue
            if _pref(t) != _pref(v):
                malas.setdefault(k, []).append(f[:-5])
    return {k: ','.join(v) for k, v in malas.items()}


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
    # La cuenta real, medida en el navegador, es 218. Si este script vuelve a
    # medir de menos (le paso con las comillas simples), que se note aqui.
    MINIMO = 200
    if base and base < MINIMO:
        fallos.append('_i18nStrMap[en] mide %d y deberia pasar de %d: '
                      'probablemente el que esta roto es este script' % (base, MINIMO))
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
    print('  EMOJI QUE SE PIERDEN AL TRADUCIR')
    print(LINEA)
    perdidos = _emoji_traducciones()
    if perdidos:
        for k, quien in sorted(perdidos.items())[:12]:
            print('  %-24s %s' % (k, quien))
        print('  %d clave(s): el icono desaparece al cambiar de idioma.' % len(perdidos))
        fallos.append('%d clave(s) pierden o cambian el emoji al traducirse' % len(perdidos))
    else:
        print('  Ninguna. Cada traducción lleva el mismo emoji que el español.')

    print()
    print(LINEA)
    print('  EMOJI SIN ICONO  (salen como emoji del sistema)')
    print(LINEA)
    usados = emoji_frontend(src)
    tabla = tabla_iconos()
    sueltos = {c: n for c, n in usados.items()
               if c not in tabla and c not in CRUDOS_A_PROPOSITO}
    print('  tabla de iconos: %d emoji · frontend usa: %d' % (len(tabla), len(usados)))
    if sueltos:
        for c, n in sorted(sueltos.items(), key=lambda kv: -kv[1])[:15]:
            print('   %s  x%d' % (c, n))
        fallos.append('%d emoji del frontend no tienen icono ni estan en la lista de crudos'
                      % len(sueltos))
    else:
        print('  Ninguno suelto.')

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
