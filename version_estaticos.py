#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Sello de version de los ficheros estaticos.

Existe para que el navegador no sirva de su cache un fichero viejo. Se calcula
con las fechas de los propios ficheros: cambia en cada despliegue por si solo,
y nadie tiene que acordarse de subir un numero a mano — que es justo el tipo de
paso que se olvida y deja al usuario con la version de la semana pasada.
"""
import glob
import hashlib
import os

_BASE = os.path.dirname(os.path.abspath(__file__))


def sello():
    partes = []
    ficheros = [os.path.join(_BASE, 'static', 'yve-icons.js')]
    ficheros += sorted(glob.glob(os.path.join(_BASE, 'static', 'i18n', '*.json')))
    for p in ficheros:
        try:
            partes.append('%s:%d' % (os.path.basename(p), int(os.path.getmtime(p))))
        except OSError:
            continue
    if not partes:
        return 'dev'
    return hashlib.md5('|'.join(partes).encode('utf-8')).hexdigest()[:8]


# Se calcula una vez al arrancar: en Render cada despliegue reinicia el proceso.
SELLO = sello()
