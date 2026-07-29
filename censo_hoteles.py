#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""El censo de hoteles: quien existe y cual esta activo.

FASE 0 de la separacion por hotel.

Por que existe: hasta ahora habia DOS nociones de hotel a la vez.
`hoteles.json` da de alta cada hotel con un `id` estable (tipo `HA3F9C2`), pero
la sesion guardaba el NOMBRE, y los filtros comparaban ese nombre con un
`contains` contra una columna de texto. O sea que "Hotel Sol" se llevaba
tambien las filas de "Hotel Sol Mar". Dos hoteles de un mismo grupo con nombres
parecidos se pisaban entre ellos.

A partir de aqui la identidad de un hotel es su **id**, en un solo sitio. El
nombre se resuelve al pintar, mirando el censo. Los nombres se editan y llevan
acentos; los ids no cambian nunca.

Este modulo NO filtra datos todavia. Solo contesta a "que hoteles hay" y "cual
esta activo". El filtrado se muda a `almacen_datos` en la fase 1.
"""
import json
import os


def _ruta():
    from tenant_dirs import datos_dir
    return os.path.join(datos_dir(), "hoteles.json")


def sembrar_desde_entorno():
    """Recrea el censo desde `YVE_HOTELES_SEED` si esta vacio. Apaño temporal.

    POR QUE HACE FALTA
    Render (plan gratuito) no tiene disco: cada despliegue levanta un sistema de
    ficheros nuevo DESDE EL REPO, asi que todo lo que la aplicacion escribe en
    runtime muere. `hoteles.json` esta commiteado como `[]`, de modo que cada
    despliegue restaura un censo vacio y hay que dar de alta los hoteles otra
    vez para poder verificar nada.

    Ojo con la conclusion facil: sacar el fichero del repo NO arregla esto.
    Pasaria de "restaurado a []" a "no existe", y `hoteles()` devuelve [] en los
    dos casos. El disco efimero es la causa; el fichero commiteado solo hace que
    la perdida sea silenciosa.

    POR QUE POR VARIABLE DE ENTORNO Y NO COMMITEANDO LOS HOTELES
    Decision del usuario, y es la buena: las variables de Render sobreviven a
    los despliegues igual que el repo, pero no meten un fixture en el codigo
    justo despues de haberlo limpiado (fase C), y se quitan de un clic el dia
    que haya persistencia de verdad. El apaño y el producto quedan separados.

    LAS DOS PROPIEDADES QUE LO HACEN SEGURO
      1. Solo siembra si el censo esta VACIO o no existe. Con hoteles dentro no
         toca nada — nunca puede pisar los de un cliente.
      2. Si algo falla (variable mal escrita, JSON roto, disco de solo lectura)
         avisa por consola y sigue. Un apaño de conveniencia no puede tumbar el
         arranque de la aplicacion.

    Se llama UNA vez al arrancar, sin peticion, asi que el tenant que resuelve
    `datos_dir()` es el de `YVE_TENANT` o `default` — el de desarrollo. Los
    tenants de cliente no pasan por aqui.
    """
    crudo = os.environ.get("YVE_HOTELES_SEED", "").strip()
    if not crudo:
        return None

    if hoteles(solo_activos=False):
        return None                      # ya hay censo: no se toca

    try:
        datos = json.loads(crudo)
    except Exception as e:
        print(f"[censo] YVE_HOTELES_SEED no es JSON valido, se ignora: {e}")
        return None
    if not isinstance(datos, list) or not datos:
        print("[censo] YVE_HOTELES_SEED tiene que ser una lista de hoteles; se ignora")
        return None

    limpios = []
    for h in datos:
        if not isinstance(h, dict):
            continue
        hid = str(h.get("id") or "").strip()
        nom = str(h.get("nombre") or "").strip()
        if not hid or not nom:
            continue                     # un hotel a medias no es identidad de nada
        limpios.append({
            "id": hid, "nombre": nom,
            "ciudad":       str(h.get("ciudad") or ""),
            "categoria":    str(h.get("categoria") or "4★"),
            "habitaciones": int(h.get("habitaciones") or 0),
            "grupo":        str(h.get("grupo") or "Principal"),
            "activo":       bool(h.get("activo", True)),
            "modulos":      h.get("modulos") or ["ar", "ap", "drr", "banco", "fb"],
            "creado":       str(h.get("creado") or ""),
        })
    if not limpios:
        print("[censo] YVE_HOTELES_SEED no traia ningun hotel con id y nombre; se ignora")
        return None

    try:
        with open(_ruta(), "w", encoding="utf-8") as fh:
            json.dump(limpios, fh, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[censo] no se pudo escribir el censo sembrado: {e}")
        return None

    print(f"[censo] sembrados {len(limpios)} hoteles desde YVE_HOTELES_SEED: "
          + ", ".join(h["nombre"] for h in limpios))
    return limpios


def hoteles(solo_activos=True):
    """Los hoteles del tenant. Lista vacia si no hay censo o esta corrupto.

    Se descartan las entradas sin `id` o sin `nombre`: un hotel a medias no
    puede ser la identidad de nada, y es mejor no verlo que verlo roto.
    """
    try:
        datos = json.load(open(_ruta(), encoding="utf-8"))
    except Exception:
        return []
    if not isinstance(datos, list):
        return []
    out = []
    for h in datos:
        if not isinstance(h, dict):
            continue
        if not str(h.get("id") or "").strip() or not str(h.get("nombre") or "").strip():
            continue
        if solo_activos and not h.get("activo", True):
            continue
        out.append(h)
    return out


def por_id(hid):
    """El registro completo de un hotel, o None. Busca tambien entre los de baja."""
    hid = str(hid or "").strip()
    if not hid:
        return None
    for h in hoteles(solo_activos=False):
        if str(h.get("id")).strip() == hid:
            return h
    return None


def normalizar(valor):
    """Devuelve el ID del hotel, venga un id o venga un nombre.

    Lo de aceptar el nombre no es por comodidad: las sesiones que ya estaban
    abiertas cuando se despliega esto guardan el NOMBRE. Sin esta tolerancia,
    quien tuviera un hotel elegido en ese momento se encontraria los paneles
    vacios sin entender por que. Con ella, la sesion se cura sola en cuanto
    vuelve a tocar el selector.

    Devuelve '' si no reconoce nada: '' significa "vista de grupo", que es el
    valor por defecto de siempre y el que menos esconde.
    """
    v = str(valor or "").strip()
    if not v:
        return ""
    lista = hoteles(solo_activos=False)
    for h in lista:
        if str(h.get("id")).strip() == v:
            return str(h["id"]).strip()
    for h in lista:
        if str(h.get("nombre", "")).strip().lower() == v.lower():
            return str(h["id"]).strip()
    return ""


def nombre_de(hid):
    """El nombre para enseñar. '' si el id no esta en el censo."""
    h = por_id(hid)
    return str(h.get("nombre", "")) if h else ""


def activo():
    """El ID del hotel activo de la sesion, ya normalizado.

    '' = vista de grupo (todos los hoteles), que es el comportamiento de
    siempre cuando no hay nada elegido.
    """
    try:
        from flask import session, has_request_context
        if has_request_context():
            v = session.get("hotel_activo")
            if v:
                return normalizar(v)
    except Exception:
        pass
    # Sin peticion (los scripts que se lanzan como subproceso) el hotel llega
    # por el entorno, igual que el tenant con YVE_TENANT.
    import os as _os
    return normalizar(_os.environ.get("YVE_HOTEL", ""))


def para_guardar():
    """El hotel con el que etiquetar un documento que entra AHORA.

    La regla, acordada con el usuario:

      0 hoteles en el censo -> ''  (todo como antes; nadie nota nada)
      1 hotel               -> ese, sin preguntar; no hay ambiguedad posible
      2 o mas               -> el hotel activo de la sesion

    Con 2 o mas y ninguno activo devuelve '' — o sea "sin asignar". NO se
    reparte a ojo ni se coge el primero: una factura sin hotel no es "del hotel
    principal", es una factura de la que no sabemos el hotel. Queda visible
    como pendiente de asignar, que es lo honesto. El paso siguiente es que el
    modal de subida lo EXIJA cuando hay mas de un hotel.
    """
    act = activo()
    if act:
        return act
    lista = hoteles()
    if len(lista) == 1:
        return str(lista[0]["id"])
    return ""


def exige_hotel():
    """¿Hace falta elegir hotel para poder guardar algo AHORA?

    Devuelve el MOTIVO (texto para enseñar) o None si se puede seguir.

    Relacion con `para_guardar()`, que va en UNA direccion y no en las dos:

        si esto da motivo  ->  para_guardar() daria ''   (siempre)
        si para_guardar() da ''  ->  esto NO tiene por que dar motivo

    La asimetria es el caso de **0 hoteles**: ahi `para_guardar()` devuelve ''
    y hay que dejar pasar, porque no hay nada que elegir y es el estado de
    cualquier tenant recien creado. O sea: '' significa dos cosas distintas
    —"no hay censo" y "hay censo pero no has elegido"— y solo la segunda se
    bloquea. Escrito aqui porque al implementarlo di por hecha la equivalencia
    en los dos sentidos y el propio test la tumbo.

    Lo que si tiene que cumplirse SIEMPRE: no se bloquea nunca cuando el hotel
    se habria podido asignar solo.

    La regla es la pactada:

      0 hoteles -> None   no hay nada que elegir, todo sigue como siempre
      1 hotel   -> None   no hay ambiguedad, se etiqueta solo
      2 o mas   -> hace falta uno activo

    Los casos de 0 y 1 son los que NO se pueden romper: son el tenant recien
    creado y el hotel suelto, o sea la mayoria.
    """
    if activo():
        return None
    lista = hoteles()
    if len(lista) < 2:
        return None
    return (f"Hay {len(lista)} hoteles y ninguno elegido. Elige el hotel antes de "
            f"procesar: un documento sin hotel no es «del hotel principal», es un "
            f"documento del que no sabemos el hotel.")


def sufijo_fichero():
    """`_<HOTEL>` para meter en el nombre de un fichero, o '' si no hay hotel.

    El DRR no es una tabla de filas: es UN informe por subida, y quien lo lee
    coge "el ultimo". Etiquetar por columna no sirve de nada ahi — el ultimo
    fichero seria el del hotel que subio mas tarde, y el otro hotel se comeria
    sus numeros. Asi que el hotel va en el NOMBRE y cada hotel tiene su ultimo.

    Sin hotel el sufijo es vacio y el nombre queda exactamente como siempre,
    que es lo que mantiene intacto el caso de 0 hoteles.
    """
    h = para_guardar()
    return f"_{h}" if h else ""


def fichero_es_de(nombre, hid):
    """Si el nombre de fichero lleva la marca del hotel `hid`.

    Se busca el id CONCRETO que interesa en vez de intentar sacar "el id que
    sea" con una expresion regular. La primera version adivinaba la forma del
    id (una H y seis de hex) y fallaba con cualquier id que no encajara: es
    dar por hecho un formato que puede cambiar. Aqui no se adivina nada.
    """
    if not hid:
        return False
    n = str(nombre or "")
    return (f"_{hid}_" in n) or n.endswith(f"_{hid}") or (f"_{hid}." in n)


def _plegar(t):
    import unicodedata
    t = unicodedata.normalize("NFKD", str(t or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return " ".join(t.lower().split())


def encaje(nombre):
    """El hotel del censo que mejor encaja con un nombre suelto. '' si ninguno.

    Gana el nombre MAS LARGO que encaje. Sin eso, "Hotel Sol Mar" se lo llevaba
    "Hotel Sol" por ser prefijo suyo — y dos hoteles hermanos de un grupo es
    justo el caso en el que hace falta acertar.
    """
    n = _plegar(nombre)
    if not n or n in ("no_encontrado", "nan", "none"):
        return ""
    censo = [(str(h["id"]), _plegar(h.get("nombre"))) for h in hoteles(solo_activos=False)]
    censo = [(i, c) for i, c in censo if c]

    # 1. Igual clavado: no hay nada que pensar.
    for i, c in censo:
        if c == n:
            return i

    # 2. El censo dentro del documento: "Hotel Sol Mar S.L." trae el nombre y
    #    una coletilla. Gana el MAS LARGO, para que "Hotel Sol Mar" no se lo
    #    lleve "Hotel Sol" por ser prefijo suyo.
    dentro = sorted([(len(c), i) for i, c in censo if c in n], reverse=True)
    if dentro:
        return dentro[0][1]

    # 3. El documento dentro del censo: el papel dice menos de lo que sabemos.
    #    Si encaja con UNO solo, ese. Si encaja con varios ("Sol" vale para
    #    "Hotel Sol" y para "Hotel Sol Mar"), no se sabe: mejor no decir nada
    #    que señalar al hotel equivocado.
    fuera = [i for i, c in censo if n in c]
    return fuera[0] if len(fuera) == 1 else ""


def para_selector():
    """[{id, nombre}] de los hoteles activos, para el desplegable."""
    return [{"id": str(h["id"]), "nombre": str(h["nombre"])} for h in hoteles()]
