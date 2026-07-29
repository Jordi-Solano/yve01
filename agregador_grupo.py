"""agregador_grupo.py — la ficha financiera de cada hotel del grupo.

Fase A del Multi-Hotel real. Sustituye a `kpis_hoteles.xlsx`, que solo lo
escribia el generador de demo.


ESTE MODULO SOLO LEE
────────────────────
No escribe un fichero, no crea uno, no borra nada, no toca la sesion. Es la
propiedad que lo hace seguro: es lo primero que mira los datos de TODOS los
hoteles a la vez, asi que un fallo aqui, si escribiera, se los llevaria todos
por delante de una vez. Sin escritura, el peor caso es un numero malo en una
pantalla. Comprobable de un vistazo, y la comprobacion va en el commit:

    grep -nE "to_excel|to_csv|to_json|\\.save\\(|os\\.remove|os\\.rename|shutil\\.|open\\([^)]*[\"'][wax]" agregador_grupo.py

tiene que salir VACIO.

Tampoco toca `oracle_*` ni el clasificador: corre despues de todo, sobre lo ya
guardado.


COMO LEE: UNA VEZ, Y PARTE
──────────────────────────
Lo que NO hace es `facturas_ap(hotel=X)` una vez por hotel. Dos razones, y la
primera es la importante:

1. **El cuadre.** Una particion de un conjunto suma el conjunto. Siempre. Si
   las cajas salen de N+1 lecturas distintas, el total y las partes vienen de
   fotos distintas del disco y basta con que entre una fila en medio para que
   dejen de cuadrar. Partiendo, el cuadre no es una comprobacion que se pasa:
   es aritmetica que no se puede incumplir. La comprobacion que si hacemos
   (`cuadre`) sirve para demostrar que la particion ES una particion.

2. **El coste.** Con 3 hoteles + sin asignar + el total son cinco pasadas
   releyendo y reconsolidando los mismos xlsx. El panel de AP ya tarda 3-4 s
   en el plan gratuito.


DE DONDE SALEN LOS NUMEROS
──────────────────────────
De las MISMAS funciones que usan los paneles, nunca de una copia:

    AP        dashboard.calcular_stats_ap
    AR / OTA  dashboard.calcular_stats
    AR Real   tab_ar_real.facturas_y_stats
    F&B       tab_fb_dashboard.resumen_fb
    Banco     dashboard.stats_banco

Es la condicion de la verificacion de la fase A: si el agregador cuenta por su
cuenta, dara numeros parecidos y nadie sabra cual de los dos esta bien. Dando
los mismos numeros por otro camino, esta bien.

Y no llama a ninguno de los nueve sitios que resuelven el hotel por la SESION
(`solo_del_hotel_activo`). Recorrer hoteles cambiando el hotel de la sesion
seria un truco que revienta con dos usuarios a la vez, y ademas cambiaria lo
que el usuario esta viendo mientras carga.
"""

import almacen_datos as _alm
import censo_hoteles as _censo

CAJA_SIN_ASIGNAR = "sin_asignar"
CAJA_DESCONOCIDO = "desconocido"


# ── Que procedencias pueden entrar en un numero de grupo (fase D) ─────────
#
# El GOP de un hotel puede venir de tres sitios, y NO valen lo mismo:
#
#   medido     el DRR lo trae. Es el dato.
#   derivado   aritmetica sobre datos del propio hotel (sus ingresos por su
#              GOP%, o por el de su presupuesto). Es suyo, aunque calculado.
#   inventado  sacado de la nada — la media del sector y compañia.
#
# Esto vive AQUI, en el agregador, y no solo en el lector del DRR, porque el
# daño de mezclarlas es de grupo: con tres hoteles y uno inventado, el GOP del
# grupo es medio inventado y NADIE puede saber cual es cual mirando la
# pantalla. Peor todavia, la mezcla parece medida.
#
# Hoy ningun camino produce `inventado` (la rama del 22% se borro en la fase D),
# asi que este filtro no quita nada. Esta puesto igual: si mañana alguien añade
# otra estimacion de la nada, la regla ya esta escrita y no depende de que se
# acuerde. La fila hotelera de la fase E entra por aqui.
PROCEDENCIAS_AGREGABLES = frozenset({"medido", "derivado"})


def agregable(procedencia):
    """True si un valor con esa procedencia puede entrar en un numero de grupo.

    Falla en CERRADO: lo desconocido no se agrega. Si mañana aparece una
    procedencia nueva sin que nadie decida que hacer con ella, se queda fuera
    del total en vez de colarse dentro — que es el fallo que no se ve.
    """
    return str(procedencia or "").strip().lower() in PROCEDENCIAS_AGREGABLES


# ── La particion ──────────────────────────────────────────────────────────

def partir_por_hotel(df, ids):
    """Parte `df` en una caja por hotel del censo, mas dos especiales.

    Devuelve un dict {clave: sub_df} con TODAS las claves siempre presentes,
    aunque la caja quede vacia: un hotel sin facturas tiene que salir con 0, no
    desaparecer del panel.

    Es una particion de verdad — cada fila cae en exactamente una caja y no se
    pierde ninguna. De ahi sale el cuadre.

    Las dos cajas especiales NO son lo mismo, y juntarlas seria mentir:

      sin_asignar   el documento no lleva etiqueta de hotel. Es lo de antes de
                    la fase 1, o lo que entro con 0/1 hoteles en el censo.

      desconocido   lleva una etiqueta que no esta en el censo. Pasa de verdad
                    y no es raro: `hoteles.json` va commiteado como [] y Render
                    no tiene disco persistente, asi que cada despliegue borra el
                    censo y los documentos sobreviven a los hoteles que los
                    firmaron. Meter esto en `sin_asignar` diria "no sabemos de
                    quien es" cuando lo que pasa es "sabemos de quien es y no
                    sabemos como se llamaba".

    Sin columna de hotel, todo cae en `sin_asignar`. Es lo correcto y no un
    apaño: si ningun documento lleva etiqueta, ninguno es de ningun hotel. Lo
    que NO se hace es devolver el df entero en cada caja, que es el fallo de
    fase 0 —un filtro que parece filtrar y no filtra— multiplicado por N.
    """
    cajas = {hid: _vacio(df) for hid in ids}
    cajas[CAJA_SIN_ASIGNAR] = _vacio(df)
    cajas[CAJA_DESCONOCIDO] = _vacio(df)

    if df is None or getattr(df, "empty", True):
        return cajas

    col = _alm.COL_HOTEL
    if col not in df.columns:
        cajas[CAJA_SIN_ASIGNAR] = df.copy()
        return cajas

    # `.map` en Python plano, no el accesor `.str` (regla 3): con `.str` los
    # nulos se propagan y las filas sin etiqueta acabarian compartiendo valor.
    etiqueta = df[col].map(_alm._txt)
    conocidos = {_alm._txt(h) for h in ids}

    cajas[CAJA_SIN_ASIGNAR] = df[etiqueta == ""].copy()
    for hid in ids:
        cajas[hid] = df[etiqueta == _alm._txt(hid)].copy()
    cajas[CAJA_DESCONOCIDO] = df[(etiqueta != "") & (~etiqueta.isin(conocidos))].copy()
    return cajas


def _vacio(df):
    """Un df vacio con las mismas columnas, para que los calculadores no fallen."""
    import pandas as pd
    if df is None or not hasattr(df, "iloc"):
        return pd.DataFrame()
    return df.iloc[0:0].copy()


# ── Las fichas de cada seccion ────────────────────────────────────────────

def _ficha_ap(df):
    from dashboard import calcular_stats_ap
    s = calcular_stats_ap(df)
    return {
        "facturas":      s.get("total", 0),
        "importe":       s.get("importe", 0),
        "cuadran":       s.get("matches", 0),
        "discrepancias": s.get("discrepancias", 0),
        "sin_po":        s.get("sin_po", 0),
        "revisar":       s.get("manual", 0),
        "aprobadas":     s.get("aprobadas", 0),
        "rechazadas":    s.get("rechazadas", 0),
    }


def _ficha_ar_ota(df):
    from dashboard import calcular_stats
    s = calcular_stats(df)
    return {
        "facturas":      s.get("total", 0),
        "importe_bruto": s.get("importe_total", 0),
        "correctas":     s.get("correctas", 0),
        "discrepancias": s.get("discrepancias", 0),
        # El numero que justifica el producto. Va con el nombre completo
        # aposta: "reclamable" a secas se confunde con el importe bruto.
        "importe_reclamable": s.get("importe_reclamable", 0),
        "di_pendientes": s.get("di_pendientes", 0),
        "sin_accion":    s.get("sin_accion", 0),
    }


def _ficha_ar_real(df):
    from tab_ar_real import facturas_y_stats
    _lista, s = facturas_y_stats(df)
    return {
        "facturas":  s.get("total_facturas", 0),
        "pendiente": s.get("pendiente", 0),
        "vencido":   s.get("vencido", 0),
        "cobrado":   s.get("cobrado_mes", 0),
    }


def _ficha_fb(df_rec, df_inv, df_ven, df_mer):
    from tab_fb_dashboard import resumen_fb
    _df, _map, r = resumen_fb(df_rec, df_inv, df_ven, df_mer)
    cob = r.get("cobertura", {})
    return {
        "ventas":            r.get("total_ventas", 0),
        "coste_escandallo":  r.get("coste_escandallo", 0),
        "food_cost_pct":     r.get("fc_teorico_pct", 0),
        "coste_mermas":      r.get("coste_mermas", 0),
        "cobertura_pct":     cob.get("pct", 0),
        "ventas_con_receta": cob.get("ventas_con_receta", 0),
    }


def _fb_del_grupo(fichas, ficha_entera):
    """El F&B del grupo: lo aditivo del df entero, el RATIO ponderado.

    Aqui hay una trampa que se ve con datos y no leyendo el codigo. El coste de
    una receta sale de `cost_map`, que se indexa por NOMBRE de ingrediente. Si
    se le pasa el inventario del grupo entero, el tomate del hotel B pisa al del
    hotel A y toda la cadena queda costeada al precio del ultimo que aparezca en
    el fichero.

    Medido con el juego de pruebas: tomate a 2 EUR en un hotel y a 9 EUR en
    otro daba un food cost de grupo del 15,0% —que es el numero del segundo
    hotel, no una media— cuando el ponderado real es 6,4%. Mas del doble, y sin
    ningun aviso.

    Asi que el ratio del grupo NO se calcula sobre el inventario aplanado: se
    pondera desde las cajas, que es lo unico que respeta que cada hotel compra a
    su precio. Sumar costes y ventas y dividir al final es la misma regla que
    hara falta para ocupacion, ADR y GOP% cuando llegue la fila hotelera.

    Lo ADITIVO (ventas, mermas) si sale del df entero, sin tocar: asi su cuadre
    sigue comparando dos caminos independientes y no se vuelve una tautologia.
    """
    coste  = sum(float(f["fb"]["coste_escandallo"])  for f in fichas)
    con_re = sum(float(f["fb"]["ventas_con_receta"]) for f in fichas)
    ventas = float(ficha_entera["ventas"])
    return dict(
        ficha_entera,
        coste_escandallo  = round(coste, 2),
        ventas_con_receta = round(con_re, 2),
        food_cost_pct     = round(coste / con_re * 100, 2) if con_re > 0 else 0.0,
        cobertura_pct     = round(con_re / ventas * 100, 1) if ventas > 0 else 0.0,
        ponderado         = True,
    )


# ── La fila hotelera, del DRR (fase E) ────────────────────────────────────
#
# Un DRR de mas de esto es viejo. El DRR es un informe DIARIO: a los siete dias
# ya no cuenta lo que esta pasando, cuenta lo que pasaba. No se descarta —sus
# numeros siguen siendo reales— pero se dice.
DIAS_PARA_VIEJO = 7

# Que periodo del DRR manda. MTD es del que se habla ("como va el mes"); si el
# DRR no lo trae, el de hoy. El presupuesto NUNCA: es objetivo, no realidad.
_PERIODOS = ("mtd", "today")


def _valor(metricas, clave):
    """El numero de una metrica en el primer periodo que lo tenga, o None."""
    from dashboard import num_drr
    fila = (metricas or {}).get(clave) or {}
    for p in _PERIODOS:
        v = num_drr(fila.get(p))
        if v is not None:
            return v, p
    return None, None


def _ficha_drr(hid):
    """La fila hotelera de un hotel: sus metricas de habitacion, o por que no.

    Tres estados, y los tres se enseñan:

        con_drr     hay DRR y es reciente
        drr_viejo   hay DRR pero de hace mas de DIAS_PARA_VIEJO dias
        sin_drr     ese hotel no ha subido ninguno

    `sin_drr` NO es cero. Es la ausencia del dato, y por eso este hotel se
    queda fuera de las medias del grupo en vez de entrar como un cero. Las dos
    formas de equivocarse aqui son simetricas y las dos mienten: contarlo como
    0 hunde la media y hace que un hotel que no ha subido un papel parezca un
    hotel que va mal; saltarselo en silencio la infla. Lo correcto es lo
    tercero — excluirlo y DECIR el denominador.
    """
    import os
    import time
    from dashboard import drr_del_hotel, _leer_drr_stats

    vacia = {"estado": "sin_drr", "dias_drr": None, "archivo": None,
             "ocupacion_pct": None, "adr": None, "revpar": None,
             "noches_ocupadas": None, "noches_disponibles": None,
             "rooms_revenue": None, "ingresos": None,
             "gop": None, "gop_pct": None, "gop_procedencia": None,
             "dias_oob": None, "periodo": None}

    try:
        ruta = drr_del_hotel(hotel=hid)
    except Exception:
        ruta = None
    if not ruta or not os.path.exists(ruta):
        return vacia

    try:
        s = _leer_drr_stats(ruta)
    except Exception:
        return vacia
    if not s or s.get("error"):
        return vacia

    m = s.get("metricas") or {}
    ocup, per   = _valor(m, "Occupancy %")
    ocupadas, _ = _valor(m, "Rooms Occupied")
    adr, _      = _valor(m, "ADR")
    revpar, _   = _valor(m, "Revenue PAR")
    rooms_rev,_ = _valor(m, "Rooms Revenue")
    ingresos, _ = _valor(m, "Total Revenue")
    gop, _      = _valor(m, "GOP")
    gop_pct, _  = _valor(m, "GOP %")

    # Noches disponibles: el DRR no las trae, pero se deducen de sus propios
    # dos numeros (ocupadas / ocupacion). Se usa esto y no las habitaciones del
    # censo a proposito: el censo no sabe de cuantos DIAS habla el DRR, y
    # multiplicar por un numero de dias supuesto seria inventar el denominador
    # justo en la metrica que mas depende de el.
    disponibles = None
    if ocupadas is not None and ocup:
        p = ocup / 100 if ocup > 1 else ocup
        if p > 0:
            disponibles = ocupadas / p

    dias = None
    try:
        dias = int((time.time() - os.path.getmtime(ruta)) // 86400)
    except Exception:
        pass
    estado = "con_drr"
    if dias is not None and dias > DIAS_PARA_VIEJO:
        estado = "drr_viejo"

    return {
        "estado": estado,
        "dias_drr": dias,
        "archivo": os.path.basename(ruta),
        "periodo": per,
        "ocupacion_pct": round(ocup, 1) if ocup is not None else None,
        "adr": round(adr, 2) if adr is not None else None,
        "revpar": round(revpar, 2) if revpar is not None else None,
        "noches_ocupadas": ocupadas,
        "noches_disponibles": round(disponibles, 1) if disponibles else None,
        "rooms_revenue": rooms_rev,
        "ingresos": ingresos,
        "gop": gop,
        "gop_pct": round(gop_pct, 2) if gop_pct is not None else None,
        # De la fase D. Lo que no sea agregable no entra en el GOP del grupo.
        "gop_procedencia": (s.get("gop_procedencia") or {}).get(per or "mtd"),
        "dias_oob": s.get("dias_oob"),
    }


def _hotelero_del_grupo(fichas):
    """Las medias del grupo, PONDERADAS, y diciendo sobre cuantos hoteles.

    Una media plana de porcentajes es mentira en cuanto los hoteles tienen
    tamaños distintos: un hotel de 400 habitaciones al 60% y uno de 20 al 100%
    no dan 80% de ocupacion del grupo, dan 62%. Asi que se suman los de arriba
    y los de abajo por separado y se divide al final:

        ocupacion = Σ noches ocupadas   / Σ noches disponibles
        ADR       = Σ ingresos de hab.  / Σ noches ocupadas
        RevPAR    = Σ ingresos de hab.  / Σ noches disponibles
        GOP %     = Σ GOP €             / Σ ingresos totales

    Y el denominador VA EN LA RESPUESTA. "Ocupacion del grupo 78%" a secas es
    una trampa si en realidad es "de los 2 hoteles que subieron el DRR". El
    "sobre 2 de 3" es lo que hace que el numero sea defendible.

    El GOP lleva su propio denominador porque puede ser distinto: un hotel
    puede tener DRR (y contar para la ocupacion) y aun asi no traer un GOP
    agregable.
    """
    total = len(fichas)
    con_datos = [f for f in fichas if f["drr"]["estado"] != "sin_drr"]

    def suma(campo, de=None):
        s, hay = 0.0, False
        for f in (de if de is not None else con_datos):
            v = f["drr"].get(campo)
            if v is not None:
                s += float(v); hay = True
        return (s if hay else None)

    ocupadas    = suma("noches_ocupadas")
    disponibles = suma("noches_disponibles")
    rooms_rev   = suma("rooms_revenue")

    # El GOP se agrega SOLO con las procedencias que la fase D deja pasar.
    con_gop = [f for f in con_datos
               if f["drr"].get("gop") is not None
               and agregable(f["drr"].get("gop_procedencia"))]
    gop_eur  = suma("gop", de=con_gop)
    ingresos = suma("ingresos", de=con_gop)

    return {
        "n_hoteles": total,
        "con_datos": len(con_datos),
        "sin_drr":   total - len(con_datos),
        "viejos":    len([f for f in con_datos if f["drr"]["estado"] == "drr_viejo"]),
        "ocupacion_pct": round(ocupadas / disponibles * 100, 1) if (ocupadas and disponibles) else None,
        "adr":           round(rooms_rev / ocupadas, 2) if (rooms_rev and ocupadas) else None,
        "revpar":        round(rooms_rev / disponibles, 2) if (rooms_rev and disponibles) else None,
        "gop_eur":       round(gop_eur, 2) if gop_eur is not None else None,
        "gop_pct":       round(gop_eur / ingresos * 100, 2) if (gop_eur is not None and ingresos) else None,
        "gop_sobre":     len(con_gop),
        "dias_oob":      int(suma("dias_oob") or 0),
        "ponderado":     True,
    }


# ── El cuadre ─────────────────────────────────────────────────────────────

# Solo lo ADITIVO. Los porcentajes y los ratios (food cost, ocupacion, ADR) no
# se suman y no tienen invariante: la media del grupo hay que PONDERARLA, y eso
# va en su fase. Meterlos aqui daria descuadres falsos todo el rato.
_ADITIVOS = [
    ("ap.facturas",             lambda f: f["ap"]["facturas"]),
    ("ap.importe",              lambda f: f["ap"]["importe"]),
    ("ar_ota.facturas",         lambda f: f["ar_ota"]["facturas"]),
    ("ar_ota.importe_bruto",    lambda f: f["ar_ota"]["importe_bruto"]),
    ("ar_ota.reclamable",       lambda f: f["ar_ota"]["importe_reclamable"]),
    ("ar_real.facturas",        lambda f: f["ar_real"]["facturas"]),
    ("ar_real.pendiente",       lambda f: f["ar_real"]["pendiente"]),
    ("fb.ventas",               lambda f: f["fb"]["ventas"]),
    ("fb.coste_mermas",         lambda f: f["fb"]["coste_mermas"]),
]

# Los importes se comparan con tolerancia porque cada caja se redondea a dos
# decimales por su cuenta y N redondeos no suman el redondeo de la suma. Un
# centimo de diferencia entre 3 hoteles es aritmetica de coma flotante; un euro
# ya es una fila perdida. El umbral esta pegado al primero, no al segundo.
_TOLERANCIA = 0.05


def _comprobar_cuadre(fichas, total):
    """Compara la suma de las cajas con el total leido del df ENTERO.

    El total NO sale de sumar las cajas —eso seria comprobar que 2+2 es 2+2—,
    sale de pasar el df sin partir por los mismos calculadores. Asi la
    comprobacion demuestra de verdad que la particion no pierde ni duplica.
    """
    filas, todo_cuadra = [], True
    for nombre, saca in _ADITIVOS:
        try:
            suma = sum(float(saca(f)) for f in fichas)
            tot  = float(saca(total))
        except (KeyError, TypeError, ValueError):
            continue
        cuadra = abs(suma - tot) <= _TOLERANCIA
        todo_cuadra = todo_cuadra and cuadra
        filas.append({"metrica": nombre,
                      "suma_cajas": round(suma, 2),
                      "total_grupo": round(tot, 2),
                      "diferencia": round(suma - tot, 2),
                      "cuadra": cuadra})
    return filas, todo_cuadra


# ── La entrada ────────────────────────────────────────────────────────────

def agregado():
    """La ficha de cada hotel, la del grupo y el cuadre.

    Solo la fila FINANCIERA (AP, AR/OTA, AR Real, F&B) y el banco del grupo.
    La fila hotelera —ocupacion, ADR, RevPAR, GOP— sale del DRR y va en su
    fase: hoy no hay DRR confirmado en produccion y no se inventa nada.
    """
    hoteles = _censo.hoteles()
    ids     = [str(h["id"]) for h in hoteles]
    nombres = {str(h["id"]): str(h.get("nombre") or "") for h in hoteles}
    # FASE F: la categoria y el numero de habitaciones salen del CENSO, que es
    # quien los sabe. El panel viejo los adivinaba del nombre del hotel
    # ("'5★' si el nombre lleva un 5 o dice boutique"), asi que un "Hotel 5 de
    # Mayo" salia como cinco estrellas. Adivinar un dato que esta escrito al
    # lado es de las formas mas tontas de mentir.
    ficha_censo = {str(h["id"]): {"categoria":    str(h.get("categoria") or "") or None,
                                  "habitaciones": h.get("habitaciones") or None,
                                  "ciudad":       str(h.get("ciudad") or "") or None}
                   for h in hoteles}

    # ── Una lectura por fuente, sin filtrar ───────────────────────────────
    from dashboard import cargar_datos_ap_sin_filtrar, cargar_datos_ar_sin_filtrar
    from tab_ar_real import reservas_normalizadas
    from tab_fb_dashboard import _xlsx

    df_ap        = cargar_datos_ap_sin_filtrar()
    df_ar, _meta = cargar_datos_ar_sin_filtrar()
    df_res       = reservas_normalizadas()
    df_ven       = _xlsx("ventas_fb_diarias.xlsx")
    df_inv       = _xlsx("inventario.xlsx")
    df_mer       = _xlsx("mermas.xlsx")
    # El recetario es del GRUPO y no se parte: una cadena comparte carta.
    # Lo que si es del hotel es el inventario, y de ahi sale que el mismo plato
    # tenga food cost distinto en dos hoteles.
    df_rec       = _xlsx("recetas.xlsx")

    # ── Y una particion por fuente ────────────────────────────────────────
    c_ap  = partir_por_hotel(df_ap,  ids)
    c_ar  = partir_por_hotel(df_ar,  ids)
    c_res = partir_por_hotel(df_res, ids)
    c_ven = partir_por_hotel(df_ven, ids)
    c_inv = partir_por_hotel(df_inv, ids)
    c_mer = partir_por_hotel(df_mer, ids)

    def _ficha(clave, nombre, con_drr=False):
        f = {
            "hotel_id": clave,
            "nombre":   nombre,
            "ap":       _ficha_ap(c_ap[clave]),
            "ar_ota":   _ficha_ar_ota(c_ar[clave]),
            "ar_real":  _ficha_ar_real(c_res[clave]),
            "fb":       _ficha_fb(df_rec, c_inv[clave], c_ven[clave], c_mer[clave]),
        }
        # La fila hotelera solo tiene sentido para un HOTEL. "Sin asignar" y
        # "desconocido" son cajas de documentos, no sitios con habitaciones:
        # preguntarles la ocupacion no significa nada.
        if con_drr:
            f["drr"] = _ficha_drr(clave)
            f["censo"] = ficha_censo.get(clave, {})
        return f

    fichas       = [_ficha(hid, nombres[hid], con_drr=True) for hid in ids]
    sin_asignar  = _ficha(CAJA_SIN_ASIGNAR, "Sin asignar")
    desconocido  = _ficha(CAJA_DESCONOCIDO, "Hotel desconocido")

    # El total, del df ENTERO y por los mismos calculadores. Que NO salga de
    # sumar las cajas es lo que hace que el cuadre demuestre algo: si el total
    # se calculara sumando, estariamos comprobando que 2+2 es 2+2.
    todas = fichas + [sin_asignar, desconocido]
    total = {
        "hotel_id": "grupo",
        "nombre":   "Grupo",
        "ap":       _ficha_ap(df_ap),
        "ar_ota":   _ficha_ar_ota(df_ar),
        "ar_real":  _ficha_ar_real(df_res),
        # La excepcion, y esta explicada en `_fb_del_grupo`: el ratio de F&B no
        # puede salir del inventario aplanado.
        "fb":       _fb_del_grupo(todas, _ficha_fb(df_rec, df_inv, df_ven, df_mer)),
    }

    cuadre, cuadra = _comprobar_cuadre(todas, total)

    return {
        "ok": True,
        "n_hoteles":   len(fichas),
        "hoteles":     fichas,
        "sin_asignar": sin_asignar,
        "desconocido": desconocido,
        "grupo":       total,
        # FASE E · la fila hotelera del grupo, ponderada y con su denominador.
        # Va aparte de `grupo` porque no es aditiva y no entra en el cuadre:
        # sumar porcentajes no significa nada.
        "hotelero":    _hotelero_del_grupo(fichas),
        # De grupo, y separado a proposito: el extracto es de la cuenta de la
        # sociedad, no del hotel. `movimientos_banco()` es la unica funcion del
        # almacen sin argumento `hotel` justamente por eso. Repartirlo entre
        # hoteles seria inventar.
        "banco":       _banco(),
        "cuadre":      cuadre,
        "cuadra":      cuadra,
    }


def _banco():
    from dashboard import stats_banco, _rdir
    df, info = _alm.movimientos_banco(reportes_dir=_rdir())
    if df is None or df.empty:
        return {"hay_datos": False}
    return dict(stats_banco(df), hay_datos=True, nivel="grupo",
                sin_conciliar=info.get("informe") is None)
