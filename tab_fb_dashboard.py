"""
tab_fb_dashboard.py — F&B Cost Control completo
Endpoints reales calculados desde recetas.xlsx, inventario.xlsx, mermas.xlsx, ventas_fb_diarias.xlsx
"""
from flask import Blueprint, jsonify, request
from pathlib import Path
import pandas as pd, json, time as _t
from datetime import datetime

fb_bp = Blueprint("fb", __name__, url_prefix="/fb")
BASE_DIR = Path(__file__).parent
from tenant_dirs import datos_dir as _t_ddir, tenant_id as _t_tid

class _TDatos:
    def __truediv__(self, other): return Path(_t_ddir()) / other
    def __str__(self): return _t_ddir()

DATOS = _TDatos()

# ── Cache ────────────────────────────────────────────────────────────────────
import os as _os_cache

_FB_CACHE: dict = {}
_FB_TTL = 180  # 3 min

def _huella(path):
    """(mtime, tamaño) del fichero, o None si no existe.

    Los dos y no solo el mtime: su resolucion puede ser de 1s y una escritura
    dentro del mismo segundo pasaria desapercibida.
    """
    try:
        st = _os_cache.stat(path)
        return (st.st_mtime, st.st_size)
    except OSError:
        return None


def _xlsx(fname, **kw):
    """Lee un Excel del tenant con cache que se invalida SOLA.

    La cache se descarta en cuanto el fichero cambia (mtime o tamaño), asi que
    ningun escritor tiene que acordarse de llamar a _invalidate(): antes,
    Procesar Archivos escribia ventas/inventario/mermas y el tab seguia
    mostrando los numeros viejos hasta 3 minutos. Poner la correccion aqui, en
    UN solo sitio, en vez de repartir llamadas por cada escritor, es el mismo
    criterio que en almacen_datos.
    El TTL se queda como red de seguridad (relojes raros, ficheros en red).
    """
    path = DATOS / fname
    key  = _t_tid() + "|" + fname
    now  = _t.time()
    huella = _huella(path)
    if key in _FB_CACHE:
        df, ts, huella_cache = _FB_CACHE[key]
        if huella == huella_cache and now - ts < _FB_TTL:
            return df
    df = pd.read_excel(path, **kw)
    _FB_CACHE[key] = (df, now, huella)
    return df

def _invalidate():
    _FB_CACHE.clear()

# ── Helpers ──────────────────────────────────────────────────────────────────
def _num_fb(v, defecto=0.0):
    """Numero tolerante. NaN, None, vacios y textos no numericos -> defecto.

    OJO con el patron `float(x or 0)`: **NaN es truthy**, asi que NO lo atrapa.
    Y un NaN que llegue a la respuesta la deja en JSON invalido: Flask lo
    serializa tal cual, el navegador falla al hacer r.json() y la pestaña se
    queda EN BLANCO con HTTP 200 (la regla del proyecto sobre NaN).
    """
    if v is None:
        return defecto
    try:
        f = float(v)
    except (TypeError, ValueError):
        s = str(v).strip().replace(",", ".")
        try:
            f = float(s)
        except (TypeError, ValueError):
            return defecto
    return defecto if f != f else f          # f != f es True solo para NaN


def _txt_ing(v):
    """Nombre de ingrediente comparable, tolerando nulos.

    MISMO criterio que hasta ahora —minusculas y sin espacios sobrantes—, solo
    que ahora no revienta con un vacio: una fila de inventario sin ingrediente
    tumbaba el endpoint con "'float' object has no attribute 'strip'".

    A PROPOSITO no quita los acentos, aunque `_clave_plato` si lo haga para los
    platos. Cambiarlo haria cruzar ingredientes que hoy no cruzan ("Cafe" vs
    "Café") y el food cost SUBIRIA: es una correccion, pero mueve un numero que
    el usuario mira, asi que va en su propio paso y con su propia comparacion.
    """
    s = "" if v is None else str(v)
    s = " ".join(s.split()).strip().lower()
    return "" if s in ("", "nan", "none", "nat", "<na>", "null") else s


def _calc_recipe_costs(df_rec, df_inv):
    """Coste teorico y FC% de cada receta, tolerando lo que venga.

    Se tolera AL LEER, mismo criterio que `_ventas_con_receta`: el recetario y
    el inventario los sube el cliente y vienen con huecos. Un hueco no puede
    tumbar la pestaña ni colar un NaN en la respuesta.

    Cada receta lleva `sin_precio` cuando no se ha podido leer su PVP: su FC%
    no es 0, es DESCONOCIDO, y quien agregue tiene que poder dejarla fuera en
    vez de contarla como la de mejor margen del menu (que es lo que pasaba).
    """
    cost_map = {}
    for fila in (df_inv.to_dict("records") if df_inv is not None and not df_inv.empty else []):
        nombre = _txt_ing(fila.get("ingrediente"))
        if not nombre:
            continue
        coste = _num_fb(fila.get("coste_unitario"), None)
        if coste is None:
            continue        # sin coste no aporta: que gane el de la propia receta
        cost_map[nombre] = coste

    results = []
    for rec in (df_rec.to_dict("records") if df_rec is not None and not df_rec.empty else []):
        ings = rec.get("ingredientes_json")
        if isinstance(ings, str):
            try:
                ings = json.loads(ings)
            except Exception:
                ings = []
        if not isinstance(ings, list):
            ings = []

        coste_total = 0.0
        for ing in ings:
            if not isinstance(ing, dict):
                continue
            unit = cost_map.get(_txt_ing(ing.get("ingrediente")))
            if unit is None:
                unit = _num_fb(ing.get("coste_unitario"), 0.0)
            coste_total += _num_fb(ing.get("cantidad"), 0.0) * unit

        precio = _num_fb(rec.get("precio_venta"), 0.0)
        fc_pct = round(coste_total / precio * 100, 2) if precio > 0 else 0.0
        results.append({
            "id":            _txt_ing(rec.get("id_receta")) and str(rec.get("id_receta")).strip() or "",
            "nombre":        str(rec.get("nombre") or "").strip() or "(sin nombre)",
            "categoria":     str(rec.get("categoria") or "").strip() or "General",
            "precio_venta":  round(precio, 2),
            "coste_teorico": round(coste_total, 3),
            "fc_pct":        fc_pct,
            "sin_precio":    precio <= 0,
            "alerta":        fc_pct > 35,
        })
    return results

# ── Ventas ↔ escandallo ────────────────────────────────────────────────────
# El TPV manda el NOMBRE del plato, no el id de la receta: el prompt de
# VENTAS_POS (lector_facturas_ap.py) no lo pide y _VEN_COL_MAP no lo genera.
# Antes, el calculo hacia sale['id_receta'] a pelo y un cliente que subiera sus
# ventas reales se cargaba el tab F&B entero con un KeyError.
#
# El cruce se resuelve AQUI, al leer, y NO al guardar: el recetario puede llegar
# despues que las ventas (o cambiar). Un id congelado en el fichero dejaria esas
# ventas huerfanas para siempre; resuelto al leer, el dia que entre el escandallo
# se encienden solas todas las ventas historicas.

import unicodedata as _ud


def _vacio(v):
    """True si el valor no dice nada: None, NaN, '' o los 'nan'/'NaT' de texto.

    Hace falta porque tras un concat de pandas conviven None, float('nan') y la
    cadena 'nan' (la que deja str(NaN)) en la misma columna.
    """
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    return str(v).strip().lower() in ("", "nan", "none", "nat", "<na>")


def _clave_plato(v):
    """Nombre de plato normalizado para cruzar ventas con escandallo.

    En Python plano y no con el accesor .str: en pandas 3 los nulos se propagan
    y todas las filas sin nombre acabarian compartiendo la MISMA clave, con lo
    que platos distintos cruzarian contra la misma receta.
    Sin acentos a proposito: "Sangria" y "Sangría" son el mismo plato.
    """
    s = "" if v is None else str(v)
    s = " ".join(s.split()).strip().lower()
    if s in ("", "nan", "none", "nat", "<na>"):
        return ""
    s = _ud.normalize("NFKD", s)
    return "".join(c for c in s if not _ud.combining(c))


def _ventas_con_receta(df_ven, recipes):
    """Devuelve (df, cobertura) con las columnas id_receta y nombre_plato SIEMPRE.

    - Si falta id_receta, se rellena cruzando el nombre del plato con el
      escandallo. Lo que no cruza se queda con id vacio: cuenta como ingreso,
      no cuenta para el coste, y no tumba nada.
    - cobertura mide EUROS, no filas: el food cost es un ratio de dinero, asi
      que lo que importa es que parte de la facturacion tiene escandallo.
    """
    df = df_ven.copy()

    # Completar por CELDA, no por columna. Las ventas pueden llegar por tres
    # puertas distintas y una de ellas (capa 1 del pipeline, la que decide por
    # el nombre del fichero) concatena el CSV EN CRUDO con lo que ya habia: el
    # resultado tiene 'plato' e 'importe' en unas filas y 'nombre_plato' y
    # 'total_venta' en otras, con NaN en los huecos. Rellenando celda a celda se
    # recuperan las dos mitades.
    for destino, alternativas in (
            ('nombre_plato',      ('plato', 'nombre', 'producto', 'item', 'descripcion', 'dish')),
            ('categoria',         ('tipo', 'category', 'grupo', 'familia')),
            ('unidades_vendidas', ('cantidad', 'qty', 'units', 'unidades')),
            ('total_venta',       ('total', 'importe', 'revenue', 'ventas')),
    ):
        valores = list(df[destino]) if destino in df.columns else [None] * len(df)
        for alt in alternativas:
            if alt not in df.columns:
                continue
            suplentes = list(df[alt])
            valores = [s if _vacio(v) else v for v, s in zip(valores, suplentes)]
        df[destino] = valores

    # Nada de NaN a partir de aqui: el JSON de Flask serializa NaN tal cual y el
    # resultado NO es JSON valido, asi que el navegador no puede leer la
    # respuesta y el tab se queda en blanco aunque el servidor devuelva 200.
    for col in ('unidades_vendidas', 'total_venta'):
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    df['nombre_plato'] = ['' if _vacio(v) else str(v).strip() for v in df['nombre_plato']]
    df['categoria'] = ['General' if _vacio(v) else str(v).strip() for v in df['categoria']]

    por_nombre = {}
    for r in recipes:
        k = _clave_plato(r.get('nombre'))
        if k and k not in por_nombre:      # primera receta gana: orden estable
            por_nombre[k] = r['id']

    ids_validos = {r['id'] for r in recipes}
    if 'id_receta' not in df.columns:
        df['id_receta'] = ''

    resueltos = []
    for fila in df.to_dict('records'):
        rid = fila.get('id_receta')
        rid = '' if rid is None else str(rid).strip()
        if rid in ('', 'nan', 'None', 'NaT'):
            rid = ''
        # un id que no existe en el escandallo no vale; probamos por nombre
        if rid not in ids_validos:
            rid = por_nombre.get(_clave_plato(fila.get('nombre_plato')), '')
        resueltos.append(rid)
    df['id_receta'] = resueltos

    def _eur(v):
        try:
            return 0.0 if v is None or pd.isna(v) else float(v)
        except (TypeError, ValueError):
            return 0.0

    ventas_cruzan, ventas_sueltas, sin_receta = 0.0, 0.0, []
    for fila in df.to_dict('records'):
        imp = _eur(fila.get('total_venta'))
        if fila.get('id_receta'):
            ventas_cruzan += imp
        else:
            ventas_sueltas += imp
            n = fila.get('nombre_plato')
            n = '' if _vacio(n) else str(n).strip()
            if n and n not in sin_receta:
                sin_receta.append(n)

    total = ventas_cruzan + ventas_sueltas
    cobertura = {
        'pct': round(ventas_cruzan / total * 100, 1) if total > 0 else 0.0,
        'ventas_con_receta': round(ventas_cruzan, 2),
        'ventas_sin_receta': round(ventas_sueltas, 2),
        'platos_sin_receta': sorted(sin_receta)[:10],
        'n_platos_sin_receta': len(sin_receta),
    }
    return df, cobertura


# ── Resultados consolidados ────────────────────────────────────────────────
@fb_bp.route("/api/resultados")
def api_resultados():
    try:
        # Early check: si no hay datos, devolver respuesta vacía
        import os as _os
        ventas_path = _os.path.join(_t_ddir(), "ventas_fb_diarias.xlsx")
        _df_check = pd.read_excel(ventas_path)
        if _df_check.empty or len(_df_check) < 1:
            return jsonify({
                "data": [], "chart": {"data": [], "labels": []},
                "total": 0, "fc_teorico": 0, "fc_real": 0, "mermas": 0,
                "correctas": 0, "discrepancias": 0, "di_pendientes": 0,
                "importe_reclamable": 0, "importe_total": 0,
                "pendientes_firma": {}, "rechazadas": 0, "sin_accion": 0,
                "meta": {}
            })
        df_rec = _xlsx("recetas.xlsx")
        df_inv = _xlsx("inventario.xlsx")
        df_mer = _xlsx("mermas.xlsx")
        df_ven = _xlsx("ventas_fb_diarias.xlsx")

        recipes = _calc_recipe_costs(df_rec, df_inv)
        recipe_map = {r['id']: r for r in recipes}

        # id_receta garantizado: sin esto el tab entero se caia con KeyError
        # cuando las ventas venian del TPV (ver _ventas_con_receta).
        df_ven, cobertura = _ventas_con_receta(df_ven, recipes)

        # Ventas: total facturado y coste de lo que SI tiene escandallo
        total_ventas   = float(pd.to_numeric(df_ven['total_venta'], errors='coerce').fillna(0).sum())
        coste_real_sum = 0.0
        for sale in df_ven.to_dict('records'):
            rec = recipe_map.get(sale.get('id_receta'))
            if not rec:
                continue
            try:
                uds = float(sale.get('unidades_vendidas') or 0)
            except (TypeError, ValueError):
                uds = 0.0
            coste_real_sum += rec['coste_teorico'] * uds

        # El food cost se calcula SOLO sobre las ventas que cruzan con receta.
        # Dividir entre TODAS las ventas daria un numero mas bajo y tranquilizador
        # justamente cuando falta escandallo: parecerian margenes buenos cuando en
        # realidad no se sabe. La cobertura va al lado para que se vea sobre que
        # parte de la facturacion esta calculado.
        base_fc = cobertura['ventas_con_receta']
        # fc_teorico y fc_real salen identicos por construccion (misma suma
        # calculada de dos formas); esta en el cajon de pendientes.
        fc_teorico_global = round(coste_real_sum / base_fc * 100, 2) if base_fc > 0 else 0
        fc_real_global    = fc_teorico_global

        # Mermas — normalizar columnas
        if 'coste_merma' not in df_mer.columns and 'coste' in df_mer.columns:
            df_mer = df_mer.rename(columns={'coste': 'coste_merma'})
        if 'coste_merma' not in df_mer.columns:
            df_mer['coste_merma'] = 0
        df_mer['coste_merma'] = pd.to_numeric(df_mer['coste_merma'], errors='coerce').fillna(0)
        coste_mermas = float(df_mer['coste_merma'].sum())

        # Categorías
        cats_summary = {}
        for _, sale in df_ven.iterrows():
            cat = str(sale.get('categoria', 'General'))
            rid = sale.get('id_receta', '')
            uds = float(sale.get('unidades_vendidas', 0) or 0)
            ven = float(sale.get('total_venta', 0) or 0)
            rec = recipe_map.get(rid, {})
            c   = rec.get('coste_teorico', 0) * uds
            if cat not in cats_summary:
                cats_summary[cat] = {'ventas': 0, 'coste': 0}
            cats_summary[cat]['ventas'] += ven
            cats_summary[cat]['coste']  += c

        categorias = []
        for cat, vals in sorted(cats_summary.items(), key=lambda x: -x[1]['ventas']):
            fc_t = round(vals['coste'] / vals['ventas'] * 100, 1) if vals['ventas'] > 0 else 0
            categorias.append({
                'nombre': cat,
                'total_ventas': round(vals['ventas'], 0),
                'fc_teorico_pct': fc_t,
                'fc_real_pct': fc_t,  # same for category level
                'alerta': fc_t > 35,
            })

        # Ranking top platos por ventas.
        # _ventas_con_receta ya garantiza id_receta y nombre_plato, asi que aqui
        # no hay que adivinar la columna. IMPORTANTE: si alguna de las dos
        # faltase, pandas 3 interpretaria la lista ['id_receta','nombre_plato']
        # como un ARRAY de etiquetas cuando el df tiene exactamente 2 filas, y
        # devolveria un ranking inventado SIN dar error. Por eso se garantizan
        # antes en vez de resolverlas aqui.
        ranking = (df_ven.groupby(['id_receta', 'nombre_plato'])['total_venta']
                   .sum().reset_index()
                   .sort_values('total_venta', ascending=False).head(8))
        ranking_top = []
        for _, row in ranking.iterrows():
            rec = recipe_map.get(row['id_receta'], {})
            ranking_top.append({
                'nombre': row['nombre_plato'],
                'total_ventas': round(float(row['total_venta']), 0),
                'fc_real_pct': rec.get('fc_pct', 0),
                'fc_teorico_pct': rec.get('fc_pct', 0),
            })

        # Ventas diarias por categoría (últimos 30 días)
        df_ven['fecha'] = pd.to_datetime(df_ven['fecha'])
        ventas_diarias = (df_ven.groupby(df_ven['fecha'].dt.strftime('%Y-%m-%d'))['total_venta']
                          .sum().reset_index().tail(30))

        return jsonify({
            'ok': True,
            'resumen': {
                'total_ventas':    round(total_ventas, 2),
                'fc_teorico_pct':  fc_teorico_global,
                'fc_real_pct':     fc_real_global,
                'coste_mermas':    round(coste_mermas, 2),
                'alerta':          fc_real_global > fc_teorico_global + 3,
                'cobertura':       cobertura,
            },
            'cobertura': cobertura,
            'categorias': categorias,
            'ranking_top': ranking_top,
            'ventas_diarias': {
                'fechas':  ventas_diarias['fecha'].tolist(),
                'totales': [round(v, 0) for v in ventas_diarias['total_venta'].tolist()],
            },
        })
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@fb_bp.route("/api/inventario")
def api_inventario():
    try:
        import os as _os2
        inv_path = _os2.path.join(_t_ddir(), "inventario.xlsx")
        _df_inv_check = pd.read_excel(inv_path)
        if _df_inv_check.empty or len(_df_inv_check) < 1:
            return jsonify({'ok': True, 'items': [], 'valor_total': 0, 'alertas_count': 0, 'criticos_count': 0, 'top_alerts': []})
        df = _xlsx("inventario.xlsx")
        df['stock_actual_kg_l'] = pd.to_numeric(df['stock_actual_kg_l'], errors='coerce').fillna(0)
        df['stock_inicial_kg_l'] = pd.to_numeric(df['stock_inicial_kg_l'], errors='coerce').fillna(0)
        df['coste_unitario'] = pd.to_numeric(df['coste_unitario'], errors='coerce').fillna(0)
        items = []
        for _, row in df.iterrows():
            stock_act = float(row.get('stock_actual_kg_l', 0) or 0)
            stock_ini = float(row.get('stock_inicial_kg_l', 0) or 0)
            pct = round(stock_act / stock_ini * 100, 0) if stock_ini > 0 else 0
            items.append({
                'ingrediente': str(row.get('ingrediente', '—')),
                'categoria': str(row.get('categoria', 'General')),
                'stock_inicial': stock_ini,
                'stock_actual': stock_act,
                'unidad': str(row.get('unidad', 'kg')),
                'coste_unitario': float(row.get('coste_unitario', 0) or 0),
                'proveedor': str(row.get('proveedor', '—')),
                'pct_restante': pct,
                'alerta': pct < 30,
                'critico': pct < 15,
            })
        alertas = [i for i in items if i['alerta']]
        criticos = [i for i in items if i['critico']]
        return jsonify({'ok': True, 'items': items,
                        'valor_total': round(sum(i['stock_actual']*i['coste_unitario'] for i in items), 2),
                        'alertas_count': len(alertas),
                        'criticos_count': len(criticos),
                        'top_alerts': [i['ingrediente'] for i in criticos[:3]]})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@fb_bp.route("/api/mermas")
def api_mermas():
    try:
        import os as _os3
        mer_path = _os3.path.join(_t_ddir(), "mermas.xlsx")
        _df_mer_check = pd.read_excel(mer_path)
        if _df_mer_check.empty or len(_df_mer_check) < 1:
            return jsonify({'ok': True, 'mermas': [], 'total_coste': 0, 'por_categoria': {}, 'total': 0, 'por_causa': {}})
        df = _xlsx("mermas.xlsx")
        # Normalizar: Claude puede devolver 'coste' en vez de 'coste_merma'
        if 'coste_merma' not in df.columns and 'coste' in df.columns:
            df = df.rename(columns={'coste': 'coste_merma'})
        if 'cantidad_merma' not in df.columns and 'cantidad' in df.columns:
            df = df.rename(columns={'cantidad': 'cantidad_merma'})
        if 'coste_merma' not in df.columns:
            df['coste_merma'] = 0
        if 'cantidad_merma' not in df.columns:
            df['cantidad_merma'] = 0
        df['coste_merma'] = pd.to_numeric(df['coste_merma'], errors='coerce').fillna(0)
        total_coste  = float(df['coste_merma'].sum())
        por_categoria = df.groupby('categoria')['coste_merma'].sum().sort_values(ascending=False).to_dict() if 'categoria' in df.columns else {}
        mermas = []
        for _, row in df.iterrows():
            mermas.append({
                'fecha': str(row.get('fecha', ''))[:10],
                'ingrediente': str(row.get('ingrediente', '—')),
                'categoria': str(row.get('categoria', 'General')),
                'cantidad': float(row.get('cantidad_merma', 0) or 0),
                'unidad': str(row.get('unidad', 'kg')),
                'causa': str(row.get('causa', '—')),
                'coste': float(row.get('coste_merma', 0) or 0),
            })
        total = round(sum(m['coste'] for m in mermas), 2)
        por_causa = {}
        for m in mermas:
            por_causa[m['causa']] = round(por_causa.get(m['causa'], 0) + m['coste'], 2)
        return jsonify({'ok': True, 'mermas': mermas, 'total_coste': total_coste, 'por_categoria': por_categoria, 'total': total, 'por_causa': por_causa})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@fb_bp.route("/api/registrar_merma", methods=["POST"])
def api_registrar_merma():
    """Registra una nueva merma en mermas.xlsx."""
    data = request.get_json(silent=True) or {}
    required = ['ingrediente', 'cantidad', 'unidad', 'causa', 'coste_unitario']
    for f in required:
        if not data.get(f): return jsonify({'ok': False, 'error': f'Falta campo: {f}'}), 400
    try:
        path = DATOS / "mermas.xlsx"
        df = pd.read_excel(path)
        cantidad = float(data['cantidad'])
        coste_u  = float(data['coste_unitario'])
        new_row = {
            'fecha': datetime.now().strftime('%Y-%m-%d'),
            'ingrediente': data['ingrediente'],
            'categoria': data.get('categoria', '—'),
            'cantidad_merma': cantidad,
            'unidad': data['unidad'],
            'causa': data['causa'],
            'coste_unitario': coste_u,
            'coste_merma': round(cantidad * coste_u, 2),
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel(path, index=False)
        _invalidate()
        return jsonify({'ok': True, 'coste': new_row['coste_merma']})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@fb_bp.route("/api/recetas")
def api_recetas():
    try:
        df_rec = _xlsx("recetas.xlsx")
        df_inv = _xlsx("inventario.xlsx")
        recipes = _calc_recipe_costs(df_rec, df_inv)
        # Add ranking and margin info
        for i, r in enumerate(sorted(recipes, key=lambda x: x.get('fc_pct', 0))):
            r['rank'] = i + 1
        # Una receta sin PVP tiene un FC% DESCONOCIDO, no un 0%. Contarla como 0
        # la convertia en "la de mejor margen del menu", que es justo lo
        # contrario de lo que pasa: no se sabe. Se dejan fuera de las medias y
        # del ranking, y se dice cuantas son.
        con_precio = [r for r in recipes if not r.get('sin_precio')]
        avg_fc = round(sum(r.get('fc_pct', 0) for r in con_precio) / len(con_precio), 1) \
            if con_precio else 0
        return jsonify({'ok': True, 'recetas': recipes, 'avg_fc_pct': avg_fc,
                       'sin_precio': len(recipes) - len(con_precio),
                       'best_margin': min(con_precio, key=lambda x: x.get('fc_pct', 100)).get('nombre', '') if con_precio else '',
                       'worst_margin': max(con_precio, key=lambda x: x.get('fc_pct', 0)).get('nombre', '') if con_precio else ''})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@fb_bp.route("/api/ejecutar")
def api_ejecutar():
    """Recalcula los datos (limpia caché)."""
    from flask import Response
    def gen():
        _invalidate()
        yield "data: Recalculando F&B...\n\n"
        yield "data: Leyendo ventas...\n\n"
        yield "data: Calculando Food Cost...\n\n"
        yield "data: FB_COMPLETO\n\n"
    return Response(gen(), mimetype='text/event-stream')
