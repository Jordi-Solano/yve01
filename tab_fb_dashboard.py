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
DATOS    = BASE_DIR / "datos-referencia"

# ── Cache ────────────────────────────────────────────────────────────────────
_FB_CACHE: dict = {}
_FB_TTL = 180  # 3 min

def _xlsx(fname, **kw):
    path = DATOS / fname
    key  = fname
    now  = _t.time()
    if key in _FB_CACHE:
        df, ts = _FB_CACHE[key]
        if now - ts < _FB_TTL: return df
    df = pd.read_excel(path, **kw)
    _FB_CACHE[key] = (df, now)
    return df

def _invalidate():
    _FB_CACHE.clear()

# ── Helpers ──────────────────────────────────────────────────────────────────
def _calc_recipe_costs(df_rec, df_inv):
    """Calcula coste teórico y FC% de cada receta."""
    # Build unit cost map from inventory
    cost_map = {}
    for _, row in df_inv.iterrows():
        cost_map[row['ingrediente'].strip().lower()] = float(row['coste_unitario'])

    results = []
    for _, rec in df_rec.iterrows():
        try:
            ings = json.loads(rec['ingredientes_json']) if isinstance(rec['ingredientes_json'], str) else []
        except Exception:
            ings = []
        coste_total = sum(
            float(ing.get('cantidad', 0)) * cost_map.get(ing.get('ingrediente','').strip().lower(),
                float(ing.get('coste_unitario', 0)))
            for ing in ings
        )
        precio = float(rec['precio_venta'])
        fc_pct = round(coste_total / precio * 100, 2) if precio > 0 else 0
        results.append({
            'id': rec['id_receta'],
            'nombre': rec['nombre'],
            'categoria': rec['categoria'],
            'precio_venta': precio,
            'coste_teorico': round(coste_total, 3),
            'fc_pct': fc_pct,
            'alerta': fc_pct > 35,
        })
    return results

# ── Resultados consolidados ────────────────────────────────────────────────
@fb_bp.route("/api/resultados")
def api_resultados():
    try:
        df_rec = _xlsx("recetas.xlsx")
        df_inv = _xlsx("inventario.xlsx")
        df_mer = _xlsx("mermas.xlsx")
        df_ven = _xlsx("ventas_fb_diarias.xlsx")

        recipes = _calc_recipe_costs(df_rec, df_inv)
        recipe_map = {r['id']: r for r in recipes}

        # Ventas: total y coste real
        total_ventas   = float(df_ven['total_venta'].sum())
        coste_real_sum = 0.0
        for _, sale in df_ven.iterrows():
            rid  = sale['id_receta']
            uds  = float(sale['unidades_vendidas'])
            rec  = recipe_map.get(rid)
            if rec: coste_real_sum += rec['coste_teorico'] * uds

        fc_teorico_global = round(
            sum(r['coste_teorico'] * float(df_ven[df_ven['id_receta']==r['id']]['unidades_vendidas'].sum())
                for r in recipes) / total_ventas * 100, 2
        ) if total_ventas > 0 else 0

        fc_real_global = round(coste_real_sum / total_ventas * 100, 2) if total_ventas > 0 else 0

        # Mermas
        df_mer['coste_merma'] = pd.to_numeric(df_mer['coste_merma'], errors='coerce').fillna(0)
        coste_mermas = float(df_mer['coste_merma'].sum())

        # Categorías
        cats_summary = {}
        for _, sale in df_ven.iterrows():
            cat = str(sale['categoria'])
            rid = sale['id_receta']
            uds = float(sale['unidades_vendidas'])
            ven = float(sale['total_venta'])
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

        # Ranking top platos por ventas
        ranking = (df_ven.groupby(['id_receta','nombre_plato'])['total_venta']
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
            },
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
        df = _xlsx("inventario.xlsx")
        df['stock_actual_kg_l'] = pd.to_numeric(df['stock_actual_kg_l'], errors='coerce').fillna(0)
        df['stock_inicial_kg_l'] = pd.to_numeric(df['stock_inicial_kg_l'], errors='coerce').fillna(0)
        df['coste_unitario'] = pd.to_numeric(df['coste_unitario'], errors='coerce').fillna(0)
        items = []
        for _, row in df.iterrows():
            pct = round(row['stock_actual_kg_l'] / row['stock_inicial_kg_l'] * 100, 0) if row['stock_inicial_kg_l'] > 0 else 0
            items.append({
                'ingrediente': str(row['ingrediente']),
                'categoria': str(row['categoria']),
                'stock_inicial': float(row['stock_inicial_kg_l']),
                'stock_actual': float(row['stock_actual_kg_l']),
                'unidad': str(row['unidad']),
                'coste_unitario': float(row['coste_unitario']),
                'proveedor': str(row['proveedor']),
                'pct_restante': pct,
                'alerta': pct < 30,
                'critico': pct < 15,
            })
        return jsonify({'ok': True, 'items': items,
                        'valor_total': round(sum(i['stock_actual']*i['coste_unitario'] for i in items), 2)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@fb_bp.route("/api/mermas")
def api_mermas():
    try:
        df = _xlsx("mermas.xlsx")
        df['coste_merma'] = pd.to_numeric(df['coste_merma'], errors='coerce').fillna(0)
        total_coste  = float(df['coste_merma'].sum())
        por_categoria = df.groupby('categoria')['coste_merma'].sum().sort_values(ascending=False).to_dict() if 'categoria' in df.columns else {}
        mermas = []
        for _, row in df.iterrows():
            mermas.append({
                'fecha': str(row['fecha'])[:10],
                'ingrediente': str(row['ingrediente']),
                'categoria': str(row['categoria']),
                'cantidad': float(row['cantidad_merma']),
                'unidad': str(row['unidad']),
                'causa': str(row['causa']),
                'coste': float(row['coste_merma']),
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
        return jsonify({'ok': True, 'recetas': recipes})
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
