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
        # Early check: si no hay datos, devolver respuesta vacía
        import os as _os
        ventas_path = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "datos-referencia", "ventas_fb_diarias.xlsx")
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

        # Ranking top platos por ventas
        # Defensivo: usar columna de nombre que exista
        nombre_col = 'nombre_plato' if 'nombre_plato' in df_ven.columns else ('plato' if 'plato' in df_ven.columns else 'id_receta')
        ranking = (df_ven.groupby(['id_receta', nombre_col])['total_venta']
                   .sum().reset_index()
                   .sort_values('total_venta', ascending=False).head(8))
        ranking = ranking.rename(columns={nombre_col: 'nombre_plato'})
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
        import os as _os2
        inv_path = _os2.path.join(_os2.path.dirname(_os2.path.abspath(__file__)), "datos-referencia", "inventario.xlsx")
        _df_inv_check = pd.read_excel(inv_path)
        if _df_inv_check.empty or len(_df_inv_check) < 1:
            return jsonify({"items": [], "total_valor": 0, "categorias": {}})
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
        mer_path = _os3.path.join(_os3.path.dirname(_os3.path.abspath(__file__)), "datos-referencia", "mermas.xlsx")
        _df_mer_check = pd.read_excel(mer_path)
        if _df_mer_check.empty or len(_df_mer_check) < 1:
            return jsonify({"mermas": [], "total": 0})
        df = _xlsx("mermas.xlsx")
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
        avg_fc = round(sum(r.get('fc_pct',0) for r in recipes) / len(recipes), 1) if recipes else 0
        return jsonify({'ok': True, 'recetas': recipes, 'avg_fc_pct': avg_fc,
                       'best_margin': min(recipes, key=lambda x: x.get('fc_pct',100)).get('nombre','') if recipes else '',
                       'worst_margin': max(recipes, key=lambda x: x.get('fc_pct',0)).get('nombre','') if recipes else ''})
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
