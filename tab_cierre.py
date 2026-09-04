"""
tab_cierre.py — Yve.01 · Cierre de mes (OLA B)
Endpoints del apartado "Cierre": asientos del mes, reconciliacion de cuentas
y su Excel. Solo lectura: nada de aqui escribe en los datos.
"""
import os
from flask import Blueprint, jsonify, request, send_file

from tenant_dirs import datos_dir as _t_ddir, procesadas_dir as _t_pdir, reportes_dir as _t_rdir

cierre_bp = Blueprint('cierre', __name__)


def _args():
    mes = (request.args.get('mes') or '').strip()[:7] or None
    try:
        import censo_hoteles
        hotel = censo_hoteles.activo() or None
    except Exception:
        hotel = None
    return mes, hotel


def _dirs():
    return {"procesadas_dir": str(_t_pdir()), "reportes_dir": str(_t_rdir()), "datos_dir": str(_t_ddir())}


@cierre_bp.route('/api/cierre/asientos')
def api_cierre_asientos():
    import cierre_mes as CM
    mes, hotel = _args()
    try:
        res, rec = CM.cierre_completo(mes, hotel, **_dirs())
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]}), 500
    limite = int(request.args.get('limite') or 400)
    return jsonify({'ok': True, 'hotel': hotel or '', **{k: v for k, v in res.items() if k != 'asientos'},
                    'asientos': res['asientos'][:limite], 'truncado': len(res['asientos']) > limite,
                    'mayor': rec['mayor'], 'reconciliacion': {k: v for k, v in rec.items() if k != 'mayor'}})


@cierre_bp.route('/api/exportar/cierre')
def api_exportar_cierre():
    import cierre_mes as CM
    mes, hotel = _args()
    res, rec = CM.cierre_completo(mes, hotel, **_dirs())
    buf, nombre = CM.exportar_excel(res, rec)
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ── Cuadre de banco por pestañas (Ola B · bloque 2) ─────────────────────────
def _cuadre(mes, hotel):
    import cuadre_banco as CB
    import almacen_datos as ALM
    import pandas as pd
    dd = str(_t_ddir())
    try:
        bk, _ = ALM.movimientos_banco(datos_dir=dd, reportes_dir=str(_t_rdir()))
        if hotel:
            bk = ALM._filtrar_hotel(bk, hotel)
    except Exception:
        bk = pd.DataFrame()
    try:
        vf = pd.read_excel(os.path.join(dd, 'ventas_fb_diarias.xlsx'))
        if hotel:
            vf = ALM._filtrar_hotel(vf, hotel)
    except Exception:
        vf = pd.DataFrame()
    return CB.cuadrar(mes, bk, vf, CB.palabras(dd), CB.manuales(dd), CB.proveedores_conocidos(dd))


@cierre_bp.route('/api/cuadre_banco')
def api_cuadre_banco():
    mes, hotel = _args()
    try:
        res = _cuadre(mes, hotel)
    except Exception as e:
        return jsonify({'ok_api': False, 'error': str(e)[:200]}), 500
    return jsonify({'ok_api': True, 'hotel': hotel or '', **res})


@cierre_bp.route('/api/cuadre_banco/asignar', methods=['POST'])
def api_cuadre_banco_asignar():
    """Pone un movimiento en una pestaña a mano (clave = clave_movimiento)."""
    import cuadre_banco as CB
    data = request.get_json(force=True, silent=True) or {}
    clave = str(data.get('clave') or '').strip()
    pestana = str(data.get('pestana') or '').strip().upper()
    if not clave:
        return jsonify({'ok': False, 'error': 'falta clave'}), 400
    if pestana and pestana not in CB.PESTANAS:
        return jsonify({'ok': False, 'error': 'pestaña desconocida'}), 400
    CB.guardar_manual(clave, pestana if pestana != 'SIN_CLASIFICAR' else '', str(_t_ddir()))
    try:
        from dashboard import _audit
        from flask_login import current_user
        _audit('CUADRE_BANCO_ASIGNAR', f'{clave} -> {pestana or "(quitar)"}', getattr(current_user, 'username', None) or 'sistema')
    except Exception:
        pass
    return jsonify({'ok': True})


@cierre_bp.route('/api/exportar/cuadre_banco')
def api_exportar_cuadre_banco():
    import cuadre_banco as CB
    mes, hotel = _args()
    buf, nombre = CB.exportar_excel(_cuadre(mes, hotel))
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ── Inventarios de cierre (Ola B · bloque 3) ────────────────────────────────
def _inventario_hotel(hotel):
    import pandas as pd
    import almacen_datos as ALM
    dd = str(_t_ddir())
    try:
        df = pd.read_excel(os.path.join(dd, 'inventario.xlsx'))
    except Exception:
        return pd.DataFrame()
    return ALM._filtrar_hotel(df, hotel) if hotel else df


def _coste_teorico_fb(mes, hotel):
    """Escandallo × unidades vendidas del mes (tab_fb_dashboard.resumen_fb)."""
    try:
        import pandas as pd
        import tab_fb_dashboard as FB
        from provisiones import _fecha, _mes_a_rango
        ini, fin, _ = _mes_a_rango(mes)
        df_rec = FB._xlsx('recetas.xlsx')
        df_inv = FB._xlsx_hotel('inventario.xlsx') if hotel else FB._xlsx('inventario.xlsx')
        df_ven = FB._xlsx_hotel('ventas_fb_diarias.xlsx') if hotel else FB._xlsx('ventas_fb_diarias.xlsx')
        df_mer = FB._xlsx_hotel('mermas.xlsx') if hotel else FB._xlsx('mermas.xlsx')
        if df_ven is None or df_ven.empty or 'fecha' not in df_ven.columns:
            return None
        mask = df_ven['fecha'].map(lambda v: (lambda f: f is not None and ini <= f <= fin)(_fecha(v)))
        df_ven = df_ven[mask]
        if df_ven.empty:
            return None
        _, _, resumen = FB.resumen_fb(df_rec, df_inv, df_ven, df_mer)
        return resumen.get('coste_escandallo')
    except Exception:
        return None


def _inventarios(mes, hotel):
    import inventarios as INV
    import almacen_datos as ALM
    dd = str(_t_ddir())
    try:
        ap = ALM.facturas_ap(str(_t_pdir()), str(_t_rdir()), hotel=hotel or None)
    except Exception:
        ap = None
    return INV.valorar(mes, _inventario_hotel(hotel), ap, _coste_teorico_fb(mes, hotel), INV.config_familias(dd))


@cierre_bp.route('/api/inventarios')
def api_inventarios():
    mes, hotel = _args()
    try:
        res = _inventarios(mes, hotel)
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]}), 500
    return jsonify({'ok': True, 'hotel': hotel or '', **res})


@cierre_bp.route('/api/exportar/inventarios')
def api_exportar_inventarios():
    import inventarios as INV
    mes, hotel = _args()
    buf, nombre = INV.exportar_excel(_inventarios(mes, hotel))
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@cierre_bp.route('/api/inventarios/hoja_recuento')
def api_hoja_recuento():
    import inventarios as INV
    mes, hotel = _args()
    from provisiones import _mes_a_rango
    _, _, mes = _mes_a_rango(mes)
    buf, nombre = INV.hoja_recuento(_inventario_hotel(hotel), mes, INV.config_familias(str(_t_ddir())))
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@cierre_bp.route('/api/inventarios/recuento', methods=['POST'])
def api_subir_recuento():
    """Sube la hoja de recuento rellenada: el `recuento` pasa a ser el stock final.

    Se conserva TODO lo demas del articulo (coste, categoria, stock inicial):
    solo cambia el stock final de las filas contadas. Los articulos nuevos
    entran con lo que traiga la hoja.
    """
    import pandas as pd
    import inventarios as INV
    f = request.files.get('archivo') or request.files.get('file')
    if not f:
        return jsonify({'ok': False, 'error': 'Falta el fichero (campo "archivo")'}), 400
    try:
        df_c, n, saltadas = INV.leer_recuento(f)
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]}), 400
    if df_c.empty:
        return jsonify({'ok': False, 'error': 'La hoja no trae ningun recuento informado', 'saltadas': saltadas}), 400
    mes, hotel = _args()
    actual = _inventario_hotel(hotel)
    por_nombre = {}
    if actual is not None and not actual.empty and 'ingrediente' in actual.columns:
        for _, r in actual.iterrows():
            por_nombre[INV._norm(r.get('ingrediente'))] = r.to_dict()
    filas = []; nuevos = 0
    for _, r in df_c.iterrows():
        base = por_nombre.get(INV._norm(r['ingrediente']))
        if base is None:
            nuevos += 1
            base = {'ingrediente': r['ingrediente']}
        fila = dict(base)
        fila['stock_actual_kg_l'] = float(r['stock_actual_kg_l'])
        for k in ('categoria', 'unidad', 'coste_unitario', 'proveedor'):
            if k in r and pd.notna(r[k]) and str(r[k]).strip():
                fila[k] = r[k]
        fila.pop('hotel_id', None)
        filas.append(fila)
    from dashboard import _guardar_fb_del_hotel, _audit
    _guardar_fb_del_hotel(pd.DataFrame(filas), 'inventario.xlsx')
    try:
        from flask_login import current_user
        _audit('INVENTARIO_RECUENTO', f'{n} articulos contados, {nuevos} nuevos', getattr(current_user, 'username', None) or 'sistema')
    except Exception:
        pass
    return jsonify({'ok': True, 'contados': n, 'nuevos': nuevos, 'saltadas': saltadas})


# ── Inmovilizado y amortizaciones (Ola B · bloque 4) ────────────────────────
def _activos_df():
    import pandas as pd
    import inmovilizado as IM
    ruta = os.path.join(str(_t_ddir()), IM.FICHERO)
    if not os.path.exists(ruta):
        return pd.DataFrame(columns=IM.COLUMNAS)
    try:
        return pd.read_excel(ruta)
    except Exception:
        return pd.DataFrame(columns=IM.COLUMNAS)


def _guardar_activos(df):
    import inmovilizado as IM
    ruta = os.path.join(str(_t_ddir()), IM.FICHERO)
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    tmp = ruta + '.tmp.xlsx'
    df.to_excel(tmp, index=False)
    os.replace(tmp, ruta)


def _inmovilizado(mes, hotel):
    import inmovilizado as IM
    import almacen_datos as ALM
    df = _activos_df()
    if hotel and not df.empty:
        df = ALM._filtrar_hotel(df, hotel)
    try:
        ap = ALM.facturas_ap(str(_t_pdir()), str(_t_rdir()), hotel=hotel or None)
    except Exception:
        ap = None
    return IM.amortizar_mes(mes, df, ap, IM.config(str(_t_ddir())))


@cierre_bp.route('/api/inmovilizado')
def api_inmovilizado():
    mes, hotel = _args()
    try:
        res = _inmovilizado(mes, hotel)
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]}), 500
    import inmovilizado as IM
    return jsonify({'ok': True, 'hotel': hotel or '', 'categorias': {k: {'vida': v[0], 'cuenta': v[1], 'nombre': v[4]} for k, v in IM.CATEGORIAS.items()}, **res})


@cierre_bp.route('/api/inmovilizado/alta', methods=['POST'])
def api_inmovilizado_alta():
    """Da de alta un activo (o lo actualiza si trae id)."""
    import pandas as pd
    import inmovilizado as IM
    data = request.get_json(force=True, silent=True) or {}
    try:
        import censo_hoteles
        data.setdefault('hotel_id', censo_hoteles.para_guardar())
        act = IM.normalizar_activo(data, IM.config(str(_t_ddir())))
    except ValueError as e:
        return jsonify({'ok': False, 'error': str(e)}), 400
    df = _activos_df()
    if not act['id']:
        from datetime import datetime
        act['id'] = 'INM-' + datetime.now().strftime('%Y%m%d%H%M%S%f')[:17]
    else:
        df = df[df['id'].astype(str) != act['id']] if 'id' in df.columns else df
    df = pd.concat([df, pd.DataFrame([act])], ignore_index=True)
    _guardar_activos(df)
    try:
        from dashboard import _audit
        from flask_login import current_user
        _audit('INMOVILIZADO_ALTA', f"{act['id']} {act['descripcion']} {act['coste']} EUR", getattr(current_user, 'username', None) or 'sistema')
    except Exception:
        pass
    return jsonify({'ok': True, 'activo': act})


@cierre_bp.route('/api/inmovilizado/baja', methods=['POST'])
def api_inmovilizado_baja():
    """Fecha de baja de un activo (deja de amortizarse desde ese mes)."""
    import inmovilizado as IM
    from provisiones import _fecha
    data = request.get_json(force=True, silent=True) or {}
    aid = str(data.get('id') or '').strip()
    fb = _fecha(data.get('fecha_baja'))
    if not aid or fb is None:
        return jsonify({'ok': False, 'error': 'Faltan id o fecha_baja'}), 400
    df = _activos_df()
    if df.empty or 'id' not in df.columns or not (df['id'].astype(str) == aid).any():
        return jsonify({'ok': False, 'error': 'Activo no encontrado'}), 404
    df['fecha_baja'] = df['fecha_baja'].astype(object) if 'fecha_baja' in df.columns else ''
    df.loc[df['id'].astype(str) == aid, 'fecha_baja'] = fb.isoformat()
    _guardar_activos(df)
    try:
        from dashboard import _audit
        from flask_login import current_user
        _audit('INMOVILIZADO_BAJA', f'{aid} baja {fb.isoformat()}', getattr(current_user, 'username', None) or 'sistema')
    except Exception:
        pass
    return jsonify({'ok': True})


@cierre_bp.route('/api/exportar/inmovilizado')
def api_exportar_inmovilizado():
    import inmovilizado as IM
    mes, hotel = _args()
    buf, nombre = IM.exportar_excel(_inmovilizado(mes, hotel))
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ── Fiscal: 303, 349, SII (Ola B · bloque 6) ────────────────────────────────
@cierre_bp.route('/api/fiscal')
def api_fiscal():
    import fiscal as FI
    mes, hotel = _args()
    try:
        res = FI.fiscal_completo(mes, hotel, **_dirs())
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]}), 500
    # cuadre con el libro del mes (477/472), protegido
    try:
        import cierre_mes as CM
        asi, _ = CM.cierre_completo(mes, hotel, **_dirs())
        rep = round(sum(a['haber'] - a['debe'] for a in asi['asientos'] if a['cuenta'] == '477'), 2)
        sop = round(sum(a['debe'] - a['haber'] for a in asi['asientos'] if a['cuenta'] == '472'), 2)
        res['libro'] = {'iva_repercutido_477': rep, 'iva_soportado_472': sop,
                        'cuadra': abs(rep - res['m303']['c27_devengado']) <= 0.011 and abs(sop - res['m303']['c45_deducible']) <= 0.011}
    except Exception:
        res['libro'] = None
    lim = int(request.args.get('limite') or 200)
    res['sii']['expedidas'] = res['sii']['expedidas'][:lim]
    res['sii']['recibidas'] = res['sii']['recibidas'][:lim]
    return jsonify({'ok': True, 'hotel': hotel or '', **res})


@cierre_bp.route('/api/exportar/fiscal')
def api_exportar_fiscal():
    import fiscal as FI
    mes, hotel = _args()
    buf, nombre = FI.exportar_excel(FI.fiscal_completo(mes, hotel, **_dirs()))
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ── Archivo de fin de mes para la central (Ola B · bloque 5) ────────────────
def _aging(hotel):
    try:
        import aging_ap as AG
        import almacen_datos as ALM
        from dashboard import cargar_datos_ap, cargar_datos
        df_ap = cargar_datos_ap()
        df_ar, _ = cargar_datos()
        bk, _ = ALM.movimientos_banco(datos_dir=str(_t_ddir()), reportes_dir=str(_t_rdir()))
        return AG.calcular_aging(df_ap, df_ar, bk)
    except Exception:
        return None


def _bloques(mes, hotel, con_fiscal=True):
    """Todos los bloques del cierre, cada uno protegido: uno que falle no tumba el paquete."""
    import cierre_mes as CM
    out = {}
    try:
        out['asientos'], out['reconciliacion'] = CM.cierre_completo(mes, hotel, **_dirs())
        out['drr'] = CM.drr_del_mes(out['asientos']['mes'], hotel)
    except Exception:
        out['asientos'] = out['reconciliacion'] = out['drr'] = None
    for clave, fn in (('banco', lambda: _cuadre(mes, hotel)), ('inventarios', lambda: _inventarios(mes, hotel)),
                      ('inmovilizado', lambda: _inmovilizado(mes, hotel)), ('aging', lambda: _aging(hotel))):
        try:
            out[clave] = fn()
        except Exception:
            out[clave] = None
    try:
        import provisiones as PV
        d = _dirs()
        out['provisiones'] = [PV.provision_albaranes(mes, hotel, d['procesadas_dir'], d['reportes_dir'], d['datos_dir']),
                              PV.provision_comisiones(mes, hotel, d['reportes_dir'], d['datos_dir'])]
    except Exception:
        out['provisiones'] = None
    out['fiscal'] = None
    if con_fiscal:
        try:
            import fiscal as FI
            out['fiscal'] = FI.resumen_para_paquete(mes, out.get('asientos'), out.get('reconciliacion'), hotel, **_dirs())
        except Exception:
            out['fiscal'] = None
    return out


@cierre_bp.route('/api/cierre/paquete')
def api_cierre_paquete():
    import paquete_cierre as PQ
    from provisiones import _mes_a_rango
    mes, hotel = _args()
    _, _, mes = _mes_a_rango(mes)
    b = _bloques(mes, hotel)
    paq = PQ.montar(mes, b['asientos'], b['reconciliacion'], b['banco'], b['provisiones'], b['inventarios'],
                    b['inmovilizado'], b['aging'], b['fiscal'], PQ.comentarios(mes, str(_t_ddir())), b['drr'])
    return jsonify({'ok': True, 'hotel': hotel or '', **paq})


@cierre_bp.route('/api/cierre/comentario', methods=['POST'])
def api_cierre_comentario():
    import paquete_cierre as PQ
    data = request.get_json(force=True, silent=True) or {}
    mes = str(data.get('mes') or '').strip()[:7]
    if len(mes) != 7:
        return jsonify({'ok': False, 'error': 'mes invalido (aaaa-mm)'}), 400
    try:
        from flask_login import current_user
        usuario = getattr(current_user, 'username', None) or 'sistema'
    except Exception:
        usuario = 'sistema'
    try:
        m = PQ.guardar_comentario(mes, str(data.get('seccion') or ''), str(data.get('texto') or ''), usuario, str(_t_ddir()))
    except ValueError as e:
        return jsonify({'ok': False, 'error': str(e)}), 400
    try:
        from dashboard import _audit
        _audit('CIERRE_COMENTARIO', f"{mes} {data.get('seccion')}", usuario)
    except Exception:
        pass
    return jsonify({'ok': True, 'comentarios': m})


@cierre_bp.route('/api/exportar/cierre_paquete')
def api_exportar_cierre_paquete():
    import paquete_cierre as PQ
    from provisiones import _mes_a_rango
    mes, hotel = _args()
    _, _, mes = _mes_a_rango(mes)
    b = _bloques(mes, hotel)
    paq = PQ.montar(mes, b['asientos'], b['reconciliacion'], b['banco'], b['provisiones'], b['inventarios'],
                    b['inmovilizado'], b['aging'], b['fiscal'], PQ.comentarios(mes, str(_t_ddir())), b['drr'])
    buf, nombre = PQ.exportar_excel(paq, b['asientos'], b['reconciliacion'], b['banco'], b['provisiones'],
                                    b['inventarios'], b['inmovilizado'], b['aging'], b['fiscal'])
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
