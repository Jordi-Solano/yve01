"""
tab_ar_real.py — Yve.01 AR Real (Cuentas a Cobrar Corporativas)
Gestión completa del ciclo AR: grupos corporativos, BEOs, facturación
"""
import os as _os, json as _json
from datetime import datetime, date, timedelta
from flask import Blueprint, jsonify, request, send_file, session
import pandas as pd

ar_real_bp = Blueprint('ar_real', __name__)
BASE_DIR = _os.path.dirname(_os.path.abspath(__file__))
from tenant_dirs import datos_dir as _t_ddir, reportes_dir as _t_rdir

class _TStr(str):
    pass

def _mk_dyn(fn):
    class _D:
        def __add__(self, o): return fn() + o
        def __radd__(self, o): return o + fn()
        def __str__(self): return fn()
        def __fspath__(self): return fn()
    return _D()

DATOS = _mk_dyn(_t_ddir)
REPORTES = _mk_dyn(_t_rdir)

def _get_clientes():
    ruta = _os.path.join(DATOS, 'clientes_credito.xlsx')
    if not _os.path.exists(ruta): return pd.DataFrame()
    return pd.read_excel(ruta)

def _get_reservas():
    ruta = _os.path.join(DATOS, 'reservas_credito.xlsx')
    if not _os.path.exists(ruta): return pd.DataFrame()
    df = pd.read_excel(ruta)
    # filtro por hotel activo de la sesión (si el df tiene columna hotel)
    try:
        _h = session.get('hotel_activo')
        if _h and 'hotel' in df.columns:
            df = df[df['hotel'].astype(str).str.contains(_h, case=False, na=False, regex=False)].copy()
    except Exception:
        pass
    for col in ['fecha_entrada','fecha_salida','fecha_emision']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    return df

def _save_reservas(df):
    ruta = _os.path.join(DATOS, 'reservas_credito.xlsx')
    df.to_excel(ruta, index=False)

def _aging_bucket(fecha_emision):
    """Returns aging bucket: 0-30, 31-60, 61-90, >90 days"""
    if pd.isna(fecha_emision): return 'Sin fecha'
    days = (date.today() - pd.Timestamp(fecha_emision).date()).days
    if days <= 30:  return '0-30 días'
    if days <= 60:  return '31-60 días'
    if days <= 90:  return '61-90 días'
    return '>90 días (VENCIDA)'

@ar_real_bp.route('/api/ar_real/clientes')
def api_clientes():
    """Lista de clientes corporativos con estado de crédito."""
    try:
        df_c = _get_clientes()
        df_r = _get_reservas()
        clientes = []
        for _, c in df_c.iterrows():
            nombre = str(c.get('nombre_cliente',''))
            # Pending invoices for this client
            pend = df_r[(df_r['cliente'] == nombre) & 
                        (df_r['estado'].isin(['FACTURADO','PENDIENTE_FACTURA']))] if len(df_r) else pd.DataFrame()
            total_pend = float(pend['total'].sum()) if len(pend) else 0
            has_overdue = False
            if len(pend) and 'fecha_emision' in pend.columns:
                for _, row in pend.iterrows():
                    if pd.notna(row.get('fecha_emision')):
                        days = (date.today() - pd.Timestamp(row['fecha_emision']).date()).days
                        if days > 60: has_overdue = True; break
            limit = float(c.get('limite_credito', 0) or 0)
            uso_pct = round(total_pend / limit * 100, 1) if limit > 0 else 0
            clientes.append({
                'nombre':    nombre,
                'NIF':       str(c.get('NIF','')),
                'email':     str(c.get('email','')),
                'telefono':  str(c.get('telefono','')),
                'dias_pago': int(c.get('dias_pago', 30) or 30),
                'limite_credito':   round(limit, 2),
                'saldo_pendiente':  round(total_pend, 2),
                'uso_credito_pct':  uso_pct,
                'tiene_vencidas':   has_overdue,
                'facturas_pendientes': len(pend),
            })
        return jsonify({'ok': True, 'clientes': clientes})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@ar_real_bp.route('/api/ar_real/facturas')
def api_facturas():
    """Lista de facturas con aging y estado."""
    try:
        df_r = _get_reservas()
        if df_r.empty:
            return jsonify({'ok': True, 'facturas': [], 'stats': {}})
        
        facturas = []
        total_pend = total_venc = total_cobr = 0
        aging = {'0-30 días': 0, '31-60 días': 0, '61-90 días': 0, '>90 días (VENCIDA)': 0}
        
        for _, row in df_r.iterrows():
            total = float(row.get('total', 0) or 0)
            estado = str(row.get('estado', ''))
            fecha_em = row.get('fecha_emision')
            bucket = _aging_bucket(fecha_em) if estado == 'FACTURADO' else 'N/A'
            
            days_pending = None
            if pd.notna(fecha_em) and estado == 'FACTURADO':
                days_pending = (date.today() - pd.Timestamp(fecha_em).date()).days
                if bucket in aging: aging[bucket] += total
                if days_pending > 60: total_venc += total
                else: total_pend += total
            elif estado == 'COBRADO':
                total_cobr += total
            
            facturas.append({
                'numero':        str(row.get('numero_reserva','')),
                'cliente':       str(row.get('cliente','')),
                'fecha_entrada': str(row.get('fecha_entrada',''))[:10] if pd.notna(row.get('fecha_entrada')) else '',
                'fecha_salida':  str(row.get('fecha_salida',''))[:10] if pd.notna(row.get('fecha_salida')) else '',
                'habitaciones':  int(row.get('habitaciones', 1) or 1),
                'importe_hab':   float(row.get('importe_habitaciones', 0) or 0),
                'importe_fb':    float(row.get('importe_fb', 0) or 0),
                'importe_extras':float(row.get('importe_extras', 0) or 0),
                'total':         round(total, 2),
                'estado':        estado,
                'fecha_emision': str(fecha_em)[:10] if pd.notna(fecha_em) else '',
                'aging_bucket':  bucket,
                'days_pending':  days_pending,
            })
        
        stats = {
            'total_facturas': len(facturas),
            'pendiente':   round(total_pend, 2),
            'vencido':     round(total_venc, 2),
            'cobrado_mes': round(total_cobr, 2),
            'aging':       aging,
        }
        return jsonify({'ok': True, 'facturas': facturas, 'stats': stats})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@ar_real_bp.route('/api/ar_real/cobrar', methods=['POST'])
def api_cobrar():
    """Marca una factura como cobrada."""
    data = request.get_json(force=True, silent=True) or {}
    numero = data.get('numero', '').strip()
    if not numero:
        return jsonify({'ok': False, 'error': 'Número de factura requerido'}), 400
    try:
        df = _get_reservas()
        mask = df['numero_reserva'].astype(str) == numero
        if not mask.any():
            return jsonify({'ok': False, 'error': f'Factura {numero} no encontrada'}), 404
        df.loc[mask, 'estado'] = 'COBRADO'
        df.loc[mask, 'fecha_cobro'] = date.today().isoformat()
        _save_reservas(df)
        total = float(df.loc[mask, 'total'].values[0])
        return jsonify({'ok': True, 'numero': numero, 'total': total,
                        'message': f'Factura {numero} marcada como cobrada'})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@ar_real_bp.route('/api/ar_real/recordatorio', methods=['POST'])
def api_recordatorio():
    """Envía email de recordatorio de pago al cliente."""
    data = request.get_json(force=True, silent=True) or {}
    numero = data.get('numero', '').strip()
    try:
        df = _get_reservas()
        df_c = _get_clientes()
        mask = df['numero_reserva'].astype(str) == numero
        if not mask.any():
            return jsonify({'ok': False, 'error': 'Factura no encontrada'}), 404
        row = df[mask].iloc[0]
        cliente_nombre = str(row['cliente'])
        total = float(row.get('total', 0))
        fecha_em = str(row.get('fecha_emision',''))[:10]
        
        # Get client email
        c_match = df_c[df_c['nombre_cliente'] == cliente_nombre]
        email_dest = str(c_match.iloc[0]['email']) if len(c_match) else ''
        dias_pago = int(c_match.iloc[0]['dias_pago']) if len(c_match) else 30
        
        if not email_dest or '@' not in email_dest:
            return jsonify({'ok': False, 'error': f'Email de {cliente_nombre} no disponible'}), 400
        
        # Send reminder via notificaciones
        from notificaciones import enviar_email, _email_html
        asunto = f'Recordatorio de pago — Factura {numero} — €{total:,.2f}'
        cuerpo = _email_html(
            f'Recordatorio de pago — {numero}',
            [
                f'Cliente: {cliente_nombre}',
                f'Importe: €{total:,.2f} (IVA incluido)',
                f'Fecha emisión: {fecha_em}',
                f'Condiciones de pago: {dias_pago} días',
                f'Estado: PENDIENTE DE COBRO',
            ],
            color='#f59e0b',
            footer_note='Por favor, realice el pago según las condiciones acordadas. Para cualquier consulta, contacte con nuestro departamento financiero.'
        )
        ok = enviar_email(email_dest, asunto, cuerpo, 'ar_recordatorio')
        if ok:
            # Log the reminder
            df.loc[mask, 'ultimo_recordatorio'] = date.today().isoformat()
            _save_reservas(df)
        return jsonify({'ok': ok, 'email': email_dest, 
                        'message': f'Recordatorio enviado a {email_dest}' if ok else 'Error enviando email'})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@ar_real_bp.route('/api/ar_real/emitir_factura', methods=['POST'])
def api_emitir_factura():
    """Emite una nueva factura corporativa y la registra."""
    data = request.get_json(force=True, silent=True) or {}
    cliente       = data.get('cliente', '').strip()
    fecha_entrada = data.get('fecha_entrada', '')
    fecha_salida  = data.get('fecha_salida', '')
    habitaciones  = int(data.get('habitaciones', 1))
    precio_noche  = float(data.get('precio_noche', 0))
    fb_extras     = float(data.get('fb_extras', 0))
    total         = float(data.get('total', 0))
    if not cliente or not fecha_entrada or not fecha_salida:
        return jsonify({'ok': False, 'error': 'Faltan datos obligatorios'}), 400
    try:
        df = _get_reservas()
        year = datetime.now().year
        last_num = len(df) + 1
        numero = f'FAC-{year}-CORP-{last_num:04d}'
        noches = max(1, (pd.to_datetime(fecha_salida) - pd.to_datetime(fecha_entrada)).days)
        importe_h = round(habitaciones * noches * precio_noche, 2)
        new_row = {
            'numero_reserva': numero,
            'cliente':        cliente,
            'fecha_entrada':  fecha_entrada,
            'fecha_salida':   fecha_salida,
            'habitaciones':   habitaciones,
            'importe_habitaciones': importe_h,
            'importe_fb':     round(fb_extras, 2),
            'importe_extras': 0.0,
            'total':          round(total, 2),
            'estado':         'FACTURADO',
            'fecha_emision':  datetime.now().strftime('%Y-%m-%d'),
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        _save_reservas(df)
        return jsonify({'ok': True, 'numero': numero, 'total': total})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

@ar_real_bp.route('/api/ar_real/stats')
@ar_real_bp.route('/api/ar_real/data')
def api_ar_real_stats():
    """Quick stats for mobile KPI bar."""
    try:
        df_r = _get_reservas()
        df_c = _get_clientes()
        if df_r.empty:
            return jsonify({'ok': True, 'pendiente': 0, 'vencido': 0, 'n_clientes': len(df_c)})
        
        pendiente = vencido = 0
        for _, row in df_r.iterrows():
            if str(row.get('estado','')) == 'FACTURADO':
                total = float(row.get('total', 0) or 0)
                fem = row.get('fecha_emision')
                if pd.notna(fem):
                    days = (date.today() - pd.Timestamp(fem).date()).days
                    if days > 60: vencido += total
                    else: pendiente += total
        
        return jsonify({
            'ok': True,
            'pendiente': round(pendiente, 2),
            'vencido':   round(vencido, 2),
            'n_clientes': len(df_c),
        })
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500
