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

def _hotel_activo():
    try:
        import censo_hoteles as _censo
        return _censo.activo()
    except Exception:
        return ""


def _get_reservas_todas():
    """El fichero ENTERO, sin filtrar. Para quien vaya a reescribirlo."""
    ruta = _os.path.join(DATOS, 'reservas_credito.xlsx')
    if not _os.path.exists(ruta):
        return pd.DataFrame()
    return pd.read_excel(ruta)


def _get_reservas():
    ruta = _os.path.join(DATOS, 'reservas_credito.xlsx')
    if not _os.path.exists(ruta): return pd.DataFrame()
    df = pd.read_excel(ruta)
    # FASE 5: se cruza por `hotel_id`, no por el NOMBRE.
    #
    # Lo que habia era un `contains` contra la columna `hotel` con el nombre
    # del censo. Es el fallo que la fase 0 vino a matar y que aqui se habia
    # quedado vivo: con "Hotel Sol" y "Hotel Sol Mar" en el mismo grupo, el
    # primero se llevaba las reservas del segundo.
    try:
        from almacen_datos import solo_del_hotel_activo as _solo
        df = _solo(df)
    except Exception:
        pass
    for col in ['fecha_entrada','fecha_salida','fecha_emision']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    return df

def _save_reservas(df):
    """Guarda `df` COMO las reservas del hotel activo, sin tocar las de los demas.

    Ojo, que aqui habia una mina: todos los que llaman a esto parten de
    `_get_reservas()`, que devuelve solo las del hotel activo. Con un
    `df.to_excel()` a pelo, guardar despues de cobrar una factura reescribia
    el fichero entero con SOLO las filas de ese hotel — o sea, borraba las
    reservas de los otros. Estaba latente porque el filtro por nombre no
    llegaba a filtrar casi nunca; al ponerlo por id se habria activado.

    Asi que se recompone: del fichero completo se quitan las filas del hotel
    activo y se pegan las que llegan. Sin hotel activo (0 hoteles o vista de
    grupo) se guarda tal cual, que es el comportamiento de siempre.
    """
    ruta = _os.path.join(DATOS, 'reservas_credito.xlsx')
    hid = _hotel_activo()
    if hid:
        try:
            from almacen_datos import COL_HOTEL as _COLH
            completo = _get_reservas_todas()
            if not completo.empty and _COLH in completo.columns:
                _col = completo[_COLH].map(lambda v: "" if v is None else str(v).strip())
                otros = completo[_col != str(hid)]
                df = pd.concat([otros, df], ignore_index=True)
        except Exception:
            pass
    df.to_excel(ruta, index=False)

def _cotejar_beo(beo, df_r):
    """Compara el total del BEO con el importe de la factura/reserva del mismo evento."""
    try:
        if df_r is None or not len(df_r):
            return {"estado": "sin_factura"}
        num = str(beo.get('numero_reserva') or '')
        total_beo = float(beo.get('total') or 0)
        fila = None
        for col in ('numero_reserva', 'numero'):
            if col in df_r.columns and num:
                m = df_r[df_r[col].astype(str) == num]
                if len(m):
                    fila = m.iloc[0]; break
        if fila is None:
            return {"estado": "sin_factura", "total_beo": round(total_beo, 2)}
        total_fac = float(fila.get('total') or fila.get('importe') or 0)
        if total_beo <= 0:
            return {"estado": "sin_importe"}
        diff = abs(total_fac - total_beo)
        pct = diff / total_beo * 100 if total_beo else 0
        return {"estado": "cuadra" if pct <= 5 else "discrepancia",
                "total_beo": round(total_beo, 2), "total_factura": round(total_fac, 2),
                "diff": round(diff, 2), "diff_pct": round(pct, 1)}
    except Exception:
        return {"estado": "error"}

def _aging_bucket(fecha_emision):
    """Returns aging bucket: 0-30, 31-60, 61-90, >90 days"""
    if pd.isna(fecha_emision): return 'Sin fecha'
    days = (date.today() - pd.Timestamp(fecha_emision).date()).days
    if days <= 30:  return '0-30 días'
    if days <= 60:  return '31-60 días'
    if days <= 90:  return '61-90 días'
    return '>90 días (VENCIDA)'

def _campo_cliente(fila, *nombres):
    """El primer campo que exista de una lista de nombres posibles.

    Hacia falta porque el codigo leia claves que NO estan en el fichero:
    `limite_credito` cuando la columna es `credito_limite`, y `NIF` cuando es
    `nif`. Resultado: el limite de credito salia SIEMPRE 0 —y con el 0, el
    porcentaje de uso tambien— y el NIF vacio. Estaba latente solo porque
    nadie habia dado de alta un cliente todavia.

    Se aceptan las dos grafias en vez de cambiar por la buena y ya: si alguien
    tiene el fichero con la otra, sigue funcionando. Un lector tolerante no
    cuesta nada; un fichero que deja de leerse, si.
    """
    for n in nombres:
        if n in fila and fila.get(n) is not None:
            v = fila.get(n)
            if not (isinstance(v, float) and v != v):      # descarta NaN
                return v
    return None


def _txt_cliente(fila, *nombres):
    v = _campo_cliente(fila, *nombres)
    s = "" if v is None else str(v).strip()
    return "" if s.lower() in ("nan", "none", "nat") else s


def _num_cliente(fila, *nombres):
    v = _campo_cliente(fila, *nombres)
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


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
            limit = _num_cliente(c, 'credito_limite', 'limite_credito')
            uso_pct = round(total_pend / limit * 100, 1) if limit > 0 else 0
            clientes.append({
                'nombre':    nombre,
                'NIF':       _txt_cliente(c, 'nif', 'NIF', 'cif'),
                'email':     _txt_cliente(c, 'email', 'correo'),
                'telefono':  _txt_cliente(c, 'telefono', 'teléfono', 'tel'),
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


@ar_real_bp.route('/api/ar_real/procesar_contrato', methods=['POST'])
def api_procesar_contrato():
    """Sube las fotos de un contrato de grupo y lo procesa (visión) -> AR Real + comisión + DI."""
    import tempfile as _tmp, shutil as _sh
    files = request.files.getlist('files') or request.files.getlist('fotos')
    if not files:
        return jsonify({'ok': False, 'error': 'No se recibieron fotos'}), 400
    carpeta = _tmp.mkdtemp(prefix='contrato_')
    guardadas = []
    try:
        for f in files:
            if not f or not f.filename:
                continue
            ext = _os.path.splitext(f.filename)[1].lower()
            if ext not in ('.jpg', '.jpeg', '.png', '.webp', '.heic'):
                continue
            ruta = _os.path.join(carpeta, _os.path.basename(f.filename))
            f.save(ruta)
            guardadas.append(ruta)
        if not guardadas:
            return jsonify({'ok': False, 'error': 'No hay imágenes válidas (jpg/png)'}), 400
        from lector_contratos_grupo import procesar_contrato_grupo
        res = procesar_contrato_grupo(guardadas)
        return jsonify(res)
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)[:200]}), 500
    finally:
        _sh.rmtree(carpeta, ignore_errors=True)


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
        # El contador va sobre el fichero ENTERO, no sobre el df filtrado.
        #
        # Al poner el filtro por hotel, `_get_reservas()` paso a devolver solo
        # las del hotel activo y este `len(df)+1` empezo a contar desde 1 en
        # cada hotel: emitir en dos hoteles daba DOS facturas distintas con el
        # mismo numero. Visto en produccion, no leyendo el codigo.
        #
        # Se deja la serie del grupo, continua, que es como estaba antes de
        # tocar nada: un numero de factura repetido es un problema contable, y
        # cambiar a una serie por hotel es una decision de producto, no algo
        # que deba caerse por un efecto colateral.
        last_num = len(_get_reservas_todas()) + 1
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
            'hotel_id':       _hotel_activo(),      # fase 5
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        _save_reservas(df)
        return jsonify({'ok': True, 'numero': numero, 'total': total})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

@ar_real_bp.route('/api/ar_real/beos')
def api_ar_real_beos():
    """BEOs generados automáticamente desde los contratos (con estado de cotejo vs factura)."""
    try:
        ruta = _os.path.join(DATOS, 'beos_generados.json')
        beos = _json.load(open(ruta, encoding='utf-8')) if _os.path.exists(ruta) else []
        # Un BEO es el catering de un evento en UN hotel (fase 5). Los que no
        # llevan etiqueta son de antes: se ven en la vista de grupo, como el
        # resto de "sin asignar", pero no se le cuelgan a ningun hotel.
        _hid = _hotel_activo()
        if _hid:
            beos = [b for b in beos if str(b.get('hotel_id') or '') == str(_hid)]
        df_r = _get_reservas()
        for b in beos:
            b['cotejo'] = _cotejar_beo(b, df_r)
        beos.sort(key=lambda b: b.get('fecha_generado', ''), reverse=True)
        return jsonify({"ok": True, "beos": beos})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e), "beos": []})


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
