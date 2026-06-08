"""
Flask blueprint para AR Real — integración al dashboard
Endpoint: /api/procesar_ar_real
"""
import os
import sys
import subprocess
from flask import Blueprint, Response
from datetime import datetime
import threading

ar_real_bp = Blueprint('ar_real', __name__)

_lock = threading.Lock()
_ar_real_running = False

@ar_real_bp.route('/api/procesar_ar_real', methods=['GET'])
def api_procesar_ar_real():
    """
    Procesa archivos de grupos corporativos del Hilton
    Genera reporte Excel consolidado con rooming list + facturas
    """
    def generar():
        global _ar_real_running
        
        with _lock:
            if _ar_real_running:
                yield "data: Proceso AR Real en curso. Espere...\n\n"
                return
            _ar_real_running = True
        
        try:
            # Ejecutar el módulo de procesamiento
            result = subprocess.run(
                [sys.executable, "ar_grupos_reales.py"],
                cwd=os.path.dirname(os.path.abspath(__file__)),
                capture_output=True,
                text=True,
                timeout=60
            )
            
            # Enviar stdout línea por línea
            for line in result.stdout.splitlines():
                yield f"data: {line}\n\n"
            
            if result.returncode != 0:
                yield f"data: ERROR: {result.stderr}\n\n"
            else:
                yield "data: AR_REAL_COMPLETO\n\n"
        
        except subprocess.TimeoutExpired:
            yield "data: ERROR: Timeout en procesamiento\n\n"
        except Exception as e:
            yield f"data: ERROR: {str(e)}\n\n"
        finally:
            _ar_real_running = False
    
    return Response(
        generar(),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no"
        }
    )

@ar_real_bp.route('/api/ar_real_status', methods=['GET'])
def api_ar_real_status():
    """
    Retorna el estado de los últimos reportes procesados
    """
    reportes = []
    reportes_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes")
    
    if os.path.exists(reportes_dir):
        for f in sorted(os.listdir(reportes_dir)):
            if f.startswith("ar_real_") and f.endswith(".xlsx"):
                filepath = os.path.join(reportes_dir, f)
                size = os.path.getsize(filepath)
                mtime = os.path.getmtime(filepath)
                reportes.append({
                    "filename": f,
                    "size_kb": round(size / 1024, 2),
                    "timestamp": datetime.fromtimestamp(mtime).isoformat()
                })
    
    return {
        "status": "running" if _ar_real_running else "idle",
        "reportes": sorted(reportes, key=lambda x: x["timestamp"], reverse=True)[:5]
    }

import pandas as pd
from flask import jsonify as _jsonify
import os as _os

@ar_real_bp.route('/api/ar_real_data', methods=['GET'])
def api_ar_real_data():
    """Retorna clientes corporativos y reservas para el dashboard AR Real."""
    base = _os.path.dirname(_os.path.abspath(__file__))
    datos = _os.path.join(base, 'datos-referencia')
    try:
        # Simple cache for small files
        import time as _t2
        _ar_cache = getattr(api_ar_real_data, '_cache', {})
        now2 = _t2.time()
        if 'data' in _ar_cache and now2 - _ar_cache.get('ts', 0) < 300:
            df_c, df_r = _ar_cache['data']
        else:
            df_c = pd.read_excel(_os.path.join(datos, 'clientes_credito.xlsx'))
            df_r = pd.read_excel(_os.path.join(datos, 'reservas_credito.xlsx'))
            api_ar_real_data._cache = {'data': (df_c, df_r), 'ts': now2}
        df_r['fecha_entrada'] = pd.to_datetime(df_r['fecha_entrada'], errors='coerce')
        df_r['fecha_emision'] = pd.to_datetime(df_r['fecha_emision'], errors='coerce')

        # KPIs
        pend_fact = float(df_r[df_r['estado']=='PENDIENTE_FACTURA']['total'].sum())
        facturado = float(df_r[df_r['estado']=='FACTURADO']['total'].sum())
        cobrado   = float(df_r[df_r['estado']=='COBRADO']['total'].sum())
        saldo_total = float(df_c['saldo_pendiente'].sum())
        
        # Clients enriched with reservation count
        clientes = []
        for _, c in df_c.iterrows():
            nombre = str(c['nombre_cliente'])
            reservas_cliente = df_r[df_r['cliente'] == nombre]
            clientes.append({
                'nombre': nombre,
                'NIF': str(c.get('NIF','')),
                'email': str(c.get('email','')),
                'dias_pago': int(c.get('dias_pago', 30)),
                'limite_credito': float(c.get('limite_credito', 0)),
                'saldo_pendiente': float(c.get('saldo_pendiente', 0)),
                'num_reservas': len(reservas_cliente),
                'reservas_pendientes': int((reservas_cliente['estado']=='PENDIENTE_FACTURA').sum()),
                'status': 'critical' if float(c.get('saldo_pendiente',0)) > float(c.get('limite_credito',0))*0.8
                          else 'warning' if float(c.get('saldo_pendiente',0)) > 0
                          else 'ok',
            })
        
        # Recent reservations
        reservas = []
        for _, r in df_r.sort_values('fecha_entrada', ascending=False).iterrows():
            fe = r['fecha_entrada']
            em = r['fecha_emision']
            reservas.append({
                'numero': str(r['numero_reserva']),
                'cliente': str(r['cliente']),
                'fecha_entrada': fe.strftime('%d/%m/%Y') if pd.notna(fe) else '—',
                'fecha_salida': pd.to_datetime(r['fecha_salida']).strftime('%d/%m/%Y') if pd.notna(r['fecha_salida']) else '—',
                'habitaciones': int(r['habitaciones']),
                'total': float(r['total']),
                'estado': str(r['estado']),
                'fecha_emision': em.strftime('%d/%m/%Y') if pd.notna(em) else '—',
            })
        
        return _jsonify({
            'kpis': {
                'pendiente_facturar': pend_fact,
                'facturado': facturado,
                'cobrado': cobrado,
                'saldo_total': saldo_total,
                'num_clientes': len(df_c),
                'reservas_pendientes': int((df_r['estado']=='PENDIENTE_FACTURA').sum()),
            },
            'clientes': clientes,
            'reservas': reservas,
        })
    except Exception as e:
        return _jsonify({'error': str(e)}), 500



@ar_real_bp.route('/api/ar_real/emitir_factura', methods=['POST'])
def api_emitir_factura():
    """Emite una nueva factura corporativa y la registra en reservas_credito.xlsx."""
    import json as _json
    from datetime import datetime
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
        ruta = BASE_DIR / 'datos-referencia' / 'reservas_credito.xlsx'
        if ruta.exists():
            df = pd.read_excel(ruta)
        else:
            df = pd.DataFrame(columns=['numero_reserva','cliente','fecha_entrada','fecha_salida',
                                       'habitaciones','importe_habitaciones','importe_fb',
                                       'importe_extras','total','estado','fecha_emision'])

        # Generate invoice number
        year = datetime.now().year
        last_num = len(df) + 1
        numero = f'FAC-{year}-CORP-{last_num:04d}'

        noches    = max(1, (pd.to_datetime(fecha_salida) - pd.to_datetime(fecha_entrada)).days)
        importe_h = round(habitaciones * noches * precio_noche, 2)

        new_row = {
            'numero_reserva':     numero,
            'cliente':            cliente,
            'fecha_entrada':      fecha_entrada,
            'fecha_salida':       fecha_salida,
            'habitaciones':       habitaciones,
            'importe_habitaciones': importe_h,
            'importe_fb':         round(fb_extras, 2),
            'importe_extras':     0.0,
            'total':              round(total, 2),
            'estado':             'EMITIDA',
            'fecha_emision':      datetime.now().strftime('%Y-%m-%d'),
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel(ruta, index=False)
        return jsonify({'ok': True, 'numero': numero, 'total': total})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500
