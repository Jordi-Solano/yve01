"""
Blueprint Flask para descarga de reportes
"""
from flask import Blueprint, send_file, jsonify
from exportador_reportes import exportar_excel, crear_reporte_calipolis_excel
from ar_real_completo import procesar_ar_real_completo

exportador_bp = Blueprint('exportador', __name__)

@exportador_bp.route('/api/exportar/<tipo>')
def api_exportar(tipo):
    """Descarga reporte en Excel"""
    valid_tipos = ['ar', 'ap', 'drr', 'multihotel', 'banco', 'fb', 'ar_real']
    
    if tipo not in valid_tipos and tipo != 'ar_real' and tipo != 'calipolis':
        return jsonify({"error": "Tipo de reporte inválido"}), 400
    
    try:
        if tipo == 'ar_real':
            output_file = procesar_ar_real_completo()
            if not output_file:
                return jsonify({"error": "No se pudo procesar AR Real"}), 500
            output = open(output_file, 'rb').read()
            from io import BytesIO
            output = BytesIO(output)
            import os
            filename = os.path.basename(output_file)
            result = (output, filename)
        elif tipo == 'banco':
            # Antes se mandaba TAL CUAL el ultimo conciliacion_*.xlsx, asi que la
            # descarga se quedaba en la foto del dia que se concilio y le faltaban
            # los movimientos subidos despues. Ahora sale del mismo sitio que la
            # pantalla y el panel: extracto real + estado del informe.
            from datetime import datetime as _dt
            from tenant_dirs import reportes_dir as _t_rdir
            import almacen_datos as _alm
            from io import BytesIO
            _df_bk, _info_bk = _alm.movimientos_banco(reportes_dir=_t_rdir())
            if _df_bk is None or _df_bk.empty:
                return jsonify({'error': 'No hay datos bancarios'}), 404
            _buf = BytesIO()
            _df_bk.to_excel(_buf, index=False)
            _buf.seek(0)
            result = (_buf, f'banco_movimientos_{_dt.now().strftime("%Y%m%d")}.xlsx')
        elif tipo == 'ar_real':
            import os as _os
            ruta = __import__('tenant_dirs').datos_dir() + '/reservas_credito.xlsx'
            if not _os.path.exists(ruta): return jsonify({'error': 'Sin datos AR Real'}), 404
            from io import BytesIO as _BIO
            from datetime import datetime as _dt
            # Se reconstruye filtrado en vez de mandar el fichero crudo (fase
            # 5). Mandandolo tal cual, con un hotel elegido te descargabas las
            # facturas de TODOS: un Excel que sale de la app y se manda por
            # correo es la peor forma de que se escape lo de otro hotel.
            import pandas as _pd
            _df = _pd.read_excel(ruta)
            try:
                from almacen_datos import solo_del_hotel_activo as _solo
                _df = _solo(_df)
            except Exception:
                pass
            _buf2 = _BIO()
            _df.to_excel(_buf2, index=False)
            _buf2.seek(0)
            result = (_buf2, f'ar_real_facturas_{_dt.now().strftime("%Y%m%d")}.xlsx')
        elif tipo == 'multihotel':
            # FASE B: sale del agregador, igual que la pantalla.
            #
            # Lo que habia aqui devolvia seis hoteles escritos a mano en
            # `exportador_reportes.py` ("Premier London Mayfair"...) con el
            # titulo fijo "Junio 2025". Lo que veias y lo que te descargabas no
            # tenian nada que ver, y un Excel que sale de la app y se manda por
            # correo es la peor forma de enterarse.
            from io import BytesIO as _BIO3
            from datetime import datetime as _dt3
            import pandas as _pd3
            from agregador_grupo import agregado as _agregado
            _ag = _agregado()
            _filas = []
            for _f in _ag['hoteles'] + [_ag['sin_asignar'], _ag['desconocido'], _ag['grupo']]:
                # Las cajas vacias de "sin asignar" y "desconocido" no se
                # escriben: en pantalla tampoco salen, y una fila a cero en un
                # Excel se lee como un hotel que no factura.
                _vacia = not (_f['ap']['facturas'] or _f['ar_ota']['facturas']
                              or _f['ar_real']['facturas'] or _f['fb']['ventas'])
                if _vacia and _f['hotel_id'] in ('sin_asignar', 'desconocido'):
                    continue
                _filas.append({
                    'Hotel': _f['nombre'],
                    'Facturas AP': _f['ap']['facturas'],
                    'Importe AP (EUR)': _f['ap']['importe'],
                    'AP con incidencia': _f['ap']['discrepancias'],
                    'Facturas OTA': _f['ar_ota']['facturas'],
                    'Bruto OTA (EUR)': _f['ar_ota']['importe_bruto'],
                    'Reclamable OTA (EUR)': _f['ar_ota']['importe_reclamable'],
                    'DI pendientes': _f['ar_ota']['di_pendientes'],
                    'Facturas AR Real': _f['ar_real']['facturas'],
                    'Por cobrar (EUR)': _f['ar_real']['pendiente'],
                    'Vencido (EUR)': _f['ar_real']['vencido'],
                    'Ventas F&B (EUR)': _f['fb']['ventas'],
                    'Food cost %': _f['fb']['food_cost_pct'],
                    'Coste mermas (EUR)': _f['fb']['coste_mermas'],
                })
            _buf3 = _BIO3()
            with _pd3.ExcelWriter(_buf3, engine='openpyxl') as _w:
                _pd3.DataFrame(_filas).to_excel(_w, index=False, sheet_name='Multi-Hotel')
                # El cuadre viaja CON el Excel. Si un dia no cuadra, quien abra
                # el fichero tiene que poder verlo sin volver a la aplicacion.
                _pd3.DataFrame(_ag['cuadre']).to_excel(_w, index=False, sheet_name='Cuadre')
            _buf3.seek(0)
            result = (_buf3, f'multihotel_{_dt3.now().strftime("%Y%m%d")}.xlsx')
        elif tipo == 'fb':
            import os as _os
            ruta = __import__('tenant_dirs').datos_dir() + '/ventas_fb_diarias.xlsx'
            if not _os.path.exists(ruta): return jsonify({'error': 'Sin datos F&B'}), 404
            from io import BytesIO
            with open(ruta,'rb') as fh: data = fh.read()
            from datetime import datetime as _dt
            result = (BytesIO(data), f'fb_ventas_{_dt.now().strftime("%Y%m%d")}.xlsx')
        elif tipo == 'calipolis':
            wb = crear_reporte_calipolis_excel()
            from io import BytesIO
            from datetime import datetime
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            filename = f"Calipolis_Consolidado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            result = (output, filename)
        else:
            result = exportar_excel(tipo)
        if not result:
            return jsonify({"error": "Error generando reporte"}), 500
        
        output, filename = result
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500
