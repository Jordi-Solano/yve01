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
                cwd="/home/claude",
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
    reportes_dir = "/home/claude/reportes"
    
    if os.path.exists(reportes_dir):
        for f in sorted(os.listdir(reportes_dir)):
            if f.startswith("ar_real_abbvie") and f.endswith(".xlsx"):
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
