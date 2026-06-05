# Integración AR Real al Dashboard

## Archivos creados

```
ar_grupos_reales.py       — Módulo principal (procesa rooming + facturas)
tab_ar_real.py            — Blueprint Flask para el endpoint API
reportes/ar_real_*.xlsx   — Reportes generados
```

## Pasos para integrar

### 1. Copiar archivos al repo

```bash
cp /home/claude/ar_grupos_reales.py C:\Users\Jo\yve01\
cp /home/claude/tab_ar_real.py C:\Users\Jo\yve01\
```

### 2. Actualizar dashboard.py

En tu `dashboard.py`, después de importar los blueprints existentes, agregar:

```python
# Al inicio, junto con otros imports
from tab_ar_real import ar_real_bp

# En la función main() o donde registras blueprints:
app.register_blueprint(ar_real_bp)
```

### 3. Agregar tab en HTML

En la sección `<div class="tabs">` del dashboard, agregar:

```html
<button class="tab-button" onclick="switchTab('ar-real')">AR Real</button>
```

Y agregar el contenedor:

```html
<div id="ar-real" class="tab-content">
    <h2>AR Real — Grupos Corporativos</h2>
    
    <div class="section">
        <h3>Procesar Facturas</h3>
        <button onclick="procesarARReal()" id="btn-ar-real" class="btn-primary">
            Procesar Archivos
        </button>
        <div id="ar-real-log" class="log-container"></div>
    </div>
    
    <div class="section">
        <h3>Reportes Generados</h3>
        <div id="ar-real-status" class="status-box"></div>
    </div>
</div>
```

### 4. JavaScript en dashboard.html

Agregar al script del dashboard:

```javascript
function procesarARReal() {
    const logDiv = document.getElementById('ar-real-log');
    logDiv.innerHTML = '';
    
    const eventSource = new EventSource('/api/procesar_ar_real');
    
    eventSource.onmessage = function(event) {
        const logLine = document.createElement('div');
        logLine.textContent = event.data;
        logDiv.appendChild(logLine);
        logDiv.scrollTop = logDiv.scrollHeight;
        
        if (event.data === 'AR_REAL_COMPLETO') {
            eventSource.close();
            cargarStatusARReal();
        }
    };
    
    eventSource.onerror = function(err) {
        console.error('Error:', err);
        eventSource.close();
    };
}

function cargarStatusARReal() {
    fetch('/api/ar_real_status')
        .then(r => r.json())
        .then(data => {
            const statusDiv = document.getElementById('ar-real-status');
            let html = `<p>Último reporte: <strong>${data.reportes[0]?.filename || 'No disponible'}</strong></p>`;
            html += `<p>Tamaño: ${data.reportes[0]?.size_kb || 0} KB</p>`;
            html += `<p>Actualizado: ${new Date(data.reportes[0]?.timestamp).toLocaleString() || '-'}</p>`;
            statusDiv.innerHTML = html;
        });
}

// Cargar status al iniciar
window.addEventListener('DOMContentLoaded', cargarStatusARReal);
```

### 5. CSS (opcional, mejorar estilo)

```css
.log-container {
    background: #f5f5f5;
    border: 1px solid #ddd;
    padding: 10px;
    max-height: 300px;
    overflow-y: auto;
    font-family: monospace;
    font-size: 12px;
    margin: 10px 0;
}

.log-container div {
    padding: 2px 0;
}

.status-box {
    background: #e8f5e9;
    border-left: 4px solid #4caf50;
    padding: 10px;
    margin: 10px 0;
}
```

## Estructura del reporte

El reporte Excel generado incluye 4 sheets:

1. **Overview** — Resumen ejecutivo (attendees, invoices, balance)
2. **Rooming List** — Detalle de cada asistente (nombre, grupo, check-in/out, forma de pago)
3. **Invoices** — Línea por línea de cada factura (qty, precio unitario, totales)
4. **Reconciliation** — Consolidado de depósitos vs consumos

## Datos de ejemplo (Hilton AbbVie)

El módulo procesa:
- **rooming.xlsx** — 67 asistentes (62 Master Account + 5 Portugal)
- **251527287_1_Abbvie_Poland_.xlsm** — Factura Poland (depósito €1,081.35, consumo €858.81)
- PDFs de facturas Portugal y Master Account (próximas iteraciones)

## Variables de entorno (opcional)

Si quieres agregar rutas personalizadas, puedes configurar:

```bash
# .env
AR_REAL_UPLOADS_PATH=/path/to/uploads
AR_REAL_OUTPUT_PATH=/home/claude/reportes
```

## Comando de test local

```bash
cd C:\Users\Jo\yve01
python ar_grupos_reales.py
```

Esto genera un archivo Excel en `/home/claude/reportes/ar_real_abbvie_YYYYMMDD_HHMMSS.xlsx`

---

**Próximos pasos:**
- [ ] Procesar PDFs de facturas (Portugal, Master Account)
- [ ] Extraer datos de BEOs (F&B charges)
- [ ] Integrar con Oracle para contabilización
- [ ] Dashboard multi-hotel para Calipolis
