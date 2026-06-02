# Yve.01

**Automatización financiera para hoteles.** Procesa facturas de OTAs y proveedores,
hace _matching_ a 3 vías, asigna cuentas del PGC, contabiliza asientos en Oracle GL
y concilia banco — con revisión humana en cada paso. Validado con el equipo financiero
de una cadena internacional 5★ en Barcelona.

Producción: https://yve01.onrender.com

---

## Arranque rápido

```bash
pip install -r requirements.txt
python dashboard.py
```

Abre **http://localhost:5001**. Toda la plataforma (login, dashboard, configuración,
administración, aprobaciones AR/AP y conciliación) corre en **un solo proceso y un solo
puerto**. No hace falta levantar varios servidores.

### Usuarios de demostración

| Usuario | Contraseña | Rol |
|---|---|---|
| `admin` | `admin123` | Administrador |
| `fc_user` | `hotel2024` | Financial Controller |
| `auditor` | `hotel2024` | Income Auditor |
| `fbmanager` | `hotel2024` | F&B Manager |

> En la pantalla de login puedes pulsar cualquier rol para rellenar las credenciales.

---

## Arquitectura

Una única app Flask (`dashboard.py`) que registra cada módulo como **blueprint**:

| Ruta | Módulo | Función |
|---|---|---|
| `/` | `dashboard.py` | Dashboard principal (AR, AP, DRR, Banco, Notificaciones) |
| `/login` · `/logout` | `login.py` | Autenticación |
| `/configuracion/` | `onboarding.py` | Alta y configuración del hotel |
| `/admin/` | `panel_admin.py` | Gestión de usuarios (solo admin) |
| `/aprobaciones-ar/` | `app_aprobacion.py` | Aprobación de facturas AR (OTAs) |
| `/aprobaciones-ap/` | `app_aprobacion_ap.py` | Aprobación de facturas AP (proveedores) |
| `/conciliacion/` | `app_conciliacion.py` | Conciliación bancaria |

La sesión y los roles se comparten entre todos los módulos. El sistema de diseño vive en
`static/yve.css` y lo cargan todas las páginas.

---

## Pipelines (línea de comandos)

```bash
# AR — Cuentas por cobrar (OTAs)
python lector_ota.py && python verificador_comisiones.py && \
python detector_doble_imposicion.py && python generador_emails.py

# AP — Cuentas por pagar (proveedores)
python lector_facturas_ap.py && python matching_ap_otras.py && \
python matching_ap_fb.py && python asignador_cuentas.py && \
python generador_emails_ap.py

# Oracle GL (modo simulación hasta tener credenciales)
python oracle_pipeline.py

# DRR — Daily Revenue Report
python lector_drr.py "ruta/al/DailyHotel.xlsm"
```

Las facturas se contabilizan en Oracle **solo** cuando tienen estado `APROBADA`
(requisito legal en España). Sin credenciales Oracle, el pipeline entra en modo
simulación automáticamente.

---

## Despliegue

`render.yaml` está configurado para Render con Gunicorn:

```bash
gunicorn dashboard:app --bind 0.0.0.0:$PORT --workers 1 --threads 8 --timeout 120
```

Variables de entorno (`.env`, nunca se versiona):

```
ANTHROPIC_API_KEY=...        # extracción de facturas con IA (con fallback regex)
SECRET_KEY=...               # firma de sesión (Render lo genera solo)
SMTP_SERVER, SMTP_PORT, SMTP_USER, SMTP_PASSWORD   # notificaciones (opcional)
# Oracle (opcional — sin esto, modo simulación):
ORACLE_BASE_URL, ORACLE_CLIENT_ID, ORACLE_CLIENT_SECRET, ORACLE_LEDGER_NAME, ORACLE_ENTITY
```

---

## Estructura

```
dashboard.py              App principal + registro de blueprints
auth.py                   Usuarios, roles y Flask-Login
login.py onboarding.py panel_admin.py
app_aprobacion*.py app_conciliacion.py     Módulos web (blueprints)
lector_*.py matching_*.py asignador_cuentas.py gestor_pos.py
verificador_comisiones.py detector_doble_imposicion.py
generador_emails*.py oracle_*.py           Lógica de negocio / pipelines
static/yve.css            Sistema de diseño compartido
datos-referencia/         Proveedores, plan de cuentas, POS, usuarios
facturas-entrada/         PDFs de entrada
facturas-procesadas/      Excels procesados
reportes/                 Salidas (verificación, matching, DRR, Oracle)
_dev_archive/             Scripts de desarrollo (no forman parte del producto)
```
