---
name: yve01
description: >
  Full context skill for Yve.01 — an AI-first hotel operations automation startup
  targeting the hospitality finance sector. Use this skill whenever the user mentions
  Yve.01, hotel finance automation, AP/AR workflows, the Hilton project, Oracle hotel
  integration, OTA commissions, albaranes, DRR, lector_drr, or anything related to
  building this startup. Also trigger for any coding task related to the project
  (invoice processing, OCR, 3-way matching, Flask dashboards, Oracle API, Trial Balance).
  This skill gives Claude complete context so the user never has to re-explain the project.
---

# Yve.01 — Contexto completo del proyecto (actualizado julio 2026)

## Qué es Yve.01

Startup de automatización financiera para hoteles (AR, AP, conciliación bancaria, DRR,
F&B cost, multi-hotel). Fundador: 16 años, Barcelona, catalán nativo, SIN experiencia
programando — Claude es su único desarrollador. Validado con el Assistant Financial
Controller del Hilton Barcelona (asesor, convive con el fundador).

- **Producción:** https://yve01.onrender.com (Render free tier, auto-deploy desde main)
- **GitHub:** https://github.com/Jordi-Solano/yve01
- **Login:** admin/admin123 · Demo sin login: /demo
- **Token GitHub (push):** <TOKEN_GITHUB — está en el prompt de sesión del proyecto y en el skill local, NO escribirlo en el repo (GitHub bloquea el push)>
- **Email:** Brevo, sender vvertex001@gmail.com, API key en Render como BREVO_API_KEY. PROBADO Y FUNCIONA.

## Cómo trabajar con el usuario (MUY IMPORTANTE)

- Mensajes CORTOS y en lenguaje SIMPLE: es estudiante, no del sector. Formato que le gusta:
  resumen directo → explicación llana → qué queda por hacer.
- Idioma de la conversación: español (usa la app en catalán).
- **Claude verifica TODO por sí mismo en el navegador del usuario** (extensión Chrome) sin
  pedirle nada: push → esperar deploy (~2-4 min) → comprobar marcador en el HTML servido →
  probar funcionalmente → arreglar si falla. El usuario espera este flujo siempre.
- Si detectas mejoras o bugs, arréglalos sin preguntar. Él confía en tu criterio técnico.
- Decisiones ya tomadas: pagará Render/persistencia cuando el producto esté listo para
  vender (no antes); las cuentas multi-usuario se harán con el primer cliente; el email ya
  está probado; no quiere botón Atajos ni Actualizar en la nav.

## Setup rápido (sandbox nuevo)

```bash
cd /tmp && git clone https://<TOKEN_GITHUB — está en el prompt de sesión del proyecto y en el skill local, NO escribirlo en el repo (GitHub bloquea el push)>@github.com/Jordi-Solano/yve01.git yve01b && cd yve01b
git config user.email "barnar749@gmail.com" && git config user.name "Jordi-Solano"
```
Push: `git push https://ghp_...@github.com/Jordi-Solano/yve01.git main`

## REGLAS TÉCNICAS CRÍTICAS (aprendidas a base de incidentes)

1. **ESCRITURA SEGURA DE ARCHIVOS**: NUNCA `open('f','w')` directo en scripts de edición.
   Un script que falló tras abrir en 'w' TRUNCÓ dashboard.py y se subió vacío a GitHub
   (Render rechazó el build y salvó la web). Patrón obligatorio:
   ```python
   tmp = 'archivo.tmp'; open(tmp,'w',encoding='utf-8').write(src)
   assert os.path.getsize(tmp) > tamaño_esperado; os.replace(tmp, 'archivo')
   ```
   Y TODOS los `assert old in src` ANTES de cualquier replace/write. Ojo con el bug de
   paréntesis `src.replace(a, b), 1` → tupla (pasó dos veces).
2. **Validar SIEMPRE antes de commit**: `python3 -m py_compile *.py` cambiados + extraer
   los <script> de dashboard.py a un .js y `node --check`.
3. **Render free tier**: el disco se BORRA en cada deploy Y en reinicios aleatorios.
   Todo dato subido es efímero. Los deploys tardan 2-4 min; verificar con un marcador:
   `fetch('/',{cache:'no-store'})` y buscar un string nuevo del HTML/JS (¡no sirve para
   cambios solo-servidor: usar el comportamiento del endpoint!). Los builds rotos se
   descartan y la web sigue con la versión anterior.
4. **CUIDADO al borrar código**: quedaron 4 huérfanos de la limpieza de Calipolis que
   rompían cosas en runtime (bnMap en switchTab, _asyncTabs del tour, _calSparkline del
   multi-hotel, applyMobileLite en cambiarIdioma, paso fantasma en _tourSteps). Antes de
   borrar una función: grep TODAS sus referencias. Tras borrar: node --check + prueba real.
5. **CSRF**: todos los POST JSON usan `_postJson()` o header X-CSRF-Token (multipart exento).
6. **SSE**: `yield 'data: msg\n\n'`, keep-alive ': ping', lotes de 4 archivos (timeout Render).
   El endpoint batch espera `?archivos=<JSON array URL-encoded>`, NO lista con comas.
7. **requestAnimationFrame se congela** si la ventana está tapada/minimizada → para
   posicionamiento usar setTimeout(fn,0). rAF solo para drags activos.
8. **pandas dtype**: columnas vacías se tipan float64 y rechazan strings → `astype(object)`
   antes de asignar texto (pasó en asignar_manual de conciliación).
9. Los mensajes del usuario a veces mezclan bugs suyos con bugs reales: reproducir SIEMPRE
   antes de arreglar.

## Sistema i18n (7 idiomas: es + en/ca/fr/de/it/pt)

- `t(key, fallback)` en dashboard.py: JSON (static/i18n/<lang>.json, claves tipo 'fb.x') →
  fallback → clave. NUNCA devuelve null.
- `_i18nStrMap` (~200+ strings × 6 idiomas): mapa texto-español-exacto → traducción, aplicado
  por un walker de text-nodes (`_applyStrMap`) + placeholders. Un **MutationObserver**
  retraduce automáticamente todo contenido nuevo del DOM (debounce 100ms + segunda pasada).
- `_tSSE(txt)`: traducción por FRAGMENTOS (74+ patrones) para mensajes SSE dinámicos con
  variables. Cableada en los 4 onmessage, _log, títulos de overlay.
- **Para añadir traducciones por script**: insertar tras el anclaje `"\n  en: {"` (y ca/fr/
  de/it/pt) dentro de _i18nStrMap. Los duplicados no rompen (gana el último del literal).
- Idioma en localStorage `yve_lang`; color/paleta en yve_accent/yve_bg (persisten en el
  navegador; explicar al usuario que los DATOS no persisten porque son del servidor).
- La página /conciliacion tiene su propio mini-i18n (_L + tt()) en app_conciliacion.py.
- Reloj del header: _CLOCK_LOCALES según yve_lang.

## Flujo "Procesar Archivos" (EL corazón del producto)

- Botón ⚡ → modal upload-modal: drag&drop + Seleccionar archivos (multiple). En móvil NO
  hay "Seleccionar carpeta" y el input acepta image/* → el móvil ofrece cámara/galería.
- Fotos: `_comprimirImagen` (canvas máx 1800px JPEG 85%, HEIC pasa tal cual) →
  /api/scan_documento (Claude Vision) con 3 reintentos y progreso en el overlay.
  Documentos: /api/upload_facturas → /api/procesar_batch_stream (SSE, lotes de 4).
- Clasificador IA (Claude Sonnet): FACTURA, BEO, TM, CONTRATO, EXTRACTO_BANCO, VENTAS_POS,
  INVENTARIO, MERMAS, COMISIONES_OTA, ROOMING → integra en el tab correcto.
- Columnas flexibles: `_normalize_cols(df, _INV_COL_MAP/_MER_COL_MAP/_VEN_COL_MAP/_BANK_COL_MAP)`.
- lector_ota.py es 100% REGEX (sin Claude): acepta moneda delante Y detrás ("28.333,33 EUR"),
  "Nº:", "Fecha:" a secas. Si un campo sale vacío, añadir patrón ahí.
- 3-way matching BEO↔Contrato↔Factura vía eventos_referencia.json (alerta si >5% desvío).
- TESTEADO E2E (los 6 tipos llegan a su tab con cifras exactas y la conciliación cruza
  factura↔banco). Receta del test: generar CSVs inline en el navegador + PDFs por base64,
  subir, procesar por SSE, verificar APIs de cada tab.

## Conciliación bancaria

- Motor en conciliacion_bancaria.py: `cargar_extracto` (normaliza columnas, deriva
  CARGO/ABONO del signo), `conciliar` con scoring: nº factura en concepto/referencia >
  nombre proveedor > solo importe (abs, tolerancia 2%, fecha laxa). Ref con importe
  distinto → estado DIFERENCIA con desvío.
- /api/conciliar (dashboard) usa el módulo directamente; facturas desde
  facturas-procesadas/facturas_ap_*.xlsx + AR. Página detalle /conciliacion/ (traducida,
  asignación manual arreglada con astype(object)).

## Tour guiado

- 9 pasos en `_tourSteps` (títulos/textos en strmap 6 idiomas). switchTab de cada paso con
  fallback de guiones (`tab-ar_real` → `tab-ar-real`); la pestaña activa se ilumina
  (data-tour-active + zIndex); paso Notificaciones enfoca #panel-notif entero.
- Burbuja: misma caja siempre (solo cambia contenido con fade), se DESLIZA entre pasos,
  arrastrable (transform+rAF) con imán a 8 zonas + centro al soltar; posición de cada paso
  ALEATORIA entre las 8 zonas (nunca centro, sin repetir, nunca tapa pestañas; si pisa la
  zona iluminada se desliza encima/debajo manteniendo carril). `_asyncTabs` es LOCAL de
  _showTourStep — no borrar.
- Final: confeti (160 piezas, colores de la paleta via _paletteColors) + onda tourRing +
  tarjeta con var(--acc). Todo el tour respeta la paleta personalizada.

## Demo Mode personalizado (para enseñar a clientes/gestorías)

- 🎭 → modal demo-setup-modal: una línea por cadena, formato "Cadena: Hotel 1, Hotel 2"
  (línea sin ':' = hotel suelto). POST /api/demo/generar {cadenas:[{nombre,hoteles}]}.
- demo_generator.py `generar_demo()`: KPIs 6 meses/hotel con estacionalidad (grupo=cadena,
  ciudad, categoría), 13+ facturas OTA (con discrepancias y DI pendientes), 8 AP con
  matching variado, extracto donde el 50% casa con AP, recetas/inventario/ventas 30 días/
  mermas coherentes, 5 clientes corporativos con aging. Seed = hash de los nombres.
- Salir (✕ banner → /api/demo/toggle) llama `limpiar_demo()` → todo vacío.
- Multi-Hotel: overview muestra TODOS los hoteles por defecto, respuesta incluye `grupos`,
  chips de filtro por cadena (renderMHGrupos/filtrarMHGrupo, ?grupo= en overview y
  rankings), cadena visible bajo el nombre del hotel.

## MULTI-TENANT (cuentas por cliente con datos aislados) — HECHO

- `tenant_dirs.py`: cada tenant tiene su árbol en `tenants/<slug>/` (datos-referencia,
  reportes, facturas-entrada/procesadas, aprobaciones), autocreado con seeds vacíos la
  primera vez. Tenant `default` = directorios raíz (compatibilidad total con admin).
  Resolución del tenant: sesión Flask → env `YVE_TENANT` (para subprocess) → default.
- Usuarios (usuarios.json GLOBAL en raíz) con campo `tenant`. Al hacer login,
  `session['tenant_id']` = tenant del usuario. `crear_usuario(..., tenant=)` y el panel
  admin (/admin/api/crear_usuario) aceptan tenant. El SIGNUP crea su propio tenant a
  partir del nombre del hotel/grupo (slug) y registra el hotel en SU hoteles.json.
- Módulos redirigidos a rutas por-tenant: dashboard (helpers `_ddir/_rdir/_edir/_pdir/
  _adir` + `_env_tenant()` para los subprocess del pipeline), F&B (caché `_FB_CACHE`
  con clave tenant|fname), AR Real, Multi-Hotel, conciliación (módulo y página),
  notificaciones (config e historial), demo generator (¡el demo es POR TENANT!),
  exportadores, verificador_comisiones (lee YVE_TENANT al correr por subprocess).
  Patrón usado en módulos: clase wrapper `_TDir/_TFile` con `__truediv__/__fspath__/
  read_text/write_text` que evalúa la ruta EN CADA USO (no al importar).
- PROBADO en producción: admin genera demo (7 facturas "Hotel Admin Demo") → cliente1
  (tenant hotel-prueba) entra y ve 0 → genera su demo y ve 7 propias ("Hotel Prueba")
  → admin vuelve y sigue viendo solo las suyas. Aislamiento total en ambas direcciones.
- OJO: los usuarios creados en runtime y los tenants viven en disco → se BORRAN en cada
  deploy/reinicio del free tier. Con la persistencia de pago esto queda resuelto.
- GOTCHA del sweep: el replace por tokens pilló una clave string "REPORTES_DIR" en el
  dict de /api/debug — al barrer constantes, revisar strings/comentarios.

## HOTEL ACTIVO (filtro por hotel dentro de un tenant) — HECHO

- Clic en una tarjeta de hotel del Multi-Hotel → toda la app muestra solo ese hotel y
  salta al tab AR. Selector 🏨 en la nav (visible con 2+ hoteles) para cambiar o volver
  a "🌍 Todos los hoteles". `session['hotel_activo']` + `_filtrar_hotel_activo(df)`
  aplicado en cargar_datos (AR), cargar_datos_ap (AP) y _get_reservas (AR Real) cuando
  el df tiene columna hotel/nombre_hotel. Banco/F&B/DRR son de grupo (sus datos no
  llevan hotel aún). Endpoints: GET/POST /api/hotel_activo.

## UI / Nav / Móvil

- Nav: fecha (pill), usuario (pill), botón ⚡ Procesar (estilo sutil con borde --acc, SIN
  brillo pulsante), ⋯ menú. NO existen ya: Atajos, Actualizar, botón 📸, toggle Vista lite.
- Móvil (≤768px): una sola vista (media queries; el modo lite se eliminó ENTERO), grids JS
  con clases fb-kpi-grid/fb-chart-grid, tablas con scroll horizontal.
- Estados vacíos: helper `_emptyState(emoji, titulo, sub, conCta)` con CTA a Procesar —
  usado en F&B (4 subtabs), Multi-Hotel, AR Real.
- Alerts nativos prohibidos → showNotification (los alert() bloqueaban el hilo).
- Push: al activar el canal pide Notification.requestPermission + notificación de prueba
  (service worker de la PWA en móvil); si deniegan, el toggle se revierte.
- Paleta personalizable: --acc/--acc2/--acc3/--acc-dark + --acc-r/g/b (usar SIEMPRE estas
  vars en nuevo UI; el logo-dot, tour, confeti ya la siguen).

## Verificación en producción (métodos que funcionan)

- Deploy listo: fetch '/' sin caché y buscar string nuevo. Para cambios server-only,
  probar el endpoint (p.ej. respuesta con campo nuevo).
- Vista móvil: iframe de 390px inyectado en la página (las media queries aplican al iframe);
  el resize_window real no cambia el viewport por el zoom del usuario.
- Fotos de prueba: dibujar factura en canvas → toBlob → FormData a /api/scan_documento.
- Tests largos en el navegador: lanzarlos en background guardando en window.__x y leer
  después (los js_exec >45s dan timeout). Los diálogos alert() congelan la evaluación.
- La consola del panel Multi-Hotel loguea "Sin datos KPI" solo si es error real (silenciado
  el caso vacío).

## Estado del negocio (fases)

- ✅ Fase 0-3 completas: AR, AP (3-way), Oracle (SIMULACIÓN — falta credenciales reales de
  un hotel), lector DRR (probado con .xlsm real del Hilton: 45 hojas, 31 días, 7.397 líneas
  Trial Balance). Extras: clasificador universal, escáner, conciliación, F&B, 7 idiomas,
  móvil, demo personalizado.
- 🔜 Fase 4 primer cliente. Pendiente para VENDER (en orden): 1) persistencia (Render de
  pago o BD — el usuario pagará cuando toque; SIN esto los tenants/usuarios nuevos se
  borran en cada deploy), 2) ✅ multi-tenant HECHO Y PROBADO, 3) Oracle real (credenciales
  del piloto), 4) piloto con hotel real + testimonio, 5) dominio propio + revisar
  pricing/legal (existen en el código, sin probar).
- Pendiente menor técnico: DRR .xlsm no cubierto por el test E2E (probado antes a mano).
- Visión Multi-Hotel acordada: ✅ HECHA (ver sección HOTEL ACTIVO).

## Arquitectura (resumen de archivos clave)

- `dashboard.py` (~11.600 líneas): app principal Flask + TODO el frontend embebido (HTML/
  CSS/JS en strings). Blueprints registrados de: auth, config, admin, aprobaciones AR/AP,
  conciliación, F&B (tab_fb_dashboard, prefix /fb), AR Real (tab_ar_real), multi-hotel
  (tab_multi_hotel), self-service, exportador (tab_exportador /api/exportar/<tipo>), demo,
  reportes PDF, blog, billing, asientos (exportador_asientos), signup, about, pdf, legal.
- Módulos: lector_ota (regex), lector_facturas_ap, verificador_comisiones,
  detector_doble_imposicion, matching_ap_fb/otras, asignador_cuentas, conciliacion_bancaria,
  demo_generator, lector_drr, oracle_* (pipeline con gate APROBADA — legal en España),
  generador_emails*, notificaciones (Brevo).
- Datos: datos-referencia/*.xlsx|json (efímeros salvo lo commiteado — actualmente VACÍOS
  a propósito para testing), facturas-entrada/ (subidas), facturas-procesadas/, reportes/.
- Los .py de Calipolis siguen en el repo pero NO se importan.

## Puertos y rutas útiles

- 5001 dashboard · 5000 aprobaciones AR · 5002 aprobaciones AP (local).
- APIs de verificación rápida: /api/stats (AR), /api/stats_ap, /api/stats_banco,
  /fb/api/resultados|inventario|mermas|recetas, /api/ar_real/clientes,
  /api/multi_hotel/overview, /api/historial_procesado, /api/archivos_estado.
