# Multi-Hotel con datos reales — mapa, riesgo y plan

*Investigación, 28 julio 2026. No se ha tocado ni una línea de código.*

---

## Lo primero, porque cambia todo lo demás

**El problema no es Multi-Hotel. Es que ningún documento sabe a qué hotel pertenece.**

Cuando subes una factura, el fichero que se guarda tiene esto:

```
archivo · numero_factura · fecha · nombre_proveedor · NIF_proveedor ·
descripcion_concepto · base_imponible · porcentaje_iva · cuota_iva ·
total_factura · moneda · tipo · estado_matching
```

Comprobado sobre la factura real que hay hoy en disco: **ninguna columna de hotel**. Ni `hotel`, ni `hotel_id`, ni `propiedad`. Y lo mismo en albaranes, órdenes de compra, ventas de F&B, extracto de banco e inventario.

Multi-Hotel no puede agregar por hotel porque **no hay nada que agregar por hotel**. Todo lo demás de este documento sale de ahí.

Hay un corolario incómodo: el selector de "hotel activo" que ya existe **no filtra AP**. `_filtrar_hotel_activo` busca una columna `hotel` o `nombre_hotel`; si no la encuentra, devuelve la tabla entera sin avisar. En el modo demo sí funciona, porque el generador de demo **sí** escribe una columna `hotel`. O sea: la función parece correcta cuando la pruebas con datos de demo y no hace nada con datos reales. Es la trampa más peligrosa del modelo actual.

**Y una segunda cosa que hay que decidir antes de construir**: la mitad de lo que Multi-Hotel enseña hoy —ocupación, ADR, RevPAR, ingresos de habitaciones, GOP— **Yve no lo puede calcular**. Yve procesa facturas, albaranes, comisiones, banco y F&B. Nunca ve una habitación. Vuelvo sobre esto en el punto 2.

---

## 1 · Qué muestra Multi-Hotel hoy y de dónde sale

### La cadena completa

```
demo_generator.py:120   ← random.Random(), datos inventados
        │
        ▼
datos-referencia/kpis_hoteles.xlsx     ← HOY: 0 filas
        │
        ▼   pd.read_excel  (tab_multi_hotel.py:21)   ← NO pasa por almacen_datos
tab_multi_hotel.py  →  4 endpoints /api/multi_hotel/*
        │
        ▼   fetch  (dashboard.py:13012)
loadMultiHotel()  →  #panel-multi_hotel
```

**El único fichero que alimenta el panel es `kpis_hoteles.xlsx`, y su único escritor es el generador de demo.** No hay ningún camino, ni parcial, por el que hoy entre un dato real.

Verificado: el fichero tiene **0 filas y 24 columnas**. Por eso ahora mismo el panel no enseña ni siquiera el demo — enseña el estado vacío "No hay hoteles en el grupo".

### KPI por KPI

Todo lo del bloque consolidado es una agregación limpia de ese Excel:

| Lo que ves | De dónde sale |
|---|---|
| REVENUE MTD | `suma(total_ingresos)` |
| GOP TOTAL y GOP% medio | `suma(gop_eur)` · `media(gop_pct)` |
| OCUPACIÓN MEDIA y ADR | `media(ocupacion_pct)` · `media(adr_eur)` |
| REVPAR MEDIO | `media(revpar_eur)` |
| Habitaciones | `suma(habitaciones)` |
| Facturas AP pendientes | `suma(facturas_ap_pendientes)` |
| Gráficos de tendencia | agregación por mes de las mismas columnas |
| Tarjetas por hotel | copia directa de cada fila |

Y hay cosas que **no salen del fichero: se las inventa el código**.

- **Las estrellas** (`tab_multi_hotel.py:86`): `'5★' if 'boutique' in nombre or '5' in nombre else '4★'`. Adivina la categoría a partir del nombre, **ignorando la columna `categoria` que sí trae el Excel**. Un "Hotel 5 de Mayo" sale como 5 estrellas.
- **Las alertas** (`:163-168`) no se leen, se fabrican con umbrales escritos a mano: `AP pendientes > 5`, `días fuera de balance > 0`.
- **El estado de Oracle** (`:103`) tiene por defecto el literal `'SIMULACION'`, y en los datos de demo es `random.choice(["OK","OK","OK","PENDIENTE"])`.

### Tres cosas rotas que aparecieron de paso

1. **El botón "⬇️ Excel" descarga otro universo.** No lee el panel ni el fichero: devuelve seis hoteles escritos a mano en `exportador_reportes.py:202-209` ("Premier London Mayfair"…) con el título fijo "Junio 2025". Lo que ves y lo que te descargas no tienen nada que ver.
2. **La columna "Categoría" de la tabla sale siempre vacía.** El frontend hace `'★'.repeat(h.stars)`, pero `h.stars` es el texto `'4★'`, no un número; `repeat` de un texto da cadena vacía. Comprobado.
3. **`multi_hotel_data.py` es código muerto** — 141 líneas que nadie importa. Y tiene sus propios umbrales (`AP > 8`) que **no coinciden** con los del módulo vivo (`AP > 5`). Igual que `dashboard_multihotel.py` (496 líneas) y `renderMHMap()`. Cuando toquemos esto, hay que borrarlos: si no, el día que alguien los lea creerá que son la verdad.

---

## 2 · Qué datos reales hay, y cuáles pasan por `almacen_datos`

`almacen_datos` es la capa que ya centraliza la lectura. Hoy cubre **AP, AR y banco**, y no cubre **F&B ni DRR ni las aprobaciones**.

| Sección | Datos reales que produce | ¿Pasa por `almacen_datos`? |
|---|---|---|
| **AP** | facturas, líneas de factura, albaranes, líneas de albarán, órdenes de compra, líneas de PO | **Sí**, seis funciones |
| **AP · aprobaciones** | quién aprobó qué y cuándo | **No** — cada consumidor hace su propio cruce, y con criterios distintos |
| **AR / OTA** | facturas de OTA, verificación de comisiones, doble imposición | **Sí**, tres funciones |
| **AR · reclamaciones** | discrepancias, emails pendientes | **No** (se derivan del informe de verificación, que sí está) |
| **Banco** | extracto + informe de conciliación | **Sí** — es la sección mejor centralizada |
| **F&B** | ventas TPV, inventario, mermas, recetas, food cost | **No, nada.** Tiene su propia caché en `tab_fb_dashboard.py` |
| **DRR** | revenue diario, trial balance, días fuera de balance, ocupación/ADR/GOP | **No, nada.** Un glob suelto en `dashboard.py:3387` |

### Lo que se podría agregar HOY sin inventar nada

Estos salen ya de `almacen_datos`, para un hotel:

- número e importe de facturas AP · desglose por estado de cruce (cuadran / discrepancia / sin albarán)
- importe bruto AR y desglose por OTA · **comisión reclamable en euros** (del informe de verificación)
- certificados de doble imposición pendientes
- saldo de banco, cargos, abonos, pendiente de conciliar, movimientos parados más de 7 días
- albaranes sin facturar · compromiso en órdenes de compra abiertas

Estos necesitan una función nueva en el almacén, pero el dato existe:

- food cost %, coste de mermas, cobertura del recetario → hoy solo por el endpoint de F&B
- días fuera de balance, revenue diario → hoy solo leyendo el informe de DRR a mano

### Lo que NO existe y no se puede calcular

| Dato | Por qué no |
|---|---|
| Habitaciones disponibles | `hotel_config.json` trae `hotel_habitaciones: 0`. Sin denominador no hay ocupación ni RevPAR |
| Habitaciones ocupadas | No hay PMS ni fichero de rooming diario |
| Ingresos de habitaciones | Solo llegan si el hotel sube un DRR que ya los trae calculados |
| ADR, RevPAR | Derivados de los dos anteriores |
| **GOP** | Requiere una cuenta de resultados departamental. **Y ojo: hoy, cuando el DRR no lo trae, `dashboard.py:3470` lo estima como `ingresos × 0,22`** con el comentario "media del sector". Es un número fabricado que viaja igual que uno medido |

**Esto es una decisión de producto, no técnica.** Multi-Hotel hoy es un cuadro de mando *hotelero* (ocupación, ADR, RevPAR, GOP). Yve es un sistema *financiero*. Se puede:

- **(A)** Redefinir Multi-Hotel alrededor de lo que Yve sí sabe: dinero por pagar, dinero reclamable, incidencias, food cost, saldo. Real desde el primer día, sin depender de nada externo.
- **(B)** Mantener las métricas de habitación y aceptar que **solo aparecen si el hotel sube su DRR**, marcándolas honestamente cuando no estén.
- **(C)** Las dos: una fila financiera siempre real, y una fila hotelera que se rellena con el DRR y se dice claramente que falta si no está.

Mi recomendación es **(C)**, empezando por la parte financiera. Es lo que diferencia a Yve: el cuadro de ocupación lo tiene cualquiera; "tienes 14.000 € reclamables a Booking repartidos en tres hoteles" no lo tiene nadie.

---

## 3 · Cómo se relacionan hoy los hoteles

| Nivel | ¿Aislado? | Cómo |
|---|---|---|
| **Tenant** | **Sí**, de verdad, en disco | `tenants/<slug>/` con su árbol completo, resuelto por la sesión |
| **Hotel — AR / OTA** | Filtro visual, parcial | El nombre del hotel viene **dentro del PDF de la OTA**, y se filtra con un `contains` |
| **Hotel — AP** | **No** | La columna no existe; el filtro pasa de largo en silencio |
| **Hotel — banco / F&B / DRR** | **No** | Datos de grupo por diseño declarado |
| **Hotel — en disco** | **No** | Un único árbol por tenant, partido por **fecha**, no por hotel |

`hoteles.json` es hoy **el censo, no la llave**: guarda id, nombre, ciudad, categoría, habitaciones y grupo, y se usa para el desplegable de hotel activo, para el panel de administración y para **la firma de los emails de reclamación**. Ningún fichero de datos apunta a esos ids. Está vacío (`[]`) en este tenant.

Hay dos usuarios que ya representan un grupo: **`solmar`** (Cadena Sol: Hotel Sol Mar + Hotel Sol Playa) y **`gestoria`** (Gestoría Nord: Hotel Pirineus + Hotel Vall d'Aran). Son el caso de uso real de Multi-Hotel, y hoy sus datos irían **al mismo fichero, mezclados y sin distinguir**.

Un riesgo concreto que ya existe: el guardado de AP deduplica por nombre de fichero. Dos hoteles del mismo grupo que suban `factura_enero.pdf` el mismo día — **uno pisa al otro**.

---

## 4 · Qué haría falta para que un documento propague

Cuatro piezas, en este orden. Ninguna es enorme; la primera es la que desbloquea todo.

**a) Etiquetar en el guardado.** `_guardar_factura_ap`, `_guardar_albaran` y `_guardar_orden_compra` inyectan una columna `hotel_id` tomada del hotel activo de la sesión. Es un cambio pequeño y en un solo sitio por tipo de documento. Sin esto, nada de lo demás sirve.

**b) Que el filtro falle en cerrado, o al menos avise.** Hoy, si no hay columna de hotel, devuelve todo. Con datos de dos hoteles mezclados, eso es enseñar a un hotel las facturas del otro. Como mínimo tiene que decirlo en pantalla.

**c) Meter en el almacén lo que falta**: `aprobaciones_ap()`, y una entrada para F&B y otra para DRR. Así el agregador tiene un solo sitio de donde leer, que es justo para lo que se hizo el almacén.

**d) Un agregador nuevo** que, para cada hotel del censo, llame al almacén, filtre por `hotel_id` y devuelva la ficha. Sustituye a `kpis_hoteles.xlsx` como fuente del panel.

Respondiendo directo a tu pregunta: **es leer de `almacen_datos` y agregar** — pero solo después de (a), porque hoy no hay por dónde agrupar. Lo nuevo de verdad es el agregador, y es la parte fácil.

**La deduplicación también hay que revisarla**: hoy la clave de AP es `(numero_factura, nombre_proveedor)`. Con dos hoteles, dos facturas legítimamente distintas pueden compartir esa clave. Tiene que pasar a incluir el hotel.

---

## 5 · Riesgo: qué se puede romper y qué proteger

**Lo que NO se toca, y hay que decirlo por escrito:**

- **Oracle.** El agregador solo lee. Ningún `oracle_*`, ni el gate, ni la simulación. La comprobación de siempre antes de cada commit: `git diff --name-only | grep "^oracle_"` vacío.
- **El clasificador.** Etiquetar con el hotel pasa **después** de clasificar y leer el documento. No se toca el prompt ni el esquema que se le pide a la IA. Este es el punto donde hay que ser tozudo: si en algún momento parece que hay que meter el hotel en el prompt, la respuesta es no — el hotel lo sabe la sesión, no el papel.

**Lo que sí puede romperse:**

| Riesgo | Por qué | Cómo se protege |
|---|---|---|
| **Añadir una columna rompe un lector** | Varios módulos leen esos xlsx por su cuenta | La columna es nueva y opcional; nadie hace `SELECT *` posicional. Aun así, probar el pipeline AP entero después |
| **Los datos viejos no tienen `hotel_id`** | Todo lo ya procesado | Tratar el vacío como "sin asignar" y **enseñarlo como tal**, no repartirlo a ciegas entre hoteles |
| **Que el filtro pase a fallar en cerrado esconda datos** | Hoy fail-open; al cerrarlo, lo que no tenga hotel desaparece | Por eso el cambio (b) va **con** un aviso visible, no solo con un filtro |
| **La deduplicación fusione facturas de hoteles distintos** | Clave sin hotel | Cambiar la clave a la vez que se añade la columna, no después |
| **Sembrar de nuevo el demo pise datos reales** | `login.py` regenera el demo si `kpis_hoteles.xlsx` está vacío | Cuando el panel deje de leer ese fichero, esa siembra hay que quitarla o acotarla |
| **Que un GOP inventado se enseñe como medido** | El `× 0,22` | Marcar el dato estimado de forma que se distinga, o no enseñarlo |

**Una nota sobre el orden**: el fallo de aislamiento por hotel **ya existe hoy** y afecta a `solmar` y `gestoria`. Conectar Multi-Hotel no lo crea — lo hace visible. Eso es bueno: es mejor verlo en un panel que descubrirlo con un cliente.

---

## 6 · Plan por fases

Todo esto se puede construir y verificar **entre despliegues**, sin persistencia: se sube el juego de prueba, se comprueba, y si el disco se borra se vuelve a subir. Es como se ha verificado todo lo demás del proyecto.

### Fase 1 — Etiquetar el documento *(base de todo)*
Columna `hotel_id` en el guardado de factura, albarán y orden de compra, tomada del hotel activo. La clave de deduplicación pasa a incluirla. Nada visible todavía.

*Verificación:* dos hoteles en el censo; subo la misma factura con cada uno activo; compruebo que salen **dos filas** con `hotel_id` distinto y que antes salía una.

### Fase 2 — Que el filtro diga la verdad
El filtro por hotel activo deja de pasar de largo en silencio. Si hay hotel activo y la tabla no tiene columna, se dice en pantalla.

*Verificación:* con un hotel activo, AP enseña solo lo suyo, y lo antiguo sin etiquetar aparece como "sin asignar" en vez de colarse.

### Fase 3 — El agregador financiero *(primer valor real)*
Función nueva que, por hotel, devuelve **solo lo que Yve sabe de verdad**: facturas AP y su importe, cuántas cuadran y cuántas tienen incidencia, importe AR, **comisión reclamable**, saldo de banco y pendiente de conciliar. Multi-Hotel deja de leer `kpis_hoteles.xlsx` para esta mitad.

*Verificación:* subo facturas a dos hoteles y compruebo que el panel las suma bien y que la suma del grupo cuadra con la suma de las fichas.

### Fase 4 — Food cost y DRR al almacén
Entrada de F&B y de DRR en `almacen_datos`, y la fila hotelera del panel (ocupación, ADR, RevPAR, días fuera de balance) alimentada del DRR real **cuando lo haya**, y marcada honestamente cuando no.

*Verificación:* subo un DRR a un hotel y no al otro; el primero enseña sus métricas y el segundo dice que faltan, sin inventarlas.

### Fase 5 — Limpieza
Borrar `multi_hotel_data.py`, `dashboard_multihotel.py`, `renderMHMap()` y `openHotelDetail()`. Arreglar el Excel que descarga otro universo y la columna de categoría vacía.

### Por dónde empezar

**Fase 1, sin duda.** No porque sea lo más vistoso, sino porque es la única que desbloquea el resto — y porque **arregla un fallo que ya tienes**: hoy el filtro de hotel activo no filtra AP, y con `solmar` o `gestoria` eso significa enseñar a un hotel lo del otro.

Si quieres ver valor antes, **1 → 3** ya da un Multi-Hotel real y útil: "tres hoteles, 47.000 € de facturas por pagar, 12.000 € reclamables a las OTAs, 4 incidencias". Eso es defendible delante de un cliente. La fase 4 la haría después, porque depende de que el hotel suba su DRR y eso no lo controlas tú.

---

## Lo que hay que decidir antes de empezar

1. **¿Qué debe enseñar Multi-Hotel?** ¿La opción (A) financiera pura, la (B) hotelera, o la (C) las dos con la hotelera marcada cuando falte? Mi recomendación es (C).
2. **¿De dónde sale el `hotel_id` al guardar?** Lo natural es el hotel activo de la sesión. Pero si el usuario se olvida de cambiarlo, etiqueta mal. ¿Prefieres que el sistema **pida** el hotel al subir cuando hay más de uno?
3. **¿Qué hacemos con lo ya procesado sin hotel?** ¿"Sin asignar" y a la vista, o una pantalla para asignarlo a posteriori?
