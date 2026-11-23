# Conciliador de Facturación Telefónica - PRODAC

Sistema web corporativo para la conciliación mensual de facturación telefónica y distribución por Centro de Costo (CC Dpto) para **PRODAC** y **PRODIMIN**.

## 🚀 Módulos Disponibles

- **Módulo Claro (`claro.html` / `index.html`)**: Tres archivos obligatorios: Excel de Claro, Inventario General y PDF mensual. Conserva Factura Detallada y genera Hoja2 de trabajo y Hoja1 final con prorrateo y tablas dinámicas nativas.
- **Módulo Movistar (`movistar.html`)**: Conciliación independiente desde el recibo PDF y el Inventario General. Fact, Traf, Cargos y su contraseña son opcionales. Genera las tres hojas requeridas con fórmulas, valores calculados y una tabla dinámica real por CC.
- **Módulo Entel (`entel.html`)**: Módulo para la conciliación mensual dinámica de Entel Perú. Detecta automáticamente la hoja y columnas de facturación de cada mes, cruza con el Inventario General filtrando operador Entel, extrae el concepto PLAN del PDF oficial (sin confundirlo con equipos ni total a pagar), alerta discrepancias respecto a la Hoja1 previa, conserva al 100% las hojas originales del proveedor y añade la nueva hoja `Resultado Entel` en la primera posición con las columnas NUMERO, MONTO y CC, y una tabla dinámica nativa de Excel en G4:H.

---

## Navegación compartida

Claro (incluido `index.html`), Movistar y Entel montan la misma plantilla `js/sidebar.js` en un contenedor `.sidebar` con `data-operator`. Todos sus tamaños, márgenes, iconos y estados se definen en `css/styles.css`; el operador activo sólo cambia colores. El ancho fijo y el margen del contenido comparten `--sidebar-width: 240px`. En ventanas bajas el panel se desplaza de forma independiente.

El menú contiene Dashboard, Cargar archivos, Resultados, Sin CC - Revisar y Descargar reporte, seguido de las tres operadoras. Cada módulo registra sus destinos con `AppSidebar.bind`; los destinos ocultos y las descargas pendientes no se activan. En pantallas de hasta 760 px se abre el mismo panel de 240 px mediante el botón de menú, sin desplazar el contenido.

`python -m unittest test_operator_layout -v` compara geometría, capturas y píxeles (neutralizando únicamente los colores de operador) en cinco tamaños de ventana, verifica el desplazamiento independiente y prueba los cinco enlaces antes y después de procesar. Guarda las capturas y `comparacion.html`, con superposición opcional, en `data/sidebar-comparison/`.

### Plantilla y componentes de las pantallas

`build_views.py` es la única plantilla de encabezado, carga, botones, estado, indicadores y tablas de los tres módulos. Sus configuraciones conservan los archivos e indicadores de cada operador. Ejecutar `python build_views.py` al editar la plantilla; las páginas generadas se sirven como HTML estático y Entel mantiene su ejecución local, incluida la apertura mediante `file://`.

`css/components.css` define todos los estilos y las rejillas de 4, 2 y 1 tarjetas. `js/moduleUI.js` presenta el periodo como Mes Año sin modificar la fecha del motor, y renderiza estados, filas y observaciones. Las incidencias se muestran junto al resumen en un único recuadro Sin CC - Revisar; el estado financiero y los datos descargados conservan su significado. REVISAR señala incidencias aunque la diferencia financiera sea cero.

`python -m unittest test_module_ui -v` verifica la plantilla, geometría y estados comunes y utiliza los documentos reales disponibles de Claro, Movistar y Entel. Guarda las seis capturas iniciales/resultados, sus variantes de laptop y pantalla pequeña, una galería `comparacion.html` y `verification.json` en `data/unified-ui/`. Si existe una copia anterior del Excel en `baseline/`, compara todos los recursos del paquete XLSX mediante XML canónico (prefijos equivalentes); únicamente excluye los metadatos de fecha de `docProps/core.xml`.

## 🛠️ Requisitos e Instalación

### Carga y diagnóstico de rendimiento

Los tres módulos usan `js/uploadUI.js`: al seleccionar sólo se validan nombre, extensión y tamaño (hasta 100 MB por archivo), se guarda el objeto File y se actualiza la tarjeta. No se leen ni envían documentos antes de pulsar Procesar archivos. El botón se habilita al completar las entradas obligatorias; durante el proceso se bloquean temporalmente los selectores y se restauran incluso ante errores. Limpiar no registra nuevos listeners.

Entel carga SheetJS, PDF.js y JSZip únicamente al procesar. En HTTP/localhost usa un Worker para leer e indexar ambos Excel; transfiere los buffers y recupera el original para generar la salida, sin releer archivos. Si el navegador no permite ese Worker (por ejemplo, apertura directa `file://`), utiliza las mismas funciones y cede el control entre lecturas. PDF.js extrae texto sin renderizar páginas y se libera cada página; el cruce cede control por lotes. Claro y Movistar procesan en el servidor.

`UploadDiagnostics.snapshot()` permite inspeccionar las mediciones del navegador: selección, lectura de PDF, lectura de Excel, cruce con inventario, conciliación, creación de Excel y tiempo total. Las respuestas de Claro y Movistar incluyen `timings` con las fases medidas en Python. Los registros son acotados y no incluyen contenido de documentos ni contraseñas. `python -m unittest test_upload_performance -v` guarda mediciones y comprobaciones en `data/upload-performance/` (ignorado por Git), incluyendo ausencia de lecturas/peticiones antes del botón, listeners únicos, recuperación tras error, uso del Worker y lectura única de cada File de Entel.

### Requisitos del Sistema
- **Claro y Movistar:** servidor Python con las dependencias indicadas a continuación. No requieren Microsoft Excel instalado en el servidor.
- **Entel:** navegador moderno y acceso a Internet para cargar las bibliotecas CDN. No requiere Python para procesar o descargar.
- Navegador web moderno (Chrome, Edge, Firefox).

### Instalación de dependencias
Ejecutar en la terminal o consola de comandos (CMD / PowerShell):

```bash
python -m pip install -r requirements.txt
```

Movistar utiliza las dependencias existentes: `pypdf`, `openpyxl` y `msoffcrypto-tool`. No requiere nuevas bibliotecas de ejecución ni Excel instalado en el servidor.

---

## ▶️ Cómo Iniciar el Servidor (Claro / Movistar / API)

1. Hacer doble clic en el archivo:
   ```text
   iniciar_servidor.bat
   ```
2. El script verificará automáticamente las dependencias e iniciará el servidor local en:
   ```text
   http://localhost:8080/movistar.html
   ```
3. O de forma manual ejecutando:
   ```bash
   python server.py
   ```

---

## 📁 Archivos de Entrada Requeridos

### Claro

Abra `http://localhost:8080/claro.html` después de iniciar el servidor. La pantalla conserva los selectores **PRODAC / PRODIMIN**, tres tarjetas numeradas y el dashboard de conciliación.

Los tres archivos son obligatorios:

1. **Excel de Claro XLSX**, con `Factura Detallada`: fuente de teléfonos, planes e importes de `Total Cargos por Línea (Sin IGV)`. Se conservan los registros repetidos y las filas Varios.
2. **Inventario General XLSX**: hoja de la empresa seleccionada, `Numero` en C y `CC Dpto` en H. El cruce compara los últimos nueve dígitos. Los números sin CC permanecen como `SIN CC - REVISAR`.
3. **Recibo mensual PDF**: valida la empresa seleccionada, obtiene periodo y Cargos del Mes, y comprueba los ajustes generales. Los conceptos no afectos al IGV no forman parte del reparto. Una discrepancia entre Varios del PDF y del Excel requiere corregir la entrada; no se altera el original para ocultarla.

La descarga es una copia del Excel cargado con exactamente tres hojas y este orden:

- **Factura Detallada**: se conserva el XML original y sus recursos, valores y formatos. No se reescribe con openpyxl. Se mantienen las filas Varios tal como fueron cargadas.
- **Hoja2**: NUMERO, MONTO y CC desde A1, antes del ajuste; incluye Varios cuando existe. Su pivote nativo en G7 usa la tabla estructurada `ClaroTrabajo`. I:J muestran `prorrateo` y `total ajustado` por CC, con fórmulas y valores precalculados.
- **Hoja1**: NUMERO, MONTO y CC desde A1, con importes finales. El pivote nativo en G7 usa `ClaroFinal`; no incluye fila Varios ni las columnas de prorrateo.

En ambas hojas A:C tiene texto negro, fondo blanco, sin bordes decorativos ni filtros visibles; se conserva el origen estructurado ampliable del pivote. Los importes tienen dos decimales y los negativos se muestran en negro. Sólo G:H en los encabezados del pivote y en Total general lleva el estilo azul claro; I:J tiene formato sencillo.

Varios se distribuye proporcionalmente por CC y después por teléfono dentro de cada CC. Cada importe se redondea a dos decimales y el residuo se aplica al último CC y a la última línea del CC. No se asigna un CC nuevo a ningún teléfono. Si faltan CC, las líneas permanecen con su importe original y el reparto se calcula sobre los CC identificados, con una observación visible. Un denominador cero impide repartir un ajuste distinto de cero.

El dashboard mantiene conciliación, líneas, registros, importes antes y después del ajuste, Varios, Cargos del Mes, diferencia, resumen por CC y detalle. Una diferencia absoluta de hasta S/0.01 muestra CUADRA. Los CC faltantes se informan por separado.

El bloque completo de resultados comienza oculto y sólo aparece al terminar correctamente el procesamiento. Limpiar, cambiar de empresa o modificar un archivo invalida y oculta el resultado anterior y restablece el periodo; las respuestas pendientes de una entrada anterior se descartan. Durante el proceso se muestra el indicador en el botón. Los errores aparecen junto a la carga de archivos sin mostrar resultados vacíos. Después de procesar, el recuadro Sin CC - Revisar muestra las incidencias o el mensaje común de que todas las líneas están conciliadas y tienen centro de costo.

Ambas tablas dinámicas tienen cachés independientes y actualizan su origen estructurado al abrir normalmente Excel o mediante **Datos → Actualizar todo**, incluyendo filas nuevas y cambios en MONTO o CC. Las modificaciones de cada tabla se reflejan en su pivote al actualizar; para regenerar el reparto completo entre ambas hojas, procese de nuevo los tres archivos. Las pruebas verifican directamente en Microsoft Excel apertura sin reparación, modificaciones, filas nuevas negativas y actualización al reabrir.

Pruebas: `python -m unittest test_claro test_claro_browser -v`. En Windows con Excel: `powershell -File test_claro_excel.ps1 -WorkbookPath ruta.xlsx -SheetName Hoja1`, y repetir con `Hoja2`. Las evidencias quedan en `data/claro-results/`, excluido de Git.

Las pruebas reales usan `CLARO_JULY_PDF`, `CLARO_JULY_REFERENCE`, `CLARO_OCTOBER_PDF`, `CLARO_OCTOBER_REFERENCE` y `CLARO_INVENTORY`. El Excel de cada mes es obligatorio. Sin el PDF de octubre sólo se verifica su Excel: esa comprobación no sustituye la aceptación completa de los tres documentos.

### Movistar

Movistar solicita solamente **dos archivos obligatorios**:

1. **Recibo PDF Movistar**, con texto seleccionable.
2. **Inventario Celulares General.xlsx**, con hoja `PRODAC` y campos `Numero` y `CC Dpto`.

Pulse **Procesar archivos**, revise los totales y **Sin CC - Revisar**, y pulse **Descargar Excel conciliado**. La descarga se llama `Movistar_Conciliado_Mes_Año.xlsx`, usando el periodo del PDF.

El desplegable **Validación opcional** admite Fact, Traf, Cargos y la contraseña mensual de esos Excel. Puede cargar uno o varios. Se comparan teléfonos y montos (Fact), teléfonos y planes (Traf), y cuotas por teléfono e importe (Cargos). Las diferencias no reemplazan las cifras oficiales ni bloquean la descarga del procesamiento PDF. La contraseña se elimina del campo al enviar y los Excel se descifran sólo en RAM. No se registra ni almacena la contraseña.

No se solicita el reporte de referencia, el mes anterior, listas maestras ni archivos adicionales. Al cambiar un archivo se invalidan los resultados anteriores para evitar descargar una conciliación desactualizada.

### 🏷️ Extracción Dinámica de Descuentos desde el PDF
- No se requiere ninguna lista maestra ni configuración previa de descuentos.
- `pypdf` extrae el texto de todas las páginas. Se unen antes de separar los Anexos y se ignoran encabezados repetidos; una descripción o un Anexo pueden continuar en la página siguiente.
- Cada Anexo aporta teléfono, plan, cargos mensuales, adicionales afectos y descuentos. Se normalizan tildes, mayúsculas, espacios y duraciones. Los descuentos de 12 meses, permanencia de 12 meses y 3 meses se exportan positivos a DESC. Los de 18 meses y otras bonificaciones netas se incorporan en MONTO. `MONTO F = MONTO - DESC`.
- Del bloque detallado `Cargos Adicionales Inafectos` se extraen equipos y financiamientos de venta, incluido cada registro de S/1.00. Se lee el teléfono después de TELEF y se conserva todo el importe.
- Se usa aritmética decimal y el redondeo neto del resumen; si éste falta, se combinan el mes actual y anterior sin duplicar las repeticiones del PDF. La diferencia conserva su signo.
- Se valida subtotal, cargos inafectos, total afecto y total facturado con tolerancia S/0.01. Un registro faltante, duplicados del PDF o una discrepancia impiden mostrar CUADRA, aunque el total general por sí solo coincida.
- El inventario se cruza por los últimos nueve dígitos. Los teléfonos ausentes o con CC ambiguo permanecen como SIN CC, conservan su importe y aparecen en la revisión web.

---

## 📱 Módulo Entel (`entel.html`)

### Cómo abrir Entel sin pasos adicionales
Abra su aplicación como siempre y seleccione **Entel** en el menú. Entel admite alojamiento estático (incluido GitHub Pages o Live Server). También puede abrir directamente **entel.html** con doble clic: sus scripts no usan módulos ES ni requieren una API.

Cargue los tres archivos y pulse **Procesar archivos**; después, **Descargar Excel conciliado**. No ejecute **python server.py** para Entel. Se necesita conexión a Internet para cargar SheetJS 0.20.1, PDF.js 3.11.174 y JSZip 3.10.1 desde sus CDN. No hay instalaciones adicionales para el usuario de Entel.

La descarga mantiene el nombre del proveedor con el sufijo ** - Conciliado.xlsx**. El procesamiento, cruce y descarga ocurren localmente: no se envían los documentos al backend.

### Arquitectura y conservación
- Claro carga **js/app.js**, envía los tres archivos y la empresa seleccionada a `/api/claro/process` y descarga el libro generado por `claro_workbook.py`. Los parsers JS antiguos no participan en ese flujo.
- Entel antes hacía un POST a **/api/entel/process** y descargaba con un token del servidor Python. Un servidor estático no ofrece esa ruta; abrir el HTML tampoco inicia Python. Eso explicaba los fallos de conexión.
- Ahora **js/entelProcessor.js** detecta encabezados de varias filas, hojas y columnas, suma en céntimos y valida PLAN con PDF.js; **js/entelWorkbook.js** añade una hoja OOXML mediante JSZip al XLSX original.
- Se preservan las entradas originales de hojas, dibujos, fórmulas, tablas dinámicas y relaciones. Sólo se actualizan los metadatos del libro, relaciones del libro, tipos de contenido y tabla de estilos para registrar el resultado. Los índices de nombres definidos locales se ajustan al insertar la primera hoja.
- Para conservar intactos los XML de las hojas del proveedor y evitar pestañas agrupadas, se mantiene la pestaña activa original. Resultado Entel es la primera pestaña: selecciónela para ver el formato nuevo.
- Si ya existe Resultado Entel, se conserva con un nombre Resultado Entel anterior N y se añade el resultado nuevo. El Excel debe ser .xlsx sin contraseña; el PDF debe tener texto seleccionable. Estas son las mismas clases de entrada admitidas por el motor Python existente.
- **entel_engine.py** y sus rutas en **server.py** permanecen disponibles para integraciones existentes. La interfaz de Entel ya no los necesita. Movistar y Claro no se modificaron durante esta corrección.

### Archivos Requeridos:
1. **Excel original entregado por Entel**: Documento mensual con detalle de líneas (e.g. `Entel - May.xlsx`).
2. **Inventario Celulares General** (`Inventario Celulares General.xlsx`): Archivo con hoja `PRODAC` o `PRODIMIN`.
3. **Recibo oficial en PDF de Entel**: Recibo oficial mensual del cual se extrae el concepto `PLAN` (independiente de cuotas de equipos y saldo a pagar).

### Características y Reglas Implementadas:
- **Detección inteligente de Excel**: Inspecciona dinámicamente las hojas, detecta la hoja de detalle de líneas (e.g. `5.16583.00.00.100000`), reconoce encabezados (Número, Plan, Cargo Fijo) y rangos de datos sin posiciones fijas.
- **Cruce con Inventario**: Busca por número normalizado de 9 dígitos, filtra operador `Entel`, y recupera Empresa, Usuario, Área, Departamento, CC Área y CC Dpto. Asigna `CC Dpto` como Centro de Costo oficial.
- **Comparación con Hoja1 previa**: Si el archivo de entrada incluye una `Hoja1` manual de meses anteriores, la preserva intacta y compara sus centros de costo con el inventario actual, señalando cualquier discrepancia en el detalle web y el CSV de observaciones.
- **Conservación obligatoria del Excel original**: Todas las hojas originales, fórmulas y formatos se mantienen intactos. Se inserta una nueva hoja denominada `Resultado Entel` como primera hoja del libro (índice 0).
- **Formato contable de Resultado Entel**: La fila 1 comienza directamente con NUMERO, MONTO y CC en A:C, con filtros. MONTO conserva el Cargo Fijo **con IGV** del proveedor, con dos decimales; CC proviene de CC Dpto del inventario. El total usa SUM sobre la tabla de datos. No incluye títulos ni columnas administrativas.
- **Tabla dinámica nativa**: Desde G4:H se muestran Etiquetas de fila y Suma de MONTO. Agrupa por CC, suma MONTO y presenta el Total general en celeste y negrita. Tiene definición OOXML, caché, registros y relaciones de una PivotTable real; no es una tabla de fórmulas que simule un pivote.
- **Actualizar en Excel**: Seleccionar la primera pestaña Resultado Entel y usar clic derecho en el resumen → Actualizar (o Datos → Actualizar todo). Si inserta filas dentro de la tabla principal, el origen estructurado incluye las líneas y CC nuevos al actualizar. La tabla dinámica no se recalcula automáticamente en cada edición; debe actualizarse.
- **Registros sin CC**: Se exportan como SIN CC y sus montos permanecen en el detalle, el resumen y el total. Empresa, usuario, área, departamento y observaciones siguen disponibles en el módulo web y en su descarga de observaciones.
- **Conciliación estricta contra PDF**: Extrae el concepto `PLAN` del recibo PDF oficial y valida:
  `Diferencia = Suma(Cargo Fijo Excel) - Plan(PDF)`
  Tolerancia de redondeo: S/ 0.01. Muestra `CONCILIADO` o `REVISAR`.
- **Cero datos hardcodeados**: Reutilizable todos los meses sin importar nombres de archivos, importes, meses ni cantidad de líneas.

### Pruebas de navegador sin backend
Para desarrollo, instalar **playwright** con pip y disponer de Chrome (o instalar Chromium con **python -m playwright install chromium**). No es una dependencia del módulo.

Ejecutar **python test_entel_browser.py**.

La prueba usa un servidor estático temporal y repite Entel mediante file://. Comprueba conciliación, tablas, filtros, descarga y conservación del paquete original, además de archivos sintéticos con estructuras e importes distintos. Las evidencias se guardan localmente en `data/browser-results/`, excluido de Git.


### Prueba de formato y actualización en Microsoft Excel
Después de la prueba de navegador, ejecutar **powershell -File test_entel_excel.ps1** en Windows con Microsoft Excel instalado. No es un requisito de uso del módulo.

La prueba abre el archivo descargado en modo de sólo lectura, verifica que no requiera reparación, reconoce una tabla dinámica nativa y comprueba su ubicación y campos. Prueba Actualizar, un monto modificado y una línea con CC nuevo dentro de la tabla principal, sin guardar los cambios de prueba. La evidencia queda en **data/format-tests/excel-verification.json**.

El generador agrega los componentes nativos directamente al ZIP existente con JSZip; SheetJS se mantiene para leer los datos. No se añadió ninguna dependencia del usuario ni un backend. La estructura de PivotTable y su caché sigue la documentación de [Microsoft Open XML](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-pivottables).

### Pruebas Automatizadas del motor Python de Entel:
Ejecutar en terminal:
```bash
python test_entel_acceptance.py
```
Verifica detección dinámica, conciliación contra PDF, funcionamiento sin Hoja1 y manejo de errores. Los documentos privados se utilizan solamente en el entorno local.

---

## 📊 Estructura del Excel Generado

El archivo descargado contiene exactamente tres hojas y en este orden:

1. **`CARGOS ADICIONALES`**:
   - Columnas: `NUMERO`, `DESCRIPCION`, `MONTO`.
   - Encabezados en A1:C1; datos desde la fila 2. Cuotas de equipos y venta extraídas del PDF (`TELEF:`). El nombre de la hoja aparece únicamente en su pestaña.
   - Fila final con fórmula `=SUM(C2:Cultima_fila)` y filtro.

2. **`CARGO MENSUAL`**:
   - Columnas: `NUMERO`, `PLAN`, `MONTO`, `DESC`, `MONTO F`.
   - Encabezados en A1:F1: N°, NUMERO, PLAN, MONTO, DESC y MONTO F; datos desde la fila 2. Una fila por Anexo del PDF, sin título ni filas vacías sobre los encabezados.
   - `MONTO = cargos mensuales + cargos adicionales afectos - descuentos incorporados`; DESC conserva los descuentos clasificados por su descripción; `MONTO F = D{row}-E{row}`.
   - Totales de MONTO, DESC, MONTO F y diferencia de comprobación contra el subtotal del PDF.

3. **`Hoja1`**:
   - Columnas A, B, C: `NUMERO`, `MONTO` (proviene de `MONTO F`), `CC` (obtenido de `CC Dpto` del inventario).
   - Desde G4:H: tabla dinámica real, con `Etiquetas de fila`, `Suma de MONTO` y `Total general`, sin IGV ni TOTAL adicionales.
   - A:C es una tabla estructurada de Excel llamada `MovistarDatos`, con encabezados y filtros en la fila 1 y una fila de totales con `SUM(MovistarDatos[MONTO])`. El pivote usa el nombre de esa tabla como origen e incluye las filas añadidas o eliminadas al actualizar. La fila de totales queda excluida del origen del pivote.
   - `movistar_workbook.py` construye la definición nativa con las clases OOXML de openpyxl, su caché, los registros actuales y las relaciones. Activa `refreshOnLoad` y descarta del caché los CC eliminados al refrescar. Después de editar MONTO o CC, use **Datos → Actualizar todo** o clic derecho en la tabla dinámica → **Actualizar**; también se actualiza al cerrar y volver a abrir. El archivo es `.xlsx`, sin macros.
   - Las fórmulas incluyen sus resultados calculados en el archivo para vistas previas; Excel puede recalcularlas y actualizar la PivotTable. La estructura sigue [Microsoft Open XML](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-pivottables).

### Verificación de Movistar

Ejecutar `python -m unittest test_movistar_dynamic test_movistar_acceptance test_server_movistar_api test_movistar_browser -v`. Las pruebas privadas se omiten si faltan los documentos en `data/movistar_nov_test/`. Para validar los Excel protegidos, proporcionar la contraseña de prueba mediante `MOVISTAR_TEST_PASSWORD`, sin incorporarla al código. Playwright 1.64.1 y Chrome/Chromium son herramientas de desarrollo, no dependencias de ejecución.

Las pruebas de navegador descargan `data/browser-results/Movistar_Conciliado_Noviembre_2026.xlsx` y verifican sus tres hojas, fórmulas, valores, tabla estructurada y caché. En Windows con Microsoft Excel instalado, ejecutar `powershell -File test_movistar_excel.ps1` para abrir esa descarga, comprobar que no requiere reparación y recalcular las fórmulas. La prueba cambia temporalmente MONTO y CC, agrega y elimina una fila, ejecuta Actualizar todo, comprueba los saldos de cada CC y restaura el total inicial. También guarda una copia local de prueba con el origen modificado y la caché pendiente, la reabre y verifica que se refresque automáticamente. La descarga original conserva sus valores. La evidencia queda en `data/browser-results/movistar-excel-verification.json`.

La prueba de reapertura inicia una instancia propia de Excel por su apertura normal y consulta el libro ya abierto. Esto permite comprobar la actualización automática: [Microsoft documenta que `Workbooks.Open` omite ese refresco](https://learn.microsoft.com/office/vba/api/excel.pivotcache.refreshonfileopen). El script de prueba no agrega macros al archivo.

Los informes de aceptación y las evidencias de documentos privados se conservan únicamente en `data/`, excluido de Git. Las pruebas públicas utilizan exclusivamente datos sintéticos.

---

## 🔒 Seguridad y Privacidad
- Las contraseñas de archivos nunca son guardadas en disco, base de datos ni `localStorage`.
- Los archivos desencriptados se manejan como flujos de bytes en memoria RAM (`io.BytesIO`) y se liberan al finalizar la solicitud.


## Publicación y servicio de Render

Antes de publicar, ejecutar `python -m unittest discover -v`. Las referencias privadas permanecen fuera de Git y sólo se usan localmente. Para comprobar el servicio público con documentos generados en memoria, configurar `APP_PUBLIC_TEST_BASE_URL` y `EXPECTED_RELEASE_COMMIT` y ejecutar `python -m unittest test_release_public -v`. Esta prueba no lee documentos corporativos; verifica salud y commit, carga sin lecturas anticipadas, navegación en escritorio y móvil, conciliación, descarga, pivotes y limpieza. Guarda capturas y resultados en `data/release-public/`, excluido de Git.

La aplicación completa se sirve con **python -u server.py**, en **0.0.0.0** y el puerto indicado por **PORT**. El archivo **.python-version** fija Python 3.14.7 y **requirements.txt** contiene las tres dependencias de ejecución en sus versiones verificadas. Las bibliotecas de Claro y Entel se cargan en el navegador desde CDN; Movistar usa la API del mismo servicio.

El Blueprint **render.yaml** define la rama **main**, el build **pip install -r requirements.txt**, el comando de arranque y **/health**. La respuesta de salud incluye el SHA **RENDER_GIT_COMMIT** para identificar qué versión está desplegada. **Procfile** usa el mismo arranque.

Para un servicio existente en Render:
1. Abrir [Render Dashboard](https://dashboard.render.com/) y seleccionar **facturacion-telefonica-prodac**.
2. En **Settings**, verificar repositorio **jakelincullampepuma-maker/facturacion-telefonica-prodac**, rama **main**, build y start anteriores. Si se administra mediante Blueprint, sincronizar el Blueprint; un archivo render.yaml por sí solo no modifica un servicio creado manualmente.
3. En **Manual Deploy**, elegir **Deploy latest commit** si el despliegue automático no se inicia. No crear otro servicio si ya existe.
4. Comprobar que **/health** muestra el mismo commit de GitHub y abrir los módulos Claro, Movistar y Entel desde la URL del servicio.

Si no existe un servicio, crear **New → Blueprint**, seleccionar el repositorio y revisar el plan **free** antes de aplicar la configuración. Los pasos y requisitos del puerto están descritos en la [documentación de Render](https://render.com/docs/web-services).

### Datos privados y pruebas

**data/**, archivos Excel/PDF/CSV/ZIP, reportes, inventarios y archivos **.env** están excluidos de Git. El servidor sólo publica las páginas y recursos de **js/** y **css/**; no permite descargar **.git**, secretos, pruebas ni documentos locales. Las descargas de la API viven en RAM, tienen caducidad de 30 minutos y están sujetas a un límite de caché de 64 MB; los reinicios de Render invalidan los enlaces temporales. La carga máxima por petición es de 100 MB.

No se publica la contraseña Movistar. Para ejecutar las pruebas privadas se proporcionan mediante el entorno **MOVISTAR_TEST_PASSWORD**, **MOVISTAR_FACT_PATH**, **MOVISTAR_TRAF_PATH**, **MOVISTAR_CARGOS_PATH**, **MOVISTAR_PDF_PATH** y **MOVISTAR_INVENTORY_PATH**. Para Entel, las tres referencias pueden colocarse localmente en **data/entel_test/** o especificarse mediante **ENTEL_TEST_DIR** para el motor Python. Nunca incorporar estas referencias al repositorio.

La prueba portable se ejecuta con **python -m unittest discover -v** después de instalar **playwright==1.64.1** y Chromium para desarrollo. GitHub Actions ejecuta las pruebas con datos sintéticos, sin contraseñas o documentos corporativos. Las pruebas reales se omiten cuando no se proporcionan sus referencias privadas.
