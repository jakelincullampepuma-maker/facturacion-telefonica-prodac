# Conciliador de Facturación Telefónica - PRODAC

Sistema web corporativo para la conciliación mensual de facturación telefónica y distribución por Centro de Costo (CC Dpto) para **PRODAC** y **PRODIMIN**.

## 🚀 Módulos Disponibles

- **Módulo Claro (`claro.html` / `index.html`)**: Conciliación mensual de Claro para PRODAC y PRODIMIN mediante lectura de factura detallada, cruce con inventario y validación con recibo PDF.
- **Módulo Movistar (`movistar.html`)**: Módulo independiente para el procesamiento de archivos protegidos de Movistar (`Fact_*.xlsx`, `Traf_*.xlsx`, `Cargos_*.xlsx`, Recibo PDF mensual), cruce con inventario general, aplicación de descuentos del 50% y generación de Excel consolidado con 3 hojas (`CARGOS ADICIONALES`, `CARGO MENSUAL`, `Hoja1`).
- **Módulo Entel (`entel.html`)**: Módulo para la conciliación mensual dinámica de Entel Perú. Detecta automáticamente la hoja y columnas de facturación de cada mes, cruza con el Inventario General filtrando operador Entel, extrae el concepto PLAN del PDF oficial (sin confundirlo con equipos ni total a pagar), alerta discrepancias respecto a la Hoja1 previa, conserva al 100% las hojas originales del proveedor y añade la nueva hoja `Resultado Entel` en la primera posición con las columnas NUMERO, MONTO y CC, y una tabla dinámica nativa de Excel en G4:H.

---

## 🛠️ Requisitos e Instalación

### Requisitos del Sistema
- **Claro y Entel:** navegador moderno y acceso a Internet para cargar las bibliotecas CDN. No requieren Python para procesar o descargar.
- **Movistar y API Python opcional:** Python y las dependencias indicadas a continuación.
- Navegador web moderno (Chrome, Edge, Firefox).

### Instalación de dependencias
Ejecutar en la terminal o consola de comandos (CMD / PowerShell):

```bash
pip install msoffcrypto-tool openpyxl pypdf
```

*(Opcionales para manipulación adicional: `pdfplumber`, `xlsxwriter`).*

---

## ▶️ Cómo Iniciar el Servidor (Movistar / API opcional)

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

En cada procesamiento mensual, el usuario carga en pantalla únicamente **5 archivos**:

1. **Archivo Fact** (`Fact_*.xlsx`): Archivo de facturación detallada mensual (protegido con contraseña).
2. **Archivo Traf** (`Traf_*.xlsx`): Archivo de tráfico y planes tarifarios (protegido con contraseña).
3. **Archivo Cargos** (`Cargos_*.xlsx`): Archivo de cargos adicionales y cuotas de equipos (protegido con contraseña).
4. **Inventario Celulares General** (`Inventario Celulares General.xlsx`): Archivo con la hoja `PRODAC` que mapea `Numero` a `CC Dpto`.
5. **Recibo mensual de Movistar en PDF**: Recibo en PDF oficial de Movistar del cual se extraen los totales y los descuentos de 12 meses por Anexo.

Además, debe ingresar:
6. **Contraseña de los tres Excel Movistar**: Campo de tipo contraseña con botón para mostrar/ocultar. La contraseña se utiliza únicamente en memoria durante el procesamiento y **nunca se almacena** en ningún soporte (ni localStorage, cookies, base de datos ni logs).

### 🏷️ Extracción Dinámica de Descuentos desde el PDF
- No se requiere ninguna lista maestra ni configuración previa de descuentos.
- El sistema examina automáticamente cada sección `Anexo` del recibo PDF, extrayendo los descuentos de 12 meses (`Dscto x Perm B2B 50% al plan x 12 m` y `Descuento de 50% al plan x 12 meses`) y asociándolos al número telefónico correspondiente en la hoja `CARGO MENSUAL`. Descuentos de 18 meses son excluidos conforme a la regla de negocio.

---

## 📱 Módulo Entel (`entel.html`)

### Cómo abrir Entel sin pasos adicionales
Abra su aplicación como siempre y seleccione **Entel** en el menú. Puede usar el mismo alojamiento estático de Claro (incluido GitHub Pages o Live Server). También puede abrir directamente **entel.html** con doble clic: sus scripts no usan módulos ES ni requieren una API.

Cargue los tres archivos y pulse **Procesar archivos**; después, **Descargar Excel conciliado**. No ejecute **python server.py** para Entel. Se necesita conexión a Internet para cargar SheetJS 0.20.1, PDF.js 3.11.174 y JSZip 3.10.1 desde sus CDN, igual que Claro carga sus bibliotecas. No hay instalaciones adicionales para el usuario.

La descarga mantiene el nombre del proveedor con el sufijo ** - Conciliado.xlsx**. El procesamiento, cruce y descarga ocurren localmente: no se envían los documentos al backend.

### Arquitectura y conservación
- Claro carga realmente **js/app.js** (sin importar los otros parsers): lee con SheetJS, extrae **Cargos del Mes** con PDF.js y descarga mediante XLSX.writeFile las hojas Detalle, Resumen CC y Conciliación.
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

La prueba usa un servidor **estático** temporal, sin importar server.py, y repite Entel mediante file://. Carga los tres documentos reales en data/entel_test, verifica mayo (16 líneas, S/ 924.60, diferencia cero, tres discrepancias), tablas, filtros, CSV y descarga. Reabre el archivo con openpyxl y compara byte por byte las entradas originales del ZIP. También prueba estructura, teléfonos e importes distintos, datos inválidos y el flujo de Claro con archivos sintéticos. Las evidencias se guardan localmente en **data/browser-results/** (ignorado por Git).

Resultado comprobado en Chrome: conciliación de mayo correcta, hojas originales conservadas, fórmulas y valores calculados presentes. Se abrió además el archivo descargado en Microsoft Excel, en modo de sólo lectura: seis hojas, Resultado Entel primero, y totales recalculados de S/ 924.60 en detalle y resumen. El recibo de mayo indica como periodo de servicio **01/Jun/2026–30/Jun/2026**; se muestra lo leído del PDF, sin fijar el mes por el nombre del archivo.

### Prueba de formato y actualización en Microsoft Excel
Después de la prueba de navegador, ejecutar **powershell -File test_entel_excel.ps1** en Windows con Microsoft Excel instalado. No es un requisito de uso del módulo.

La prueba abre el archivo descargado en modo de sólo lectura, verifica que no requiera reparación, reconoce una tabla dinámica nativa y comprueba su ubicación y campos. Prueba Actualizar, un monto modificado y una línea con CC nuevo dentro de la tabla principal, sin guardar los cambios de prueba. La evidencia queda en **data/format-tests/excel-verification.json**.

El generador agrega los componentes nativos directamente al ZIP existente con JSZip; SheetJS se mantiene para leer los datos. No se añadió ninguna dependencia del usuario ni un backend. La estructura de PivotTable y su caché sigue la documentación de [Microsoft Open XML](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-pivottables).

### Pruebas Automatizadas del motor Python de Entel:
Ejecutar en terminal:
```bash
python test_entel_acceptance.py
```
Verifica ausencia de constantes fijas, conciliación exacta del caso de prueba de Mayo (16 líneas, S/ 924.60, S/ 0.00 de diferencia, 3 discrepancias de CC detectadas), funcionamiento sin Hoja1 y manejo de errores.

---

## 📊 Estructura del Excel Generado

El archivo descargado contiene exactamente tres hojas y en este orden:

1. **`CARGOS ADICIONALES`**:
   - Columnas: `NUMERO`, `DESCRIPCION`, `MONTO`.
   - Cuotas de financiamiento de equipos extraídas del archivo Cargos (`TELEF:`).
   - Fila final con fórmula de suma: `=SUM(MONTO)`.

2. **`CARGO MENSUAL`**:
   - Columnas: `NUMERO`, `PLAN`, `MONTO`, `DESC`, `MONTO F`.
   - Incluye todas las líneas telefónicas procesadas en Fact con su plan de Traf.
   - Cálculo dinámico: `MONTO = Columna T + Columna AC`, `DESC = 50%` (si aplica), `MONTO F = =D{row}-E{row}`.
   - Totales de MONTO, DESC, MONTO F y diferencia de comprobación contra el subtotal del PDF.

3. **`Hoja1`**:
   - Columnas A, B, C: `NUMERO`, `MONTO` (proviene de `MONTO F`), `CC` (obtenido de `CC Dpto` del inventario).
   - Columnas G a J: Tabla dinámica/resumen agrupada por Centro de Costo:
     - `Etiquetas de fila` (CC)
     - `Suma de MONTO`
     - `IGV` (`=H{row}*0.18`)
     - `TOTAL` (`=H{row}+I{row}`)
     - Fila `Total general` con sumas dinámicas.

---

## 🔒 Seguridad y Privacidad
- Las contraseñas de archivos nunca son guardadas en disco, base de datos ni `localStorage`.
- Los archivos desencriptados se manejan como flujos de bytes en memoria RAM (`io.BytesIO`) y se liberan al finalizar la solicitud.


## Publicación y servicio de Render

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
