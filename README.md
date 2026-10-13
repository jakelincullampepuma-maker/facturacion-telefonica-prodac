# Conciliador de Facturación Telefónica - PRODAC

Sistema web corporativo para la conciliación mensual de facturación telefónica y distribución por Centro de Costo (CC Dpto) para **PRODAC** y **PRODIMIN**.

## 🚀 Módulos Disponibles

- **Módulo Claro (`claro.html` / `index.html`)**: Conciliación mensual de Claro para PRODAC y PRODIMIN mediante lectura de factura detallada, cruce con inventario y validación con recibo PDF.
- **Módulo Movistar (`movistar.html`)**: Módulo independiente para el procesamiento de archivos protegidos de Movistar (`Fact_*.xlsx`, `Traf_*.xlsx`, `Cargos_*.xlsx`, Recibo PDF mensual), cruce con inventario general, aplicación de descuentos del 50% y generación de Excel consolidado con 3 hojas (`CARGOS ADICIONALES`, `CARGO MENSUAL`, `Hoja1`).
- **Módulo Entel (`entel.html`)**: Módulo preparado para futuras implementaciones de Entel.

---

## 🛠️ Requisitos e Instalación

### Requisitos del Sistema
- Python 3.8 o superior instalado en el equipo.
- Navegador web moderno (Chrome, Edge, Firefox).

### Instalación de dependencias
Ejecutar en la terminal o consola de comandos (CMD / PowerShell):

```bash
pip install msoffcrypto-tool openpyxl pypdf
```

*(Opcionales para manipulación adicional: `pdfplumber`, `xlsxwriter`).*

---

## ▶️ Cómo Iniciar el Servidor

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
