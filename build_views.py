"""Plantilla visual única. Regenerar las cuatro páginas con python build_views.py.

Sólo genera HTML estático: Entel sigue funcionando sin servidor de procesamiento.
Los identificadores se conservan para los controladores de cada operador.
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parent
VIEWS = {
    'claro': dict(name='Claro', result='claroResultsSection', review='claroMissingCcSection',
        files=[('claro','Excel de Claro','XLSX · Factura Detallada'),('inv','Inventario General','XLSX'),('pdf','Recibo PDF de Claro','PDF')],
        kpis=[('lineas','Líneas procesadas','Números únicos y registros de cargos','lineas'),('total-lineas','Total antes del ajuste','Antes del ajuste o Varios',''),('ajustes','Ajustes o Varios','Ajustes del periodo',''),('total','Total final','Líneas + ajustes',''),('cargos-mes','Cargos del Mes','Según recibo PDF',''),('diferencia','Diferencia','','diferencia')]),
    'movistar': dict(name='Movistar', result='section-resultados-movistar', review='section-sin-cc-movistar',
        files=[('pdf','Recibo PDF Movistar','PDF'),('inventory','Inventario Celulares General','XLSX · Hoja PRODAC')],
        kpis=[('lineas','Líneas / Anexos','','lineas'),('monto','MONTO total','','monto'),('desc','Descuentos (DESC)','','desc'),('monto-f','MONTO F (Subtotal)','','subtotal'),('igv','IGV (18%)','18% sobre Subtotal',''),('total-afecto','Total afecto','Subtotal + IGV',''),('inafectos','Cargos inafectos','','inafectos'),('redondeo','Redondeo','Neto en recibo PDF',''),('total-pdf','Total recibo PDF','','calc'),('total-calculado','Total calculado','Total afecto + inafectos + redondeo',''),('diferencia','Diferencia','','diferencia')]),
    'entel': dict(name='Entel', result='section-resultados-entel', review='section-sin-cc-entel',
        files=[('excel','Excel Facturación Entel','XLSX · Original del proveedor'),('inventory','Inventario Celulares General','XLSX · PRODAC / PRODIMIN'),('pdf','Recibo PDF mensual Entel','PDF')],
        kpis=[('lineas','Líneas facturadas','','lineas'),('conciliadas','Líneas conciliadas','','conciliadas'),('observaciones','Observaciones','','observaciones'),('excel-plan','Total Planes Excel','Suma de cargos fijos',''),('pdf-plan','Concepto PLAN PDF','Según recibo oficial',''),('equipos','Equipos inafectos','Según recibo oficial',''),('total-factura','Total recibo factura','Según recibo oficial',''),('diferencia','Diferencia','','diferencia')]),
}


def icon(kind='file'):
    paths = {'file':'<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M8 13h8M8 17h8"/>',
             'chart':'<path d="M6 20v-6M12 20V4M18 20V10"/>',
             'check':'<path d="m5 12 4 4L19 6"/>',
             'remove':'<path d="m6 6 12 12M6 18 18 6"/>',
             'download':'<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M12 3v12m-5-5 5 5 5-5"/>',
             'clear':'<path d="M3 6h18M19 6v14H5V6m4 0V3h6v3"/>',
             'play':'<path d="m8 4 12 8-12 8z"/>',
             'calendar':'<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>',
             'warning':'<path d="m12 3 10 18H2L12 3zM12 9v4m0 4h.01"/>'}
    return f'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">{paths[kind]}</svg>'


def file_card(key, label, fmt, number):
    pdf = key == 'pdf'
    return f'''<div class="upload-file-card empty-border" id="card-wrap-{key}">
      <div class="card-step-num">{number:02}</div><span class="card-check-corner d-none" id="check-icon-{key}">{icon('check')}</span>
      <h3 class="card-file-heading">{label}</h3><p class="card-file-format">{fmt}</p>
      <div class="card-drop-box" id="dropzone-{key}" tabindex="0" role="button" aria-label="Seleccionar {label}">
        <div class="card-file-preview-left"><span class="icon-file-badge {'icon-pdf-red' if pdf else 'icon-excel-green'}">{'PDF' if pdf else 'X'}</span><span class="card-filename-text" id="name-{key}-file">Seleccionar archivo</span></div>
        <button type="button" class="btn-remove-card-file d-none" id="btn-remove-{key}" title="Retirar archivo" aria-label="Retirar {label}">{icon('remove')}</button>
      </div><input type="file" id="input-{key}-file" accept="{'.pdf' if pdf else '.xlsx'}" class="d-none">
    </div>'''


def table(body, headings):
    return f'<div class="table-scroll"><table class="clean-corp-table"><thead><tr>'+''.join(f'<th>{h}</th>' for h in headings)+f'</tr></thead><tbody id="{body}"></tbody></table></div>'


def optional_files():
    cards = ''.join(file_card(k, f'Archivo {k.title()} (Opcional)', 'XLSX · Protegido', n) for n,k in enumerate(['fact','traf','cargos'],1))
    return f'''<details class="optional-validation-details" id="details-optional-val"><summary>Validación opcional</summary><div class="optional-validation-content">
      <p>Compare los datos del PDF con los reportes del portal de Movistar.</p><div class="upload-cards-row">{cards}</div>
      <div class="password-box"><label for="input-movistar-password">Contraseña de los archivos protegidos (si cargó Fact, Traf o Cargos)</label>
      <div class="password-field-wrap"><input type="password" id="input-movistar-password" placeholder="Contraseña mensual" autocomplete="new-password"><button type="button" id="btn-toggle-password" aria-label="Mostrar u ocultar contraseña"><span id="icon-eye-open">Mostrar</span><span id="icon-eye-closed" class="d-none">Ocultar</span></button></div></div>
    </div></details>'''


def detail(operator):
    if operator == 'movistar': return ''
    if operator == 'claro':
        return f'''<section class="table-widget-card detail-section d-none" id="full-detail-section"><div class="table-widget-header"><h2 class="table-widget-title">Detalle completo de facturación · <span id="label-full-company">PRODAC</span></h2><button type="button" class="btn-white-action" id="btn-close-full-view">Cerrar vista extendida</button></div>
        <div class="table-filter-bar"><input class="filter-input-search" id="filter-search-input" placeholder="Buscar por número, usuario o CC..." aria-label="Buscar líneas"></div>
        {table('tbody-full-detail-list',['N°','Línea','Usuario','Centro de costo','Descripción','Cargo del mes','Observación'])}</section>'''
    return f'''<section class="table-widget-card detail-section" id="section-detalle-entel"><div class="table-widget-header"><h2 class="table-widget-title">Detalle de Líneas Conciliadas – Hoja Resultado Entel</h2><span class="count-badge" id="badge-filtered-count"></span></div>
      <div class="table-filter-bar"><input class="filter-input-search" id="input-search-lines" placeholder="Buscar por número, usuario, área, dpto o CC..." aria-label="Buscar líneas"><select class="filter-select" id="select-filter-status" aria-label="Filtrar estado"><option value="all">Todos los estados</option><option value="conciliado">Conciliadas</option><option value="sin_cc">Sin CC</option><option value="observacion">Con observaciones</option></select></div>
      {table('tbody-lineas-entel',['Número','Plan','Cargo fijo','Empresa','Usuario','Área / Dpto','CC Dpto','PRS','Estado','Observación'])}</section>'''


def render(operator):
    v = VIEWS[operator]; name = v['name']
    company = '<button type="button" class="btn-company-tab active" id="tab-prodac" data-company="PRODAC">PRODAC</button><button type="button" class="btn-company-tab" id="tab-prodimin" data-company="PRODIMIN">PRODIMIN</button>' if operator == 'claro' else ''
    cards = ''.join(file_card(*f,n) for n,f in enumerate(v['files'],1))
    kpis = ''
    for key,label,sub,subid in v['kpis']:
        subattr = f' id="kpi-sub-{subid}"' if subid else ''
        if subid == 'subtotal': subattr = ' id="subtotal-comparison"'
        kpis += f'<div class="card-kpi-concil"><div class="kpi-icon-wrapper">{icon("chart" if key=="diferencia" else "file")}</div><div class="kpi-content-box"><span class="kpi-label-small">{label}</span><span class="kpi-value-bold" id="kpi-val-{key}"></span><span class="kpi-subtext-ref"{subattr}>{sub}</span></div></div>'
    badges = '<div class="breakdown-badges">'+''.join(f'<span class="count-badge" id="badge-{k}"></span>' for k in ['desc-12m','desc-3m','desc-perm','desc-18m','inaf-equipos','inaf-ventas'])+'</div>' if operator == 'movistar' else ''
    reviewbody = {'claro':'tbody-sin-cc','movistar':'tbody-sin-cc-movistar','entel':'tbody-observaciones-entel'}[operator]
    reviewcols = ['Línea','Usuario','Centro de costo','Cargo del mes','Observación'] if operator == 'claro' else ['Número','Plan','Monto final','Observación'] if operator == 'movistar' else ['Número','Usuario','CC Dpto','Cargo fijo','Observación']
    reviewbutton = '<button type="button" class="btn-white-action" id="btn-ver-todos-detalle">Ver todos</button>' if operator == 'claro' else '<button type="button" class="btn-white-action d-none" id="btn-download-sin-cc">Descargar observaciones</button>' if operator == 'movistar' else '<button type="button" class="btn-white-action d-none" id="btn-download-obs">Descargar observaciones (CSV)</button>'
    notes = '<div class="review-notes" id="warnings" hidden></div>' if operator == 'claro' else '<div class="review-notes d-none" id="movistar-warnings-container"></div><div class="review-notes d-none" id="movistar-optional-warnings-container"></div>' if operator == 'movistar' else ''
    emptyid = 'empty-sin-cc-msg' if operator != 'entel' else 'empty-obs-msg'
    error = '<p id="process-message" class="process-error" role="alert" hidden></p>' if operator == 'claro' else f'<div class="process-error" id="{operator}-alert-container" role="alert"></div>'
    scripts = '<script type="module" src="js/app.js"></script>' if operator=='claro' else '<script src="js/movistar.js"></script>' if operator=='movistar' else '<script src="js/entelWorkbook.js"></script><script src="js/entelProcessor.js"></script><script src="js/entel.js"></script>'
    return f'''<!doctype html>
<!-- Generado por build_views.py: editar la plantilla compartida, no esta página. -->
<html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Módulo {name} | Conciliación Telefónica PRODAC</title>
<link rel="icon" type="image/svg+xml" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect width='32' height='32' rx='7' fill='%230f2746'/%3E%3Cpath d='M9 24V8h7a6 6 0 0 1 0 12h-3v4zm4-8h3a2 2 0 0 0 0-4h-3z' fill='white'/%3E%3C/svg%3E">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin><link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<link rel="stylesheet" href="css/styles.css"><link rel="stylesheet" href="css/components.css"></head><body>
<aside class="sidebar" data-operator="{operator}"></aside><script src="js/sidebar.js"></script>
<main class="main-wrapper module-page" data-operator="{operator}">
  <header class="top-header-row module-header"><div class="header-title-group"><h1>Módulo {name}</h1><p>Conciliación mensual por centro de costo</p><div class="company-tabs-container company-slot">{company}</div></div><div class="period-badge-btn">{icon('calendar')}<span id="period-display-text">Periodo: pendiente de detectar</span></div></header>
  <section class="upload-panel-card" id="section-carga-archivos"><div class="upload-panel-header"><div class="upload-panel-title"><span class="badge-num-step">1</span><h2>Carga de archivos – {name}</h2></div><p class="upload-panel-subtitle">Selecciona los archivos para procesar la conciliación.</p></div>
    <div class="upload-cards-row" style="--file-columns:{len(v['files'])}">{cards}</div>
    <div class="upload-buttons-bar"><button type="button" class="btn-blue-action" id="btn-process-files" disabled>{icon('play')}<span id="process-button-label" aria-live="polite">Procesar archivos</span></button><button type="button" class="btn-white-action" id="btn-clear-files">{icon('clear')}<span>Limpiar</span></button><button type="button" class="btn-outline-download" id="btn-export-excel" disabled>{icon('download')}<span>Descargar Excel conciliado</span></button></div>
    {optional_files() if operator=='movistar' else ''}{error}
  </section>
  <section class="module-results{'' if operator=='claro' else ' d-none'}" id="{v['result']}" aria-label="Resultados de conciliación" {'hidden' if operator=='claro' else ''}>
    <div class="conciliation-banner-card" id="conciliation-banner" role="status"><div class="banner-left-group"><div class="banner-check-circle" id="banner-status-circle"></div><div><h2 class="banner-title-text" id="banner-title-text"></h2><p class="banner-sub-text" id="banner-sub-text"></p></div></div><span class="banner-pill-cuadra" id="banner-pill-badge"></span></div>
    {badges}<div class="conciliation-kpis-grid">{kpis}</div>{detail(operator)}
    <div class="two-columns-bottom-grid"><section class="table-widget-card" id="section-resumen-cc"><div class="table-widget-header"><h2 class="table-widget-title">{icon('chart')}Resumen por CC Dpto.</h2></div>{table('tbody-resumen-cc',['Centro de costo','Líneas','Costo sin IGV','% del total'])}</section>
      <section class="table-widget-card" id="{v['review']}"><div class="table-widget-header"><h2 class="table-widget-title">{icon('warning')}Sin CC - Revisar</h2>{reviewbutton}</div>{table(reviewbody,reviewcols)}{notes}<p class="empty-review d-none" id="{emptyid}">Todas las líneas están conciliadas y tienen centro de costo.</p></section>
    </div>
  </section>
</main><script src="js/moduleUI.js"></script><script src="js/uploadUI.js"></script>{scripts}</body></html>
'''


if __name__ == '__main__':
    for op in VIEWS:
        (ROOT / f'{op}.html').write_text(render(op), encoding='utf-8')
    (ROOT / 'index.html').write_text(render('claro'), encoding='utf-8')
