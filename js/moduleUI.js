/* Presentación compartida. No lee archivos ni modifica los resultados del motor. */
(function () {
  'use strict';
  const q = id => document.getElementById(id);
  const months = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  function period(value) {
    const raw = String(value || '');
    const named = raw.match(/(?:\d{1,2}[\/\-])?([a-záéíóú]+)[\s\/\-]+(\d{4})/i);
    if (named) {
      const prefix = named[1].toLowerCase().slice(0,3);
      const index = months.findIndex(m => m.toLowerCase().startsWith(prefix));
      if (index >= 0) return months[index] + ' ' + named[2];
    }
    const date = raw.match(/\b\d{1,2}[\/-](\d{1,2})[\/-](\d{4})\b/);
    if (date && +date[1] >= 1 && +date[1] <= 12) return months[+date[1]-1] + ' ' + date[2];
    return raw || 'pendiente de detectar';
  }
  function setPeriod(value) { q('period-display-text').textContent = 'Periodo: ' + period(value); }
  function status({cuadra, observations = 0, missing = 0, difference, detail = ''}) {
    const type = !cuadra ? 'error' : observations || missing ? 'warning' : 'success';
    const labels = {success:['La conciliación cuadra','CUADRA'],warning:['Hay observaciones por revisar','REVISAR'],error:['La conciliación no cuadra','NO CUADRA']};
    q('conciliation-banner').dataset.state = type;
    q('banner-title-text').textContent = labels[type][0];
    q('banner-pill-badge').textContent = labels[type][1];
    const notes = [];
    if (!cuadra) notes.push('Diferencia: ' + difference + '.');
    if (missing) notes.push(missing + ' línea(s) sin centro de costo.');
    if (observations) notes.push(observations + ' observación(es) o discrepancia(s) por revisar.');
    if (detail) notes.push(detail);
    q('banner-sub-text').textContent = notes.join(' ');
    const paths = {success:'<path d="m5 12 4 4L19 6"/>', warning:'<path d="M12 6v7m0 4h.01"/>',error:'<path d="m6 6 12 12M6 18 18 6"/>'};
    q('banner-status-circle').innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" aria-hidden="true">'+paths[type]+'</svg>';
  }
  function rows(id, values) {
    const body = q(id), fragment = document.createDocumentFragment();
    for (const cells of values) {
      const row = document.createElement('tr');
      for (const value of cells) { const cell = document.createElement('td'); cell.textContent = value ?? ''; row.append(cell); }
      fragment.append(row);
    }
    body.replaceChildren(fragment);
  }
  function review({body, values, notes = [], empty = 'empty-sin-cc-msg', button}) {
    rows(body, values);
    const host = q(body).closest('.table-widget-card');
    host.querySelector('.table-scroll').hidden = !values.length;
    q(empty).classList.toggle('d-none', Boolean(values.length || notes.length));
    if (button) q(button).classList.toggle('d-none', !values.length);
  }
  function difference(id, value) {
    q(id).className = 'kpi-value-bold ' + (Number(value) === 0 ? 'text-green' : 'text-red');
  }
  const files = cards => cards.map(card => ({...card, emptyLabel:'Seleccionar archivo'}));
  window.ModuleUI = {period, setPeriod, status, rows, review, difference, files};
})();
