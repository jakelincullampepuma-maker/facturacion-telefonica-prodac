/* Una plantilla y un controlador de navegación para todas las operadoras. */
(function () {
  'use strict';
  const sidebar = document.querySelector('.sidebar');
  if (!sidebar) return;
  const icons = {
    dashboard: '<path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/>',
    upload: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><path d="M12 18v-6m-3 3 3-3 3 3"/>',
    results: '<path d="M18 20V10M12 20V4M6 20v-6"/>',
    'sin-cc': '<path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><path d="M12 9v4m0 4h.01"/>',
    export: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5m-5 5V3"/>'
  };
  const svg = body => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">${body}</svg>`;
  const links = [['dashboard','Dashboard'], ['upload','Cargar archivos'], ['results','Resultados'],
                 ['sin-cc','Sin CC - Revisar'], ['export','Descargar reporte']];
  const operators = [['claro','Claro'], ['movistar','Movistar'], ['entel','Entel']];
  sidebar.setAttribute('aria-label', 'Navegación principal');
  sidebar.id = 'shared-sidebar';
  sidebar.innerHTML = `
    <div class="sidebar-header">
      <a href="index.html" class="sidebar-logo" aria-label="PRODAC, inicio">
        <svg class="prodac-logo-svg" viewBox="0 0 135 38" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
          <path d="M19 3L34 11.5V28.5L19 37L4 28.5V11.5L19 3Z" stroke="#00A3E0" stroke-width="3.5" stroke-linejoin="round"/>
          <path d="M19 12L25 15.5V22.5L19 26L13 22.5V15.5L19 12Z" stroke="#00A3E0" stroke-width="2.5" stroke-linejoin="round"/>
          <path d="M13 19V32" stroke="#00A3E0" stroke-width="3.2" stroke-linecap="round"/>
          <text x="42" y="27" fill="#FFFFFF" font-family="Inter, system-ui, sans-serif" font-size="22" font-weight="800" letter-spacing="-0.5px">prodac</text>
        </svg>
      </a>
    </div>
    <nav class="sidebar-nav" aria-label="Secciones del módulo">
      ${links.map(([id,label]) => `<button type="button" class="nav-item sidebar-link${id === 'dashboard' ? ' active' : ''}" id="nav-${id}">${svg(icons[id])}<span>${label}</span></button>`).join('')}
      <div class="operators-section">
        <div class="sidebar-section-title">Páginas de Operadoras</div>
        ${operators.map(([id,label]) => `<a href="${id}.html" class="operator-item operator-link${sidebar.dataset.operator === id ? ' active' : ''}" data-operator="${id}"${sidebar.dataset.operator === id ? ' aria-current="page"' : ''}>
          <span class="op-left"><span class="operator-icon" aria-hidden="true">${id === 'entel' ? 'e)' : id === 'movistar' ? 'M' : svg('<circle cx="12" cy="12" r="2"/><path d="M16.24 7.76a6 6 0 0 1 0 8.49m-8.48-.01a6 6 0 0 1 0-8.49m11.31-2.82a10 10 0 0 1 0 14.14m-14.14 0a10 10 0 0 1 0-14.14"/>')}</span><span>${label}</span></span>
          <span class="op-badge-soon">Habilitado</span>
        </a>`).join('')}
      </div>
    </nav>
    <div class="sidebar-footer">Conectando negocios con el mundo</div>`;

  const toggle = document.createElement('button');
  toggle.type = 'button'; toggle.className = 'sidebar-toggle'; toggle.id = 'sidebar-menu-button';
  toggle.setAttribute('aria-label', 'Abrir o cerrar menú'); toggle.setAttribute('aria-controls', sidebar.id); toggle.setAttribute('aria-expanded', 'false');
  toggle.innerHTML = svg('<path d="M4 6h16M4 12h16M4 18h16"/>');
  const backdrop = document.createElement('button'); backdrop.type = 'button'; backdrop.className = 'sidebar-backdrop';
  backdrop.tabIndex = -1; backdrop.setAttribute('aria-label', 'Cerrar menú');
  sidebar.after(toggle, backdrop);
  const close = () => { sidebar.classList.remove('is-open'); backdrop.classList.remove('is-open'); toggle.setAttribute('aria-expanded','false'); };
  toggle.onclick = () => { const open = toggle.getAttribute('aria-expanded') !== 'true'; sidebar.classList.toggle('is-open',open); backdrop.classList.toggle('is-open',open); toggle.setAttribute('aria-expanded',String(open)); };
  backdrop.onclick = close;
  document.addEventListener('keydown', event => { if (event.key === 'Escape') close(); });
  function visible(element) { return Boolean(element && element.getClientRects().length); }
  window.AppSidebar = {
    bind({results, missing, download}) {
      const target = value => typeof value === 'function' ? value() : document.getElementById(value);
      for (const [id] of links) document.getElementById('nav-' + id).onclick = () => {
        close();
        if (id === 'export') {
          const button = target(download);
          if (button && !button.disabled) button.click();
          return;
        }
        if (id === 'dashboard') window.scrollTo({top: 0, behavior: 'smooth'});
        else {
          if (id !== 'upload' && !visible(target(results))) return;
          const section = target(id === 'upload' ? 'section-carga-archivos' : id === 'results' ? results : missing);
          if (!visible(section)) return;
          section.scrollIntoView({behavior: 'smooth', block: 'start'});
        }
        for (const button of sidebar.querySelectorAll('.sidebar-link')) {
          const active = button.id === 'nav-' + id;
          button.classList.toggle('active', active);
          if (active) button.setAttribute('aria-current', 'location'); else button.removeAttribute('aria-current');
        }
      };
    }
  };
})();
