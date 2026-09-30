/* ------------------------------------------------------------------
   Tablero de control — Graduados · UNIMINUTO Sede Tolima-Huila
   data.json (generado por build_data.py desde Graduados_SAP_<fecha>.xlsx):
   { corte, anios, periodos, anio_parcial, dims:{cu:[...],...},
     cols:[anio,sem,cu,modalidad,programa,area,nivel,genero,edad,estrato,n],
     rows:[[2025,1,0,0,3,0,0,0,0,1,12],...], prog_meta, calidad }
   Cada fila es un conteo agregado: no hay datos personales.
------------------------------------------------------------------- */

const COL = { ink: '#16324F', accent: '#E2962B', teal: '#2F7A6D', purple: '#8A5FBF', nat: '#8A9BAE', danger: '#C1432B', grid: '#EDEFEC', muted: '#6B7280' };
const CU_COLORS = { 'Ibagué': '#E2962B', 'Neiva': '#2F7A6D', 'Garzón': '#C1432B', 'Pitalito': '#3E6FA8', 'Lérida': '#8A5FBF', 'La Dorada': '#D1709B' };
const NIVEL_COLORS = { 'Pregrado': '#16324F', 'Especialización': '#E2962B', 'Maestría': '#8A5FBF', 'Técnico Profesional': '#2F7A6D', 'Sin dato': '#B8BEC6' };
const POSGRADO = ['Especialización', 'Maestría'];
const posg = m => sum(POSGRADO.map(n => g(m, n)));
const posg2 = (m, a) => sum(POSGRADO.map(n => g2(m, a, n)));
const abrev = p => p.replace('Especialización en ', 'Esp. ').replace('Maestría en ', 'Maestría ');
const SEM_COLORS = { 1: '#16324F', 2: '#E2962B' };
const GEN_COLORS = { 'Femenino': '#2F7A6D', 'Masculino': '#16324F', 'Sin dato': '#B8BEC6' };
const AREA_COLORS = ['#16324F', '#E2962B', '#2F7A6D', '#8A5FBF', '#3E6FA8', '#C1432B', '#6E8B3D', '#B8BEC6'];
const EST_COLORS = ['#C1432B', '#E2962B', '#2F7A6D', '#3E6FA8', '#16324F', '#8A5FBF', '#B8BEC6'];
const TABS = [
  ['resumen', 'Panorama'],
  ['centros', 'Centros Universitarios'],
  ['programas', 'Programas'],
  ['perfil', 'Perfil del graduado'],
  ['metodologia', 'Datos y metodología'],
];
/* dimensiones filtrables: [clave, título, abierto, lista con buscador] */
const FILTERS = [
  ['cu', 'Centro Universitario', true, false],
  ['nivel', 'Nivel de formación', true, false],
  ['modalidad', 'Modalidad', false, false],
  ['area', 'Área de conocimiento', false, false],
  ['programa', 'Programa', false, true],
  ['genero', 'Género', false, false],
  ['edad', 'Rango de edad', false, false],
  ['estrato', 'Estrato socioeconómico', false, false],
];
const DIMS = ['cu', 'modalidad', 'programa', 'area', 'nivel', 'genero', 'edad', 'estrato'];

let DATA = null, IDX = {};
const charts = {};
const state = { tab: 'resumen', years: new Set(), sems: new Set([1, 2]), f: {}, search: '' };
DIMS.forEach(d => state.f[d] = new Set());

/* ---------------- utilidades ---------------- */
const el = id => document.getElementById(id);
const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const fmt = (v, d = 0) => (v == null || !isFinite(v)) ? '—' : v.toLocaleString('es-CO', { minimumFractionDigits: d, maximumFractionDigits: d });
const fmtP = (v, d = 1) => (v == null || !isFinite(v)) ? '—' : fmt(v * 100, d) + '%';
const fmtPS = (v, d = 1) => (v == null || !isFinite(v)) ? '—' : (v > 0 ? '+' : v < 0 ? '−' : '') + fmt(Math.abs(v) * 100, d) + '%';
const sum = a => a.reduce((s, x) => s + x, 0);
const isPartial = y => DATA.anio_parcial.includes(y);
const ylab = y => isPartial(y) ? `${y}*` : String(y);
const alpha = (hex, a) => hex + Math.round(a * 255).toString(16).padStart(2, '0');
const colFor = (y, base) => isPartial(y) ? alpha(base, 0.5) : base;
const yrs = () => DATA.anios.filter(y => state.years.has(y));
const short = (s, n = 42) => s.length > n ? s.slice(0, n - 1) + '…' : s;
const corteDate = () => { const [a, m, d] = DATA.corte.split('-').map(Number); return new Date(a, m - 1, d); };
const fmtCorte = () => corteDate().toLocaleDateString('es-CO', { day: 'numeric', month: 'long', year: 'numeric' });
const gapSpan = (v, note = '') => v == null ? '<span class="flat">—</span>' : `<span class="gap ${v >= 0 ? 'pos' : 'neg'}"${note ? ` title="${note}"` : ''}>${fmtPS(v)}</span>`;
const pTag = y => isPartial(y) ? `<span class="ptag">solo ${y}-1</span>` : '';

/* ---------------- filtrado y agregación ---------------- */
function pass(r, skip) {
  if (skip !== 'years' && !state.years.has(r[0])) return false;
  if (skip !== 'sems' && !state.sems.has(r[1])) return false;
  for (let i = 0; i < DIMS.length; i++) {
    const d = DIMS[i]; if (d === skip) continue;
    const s = state.f[d]; if (s.size && !s.has(DATA.dims[d][r[IDX[d]]])) return false;
  }
  return true;
}
const filtered = skip => DATA.rows.filter(r => pass(r, skip));
const total = rows => { let t = 0; for (const r of rows) t += r[10]; return t; };
const keyOf = (r, k) => k === 'anio' ? r[0] : k === 'sem' ? r[1] : DATA.dims[k][r[IDX[k]]];
function by(rows, k) { const m = new Map(); for (const r of rows) { const v = keyOf(r, k); m.set(v, (m.get(v) || 0) + r[10]); } return m; }
function cross(rows, k1, k2) {
  const m = new Map();
  for (const r of rows) { const a = keyOf(r, k1), b = keyOf(r, k2); if (!m.has(a)) m.set(a, new Map()); const mm = m.get(a); mm.set(b, (mm.get(b) || 0) + r[10]); }
  return m;
}
const g = (m, k) => (m && m.get(k)) || 0;
const g2 = (m, a, b) => g(m.get(a), b);
const ranked = m => [...m.entries()].filter(([, v]) => v > 0).sort((a, b) => b[1] - a[1]);

/* ---------------- Chart.js: etiquetas propias (sin plugin externo) ---------------- */
const InlineLabels = {
  id: 'inlineLabels',
  afterDatasetsDraw(chart, _a, o) {
    if (!o || !o.enabled) return;
    const ctx = chart.ctx; const horiz = chart.options.indexAxis === 'y';
    ctx.save(); ctx.textBaseline = 'middle';
    chart.data.datasets.forEach((ds, di) => {
      if (ds.labels === false || !chart.isDatasetVisible(di)) return;
      const meta = chart.getDatasetMeta(di);
      const type = ds.type || chart.config.type;
      meta.data.forEach((pt, i) => {
        const v = ds.data[i]; if (v == null || !isFinite(v)) return;
        const f = ds.labelFmt || o.fmt || (x => fmt(x));
        ctx.font = `600 ${o.size || 10.5}px 'IBM Plex Sans', sans-serif`;
        ctx.fillStyle = ds.labelColor || o.color || '#3B4B5C';
        if (type === 'doughnut') {
          const tot = sum(ds.data); if (!tot || v / tot < 0.045) return;
          const p = pt.tooltipPosition(); ctx.textAlign = 'center'; ctx.fillStyle = '#fff';
          ctx.fillText(fmtP(v / tot, 0), p.x, p.y); return;
        }
        if (ds.inside) {                                   // etiqueta centrada dentro del segmento
          if (o.minInside != null && v < o.minInside) return;
          const cx = horiz ? (pt.x + pt.base) / 2 : pt.x, cy = horiz ? pt.y : (pt.y + pt.base) / 2;
          const span = horiz ? Math.abs(pt.x - pt.base) : Math.abs(pt.y - pt.base);
          const txt = f(v, i); if (span < ctx.measureText(txt).width + 6 && horiz) return; if (!horiz && span < 13) return;
          ctx.textAlign = 'center'; ctx.fillStyle = ds.labelColor || '#fff'; ctx.fillText(txt, cx, cy); return;
        }
        const txt = f(v, i);
        if (type === 'line') {
          const pos = typeof ds.labelPos === 'function' ? ds.labelPos(i) : (ds.labelPos || 'top');
          ctx.textAlign = 'center';
          const yy = pt.y + (pos === 'top' ? -12 : 12);
          ctx.lineWidth = 3.5; ctx.strokeStyle = 'rgba(255,255,255,0.95)'; ctx.lineJoin = 'round';
          ctx.strokeText(txt, pt.x, yy);
          ctx.fillStyle = ds.labelColor || ds.borderColor;
          ctx.fillText(txt, pt.x, yy);
        } else if (horiz) {
          ctx.textAlign = 'left'; ctx.fillText(txt, pt.x + 5, pt.y);
        } else {
          ctx.textAlign = 'center'; ctx.fillText(txt, pt.x, pt.y + (v < 0 ? 9 : -9));
        }
      });
    });
    ctx.restore();
  }
};
const StackTotals = {
  id: 'stackTotals',
  afterDatasetsDraw(chart, _a, o) {
    if (!o || !o.enabled) return;
    const ctx = chart.ctx; const n = chart.data.labels.length;
    ctx.save(); ctx.font = `700 ${o.size || 11.5}px 'Space Grotesk', sans-serif`; ctx.fillStyle = COL.ink; ctx.textAlign = 'center';
    for (let i = 0; i < n; i++) {
      let tot = 0, top = Infinity, x = null;
      chart.data.datasets.forEach((ds, di) => {
        if (!chart.isDatasetVisible(di)) return;
        const e = chart.getDatasetMeta(di).data[i]; tot += ds.data[i] || 0;
        if (e && (ds.data[i] || 0) > 0) { top = Math.min(top, e.y); x = e.x; }
      });
      if (x != null && tot > 0) ctx.fillText(fmt(tot), x, top - 9);
    }
    ctx.restore();
  }
};
Chart.register(InlineLabels, StackTotals);
Chart.defaults.font.family = "'IBM Plex Sans', sans-serif";
Chart.defaults.color = COL.muted;
Chart.defaults.locale = 'es-CO';
Chart.defaults.elements.line.cubicInterpolationMode = 'monotone';
const baseGrid = { color: COL.grid };
const tick = { font: { size: 11 }, color: COL.muted };
const xTick = { font: { size: 11.5 }, color: COL.ink };
const legendOpts = (pos = 'bottom') => ({ position: pos, labels: { boxWidth: 10, boxHeight: 10, font: { size: 11 }, color: COL.ink, padding: 14 } });
const lblSize = n => n > 16 ? 9 : n > 11 ? 10 : 11.5;

function ensureChart(id, config) {
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(el(id).getContext('2d'), config);
}

/* ---------------- arranque ---------------- */
async function boot() {
  DATA = window.__EMBEDDED_DATA__ || await (await fetch('data.json')).json();
  DATA.cols.forEach((c, i) => IDX[c] = i);
  DATA.anios.forEach(y => state.years.add(y));
  const pUlt = DATA.periodos[DATA.periodos.length - 1];
  el('as-of').innerHTML = `Corte de datos: <b>${fmtCorte()}</b> · último periodo ${pUlt}` + (DATA.anio_parcial.length ? ` · <span class="pending">${DATA.anio_parcial[0]}-2 pendiente</span>` : '');
  el('brand-sub').textContent = `Grados ${DATA.periodos[0]} a ${pUlt}`;
  el('side-note').innerHTML = `Todos los filtros se combinan entre sí y aplican a todas las pestañas. Las cifras junto a cada opción muestran cuántos graduados quedarían al elegirla con los demás filtros activos.${DATA.anio_parcial.length ? `<br><br>${DATA.anio_parcial[0]}* solo incluye el periodo ${DATA.anio_parcial[0]}-1; el ${DATA.anio_parcial[0]}-2 se incorporará cuando SAP lo reporte.` : ''}`;
  el('tabs').innerHTML = TABS.map(([k, l]) => `<button class="tab" data-tab="${k}">${l}</button>`).join('');
  el('tabs').addEventListener('click', e => {
    const b = e.target.closest('.tab'); if (!b) return;
    state.tab = b.dataset.tab; render(); window.scrollTo({ top: 0, behavior: 'smooth' });
  });
  el('credit').innerHTML = `Fuente: listado de graduados SAP (${esc(DATA.fuente)}), corte al ${fmtCorte()}. Datos agregados por periodo, Centro Universitario, programa y variables sociodemográficas, sin información personal. Elaborado para UNIMINUTO Sede Tolima-Huila.`;
  buildSidebar(); render();
}

/* ---------------- sidebar ---------------- */
function group(title, key, inner, open = true) {
  const d = document.createElement('details'); d.className = 'filter-group'; d.open = open;
  d.innerHTML = `<summary><span>${title}<span class="count-badge" data-badge="${key}" style="display:none"></span></span><span class="chev">▸</span></summary>`;
  d.appendChild(inner); return d;
}

function buildSidebar() {
  const sb = el('sidebar'); sb.querySelectorAll('.filter-group').forEach(n => n.remove());
  const before = el('clear-all');

  // Año
  const yw = document.createElement('div');
  const yl = document.createElement('div'); yl.className = 'chip-list';
  DATA.anios.forEach(y => {
    const b = document.createElement('button');
    b.className = 'chip' + (state.years.has(y) ? ' active' : ''); b.textContent = ylab(y);
    b.setAttribute('aria-pressed', state.years.has(y));
    b.title = isPartial(y) ? `Solo ${y}-1, datos al ${fmtCorte()}` : '';
    b.onclick = () => { state.years.has(y) ? state.years.delete(y) : state.years.add(y); b.classList.toggle('active'); b.setAttribute('aria-pressed', state.years.has(y)); render(); };
    yl.appendChild(b);
  });
  const q = document.createElement('div'); q.className = 'quick';
  const setYears = list => { state.years = new Set(list); buildSidebar(); render(); };
  const completos = DATA.anios.filter(y => !isPartial(y));
  [['Todos', DATA.anios], ['Desde 2016', DATA.anios.filter(y => y >= 2016)], ['Últimos 5 completos', completos.slice(-5)]].forEach(([l, list]) => {
    const b = document.createElement('button'); b.textContent = l; b.onclick = () => setYears(list); q.appendChild(b);
  });
  yw.append(yl, q);
  sb.insertBefore(group('Año de grado', 'years', yw), before);

  // Semestre
  const sl = document.createElement('div'); sl.className = 'chip-list';
  [1, 2].forEach(s => {
    const b = document.createElement('button');
    b.className = 'chip' + (state.sems.has(s) ? ' active' : ''); b.textContent = `Semestre ${s}`;
    b.onclick = () => { state.sems.has(s) ? state.sems.delete(s) : state.sems.add(s); b.classList.toggle('active'); render(); };
    sl.appendChild(b);
  });
  sb.insertBefore(group('Semestre de grado', 'sems', sl), before);

  // Dimensiones categóricas
  FILTERS.forEach(([k, title, open, searchable]) => {
    const w = document.createElement('div');
    const cl = document.createElement('div'); cl.className = 'check-list' + (searchable ? ' scroll' : '');
    if (searchable) {
      const s = document.createElement('input'); s.type = 'search'; s.className = 'search'; s.placeholder = 'Buscar programa…'; s.value = state.search;
      s.setAttribute('aria-label', 'Buscar programa');
      s.oninput = () => { state.search = s.value; filterSearch(cl); };
      w.appendChild(s);
    }
    DATA.dims[k].forEach(v => {
      const r = document.createElement('label'); r.className = 'check-row'; r.dataset.dim = k; r.dataset.val = v;
      const cb = document.createElement('input'); cb.type = 'checkbox'; cb.checked = state.f[k].has(v);
      cb.onchange = () => { cb.checked ? state.f[k].add(v) : state.f[k].delete(v); render(); };
      r.appendChild(cb);
      const sw = k === 'cu' ? CU_COLORS[v] : k === 'nivel' ? NIVEL_COLORS[v] : k === 'genero' ? GEN_COLORS[v] : null;
      r.insertAdjacentHTML('beforeend', `${sw ? `<span class="sw" style="background:${sw}"></span>` : ''}<span class="lbl">${esc(v)}</span><span class="n"></span>`);
      cl.appendChild(r);
    });
    w.appendChild(cl);
    if (searchable) filterSearch(cl);
    sb.insertBefore(group(title, k, w, open || state.f[k].size > 0), before);
  });

  before.onclick = () => { state.years = new Set(DATA.anios); state.sems = new Set([1, 2]); DIMS.forEach(d => state.f[d].clear()); state.search = ''; buildSidebar(); render(); };
}

function filterSearch(cl) {
  const q = state.search.trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  cl.querySelectorAll('.check-row').forEach(r => {
    const t = r.dataset.val.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    r.style.display = !q || t.includes(q) ? '' : 'none';
  });
}

/* conteos cruzados junto a cada opción */
function syncCounts() {
  FILTERS.forEach(([k]) => {
    const m = by(filtered(k), k);
    document.querySelectorAll(`.check-row[data-dim="${k}"]`).forEach(r => {
      const n = g(m, r.dataset.val);
      r.querySelector('.n').textContent = fmt(n);
      r.classList.toggle('zero', n === 0 && !state.f[k].has(r.dataset.val));
    });
  });
  const set = (k, n) => { const b = document.querySelector(`[data-badge="${k}"]`); if (b) { b.style.display = n ? 'inline-block' : 'none'; b.textContent = n; } };
  set('years', state.years.size === DATA.anios.length ? 0 : state.years.size);
  set('sems', state.sems.size === 2 ? 0 : state.sems.size);
  DIMS.forEach(d => set(d, state.f[d].size));
}

function rangeText(list) {
  if (!list.length) return 'ninguno';
  const consecutive = list.every((y, i) => i === 0 || DATA.anios.indexOf(y) === DATA.anios.indexOf(list[i - 1]) + 1);
  return consecutive && list.length > 2 ? `${ylab(list[0])}–${ylab(list[list.length - 1])}` : list.map(ylab).join(', ');
}

/* ---------------- render ---------------- */
function render() {
  syncCounts();
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.dataset.tab === state.tab));
  const Y = yrs();
  const parts = [`<span class="fs-item">Años: <b>${rangeText(Y)}</b></span>`];
  if (state.sems.size < 2) parts.push(`<span class="fs-item">Semestre: <b>${[...state.sems].join(', ') || 'ninguno'}</b></span>`);
  FILTERS.forEach(([k, t]) => { if (state.f[k].size) parts.push(`<span class="fs-item">${t}: <b>${[...state.f[k]].map(v => esc(short(v, 34))).join(', ')}</b></span>`); });
  if (!parts.slice(1).length) parts.push('<span class="fs-item">Centros Universitarios: <b>todos</b></span>');
  if (Y.some(isPartial)) parts.push(`<span class="fs-item"><span class="ptag">*</span> ${DATA.anio_parcial[0]} solo con el periodo ${DATA.anio_parcial[0]}-1</span>`);
  el('filter-summary').innerHTML = parts.join(' · ');
  const rows = filtered();
  const empty = !total(rows) && state.tab !== 'metodologia';
  el('empty-state').style.display = empty ? 'block' : 'none';
  document.querySelectorAll('.view').forEach(v => v.classList.toggle('active', !empty && v.id === 'view-' + state.tab));
  if (empty) return;
  ({ resumen: renderResumen, centros: renderCentros, programas: renderProgramas, perfil: renderPerfil, metodologia: renderMetodologia })[state.tab](rows);
}

function kpi(label, value, delta, color) {
  return `<div class="kpi"${color ? ` style="border-left-color:${color}"` : ''}><div class="label">${label}</div><div class="num">${value}</div>${delta ? `<div class="delta">${delta}</div>` : ''}</div>`;
}

/* variación anual; el año parcial se compara semestre a semestre */
function yoy(y) {
  const base = filtered('years');
  const ys = by(base, 'anio');
  if (isPartial(y)) {
    const ps = cross(base, 'anio', 'sem');
    const a = g2(ps, y, 1), b = g2(ps, y - 1, 1);
    return b ? { v: a / b - 1, note: `${y}-1 frente a ${y - 1}-1` } : null;
  }
  const a = g(ys, y), b = g(ys, y - 1);
  return b ? { v: a / b - 1, note: `${y} frente a ${y - 1}` } : null;
}

/* ================================================================
   PANORAMA
================================================================ */
function renderResumen(rows) {
  const Y = yrs();
  const tot = total(rows);
  const byY = by(rows, 'anio'), ys = cross(rows, 'anio', 'sem');
  const comp = Y.filter(y => !isPartial(y));
  const promComp = comp.length ? sum(comp.map(y => g(byY, y))) / comp.length : null;
  const best = [...comp].sort((a, b) => g(byY, b) - g(byY, a))[0];
  const gen = by(rows, 'genero'), niv = by(rows, 'nivel');

  // último periodo con datos
  const perRows = cross(rows, 'anio', 'sem');
  const per = DATA.periodos.map(p => p.split('-').map(Number)).filter(([a, s]) => state.years.has(a) && state.sems.has(s));
  const withData = per.filter(([a, s]) => g2(perRows, a, s) > 0);
  const last = withData[withData.length - 1];
  let lastTxt = '', lastVal = '—';
  if (last) {
    const [a, s] = last; lastVal = fmt(g2(perRows, a, s));
    const prev = g2(cross(filtered('years'), 'anio', 'sem'), a - 1, s);
    lastTxt = `Periodo ${a}-${s}` + (prev ? ` · frente a ${a - 1}-${s}: ${gapSpan(g2(perRows, a, s) / prev - 1)}` : '');
  }

  el('kpi-res').innerHTML = [
    kpi('Graduados en la selección', fmt(tot), promComp != null ? `Promedio anual (años completos): ${fmt(promComp)}` : 'Solo año en curso'),
    kpi('Último periodo de grado', lastVal, lastTxt),
    kpi('Mujeres graduadas', fmtP(g(gen, 'Femenino') / tot), `${fmt(g(gen, 'Femenino'))} mujeres · ${fmt(g(gen, 'Masculino'))} hombres`),
    kpi('Año con más graduados', best ? String(best) : '—', best ? `${fmt(g(byY, best))} graduados · posgrado: ${fmtP(posg(niv) / tot)} de la selección` : 'Sin años completos seleccionados'),
  ].join('');

  // Anual apilado por semestre
  const labels = Y.map(y => { const d = yoy(y); return [ylab(y), d ? fmtPS(d.v, 0) : '']; });
  const sems = [1, 2].filter(s => state.sems.has(s));
  ensureChart('c-res-anio', {
    type: 'bar',
    data: { labels, datasets: sems.map(s => ({ label: `Semestre ${s}`, data: Y.map(y => g2(ys, y, s)), backgroundColor: Y.map(y => colFor(y, SEM_COLORS[s])), borderColor: '#fff', borderWidth: { top: 1 }, borderRadius: 2, stack: 's', inside: true, labelColor: '#fff' })) },
    options: {
      responsive: true, maintainAspectRatio: false, layout: { padding: { top: 22 } },
      plugins: {
        legend: legendOpts(), stackTotals: { enabled: true, size: Y.length > 16 ? 10 : 11.5 }, inlineLabels: { enabled: Y.length <= 12, size: 10, minInside: 1 },
        tooltip: { mode: 'index', intersect: false, callbacks: { title: c => ylab(Y[c[0].dataIndex]), label: c => ` ${c.dataset.label}: ${fmt(c.parsed.y)}`, footer: c => { const y = Y[c[0].dataIndex]; const d = yoy(y); return `Total: ${fmt(g(byY, y))}` + (d ? `\nVariación (${d.note}): ${fmtPS(d.v)}` : ''); } } }
      },
      scales: { x: { stacked: true, grid: { display: false }, ticks: { ...xTick, font: { size: Y.length > 16 ? 10 : 11.5 } } }, y: { stacked: true, beginAtZero: true, grid: baseGrid, ticks: tick } }
    }
  });

  // Serie semestral
  const pl = per.map(([a, s]) => `${a}-${s}`), pv = per.map(([a, s]) => g2(perRows, a, s));
  ensureChart('c-res-per', {
    type: 'line',
    data: { labels: pl, datasets: [{ label: 'Graduados', data: pv, borderColor: COL.teal, backgroundColor: alpha(COL.teal, 0.12), fill: true, borderWidth: 2.5, pointRadius: pl.length > 20 ? 2 : 3.5, pointBackgroundColor: COL.teal, tension: .25, labelPos: 'top' }] },
    options: {
      responsive: true, maintainAspectRatio: false, layout: { padding: { top: 16 } },
      plugins: { legend: { display: false }, inlineLabels: { enabled: pl.length <= 14, size: 10 }, tooltip: { callbacks: { label: c => ` Graduados: ${fmt(c.parsed.y)}` } } },
      scales: { x: { grid: { display: false }, ticks: { ...tick, maxRotation: 60, autoSkip: true } }, y: { beginAtZero: true, grid: baseGrid, ticks: tick } }
    }
  });

  // Nivel por año
  const yn = cross(rows, 'anio', 'nivel');
  const niveles = DATA.dims.nivel.filter(n => g(niv, n) > 0);
  ensureChart('c-res-niv', {
    type: 'bar',
    data: { labels: Y.map(ylab), datasets: niveles.map(n => ({ label: n, data: Y.map(y => g2(yn, y, n)), backgroundColor: Y.map(y => colFor(y, NIVEL_COLORS[n])), borderRadius: 2, stack: 's' })) },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: legendOpts(), tooltip: { mode: 'index', intersect: false, filter: i => i.parsed.y > 0, callbacks: { label: c => ` ${c.dataset.label}: ${fmt(c.parsed.y)}` } } },
      scales: { x: { stacked: true, grid: { display: false }, ticks: { ...tick, maxRotation: 60 } }, y: { stacked: true, beginAtZero: true, grid: baseGrid, ticks: tick } }
    }
  });

  // Lectura rápida
  const ins = [];
  const cuR = ranked(by(rows, 'cu')), prR = ranked(by(rows, 'programa'));
  ins.push(`En ${rangeText(Y)} se registran <b>${fmt(tot)}</b> graduados con los filtros aplicados; <b>${fmtP(g(niv, 'Pregrado') / tot)}</b> son de pregrado y <b>${fmtP(posg(niv) / tot)}</b> de posgrado (especialización y maestría).`);
  if (cuR.length > 1) ins.push(`<b>${esc(cuR[0][0])}</b> aporta el ${fmtP(cuR[0][1] / tot)} de los graduados y <b>${esc(cuR[1][0])}</b> el ${fmtP(cuR[1][1] / tot)}; entre los dos concentran el <b>${fmtP((cuR[0][1] + cuR[1][1]) / tot)}</b>.`);
  if (prR.length > 1) ins.push(`El programa con más graduados es <b>${esc(prR[0][0])}</b> (${fmt(prR[0][1])}), seguido de ${esc(prR[1][0])} (${fmt(prR[1][1])}).`);
  if (comp.length >= 3) {
    const worst = [...comp].filter(y => g(byY, y) >= 0.1 * g(byY, best)).sort((a, b) => g(byY, a) - g(byY, b))[0];
    ins.push(`Entre los años completos seleccionados, el máximo de grados se dio en <b>${best}</b> (${fmt(g(byY, best))})` + (worst && worst !== best ? ` y el mínimo${worst !== comp[0] ? '' : ''} en <b>${worst}</b> (${fmt(g(byY, worst))})${comp.some(y => g(byY, y) < 0.1 * g(byY, best)) ? ', sin contar los primeros años con muy pocos registros' : ''}` : '') + '.');
    const last3 = comp.slice(-3);
    if (last3.length === 3 && last3.includes(best) === false) {
      ins.push(`Los últimos tres años completos (${last3.join(', ')}) promedian <b>${fmt(sum(last3.map(y => g(byY, y))) / 3)}</b> graduados por año, ${fmtPS(sum(last3.map(y => g(byY, y))) / 3 / g(byY, best) - 1)} frente al año pico (${best}).`);
    }
  }
  const s1 = sum(Y.map(y => g2(ys, y, 1))), s2 = sum(Y.map(y => g2(ys, y, 2)));
  if (state.sems.size === 2 && s1 + s2) ins.push(`El <b>${fmtP(s1 / (s1 + s2))}</b> de los grados de la selección ocurre en el primer semestre del año.`);
  Y.filter(isPartial).forEach(y => {
    const d = yoy(y); if (d) ins.push(`<b>${y}-1</b> (corte ${fmtCorte()}): ${fmt(g2(ys, y, 1))} graduados, ${gapSpan(d.v)} frente a ${y - 1}-1. Los datos de ${y}-2 aún no están disponibles.`);
  });
  el('ins-res').innerHTML = ins.map(t => `<li>${t}</li>`).join('');

  // Tabla anual
  const yg = cross(rows, 'anio', 'genero');
  let h = `<thead><tr><th>Año</th><th class="r">Semestre 1</th><th class="r">Semestre 2</th><th class="r">Total</th><th class="r">Var. anual</th><th class="r">% mujeres</th><th class="r">% posgrado</th><th class="r">Participación</th></tr></thead><tbody>`;
  [...Y].reverse().forEach(y => {
    const t = g(byY, y); const d = yoy(y);
    h += `<tr><td class="name">${ylab(y)}${pTag(y)}</td><td class="r cell-sc">${fmt(g2(ys, y, 1))}</td><td class="r cell-sc">${isPartial(y) ? '<span class="pending">pendiente</span>' : fmt(g2(ys, y, 2))}</td><td class="r cell-sc">${fmt(t)}</td><td class="r">${d ? gapSpan(d.v, d.note) : '<span class="flat">—</span>'}</td><td class="r">${t ? fmtP(g2(yg, y, 'Femenino') / t) : '—'}</td><td class="r">${t ? fmtP(posg2(yn, y) / t) : '—'}</td><td class="r"><div class="barcell"><span>${fmtP(t / tot)}</span><span class="bt"><span class="bf" style="width:${(t / Math.max(...Y.map(z => g(byY, z))) * 100).toFixed(1)}%;display:block"></span></span></div></td></tr>`;
  });
  h += `<tr class="total"><td>Total selección</td><td class="r">${fmt(s1)}</td><td class="r">${fmt(s2)}</td><td class="r">${fmt(tot)}</td><td></td><td class="r">${fmtP(g(gen, 'Femenino') / tot)}</td><td class="r">${fmtP(posg(niv) / tot)}</td><td class="r">100%</td></tr>`;
  el('t-res').innerHTML = h + '</tbody>';
}

/* ================================================================
   CENTROS UNIVERSITARIOS
================================================================ */
function renderCentros(rows) {
  const Y = yrs();
  const tot = total(rows);
  const cu = by(rows, 'cu'), R = ranked(cu);
  const sedeTot = total(filtered('cu'));
  const cy = cross(rows, 'cu', 'anio');
  const comp = Y.filter(y => !isPartial(y));
  const lastC = comp[comp.length - 1], prevC = comp[comp.length - 2];
  const growthCU = lastC && prevC ? R.map(([c]) => ({ c, v: g2(cy, c, prevC) ? g2(cy, c, lastC) / g2(cy, c, prevC) - 1 : null })).filter(x => x.v != null).sort((a, b) => b.v - a.v) : [];

  el('kpi-cu').innerHTML = [
    kpi('Graduados de los centros seleccionados', fmt(tot), state.f.cu.size ? `${fmtP(tot / sedeTot)} de la sede (${fmt(sedeTot)})` : `Total de la sede en ${rangeText(Y)}`),
    kpi('Centro con más graduados', R.length ? esc(R[0][0]) : '—', R.length ? `${fmt(R[0][1])} graduados · ${fmtP(R[0][1] / tot)} de la selección` : ''),
    kpi('Centros con graduados', `${R.length}<small> de ${DATA.dims.cu.length}</small>`, 'Con al menos un graduado en la selección'),
    kpi(lastC && prevC ? `Mayor crecimiento ${prevC}–${lastC}` : 'Mayor crecimiento anual', growthCU.length ? esc(growthCU[0].c) : '—', growthCU.length ? `${gapSpan(growthCU[0].v)} · ${fmt(g2(cy, growthCU[0].c, prevC))} → ${fmt(g2(cy, growthCU[0].c, lastC))}` : 'Se necesitan dos años completos'),
  ].join('');

  ensureChart('c-cu-stack', {
    type: 'bar',
    data: { labels: Y.map(ylab), datasets: R.map(([c]) => ({ label: c, data: Y.map(y => g2(cy, c, y)), backgroundColor: Y.map(y => colFor(y, CU_COLORS[c] || COL.nat)), borderColor: '#fff', borderWidth: { top: 1 }, stack: 's' })) },
    options: {
      responsive: true, maintainAspectRatio: false, layout: { padding: { top: 22 } },
      plugins: { legend: legendOpts(), stackTotals: { enabled: true, size: Y.length > 16 ? 10 : 11.5 }, tooltip: { mode: 'index', intersect: false, filter: i => i.parsed.y > 0, callbacks: { label: c => ` ${c.dataset.label}: ${fmt(c.parsed.y)}` } } },
      scales: { x: { stacked: true, grid: { display: false }, ticks: { ...xTick, font: { size: Y.length > 16 ? 10 : 11.5 } } }, y: { stacked: true, beginAtZero: true, grid: baseGrid, ticks: tick } }
    }
  });
  ensureChart('c-cu-bar', {
    type: 'bar',
    data: { labels: R.map(r => r[0]), datasets: [{ label: 'Graduados', data: R.map(r => r[1]), backgroundColor: R.map(r => CU_COLORS[r[0]] || COL.nat), borderRadius: 2 }] },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false, layout: { padding: { right: 48 } },
      plugins: { legend: { display: false }, inlineLabels: { enabled: true }, tooltip: { callbacks: { label: c => ` Graduados: ${fmt(c.parsed.x)} (${fmtP(c.parsed.x / tot)})` } } },
      scales: { x: { beginAtZero: true, grid: baseGrid, ticks: tick }, y: { grid: { display: false }, ticks: xTick } }
    }
  });
  ensureChart('c-cu-pie', {
    type: 'doughnut',
    data: { labels: R.map(r => r[0]), datasets: [{ data: R.map(r => r[1]), backgroundColor: R.map(r => CU_COLORS[r[0]] || COL.nat), borderColor: '#fff', borderWidth: 2 }] },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '52%',
      plugins: { legend: legendOpts('right'), inlineLabels: { enabled: true, size: 11 }, tooltip: { callbacks: { label: c => ` ${c.label}: ${fmt(c.parsed)} (${fmtP(c.parsed / sum(c.dataset.data))})` } } }
    }
  });

  // Nivel por CU (100 %)
  const cn = cross(rows, 'cu', 'nivel');
  const niveles = DATA.dims.nivel.filter(n => R.some(([c]) => g2(cn, c, n) > 0));
  ensureChart('c-cu-niv', {
    type: 'bar',
    data: { labels: R.map(r => r[0]), datasets: niveles.map(n => ({ label: n, data: R.map(([c, t]) => g2(cn, c, n) / t * 100), raw: R.map(([c]) => g2(cn, c, n)), backgroundColor: NIVEL_COLORS[n], stack: 's', inside: true, labelFmt: v => fmt(v, 0) + '%' })) },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: legendOpts(), inlineLabels: { enabled: true, size: 10.5, minInside: 4 }, tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${fmt(c.parsed.x, 1)}% (${fmt(c.dataset.raw[c.dataIndex])})` } } },
      scales: { x: { stacked: true, max: 100, grid: baseGrid, ticks: { ...tick, callback: v => v + '%' } }, y: { stacked: true, grid: { display: false }, ticks: xTick } }
    }
  });

  // Mapa de calor
  const max = Math.max(1, ...R.flatMap(([c]) => Y.map(y => g2(cy, c, y))));
  const heat = v => { if (!v) return 'color:#B8BEC6'; const a = 0.1 + 0.75 * (v / max); return `background:rgba(47,122,109,${a.toFixed(2)});color:${a > 0.5 ? '#fff' : '#16324F'}`; };
  let h = `<thead><tr><th>Centro Universitario</th>${Y.map(y => `<th class="r">${ylab(y)}</th>`).join('')}<th class="r">Total</th><th class="r">Participación</th></tr></thead><tbody>`;
  R.forEach(([c, t]) => {
    h += `<tr><td class="name">${esc(c)}</td>${Y.map(y => { const v = g2(cy, c, y); return `<td class="heat" style="${heat(v)}">${v ? fmt(v) : '–'}</td>`; }).join('')}<td class="r cell-sc">${fmt(t)}</td><td class="r"><div class="barcell"><span>${fmtP(t / tot)}</span><span class="bt"><span class="bf" style="width:${(t / tot * 100).toFixed(1)}%;display:block;background:${CU_COLORS[c] || COL.nat}"></span></span></div></td></tr>`;
  });
  const byY = by(rows, 'anio');
  h += `<tr class="total"><td>Total</td>${Y.map(y => `<td class="r">${fmt(g(byY, y))}</td>`).join('')}<td class="r">${fmt(tot)}</td><td class="r">100%</td></tr>`;
  el('t-cu').innerHTML = h + '</tbody>';
}

/* ================================================================
   PROGRAMAS
================================================================ */
function renderProgramas(rows) {
  const Y = yrs();
  const tot = total(rows);
  const pr = by(rows, 'programa'), R = ranked(pr);
  const meta = p => DATA.prog_meta[p] || { nivel: 'Sin dato', area: 'Sin dato' };
  const top5 = sum(R.slice(0, 5).map(r => r[1]));
  const niv = by(rows, 'nivel');
  const posgR = R.filter(([p]) => POSGRADO.includes(meta(p).nivel));

  el('kpi-prog').innerHTML = [
    kpi('Programas con graduados', fmt(R.length), `${fmt(R.length - posgR.length)} de pregrado y técnicos · ${fmt(posgR.length)} de posgrado`),
    kpi('Programa con más graduados', R.length ? `<span style="font-size:18px">${esc(short(abrev(R[0][0]), 34))}</span>` : '—', R.length ? `${fmt(R[0][1])} graduados · ${fmtP(R[0][1] / tot)} de la selección` : ''),
    kpi('Concentración en los 5 primeros', fmtP(top5 / tot), `${fmt(top5)} de ${fmt(tot)} graduados`),
    kpi('Posgrado con más graduados', posgR.length ? `<span style="font-size:18px">${esc(short(abrev(posgR[0][0]), 34))}</span>` : '—', posgR.length ? `${fmt(posgR[0][1])} graduados · posgrado total: ${fmtP(posg(niv) / tot)}` : 'Sin posgrados en la selección'),
  ].join('');

  const T = R.slice(0, 15);
  el('c-prog-top-sub').textContent = `${T.length < R.length ? `Los ${T.length} primeros de ${R.length}` : `${R.length} ${R.length === 1 ? 'programa' : 'programas'}`} · color según nivel de formación`;
  ensureChart('c-prog-top', {
    type: 'bar',
    data: { labels: T.map(r => short(abrev(r[0]), 34)), datasets: [{ label: 'Graduados', data: T.map(r => r[1]), backgroundColor: T.map(r => NIVEL_COLORS[meta(r[0]).nivel] || COL.nat), borderRadius: 2 }] },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false, layout: { padding: { right: 44 } },
      plugins: { legend: { display: false }, inlineLabels: { enabled: true, size: 10.5 }, tooltip: { callbacks: { title: c => T[c[0].dataIndex][0], label: c => ` ${fmt(c.parsed.x)} graduados (${fmtP(c.parsed.x / tot)})`, afterLabel: c => ` ${meta(T[c.dataIndex][0]).nivel}` } } },
      scales: { x: { beginAtZero: true, grid: baseGrid, ticks: tick }, y: { grid: { display: false }, ticks: { font: { size: 10.5 }, color: COL.ink } } }
    }
  });

  const ar = ranked(by(rows, 'area'));
  ensureChart('c-prog-area', {
    type: 'doughnut',
    data: { labels: ar.map(r => r[0]), datasets: [{ data: ar.map(r => r[1]), backgroundColor: ar.map((_, i) => AREA_COLORS[i % AREA_COLORS.length]), borderColor: '#fff', borderWidth: 2 }] },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '50%',
      plugins: { legend: legendOpts('bottom'), inlineLabels: { enabled: true, size: 11 }, tooltip: { callbacks: { label: c => ` ${c.label}: ${fmt(c.parsed)} (${fmtP(c.parsed / sum(c.dataset.data))})` } } }
    }
  });

  const py = cross(rows, 'programa', 'anio');
  const pal = [COL.ink, COL.accent, COL.teal, COL.purple, '#3E6FA8'];
  ensureChart('c-prog-evo', {
    type: 'line',
    data: { labels: Y.map(ylab), datasets: R.slice(0, 5).map(([p], i) => ({ label: short(p, 48), data: Y.map(y => g2(py, p, y)), borderColor: pal[i], backgroundColor: pal[i], borderWidth: 2.2, pointRadius: 3, tension: .25, segment: { borderDash: c => isPartial(Y[c.p1DataIndex]) ? [5, 4] : undefined } })) },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: legendOpts(), tooltip: { mode: 'index', intersect: false, callbacks: { title: c => ylab(Y[c[0].dataIndex]), label: c => ` ${c.dataset.label}: ${fmt(c.parsed.y)}` } } },
      scales: { x: { grid: { display: false }, ticks: { ...tick, maxRotation: 60 } }, y: { beginAtZero: true, grid: baseGrid, ticks: tick } }
    }
  });

  // Lectura rápida
  const comp = Y.filter(y => !isPartial(y));
  const lastC = comp[comp.length - 1], prevC = comp[comp.length - 2];
  const ins = [];
  ins.push(R.length === 1 ? `Un solo programa registra graduados en la selección.` : `<b>${fmt(R.length)}</b> programas registran graduados en la selección; los cinco primeros concentran el <b>${fmtP(top5 / tot)}</b> del total.`);
  if (ar.length) ins.push(`El área de <b>${esc(ar[0][0])}</b> reúne el ${fmtP(ar[0][1] / tot)} de los graduados${ar[1] ? `, seguida de ${esc(ar[1][0])} (${fmtP(ar[1][1] / tot)})` : ''}.`);
  if (lastC && prevC) {
    const ch = R.map(([p]) => ({ p, a: g2(py, p, prevC), b: g2(py, p, lastC) })).filter(x => x.a >= 20);
    const up = [...ch].sort((x, y) => (y.b - y.a) - (x.b - x.a))[0], dn = [...ch].sort((x, y) => (x.b - x.a) - (y.b - y.a))[0];
    if (up && up.b > up.a) ins.push(`Mayor aumento de graduados entre ${prevC} y ${lastC}: <b>${esc(up.p)}</b> (${fmt(up.a)} → ${fmt(up.b)}).`);
    if (dn && dn.b < dn.a) ins.push(`Mayor disminución entre ${prevC} y ${lastC}: <b>${esc(dn.p)}</b> (${fmt(dn.a)} → ${fmt(dn.b)}).`);
    const sinRecientes = R.filter(([p]) => sum(comp.slice(-3).map(y => g2(py, p, y))) === 0 && !Y.filter(isPartial).some(y => g2(py, p, y)));
    if (sinRecientes.length && comp.length >= 3) ins.push(`<b>${sinRecientes.length}</b> programas con graduados en la selección no registran grados en los últimos tres años completos (${comp.slice(-3).join(', ')}), lo que sugiere programas cerrados o sin cohortes activas.`);
  }
  el('ins-prog').innerHTML = ins.map(t => `<li>${t}</li>`).join('');

  // Tabla
  const YC = Y.slice(-8);
  el('t-prog-sub').textContent = `Programas ordenados por número de graduados · ${YC.length < Y.length ? `se muestran los últimos ${YC.length} años seleccionados; el total suma todos los años de la selección` : 'años seleccionados'}`;
  const max = Math.max(1, ...R.flatMap(([p]) => YC.map(y => g2(py, p, y))));
  const heat = v => { if (!v) return 'color:#B8BEC6'; const a = 0.1 + 0.75 * Math.sqrt(v / max); return `background:rgba(22,50,79,${a.toFixed(2)});color:${a > 0.5 ? '#fff' : '#16324F'}`; };
  let h = `<thead><tr><th>Programa</th>${YC.map(y => `<th class="r">${ylab(y)}</th>`).join('')}<th class="r">Total</th><th class="r">Participación</th></tr></thead><tbody>`;
  R.forEach(([p, t]) => {
    const m = meta(p);
    h += `<tr><td class="name wrap"><span class="lvl" style="background:${NIVEL_COLORS[m.nivel] || COL.nat}"></span>${esc(p)}<small>${esc(m.nivel)} · ${esc(m.area)}</small></td>${YC.map(y => { const v = g2(py, p, y); return `<td class="heat" style="${heat(v)}">${v ? fmt(v) : '–'}</td>`; }).join('')}<td class="r cell-sc">${fmt(t)}</td><td class="r"><div class="barcell"><span>${fmtP(t / tot)}</span><span class="bt"><span class="bf" style="width:${Math.min(100, t / R[0][1] * 100).toFixed(1)}%;display:block"></span></span></div></td></tr>`;
  });
  const byY = by(rows, 'anio');
  h += `<tr class="total"><td>Total</td>${YC.map(y => `<td class="r">${fmt(g(byY, y))}</td>`).join('')}<td class="r">${fmt(tot)}</td><td class="r">100%</td></tr>`;
  el('t-prog').innerHTML = h + '</tbody>';
}

/* ================================================================
   PERFIL DEL GRADUADO
================================================================ */
function renderPerfil(rows) {
  const Y = yrs();
  const tot = total(rows);
  const gen = by(rows, 'genero'), edad = by(rows, 'edad'), est = by(rows, 'estrato');
  const conEst = sum(DATA.dims.estrato.filter(e => e !== 'Sin dato').map(e => g(est, e)));
  const e12 = g(est, 'Estrato 1') + g(est, 'Estrato 2');
  const e3p = conEst - e12;
  const edR = ranked(edad).filter(([k]) => k !== 'Sin dato');
  const joven = g(edad, '18 a 30');

  el('kpi-per').innerHTML = [
    kpi('Mujeres graduadas', fmtP(g(gen, 'Femenino') / tot), `${fmt(g(gen, 'Femenino'))} de ${fmt(tot)} graduados`, COL.teal),
    kpi('Rango de edad predominante', edR.length ? edR[0][0] : '—', edR.length ? `${fmtP(edR[0][1] / tot)} de los graduados · mayores de 30: ${fmtP((tot - joven - g(edad, 'Sin dato')) / tot)}` : ''),
    kpi('Graduados de estratos 1 y 2', fmtP(e12 / (conEst || 1)), `${fmt(e12)} graduados con estrato registrado`),
    kpi('Graduados de estrato 3 o superior', fmtP(e3p / (conEst || 1)), `${fmt(e3p)} graduados`),
  ].join('');

  const G = ranked(gen);
  ensureChart('c-per-gen', {
    type: 'doughnut',
    data: { labels: G.map(r => r[0]), datasets: [{ data: G.map(r => r[1]), backgroundColor: G.map(r => GEN_COLORS[r[0]] || COL.nat), borderColor: '#fff', borderWidth: 2 }] },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '55%',
      plugins: { legend: legendOpts('right'), inlineLabels: { enabled: true, size: 12 }, tooltip: { callbacks: { label: c => ` ${c.label}: ${fmt(c.parsed)} (${fmtP(c.parsed / sum(c.dataset.data))})` } } }
    }
  });

  const yg = cross(rows, 'anio', 'genero'), byY = by(rows, 'anio');
  const fem = Y.map(y => g(byY, y) >= 10 ? g2(yg, y, 'Femenino') / g(byY, y) * 100 : null);
  ensureChart('c-per-fem', {
    type: 'line',
    data: { labels: Y.map(ylab), datasets: [{ label: '% mujeres', data: fem, borderColor: COL.teal, backgroundColor: COL.teal, borderWidth: 2.5, pointRadius: 3.5, tension: .25, spanGaps: true, labelPos: 'top' }] },
    options: {
      responsive: true, maintainAspectRatio: false, layout: { padding: { top: 16 } },
      plugins: { legend: { display: false }, inlineLabels: { enabled: Y.length <= 12, size: 10, fmt: v => fmt(v, 0) + '%' }, tooltip: { callbacks: { title: c => ylab(Y[c[0].dataIndex]), label: c => ` Mujeres: ${fmt(c.parsed.y, 1)}% de ${fmt(g(byY, Y[c.dataIndex]))} graduados` } } },
      scales: { x: { grid: { display: false }, ticks: { ...tick, maxRotation: 60 } }, y: { min: 0, max: 100, grid: baseGrid, ticks: { ...tick, callback: v => v + '%' } } }
    }
  });

  const eg = cross(rows, 'edad', 'genero');
  const E = DATA.dims.edad.filter(e => g(edad, e) > 0 && e !== 'Sin dato');
  ensureChart('c-per-edad', {
    type: 'bar',
    data: { labels: E, datasets: ['Femenino', 'Masculino'].filter(s => g(gen, s)).map(s => ({ label: s, data: E.map(e => g2(eg, e, s)), backgroundColor: GEN_COLORS[s], borderRadius: 2 })) },
    options: {
      responsive: true, maintainAspectRatio: false, layout: { padding: { top: 18 } },
      plugins: { legend: legendOpts(), inlineLabels: { enabled: true, size: 10 }, tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${fmt(c.parsed.y)}` } } },
      scales: { x: { grid: { display: false }, ticks: xTick }, y: { beginAtZero: true, grid: baseGrid, ticks: tick } }
    }
  });

  const ES = DATA.dims.estrato.filter(e => g(est, e) > 0);
  ensureChart('c-per-est', {
    type: 'bar',
    data: { labels: ES, datasets: [{ label: 'Graduados', data: ES.map(e => g(est, e)), backgroundColor: ES.map(e => EST_COLORS[DATA.dims.estrato.indexOf(e)]), borderRadius: 2 }] },
    options: {
      responsive: true, maintainAspectRatio: false, layout: { padding: { top: 18 } },
      plugins: { legend: { display: false }, inlineLabels: { enabled: true, fmt: v => `${fmt(v)} · ${fmtP(v / tot, 0)}` }, tooltip: { callbacks: { label: c => ` ${fmt(c.parsed.y)} graduados (${fmtP(c.parsed.y / tot)})` } } },
      scales: { x: { grid: { display: false }, ticks: xTick }, y: { beginAtZero: true, grid: baseGrid, ticks: tick } }
    }
  });

  const CR = ranked(by(rows, 'cu'));
  const ce = cross(rows, 'cu', 'estrato');
  ensureChart('c-per-estcu', {
    type: 'bar',
    data: { labels: CR.map(r => r[0]), datasets: ES.map(e => ({ label: e, data: CR.map(([c, t]) => g2(ce, c, e) / t * 100), raw: CR.map(([c]) => g2(ce, c, e)), backgroundColor: EST_COLORS[DATA.dims.estrato.indexOf(e)], stack: 's', inside: true, labelFmt: v => fmt(v, 0) + '%' })) },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: legendOpts(), inlineLabels: { enabled: true, size: 10.5, minInside: 5 }, tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${fmt(c.parsed.x, 1)}% (${fmt(c.dataset.raw[c.dataIndex])})` } } },
      scales: { x: { stacked: true, max: 100, grid: baseGrid, ticks: { ...tick, callback: v => v + '%' } }, y: { stacked: true, grid: { display: false }, ticks: xTick } }
    }
  });

  // Lectura rápida
  const ins = [];
  ins.push(`<b>${fmtP(g(gen, 'Femenino') / tot)}</b> de los graduados de la selección son mujeres: por cada graduado hay <b>${fmt(g(gen, 'Femenino') / Math.max(1, g(gen, 'Masculino')), 1)}</b> graduadas.`);
  const validF = Y.map((y, i) => [y, fem[i]]).filter(([y, v]) => v != null && g(byY, y) >= 100);
  if (validF.length >= 2) { const [a, fa] = validF[0], [b, fb] = validF[validF.length - 1]; ins.push(`La participación femenina pasó de <b>${fmt(fa, 1)}%</b> en ${ylab(a)} a <b>${fmt(fb, 1)}%</b> en ${ylab(b)}.`); }
  ins.push(`El <b>${fmtP(joven / tot)}</b> de los graduados está en el rango de 18 a 30 años; el <b>${fmtP((g(edad, '41 a 45') + g(edad, 'Más de 45')) / tot)}</b> supera los 40.`);
  ins.push(`El <b>${fmtP(e12 / (conEst || 1))}</b> de los graduados con estrato registrado pertenece a los estratos 1 y 2.`);
  if (CR.length > 1) {
    const s12 = CR.map(([c, t]) => ({ c, v: (g2(ce, c, 'Estrato 1') + g2(ce, c, 'Estrato 2')) / t })).sort((a, b) => b.v - a.v);
    ins.push(`<b>${esc(s12[0].c)}</b> tiene la mayor proporción de graduados de estratos 1 y 2 (${fmtP(s12[0].v)}) y <b>${esc(s12[s12.length - 1].c)}</b> la menor (${fmtP(s12[s12.length - 1].v)}).`);
  }
  el('ins-per').innerHTML = ins.map(t => `<li>${t}</li>`).join('');
}

/* ================================================================
   METODOLOGÍA
================================================================ */
function renderMetodologia() {
  const q = DATA.calidad;
  const excl = Object.entries(q.excluidos_por_centro).map(([k, v]) => `${esc(k)} (${fmt(v)})`).join(', ');
  const p = DATA.anio_parcial[0];
  el('prose').innerHTML = `
    <h4>Fuente</h4>
    <p>Listado de graduados exportado de SAP (<b>${esc(DATA.fuente)}</b>), hoja «Listado Estudiantes», con corte al <b>${fmtCorte()}</b>. Cada registro corresponde a un grado e incluye Centro Universitario, programa, nivel de formación, modalidad, área de conocimiento, año y semestre de grado, género, rango de edad y estrato socioeconómico.</p>
    <h4>Cobertura</h4>
    <ul>
      <li>Periodos de grado desde <b>${DATA.periodos[0]}</b> hasta <b>${DATA.periodos[DATA.periodos.length - 1]}</b>. Los años anteriores a 2014 tienen pocos registros en SAP; no hay grados registrados en 2007 y 2008.</li>
      ${p ? `<li><b>${p} es un año parcial</b>: solo contiene el periodo ${p}-1. En el tablero aparece como «${p}*», con color atenuado, y su variación se calcula frente a ${p - 1}-1 (mismo semestre) para que la comparación sea justa. Los datos de ${p}-2 se incorporarán cuando estén disponibles.</li>` : ''}
      <li>Centros Universitarios incluidos: ${DATA.dims.cu.map(esc).join(', ')}.</li>
    </ul>
    <h4>Depuración aplicada</h4>
    <ul>
      <li>Registros en el archivo: <b>${fmt(q.registros_archivo)}</b>. Registros incluidos en el tablero: <b>${fmt(q.registros_tablero)}</b>.</li>
      <li>Se excluyen los centros que no hacen parte del alcance de la Sede Tolima-Huila: ${excl}.</li>
      ${q.registros_vacios ? `<li>Se descartan ${fmt(q.registros_vacios)} filas vacías del final del archivo.</li>` : ''}
      <li>Se unifican denominaciones de programa que SAP registra con distinta escritura (por ejemplo «Trabajo social» y «Trabajo Social», o «Especialización En Gerencia De Proyectos» y «Especialización en Gerencia de Proyectos»). También se normalizan tildes en niveles y modalidades.</li>
    </ul>
    <h4>Cómo leer los filtros</h4>
    <ul>
      <li>Todos los filtros se combinan entre sí. Dentro de un mismo filtro, elegir varias opciones suma sus graduados; dejarlo vacío equivale a «todos».</li>
      <li>El número junto a cada opción indica cuántos graduados quedarían al seleccionarla con el resto de filtros activos. Las opciones en gris no tienen graduados en la combinación actual.</li>
      <li>El rango de edad es el que reporta SAP en el listado de graduados.</li>
    </ul>
    <h4>Privacidad</h4>
    <p>El archivo publicado (<code>data.json</code>) solo contiene conteos agregados por combinación de variables. No incluye nombres, documentos, códigos ni ningún otro dato que identifique a una persona.</p>
    <h4>Actualización</h4>
    <p>Para incorporar un nuevo corte (por ejemplo, ${p ? `${p}-2` : 'el siguiente periodo'}): ejecutar <code>python build_data.py Graduados_SAP_&lt;fecha&gt;.xlsx AAAA-MM-DD</code> y subir el nuevo <code>data.json</code> al repositorio. El tablero detecta automáticamente los periodos y marca como parcial un año que solo tenga el primer semestre.</p>`;
}

boot();
