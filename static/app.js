'use strict';
/* =====================================================================
   AutoBOQ 2D – on-screen takeoff
   Drawing is shown exactly as drawn. The user measures on it.
   World coordinates = drawing units. Metres = units × S.unitToM
   ===================================================================== */

const $ = (id) => document.getElementById(id);
const PALETTE = ['#e11d48', '#2563eb', '#16a34a', '#d97706', '#7c3aed', '#0891b2', '#db2777', '#65a30d', '#ea580c', '#4f46e5'];
const DIAS = [8, 10, 12, 16, 20, 25, 32];

const S = {
  dwg: null,            // drawing data from server
  unitToM: 1,           // metres per drawing unit
  unitLabel: '',
  ox: 0, oy: 0,         // local origin (bounds min) for precision
  view: { s: 1, tx: 0, ty: 0 },
  paths: [],            // [{c, stroke: Path2D}] cached
  fillPaths: [],
  img: null,
  snapGrid: null,
  tool: 'pan',
  pts: [],              // points of measurement in progress (world)
  cursor: null,         // {x,y,snapped}
  snapOn: true, orthoOn: false, shift: false,
  hoverText: null,
  groups: [],           // BOQ items
  activeGid: null,
  selMeas: null,        // {gid, mid}
  history: [],
  dsr: [], dsrByCode: {},
  steel: [],
};

// ===================================================================== init
window.addEventListener('DOMContentLoaded', async () => {
  initTabs();
  initToolbar();
  initCanvas();
  initSteel();
  await loadDsr();
  ['meta-project', 'meta-client', 'meta-dwg', 'meta-cont', 'meta-gst'].forEach(id => $(id).addEventListener('input', () => { renderBoq(); save(); }));
  restoreSession();
  const qs = new URLSearchParams(location.search);
  if (qs.get('sample')) openSample(qs.get('sample'), 0);
});

function initTabs() {
  document.querySelectorAll('.tab').forEach(b => b.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(x => x.classList.toggle('active', x === b));
    document.querySelectorAll('.tabpane').forEach(p => p.classList.toggle('active', p.id === 'tab-' + b.dataset.tab));
    if (b.dataset.tab === 'boq') renderBoq();
    if (b.dataset.tab === 'takeoff') { resizeCanvas(); draw(); }
  }));
}

async function loadDsr() {
  try {
    const r = await fetch('/api/dsr-library');
    S.dsr = (await r.json()).items;
  } catch (e) { S.dsr = []; }
  S.dsr.forEach(d => S.dsrByCode[d.code] = d);
  const sel = $('dsr-select');
  const cats = {};
  S.dsr.forEach(d => (cats[d.category] = cats[d.category] || []).push(d));
  let html = '<option value="CUSTOM">✎ Custom item (type your own)…</option>';
  for (const c of Object.keys(cats)) {
    html += `<optgroup label="${esc(c)}">`;
    for (const d of cats[c]) html += `<option value="${d.code}">${d.code} · ${esc(shortDesc(d.description))} (${d.unit})</option>`;
    html += '</optgroup>';
  }
  sel.innerHTML = html;
  sel.value = 'DSR-F01';
  sel.addEventListener('change', () => $('custom-fields').classList.toggle('hidden', sel.value !== 'CUSTOM'));
  $('btn-add-item').addEventListener('click', addItemFromPanel);
}

function shortDesc(s) { return s.length > 70 ? s.slice(0, 68) + '…' : s; }
function esc(s) { return String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function fmt(n, d = 2) { return (+n || 0).toLocaleString('en-IN', { minimumFractionDigits: d, maximumFractionDigits: d }); }

// ===================================================================== toolbar & file open
function initToolbar() {
  $('file-input').addEventListener('change', e => { const f = e.target.files[0]; if (f) openFile(f, 0); e.target.value = ''; });
  $('sample-select').addEventListener('change', e => { if (e.target.value) openSample(e.target.value, 0); });
  $('page-select').addEventListener('change', e => {
    const p = +e.target.value;
    if (S.lastFile) openFile(S.lastFile, p); else if (S.lastSample) openSample(S.lastSample, p);
  });
  $('unit-select').addEventListener('change', onUnitChange);
  document.querySelectorAll('.tool').forEach(b => b.addEventListener('click', () => setTool(b.dataset.tool)));
  $('tool-calibrate').addEventListener('click', () => setTool('calibrate'));
  $('btn-snap').addEventListener('click', () => { S.snapOn = !S.snapOn; $('btn-snap').classList.toggle('on', S.snapOn); });
  $('btn-ortho').addEventListener('click', () => { S.orthoOn = !S.orthoOn; $('btn-ortho').classList.toggle('on', S.orthoOn); });
  $('btn-fit').addEventListener('click', fitView);
  $('btn-export').addEventListener('click', () => exportDocument('excel'));
  if ($('btn-export-pdf')) $('btn-export-pdf').addEventListener('click', () => exportDocument('pdf'));
}

async function openFile(file, page) {
  S.lastFile = file; S.lastSample = null;
  const fd = new FormData();
  fd.append('file', file);
  fd.append('page', String(page || 0));
  await loadDrawing(() => fetch('/api/open', { method: 'POST', body: fd }));
}

async function openSample(key, page) {
  S.lastSample = key; S.lastFile = null;
  await loadDrawing(() => fetch(`/api/sample/${key}?page=${page || 0}`));
}

async function loadDrawing(fetcher) {
  $('loading').classList.remove('hidden');
  try {
    const r = await fetcher();
    const txt = await r.text();
    let data;
    try { data = JSON.parse(txt); } catch (e) { throw new Error('Server returned an invalid response: ' + txt.slice(0, 200)); }
    if (!r.ok) throw new Error(data.detail || ('Error ' + r.status));
    setDrawing(data);
  } catch (e) {
    alert('Could not open the drawing:\n\n' + e.message);
  } finally {
    $('loading').classList.add('hidden');
  }
}

function setDrawing(d) {
  const prevName = S.dwg ? S.dwg.filename : null;
  S.dwg = d;
  S.ox = d.bounds.minx; S.oy = d.bounds.miny;
  S.pts = []; S.hoverText = null;
  $('empty-msg').classList.add('hidden');
  $('file-name').textContent = `${d.filename}  ·  ${d.format}  ·  ${d.layout}`;
  if (!$('meta-dwg').value) $('meta-dwg').value = d.filename.replace(/\.[^.]+$/, '');

  // page selector
  const ps = $('page-select');
  if (d.page_count > 1) {
    ps.innerHTML = Array.from({ length: d.page_count }, (_, i) => `<option value="${i}">Page ${i + 1}</option>`).join('');
    ps.value = d.page; ps.classList.remove('hidden');
  } else ps.classList.add('hidden');

  buildUnitSelect(d);
  buildPaths(d);
  buildSnapGrid(d.snaps);

  S.img = null;
  if (d.image) {
    const im = new Image();
    im.onload = () => { S.img = im; draw(); };
    im.src = d.image.src;
  }

  // restore saved takeoff for this file, if any
  if (d.filename !== prevName) loadSavedFor(d.filename);

  $('st-stats').textContent = `${d.stats.polylines} lines · ${d.stats.fills} fills · ${d.stats.texts} texts · ${d.stats.snaps} snap points`;
  resizeCanvas();
  fitView();
  renderItems();
  setTool(S.tool);
}

// ===================================================================== units
const DXF_UNITS = [['mm', 0.001], ['cm', 0.01], ['m', 1], ['inch', 0.0254], ['feet', 0.3048]];
const PDF_SCALES = [20, 25, 50, 75, 100, 125, 150, 200, 250, 300, 400, 500, 750, 1000];
const PT_M = 0.0254 / 72;

function buildUnitSelect(d) {
  const sel = $('unit-select');
  let html = '';
  if (d.format === 'PDF') {
    for (const s of PDF_SCALES) html += `<option value="${PT_M * s}">Scale 1:${s}</option>`;
    html += '<option value="cal" disabled>Calibrated</option>';
    sel.innerHTML = html;
    sel.value = String(PT_M * 100);
    S.unitToM = PT_M * 100; S.unitLabel = 'PDF 1:100';
  } else {
    for (const [n, f] of DXF_UNITS) html += `<option value="${f}">Drawing units = ${n}</option>`;
    html += '<option value="cal" disabled>Calibrated</option>';
    sel.innerHTML = html;
    sel.value = String(d.unit_to_m);
    S.unitToM = d.unit_to_m; S.unitLabel = d.unit_name;
  }
  updateUnitStatus(d.unit_source);
}

function updateUnitStatus(src) {
  $('st-units').innerHTML = `Units: <b>${esc(S.unitLabel)}</b>` + (src ? ` <span class="muted">(${esc(src)})</span>` : '');
}

function onUnitChange() {
  const v = $('unit-select').value;
  if (v === 'cal') return;
  S.unitToM = parseFloat(v);
  S.unitLabel = $('unit-select').selectedOptions[0].textContent.replace('Drawing units = ', '');
  updateUnitStatus('set by you');
  recomputeAll();
}

function setCalibrated(factor) {
  S.unitToM = factor;
  S.unitLabel = `calibrated (1 unit = ${(factor * 1000).toPrecision(5)} mm)`;
  const sel = $('unit-select');
  const o = sel.querySelector('option[value="cal"]');
  o.disabled = false; o.textContent = 'Calibrated'; sel.value = 'cal'; o.disabled = true;
  updateUnitStatus('calibrated by 2 points');
  recomputeAll();
}

// ===================================================================== geometry caches
function buildPaths(d) {
  const ox = S.ox, oy = S.oy;
  S.paths = d.strokes.map(s => {
    const p = new Path2D();
    for (const a of s.p) {
      p.moveTo(a[0] - ox, a[1] - oy);
      for (let i = 2; i < a.length; i += 2) p.lineTo(a[i] - ox, a[i + 1] - oy);
    }
    return { c: displayColor(s.c), p };
  });
  S.fillPaths = d.fills.map(s => {
    const p = new Path2D();
    for (const a of s.p) {
      p.moveTo(a[0] - ox, a[1] - oy);
      for (let i = 2; i < a.length; i += 2) p.lineTo(a[i] - ox, a[i + 1] - oy);
      p.closePath();
    }
    return { c: displayColor(s.c), p };
  });
}

// CAD colours are designed for black screens: darken very light ones for a white sheet
function displayColor(hex) {
  const r = parseInt(hex.slice(1, 3), 16), g = parseInt(hex.slice(3, 5), 16), b = parseInt(hex.slice(5, 7), 16);
  const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  if (lum > 0.92) return '#111111';
  if (lum > 0.55) { const k = 0.55 / lum; return `rgb(${Math.round(r * k)},${Math.round(g * k)},${Math.round(b * k)})`; }
  return hex;
}

function buildSnapGrid(snaps) {
  const b = S.dwg.bounds;
  const span = Math.max(b.maxx - b.minx, b.maxy - b.miny, 1e-6);
  const cell = span / 256;
  const grid = new Map();
  for (let i = 0; i < snaps.length; i += 2) {
    const x = snaps[i], y = snaps[i + 1];
    const k = Math.floor(x / cell) + ',' + Math.floor(y / cell);
    let arr = grid.get(k); if (!arr) grid.set(k, arr = []);
    arr.push(x, y);
  }
  S.snapGrid = { cell, grid };
}

function findSnap(wx, wy, tolW) {
  if (!S.snapGrid) return null;
  const { cell, grid } = S.snapGrid;
  const r = Math.ceil(tolW / cell);
  const cx = Math.floor(wx / cell), cy = Math.floor(wy / cell);
  let best = null, bd = tolW;
  if (r > 6) return null;
  for (let i = cx - r; i <= cx + r; i++) for (let j = cy - r; j <= cy + r; j++) {
    const arr = grid.get(i + ',' + j); if (!arr) continue;
    for (let k = 0; k < arr.length; k += 2) {
      const d = Math.hypot(arr[k] - wx, arr[k + 1] - wy);
      if (d < bd) { bd = d; best = { x: arr[k], y: arr[k + 1] }; }
    }
  }
  return best;
}

// ===================================================================== canvas & view
let cv, ctx, dpr = 1;
let drag = null;

function initCanvas() {
  cv = $('cad'); ctx = cv.getContext('2d');
  new ResizeObserver(() => { resizeCanvas(); draw(); }).observe($('canvas-wrap'));

  cv.addEventListener('mousedown', e => {
    if (e.button === 1 || e.button === 0) drag = { x: e.clientX, y: e.clientY, tx: S.view.tx, ty: S.view.ty, moved: false, btn: e.button };
    if (e.button === 1) e.preventDefault();
  });
  window.addEventListener('mousemove', e => {
    if (drag) {
      const dx = e.clientX - drag.x, dy = e.clientY - drag.y;
      if (Math.hypot(dx, dy) > 4) drag.moved = true;
      if (drag.moved) { S.view.tx = drag.tx + dx; S.view.ty = drag.ty + dy; draw(); return; }
    }
    if (e.target !== cv || !S.dwg) return;
    S.shift = e.shiftKey;
    updateCursor(e);
  });
  window.addEventListener('mouseup', e => {
    if (!drag) return;
    const wasClick = !drag.moved && drag.btn === 0 && e.target === cv;
    drag = null;
    if (wasClick) onClick(e);
  });
  cv.addEventListener('dblclick', e => { e.preventDefault(); if (S.tool === 'line' || S.tool === 'poly') finishShape(true); });
  cv.addEventListener('contextmenu', e => { e.preventDefault(); finishShape(false); });
  cv.addEventListener('wheel', e => {
    e.preventDefault();
    if (!S.dwg) return;
    const r = cv.getBoundingClientRect(), mx = e.clientX - r.left, my = e.clientY - r.top;
    const f = e.deltaY < 0 ? 1.2 : 1 / 1.2;
    S.view.tx = mx - (mx - S.view.tx) * f;
    S.view.ty = my - (my - S.view.ty) * f;
    S.view.s *= f;
    updateCursor(e);
  }, { passive: false });

  window.addEventListener('keydown', e => {
    if (['INPUT', 'SELECT', 'TEXTAREA'].includes(document.activeElement.tagName)) return;
    if (!$('tab-takeoff').classList.contains('active')) return;
    const k = e.key.toLowerCase();
    if (e.key === 'Escape') { S.pts = []; setTool('pan'); }
    else if (e.key === 'Enter') finishShape(false);
    else if (e.key === 'Backspace') { e.preventDefault(); S.pts.pop(); draw(); }
    else if ((e.ctrlKey || e.metaKey) && k === 'z') { e.preventDefault(); undo(); }
    else if (k === 'l') setTool('line');
    else if (k === 'a') setTool('poly');
    else if (k === 'r') setTool('rect');
    else if (k === 'c') setTool('count');
    else if (k === 't') setTool('text');
    else if (k === 'f') fitView();
    else if (e.key === 'Shift') { S.shift = true; draw(); }
  });
  window.addEventListener('keyup', e => { if (e.key === 'Shift') { S.shift = false; draw(); } });
}

function resizeCanvas() {
  if (!cv) return;
  dpr = window.devicePixelRatio || 1;
  const w = cv.clientWidth, h = cv.clientHeight;
  if (cv.width !== Math.round(w * dpr) || cv.height !== Math.round(h * dpr)) {
    cv.width = Math.round(w * dpr); cv.height = Math.round(h * dpr);
  }
}

function w2s(x, y) { return { x: (x - S.ox) * S.view.s + S.view.tx, y: S.view.ty - (y - S.oy) * S.view.s }; }
function s2w(x, y) { return { x: (x - S.view.tx) / S.view.s + S.ox, y: (S.view.ty - y) / S.view.s + S.oy }; }

function fitView() {
  if (!S.dwg) return;
  const b = S.dwg.bounds, W = cv.clientWidth, H = cv.clientHeight, pad = 30;
  const bw = Math.max(b.maxx - b.minx, 1e-6), bh = Math.max(b.maxy - b.miny, 1e-6);
  S.view.s = Math.min((W - 2 * pad) / bw, (H - 2 * pad) / bh);
  S.view.tx = (W - bw * S.view.s) / 2;
  S.view.ty = H - (H - bh * S.view.s) / 2;
  draw();
}

function updateCursor(e) {
  const r = cv.getBoundingClientRect();
  let w = s2w(e.clientX - r.left, e.clientY - r.top);
  let snapped = false;
  if (S.snapOn && S.tool !== 'pan' && S.tool !== 'text') {
    const sp = findSnap(w.x, w.y, 9 / S.view.s);
    if (sp) { w = sp; snapped = true; }
  }
  if ((S.orthoOn || S.shift) && S.pts.length && (S.tool === 'line' || S.tool === 'poly' || S.tool === 'calibrate')) {
    const last = S.pts[S.pts.length - 1];
    if (Math.abs(w.x - last.x) > Math.abs(w.y - last.y)) w = { x: w.x, y: last.y }; else w = { x: last.x, y: w.y };
    snapped = false;
  }
  S.cursor = { x: w.x, y: w.y, snapped };
  const m = S.unitToM;
  $('st-coords').textContent = `X ${((w.x - S.ox) * m).toFixed(3)} m   Y ${((w.y - S.oy) * m).toFixed(3)} m`;
  if (S.tool === 'text') S.hoverText = hitText(w.x, w.y);
  draw();
}

// ===================================================================== tools
const HINTS = {
  pan: 'Drag to move the drawing, mouse wheel to zoom. Choose a tool to measure.',
  line: 'LENGTH: click points along the wall/kerb/pipe. Enter, right-click or double-click to finish. Shift = straight.',
  poly: 'AREA: click each corner of the room/zone. Enter, right-click or click the first point to close.',
  rect: 'RECTANGLE: click two opposite corners. Gives L × B.',
  count: 'COUNT: click each tree / door / column / light. Right-click to finish this count row.',
  text: 'PICK TEXT: click a text like "4.50 x 3.60" or "BEDROOM 12\'-0" x 10\'-0"" to add L × B.',
  calibrate: 'CALIBRATE: click the two ends of a dimension you know (e.g. a 4500 mm wall).',
};

function setTool(t) {
  S.tool = t; S.pts = []; S.hoverText = null;
  document.querySelectorAll('.tool').forEach(b => b.classList.toggle('active', b.dataset.tool === t));
  $('tool-calibrate').classList.toggle('primary', t === 'calibrate');
  $('st-hint').textContent = S.dwg ? HINTS[t] : 'Open a drawing to start.';
  cv.style.cursor = t === 'pan' ? 'grab' : 'crosshair';
  draw();
}

function needActiveItem() {
  if (activeGroup()) return true;
  alert('First choose the BOQ item on the right (e.g. "Vitrified flooring") and press "＋ Add as BOQ item".\nThen measure on the drawing.');
  return false;
}

function onClick(e) {
  if (!S.dwg || !S.cursor) return;
  const p = { x: S.cursor.x, y: S.cursor.y };
  switch (S.tool) {
    case 'pan': selectMeasAt(p); break;
    case 'line':
      if (!S.pts.length && !needActiveItem()) return;
      S.pts.push(p); break;
    case 'poly':
      if (!S.pts.length && !needActiveItem()) return;
      if (S.pts.length >= 3) {
        const f = w2s(S.pts[0].x, S.pts[0].y), c = w2s(p.x, p.y);
        if (Math.hypot(f.x - c.x, f.y - c.y) < 10) { finishShape(false); return; }
      }
      S.pts.push(p); break;
    case 'rect':
      if (!S.pts.length && !needActiveItem()) return;
      S.pts.push(p);
      if (S.pts.length === 2) finishShape(false);
      break;
    case 'count': {
      if (!needActiveItem()) return;
      const g = activeGroup();
      let m = g.meas[g.meas.length - 1];
      if (!m || m.geo.type !== 'count' || m.closed) {
        m = newMeas(g, { type: 'count', pts: [] }, 'Counted on drawing');
      }
      m.geo.pts.push([p.x, p.y]);
      recomputeMeas(g, m); afterChange(); break;
    }
    case 'text': pickText(p); break;
    case 'calibrate':
      S.pts.push(p);
      if (S.pts.length === 2) finishShape(false);
      break;
  }
  draw();
}

function finishShape(fromDbl) {
  const t = S.tool;
  if (t === 'count') {
    const g = activeGroup();
    if (g && g.meas.length) g.meas[g.meas.length - 1].closed = true;
    S.pts = []; draw(); return;
  }
  let pts = S.pts.slice();
  // double-click adds the same point twice – remove duplicates
  pts = pts.filter((p, i) => i === 0 || Math.hypot(p.x - pts[i - 1].x, p.y - pts[i - 1].y) > 1e-9);
  S.pts = [];
  const g = activeGroup();
  if (t === 'line' && pts.length >= 2 && g) {
    newMeas(g, { type: 'line', pts: pts.map(p => [p.x, p.y]) }, `Length ${g.meas.length + 1}`);
    afterChange();
  } else if (t === 'poly' && pts.length >= 3 && g) {
    newMeas(g, { type: 'poly', pts: pts.map(p => [p.x, p.y]) }, `Area ${g.meas.length + 1}`);
    afterChange();
  } else if (t === 'rect' && pts.length === 2 && g) {
    const [a, b] = pts;
    if (Math.abs(a.x - b.x) < 1e-9 || Math.abs(a.y - b.y) < 1e-9) { draw(); return; }
    const x0 = Math.min(a.x, b.x), x1 = Math.max(a.x, b.x), y0 = Math.min(a.y, b.y), y1 = Math.max(a.y, b.y);
    newMeas(g, { type: 'rect', pts: [[x0, y0], [x1, y0], [x1, y1], [x0, y1]] }, `Room ${g.meas.length + 1}`);
    afterChange();
  } else if (t === 'calibrate' && pts.length === 2) {
    const du = Math.hypot(pts[1].x - pts[0].x, pts[1].y - pts[0].y);
    if (du > 0) {
      const cur = du * S.unitToM;
      const ans = prompt(`The two points are ${cur.toFixed(3)} m apart at the current scale.\n\nEnter the TRUE length from the drawing (in metres, e.g. 4.5):`, cur.toFixed(3));
      const v = parseFloat(ans);
      if (v > 0) setCalibrated(v / du);
    }
    setTool('pan');
  }
  draw();
}

// ===================================================================== text pick
function textBox(t) {
  const len = t.t.length * t.h * 0.62;
  let x0 = t.a === 'c' ? -len / 2 : t.a === 'r' ? -len : 0;
  let y0 = t.b === 'm' ? -t.h / 2 : t.b === 't' ? -t.h : -t.h * 0.25;
  return { x0, y0, x1: x0 + len, y1: y0 + t.h * 1.25 };
}

function hitText(wx, wy) {
  const texts = S.dwg.texts; let best = null, bestArea = Infinity;
  const pad = 4 / S.view.s;
  for (const t of texts) {
    const r = -t.r * Math.PI / 180;
    const dx = wx - t.x, dy = wy - t.y;
    const lx = dx * Math.cos(r) - dy * Math.sin(r), ly = dx * Math.sin(r) + dy * Math.cos(r);
    const b = textBox(t);
    if (lx >= b.x0 - pad && lx <= b.x1 + pad && ly >= b.y0 - pad && ly <= b.y1 + pad) {
      const a = (b.x1 - b.x0) * (b.y1 - b.y0);
      if (a < bestArea) { bestArea = a; best = t; }
    }
  }
  return best;
}

function parseDims(s) {
  s = s.replace(/×/g, 'x').replace(/\s+/g, ' ');
  // feet-inch: 12'-0" x 10'-6"
  let m = s.match(/(\d+)\s*'\s*-?\s*(\d+(?:\.\d+)?)?\s*"?\s*[xX*]\s*(\d+)\s*'\s*-?\s*(\d+(?:\.\d+)?)?\s*"?/);
  if (m) {
    const a = (+m[1] * 12 + (+m[2] || 0)) * 0.0254, b = (+m[3] * 12 + (+m[4] || 0)) * 0.0254;
    return [a, b];
  }
  m = s.match(/(\d+(?:\.\d+)?)\s*(mm|m)?\s*[xX*]\s*(\d+(?:\.\d+)?)\s*(mm|m)?/);
  if (m) {
    let a = +m[1], b = +m[3];
    const unit = (m[2] || m[4] || '').toLowerCase();
    if (unit === 'mm' || (unit === '' && a >= 100 && b >= 100)) { a /= 1000; b /= 1000; }
    return [a, b];
  }
  return null;
}

function pickText(p) {
  const t = hitText(p.x, p.y);
  if (!t) return;
  if (!needActiveItem()) return;
  const g = activeGroup();
  const dims = parseDims(t.t);
  if (dims) {
    const m = newMeas(g, { type: 'text', pts: [[t.x, t.y]], dims }, t.t);
    afterChange();
  } else {
    const ans = prompt(`Text: "${t.t}"\n\nNo L × B found in this text. Type the size (e.g. 4.5 x 3.6), or Cancel:`);
    const d2 = ans ? parseDims(ans) : null;
    if (d2) { newMeas(g, { type: 'text', pts: [[t.x, t.y]], dims: d2 }, t.t); afterChange(); }
  }
}

// ===================================================================== BOQ items & measurements
let gSeq = 1, mSeq = 1;
function activeGroup() { return S.groups.find(g => g.id === S.activeGid) || null; }

function addItemFromPanel() {
  const code = $('dsr-select').value;
  let g;
  if (code === 'CUSTOM') {
    const desc = $('cust-desc').value.trim();
    if (!desc) { alert('Type the item description.'); return; }
    const mode = $('cust-mode').value;
    const unit = { AREA: 'Sq.m', LENGTH: 'Rmt', COUNT: 'Nos', VOLUME: 'Cu.m', WALL_AREA: 'Sq.m', WALL_VOLUME: 'Cu.m' }[mode];
    g = makeGroup({ code: 'CUSTOM-' + gSeq, category: 'Other Works', description: desc, unit, rate: +$('cust-rate').value || 0, calc_mode: mode, is_ref: 'Custom', default_width: 0.23, default_depth: 1 });
    g.title = desc;
  } else {
    const d = S.dsrByCode[code];
    g = makeGroup(d);
  }
  S.groups.push(g);
  S.activeGid = g.id;
  if (S.tool === 'pan') setTool(g.mode === 'COUNT' ? 'count' : (g.mode === 'LENGTH' || g.mode.startsWith('WALL')) ? 'line' : 'poly');
  afterChange();
}

function makeGroup(d) {
  const mode = d.calc_mode;
  let b = 1, h = 1;
  if (mode === 'VOLUME') { b = d.default_width || 1; h = d.default_depth || 0.15; }
  if (mode === 'WALL_AREA') { h = (d.default_depth && d.default_depth > 1) ? d.default_depth : 3.0; }
  if (mode === 'WALL_VOLUME') { b = (d.default_width && d.default_width < 1) ? d.default_width : 0.23; h = (d.default_depth && d.default_depth > 1) ? d.default_depth : 3.0; }
  return {
    id: 'G' + (gSeq++), code: d.code, category: d.category, description: d.description,
    title: shortDesc(d.description), unit: d.unit, rate: d.rate, mode, is_ref: d.is_ref || 'IS 1200',
    b, h, color: PALETTE[(gSeq - 2) % PALETTE.length], meas: [], open: true,
  };
}

function newMeas(g, geo, label) {
  const m = { id: 'M' + (mSeq++), geo, label, nos: 1, L: 1, B: 1, D: 1, manual: false };
  g.meas.push(m);
  recomputeMeas(g, m);
  S.history.push({ gid: g.id, mid: m.id });
  S.selMeas = { gid: g.id, mid: m.id };
  return m;
}

function geoMetrics(geo) {
  const f = S.unitToM, p = geo.pts;
  const len = (pts, closed) => { let s = 0; for (let i = 1; i < pts.length; i++) s += Math.hypot(pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1]); if (closed && pts.length > 2) s += Math.hypot(pts[0][0] - pts[pts.length - 1][0], pts[0][1] - pts[pts.length - 1][1]); return s; };
  const area = pts => { let a = 0; for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) a += (pts[j][0] + pts[i][0]) * (pts[j][1] - pts[i][1]); return Math.abs(a / 2); };
  switch (geo.type) {
    case 'line': return { len: len(p, false) * f };
    case 'poly': return { area: area(p) * f * f, per: len(p, true) * f };
    case 'rect': return { l: Math.abs(p[1][0] - p[0][0]) * f, b: Math.abs(p[2][1] - p[1][1]) * f, per: len(p, true) * f };
    case 'text': return { l: geo.dims[0], b: geo.dims[1], per: 2 * (geo.dims[0] + geo.dims[1]) };
    case 'count': return { n: p.length };
    default: return {};
  }
}

// Fill Nos/L/B/D so that Qty = Nos × L × B × D (same as the Excel measurement sheet)
function recomputeMeas(g, m) {
  if (m.manual) return;
  const k = geoMetrics(m.geo), t = m.geo.type;
  let L = 1, B = 1, D = 1, nos = m.nos === undefined ? 1 : m.nos;
  if (t === 'count') nos = k.n * (m.sign || 1);
  const rectLike = (t === 'rect' || t === 'text');
  switch (g.mode) {
    case 'COUNT': L = B = D = 1; break;
    case 'LENGTH':
      L = t === 'line' ? k.len : (k.per || 1); break;
    case 'AREA':
      if (rectLike) { L = k.l; B = k.b; } else if (t === 'poly') { L = k.area; } else if (t === 'line') { L = k.len; B = g.b; } break;
    case 'VOLUME':
      if (rectLike) { L = k.l; B = k.b; } else if (t === 'poly') { L = k.area; } else if (t === 'line') { L = k.len; B = g.b; }
      D = g.h; break;
    case 'WALL_AREA':
      L = t === 'line' ? k.len : (k.per || 1); B = g.h; break;
    case 'WALL_VOLUME':
      L = t === 'line' ? k.len : (k.per || 1); B = g.b; D = g.h; break;
  }
  Object.assign(m, { nos, L: +L.toFixed(3), B: +B.toFixed(3), D: +D.toFixed(3) });
}

function qtyOf(m) { return m.nos * m.L * m.B * m.D; }
function groupQty(g) { return g.meas.reduce((s, m) => s + qtyOf(m), 0); }

function recomputeAll() { S.groups.forEach(g => g.meas.forEach(m => recomputeMeas(g, m))); afterChange(); }

function afterChange() { renderItems(); draw(); save(); }

function undo() {
  const h = S.history.pop(); if (!h) return;
  const g = S.groups.find(x => x.id === h.gid); if (!g) return;
  const m = g.meas.find(x => x.id === h.mid);
  if (m && m.geo.type === 'count' && m.geo.pts.length > 1) { m.geo.pts.pop(); recomputeMeas(g, m); S.history.push(h); }
  else g.meas = g.meas.filter(x => x.id !== h.mid);
  afterChange();
}

function selectMeasAt(p) {
  // click inside a measured area / near a measured line to highlight its row
  const tol = 6 / S.view.s;
  for (const g of S.groups) for (const m of g.meas) {
    const pts = m.geo.pts;
    if ((m.geo.type === 'poly' || m.geo.type === 'rect') && pip(p.x, p.y, pts)) return selectMeas(g, m);
    if (m.geo.type === 'line') for (let i = 1; i < pts.length; i++) if (distSeg(p, pts[i - 1], pts[i]) < tol) return selectMeas(g, m);
    if (m.geo.type === 'count' || m.geo.type === 'text') for (const q of pts) if (Math.hypot(q[0] - p.x, q[1] - p.y) < tol * 2) return selectMeas(g, m);
  }
  S.selMeas = null; renderItems(); draw();
}
function selectMeas(g, m) { S.activeGid = g.id; g.open = true; S.selMeas = { gid: g.id, mid: m.id }; renderItems(); draw(); const el = document.querySelector(`tr[data-mid="${m.id}"]`); if (el) el.scrollIntoView({ block: 'nearest' }); }
function pip(x, y, pts) { let ins = false; for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) { const xi = pts[i][0], yi = pts[i][1], xj = pts[j][0], yj = pts[j][1]; if (((yi > y) !== (yj > y)) && (x < (xj - xi) * (y - yi) / ((yj - yi) || 1e-12) + xi)) ins = !ins; } return ins; }
function distSeg(p, a, b) { const dx = b[0] - a[0], dy = b[1] - a[1]; const l2 = dx * dx + dy * dy || 1e-12; let t = ((p.x - a[0]) * dx + (p.y - a[1]) * dy) / l2; t = Math.max(0, Math.min(1, t)); return Math.hypot(p.x - (a[0] + t * dx), p.y - (a[1] + t * dy)); }

// ===================================================================== right panel rendering
function paramFields(g) {
  const f = (key, lbl) => `<label>${lbl}<input class="inp" type="number" step="0.01" value="${g[key]}" onchange="setGroupParam('${g.id}','${key}',this.value)"></label>`;
  switch (g.mode) {
    case 'VOLUME': return f('b', 'Width for lengths (m)') + f('h', 'Depth (m)');
    case 'WALL_AREA': return f('h', 'Height (m)');
    case 'WALL_VOLUME': return f('b', 'Thickness (m)') + f('h', 'Height (m)');
    case 'AREA': return f('b', 'Width for lengths (m)');
    default: return '';
  }
}

function renderItems() {
  const box = $('items-list');
  if (!S.groups.length) {
    box.innerHTML = '<div class="muted small pad">No items yet. Pick an item above and press <b>Add</b>, then measure on the plan.</div>';
    return;
  }
  box.innerHTML = S.groups.map(g => {
    const active = g.id === S.activeGid;
    const rows = g.meas.map(m => {
      const sel = S.selMeas && S.selMeas.mid === m.id;
      const q = qtyOf(m);
      const inp = (k, step) => `<td><input value="${m[k]}" type="number" step="${step}" onchange="editMeas('${g.id}','${m.id}','${k}',this.value)"></td>`;
      return `<tr data-mid="${m.id}" class="${sel ? 'sel' : ''} ${q < 0 ? 'ded' : ''}" onclick="selectMeasById('${g.id}','${m.id}')">
        <td><input class="lbl-in" value="${esc(m.label)}" onchange="editMeas('${g.id}','${m.id}','label',this.value)"></td>
        ${inp('nos', 1)}${inp('L', 0.01)}${inp('B', 0.01)}${inp('D', 0.01)}
        <td class="q">${fmt(q, 3)}</td>
        <td><button class="x" title="Delete" onclick="event.stopPropagation();delMeas('${g.id}','${m.id}')">×</button></td></tr>`;
    }).join('');
    return `<div class="item ${active ? 'active' : ''}">
      <div class="item-h" onclick="activateGroup('${g.id}')">
        <span class="swatch" style="background:${g.color}"></span>
        <span class="item-title" title="${esc(g.description)}">${esc(g.code.startsWith('CUSTOM') ? '' : g.code + ' · ')}${esc(g.title)}</span>
        <span class="item-qty">${fmt(groupQty(g), 2)} ${esc(g.unit)}</span>
        <button class="x" title="Delete item" onclick="event.stopPropagation();delGroup('${g.id}')">×</button>
      </div>
      ${g.open ? `<div class="item-body">
        ${paramFields(g) ? `<div class="item-params">${paramFields(g)}</div>` : ''}
        <table class="mt-tbl"><thead><tr><th>Description</th><th>Nos</th><th>L</th><th>B</th><th>D</th><th>Qty</th><th></th></tr></thead>
        <tbody>${rows || '<tr><td colspan="7" class="muted small">Nothing measured yet – use the tools on the drawing.</td></tr>'}</tbody></table>
        <div class="item-foot">
          <button class="btn" onclick="addManualRow('${g.id}')">＋ Manual row</button>
          <button class="btn danger" onclick="addDeduction('${g.id}')">－ Deduction row</button>
        </div></div>` : ''}
    </div>`;
  }).join('');
}

window.activateGroup = (gid) => {
  const g = S.groups.find(x => x.id === gid);
  if (S.activeGid === gid) g.open = !g.open; else { S.activeGid = gid; g.open = true; }
  renderItems(); draw();
};
window.selectMeasById = (gid, mid) => { S.selMeas = { gid, mid }; S.activeGid = gid; renderItems(); draw(); };
window.setGroupParam = (gid, key, v) => { const g = S.groups.find(x => x.id === gid); g[key] = parseFloat(v) || 0; g.meas.forEach(m => recomputeMeas(g, m)); afterChange(); };
window.editMeas = (gid, mid, key, v) => {
  const g = S.groups.find(x => x.id === gid), m = g.meas.find(x => x.id === mid);
  if (key === 'label') m.label = v; else { m[key] = parseFloat(v) || 0; if (key !== 'nos') m.manual = true; }
  afterChange();
};
window.delMeas = (gid, mid) => { const g = S.groups.find(x => x.id === gid); g.meas = g.meas.filter(x => x.id !== mid); afterChange(); };
window.delGroup = (gid) => {
  const g = S.groups.find(x => x.id === gid);
  if (g.meas.length && !confirm(`Delete "${g.title}" and its ${g.meas.length} measurement(s)?`)) return;
  S.groups = S.groups.filter(x => x.id !== gid);
  if (S.activeGid === gid) S.activeGid = S.groups.length ? S.groups[S.groups.length - 1].id : null;
  afterChange();
};
window.addManualRow = (gid) => { const g = S.groups.find(x => x.id === gid); const m = newMeas(g, { type: 'manual', pts: [] }, 'Manual entry'); m.manual = true; afterChange(); };
window.addDeduction = (gid) => { const g = S.groups.find(x => x.id === gid); const m = newMeas(g, { type: 'manual', pts: [] }, 'Deduct: opening'); m.manual = true; m.nos = -1; afterChange(); };

// ===================================================================== drawing
function draw() {
  if (!ctx) return;
  const W = cv.clientWidth, H = cv.clientHeight;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, W, H);
  if (!S.dwg) return;
  const v = S.view;

  // 1. PDF page image
  if (S.img && S.dwg.image) {
    const im = S.dwg.image, a = w2s(im.minx, im.maxy), b = w2s(im.maxx, im.miny);
    ctx.imageSmoothingEnabled = true;
    ctx.drawImage(S.img, a.x, a.y, b.x - a.x, b.y - a.y);
  }

  // 2. CAD vector geometry (true drawing)
  ctx.setTransform(dpr * v.s, 0, 0, -dpr * v.s, dpr * v.tx, dpr * v.ty);
  ctx.globalAlpha = 0.28;
  for (const f of S.fillPaths) { ctx.fillStyle = f.c; ctx.fill(f.p); }
  ctx.globalAlpha = 1;
  ctx.lineWidth = 1 / v.s;
  ctx.lineJoin = 'round';
  for (const p of S.paths) { ctx.strokeStyle = p.c; ctx.stroke(p.p); }

  // 3. CAD texts (PDF texts are already in the image)
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  if (S.dwg.format !== 'PDF') drawTexts(W, H);

  // 4. Measurements overlay
  for (const g of S.groups) for (const m of g.meas) drawMeas(g, m);

  // 5. Text hover box
  if (S.hoverText) {
    const t = S.hoverText, b = textBox(t);
    ctx.save();
    const s = w2s(t.x, t.y);
    ctx.translate(s.x, s.y); ctx.rotate(-t.r * Math.PI / 180);
    ctx.strokeStyle = '#2563eb'; ctx.lineWidth = 2; ctx.fillStyle = 'rgba(37,99,235,.12)';
    ctx.fillRect(b.x0 * v.s, -b.y1 * v.s, (b.x1 - b.x0) * v.s, (b.y1 - b.y0) * v.s);
    ctx.strokeRect(b.x0 * v.s, -b.y1 * v.s, (b.x1 - b.x0) * v.s, (b.y1 - b.y0) * v.s);
    ctx.restore();
  }

  // 6. In-progress shape
  drawPreview();

  // 7. Snap marker
  if (S.cursor && S.cursor.snapped) {
    const s = w2s(S.cursor.x, S.cursor.y);
    ctx.strokeStyle = '#16a34a'; ctx.lineWidth = 2; ctx.strokeRect(s.x - 6, s.y - 6, 12, 12);
  }
}

function drawTexts(W, H) {
  const v = S.view;
  ctx.fillStyle = '#1f2937';
  for (const t of S.dwg.texts) {
    const px = t.h * v.s;
    if (px < 3) continue;
    const s = w2s(t.x, t.y);
    const maxLen = t.t.length * px;
    if (s.x < -maxLen || s.x > W + maxLen || s.y < -maxLen || s.y > H + maxLen) continue;
    ctx.save();
    ctx.translate(s.x, s.y);
    if (t.r) ctx.rotate(-t.r * Math.PI / 180);
    ctx.font = `${px.toFixed(1)}px Arial, sans-serif`;
    ctx.textAlign = t.a === 'c' ? 'center' : t.a === 'r' ? 'right' : 'left';
    ctx.textBaseline = t.b === 'm' ? 'middle' : t.b === 't' ? 'top' : 'alphabetic';
    ctx.fillText(t.t, 0, 0);
    ctx.restore();
  }
}

function hexA(hex, a) { const r = parseInt(hex.slice(1, 3), 16), g = parseInt(hex.slice(3, 5), 16), b = parseInt(hex.slice(5, 7), 16); return `rgba(${r},${g},${b},${a})`; }

function drawMeas(g, m) {
  const sel = S.selMeas && S.selMeas.mid === m.id;
  const act = g.id === S.activeGid;
  const pts = m.geo.pts.map(p => w2s(p[0], p[1]));
  const ded = qtyOf(m) < 0;
  const col = ded ? '#dc2626' : g.color;
  ctx.save();
  ctx.lineWidth = sel ? 4 : 2.5;
  ctx.strokeStyle = col;
  if (ded) ctx.setLineDash([6, 4]);
  const t = m.geo.type;
  if ((t === 'poly' || t === 'rect') && pts.length) {
    ctx.beginPath(); pts.forEach((p, i) => i ? ctx.lineTo(p.x, p.y) : ctx.moveTo(p.x, p.y)); ctx.closePath();
    ctx.fillStyle = hexA(col, sel ? 0.38 : act ? 0.25 : 0.15); ctx.fill(); ctx.stroke();
    const c = centroid(pts); label(c.x, c.y, `${fmt(qtyOf(m), 2)} ${g.unit}`, col);
  } else if (t === 'line' && pts.length) {
    ctx.beginPath(); pts.forEach((p, i) => i ? ctx.lineTo(p.x, p.y) : ctx.moveTo(p.x, p.y)); ctx.stroke();
    pts.forEach(p => { ctx.fillStyle = col; ctx.fillRect(p.x - 2.5, p.y - 2.5, 5, 5); });
    const mid = pts[Math.floor(pts.length / 2)];
    label(mid.x, mid.y - 12, `${fmt(m.L, 2)} m`, col);
  } else if (t === 'count') {
    pts.forEach((p, i) => {
      ctx.beginPath(); ctx.arc(p.x, p.y, 9, 0, Math.PI * 2);
      ctx.fillStyle = hexA(col, 0.85); ctx.fill();
      ctx.fillStyle = '#fff'; ctx.font = 'bold 10px Arial'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText(String(i + 1), p.x, p.y + 0.5);
    });
  } else if (t === 'text' && pts.length) {
    ctx.beginPath(); ctx.arc(pts[0].x, pts[0].y, 6, 0, Math.PI * 2); ctx.fillStyle = col; ctx.fill();
    label(pts[0].x, pts[0].y - 14, `${fmt(m.L, 2)} × ${fmt(m.B, 2)}`, col);
  }
  ctx.restore();
}

function centroid(pts) { let x = 0, y = 0; pts.forEach(p => { x += p.x; y += p.y; }); return { x: x / pts.length, y: y / pts.length }; }

function label(x, y, txt, col) {
  ctx.save();
  ctx.font = 'bold 12px Segoe UI, Arial'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
  const w = ctx.measureText(txt).width + 10;
  ctx.fillStyle = 'rgba(255,255,255,.92)'; ctx.strokeStyle = col; ctx.lineWidth = 1.5; ctx.setLineDash([]);
  ctx.fillRect(x - w / 2, y - 9, w, 18); ctx.strokeRect(x - w / 2, y - 9, w, 18);
  ctx.fillStyle = '#0f172a'; ctx.fillText(txt, x, y);
  ctx.restore();
}

function drawPreview() {
  const lb = $('live-box');
  if (!S.cursor || !['line', 'poly', 'rect', 'calibrate'].includes(S.tool) || !S.pts.length) { lb.classList.add('hidden'); return; }
  const f = S.unitToM;
  const all = [...S.pts, S.cursor];
  const sp = all.map(p => w2s(p.x, p.y));
  ctx.save();
  ctx.strokeStyle = S.tool === 'calibrate' ? '#d97706' : '#2563eb';
  ctx.lineWidth = 2; ctx.setLineDash([7, 5]);
  let info = '';
  if (S.tool === 'rect') {
    const a = sp[0], b = sp[1];
    ctx.fillStyle = 'rgba(37,99,235,.15)';
    ctx.fillRect(Math.min(a.x, b.x), Math.min(a.y, b.y), Math.abs(b.x - a.x), Math.abs(b.y - a.y));
    ctx.strokeRect(Math.min(a.x, b.x), Math.min(a.y, b.y), Math.abs(b.x - a.x), Math.abs(b.y - a.y));
    const L = Math.abs(all[1].x - all[0].x) * f, B = Math.abs(all[1].y - all[0].y) * f;
    info = `L ${L.toFixed(3)} m × B ${B.toFixed(3)} m\nArea ${(L * B).toFixed(3)} m²`;
  } else {
    ctx.beginPath(); sp.forEach((p, i) => i ? ctx.lineTo(p.x, p.y) : ctx.moveTo(p.x, p.y));
    if (S.tool === 'poly' && sp.length > 2) { ctx.closePath(); ctx.fillStyle = 'rgba(37,99,235,.15)'; ctx.fill(); }
    ctx.stroke();
    let len = 0; for (let i = 1; i < all.length; i++) len += Math.hypot(all[i].x - all[i - 1].x, all[i].y - all[i - 1].y);
    const seg = Math.hypot(all[all.length - 1].x - all[all.length - 2].x, all[all.length - 1].y - all[all.length - 2].y) * f;
    info = `Segment ${seg.toFixed(3)} m\nTotal   ${(len * f).toFixed(3)} m`;
    if (S.tool === 'poly' && all.length > 2) {
      let a = 0; for (let i = 0, j = all.length - 1; i < all.length; j = i++) a += (all[j].x + all[i].x) * (all[j].y - all[i].y);
      info += `\nArea    ${(Math.abs(a / 2) * f * f).toFixed(3)} m²`;
    }
  }
  sp.slice(0, -1).forEach(p => { ctx.fillStyle = '#2563eb'; ctx.fillRect(p.x - 3, p.y - 3, 6, 6); });
  ctx.restore();
  lb.textContent = info; lb.classList.remove('hidden');
}

// ===================================================================== BOQ & abstract
function meta() {
  return {
    project_name: $('meta-project').value, client_name: $('meta-client').value, drawing_no: $('meta-dwg').value,
    contingency_pct: +$('meta-cont').value || 0, gst_pct: +$('meta-gst').value || 0,
  };
}

function renderBoq() {
  const gs = S.groups.filter(g => g.meas.length);
  let html = '<tr><th>S.No</th><th>Code</th><th>Description of item</th><th class="n">Qty</th><th>Unit</th><th class="n">Rate ₹</th><th class="n">Amount ₹</th></tr>';
  const cats = {};
  gs.forEach(g => (cats[g.category] = cats[g.category] || []).push(g));
  let sn = 0;
  for (const c of Object.keys(cats)) {
    html += `<tr class="cat"><td colspan="7">${esc(c)}</td></tr>`;
    for (const g of cats[c]) {
      sn++; const q = groupQty(g);
      html += `<tr><td>${sn}</td><td>${esc(g.code)}</td><td class="desc">${esc(g.description)}</td><td class="n">${fmt(q, 2)}</td><td>${esc(g.unit)}</td>
        <td class="n"><input class="inp rate" type="number" value="${g.rate}" onchange="setRate('${g.id}',this.value)"></td><td class="n">${fmt(q * g.rate, 2)}</td></tr>`;
    }
  }
  if (!gs.length) html += '<tr><td colspan="7" class="muted">No measured items yet. Measure on the drawing in tab 1.</td></tr>';
  $('boq-table').innerHTML = html;

  const mt = meta();
  let ab = '<tr><th>S.No</th><th>Category</th><th class="n">Amount ₹</th></tr>';
  let tot = 0, i = 0;
  for (const c of Object.keys(cats)) { const a = cats[c].reduce((s, g) => s + groupQty(g) * g.rate, 0); tot += a; ab += `<tr><td>${++i}</td><td>${esc(c)}</td><td class="n">${fmt(a)}</td></tr>`; }
  const cont = tot * mt.contingency_pct / 100, gst = (tot + cont) * mt.gst_pct / 100;
  ab += `<tr class="tot"><td></td><td>Sub-total</td><td class="n">${fmt(tot)}</td></tr>
    <tr><td></td><td>Contingency @ ${mt.contingency_pct}%</td><td class="n">${fmt(cont)}</td></tr>
    <tr><td></td><td>GST @ ${mt.gst_pct}%</td><td class="n">${fmt(gst)}</td></tr>
    <tr class="grand"><td></td><td>GRAND TOTAL</td><td class="n">₹ ${fmt(tot + cont + gst)}</td></tr>`;
  $('abstract-table').innerHTML = ab;
}
window.setRate = (gid, v) => { const g = S.groups.find(x => x.id === gid); g.rate = parseFloat(v) || 0; renderBoq(); save(); };

async function exportDocument(kind = 'excel') {
  const gs = S.groups.filter(g => g.meas.length);
  if (!gs.length && !S.steel.length) { alert('Nothing to export yet. Measure some items first.'); return; }
  const payload = {
    ...meta(),
    project_type: 'Landscape, Residential & Commercial',
    takeoff_groups: gs.map(g => ({
      group_id: g.id, dsr_code: g.code, category: g.category, display_name: g.title, description: g.description,
      unit: g.unit, rate: g.rate, calc_mode: g.mode, is_ref: g.is_ref, layer: (S.dwg && S.dwg.filename) || '', included: true,
      measurements: g.meas.map(m => ({ label: m.label, nos: m.nos, length: m.L, breadth: m.B, depth: m.D, raw_area: 0 })),
    })),
    bbs_rows: steelToBbsRows(),
  };
  const isPdf = kind === 'pdf';
  const btn = isPdf ? $('btn-export-pdf') : $('btn-export');
  const origTxt = btn ? btn.textContent : '';
  if (btn) { btn.disabled = true; btn.textContent = isPdf ? '⏳ Exporting PDF…' : '⏳ Exporting…'; }
  try {
    const r = await fetch(isPdf ? '/api/export-pdf' : '/api/export-excel', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload),
    });
    if (!r.ok) { alert((isPdf ? 'PDF' : 'Excel') + ' export failed: ' + (await r.text()).slice(0, 300)); return; }
    const blob = await r.blob();
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = (payload.project_name || 'BOQ').replace(/[^\w]+/g, '_').slice(0, 40) + (isPdf ? '_BOQ.pdf' : '_BOQ.xlsx');
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = origTxt; }
  }
}

// ===================================================================== steel (tender kg)
function initSteel() {
  const opts = DIAS.map(d => `<option value="${d}">${d} mm</option>`).join('');
  ['ft-dx', 'ft-dy', 'ft-tdx', 'ft-tdy', 'ob-dia'].forEach(id => { if ($(id)) $(id).innerHTML = opts; });
  $('ft-dx').value = '12'; $('ft-dy').value = '12';
  if ($('ft-tdx')) $('ft-tdx').value = '10';
  if ($('ft-tdy')) $('ft-tdy').value = '10';
  $('ob-dia').value = '16';
  if ($('ft-top-mat')) {
    $('ft-top-mat').addEventListener('change', () => {
      $('ft-top-mat-fields').classList.toggle('hidden', !$('ft-top-mat').checked);
      footingPreview();
    });
  }
  document.querySelectorAll('#tab-steel .steel-grid input, #tab-steel .steel-grid select').forEach(el => el.addEventListener('input', footingPreview));
  $('btn-add-footing').addEventListener('click', addFooting);
  $('btn-add-bar').addEventListener('click', addOtherBar);
  $('btn-steel-to-boq').addEventListener('click', steelToBoq);
  $('steel-rate').addEventListener('input', renderSteel);
  $('ob-shape').addEventListener('change', onShapeOrAngleChange);
  $('ob-angle').addEventListener('change', updateShapePreview);
  updateShapePreview();
  footingPreview(); renderSteel();
}

const unitWt = d => d * d / 162.2;

function footingCalc() {
  const v = id => parseFloat($(id).value) || 0;
  const L = v('ft-L'), B = v('ft-B'), D = v('ft-D'), c = v('ft-cover'), nos = Math.max(1, Math.round(v('ft-nos')));
  const dx = +$('ft-dx').value, dy = +$('ft-dy').value, sx = v('ft-sx') || 150, sy = v('ft-sy') || 150;
  const legIn = $('ft-leg').value.trim();
  const leg = legIn === '' ? Math.max(0, D - 2 * c) : parseFloat(legIn) || 0;
  // Longitudinal bars run along L, spaced across B
  const nX = Math.floor((B - 2 * c) / sx) + 1;
  const cutX = (L - 2 * c + 2 * leg - (leg > 0 ? 2 * 2 * dx : 0)) / 1000; // 2 × 90° bends @ 2d each
  const nY = Math.floor((L - 2 * c) / sy) + 1;
  const cutY = (B - 2 * c + 2 * leg - (leg > 0 ? 2 * 2 * dy : 0)) / 1000;
  const kgX = nX * cutX * unitWt(dx) * nos, kgY = nY * cutY * unitWt(dy) * nos;
  const hasTop = $('ft-top-mat') && $('ft-top-mat').checked;
  const tdx = hasTop ? +$('ft-tdx').value : 0, tdy = hasTop ? +$('ft-tdy').value : 0;
  const tsx = hasTop ? (v('ft-tsx') || 150) : 150, tsy = hasTop ? (v('ft-tsy') || 150) : 150;
  const tnX = hasTop ? Math.floor((B - 2 * c) / tsx) + 1 : 0;
  const tcutX = hasTop ? (L - 2 * c + 2 * leg - (leg > 0 ? 2 * 2 * tdx : 0)) / 1000 : 0;
  const tnY = hasTop ? Math.floor((L - 2 * c) / tsy) + 1 : 0;
  const tcutY = hasTop ? (B - 2 * c + 2 * leg - (leg > 0 ? 2 * 2 * tdy : 0)) / 1000 : 0;
  const tkgX = tnX * tcutX * unitWt(tdx) * nos, tkgY = tnY * tcutY * unitWt(tdy) * nos;
  return { mark: $('ft-mark').value || 'F', L, B, D, c, nos, dx, dy, sx, sy, leg, nX, nY, cutX, cutY, kgX, kgY, hasTop, tdx, tdy, tsx, tsy, tnX, tnY, tcutX, tcutY, tkgX, tkgY };
}

function footingPreview() {
  const f = footingCalc();
  const topTxt = f.hasTop ? `\n\nTOP MAT Lx – Ø${f.tdx} @ ${f.tsx} c/c: ${f.tnX} bars × ${f.tcutX.toFixed(3)} m = ${f.tkgX.toFixed(2)} kg\nTOP MAT Ly – Ø${f.tdy} @ ${f.tsy} c/c: ${f.tnY} bars × ${f.tcutY.toFixed(3)} m = ${f.tkgY.toFixed(2)} kg` : '';
  $('ft-calc').textContent =
`LONGITUDINAL (along L) – Ø${f.dx} @ ${f.sx} c/c
  No. of bars  = (B − 2c) / s + 1 = (${f.B} − ${2 * f.c}) / ${f.sx} + 1 = ${f.nX}
  Cut length   = L − 2c + 2×leg − bends = ${f.cutX.toFixed(3)} m
  Weight       = ${f.nX} × ${f.cutX.toFixed(3)} × ${unitWt(f.dx).toFixed(3)} kg/m × ${f.nos} nos = ${f.kgX.toFixed(2)} kg

TRANSVERSE (along B) – Ø${f.dy} @ ${f.sy} c/c
  No. of bars  = (L − 2c) / s + 1 = (${f.L} − ${2 * f.c}) / ${f.sy} + 1 = ${f.nY}
  Cut length   = B − 2c + 2×leg − bends = ${f.cutY.toFixed(3)} m
  Weight       = ${f.nY} × ${f.cutY.toFixed(3)} × ${unitWt(f.dy).toFixed(3)} kg/m × ${f.nos} nos = ${f.kgY.toFixed(2)} kg${topTxt}

TOTAL for ${f.nos} × ${f.mark} = ${(f.kgX + f.kgY + f.tkgX + f.tkgY).toFixed(2)} kg   (unit wt = d²/162.2)`;
  drawFootingSvg(f);
}

function drawFootingSvg(f) {
  const W = 420, H = 330, pad = 40;
  const sc = Math.min((W - 2 * pad) / Math.max(f.L, 1), (H - 2 * pad - 30) / Math.max(f.B, 1));
  const w = f.L * sc, h = f.B * sc, x0 = (W - w) / 2, y0 = 30 + (H - 30 - h) / 2, c = f.c * sc;
  let s = `<text x="${W / 2}" y="18" text-anchor="middle" font-size="13" font-weight="700" fill="#0f172a">PLAN – Footing ${esc(f.mark)} (${f.L} × ${f.B} mm)</text>`;
  s += `<rect x="${x0}" y="${y0}" width="${w}" height="${h}" fill="#e5e7eb" stroke="#334155" stroke-width="2"/>`;
  const nY = Math.min(f.nY, 40), nX = Math.min(f.nX, 40);
  for (let i = 0; i < nX; i++) { const y = y0 + c + (nX > 1 ? i * (h - 2 * c) / (nX - 1) : (h - 2 * c) / 2); s += `<line x1="${x0 + c}" y1="${y}" x2="${x0 + w - c}" y2="${y}" stroke="#dc2626" stroke-width="2"/>`; }
  for (let i = 0; i < nY; i++) { const x = x0 + c + (nY > 1 ? i * (w - 2 * c) / (nY - 1) : (w - 2 * c) / 2); s += `<line x1="${x}" y1="${y0 + c}" x2="${x}" y2="${y0 + h - c}" stroke="#2563eb" stroke-width="1.6"/>`; }
  s += `<text x="${x0 + w / 2}" y="${y0 + h + 18}" text-anchor="middle" font-size="12" fill="#334155">L = ${f.L}</text>`;
  s += `<text x="${x0 - 8}" y="${y0 + h / 2}" text-anchor="end" font-size="12" fill="#334155">B = ${f.B}</text>`;
  s += `<text x="${x0 + w + 6}" y="${y0 + 12}" font-size="11" fill="#dc2626" font-weight="700">Ø${f.dx}@${f.sx} (${f.nX} nos)</text>`;
  s += `<text x="${x0 + w + 6}" y="${y0 + 26}" font-size="11" fill="#2563eb" font-weight="700">Ø${f.dy}@${f.sy} (${f.nY} nos)</text>`;
  s += `<text x="${W / 2}" y="${H - 6}" text-anchor="middle" font-size="11" fill="#64748b">Red = longitudinal (along L) · Blue = transverse (along B)</text>`;
  $('ft-svg').innerHTML = s;
}

function shapeSvg(code, angle = 90, w = 64, h = 26) {
  const ang = Number(angle);
  let inner = '';
  if (code === 'SC-00') {
    inner = '<line x1="4" y1="13" x2="60" y2="13" stroke="#dc2626" stroke-width="2.5"/>';
  } else if (code === 'SC-01') {
    if (ang === 45) inner = '<polyline points="10,5 20,20 58,20" fill="none" stroke="#dc2626" stroke-width="2.5"/>';
    else if (ang === 135) inner = '<polyline points="20,5 10,20 58,20" fill="none" stroke="#dc2626" stroke-width="2.5"/>';
    else if (ang === 180) inner = '<path d="M 16,10 A 5,5 0 0,0 16,20 L 58,20" fill="none" stroke="#dc2626" stroke-width="2.5"/>';
    else inner = '<polyline points="10,4 10,20 58,20" fill="none" stroke="#dc2626" stroke-width="2.5"/>';
  } else if (code === 'SC-02') {
    if (ang === 45) inner = '<polyline points="8,5 16,20 48,20 56,5" fill="none" stroke="#dc2626" stroke-width="2.5"/>';
    else if (ang === 135) inner = '<polyline points="16,5 8,20 56,20 48,5" fill="none" stroke="#dc2626" stroke-width="2.5"/>';
    else inner = '<polyline points="10,5 10,20 54,20 54,5" fill="none" stroke="#dc2626" stroke-width="2.5"/>';
  } else if (code === 'SC-41') {
    if (ang === 90) inner = '<polyline points="3,6 18,6 18,20 46,20 46,6 61,6" fill="none" stroke="#dc2626" stroke-width="2.5"/>';
    else if (ang === 135) inner = '<polyline points="3,6 22,6 15,20 49,20 42,6 61,6" fill="none" stroke="#dc2626" stroke-width="2.5"/>';
    else inner = '<polyline points="3,6 14,6 22,20 42,20 50,6 61,6" fill="none" stroke="#dc2626" stroke-width="2.5"/>';
  } else if (code === 'SC-51') {
    if (ang === 90) inner = '<rect x="16" y="4" width="32" height="18" fill="none" stroke="#2563eb" stroke-width="2.2"/><line x1="16" y1="4" x2="25" y2="4" stroke="#dc2626" stroke-width="2.2"/><line x1="16" y1="4" x2="16" y2="13" stroke="#dc2626" stroke-width="2.2"/>';
    else inner = '<rect x="16" y="4" width="32" height="18" fill="none" stroke="#2563eb" stroke-width="2.2"/><line x1="16" y1="4" x2="24" y2="12" stroke="#dc2626" stroke-width="2.2"/><line x1="16" y1="4" x2="21" y2="14" stroke="#dc2626" stroke-width="2.2"/>';
  } else {
    inner = '<line x1="4" y1="13" x2="60" y2="13" stroke="#dc2626" stroke-width="2.5"/>';
  }
  return `<svg viewBox="0 0 64 26" width="${w}" height="${h}" style="vertical-align:middle;background:#fff;border:1px solid #cbd5e1;border-radius:4px">${inner}</svg>`;
}

function onShapeOrAngleChange() {
  const sc = $('ob-shape').value;
  if (sc === 'SC-00') $('ob-angle').value = '0';
  else if (sc === 'SC-41') $('ob-angle').value = '45';
  else if (sc === 'SC-51') $('ob-angle').value = '135';
  else $('ob-angle').value = '90';
  updateShapePreview();
}

function updateShapePreview() {
  const box = $('ob-shape-preview');
  if (!box) return;
  const sc = $('ob-shape').value;
  const ang = +$('ob-angle').value;
  const dedMap = { 0: '0d (Straight)', 45: '1d per bend', 90: '2d per bend', 135: '3d per bend (Hook = 10d)', 180: '4d per bend' };
  const crankMap = { 45: 'Crank Extra = +0.42×Hc', 90: 'Step Extra = +1.00×Hc', 135: 'Crank Extra = +1.30×Hc' };
  const extraTxt = sc === 'SC-41' ? ` · ${crankMap[ang] || ''}` : '';
  box.innerHTML = `${shapeSvg(sc, ang, 110, 42)}
    <div style="font-size:12px"><b>Live Shape Preview: ${esc(sc)} @ ${ang}°</b><br>
    <span class="muted">IS 2502 Bend Deduction: ${dedMap[ang] || '2d'}${extraTxt}</span></div>`;
}

function addFooting() {
  const f = footingCalc();
  S.steel.push({ member: `Footing ${f.mark}`, shape: 'SC-02', angle: 90, barType: 'BOT-MESH-LX', desc: `Bottom Longitudinal Ø${f.dx}@${f.sx} (along L)`, nos: f.nos, bars: f.nX, dia: f.dx, spacing: f.sx, cut: +f.cutX.toFixed(3), leg: f.leg / 1000, kind: 'footing' });
  S.steel.push({ member: `Footing ${f.mark}`, shape: 'SC-02', angle: 90, barType: 'BOT-MESH-LY', desc: `Bottom Transverse Ø${f.dy}@${f.sy} (along B)`, nos: f.nos, bars: f.nY, dia: f.dy, spacing: f.sy, cut: +f.cutY.toFixed(3), leg: f.leg / 1000, kind: 'footing' });
  if (f.hasTop) {
    S.steel.push({ member: `Footing ${f.mark}`, shape: 'SC-02', angle: 90, barType: 'TOP-MAT-LX', desc: `Top Mat Longitudinal Ø${f.tdx}@${f.tsx} (along L)`, nos: f.nos, bars: f.tnX, dia: f.tdx, spacing: f.tsx, cut: +f.tcutX.toFixed(3), leg: f.leg / 1000, kind: 'footing' });
    S.steel.push({ member: `Footing ${f.mark}`, shape: 'SC-02', angle: 90, barType: 'TOP-MAT-LY', desc: `Top Mat Transverse Ø${f.tdy}@${f.tsy} (along B)`, nos: f.nos, bars: f.tnY, dia: f.tdy, spacing: f.tsy, cut: +f.tcutY.toFixed(3), leg: f.leg / 1000, kind: 'footing' });
  }
  renderSteel(); save();
}

function addOtherBar() {
  const shape = $('ob-shape').value || 'SC-00';
  const angle = +$('ob-angle').value || 0;
  const barType = $('ob-type').value || 'MAIN';
  const desc = $('ob-desc').value.trim() || `Ø${$('ob-dia').value} bars`;
  S.steel.push({ member: $('ob-member').value || 'Member', shape, angle, barType, desc, nos: +$('ob-nos').value || 1, bars: +$('ob-bars').value || 1, dia: +$('ob-dia').value, spacing: 0, cut: +$('ob-len').value || 0, leg: 0, kind: 'other' });
  renderSteel(); save();
}

window.updateSteelRowShape = (i, field, val) => {
  if (!S.steel[i]) return;
  if (field === 'angle') S.steel[i].angle = +val;
  else {
    S.steel[i].shape = val;
    if (val === 'SC-00') S.steel[i].angle = 0;
    else if (val === 'SC-41' && !S.steel[i].angle) S.steel[i].angle = 45;
    else if (val === 'SC-51') S.steel[i].angle = 135;
    else if ((val === 'SC-01' || val === 'SC-02') && !S.steel[i].angle) S.steel[i].angle = 90;
  }
  renderSteel(); save();
};

function renderSteel() {
  let html = '<tr><th>Member</th><th>Shape</th><th>Figure</th><th>Angle</th><th>Bar Type</th><th>Bar Description</th><th class="n">Members</th><th class="n">Bars/mem</th><th class="n">Dia</th><th class="n">Cut len (m)</th><th class="n">Total len (m)</th><th class="n">kg/m</th><th class="n">Weight (kg)</th><th></th></tr>';
  let tot = 0; const byDia = {};
  const shapes = ['SC-00', 'SC-01', 'SC-02', 'SC-41', 'SC-51'];
  const angles = [0, 45, 90, 135, 180];
  S.steel.forEach((r, i) => {
    const sc = r.shape || (r.leg > 0 ? 'SC-02' : 'SC-00');
    const ang = r.angle !== undefined ? r.angle : (sc === 'SC-41' ? 45 : sc === 'SC-51' ? 135 : sc === 'SC-00' ? 0 : 90);
    const bt = r.barType || (r.kind === 'footing' ? 'MESH' : 'MAIN');
    const tl = r.nos * r.bars * r.cut, kg = tl * unitWt(r.dia); tot += kg; byDia[r.dia] = (byDia[r.dia] || 0) + kg;
    const scSel = `<select class="sel" style="padding:2px 4px;font-size:11.5px" onchange="updateSteelRowShape(${i},'shape',this.value)">${shapes.map(s => `<option value="${s}" ${s === sc ? 'selected' : ''}>${s}</option>`).join('')}</select>`;
    const angSel = `<select class="sel" style="padding:2px 4px;font-size:11.5px" onchange="updateSteelRowShape(${i},'angle',this.value)">${angles.map(a => `<option value="${a}" ${a === ang ? 'selected' : ''}>${a}°</option>`).join('')}</select>`;
    html += `<tr><td>${esc(r.member)}</td><td>${scSel}</td><td>${shapeSvg(sc, ang)}</td><td>${angSel}</td><td>${esc(bt)}</td><td>${esc(r.desc)}</td><td class="n">${r.nos}</td><td class="n">${r.bars}</td><td class="n">${r.dia}</td><td class="n">${r.cut.toFixed(3)}</td><td class="n">${fmt(tl)}</td><td class="n">${unitWt(r.dia).toFixed(3)}</td><td class="n">${fmt(kg)}</td><td><button class="x" onclick="delSteel(${i})">×</button></td></tr>`;
  });
  const dia = Object.keys(byDia).sort((a, b) => a - b).map(d => `Ø${d}: ${fmt(byDia[d])} kg`).join(' · ');
  html += `<tr class="tot"><td colspan="12">Total steel ${dia ? '(' + dia + ')' : ''}</td><td class="n">${fmt(tot)} kg</td><td></td></tr>`;
  html += `<tr class="tot"><td colspan="12">= ${(tot / 1000).toFixed(3)} MT × ₹${$('steel-rate').value}/kg</td><td class="n">₹ ${fmt(tot * (+$('steel-rate').value || 0))}</td><td></td></tr>`;
  $('steel-table').innerHTML = html;
  S.steelTotal = tot;
}
window.delSteel = (i) => { S.steel.splice(i, 1); renderSteel(); save(); };

function steelToBoq() {
  if (!S.steel.length) { alert('Add footing / bars first.'); return; }
  let g = S.groups.find(x => x.code === 'STEEL-FE500D');
  if (!g) {
    g = makeGroup({ code: 'STEEL-FE500D', category: 'Civil - RCC Reinforcement', description: 'Providing and placing TMT Fe-500D reinforcement steel including cutting, bending, binding with annealed wire and placing in position complete as per drawing (IS 1786 / IS 2502).', unit: 'kg', rate: +$('steel-rate').value || 0, calc_mode: 'LENGTH', is_ref: 'IS 1200 (Part 2) / IS 2502' });
    g.title = 'Reinforcement steel Fe-500D (kg)';
    S.groups.push(g);
  }
  g.rate = +$('steel-rate').value || g.rate;
  g.meas = [{ id: 'M' + (mSeq++), geo: { type: 'manual', pts: [] }, label: 'As per steel list (tab 3)', nos: 1, L: +S.steelTotal.toFixed(2), B: 1, D: 1, manual: true }];
  afterChange();
  alert(`Steel ${fmt(S.steelTotal)} kg added to the BOQ.`);
}

function steelToBbsRows() {
  return S.steel.map((r, i) => ({
    member_group: r.member, bar_mark: `B${i + 1}`, description: r.desc,
    shape_code: r.shape || (r.leg > 0 ? 'SC-02' : 'SC-00'), dia_mm: r.dia, spacing_mm: r.spacing || 0,
    num_members: r.nos, bars_per_member: r.bars,
    a_m: r.leg > 0 ? r.leg : r.cut, b_m: r.leg > 0 ? +(r.cut - 2 * r.leg + 4 * r.dia / 1000).toFixed(3) : 0, c_m: r.leg > 0 ? r.leg : 0, d_m: 0,
    bend_deduction_m: r.leg > 0 ? +(4 * r.dia / 1000).toFixed(3) : 0,
  }));
}

// ===================================================================== autosave (browser storage)
function save() {
  try {
    const key = 'autoboq3:' + (S.dwg ? S.dwg.filename : '_');
    localStorage.setItem(key, JSON.stringify({ groups: S.groups, unitToM: S.unitToM, unitLabel: S.unitLabel }));
    localStorage.setItem('autoboq3:common', JSON.stringify({ meta: meta(), steel: S.steel, rate: $('steel-rate').value }));
  } catch (e) { /* storage full – ignore */ }
}

function loadSavedFor(fname) {
  S.groups = []; S.activeGid = null; S.history = []; S.selMeas = null;
  try {
    const raw = localStorage.getItem('autoboq3:' + fname);
    if (!raw) return;
    const d = JSON.parse(raw);
    if (!d.groups || !d.groups.length) return;
    if (confirm(`You measured this drawing before (${d.groups.length} BOQ items). Restore that takeoff?`)) {
      S.groups = d.groups;
      S.activeGid = S.groups[0].id;
      const num = s => parseInt(String(s).replace(/\D/g, ''), 10) || 0;
      gSeq = Math.max(0, ...S.groups.map(g => num(g.id))) + 1;
      mSeq = Math.max(0, ...S.groups.flatMap(g => g.meas.map(m => num(m.id)))) + 1;
      if (d.unitToM && Math.abs(d.unitToM - S.unitToM) > 1e-12) {
        S.unitToM = d.unitToM; S.unitLabel = d.unitLabel || 'restored';
        const sel = $('unit-select');
        const o = [...sel.options].find(o => Math.abs(parseFloat(o.value) - d.unitToM) < 1e-12);
        if (o) sel.value = o.value; else { const c = sel.querySelector('option[value="cal"]'); c.disabled = false; sel.value = 'cal'; c.disabled = true; }
        updateUnitStatus('restored');
      }
    }
  } catch (e) { /* ignore */ }
}

function restoreSession() {
  try {
    const c = JSON.parse(localStorage.getItem('autoboq3:common') || 'null');
    if (!c) return;
    if (c.meta) { $('meta-project').value = c.meta.project_name || $('meta-project').value; $('meta-client').value = c.meta.client_name || ''; $('meta-dwg').value = c.meta.drawing_no || ''; $('meta-cont').value = c.meta.contingency_pct ?? 3; $('meta-gst').value = c.meta.gst_pct ?? 18; }
    if (c.steel) S.steel = c.steel;
    if (c.rate) $('steel-rate').value = c.rate;
    renderSteel();
  } catch (e) { /* ignore */ }
}
