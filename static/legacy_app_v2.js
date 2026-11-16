/**
 * AutoBOQ 2D Wizard — 3-Step Guided 2D Plan Takeoff, Scale Calibration, Room/Wall/Count Tracer & BOQ Studio
 */

let dsrLibrary = [];
let dsrByCode = {};
let currentDrawing = null;
let currentSampleKey = 'mixed';
let currentWizardStep = 1;

// Canvas & Interactive Tool State
let canvas, ctx;
let bgImageObj = null;
let viewScale = 12;
let viewOffsetX = 40;
let viewOffsetY = 40;
let isPanning = false;
let panStartX = 0;
let panStartY = 0;
let mouseMovedDist = 0;
let showLabels = true;
let highlightedGroupId = null;
let hoveredEntity = null;

// Interactive Takeoff Tools: 'INSPECT' | 'TRACE_AREA' | 'TRACE_WALL' | 'COUNT' | 'CALIBRATE'
let activeCanvasTool = 'INSPECT';
let traceStartPt = null;
let currentWorldCursor = { x: 0, y: 0 };

const inrFormatter = new Intl.NumberFormat('en-IN', {
  style: 'currency',
  currency: 'INR',
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

const numFormatter = new Intl.NumberFormat('en-IN', {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

window.addEventListener('DOMContentLoaded', async () => {
  initCanvas();
  await fetchDsrLibrary();
  await loadSample('mixed');
});

// ============================================================================
// 3-STEP WIZARD NAVIGATION
// ============================================================================
function goToWizardStep(stepNum) {
  currentWizardStep = stepNum;
  [1, 2, 3].forEach(n => {
    const container = document.getElementById(`step-${n}-container`);
    const btn = document.getElementById(`wiz-step-${n}`);
    if (!container || !btn) return;
    if (n === stepNum) {
      container.classList.remove('hidden');
      btn.className = 'px-4 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-2 transition bg-white text-indigo-700 shadow-sm';
    } else {
      container.classList.add('hidden');
      btn.className = 'px-4 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-2 transition text-slate-600 hover:text-slate-900';
    }
  });

  if (stepNum === 2) {
    setTimeout(() => {
      resizeCanvas();
      resetView();
      updateQuickFootingCalc();
    }, 40);
  } else if (stepNum === 3) {
    renderAllTables();
  }
}

function switchStep2Tab(key) {
  ['items', 'measurements', 'steel'].forEach(k => {
    const btn = document.getElementById(`s2-tab-${k}`);
    const panel = document.getElementById(`s2-panel-${k}`);
    if (!btn || !panel) return;
    if (k === key) {
      btn.className = 'px-3.5 py-2 text-xs font-bold rounded-t-lg border-b-2 border-indigo-600 text-indigo-700 bg-white';
      panel.classList.remove('hidden');
    } else {
      btn.className = 'px-3.5 py-2 text-xs font-semibold rounded-t-lg border-b-2 border-transparent text-slate-500 hover:text-slate-800';
      panel.classList.add('hidden');
    }
  });
  if (key === 'steel') updateQuickFootingCalc();
}

function switchStep3Tab(key) {
  ['boq', 'abstract', 'bbs'].forEach(k => {
    const btn = document.getElementById(`s3-tab-${k}`);
    const panel = document.getElementById(`s3-panel-${k}`);
    if (!btn || !panel) return;
    if (k === key) {
      btn.className = 'px-4 py-2 rounded-xl text-xs font-bold bg-indigo-600 text-white shadow-sm';
      panel.classList.remove('hidden');
    } else {
      btn.className = 'px-4 py-2 rounded-xl text-xs font-semibold bg-slate-100 text-slate-700 hover:bg-slate-200';
      panel.classList.add('hidden');
    }
  });
}

// ============================================================================
// DSR LIBRARY & FILE LOADING
// ============================================================================
async function fetchDsrLibrary() {
  const res = await fetch('/api/dsr-library');
  const data = await res.json();
  dsrLibrary = data.items || [];
  dsrByCode = {};
  dsrLibrary.forEach(item => {
    dsrByCode[item.code] = item;
  });
  populateDsrDropdowns();
}

function populateDsrDropdowns() {
  const html = dsrLibrary.map(it =>
    `<option value="${it.code}">[${it.code}] ${it.category.split('-').pop().trim()} — ${it.description.slice(0, 56)}... (${it.unit} @ ₹${it.rate})</option>`
  ).join('');

  const modalSel = document.getElementById('modal-dsr-select');
  const targetSel = document.getElementById('active-target-dsr');
  if (modalSel) modalSel.innerHTML = html;
  if (targetSel) {
    targetSel.innerHTML = html;
    targetSel.addEventListener('change', () => {
      const dsr = dsrByCode[targetSel.value];
      if (dsr) {
        document.getElementById('active-target-b').value = dsr.default_width || 1.0;
        document.getElementById('active-target-d').value = dsr.default_depth || 1.0;
      }
    });
  }
  onModalDsrChange();
}

function onModalDsrChange() {
  const sel = document.getElementById('modal-dsr-select');
  if (!sel) return;
  const it = dsrByCode[sel.value];
  if (!it) return;
  document.getElementById('modal-category').value = it.category;
  document.getElementById('modal-desc').value = it.description;
  document.getElementById('modal-unit').value = it.unit;
  document.getElementById('modal-rate').value = it.rate;
  document.getElementById('modal-wid').value = it.default_width || 1.0;
  document.getElementById('modal-dep').value = it.default_depth || 1.0;
}

async function loadSample(sampleKey) {
  currentSampleKey = sampleKey;
  ['mixed', 'pdf_sample', 'landscape', 'residential', 'commercial'].forEach(k => {
    const btn = document.getElementById(`btn-sample-${k}`);
    if (!btn) return;
    if (k === sampleKey) {
      btn.className = 'px-3 py-2 rounded-lg text-xs font-bold bg-indigo-600 text-white shadow-sm';
    } else {
      btn.className = 'px-3 py-2 rounded-lg text-xs font-medium bg-white text-slate-700 border border-slate-200 hover:bg-slate-50';
    }
  });

  const scale = parseFloat(document.getElementById('unit-scale-select').value || '1.0');
  const res = await fetch(`/api/sample/${sampleKey}?unit_scale=${scale}&grouping_mode=SMART`);
  const data = await res.json();
  applyDrawingData(data);
}

async function onUnitScaleChange() {
  const scale = parseFloat(document.getElementById('unit-scale-select').value || '1.0');
  if (currentSampleKey) {
    const res = await fetch(`/api/sample/${currentSampleKey}?unit_scale=${scale}&grouping_mode=SMART`);
    const data = await res.json();
    applyDrawingData(data);
  }
}

async function handleFileUpload(ev) {
  const file = ev.target.files[0];
  if (!file) return;
  currentSampleKey = null;

  const scale = parseFloat(document.getElementById('unit-scale-select').value || '1.0');
  const formData = new FormData();
  formData.append('file', file);
  formData.append('unit_scale', scale.toString());
  formData.append('project_type', 'All');
  formData.append('grouping_mode', 'SMART');

  const res = await fetch('/api/upload-dxf', {
    method: 'POST',
    body: formData,
  });

  if (!res.ok) {
    const err = await res.json();
    alert(err.detail || 'Error uploading drawing file.');
    return;
  }
  const data = await res.json();
  applyDrawingData(data);
  goToWizardStep(2);
}

function applyDrawingData(data) {
  currentDrawing = data;
  highlightedGroupId = null;
  traceStartPt = null;

  // Load crisp PDF background image if available
  bgImageObj = null;
  if (data.clean_plan && data.clean_plan.background_image_b64) {
    const img = new Image();
    img.onload = () => {
      bgImageObj = img;
      drawCadCanvas();
    };
    img.src = data.clean_plan.background_image_b64;
  }

  if (data.meta) {
    document.getElementById('meta-project-name').value = data.meta.project_name || '';
    document.getElementById('meta-client-name').value = data.meta.client_name || '';
    document.getElementById('meta-drawing-no').value = data.meta.drawing_no || '';
  }
  document.getElementById('canvas-filename-badge').textContent = data.filename || 'drawing.dxf';
  document.getElementById('bounds-info').textContent =
    `Plan Extents: ${data.bounds.width.toFixed(1)}m × ${data.bounds.height.toFixed(1)}m`;

  resetView();
  renderAllTables();
}

// ============================================================================
// INTERACTIVE 2D PLAN CANVAS & PRECISION MEASUREMENT TOOLS
// ============================================================================
function setCanvasTool(toolName) {
  activeCanvasTool = toolName;
  traceStartPt = null;

  ['INSPECT', 'TRACE_AREA', 'TRACE_WALL', 'COUNT', 'CALIBRATE'].forEach(t => {
    const btn = document.getElementById(`tool-btn-${t}`);
    if (!btn) return;
    if (t === toolName) {
      btn.className = 'px-2.5 py-1.5 rounded-lg font-bold bg-indigo-600 text-white flex items-center gap-1 shadow-sm';
    } else {
      btn.className = 'px-2.5 py-1.5 rounded-lg font-medium bg-slate-100 text-slate-700 hover:bg-slate-200 flex items-center gap-1';
    }
  });

  const hintEl = document.getElementById('tool-status-hint');
  if (toolName === 'INSPECT') {
    hintEl.textContent = 'Click any Room Polygon, Text Dimension, or Wall on the plan to inspect or add to BOQ';
    // Auto-select appropriate DSR item when user clicks
  } else if (toolName === 'TRACE_AREA') {
    hintEl.textContent = '📐 Click Corner 1 of a Room/Zone on the plan, then click Corner 2 to measure exact L × B Area';
  } else if (toolName === 'TRACE_WALL') {
    hintEl.textContent = '🧱 Click Start Point of a Wall/Line on the plan, then click End Point to measure Running Length';
  } else if (toolName === 'COUNT') {
    hintEl.textContent = '📍 Click on each Tree, Door, Window, Column, or Light on the plan to count (1 click = +1 No)';
  } else if (toolName === 'CALIBRATE') {
    hintEl.textContent = '📏 Click Point 1 and Point 2 of a known dimension on the drawing to calibrate 100% exact scale';
  }
  drawCadCanvas();
}

function resizeCanvas() {
  if (!canvas) return;
  const cw = canvas.clientWidth || 700;
  const ch = canvas.clientHeight || 500;
  canvas.width = cw * window.devicePixelRatio;
  canvas.height = ch * window.devicePixelRatio;
  ctx.setTransform(window.devicePixelRatio, 0, 0, window.devicePixelRatio, 0, 0);
}

function initCanvas() {
  canvas = document.getElementById('cad-canvas');
  ctx = canvas.getContext('2d');

  const resizeObserver = new ResizeObserver(() => {
    resizeCanvas();
    if (currentDrawing) drawCadCanvas();
  });
  resizeObserver.observe(document.getElementById('canvas-container'));

  canvas.addEventListener('mousedown', e => {
    isPanning = true;
    mouseMovedDist = 0;
    panStartX = e.clientX - viewOffsetX;
    panStartY = e.clientY - viewOffsetY;
  });

  window.addEventListener('mousemove', e => {
    if (!canvas || currentWizardStep !== 2) return;
    const rect = canvas.getBoundingClientRect();
    const mx = e.clientX - rect.left;
    const my = e.clientY - rect.top;

    if (isPanning) {
      const dx = (e.clientX - panStartX) - viewOffsetX;
      const dy = (e.clientY - panStartY) - viewOffsetY;
      mouseMovedDist += Math.hypot(dx, dy);
      viewOffsetX = e.clientX - panStartX;
      viewOffsetY = e.clientY - panStartY;
      drawCadCanvas();
      return;
    }

    if (mx >= 0 && my >= 0 && mx <= rect.width && my <= rect.height && currentDrawing) {
      const rawWorld = screenToWorld(mx, my);
      currentWorldCursor = snapToNearestVertex(rawWorld.x, rawWorld.y);
      document.getElementById('cursor-coords').textContent =
        `X: ${currentWorldCursor.x.toFixed(2)}m, Y: ${currentWorldCursor.y.toFixed(2)}m`;
      checkHoverEntity(currentWorldCursor.x, currentWorldCursor.y);
      if (traceStartPt) drawCadCanvas();
    }
  });

  window.addEventListener('mouseup', () => {
    isPanning = false;
  });

  canvas.addEventListener('wheel', e => {
    e.preventDefault();
    const rect = canvas.getBoundingClientRect();
    const mx = e.clientX - rect.left;
    const my = e.clientY - rect.top;

    const zoomFactor = e.deltaY < 0 ? 1.15 : 0.87;
    const newScale = Math.min(Math.max(viewScale * zoomFactor, 1.5), 350);

    viewOffsetX = mx - (mx - viewOffsetX) * (newScale / viewScale);
    viewOffsetY = my - (my - viewOffsetY) * (newScale / viewScale);
    viewScale = newScale;
    drawCadCanvas();
  }, { passive: false });

  canvas.addEventListener('click', () => {
    if (mouseMovedDist > 6) return; // Was a drag pan, not a click
    handleCanvasClick(currentWorldCursor.x, currentWorldCursor.y);
  });
}

function snapToNearestVertex(wx, wy) {
  if (!currentDrawing || !currentDrawing.clean_plan || !currentDrawing.clean_plan.snap_points) {
    return { x: wx, y: wy, snapped: false };
  }
  const snapTol = Math.max(0.15, 8 / viewScale);
  let bestDist = snapTol;
  let bestPt = null;
  for (const pt of currentDrawing.clean_plan.snap_points) {
    const d = Math.hypot(wx - pt[0], wy - pt[1]);
    if (d < bestDist) {
      bestDist = d;
      bestPt = pt;
    }
  }
  if (bestPt) {
    return { x: bestPt[0], y: bestPt[1], snapped: true };
  }
  return { x: wx, y: wy, snapped: false };
}

function handleCanvasClick(wx, wy) {
  if (!currentDrawing) return;

  // 1. INSPECT / 1-CLICK ROOM OR LINE SELECTION
  if (activeCanvasTool === 'INSPECT') {
    if (hoveredEntity && hoveredEntity.group_id) {
      toggleHighlightGroup(hoveredEntity.group_id);
    }
    return;
  }

  // 2. COUNT STAMP TOOL (1 Click = +1 Count)
  if (activeCanvasTool === 'COUNT') {
    addInteractiveMeasurementToTarget({
      kind: 'COUNT',
      label: `Counted Symbol @ (${wx.toFixed(1)}m, ${wy.toFixed(1)}m)`,
      length: 1.0,
      breadth: 1.0,
      area: 0.0,
      centroid: [wx, wy],
      canvasEntity: {
        id: `USR-CNT-${Date.now()}`,
        type: 'INSERT',
        x: wx,
        y: wy,
        block_name: 'COUNT',
      },
    });
    return;
  }

  // 3. TWO-CLICK TOOLS: TRACE_AREA, TRACE_WALL, CALIBRATE
  if (!traceStartPt) {
    traceStartPt = { x: wx, y: wy };
    drawCadCanvas();
    return;
  }

  const p1 = traceStartPt;
  const p2 = { x: wx, y: wy };
  traceStartPt = null;

  if (activeCanvasTool === 'CALIBRATE') {
    const measuredDist = Math.hypot(p2.x - p1.x, p2.y - p1.y);
    if (measuredDist < 1e-4) return;
    const userRealMeters = prompt(
      `Measured distance between the 2 clicked points is currently ${measuredDist.toFixed(3)} m.\n\nEnter the ACTUAL real-world length in Meters (e.g. 4.50):`,
      measuredDist.toFixed(2)
    );
    if (!userRealMeters) {
      drawCadCanvas();
      return;
    }
    const targetMeters = parseFloat(userRealMeters);
    if (targetMeters > 0) {
      applyScaleCalibrationFactor(targetMeters / measuredDist);
    }
    return;
  }

  if (activeCanvasTool === 'TRACE_AREA') {
    const l = Math.max(0.1, Math.abs(p2.x - p1.x));
    const b = Math.max(0.1, Math.abs(p2.y - p1.y));
    const area = +(l * b).toFixed(3);
    const minX = Math.min(p1.x, p2.x), maxX = Math.max(p1.x, p2.x);
    const minY = Math.min(p1.y, p2.y), maxY = Math.max(p1.y, p2.y);
    const pts = [[minX, minY], [maxX, minY], [maxX, maxY], [minX, maxY]];

    addInteractiveMeasurementToTarget({
      kind: 'AREA',
      label: `Traced Room/Zone (${l.toFixed(2)}m × ${b.toFixed(2)}m)`,
      length: +l.toFixed(3),
      breadth: +b.toFixed(3),
      area: area,
      centroid: [+(minX + maxX) / 2, +(minY + maxY) / 2],
      canvasEntity: {
        id: `USR-AREA-${Date.now()}`,
        type: 'POLYGON',
        points: pts,
        closed: true,
        area: +area.toFixed(2),
        perimeter: +(2 * (l + b)).toFixed(2),
        centroid: [(minX + maxX) / 2, (minY + maxY) / 2],
      },
    });
    return;
  }

  if (activeCanvasTool === 'TRACE_WALL') {
    const len = Math.hypot(p2.x - p1.x, p2.y - p1.y);
    if (len < 0.05) return;
    addInteractiveMeasurementToTarget({
      kind: 'LENGTH',
      label: `Traced Wall / Run (${len.toFixed(2)}m)`,
      length: +len.toFixed(3),
      breadth: parseFloat(document.getElementById('active-target-b').value) || 0.23,
      area: 0.0,
      centroid: [(p1.x + p2.x) / 2, (p1.y + p2.y) / 2],
      canvasEntity: {
        id: `USR-WALL-${Date.now()}`,
        type: 'POLYLINE',
        points: [[p1.x, p1.y], [p2.x, p2.y]],
        closed: false,
        length: +len.toFixed(2),
        centroid: [(p1.x + p2.x) / 2, (p1.y + p2.y) / 2],
      },
    });
  }
}

function applyScaleCalibrationFactor(factor) {
  if (!currentDrawing || factor <= 0) return;
  // Scale bounds, canvas entities, text labels, and takeoff measurements by factor
  const b = currentDrawing.bounds;
  b.min_x = +(b.min_x * factor).toFixed(2);
  b.min_y = +(b.min_y * factor).toFixed(2);
  b.max_x = +(b.max_x * factor).toFixed(2);
  b.max_y = +(b.max_y * factor).toFixed(2);
  b.width = +(b.width * factor).toFixed(2);
  b.height = +(b.height * factor).toFixed(2);

  (currentDrawing.canvas_entities || []).forEach(ent => {
    if (ent.points) {
      ent.points = ent.points.map(p => [+(p[0] * factor).toFixed(3), +(p[1] * factor).toFixed(3)]);
    }
    if (ent.x !== undefined) ent.x = +(ent.x * factor).toFixed(3);
    if (ent.y !== undefined) ent.y = +(ent.y * factor).toFixed(3);
    if (ent.centroid) ent.centroid = [+(ent.centroid[0] * factor).toFixed(3), +(ent.centroid[1] * factor).toFixed(3)];
    if (ent.area) ent.area = +(ent.area * factor * factor).toFixed(2);
    if (ent.length) ent.length = +(ent.length * factor).toFixed(2);
  });

  (currentDrawing.text_labels || []).forEach(t => {
    t.x = +(t.x * factor).toFixed(3);
    t.y = +(t.y * factor).toFixed(3);
  });

  (currentDrawing.takeoff_groups || []).forEach(grp => {
    grp.measurements.forEach(m => {
      if (grp.calc_mode !== 'COUNT') {
        m.length = +(m.length * factor).toFixed(3);
        if (grp.calc_mode === 'AREA' || grp.calc_mode === 'VOLUME') {
          m.breadth = +(m.breadth * factor).toFixed(3);
          m.raw_area = +(m.raw_area * factor * factor).toFixed(3);
        }
      }
    });
  });

  document.getElementById('bounds-info').textContent =
    `Plan Extents: ${b.width.toFixed(1)}m × ${b.height.toFixed(1)}m (Calibrated)`;
  setCanvasTool('INSPECT');
  resetView();
  renderAllTables();
}

function addInteractiveMeasurementToTarget(item) {
  const dsrCode = document.getElementById('active-target-dsr').value || 'DSR-F01';
  const dsr = dsrByCode[dsrCode];
  if (!dsr) return;

  const userB = parseFloat(document.getElementById('active-target-b').value) || dsr.default_width || 1.0;
  const userD = parseFloat(document.getElementById('active-target-d').value) || dsr.default_depth || 1.0;

  let grp = currentDrawing.takeoff_groups.find(g => g.dsr_code === dsrCode);
  if (!grp) {
    grp = {
      group_id: `USR::${dsrCode}::${Date.now()}`,
      source_type: 'User Traced on 2D Plan',
      detection_basis: 'User Traced on 2D Plan',
      layer: 'USER-TAKEOFF',
      block_name: null,
      display_name: `${dsr.category.split('-').pop().trim()} (${dsrCode})`,
      entity_kind: item.kind,
      color: '#10b981',
      dsr_code: dsr.code,
      category: dsr.category,
      description: dsr.description,
      unit: dsr.unit,
      rate: dsr.rate,
      calc_mode: dsr.calc_mode,
      default_width: userB,
      default_depth: userD,
      is_ref: dsr.is_ref,
      included: true,
      measurements: [],
    };
    currentDrawing.takeoff_groups.unshift(grp);
  }

  const mId = item.canvasEntity.id;
  grp.measurements.push({
    ent_id: mId,
    label: item.label,
    nos: 1,
    length: item.length,
    breadth: item.kind === 'AREA' ? item.breadth : userB,
    depth: userD,
    raw_area: item.area || 0,
    raw_length: item.length || 0,
    centroid: item.centroid,
  });

  item.canvasEntity.group_id = grp.group_id;
  item.canvasEntity.layer = 'USER-TAKEOFF';
  item.canvasEntity.color = grp.color;
  item.canvasEntity.detection_basis = 'User Traced on 2D Plan';
  currentDrawing.canvas_entities.push(item.canvasEntity);

  highlightedGroupId = grp.group_id;
  renderAllTables();
  drawCadCanvas();
}

function clearAllGuessedItems() {
  if (!currentDrawing) return;
  currentDrawing.takeoff_groups = [];
  highlightedGroupId = null;
  renderAllTables();
  drawCadCanvas();
}

function deleteTakeoffGroup(gIdx) {
  if (!currentDrawing) return;
  const removed = currentDrawing.takeoff_groups.splice(gIdx, 1)[0];
  if (removed && highlightedGroupId === removed.group_id) {
    highlightedGroupId = null;
  }
  renderAllTables();
  drawCadCanvas();
}

function deleteSingleMeasurement(gIdx, mIdx) {
  if (!currentDrawing) return;
  const grp = currentDrawing.takeoff_groups[gIdx];
  if (!grp) return;
  grp.measurements.splice(mIdx, 1);
  if (grp.measurements.length === 0) {
    currentDrawing.takeoff_groups.splice(gIdx, 1);
  }
  renderAllTables();
  drawCadCanvas();
}

// ============================================================================
// COORDINATE TRANSFORMS & CRISP 2D PLAN RENDERING
// ============================================================================
function worldToScreen(wx, wy) {
  const ch = canvas.clientHeight;
  return {
    x: viewOffsetX + wx * viewScale,
    y: ch - (viewOffsetY + wy * viewScale),
  };
}

function screenToWorld(sx, sy) {
  const ch = canvas.clientHeight;
  return {
    x: (sx - viewOffsetX) / viewScale,
    y: (ch - sy - viewOffsetY) / viewScale,
  };
}

function resetView() {
  if (!currentDrawing || !canvas) return;
  const b = currentDrawing.bounds;
  const cw = canvas.clientWidth || 700;
  const ch = canvas.clientHeight || 500;
  const pad = 42;
  const w = Math.max(b.width, 1);
  const h = Math.max(b.height, 1);

  viewScale = Math.min((cw - pad * 2) / w, (ch - pad * 2) / h);
  viewOffsetX = (cw - w * viewScale) / 2 - b.min_x * viewScale;
  viewOffsetY = (ch - h * viewScale) / 2 - b.min_y * viewScale;
  drawCadCanvas();
}

function toggleLabels() {
  showLabels = !showLabels;
  document.getElementById('btn-toggle-labels').textContent = `Text Labels: ${showLabels ? 'ON' : 'OFF'}`;
  drawCadCanvas();
}

function toggleHighlightGroup(groupId) {
  highlightedGroupId = highlightedGroupId === groupId ? null : groupId;
  renderLayerPills();
  drawCadCanvas();
  highlightTableRow(highlightedGroupId);
}

function clearHighlight() {
  highlightedGroupId = null;
  renderLayerPills();
  drawCadCanvas();
  highlightTableRow(null);
}

function checkHoverEntity(wx, wy) {
  if (!currentDrawing) return;
  let found = null;

  for (const ent of currentDrawing.canvas_entities) {
    if (ent.type === 'INSERT') {
      const dist = Math.hypot(wx - ent.x, wy - ent.y);
      if (dist < Math.max(0.7, 10 / viewScale)) {
        found = ent;
        break;
      }
    }
  }

  if (!found) {
    for (const ent of currentDrawing.canvas_entities) {
      if (ent.type === 'POLYGON' && pointInPolygon(wx, wy, ent.points)) {
        found = ent;
        break;
      }
    }
  }

  if (found !== hoveredEntity) {
    hoveredEntity = found;
    const tip = document.getElementById('hover-tooltip');
    if (found && found.group_id) {
      const grp = currentDrawing.takeoff_groups.find(g => g.group_id === found.group_id);
      const basis = found.detection_basis || (grp ? grp.detection_basis : '') || 'Plan Geometry';
      document.getElementById('hover-title').textContent = grp ? grp.display_name : found.layer;
      document.getElementById('hover-subtitle').textContent = `Identified via: ${basis}`;
      if (found.type === 'POLYGON') {
        document.getElementById('hover-metrics').textContent = `Exact Area: ${found.area} Sq.m | Perimeter: ${found.perimeter} m`;
      } else if (found.type === 'POLYLINE') {
        document.getElementById('hover-metrics').textContent = `Running Length: ${found.length} Rmt`;
      } else if (found.type === 'INSERT') {
        document.getElementById('hover-metrics').textContent = `Symbol: ${found.block_name} @ (${found.x}m, ${found.y}m)`;
      }
      tip.classList.remove('hidden');
    } else {
      tip.classList.add('hidden');
    }
    drawCadCanvas();
  }
}

function pointInPolygon(x, y, pts) {
  let inside = false;
  for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) {
    const xi = pts[i][0], yi = pts[i][1];
    const xj = pts[j][0], yj = pts[j][1];
    const intersect = ((yi > y) !== (yj > y)) && (x < (xj - xi) * (y - yi) / ((yj - yi) || 1e-9) + xi);
    if (intersect) inside = !inside;
  }
  return inside;
}

function drawCadCanvas() {
  if (!ctx || !canvas || !currentDrawing) return;
  const cw = canvas.clientWidth;
  const ch = canvas.clientHeight;
  ctx.clearRect(0, 0, cw, ch);

  // 0. If PDF crisp background image exists, draw it aligned to world bounds!
  if (bgImageObj) {
    const b = currentDrawing.bounds;
    const topLeft = worldToScreen(b.min_x, b.max_y);
    const botRight = worldToScreen(b.max_x, b.min_y);
    ctx.save();
    ctx.globalAlpha = 0.88;
    ctx.drawImage(bgImageObj, topLeft.x, topLeft.y, botRight.x - topLeft.x, botRight.y - topLeft.y);
    ctx.restore();
  }

  const activeGroupIds = new Set(
    (currentDrawing.takeoff_groups || []).filter(g => g.included).map(g => g.group_id)
  );
  const entities = currentDrawing.canvas_entities || [];

  // 1. Draw Closed Polygons (Rooms / Lawns / Pavers)
  for (const ent of entities) {
    if (ent.type !== 'POLYGON') continue;
    const isIncluded = activeGroupIds.has(ent.group_id);
    const isDimmed = (highlightedGroupId && ent.group_id !== highlightedGroupId) || !isIncluded;
    const isHigh = highlightedGroupId && ent.group_id === highlightedGroupId;
    const isHov = hoveredEntity && hoveredEntity.id === ent.id;

    ctx.beginPath();
    ent.points.forEach((pt, idx) => {
      const s = worldToScreen(pt[0], pt[1]);
      if (idx === 0) ctx.moveTo(s.x, s.y);
      else ctx.lineTo(s.x, s.y);
    });
    ctx.closePath();

    ctx.fillStyle = hexToRgba(ent.color, isDimmed ? 0.04 : (isHigh || isHov ? 0.45 : 0.22));
    ctx.fill();

    ctx.lineWidth = isHigh || isHov ? 2.6 : 1.5;
    ctx.strokeStyle = isDimmed ? 'rgba(148,163,184,0.3)' : ent.color;
    ctx.stroke();

    if (showLabels && !isDimmed && ent.centroid) {
      const c = worldToScreen(ent.centroid[0], ent.centroid[1]);
      ctx.font = '700 10.5px JetBrains Mono, monospace';
      ctx.fillStyle = '#ffffff';
      ctx.textAlign = 'center';
      ctx.fillText(`${ent.area} m²`, c.x, c.y + 13);
    }
  }

  // 2. Draw Walls & Linear Runs
  for (const ent of entities) {
    if (ent.type !== 'POLYLINE' && ent.type !== 'LINE') continue;
    const isIncluded = activeGroupIds.has(ent.group_id);
    const isDimmed = (highlightedGroupId && ent.group_id !== highlightedGroupId) || !isIncluded;
    const isHigh = highlightedGroupId && ent.group_id === highlightedGroupId;

    ctx.beginPath();
    ent.points.forEach((pt, idx) => {
      const s = worldToScreen(pt[0], pt[1]);
      if (idx === 0) ctx.moveTo(s.x, s.y);
      else ctx.lineTo(s.x, s.y);
    });
    if (ent.closed) ctx.closePath();

    ctx.lineWidth = isHigh ? 3.8 : 2.2;
    ctx.strokeStyle = isDimmed ? 'rgba(148,163,184,0.3)' : ent.color;
    ctx.stroke();
  }

  // 3. Draw Symbols / Blocks / Circles
  for (const ent of entities) {
    if (ent.type !== 'INSERT') continue;
    const isIncluded = activeGroupIds.has(ent.group_id);
    const isDimmed = (highlightedGroupId && ent.group_id !== highlightedGroupId) || !isIncluded;
    const isHigh = highlightedGroupId && ent.group_id === highlightedGroupId;
    const isHov = hoveredEntity && hoveredEntity.id === ent.id;

    const s = worldToScreen(ent.x, ent.y);
    const rad = Math.max(5, Math.min(14, 0.55 * viewScale));

    ctx.beginPath();
    ctx.arc(s.x, s.y, rad, 0, Math.PI * 2);
    ctx.fillStyle = hexToRgba(ent.color, isDimmed ? 0.08 : (isHigh || isHov ? 0.75 : 0.40));
    ctx.fill();
    ctx.lineWidth = isHigh || isHov ? 2.5 : 1.5;
    ctx.strokeStyle = isDimmed ? 'rgba(148,163,184,0.3)' : ent.color;
    ctx.stroke();
  }

  // 4. Draw Room / Zone Text Annotations
  if (showLabels && currentDrawing.text_labels) {
    ctx.font = '600 10.5px Inter, sans-serif';
    ctx.textAlign = 'left';
    for (const lbl of currentDrawing.text_labels) {
      const s = worldToScreen(lbl.x, lbl.y);
      ctx.fillStyle = '#f8fafc';
      ctx.fillText(lbl.text, s.x, s.y);
    }
  }

  // 5. Draw Live Rubberband Preview when Tracing Area, Wall, or Calibrating Scale
  if (traceStartPt) {
    const p1 = worldToScreen(traceStartPt.x, traceStartPt.y);
    const p2 = worldToScreen(currentWorldCursor.x, currentWorldCursor.y);

    ctx.save();
    if (activeCanvasTool === 'TRACE_AREA') {
      const rx = Math.min(p1.x, p2.x);
      const ry = Math.min(p1.y, p2.y);
      const rw = Math.abs(p2.x - p1.x);
      const rh = Math.abs(p2.y - p1.y);
      ctx.fillStyle = 'rgba(16, 185, 129, 0.28)';
      ctx.fillRect(rx, ry, rw, rh);
      ctx.strokeStyle = '#10b981';
      ctx.lineWidth = 2.2;
      ctx.setLineDash([5, 4]);
      ctx.strokeRect(rx, ry, rw, rh);

      const l = Math.abs(currentWorldCursor.x - traceStartPt.x);
      const b = Math.abs(currentWorldCursor.y - traceStartPt.y);
      ctx.fillStyle = '#ffffff';
      ctx.font = '700 11px JetBrains Mono, monospace';
      ctx.fillText(`${l.toFixed(2)}m × ${b.toFixed(2)}m = ${(l * b).toFixed(2)} m²`, rx + 6, ry - 6);
    } else {
      ctx.beginPath();
      ctx.moveTo(p1.x, p1.y);
      ctx.lineTo(p2.x, p2.y);
      ctx.strokeStyle = activeCanvasTool === 'CALIBRATE' ? '#f59e0b' : '#38bdf8';
      ctx.lineWidth = 2.8;
      ctx.setLineDash([6, 4]);
      ctx.stroke();

      const dist = Math.hypot(currentWorldCursor.x - traceStartPt.x, currentWorldCursor.y - traceStartPt.y);
      ctx.fillStyle = '#fbbf24';
      ctx.font = '700 11px JetBrains Mono, monospace';
      ctx.fillText(`L = ${dist.toFixed(2)} m`, (p1.x + p2.x) / 2 + 8, (p1.y + p2.y) / 2 - 8);
    }
    ctx.restore();
  }

  // Draw snap vertex indicator if cursor is snapped
  if (currentWorldCursor && currentWorldCursor.snapped) {
    const sp = worldToScreen(currentWorldCursor.x, currentWorldCursor.y);
    ctx.strokeStyle = '#fbbf24';
    ctx.lineWidth = 2;
    ctx.strokeRect(sp.x - 5, sp.y - 5, 10, 10);
  }
}

function renderLayerPills() {
  const container = document.getElementById('layer-pills-container');
  if (!container || !currentDrawing) return;

  // Show both Takeoff Groups AND Clickable Text Callouts with dimensions
  let html = currentDrawing.takeoff_groups.map(grp => {
    const isSelected = highlightedGroupId === grp.group_id;
    const count = grp.measurements.length;
    return `
      <button onclick="toggleHighlightGroup('${grp.group_id}')"
        class="px-2.5 py-1 rounded-lg text-[11px] font-semibold flex items-center gap-1.5 border transition ${
          isSelected
            ? 'bg-indigo-600 border-indigo-600 text-white shadow-sm'
            : 'bg-slate-100 border-slate-200 text-slate-700 hover:bg-slate-200'
        }">
        <span class="w-2.5 h-2.5 rounded-full shrink-0" style="background:${grp.color}"></span>
        <span class="truncate max-w-[175px]">${grp.display_name}</span>
        <span class="text-[10px] px-1.5 py-0.2 rounded bg-white/90 text-slate-800 mono">${count} &bull; ${numFormatter.format(grp.total_qty)} ${grp.unit}</span>
      </button>
    `;
  }).join('');

  container.innerHTML = html;
}

function highlightTableRow(groupId) {
  document.querySelectorAll('[data-group-row]').forEach(row => {
    if (groupId && row.getAttribute('data-group-row') === groupId) {
      row.classList.add('bg-indigo-50', 'ring-1', 'ring-indigo-400');
      row.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    } else {
      row.classList.remove('bg-indigo-50', 'ring-1', 'ring-indigo-400');
    }
  });
}

// ============================================================================
// RECALCULATION & LIVE SYNCHRONIZATION
// ============================================================================
function recalculateClientSide() {
  if (!currentDrawing) return;
  let totalEntities = 0;

  for (const grp of currentDrawing.takeoff_groups) {
    const mode = grp.calc_mode || 'AREA';
    const w = parseFloat(grp.default_width) || 1.0;
    const d = parseFloat(grp.default_depth) || 1.0;
    const rate = parseFloat(grp.rate) || 0.0;

    let grpQty = 0;
    for (const m of grp.measurements) {
      totalEntities += 1;
      const nos = parseFloat(m.nos) || 1;
      let qty = 0;

      if (mode === 'COUNT') {
        m.breadth = 1.0;
        m.depth = 1.0;
        qty = nos;
      } else if (mode === 'AREA') {
        m.depth = 1.0;
        qty = nos * (parseFloat(m.length) || 0) * (parseFloat(m.breadth) || 1);
      } else if (mode === 'VOLUME') {
        qty = nos * (parseFloat(m.length) || 0) * (parseFloat(m.breadth) || 1) * (parseFloat(m.depth) || d);
      } else if (mode === 'LENGTH') {
        m.breadth = 1.0;
        m.depth = 1.0;
        qty = nos * (parseFloat(m.length) || 0);
      } else if (mode === 'WALL_AREA') {
        m.breadth = 1.0;
        qty = nos * (parseFloat(m.length) || 0) * (parseFloat(m.depth) || d);
      } else if (mode === 'WALL_VOLUME') {
        qty = nos * (parseFloat(m.length) || 0) * (parseFloat(m.breadth) || w) * (parseFloat(m.depth) || d);
      } else {
        qty = nos * (parseFloat(m.length) || 0) * (parseFloat(m.breadth) || 1) * (parseFloat(m.depth) || 1);
      }

      m.quantity = Math.round(qty * 1000) / 1000;
      grpQty += m.quantity;
    }

    grp.total_qty = Math.round(grpQty * 100) / 100;
    grp.amount = grp.included ? Math.round(grp.total_qty * rate * 100) / 100 : 0;
  }

  const activeGroups = currentDrawing.takeoff_groups.filter(g => g.included);
  const baseTotal = activeGroups.reduce((acc, g) => acc + g.amount, 0);
  const contPct = (parseFloat(document.getElementById('meta-contingency').value) || 0) / 100.0;
  const gstPct = (parseFloat(document.getElementById('meta-gst').value) || 0) / 100.0;
  const contAmt = baseTotal * contPct;
  const taxable = baseTotal + contAmt;
  const gstAmt = taxable * gstPct;
  const grandTotal = Math.round(taxable + gstAmt);

  const lCountEl = document.getElementById('kpi-layer-count');
  const eCountEl = document.getElementById('kpi-entity-count');
  const gtEl = document.getElementById('header-grand-total');
  if (lCountEl) lCountEl.textContent = activeGroups.length;
  if (eCountEl) eCountEl.textContent = totalEntities;
  if (gtEl) gtEl.textContent = inrFormatter.format(grandTotal);
}

function renderAllTables() {
  if (!currentDrawing) return;
  recalculateClientSide();
  renderMappingTable();
  renderTakeoffTable();
  renderBoqTable();
  renderAbstractTable();
  renderBbsTable();
  renderLayerPills();
}

// STEP 2 TAB 1: TAKEOFF ITEMS & DEPTH CONFIG
function renderMappingTable() {
  const tbody = document.getElementById('mapping-table-body');
  if (!tbody) return;

  if (currentDrawing.takeoff_groups.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="8" class="py-8 text-center text-slate-400 text-xs">
          Takeoff list is empty. Use the <strong>📐 2-Click Box Area</strong>, <strong>🧱 2-Click Wall Length</strong>, or <strong>📍 Click to Count</strong> tools on the 2D drawing to add exact measurements.
        </td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = currentDrawing.takeoff_groups.map((grp, gIdx) => {
    const optionsHtml = dsrLibrary.map(it =>
      `<option value="${it.code}" ${it.code === grp.dsr_code ? 'selected' : ''}>
        [${it.code}] ${it.category.split('-').pop().trim()} — ${it.description.slice(0, 45)}... (${it.unit})
      </option>`
    ).join('');

    return `
      <tr data-group-row="${grp.group_id}" class="hover:bg-slate-50 transition">
        <td class="py-2 px-2 text-center">
          <input type="checkbox" ${grp.included ? 'checked' : ''}
            onchange="onToggleGroupInclude(${gIdx}, this.checked)"
            class="rounded border-slate-300 text-indigo-600 focus:ring-0 cursor-pointer" />
        </td>
        <td class="py-2 px-2">
          <button onclick="toggleHighlightGroup('${grp.group_id}')" class="flex items-center gap-2 text-left hover:underline">
            <span class="w-2.5 h-2.5 rounded-full shrink-0" style="background:${grp.color}"></span>
            <div>
              <div class="font-bold text-slate-800 leading-tight">${grp.display_name}</div>
              <div class="text-[10px] text-slate-500 mt-0.5">
                <span class="px-1.5 py-0.2 rounded bg-slate-100 text-slate-700 font-medium">${grp.detection_basis || grp.source_type}</span>
                &bull; ${grp.measurements.length} rows
              </div>
            </div>
          </button>
        </td>
        <td class="py-2 px-2">
          <select onchange="onChangeGroupDsr(${gIdx}, this.value)"
            class="w-full bg-white border border-slate-300 rounded-lg px-2 py-1 text-[11px] text-slate-800 focus:border-indigo-500 focus:outline-none">
            ${optionsHtml}
          </select>
        </td>
        <td class="py-2 px-1.5 text-right">
          <input type="number" step="0.05" value="${grp.default_width}"
            onchange="onChangeGroupDim(${gIdx}, 'width', this.value)"
            class="w-14 bg-white border border-slate-300 rounded px-1 py-0.5 text-right mono text-xs text-slate-900" />
        </td>
        <td class="py-2 px-1.5 text-right">
          <input type="number" step="0.05" value="${grp.default_depth}"
            onchange="onChangeGroupDim(${gIdx}, 'depth', this.value)"
            class="w-14 bg-white border border-slate-300 rounded px-1 py-0.5 text-right mono text-xs text-slate-900" />
        </td>
        <td class="py-2 px-2 text-right mono font-bold text-indigo-700">
          ${numFormatter.format(grp.total_qty)} <span class="text-[10px] text-slate-500 font-normal">${grp.unit}</span>
        </td>
        <td class="py-2 px-2 text-right mono font-bold text-slate-900">
          ${inrFormatter.format(grp.amount)}
        </td>
        <td class="py-2 px-1.5 text-center">
          <button onclick="deleteTakeoffGroup(${gIdx})" class="text-slate-400 hover:text-rose-600 font-bold text-sm" title="Remove this group">&times;</button>
        </td>
      </tr>
    `;
  }).join('');
}

function onToggleGroupInclude(gIdx, checked) {
  currentDrawing.takeoff_groups[gIdx].included = checked;
  renderAllTables();
  drawCadCanvas();
}

function onChangeGroupDsr(gIdx, newCode) {
  const grp = currentDrawing.takeoff_groups[gIdx];
  const dsr = dsrByCode[newCode];
  if (!dsr) return;
  grp.dsr_code = dsr.code;
  grp.category = dsr.category;
  grp.description = dsr.description;
  grp.unit = dsr.unit;
  grp.rate = dsr.rate;
  grp.calc_mode = dsr.calc_mode;
  grp.default_width = dsr.default_width;
  grp.default_depth = dsr.default_depth;
  grp.is_ref = dsr.is_ref;

  for (const m of grp.measurements) {
    if (grp.calc_mode === 'WALL_VOLUME') {
      m.breadth = grp.default_width;
      m.depth = grp.default_depth;
    } else if (grp.calc_mode === 'WALL_AREA' || grp.calc_mode === 'VOLUME') {
      m.depth = grp.default_depth;
    }
  }
  renderAllTables();
}

function onChangeGroupDim(gIdx, field, val) {
  const num = parseFloat(val) || 1.0;
  const grp = currentDrawing.takeoff_groups[gIdx];
  if (field === 'width') {
    grp.default_width = num;
    grp.measurements.forEach(m => { m.breadth = num; });
  } else {
    grp.default_depth = num;
    grp.measurements.forEach(m => { m.depth = num; });
  }
  renderAllTables();
}

// STEP 2 TAB 2: DETAILED L x B x D MEASUREMENT SHEET
function renderTakeoffTable() {
  const tbody = document.getElementById('takeoff-table-body');
  if (!tbody) return;

  let html = '';
  let itemNo = 0;

  currentDrawing.takeoff_groups.forEach((grp, gIdx) => {
    if (!grp.included) return;
    itemNo += 1;

    html += `
      <tr class="bg-slate-100 border-t border-slate-300">
        <td class="py-1.5 px-2 font-bold text-indigo-700 mono">${itemNo}.0</td>
        <td colspan="5" class="py-1.5 px-2 font-bold text-slate-900">
          [${grp.dsr_code}] ${grp.display_name}
        </td>
        <td class="py-1.5 px-2 text-right font-bold text-indigo-700 mono">${numFormatter.format(grp.total_qty)} ${grp.unit}</td>
        <td></td>
      </tr>
    `;

    grp.measurements.forEach((m, mIdx) => {
      html += `
        <tr class="hover:bg-slate-50">
          <td class="py-1 px-2 text-slate-400 mono text-[11px] pl-3">${itemNo}.${mIdx + 1}</td>
          <td class="py-1 px-2">
            <input type="text" value="${m.label}"
              onchange="onEditMeasurement(${gIdx}, ${mIdx}, 'label', this.value)"
              class="w-full bg-transparent border-b border-transparent focus:border-indigo-400 text-slate-800 focus:outline-none" />
          </td>
          <td class="py-1 px-1.5 text-right">
            <input type="number" step="1" value="${m.nos}"
              onchange="onEditMeasurement(${gIdx}, ${mIdx}, 'nos', this.value)"
              class="w-12 bg-white border border-slate-200 rounded px-1 py-0.5 text-right mono text-slate-900" />
          </td>
          <td class="py-1 px-1.5 text-right">
            <input type="number" step="0.05" value="${m.length}"
              onchange="onEditMeasurement(${gIdx}, ${mIdx}, 'length', this.value)"
              class="w-14 bg-white border border-slate-200 rounded px-1 py-0.5 text-right mono text-slate-900" />
          </td>
          <td class="py-1 px-1.5 text-right">
            <input type="number" step="0.05" value="${m.breadth}"
              onchange="onEditMeasurement(${gIdx}, ${mIdx}, 'breadth', this.value)"
              class="w-14 bg-white border border-slate-200 rounded px-1 py-0.5 text-right mono text-slate-900" />
          </td>
          <td class="py-1 px-1.5 text-right">
            <input type="number" step="0.05" value="${m.depth}"
              onchange="onEditMeasurement(${gIdx}, ${mIdx}, 'depth', this.value)"
              class="w-14 bg-white border border-slate-200 rounded px-1 py-0.5 text-right mono text-slate-900" />
          </td>
          <td class="py-1 px-2 text-right mono font-semibold text-slate-900">${numFormatter.format(m.quantity)}</td>
          <td class="py-1 px-1.5 text-center">
            <button onclick="deleteSingleMeasurement(${gIdx}, ${mIdx})" class="text-slate-400 hover:text-rose-600 font-bold">&times;</button>
          </td>
        </tr>
      `;
    });
  });

  tbody.innerHTML = html;
}

function onEditMeasurement(gIdx, mIdx, field, val) {
  const m = currentDrawing.takeoff_groups[gIdx].measurements[mIdx];
  if (field === 'label') {
    m.label = val;
  } else {
    m[field] = parseFloat(val) || 0;
    m.raw_area = 0;
  }
  renderAllTables();
}

// ============================================================================
// STEP 3 TABLES: BOQ, ABSTRACT & BBS
// ============================================================================
function renderBoqTable() {
  const tbody = document.getElementById('boq-table-body');
  if (!tbody) return;

  const byCat = groupIncludedByCategory();
  let html = '';
  let sno = 0;

  Object.keys(byCat).forEach((cat, cIdx) => {
    const items = byCat[cat];
    let catTotal = 0;

    html += `
      <tr class="bg-indigo-50/80 border-y border-indigo-100">
        <td colspan="7" class="py-2 px-4 font-bold text-indigo-900 uppercase tracking-wider text-xs">
          Part ${String.fromCharCode(65 + cIdx)} &mdash; ${cat}
        </td>
      </tr>
    `;

    items.forEach(grp => {
      sno += 1;
      catTotal += grp.amount;
      html += `
        <tr class="hover:bg-slate-50">
          <td class="py-2.5 px-3 text-center font-bold text-slate-500 mono">${sno}</td>
          <td class="py-2.5 px-3">
            <div class="font-bold text-indigo-700 mono">${grp.dsr_code}</div>
            <div class="text-[10px] text-slate-500">${grp.is_ref || 'IS 1200'}</div>
          </td>
          <td class="py-2.5 px-4">
            <div class="font-bold text-slate-900 mb-0.5">${grp.display_name}</div>
            <div class="text-slate-600 text-xs leading-relaxed">${grp.description}</div>
          </td>
          <td class="py-2.5 px-3 text-right font-bold text-indigo-700 mono">${numFormatter.format(grp.total_qty)}</td>
          <td class="py-2.5 px-2 text-center font-semibold text-slate-700">${grp.unit}</td>
          <td class="py-2.5 px-3 text-right mono text-slate-800">${inrFormatter.format(grp.rate)}</td>
          <td class="py-2.5 px-4 text-right font-bold text-slate-900 mono">${inrFormatter.format(grp.amount)}</td>
        </tr>
      `;
    });

    html += `
      <tr class="bg-slate-100 font-bold border-t border-slate-300">
        <td colspan="6" class="py-2 px-4 text-right text-slate-700">Sub-Total for Part ${String.fromCharCode(65 + cIdx)} (${cat}):</td>
        <td class="py-2 px-4 text-right text-emerald-700 mono">${inrFormatter.format(catTotal)}</td>
      </tr>
    `;
  });

  tbody.innerHTML = html;
}

function renderAbstractTable() {
  const tbody = document.getElementById('abstract-table-body');
  if (!tbody) return;

  const byCat = groupIncludedByCategory();
  const cats = Object.keys(byCat);
  let baseTotal = 0;
  const catSums = cats.map(c => {
    const sum = byCat[c].reduce((acc, g) => acc + g.amount, 0);
    baseTotal += sum;
    return { cat: c, count: byCat[c].length, sum };
  });

  let html = '';
  catSums.forEach((row, idx) => {
    const share = baseTotal > 0 ? ((row.sum / baseTotal) * 100).toFixed(1) : '0.0';
    html += `
      <tr class="hover:bg-slate-50">
        <td class="py-3 px-4 text-center font-bold text-indigo-700 mono">Part ${String.fromCharCode(65 + idx)}</td>
        <td class="py-3 px-4 font-bold text-slate-800">${row.cat}</td>
        <td class="py-3 px-4 text-center mono text-slate-600">${row.count}</td>
        <td class="py-3 px-4 text-right font-bold text-slate-900 mono">${inrFormatter.format(row.sum)}</td>
        <td class="py-3 px-4 text-right mono text-indigo-600 font-semibold">${share}%</td>
      </tr>
    `;
  });

  const contPct = parseFloat(document.getElementById('meta-contingency').value) || 0;
  const gstPct = parseFloat(document.getElementById('meta-gst').value) || 0;
  const contAmt = baseTotal * (contPct / 100.0);
  const taxable = baseTotal + contAmt;
  const gstAmt = taxable * (gstPct / 100.0);
  const grandTotal = Math.round(taxable + gstAmt);

  html += `
    <tr class="bg-slate-100 font-bold border-t-2 border-slate-300 text-slate-900">
      <td colspan="3" class="py-2.5 px-4 text-right">(A) SUB-TOTAL BASE WORK COST:</td>
      <td class="py-2.5 px-4 text-right mono">${inrFormatter.format(baseTotal)}</td>
      <td class="py-2.5 px-4 text-right mono">100.0%</td>
    </tr>
    <tr class="text-slate-700">
      <td colspan="3" class="py-2 px-4 text-right">(B) Add: Physical &amp; Unforeseen Contingencies @ ${contPct.toFixed(1)}% on (A):</td>
      <td class="py-2 px-4 text-right mono">${inrFormatter.format(contAmt)}</td>
      <td class="py-2 px-4 text-right mono text-slate-500">+${contPct.toFixed(1)}%</td>
    </tr>
    <tr class="bg-slate-50 font-bold text-slate-900">
      <td colspan="3" class="py-2 px-4 text-right">(C) TAXABLE PROJECT COST [(A) + (B)]:</td>
      <td class="py-2 px-4 text-right mono">${inrFormatter.format(taxable)}</td>
      <td></td>
    </tr>
    <tr class="text-slate-700">
      <td colspan="3" class="py-2 px-4 text-right">(D) Add: Goods &amp; Services Tax (CGST + SGST) @ ${gstPct.toFixed(1)}% on (C):</td>
      <td class="py-2 px-4 text-right mono">${inrFormatter.format(gstAmt)}</td>
      <td class="py-2 px-4 text-right mono text-slate-500">+${gstPct.toFixed(1)}%</td>
    </tr>
    <tr class="bg-emerald-50 border-t-2 border-emerald-500 text-emerald-900 font-bold text-sm">
      <td colspan="3" class="py-3.5 px-4 text-right">(E) GRAND TOTAL ESTIMATED PROJECT COST (INCL. GST):</td>
      <td class="py-3.5 px-4 text-right mono text-base text-emerald-700">${inrFormatter.format(grandTotal)}</td>
      <td class="py-3.5 px-4 text-center text-xs uppercase">IS 1200</td>
    </tr>
  `;

  tbody.innerHTML = html;
}

function groupIncludedByCategory() {
  const map = {};
  if (!currentDrawing) return map;
  for (const g of currentDrawing.takeoff_groups) {
    if (!g.included || !g.measurements.length) continue;
    const cat = g.category || 'General Works';
    if (!map[cat]) map[cat] = [];
    map[cat].push(g);
  }
  return map;
}

// ============================================================================
// BAR BENDING SCHEDULE (BBS) & VISUAL STRUCTURAL STEEL CALCULATOR
// ============================================================================
const BBS_SHAPES = {
  'SC-00': { name: 'Straight', b90: 0, b45: 0, b135: 0, svg: 'M 10 18 L 90 18' },
  'SC-01': { name: 'L-Bend (90°)', b90: 1, b45: 0, b135: 0, svg: 'M 15 6 L 15 26 L 88 26' },
  'SC-02': { name: 'U-Bar (2×90°)', b90: 2, b45: 0, b135: 0, svg: 'M 15 8 L 15 26 L 85 26 L 85 8' },
  'SC-51': { name: 'Stirrup (135°)', b90: 3, b45: 0, b135: 2, svg: 'M 24 6 L 76 6 L 76 28 L 24 28 L 24 6 M 24 6 L 32 14' },
  'SC-41': { name: 'Cranked Slab Bar', b90: 0, b45: 2, b135: 0, svg: 'M 6 10 L 24 10 L 36 26 L 64 26 L 76 10 L 94 10' },
  'SC-99': { name: 'Chair Bar (4×90°)', b90: 4, b45: 0, b135: 0, svg: 'M 8 26 L 28 26 L 28 10 L 72 10 L 72 26 L 92 26' },
};

function recalculateBbsClientSide() {
  if (!currentDrawing || !currentDrawing.bbs || !currentDrawing.bbs.bbs_rows) return;
  const rows = currentDrawing.bbs.bbs_rows;
  const diaWeights = { 8: 0, 10: 0, 12: 0, 16: 0, 20: 0, 25: 0, 32: 0 };

  rows.forEach(r => {
    const dia = parseInt(r.dia_mm, 10) || 12;
    const sc = r.shape_code || 'SC-00';
    const sh = BBS_SHAPES[sc] || BBS_SHAPES['SC-00'];
    const mbrs = parseInt(r.num_members, 10) || 1;
    const bpm = parseInt(r.bars_per_member, 10) || 1;
    r.total_bars = mbrs * bpm;

    const a = parseFloat(r.a_m) || 0;
    const b = parseFloat(r.b_m) || 0;
    const c = parseFloat(r.c_m) || 0;
    const d = parseFloat(r.d_m) || 0;
    const dMeters = dia / 1000.0;

    const ded = (sh.b90 * 2 + sh.b45 * 1 + sh.b135 * 3) * dMeters;
    r.bend_deduction_m = Math.round(ded * 1000) / 1000;

    const gross = sc === 'SC-51' ? 2.0 * (a + b) + c + d : a + b + c + d;
    r.cutting_length_m = Math.max(0.10, Math.round((gross - r.bend_deduction_m) * 1000) / 1000);
    r.total_length_m = Math.round(r.total_bars * r.cutting_length_m * 100) / 100;
    r.unit_weight = Math.round(((dia * dia) / 162.2) * 1000) / 1000;
    r.total_weight_kg = Math.round(r.total_length_m * r.unit_weight * 100) / 100;
    r.svg_path = sh.svg;

    if (diaWeights[dia] === undefined) diaWeights[dia] = 0;
    diaWeights[dia] = Math.round((diaWeights[dia] + r.total_weight_kg) * 100) / 100;
  });

  const netKg = Object.values(diaWeights).reduce((acc, v) => acc + v, 0);
  const wastageKg = netKg * 0.03;
  const grossKg = netKg + wastageKg;
  const grossMt = grossKg / 1000.0;

  currentDrawing.bbs.summary = {
    dia_weights_kg: diaWeights,
    net_weight_kg: Math.round(netKg * 100) / 100,
    wastage_kg: Math.round(wastageKg * 100) / 100,
    gross_weight_kg: Math.round(grossKg * 100) / 100,
    gross_weight_mt: Math.round(grossMt * 1000) / 1000,
  };
}

function renderBbsTable() {
  if (!currentDrawing || !currentDrawing.bbs) return;
  recalculateBbsClientSide();
  updateQuickFootingCalc();

  const sum = currentDrawing.bbs.summary;
  const cardsContainer = document.getElementById('bbs-dia-cards');
  if (cardsContainer) {
    const dias = [8, 10, 12, 16, 20, 25];
    let cardsHtml = dias.map(d => {
      const wt = sum.dia_weights_kg[d] || 0;
      return `
        <div class="bg-white border border-slate-200 rounded-xl p-2.5 shadow-xs">
          <div class="text-[10px] text-slate-500 font-semibold">${d} mm Ø TMT</div>
          <div class="text-xs font-bold text-indigo-700 mono mt-0.5">${numFormatter.format(wt)} kg</div>
          <div class="text-[9px] text-slate-400 mono">${((d * d) / 162.2).toFixed(3)} kg/m</div>
        </div>
      `;
    }).join('');

    cardsHtml += `
      <div class="bg-white border border-slate-300 rounded-xl p-2.5 shadow-xs">
        <div class="text-[10px] text-slate-600 font-semibold">Gross (+3% Wastage)</div>
        <div class="text-xs font-bold text-slate-900 mono mt-0.5">${numFormatter.format(sum.gross_weight_kg)} kg</div>
        <div class="text-[9px] text-slate-400 mono">Net: ${numFormatter.format(sum.net_weight_kg)} kg</div>
      </div>
      <div class="bg-amber-50 border border-amber-300 rounded-xl p-2.5 shadow-xs">
        <div class="text-[10px] text-amber-800 font-bold uppercase">Total Steel (MT)</div>
        <div class="text-sm font-bold text-amber-900 mono mt-0.5">${sum.gross_weight_mt.toFixed(3)} MT</div>
        <div class="text-[9px] text-amber-700">Fe-500D IS 1786</div>
      </div>
    `;
    cardsContainer.innerHTML = cardsHtml;
  }

  const tbody = document.getElementById('bbs-table-body');
  if (!tbody) return;

  const rows = currentDrawing.bbs.bbs_rows || [];
  let html = '';
  let lastMember = null;

  rows.forEach((r, idx) => {
    if (r.member_group !== lastMember) {
      lastMember = r.member_group;
      html += `
        <tr class="bg-slate-100 border-y border-slate-300">
          <td colspan="14" class="py-1.5 px-3 font-bold text-slate-800 text-xs uppercase">
            ${r.member_group}
          </td>
        </tr>
      `;
    }

    html += `
      <tr class="hover:bg-slate-50">
        <td class="py-1.5 px-2.5 font-bold text-indigo-700 mono">${r.bar_mark}</td>
        <td class="py-1.5 px-2.5 text-slate-800">${r.description}</td>
        <td class="py-1.5 px-2 text-center">
          <svg viewBox="0 0 100 34" class="w-24 h-7 mx-auto bg-slate-900 rounded border border-slate-300">
            <path d="${r.svg_path || 'M 10 18 L 90 18'}" fill="none" stroke="#f59e0b" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
        </td>
        <td class="py-1.5 px-1.5 text-center font-bold mono text-slate-900">${r.dia_mm} mm</td>
        <td class="py-1.5 px-1.5 text-right mono">${r.num_members}</td>
        <td class="py-1.5 px-1.5 text-right mono">${r.bars_per_member}</td>
        <td class="py-1.5 px-1.5 text-right font-bold mono">${r.total_bars}</td>
        <td class="py-1.5 px-1.5 text-right mono">${r.a_m}</td>
        <td class="py-1.5 px-1.5 text-right mono">${r.b_m}</td>
        <td class="py-1.5 px-1.5 text-right mono">${r.c_m}</td>
        <td class="py-1.5 px-1.5 text-right mono">${r.d_m}</td>
        <td class="py-1.5 px-1.5 text-right mono text-rose-600">-${r.bend_deduction_m.toFixed(3)}</td>
        <td class="py-1.5 px-2 text-right font-bold text-indigo-700 mono">${r.cutting_length_m.toFixed(3)}</td>
        <td class="py-1.5 px-2.5 text-right font-bold text-emerald-700 mono">${numFormatter.format(r.total_weight_kg)}</td>
      </tr>
    `;
  });

  tbody.innerHTML = html;
}

let lastQuickSteelResult = null;

function updateQuickFootingCalc() {
  const svg = document.getElementById('app-footing-svg');
  const box = document.getElementById('qf-breakdown');
  if (!svg || !box) return;

  const lx = parseFloat(document.getElementById('qf-lx').value) || 2.0;
  const ly = parseFloat(document.getElementById('qf-ly').value) || 1.5;
  const d = parseFloat(document.getElementById('qf-d').value) || 0.45;
  const dia = parseInt(document.getElementById('qf-dia').value, 10) || 12;
  const spc = Math.max(0.05, parseFloat(document.getElementById('qf-spc').value) || 0.15);
  const cov = parseFloat(document.getElementById('qf-cov').value) || 0.05;
  const nos = parseInt(document.getElementById('qf-nos').value, 10) || 1;

  const netX = Math.max(0.1, lx - 2 * cov);
  const netY = Math.max(0.1, ly - 2 * cov);
  const leg = Math.max(0.05, d - 2 * cov);

  const cutLenX = +(netX + 2 * leg).toFixed(3);
  const numBarsX = Math.floor(netY / spc) + 1;
  const totLenX = +(numBarsX * cutLenX).toFixed(2);

  const cutLenY = +(netY + 2 * leg).toFixed(3);
  const numBarsY = Math.floor(netX / spc) + 1;
  const totLenY = +(numBarsY * cutLenY).toFixed(2);

  const unitWt = (dia * dia) / 162.2;
  const totLenOne = +(totLenX + totLenY).toFixed(2);
  const totKgAll = +(totLenOne * unitWt * nos).toFixed(2);

  lastQuickSteelResult = {
    lx, ly, d, dia, spc, cov, nos,
    cutLenX, numBarsX, totLenX,
    cutLenY, numBarsY, totLenY,
    unitWt: +unitWt.toFixed(3),
    totKgAll,
  };

  box.innerHTML = `
    <div class="text-rose-400 font-bold">1. Longitudinal Rod (Along Lx = ${lx}m):</div>
    <div class="text-slate-300 pl-2">
      Cut Len = (${lx} - 2&times;${cov}) + 2&times;${leg.toFixed(2)} = <span class="text-white font-bold">${cutLenX}m</span>
      &bull; Bars = (${netY.toFixed(2)}/${spc})+1 = <span class="text-white font-bold">${numBarsX} Nos</span> &rarr; <span class="text-rose-300 font-bold">${totLenX} m</span>
    </div>
    <div class="text-sky-400 font-bold pt-1">2. Latitude Rod (Along Ly = ${ly}m):</div>
    <div class="text-slate-300 pl-2">
      Cut Len = (${ly} - 2&times;${cov}) + 2&times;${leg.toFixed(2)} = <span class="text-white font-bold">${cutLenY}m</span>
      &bull; Bars = (${netX.toFixed(2)}/${spc})+1 = <span class="text-white font-bold">${numBarsY} Nos</span> &rarr; <span class="text-sky-300 font-bold">${totLenY} m</span>
    </div>
    <div class="border-t border-slate-700 pt-1.5 mt-1 flex items-center justify-between">
      <span class="text-slate-300">Direct Qty (${nos} Nos &times; ${totLenOne}m &times; ${unitWt.toFixed(3)} kg/m):</span>
      <span class="text-sm font-bold text-emerald-400">${numFormatter.format(totKgAll)} kg</span>
    </div>
  `;

  svg.innerHTML = `
    <text x="14" y="16" fill="#94a3b8" font-size="10" font-weight="bold">PLAN VIEW (TOP) — LONGITUDINAL (RED) vs LATITUDE (BLUE)</text>
    <rect x="20" y="24" width="210" height="125" fill="rgba(30,41,59,0.5)" stroke="#94a3b8" stroke-width="1.8"/>
    <rect x="30" y="34" width="190" height="105" fill="none" stroke="#64748b" stroke-dasharray="3,3" stroke-width="1"/>
    <g stroke="#f43f5e" stroke-width="2">
      <line x1="30" y1="46" x2="220" y2="46"/><line x1="30" y1="66" x2="220" y2="66"/>
      <line x1="30" y1="86" x2="220" y2="86"/><line x1="30" y1="106" x2="220" y2="106"/>
      <line x1="30" y1="126" x2="220" y2="126"/>
    </g>
    <g stroke="#38bdf8" stroke-width="2">
      <line x1="50" y1="34" x2="50" y2="139"/><line x1="80" y1="34" x2="80" y2="139"/>
      <line x1="110" y1="34" x2="110" y2="139"/><line x1="140" y1="34" x2="140" y2="139"/>
      <line x1="170" y1="34" x2="170" y2="139"/><line x1="200" y1="34" x2="200" y2="139"/>
    </g>
    <rect x="102" y="68" width="46" height="36" fill="#0f172a" stroke="#fbbf24" stroke-width="1.5"/>
    <text x="125" y="164" fill="#f43f5e" font-size="10.5" font-weight="bold" text-anchor="middle">&larr; Longitudinal Lx = ${lx} m (${numBarsX} Bars) &rarr;</text>
    <text x="242" y="88" fill="#38bdf8" font-size="10.5" font-weight="bold" transform="rotate(90, 242, 88)" text-anchor="middle">&larr; Latitude Ly = ${ly} m &rarr;</text>

    <rect x="265" y="24" width="245" height="60" rx="4" fill="rgba(244,63,94,0.12)" stroke="#f43f5e" stroke-width="1"/>
    <text x="274" y="40" fill="#f43f5e" font-size="10.5" font-weight="bold">LONGITUDINAL ROD (X-SIDE): ${cutLenX} m</text>
    <text x="274" y="55" fill="#e2e8f0" font-size="10">Placed across Ly (${ly}m) &rarr; ${numBarsX} Bars &times; ${cutLenX}m = ${totLenX}m</text>
    <text x="274" y="70" fill="#fbbf24" font-size="10">L-Bend both sides = 2 &times; ${leg.toFixed(2)}m</text>

    <rect x="265" y="92" width="245" height="60" rx="4" fill="rgba(56,189,248,0.12)" stroke="#38bdf8" stroke-width="1"/>
    <text x="274" y="108" fill="#38bdf8" font-size="10.5" font-weight="bold">LATITUDE ROD (Y-SIDE): ${cutLenY} m</text>
    <text x="274" y="123" fill="#e2e8f0" font-size="10">Placed across Lx (${lx}m) &rarr; ${numBarsY} Bars &times; ${cutLenY}m = ${totLenY}m</text>
    <text x="274" y="138" fill="#fbbf24" font-size="10">Unit Weight (${dia}mm) = ${unitWt.toFixed(3)} kg/m</text>

    <text x="14" y="182" fill="#94a3b8" font-size="10" font-weight="bold">SECTION VIEW — U-BAR CUTTING LENGTH = (L - 2&times;Cover) + 2&times;(D - 2&times;Cover)</text>
    <rect x="20" y="188" width="290" height="36" fill="rgba(30,41,59,0.55)" stroke="#94a3b8" stroke-width="1.5"/>
    <path d="M 32 194 L 32 218 L 298 218 L 298 194" fill="none" stroke="#f43f5e" stroke-width="2.5"/>
    <g fill="#38bdf8">
      <circle cx="55" cy="214" r="3"/><circle cx="90" cy="214" r="3"/><circle cx="125" cy="214" r="3"/>
      <circle cx="160" cy="214" r="3"/><circle cx="195" cy="214" r="3"/><circle cx="230" cy="214" r="3"/><circle cx="265" cy="214" r="3"/>
    </g>
    <text x="325" y="204" fill="#10b981" font-size="11" font-weight="bold">TOTAL = ${totKgAll} kg (${nos} Footings)</text>
    <text x="325" y="219" fill="#94a3b8" font-size="10">Cover = ${(cov*1000).toFixed(0)}mm | Spacing = ${(spc*1000).toFixed(0)}mm c/c</text>
  `;
}

function pushQuickSteelToBoq() {
  if (!currentDrawing || !lastQuickSteelResult) return;
  const r = lastQuickSteelResult;
  const customId = `STEEL::TENDER::${Date.now()}`;
  currentDrawing.takeoff_groups.unshift({
    group_id: customId,
    source_type: 'Direct Structural Steel (Tender)',
    detection_basis: `Footing ${r.lx}x${r.ly}x${r.d}m (${r.dia}mm Ø @ ${r.spc*1000}mm c/c)`,
    layer: 'S-REBAR-TENDER',
    block_name: null,
    display_name: `TMT Fe-500D Reinforcement Steel (${r.dia}mm Ø Footing Mesh ${r.lx}×${r.ly}m)`,
    entity_kind: 'MANUAL',
    color: '#f59e0b',
    dsr_code: 'DSR-S01',
    category: 'RCC & Structural Works',
    description: `Supplying, cutting, bending and placing in position High Yield Strength Deformed (HYSD) Fe-500D TMT reinforcement bars (${r.dia}mm Ø) for Footings/RCC works including 18-gauge annealed binding wire (IS 1786 / CPWD DSR 5.22.6).`,
    unit: 'Kg',
    rate: 88.0,
    calc_mode: 'CUSTOM',
    default_width: 1.0,
    default_depth: 1.0,
    is_ref: 'IS 1200 (Part 5) / CPWD 5.22',
    included: true,
    measurements: [
      {
        ent_id: `ST-X-${Date.now()}`,
        label: `Longitudinal X-Bars (${r.nos} Footings × ${r.numBarsX} Bars)`,
        nos: r.nos * r.numBarsX,
        length: r.cutLenX,
        breadth: r.unitWt,
        depth: 1.0,
        raw_area: 0,
        raw_length: r.totLenX,
        centroid: [0, 0],
      },
      {
        ent_id: `ST-Y-${Date.now()}`,
        label: `Latitude Y-Bars (${r.nos} Footings × ${r.numBarsY} Bars)`,
        nos: r.nos * r.numBarsY,
        length: r.cutLenY,
        breadth: r.unitWt,
        depth: 1.0,
        raw_area: 0,
        raw_length: r.totLenY,
        centroid: [0, 0],
      },
    ],
  });
  renderAllTables();
  switchStep2Tab('items');
}

// ============================================================================
// MANUAL ITEM MODAL & EXCEL EXPORT
// ============================================================================
function openAddManualModal() {
  document.getElementById('custom-item-modal').classList.remove('hidden');
}

function closeAddManualModal() {
  document.getElementById('custom-item-modal').classList.add('hidden');
}

function confirmAddManualItem() {
  if (!currentDrawing) return;
  const code = document.getElementById('modal-dsr-select').value || 'DSR-CUST';
  const cat = document.getElementById('modal-category').value || 'General Works';
  const label = document.getElementById('modal-label').value || 'Manual Entry';
  const desc = document.getElementById('modal-desc').value || 'Custom item';
  const nos = parseFloat(document.getElementById('modal-nos').value) || 1;
  const len = parseFloat(document.getElementById('modal-len').value) || 1;
  const wid = parseFloat(document.getElementById('modal-wid').value) || 1;
  const dep = parseFloat(document.getElementById('modal-dep').value) || 1;
  const unit = document.getElementById('modal-unit').value || 'Sq.m';
  const rate = parseFloat(document.getElementById('modal-rate').value) || 0;

  const customId = `MANUAL::${Date.now()}`;
  currentDrawing.takeoff_groups.unshift({
    group_id: customId,
    source_type: 'Manual Entry',
    detection_basis: 'Manual Entry',
    layer: 'MANUAL',
    block_name: null,
    display_name: `${label} (${code})`,
    entity_kind: 'MANUAL',
    color: '#4f46e5',
    dsr_code: code,
    category: cat,
    description: desc,
    unit: unit,
    rate: rate,
    calc_mode: 'CUSTOM',
    default_width: wid,
    default_depth: dep,
    is_ref: 'IS 1200 / Custom',
    included: true,
    measurements: [
      {
        ent_id: `M-${Date.now()}`,
        label: label,
        nos: nos,
        length: len,
        breadth: wid,
        depth: dep,
        raw_area: 0,
        raw_length: 0,
        centroid: [0, 0],
      },
    ],
  });

  closeAddManualModal();
  renderAllTables();
}

async function exportExcelWorkbook() {
  if (!currentDrawing) return;
  const btn = document.getElementById('btn-export-excel');
  const origHtml = btn.innerHTML;
  btn.textContent = 'Generating Excel...';
  btn.disabled = true;

  try {
    const payload = {
      project_name: document.getElementById('meta-project-name').value,
      client_name: document.getElementById('meta-client-name').value,
      drawing_no: document.getElementById('meta-drawing-no').value,
      project_type: currentDrawing.meta ? currentDrawing.meta.project_type : 'All',
      contingency_pct: parseFloat(document.getElementById('meta-contingency').value) || 3.0,
      gst_pct: parseFloat(document.getElementById('meta-gst').value) || 18.0,
      takeoff_groups: currentDrawing.takeoff_groups,
      bbs_rows: currentDrawing.bbs ? currentDrawing.bbs.bbs_rows : [],
    };

    const res = await fetch('/api/export-excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });

    if (!res.ok) throw new Error('Failed to generate Excel workbook');
    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const cleanName = (payload.project_name || 'AutoBOQ').replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 35);
    a.download = `${cleanName}_BOQ_Takeoff_Abstract.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
  } catch (err) {
    alert('Excel export error: ' + err.message);
  } finally {
    btn.innerHTML = origHtml;
    btn.disabled = false;
  }
}

function hexToRgba(hex, alpha) {
  const clean = (hex || '#10b981').replace('#', '');
  const num = parseInt(clean, 16);
  const r = (num >> 16) & 255;
  const g = (num >> 8) & 255;
  const b = num & 255;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}
