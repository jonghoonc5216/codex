const canvas = document.getElementById("scene");
const ctx = canvas.getContext("2d", { alpha: false });

const els = {
  fileName: document.getElementById("fileName"),
  epsgSelect: document.getElementById("epsgSelect"),
  openBtn: document.getElementById("openBtn"),
  projectSaveBtn: document.getElementById("projectSaveBtn"),
  projectLoadBtn: document.getElementById("projectLoadBtn"),
  editBtn: document.getElementById("editBtn"),
  undoBtn: document.getElementById("undoBtn"),
  exportDxfBtn: document.getElementById("exportDxfBtn"),
  blenderBtn: document.getElementById("blenderBtn"),
  googleEarthBtn: document.getElementById("googleEarthBtn"),
  dxfInput: document.getElementById("dxfInput"),
  stateInput: document.getElementById("stateInput"),
  regenBtn: document.getElementById("regenBtn"),
  fitBtn: document.getElementById("fitBtn"),
  viewBtn: document.getElementById("viewBtn"),
  kmlLink: document.getElementById("kmlLink"),
  centerLat: document.getElementById("centerLat"),
  centerLon: document.getElementById("centerLon"),
  featureCount: document.getElementById("featureCount"),
  layerCount: document.getElementById("layerCount"),
  layerList: document.getElementById("layerList"),
  statusText: document.getElementById("statusText"),
  cursorText: document.getElementById("cursorText"),
  zoomRange: document.getElementById("zoomRange"),
  yawRange: document.getElementById("yawRange"),
  pitchRange: document.getElementById("pitchRange"),
  heightRange: document.getElementById("heightRange"),
  zoomValue: document.getElementById("zoomValue"),
  yawValue: document.getElementById("yawValue"),
  pitchValue: document.getElementById("pitchValue"),
  heightValue: document.getElementById("heightValue"),
  slopeCellRange: document.getElementById("slopeCellRange"),
  slopeCellValue: document.getElementById("slopeCellValue"),
  slopeBreaksInput: document.getElementById("slopeBreaksInput"),
  slopeOpacityRange: document.getElementById("slopeOpacityRange"),
  slopeOpacityValue: document.getElementById("slopeOpacityValue"),
  slopeAnalyzeBtn: document.getElementById("slopeAnalyzeBtn"),
  slopeToggleBtn: document.getElementById("slopeToggleBtn"),
  slopeClearBtn: document.getElementById("slopeClearBtn"),
  slopeLegend: document.getElementById("slopeLegend"),
  slopeStatus: document.getElementById("slopeStatus"),
  satelliteToggle: document.getElementById("satelliteToggle"),
  terrainToggle: document.getElementById("terrainToggle"),
  satelliteOpacityRange: document.getElementById("satelliteOpacityRange"),
  satelliteOpacityValue: document.getElementById("satelliteOpacityValue"),
  satelliteZoomSelect: document.getElementById("satelliteZoomSelect"),
  satelliteStatus: document.getElementById("satelliteStatus"),
  editPanel: document.getElementById("editPanel"),
  selectionText: document.getElementById("selectionText"),
  vertexXInput: document.getElementById("vertexXInput"),
  vertexYInput: document.getElementById("vertexYInput"),
  vertexZInput: document.getElementById("vertexZInput"),
  applyVertexBtn: document.getElementById("applyVertexBtn"),
  allZInput: document.getElementById("allZInput"),
  applyAllZBtn: document.getElementById("applyAllZBtn"),
  revertFeatureBtn: document.getElementById("revertFeatureBtn"),
  moveXInput: document.getElementById("moveXInput"),
  moveYInput: document.getElementById("moveYInput"),
  moveZInput: document.getElementById("moveZInput"),
  rotateInput: document.getElementById("rotateInput"),
  scaleInput: document.getElementById("scaleInput"),
  applyTransformBtn: document.getElementById("applyTransformBtn"),
  hatchStyleControls: document.getElementById("hatchStyleControls"),
  hatchColorInput: document.getElementById("hatchColorInput"),
  hatchOpacityRange: document.getElementById("hatchOpacityRange"),
  hatchOpacityValue: document.getElementById("hatchOpacityValue"),
  applyHatchStyleBtn: document.getElementById("applyHatchStyleBtn"),
  duplicateBtn: document.getElementById("duplicateBtn"),
  deleteBtn: document.getElementById("deleteBtn"),
  cadToggle: document.getElementById("cadToggle"),
  gridToggle: document.getElementById("gridToggle"),
  labelToggle: document.getElementById("labelToggle"),
  elevationToggle: document.getElementById("elevationToggle"),
  weightToggle: document.getElementById("weightToggle"),
  hatchToggle: document.getElementById("hatchToggle"),
};

let project = null;
let currentSourceId = null;
let bounds = null;
let visibleLayers = new Set();
let layerInputs = new Map();
let drag = null;
let drawPending = false;
let editMode = false;
let selectedFeatureId = null;
let selectedVertexIndex = -1;
let viewBeforeEdit = null;
let elevationLabelCells = new Set();
const hatchPatterns = new Map();
const geometryEdits = new Map();
const originalGeometry = new Map();
const undoStack = [];
let nextFeatureId = 1;
let addedFeatureIds = new Set();
let deletedFeatureIds = new Set();
const satellite = {
  manifest: null,
  images: new Map(),
  generation: 0,
  loaded: 0,
  failed: 0,
};
const cadImages = {
  images: new Map(),
  generation: 0,
  loaded: 0,
  failed: 0,
};
const DEFAULT_SLOPE_BREAKS = [5, 10, 16, 17, 25, 30];
const SLOPE_COLORS = ["#48b86f", "#96c75a", "#f2c94c", "#f2994a", "#eb5757", "#9b51e0", "#6d3a75"];
const MAX_SLOPE_CELLS = 320000;
let slopeAnalysis = {
  visible: false,
  cells: [],
  stats: [],
  breaks: [...DEFAULT_SLOPE_BREAKS],
  requestedCellSize: 2,
  cellSize: 2,
  sampleCount: 0,
};

const view = {
  zoom: 1,
  yaw: -28,
  pitch: 62,
  heightScale: 1,
  panX: 0,
  panY: 0,
};

const INITIAL_CRS_OPTIONS = [
  { epsg: 5179, name: "Korea 2000 / Unified CS" },
  { epsg: 5185, name: "Korea 2000 / West Belt 2010" },
  { epsg: 5186, name: "Korea 2000 / Central Belt 2010" },
  { epsg: 5187, name: "Korea 2000 / East Belt 2010" },
  { epsg: 5188, name: "Korea 2000 / East Sea Belt 2010" },
];

function resize() {
  const dpr = window.devicePixelRatio || 1;
  canvas.width = Math.floor(window.innerWidth * dpr);
  canvas.height = Math.floor(window.innerHeight * dpr);
  canvas.style.width = `${window.innerWidth}px`;
  canvas.style.height = `${window.innerHeight}px`;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  scheduleDraw();
}

function scheduleDraw() {
  if (drawPending) return;
  drawPending = true;
  requestAnimationFrame(() => {
    drawPending = false;
    draw();
  });
}

function metersWidth() {
  if (!bounds) return 1;
  return Math.max(bounds.xmax - bounds.xmin, bounds.ymax - bounds.ymin, 1);
}

function fitView() {
  if (!bounds) return;
  const sideSpace = window.innerWidth > 760 ? 660 : 80;
  const topSpace = 140;
  const availableW = Math.max(320, window.innerWidth - sideSpace);
  const availableH = Math.max(260, window.innerHeight - topSpace);
  const width = Math.max(bounds.xmax - bounds.xmin, 1);
  const height = Math.max(bounds.ymax - bounds.ymin, 1);
  view.zoom = Math.max(0.05, Math.min(12, Math.min(availableW / width, availableH / height) * 0.78));
  view.panX = 0;
  view.panY = 0;
  syncControls();
  scheduleDraw();
}

function normalizeYaw(value) {
  let result = value;
  while (result > 180) result -= 360;
  while (result < -180) result += 360;
  return result;
}

function syncControls() {
  view.yaw = normalizeYaw(view.yaw);
  els.zoomRange.value = String(view.zoom);
  els.yawRange.value = String(view.yaw);
  els.pitchRange.value = String(view.pitch);
  els.heightRange.value = String(view.heightScale);
  els.zoomValue.textContent = `${view.zoom.toFixed(2)}x`;
  els.yawValue.textContent = `${Math.round(view.yaw)}°`;
  els.pitchValue.textContent = `${Math.round(view.pitch)}°`;
  els.heightValue.textContent = `${view.heightScale.toFixed(1)}x`;
}

function formatNum(value, digits = 3) {
  return Number(value).toLocaleString("ko-KR", {
    maximumFractionDigits: digits,
    minimumFractionDigits: digits,
  });
}

function syncSlopeControls() {
  const cellSize = Number(els.slopeCellRange.value || 2);
  const opacity = Number(els.slopeOpacityRange.value || 0.58);
  els.slopeCellValue.textContent = `${cellSize}m`;
  els.slopeOpacityValue.textContent = `${Math.round(opacity * 100)}%`;
}

function parseSlopeBreaks() {
  const values = String(els.slopeBreaksInput.value || "")
    .split(/[,\s/]+/)
    .map((value) => Number(value.trim()))
    .filter((value) => Number.isFinite(value) && value > 0)
    .sort((first, second) => first - second);
  const unique = values.filter((value, index) => index === 0 || value !== values[index - 1]);
  if (!unique.length) return [...DEFAULT_SLOPE_BREAKS];
  return unique.slice(0, 10);
}

function slopeClassIndex(degrees, breaks) {
  for (let index = 0; index < breaks.length; index += 1) {
    if (degrees <= breaks[index]) return index;
  }
  return breaks.length;
}

function slopeClassLabel(index, breaks) {
  if (index === 0) return `0-${breaks[0]}도`;
  if (index >= breaks.length) return `${Math.floor(breaks[breaks.length - 1]) + 1}도 이상`;
  const lower = Math.floor(breaks[index - 1]) + 1;
  const upper = breaks[index];
  return lower === upper ? `${upper}도` : `${lower}-${upper}도`;
}

function slopeExportFeatures() {
  if (!slopeAnalysis.cells.length) return [];
  const opacity = Math.max(0.2, Math.min(0.85, Number(els.slopeOpacityRange.value || 0.58)));
  const rectangles = [];
  const active = new Map();

  const closeRun = (key) => {
    const run = active.get(key);
    if (!run) return;
    rectangles.push(run);
    active.delete(key);
  };

  const sortedCells = [...slopeAnalysis.cells].sort((first, second) => {
    if (Math.abs(first.y - second.y) > 0.0001) return first.y - second.y;
    return first.x - second.x;
  });
  let rowKey = "";
  let rowRuns = [];
  let rowRun = null;

  const flushRowRun = () => {
    if (!rowRun) return;
    rowRuns.push(rowRun);
    rowRun = null;
  };

  const mergeRowRuns = () => {
    if (!rowRuns.length) return;
    const seen = new Set();
    for (const run of rowRuns) {
      const key = `${run.classIndex}:${run.x0.toFixed(4)}:${run.x1.toFixed(4)}`;
      seen.add(key);
      const existing = active.get(key);
      if (existing && Math.abs(existing.y1 - run.y0) < 0.0001) {
        existing.y1 = run.y1;
      } else {
        closeRun(key);
        active.set(key, run);
      }
    }
    for (const key of Array.from(active.keys())) {
      if (!seen.has(key)) closeRun(key);
    }
    rowRuns = [];
  };

  for (const cell of sortedCells) {
    const nextRowKey = `${cell.y.toFixed(4)}:${cell.h.toFixed(4)}`;
    if (rowKey && nextRowKey !== rowKey) {
      flushRowRun();
      mergeRowRuns();
    }
    rowKey = nextRowKey;
    const sameRun =
      rowRun &&
      rowRun.classIndex === cell.classIndex &&
      Math.abs(rowRun.x1 - cell.x) < 0.0001;
    if (sameRun) {
      rowRun.x1 = cell.x + cell.w;
    } else {
      flushRowRun();
      rowRun = {
        classIndex: cell.classIndex,
        x0: cell.x,
        x1: cell.x + cell.w,
        y0: cell.y,
        y1: cell.y + cell.h,
      };
    }
  }
  flushRowRun();
  mergeRowRuns();
  for (const key of Array.from(active.keys())) closeRun(key);

  return rectangles.map((rect, index) => {
    const label = slopeClassLabel(rect.classIndex, slopeAnalysis.breaks);
    const points = [
      [rect.x0, rect.y0, 0],
      [rect.x1, rect.y0, 0],
      [rect.x1, rect.y1, 0],
      [rect.x0, rect.y1, 0],
      [rect.x0, rect.y0, 0],
    ];
    return {
      id: `slope-${index}`,
      kind: "hatch",
      layer: `경사분석 ${label}`,
      color: SLOPE_COLORS[rect.classIndex % SLOPE_COLORS.length],
      text: `경사 ${label}`,
      closed: true,
      pattern: "SOLID",
      opacity,
      points,
      paths: [points],
    };
  });
}

function isContourLayerName(layerName) {
  return /등고|contour/i.test(String(layerName || ""));
}

function finitePointZ(point) {
  const z = Number(point?.[2]);
  return Number.isFinite(z) ? z : null;
}

function collectContourSamples() {
  if (!project) return { samples: [], mode: "none" };
  const polylineFeatures = project.features.filter(
    (feature) => feature.kind === "polyline" && Array.isArray(feature.points) && feature.points.length > 1,
  );
  let terrainFeatures = polylineFeatures.filter((feature) =>
    isContourLayerName(feature.layer) && feature.points.some((point) => finitePointZ(point) !== null),
  );
  let mode = "등고선 레이어";
  if (!terrainFeatures.length) {
    terrainFeatures = polylineFeatures.filter((feature) =>
      feature.points.some((point) => Math.abs(finitePointZ(point) || 0) > 0.001),
    );
    mode = "Z값 폴리라인";
  }

  const rawSamples = [];
  for (const feature of terrainFeatures) {
    const points = feature.points;
    const step = Math.max(1, Math.floor(points.length / 800));
    for (let index = 0; index < points.length; index += step) {
      const point = points[index];
      const z = finitePointZ(point);
      if (z === null) continue;
      const x = Number(point[0]);
      const y = Number(point[1]);
      if (!Number.isFinite(x) || !Number.isFinite(y)) continue;
      rawSamples.push([x, y, z]);
    }
    const last = points[points.length - 1];
    const lastZ = finitePointZ(last);
    const lastX = Number(last?.[0]);
    const lastY = Number(last?.[1]);
    if (lastZ !== null && Number.isFinite(lastX) && Number.isFinite(lastY)) {
      rawSamples.push([lastX, lastY, lastZ]);
    }
  }
  const maxSamples = 90000;
  if (rawSamples.length <= maxSamples) return { samples: rawSamples, mode };
  const stride = Math.ceil(rawSamples.length / maxSamples);
  return {
    samples: rawSamples.filter((_, index) => index % stride === 0),
    mode: `${mode} 일부 샘플`,
  };
}

function sampleBounds(samples) {
  return samples.reduce(
    (box, point) => ({
      xmin: Math.min(box.xmin, point[0]),
      xmax: Math.max(box.xmax, point[0]),
      ymin: Math.min(box.ymin, point[1]),
      ymax: Math.max(box.ymax, point[1]),
      minZ: Math.min(box.minZ, point[2]),
      maxZ: Math.max(box.maxZ, point[2]),
    }),
    { xmin: Infinity, xmax: -Infinity, ymin: Infinity, ymax: -Infinity, minZ: Infinity, maxZ: -Infinity },
  );
}

function buildElevationSampler(samples) {
  const binSize = 40;
  const bins = new Map();
  for (const sample of samples) {
    const key = `${Math.floor(sample[0] / binSize)},${Math.floor(sample[1] / binSize)}`;
    if (!bins.has(key)) bins.set(key, []);
    bins.get(key).push(sample);
  }
  return (x, y) => {
    const cellX = Math.floor(x / binSize);
    const cellY = Math.floor(y / binSize);
    const candidates = [];
    for (let ring = 0; ring < 25; ring += 1) {
      for (let offsetX = -ring; offsetX <= ring; offsetX += 1) {
        for (let offsetY = -ring; offsetY <= ring; offsetY += 1) {
          if (ring && Math.abs(offsetX) !== ring && Math.abs(offsetY) !== ring) continue;
          const nearby = bins.get(`${cellX + offsetX},${cellY + offsetY}`);
          if (nearby) candidates.push(...nearby);
        }
      }
      if (candidates.length >= 16) break;
    }
    if (!candidates.length) return 0;
    const nearest = candidates
      .map((sample) => ({
        distance: (sample[0] - x) ** 2 + (sample[1] - y) ** 2,
        z: sample[2],
      }))
      .sort((first, second) => first.distance - second.distance)
      .slice(0, 12);
    if (nearest[0].distance < 0.01) return nearest[0].z;
    let weighted = 0;
    let totalWeight = 0;
    for (const item of nearest) {
      const weight = 1 / (item.distance + 4);
      weighted += item.z * weight;
      totalWeight += weight;
    }
    return totalWeight ? weighted / totalWeight : 0;
  };
}

function updateSlopeLegend() {
  els.slopeLegend.innerHTML = "";
  const breaks = slopeAnalysis.breaks;
  const totalArea = slopeAnalysis.stats.reduce((sum, item) => sum + item.area, 0);
  for (let index = 0; index <= breaks.length; index += 1) {
    const stat = slopeAnalysis.stats[index] || { area: 0, count: 0 };
    const row = document.createElement("div");
    row.className = "slope-legend-row";
    const chip = document.createElement("span");
    chip.className = "slope-chip";
    chip.style.background = SLOPE_COLORS[index % SLOPE_COLORS.length];
    const label = document.createElement("span");
    label.textContent = slopeClassLabel(index, breaks);
    const area = document.createElement("span");
    const percent = totalArea ? (stat.area / totalArea) * 100 : 0;
    area.textContent = `${(stat.area / 10000).toFixed(2)}ha · ${percent.toFixed(1)}%`;
    row.append(chip, label, area);
    els.slopeLegend.append(row);
  }
}

function clearSlopeAnalysis() {
  slopeAnalysis = {
    visible: false,
    cells: [],
    stats: [],
    breaks: parseSlopeBreaks(),
    requestedCellSize: Number(els.slopeCellRange.value || 2),
    cellSize: Number(els.slopeCellRange.value || 2),
    sampleCount: 0,
  };
  els.slopeToggleBtn.disabled = true;
  els.slopeClearBtn.disabled = true;
  els.slopeToggleBtn.textContent = "숨김";
  els.slopeLegend.innerHTML = "";
  els.slopeStatus.textContent = "경사분석 결과를 삭제했습니다.";
  scheduleDraw();
}

function toggleSlopeAnalysis() {
  slopeAnalysis.visible = !slopeAnalysis.visible;
  els.slopeToggleBtn.textContent = slopeAnalysis.visible ? "숨김" : "표시";
  scheduleDraw();
}

async function analyzeSlope() {
  if (!project || !bounds) {
    showError(new Error("먼저 DXF 파일을 불러와 주세요."));
    return;
  }
  syncSlopeControls();
  const requestedCellSize = Math.max(1, Number(els.slopeCellRange.value || 2));
  const breaks = parseSlopeBreaks();
  els.slopeBreaksInput.value = breaks.join(",");
  els.slopeAnalyzeBtn.disabled = true;
  els.slopeStatus.textContent = "등고선 Z값을 모아 경사면을 계산하는 중입니다.";
  await new Promise((resolve) => requestAnimationFrame(resolve));
  try {
    const { samples, mode } = collectContourSamples();
    if (samples.length < 12) {
      throw new Error("경사분석에 사용할 등고선 Z값을 충분히 찾지 못했습니다.");
    }
    const box = sampleBounds(samples);
    const width = Math.max(1, box.xmax - box.xmin);
    const height = Math.max(1, box.ymax - box.ymin);
    const autoCellSize = Math.ceil(Math.sqrt((width * height) / MAX_SLOPE_CELLS));
    const cellSize = Math.max(requestedCellSize, autoCellSize, 1);
    const columns = Math.max(1, Math.ceil(width / cellSize));
    const rows = Math.max(1, Math.ceil(height / cellSize));
    const sampler = buildElevationSampler(samples);
    const nodeColumns = columns + 1;
    const nodeRows = rows + 1;
    const elevations = new Float64Array(nodeColumns * nodeRows);
    for (let row = 0; row < nodeRows; row += 1) {
      const y = Math.min(box.ymax, box.ymin + row * cellSize);
      for (let column = 0; column < nodeColumns; column += 1) {
        const x = Math.min(box.xmax, box.xmin + column * cellSize);
        elevations[row * nodeColumns + column] = sampler(x, y);
      }
      if (row % 40 === 0) {
        els.slopeStatus.textContent = `표고 격자 계산 중 ${row.toLocaleString("ko-KR")}/${nodeRows.toLocaleString("ko-KR")}`;
        await new Promise((resolve) => setTimeout(resolve, 0));
      }
    }

    const cells = [];
    const stats = Array.from({ length: breaks.length + 1 }, () => ({ count: 0, area: 0 }));
    for (let row = 0; row < rows; row += 1) {
      const y = box.ymin + row * cellSize;
      const cellHeight = Math.min(cellSize, box.ymax - y);
      if (cellHeight <= 0) continue;
      for (let column = 0; column < columns; column += 1) {
        const x = box.xmin + column * cellSize;
        const cellWidth = Math.min(cellSize, box.xmax - x);
        if (cellWidth <= 0) continue;
        const topLeft = row * nodeColumns + column;
        const z00 = elevations[topLeft];
        const z10 = elevations[topLeft + 1];
        const z01 = elevations[topLeft + nodeColumns];
        const z11 = elevations[topLeft + nodeColumns + 1];
        const dzdx = ((z10 + z11) - (z00 + z01)) / (2 * cellWidth);
        const dzdy = ((z01 + z11) - (z00 + z10)) / (2 * cellHeight);
        const slopeDegrees = Math.atan(Math.hypot(dzdx, dzdy)) * 180 / Math.PI;
        const classIndex = slopeClassIndex(slopeDegrees, breaks);
        const area = cellWidth * cellHeight;
        stats[classIndex].count += 1;
        stats[classIndex].area += area;
        cells.push({
          x,
          y,
          w: cellWidth,
          h: cellHeight,
          z00,
          z10,
          z11,
          z01,
          classIndex,
          slope: slopeDegrees,
        });
      }
      if (row % 40 === 0) {
        els.slopeStatus.textContent = `경사도 분류 중 ${row.toLocaleString("ko-KR")}/${rows.toLocaleString("ko-KR")}`;
        await new Promise((resolve) => setTimeout(resolve, 0));
      }
    }
    slopeAnalysis = {
      visible: true,
      cells,
      stats,
      breaks,
      requestedCellSize,
      cellSize,
      sampleCount: samples.length,
    };
    updateSlopeLegend();
    els.slopeToggleBtn.disabled = false;
    els.slopeClearBtn.disabled = false;
    els.slopeToggleBtn.textContent = "숨김";
    const cellNotice = cellSize > requestedCellSize
      ? ` · 도면 범위가 커서 ${requestedCellSize}m 요청을 ${cellSize}m로 자동 조정`
      : "";
    els.slopeStatus.textContent =
      `${mode} ${samples.length.toLocaleString("ko-KR")}점 · ${cellSize}m 셀 ${cells.length.toLocaleString("ko-KR")}개${cellNotice}`;
    scheduleDraw();
  } catch (error) {
    showError(error);
  } finally {
    els.slopeAnalyzeBtn.disabled = false;
  }
}

function ensureCrsOptions(options, selectedEpsg) {
  if (!els.epsgSelect.options.length) {
    const automatic = document.createElement("option");
    automatic.value = "auto";
    automatic.textContent = "자동 감지";
    els.epsgSelect.append(automatic);
    for (const option of options) {
      const item = document.createElement("option");
      item.value = option.epsg;
      item.textContent = `EPSG:${option.epsg} ${option.name.replace("Korea 2000 / ", "")}`;
      els.epsgSelect.append(item);
    }
  }
  els.epsgSelect.value = String(selectedEpsg);
}

function setProjectActionsEnabled(enabled) {
  for (const control of [
    els.projectSaveBtn,
    els.editBtn,
    els.exportDxfBtn,
    els.blenderBtn,
    els.googleEarthBtn,
    els.slopeAnalyzeBtn,
    els.regenBtn,
    els.fitBtn,
    els.viewBtn,
  ]) {
    control.disabled = !enabled;
  }
}

function initializeEmptyProject() {
  project = null;
  currentSourceId = null;
  bounds = null;
  visibleLayers.clear();
  layerInputs.clear();
  cadImages.generation += 1;
  cadImages.images.clear();
  cadImages.loaded = 0;
  cadImages.failed = 0;
  geometryEdits.clear();
  originalGeometry.clear();
  undoStack.length = 0;
  addedFeatureIds = new Set();
  deletedFeatureIds = new Set();
  selectedFeatureId = null;
  selectedVertexIndex = -1;
  slopeAnalysis = {
    visible: false,
    cells: [],
    stats: [],
    breaks: parseSlopeBreaks(),
    requestedCellSize: Number(els.slopeCellRange.value || 2),
    cellSize: Number(els.slopeCellRange.value || 2),
    sampleCount: 0,
  };
  editMode = false;
  els.editBtn.classList.remove("active");
  els.editBtn.setAttribute("aria-pressed", "false");
  els.editPanel.hidden = true;
  canvas.classList.remove("editing");
  els.fileName.textContent = "DXF를 불러오세요";
  els.centerLat.textContent = "-";
  els.centerLon.textContent = "-";
  els.featureCount.textContent = "-";
  els.layerCount.textContent = "-";
  els.layerList.innerHTML = "";
  els.cursorText.textContent = "X -, Y -";
  els.kmlLink.removeAttribute("href");
  els.slopeLegend.innerHTML = "";
  els.slopeStatus.textContent = "등고선 Z값이 있는 DXF를 불러온 뒤 분석하세요.";
  els.slopeToggleBtn.disabled = true;
  els.slopeClearBtn.disabled = true;
  syncSlopeControls();
  ensureCrsOptions(INITIAL_CRS_OPTIONS, "auto");
  setProjectActionsEnabled(false);
  updateUndoButton();
  els.statusText.textContent = "빈 프로젝트 · DXF 열기를 눌러 파일을 선택하세요.";
  scheduleDraw();
}

function buildLayers(savedVisibleLayers = null) {
  visibleLayers.clear();
  layerInputs.clear();
  els.layerList.innerHTML = "";
  const restored = savedVisibleLayers ? new Set(savedVisibleLayers) : null;
  for (const layer of project.layers) {
    const visible = restored ? restored.has(layer.name) : true;
    if (visible) visibleLayers.add(layer.name);
    const row = document.createElement("label");
    row.className = "layer-item";
    const input = document.createElement("input");
    input.type = "checkbox";
    input.checked = visible;
    input.addEventListener("change", () => {
      if (input.checked) visibleLayers.add(layer.name);
      else visibleLayers.delete(layer.name);
      scheduleDraw();
    });
    layerInputs.set(layer.name, input);
    const swatch = document.createElement("span");
    swatch.className = "swatch";
    swatch.style.background = layer.color;
    const name = document.createElement("span");
    name.className = "layer-name";
    name.textContent = layer.name;
    const count = document.createElement("span");
    count.className = "layer-count";
    count.textContent = layer.count.toLocaleString("ko-KR");
    row.append(input, swatch, name, count);
    els.layerList.append(row);
  }
}

function clonePoints(points) {
  return points.map((point) => [...point]);
}

function featureSnapshot(feature) {
  return {
    points: clonePoints(feature.points),
    paths: feature.paths
      ? feature.paths.map((path) => clonePoints(path))
      : undefined,
    color: feature.color,
    opacity: feature.opacity,
  };
}

function restoreFeatureSnapshot(feature, snapshot) {
  feature.points = clonePoints(snapshot.points || []);
  feature.paths = snapshot.paths
    ? snapshot.paths.map((path) => clonePoints(path))
    : undefined;
  if (snapshot.color) feature.color = snapshot.color;
  if (snapshot.opacity !== undefined) feature.opacity = snapshot.opacity;
}

function refreshHatchPoints(feature) {
  if (feature.kind === "hatch" && feature.paths) {
    feature.points = feature.paths.flatMap((path) => clonePoints(path));
  }
}

function mapFeaturePoints(feature, mapper) {
  if (feature.kind === "hatch" && feature.paths) {
    feature.paths = feature.paths.map((path) =>
      path.map((point, index) => mapper(point, index)),
    );
    refreshHatchPoints(feature);
  } else {
    feature.points = feature.points.map((point, index) => mapper(point, index));
  }
}

function setFeaturePoint(feature, flatIndex, values) {
  if (feature.kind !== "hatch" || !feature.paths) {
    feature.points[flatIndex] = [...values];
    if (feature.closed && flatIndex === 0 && feature.points.length > 1) {
      feature.points[feature.points.length - 1] = [...values];
    }
    return;
  }
  let offset = 0;
  for (const path of feature.paths) {
    if (flatIndex < offset + path.length) {
      const pathIndex = flatIndex - offset;
      const wasClosed =
        path.length > 1 &&
        path[0].slice(0, 2).every(
          (value, axis) => Math.abs(value - path[path.length - 1][axis]) < 0.000001,
        );
      path[pathIndex] = [...values];
      if (pathIndex === 0 && wasClosed) {
        path[path.length - 1] = [...values];
      }
      refreshHatchPoints(feature);
      return;
    }
    offset += path.length;
  }
}

function restoreProjectChanges(state) {
  geometryEdits.clear();
  originalGeometry.clear();
  addedFeatureIds = new Set();
  deletedFeatureIds = new Set(state?.deletedFeatureIds || []);
  if (deletedFeatureIds.size) {
    project.features = project.features.filter((feature) => !deletedFeatureIds.has(feature.id));
  }
  for (const savedFeature of state?.addedFeatures || []) {
    if (!savedFeature || !Array.isArray(savedFeature.points)) continue;
    const feature = cloneFeature(savedFeature);
    project.features.push(feature);
    addedFeatureIds.add(feature.id);
  }
  for (const edit of state?.geometryEdits || []) {
    const feature = project.features.find((item) => item.id === edit.id);
    if (!feature || !["polyline", "hatch"].includes(feature.kind) || !Array.isArray(edit.points)) continue;
    originalGeometry.set(feature.id, featureSnapshot(feature));
    restoreFeatureSnapshot(feature, edit);
    geometryEdits.set(feature.id, { id: feature.id, ...featureSnapshot(feature) });
  }
  nextFeatureId = Math.max(0, ...project.features.map((feature) => Number(feature.id) || 0)) + 1;
  undoStack.length = 0;
  updateUndoButton();
}

function applyProject(data, options = {}) {
  project = data;
  currentSourceId = data.sourceId;
  bounds = data.bounds.projected;
  setProjectActionsEnabled(true);
  const automaticCrs = data.crsDetection?.mode === "auto";
  ensureCrsOptions(data.crsOptions, automaticCrs ? "auto" : data.epsg);
  const automaticOption = els.epsgSelect.querySelector('option[value="auto"]');
  if (automaticOption) {
    automaticOption.textContent = automaticCrs
      ? `자동 감지 → EPSG:${data.epsg}`
      : "자동 감지";
  }
  els.fileName.textContent =
    `${data.fileName} · ${automaticCrs ? "자동 " : ""}EPSG:${data.epsg}`;
  els.centerLat.textContent = formatNum(data.center.wgs84.lat, 6);
  els.centerLon.textContent = formatNum(data.center.wgs84.lon, 6);
  els.featureCount.textContent = data.stats.features.toLocaleString("ko-KR");
  els.layerCount.textContent = data.stats.layers.toLocaleString("ko-KR");
  els.kmlLink.href = `/api/export.kml?epsg=${data.epsg}&source=${encodeURIComponent(currentSourceId)}`;

  const saved = options.savedState || null;
  restoreProjectChanges(saved);
  buildLayers(saved?.visibleLayers || null);
  loadCadImages();
  slopeAnalysis = {
    visible: false,
    cells: [],
    stats: [],
    breaks: parseSlopeBreaks(),
    requestedCellSize: Number(els.slopeCellRange.value || 2),
    cellSize: Number(els.slopeCellRange.value || 2),
    sampleCount: 0,
  };
  els.slopeLegend.innerHTML = "";
  els.slopeStatus.textContent = "경사분석 버튼을 누르면 등고선 Z값으로 색상면을 만듭니다.";
  els.slopeToggleBtn.disabled = true;
  els.slopeClearBtn.disabled = true;
  els.slopeToggleBtn.textContent = "숨김";
  syncSlopeControls();
  if (saved) {
    Object.assign(view, saved.view || {});
    els.gridToggle.checked = saved.toggles?.grid ?? true;
    els.cadToggle.checked = saved.toggles?.cad ?? true;
    els.labelToggle.checked = saved.toggles?.labels ?? false;
    els.elevationToggle.checked = saved.toggles?.elevations ?? false;
    els.weightToggle.checked = saved.toggles?.weight ?? true;
    els.hatchToggle.checked = saved.toggles?.hatch ?? true;
    els.satelliteToggle.checked = false;
    els.terrainToggle.checked = false;
    els.satelliteOpacityRange.value = String(saved.satellite?.opacity ?? 0.82);
    els.satelliteZoomSelect.value = String(saved.satellite?.zoom ?? 17);
    els.slopeBreaksInput.value = saved.slope?.breaks || els.slopeBreaksInput.value;
    els.slopeCellRange.value = String(saved.slope?.cellSize ?? els.slopeCellRange.value);
    els.slopeOpacityRange.value = String(saved.slope?.opacity ?? els.slopeOpacityRange.value);
    syncSlopeControls();
    syncControls();
    scheduleDraw();
  } else if (options.resetView !== false) {
    view.yaw = -28;
    view.pitch = 62;
    view.heightScale = 1;
    fitView();
  } else {
    syncControls();
    scheduleDraw();
  }
  const hatchCount = data.features.filter((feature) => feature.kind === "hatch").length;
  const imageCount = data.features.filter((feature) => feature.kind === "image").length;
  const missingImageCount = data.features.filter((feature) => feature.kind === "image" && !feature.image?.available).length;
  selectedFeatureId = null;
  selectedVertexIndex = -1;
  updateSelectionPanel();
  const imageNotice = imageCount
    ? ` · 이미지 ${imageCount.toLocaleString("ko-KR")}개${missingImageCount ? `(${missingImageCount.toLocaleString("ko-KR")}개 파일 없음)` : ""}`
    : "";
  els.statusText.textContent =
    `${data.stats.points.toLocaleString("ko-KR")}개 좌표 · 해치 ${hatchCount.toLocaleString("ko-KR")}개${imageNotice} · ${data.crsDetection?.reason || `EPSG:${data.epsg}`}`;
  els.featureCount.textContent = project.features.length.toLocaleString("ko-KR");
}

async function fetchProject({
  sourceId = currentSourceId,
  epsg = els.epsgSelect.value || "auto",
  resetView = true,
  savedState = null,
  regen = false,
} = {}) {
  els.statusText.textContent = "DXF 분석 중";
  const params = new URLSearchParams({ epsg: String(epsg) });
  if (sourceId) params.set("source", sourceId);
  if (regen) params.set("regen", "1");
  const response = await fetch(`/api/project?${params}`);
  if (!response.ok) {
    const error = await response.json().catch(() => ({ error: response.statusText }));
    throw new Error(error.error || response.statusText);
  }
  applyProject(await response.json(), { resetView, savedState });
}

async function regenerateProject() {
  if (!project || !currentSourceId) return;
  const savedState = {
    view: { ...view },
    visibleLayers: [...visibleLayers],
    geometryEdits: [...geometryEdits.values()].map((edit) => ({
      ...edit,
      points: clonePoints(edit.points),
      paths: edit.paths
        ? edit.paths.map((path) => clonePoints(path))
        : undefined,
    })),
    addedFeatures: project.features
      .filter((feature) => addedFeatureIds.has(feature.id))
      .map(cloneFeature),
    deletedFeatureIds: [...deletedFeatureIds],
    toggles: {
      grid: els.gridToggle.checked,
      cad: els.cadToggle.checked,
      labels: els.labelToggle.checked,
      elevations: els.elevationToggle.checked,
      weight: els.weightToggle.checked,
      hatch: els.hatchToggle.checked,
    },
    satellite: {
      opacity: Number(els.satelliteOpacityRange.value),
      zoom: Number(els.satelliteZoomSelect.value),
    },
    slope: {
      breaks: els.slopeBreaksInput.value,
      cellSize: Number(els.slopeCellRange.value || 2),
      opacity: Number(els.slopeOpacityRange.value || 0.58),
    },
  };
  await fetchProject({
    sourceId: currentSourceId,
    epsg: project.epsg,
    resetView: false,
    savedState,
    regen: true,
  });
  els.statusText.textContent = "리젠 완료 · 폴리라인 곡률을 다시 계산했습니다.";
}

function transformPoint(point) {
  const centerX = (bounds.xmin + bounds.xmax) / 2;
  const centerY = (bounds.ymin + bounds.ymax) / 2;
  const x = point[0] - centerX + view.panX;
  const y = point[1] - centerY + view.panY;
  const z = (point[2] || 0) * view.heightScale;
  const yaw = (view.yaw * Math.PI) / 180;
  const pitch = (view.pitch * Math.PI) / 180;
  const cosYaw = Math.cos(yaw);
  const sinYaw = Math.sin(yaw);
  const rotatedX = x * cosYaw - y * sinYaw;
  const rotatedY = x * sinYaw + y * cosYaw;
  const screenY = rotatedY * Math.cos(pitch) + z * Math.sin(pitch);
  const depth = rotatedY * Math.sin(pitch) - z * Math.cos(pitch);
  const perspective = 1 + depth / (metersWidth() * 8);
  const scale = view.zoom / Math.max(0.72, perspective);
  return {
    x: window.innerWidth / 2 + rotatedX * scale,
    y: window.innerHeight / 2 - screenY * scale,
    depth,
  };
}

function screenToProjected(screenX, screenY) {
  const centerX = (bounds.xmin + bounds.xmax) / 2;
  const centerY = (bounds.ymin + bounds.ymax) / 2;
  const screenLocalX = (screenX - window.innerWidth / 2) / view.zoom;
  const screenLocalY = -(screenY - window.innerHeight / 2) / view.zoom;
  const pitch = (view.pitch * Math.PI) / 180;
  const yaw = (view.yaw * Math.PI) / 180;
  const rotatedY = Math.abs(Math.cos(pitch)) < 0.05 ? 0 : screenLocalY / Math.cos(pitch);
  const x = screenLocalX * Math.cos(yaw) + rotatedY * Math.sin(yaw);
  const y = -screenLocalX * Math.sin(yaw) + rotatedY * Math.cos(yaw);
  return {
    x: centerX + x - view.panX,
    y: centerY + y - view.panY,
  };
}

function syncSatelliteControls() {
  const opacity = Number(els.satelliteOpacityRange.value);
  els.satelliteOpacityValue.textContent = `${Math.round(opacity * 100)}%`;
}

function updateSatelliteStatus() {
  if (!els.satelliteToggle.checked) {
    els.satelliteStatus.textContent = "위성영상 꺼짐";
    return;
  }
  if (!satellite.manifest) {
    els.satelliteStatus.textContent = "위성영상 목록 불러오는 중";
    return;
  }
  const total = satellite.manifest.tileCount;
  const completed = satellite.loaded + satellite.failed;
  if (completed < total) {
    els.satelliteStatus.textContent = `브이월드 ${completed}/${total}장 불러오는 중`;
  } else if (satellite.failed) {
    els.satelliteStatus.textContent = `브이월드 ${satellite.loaded}장 · 실패 ${satellite.failed}장`;
  } else {
    const terrain = satellite.manifest.terrain;
    const terrainText = els.terrainToggle.checked && terrain?.sampleCount
      ? ` · 지형 ${Math.round(terrain.minElevation)}~${Math.round(terrain.maxElevation)}m`
      : "";
    els.satelliteStatus.textContent = `브이월드 ${satellite.loaded}장 완료${terrainText}`;
  }
}

function loadSatelliteImage(tile, generation) {
  return new Promise((resolve) => {
    const image = new Image();
    image.decoding = "async";
    image.onload = () => {
      if (satellite.generation === generation) {
        satellite.images.set(tile.url, image);
        satellite.loaded += 1;
        updateSatelliteStatus();
        scheduleDraw();
      }
      resolve();
    };
    image.onerror = () => {
      if (satellite.generation === generation) {
        satellite.failed += 1;
        updateSatelliteStatus();
      }
      resolve();
    };
    image.src = `${tile.url}?v=1`;
  });
}

async function loadSatellite() {
  satellite.generation += 1;
  const generation = satellite.generation;
  satellite.manifest = null;
  satellite.images.clear();
  satellite.loaded = 0;
  satellite.failed = 0;
  updateSatelliteStatus();
  scheduleDraw();
  if (!els.satelliteToggle.checked || !project) return;
  const params = new URLSearchParams({
    epsg: String(project.epsg),
    source: currentSourceId,
    zoom: els.satelliteZoomSelect.value,
  });
  const response = await fetch(`/api/satellite/manifest?${params}`);
  if (!response.ok) {
    const error = await response.json().catch(() => ({ error: response.statusText }));
    throw new Error(error.error || response.statusText);
  }
  if (satellite.generation !== generation) return;
  satellite.manifest = await response.json();
  updateSatelliteStatus();
  const queue = [...satellite.manifest.tiles];
  const worker = async () => {
    while (queue.length && satellite.generation === generation) {
      const tile = queue.shift();
      await loadSatelliteImage(tile, generation);
    }
  };
  await Promise.all(Array.from({ length: Math.min(8, queue.length) }, worker));
}

function showSatelliteError(error) {
  els.satelliteStatus.textContent = `위성영상 오류: ${error.message}`;
  console.error(error);
}

function cadImageUrl(feature) {
  const params = new URLSearchParams({
    source: currentSourceId || "",
    id: String(feature.id),
  });
  return `/api/image?${params}`;
}

function loadCadImage(feature, generation) {
  if (!feature.image?.available) return;
  const url = cadImageUrl(feature);
  if (cadImages.images.has(url)) return;
  const image = new Image();
  image.decoding = "async";
  image.onload = () => {
    if (cadImages.generation === generation) {
      cadImages.images.set(url, image);
      cadImages.loaded += 1;
      scheduleDraw();
    }
  };
  image.onerror = () => {
    if (cadImages.generation === generation) {
      cadImages.failed += 1;
      scheduleDraw();
    }
  };
  image.src = url;
}

function loadCadImages() {
  cadImages.generation += 1;
  cadImages.images.clear();
  cadImages.loaded = 0;
  cadImages.failed = 0;
  const generation = cadImages.generation;
  for (const feature of project?.features || []) {
    if (feature.kind === "image") loadCadImage(feature, generation);
  }
}

function drawImageTriangle(image, source, destination) {
  const [s0, s1, s2] = source;
  const [d0, d1, d2] = destination;
  const denominator =
    s0.x * (s1.y - s2.y) +
    s1.x * (s2.y - s0.y) +
    s2.x * (s0.y - s1.y);
  if (Math.abs(denominator) < 0.001) return;
  const a =
    (d0.x * (s1.y - s2.y) +
      d1.x * (s2.y - s0.y) +
      d2.x * (s0.y - s1.y)) /
    denominator;
  const c =
    (d0.x * (s2.x - s1.x) +
      d1.x * (s0.x - s2.x) +
      d2.x * (s1.x - s0.x)) /
    denominator;
  const e =
    (d0.x * (s1.x * s2.y - s2.x * s1.y) +
      d1.x * (s2.x * s0.y - s0.x * s2.y) +
      d2.x * (s0.x * s1.y - s1.x * s0.y)) /
    denominator;
  const b =
    (d0.y * (s1.y - s2.y) +
      d1.y * (s2.y - s0.y) +
      d2.y * (s0.y - s1.y)) /
    denominator;
  const d =
    (d0.y * (s2.x - s1.x) +
      d1.y * (s0.x - s2.x) +
      d2.y * (s1.x - s0.x)) /
    denominator;
  const f =
    (d0.y * (s1.x * s2.y - s2.x * s1.y) +
      d1.y * (s2.x * s0.y - s0.x * s2.y) +
      d2.y * (s0.x * s1.y - s1.x * s0.y)) /
    denominator;
  ctx.save();
  ctx.beginPath();
  ctx.moveTo(d0.x, d0.y);
  ctx.lineTo(d1.x, d1.y);
  ctx.lineTo(d2.x, d2.y);
  ctx.closePath();
  ctx.clip();
  ctx.transform(a, b, c, d, e, f);
  ctx.drawImage(image, 0, 0);
  ctx.restore();
}

function drawCadImage(feature) {
  const points = feature.points.slice(0, 4).map(transformPoint);
  if (points.length < 4) return;
  const minX = Math.min(...points.map((point) => point.x));
  const maxX = Math.max(...points.map((point) => point.x));
  const minY = Math.min(...points.map((point) => point.y));
  const maxY = Math.max(...points.map((point) => point.y));
  if (maxX < -40 || minX > window.innerWidth + 40 || maxY < -40 || minY > window.innerHeight + 40) {
    return;
  }
  const url = cadImageUrl(feature);
  const image = cadImages.images.get(url);
  ctx.save();
  ctx.globalAlpha = Math.max(0.15, Math.min(1, Number(feature.opacity ?? 1)));
  if (image) {
    const width = image.naturalWidth || Number(feature.image?.pixelWidth) || 1;
    const height = image.naturalHeight || Number(feature.image?.pixelHeight) || 1;
    const source = [
      { x: 0, y: height },
      { x: width, y: height },
      { x: width, y: 0 },
      { x: 0, y: 0 },
    ];
    drawImageTriangle(image, [source[0], source[1], source[2]], [points[0], points[1], points[2]]);
    drawImageTriangle(image, [source[0], source[2], source[3]], [points[0], points[2], points[3]]);
  } else {
    ctx.beginPath();
    ctx.moveTo(points[0].x, points[0].y);
    for (const point of points.slice(1)) ctx.lineTo(point.x, point.y);
    ctx.closePath();
    ctx.strokeStyle = feature.image?.available ? "rgba(243, 201, 105, 0.7)" : "rgba(235, 87, 87, 0.75)";
    ctx.lineWidth = 1.5;
    ctx.setLineDash([6, 4]);
    ctx.stroke();
  }
  ctx.restore();
}

function drawSatellite() {
  if (!els.satelliteToggle.checked || !satellite.manifest) return;
  const opacity = Number(els.satelliteOpacityRange.value);
  ctx.save();
  ctx.globalAlpha = opacity;
  ctx.imageSmoothingEnabled = true;
  for (const tile of satellite.manifest.tiles) {
    const image = satellite.images.get(tile.url);
    if (!image) continue;
    const width = image.naturalWidth || 256;
    const height = image.naturalHeight || 256;
    if (els.terrainToggle.checked && tile.terrain?.points?.length) {
      const size = tile.terrain.size;
      const transformed = tile.terrain.points.map(transformPoint);
      const triangles = [];
      for (let row = 0; row < size - 1; row += 1) {
        for (let column = 0; column < size - 1; column += 1) {
          const topLeft = row * size + column;
          const topRight = topLeft + 1;
          const bottomLeft = topLeft + size;
          const bottomRight = bottomLeft + 1;
          const u0 = (column / (size - 1)) * width;
          const u1 = ((column + 1) / (size - 1)) * width;
          const v0 = (row / (size - 1)) * height;
          const v1 = ((row + 1) / (size - 1)) * height;
          triangles.push({
            source: [{ x: u0, y: v0 }, { x: u1, y: v0 }, { x: u1, y: v1 }],
            destination: [transformed[topLeft], transformed[topRight], transformed[bottomRight]],
          });
          triangles.push({
            source: [{ x: u0, y: v0 }, { x: u1, y: v1 }, { x: u0, y: v1 }],
            destination: [transformed[topLeft], transformed[bottomRight], transformed[bottomLeft]],
          });
        }
      }
      triangles.sort(
        (first, second) =>
          second.destination.reduce((sum, point) => sum + point.depth, 0) -
          first.destination.reduce((sum, point) => sum + point.depth, 0),
      );
      for (const triangle of triangles) {
        drawImageTriangle(image, triangle.source, triangle.destination);
      }
    } else {
      const corners = tile.corners.map((corner) => transformPoint([corner[0], corner[1], -1]));
      const [nw, ne, se, sw] = corners;
      drawImageTriangle(
        image,
        [{ x: 0, y: 0 }, { x: width, y: 0 }, { x: width, y: height }],
        [nw, ne, se],
      );
      drawImageTriangle(
        image,
        [{ x: 0, y: 0 }, { x: width, y: height }, { x: 0, y: height }],
        [nw, se, sw],
      );
    }
  }
  ctx.restore();
}

function drawBackground() {
  const gradient = ctx.createLinearGradient(0, 0, 0, window.innerHeight);
  gradient.addColorStop(0, "#182029");
  gradient.addColorStop(0.55, "#111820");
  gradient.addColorStop(1, "#0b1015");
  ctx.fillStyle = gradient;
  ctx.fillRect(0, 0, window.innerWidth, window.innerHeight);

  if (!bounds) return;

  const centerX = (bounds.xmin + bounds.xmax) / 2;
  const centerY = (bounds.ymin + bounds.ymax) / 2;
  const width = Math.max(bounds.xmax - bounds.xmin, bounds.ymax - bounds.ymin);
  const pad = width * 0.45;
  const plane = [
    [centerX - width / 2 - pad, centerY - width / 2 - pad, -2],
    [centerX + width / 2 + pad, centerY - width / 2 - pad, -2],
    [centerX + width / 2 + pad, centerY + width / 2 + pad, -2],
    [centerX - width / 2 - pad, centerY + width / 2 + pad, -2],
  ].map(transformPoint);
  ctx.beginPath();
  ctx.moveTo(plane[0].x, plane[0].y);
  for (const point of plane.slice(1)) ctx.lineTo(point.x, point.y);
  ctx.closePath();
  ctx.fillStyle = "#1d2929";
  ctx.fill();
  drawSatellite();

  if (!els.gridToggle.checked) return;
  const rawStep = width / 12;
  const stepPower = Math.pow(10, Math.floor(Math.log10(rawStep)));
  const step = Math.ceil(rawStep / stepPower) * stepPower;
  ctx.save();
  ctx.strokeStyle = "rgba(189, 205, 214, 0.12)";
  ctx.lineWidth = 1;
  for (let x = Math.floor((centerX - width) / step) * step; x <= centerX + width; x += step) {
    const start = transformPoint([x, centerY - width, -1]);
    const end = transformPoint([x, centerY + width, -1]);
    ctx.beginPath();
    ctx.moveTo(start.x, start.y);
    ctx.lineTo(end.x, end.y);
    ctx.stroke();
  }
  for (let y = Math.floor((centerY - width) / step) * step; y <= centerY + width; y += step) {
    const start = transformPoint([centerX - width, y, -1]);
    const end = transformPoint([centerX + width, y, -1]);
    ctx.beginPath();
    ctx.moveTo(start.x, start.y);
    ctx.lineTo(end.x, end.y);
    ctx.stroke();
  }
  ctx.restore();
}

function drawSlopeAnalysis() {
  if (!slopeAnalysis.visible || !slopeAnalysis.cells.length) return;
  const opacity = Number(els.slopeOpacityRange.value || 0.58);
  ctx.save();
  ctx.globalAlpha = Math.max(0.2, Math.min(0.85, opacity));
  ctx.lineWidth = 0;
  for (const cell of slopeAnalysis.cells) {
    const p0 = transformPoint([cell.x, cell.y, cell.z00]);
    const p1 = transformPoint([cell.x + cell.w, cell.y, cell.z10]);
    const p2 = transformPoint([cell.x + cell.w, cell.y + cell.h, cell.z11]);
    const p3 = transformPoint([cell.x, cell.y + cell.h, cell.z01]);
    const minX = Math.min(p0.x, p1.x, p2.x, p3.x);
    const maxX = Math.max(p0.x, p1.x, p2.x, p3.x);
    const minY = Math.min(p0.y, p1.y, p2.y, p3.y);
    const maxY = Math.max(p0.y, p1.y, p2.y, p3.y);
    if (maxX < -20 || minX > window.innerWidth + 20 || maxY < -20 || minY > window.innerHeight + 20) {
      continue;
    }
    ctx.beginPath();
    ctx.moveTo(p0.x, p0.y);
    ctx.lineTo(p1.x, p1.y);
    ctx.lineTo(p2.x, p2.y);
    ctx.lineTo(p3.x, p3.y);
    ctx.closePath();
    ctx.fillStyle = SLOPE_COLORS[cell.classIndex % SLOPE_COLORS.length];
    ctx.fill();
  }
  ctx.restore();
}

function hatchPattern(color) {
  if (hatchPatterns.has(color)) return hatchPatterns.get(color);
  const tile = document.createElement("canvas");
  tile.width = 10;
  tile.height = 10;
  const tileContext = tile.getContext("2d");
  tileContext.strokeStyle = color;
  tileContext.globalAlpha = 0.48;
  tileContext.lineWidth = 1;
  tileContext.beginPath();
  tileContext.moveTo(-2, 10);
  tileContext.lineTo(10, -2);
  tileContext.moveTo(4, 12);
  tileContext.lineTo(12, 4);
  tileContext.stroke();
  const pattern = ctx.createPattern(tile, "repeat");
  hatchPatterns.set(color, pattern);
  return pattern;
}

function drawHatch(feature) {
  if (!els.hatchToggle.checked) return;
  const paths = feature.paths || [feature.points];
  ctx.save();
  ctx.beginPath();
  for (const path of paths) {
    const points = path.map(transformPoint);
    if (!points.length) continue;
    ctx.moveTo(points[0].x, points[0].y);
    for (const point of points.slice(1)) ctx.lineTo(point.x, point.y);
    ctx.closePath();
  }
  const solid = (feature.pattern || "").toUpperCase() === "SOLID";
  const opacity = Number.isFinite(Number(feature.opacity))
    ? Number(feature.opacity)
    : solid ? 0.24 : 0.55;
  ctx.globalAlpha = Math.max(0.05, Math.min(0.9, opacity));
  ctx.fillStyle = solid ? feature.color : hatchPattern(feature.color);
  ctx.fill("evenodd");
  ctx.globalAlpha = feature.id === selectedFeatureId ? 1 : 0.78;
  ctx.strokeStyle = feature.color;
  ctx.lineWidth = feature.id === selectedFeatureId ? 2.2 : 0.8;
  ctx.stroke();
  ctx.restore();
}

function drawElevationLabel(feature, points) {
  if (!els.elevationToggle.checked || !feature.layer.startsWith("등고선") || view.zoom < 0.18) return;
  const sourcePoint = feature.points.find((point) => point[2]) || feature.points[0];
  const elevation = sourcePoint?.[2] || 0;
  if (!elevation) return;
  let point = null;
  let bestDistance = Infinity;
  const step = Math.max(1, Math.floor(points.length / 80));
  for (let index = 0; index < points.length; index += step) {
    const candidate = points[index];
    if (
      candidate.x < 310 ||
      candidate.x > window.innerWidth - 270 ||
      candidate.y < 78 ||
      candidate.y > window.innerHeight - 38
    ) {
      continue;
    }
    const distance =
      (candidate.x - window.innerWidth / 2) ** 2 +
      (candidate.y - window.innerHeight / 2) ** 2;
    if (distance < bestDistance) {
      point = candidate;
      bestDistance = distance;
    }
  }
  if (!point) return;
  const cell = `${Math.floor(point.x / 64)}:${Math.floor(point.y / 26)}`;
  if (elevationLabelCells.has(cell)) return;
  elevationLabelCells.add(cell);
  const text = Number.isInteger(elevation) ? `${elevation}m` : `${elevation.toFixed(1)}m`;
  ctx.save();
  ctx.font = "11px 'Segoe UI', 'Malgun Gothic', sans-serif";
  const width = ctx.measureText(text).width + 8;
  ctx.fillStyle = "rgba(9, 14, 18, 0.82)";
  ctx.fillRect(point.x - width / 2, point.y - 15, width, 16);
  ctx.fillStyle = "#f3c969";
  ctx.textAlign = "center";
  ctx.fillText(text, point.x, point.y - 3);
  ctx.restore();
}

function drawFeature(feature) {
  if (!visibleLayers.has(feature.layer)) return;
  if (feature.kind === "image") {
    drawCadImage(feature);
    return;
  }
  if (feature.kind === "hatch") {
    drawHatch(feature);
    return;
  }
  const points = feature.points.map(transformPoint);
  if (!points.length) return;
  if (feature.kind === "text") {
    if (!els.labelToggle.checked) return;
    const point = points[0];
    ctx.fillStyle = feature.color;
    ctx.font = "12px 'Segoe UI', 'Malgun Gothic', sans-serif";
    ctx.fillText(feature.text || "", point.x + 4, point.y - 4);
    return;
  }
  if (feature.kind === "point") {
    const point = points[0];
    ctx.fillStyle = feature.color;
    ctx.beginPath();
    ctx.arc(point.x, point.y, 2.4, 0, Math.PI * 2);
    ctx.fill();
    return;
  }
  ctx.beginPath();
  ctx.moveTo(points[0].x, points[0].y);
  for (const point of points.slice(1)) ctx.lineTo(point.x, point.y);
  ctx.strokeStyle = feature.color;
  ctx.globalAlpha = 0.92;
  ctx.lineWidth = els.weightToggle.checked ? Math.max(1.1, Math.min(3.2, view.zoom * 1.2)) : 1;
  ctx.lineJoin = "round";
  ctx.lineCap = "round";
  ctx.stroke();
  if (feature.id === selectedFeatureId) {
    ctx.globalAlpha = 1;
    ctx.strokeStyle = "#f3c969";
    ctx.lineWidth = 3;
    ctx.stroke();
  }
  ctx.globalAlpha = 1;
}

function selectedFeature() {
  if (!project || selectedFeatureId === null) return null;
  return project.features.find((feature) => feature.id === selectedFeatureId) || null;
}

function drawSelectionHandles() {
  if (!editMode) return;
  const feature = selectedFeature();
  if (!feature || !["polyline", "hatch"].includes(feature.kind)) return;
  const step = Math.max(1, Math.ceil(feature.points.length / 90));
  ctx.save();
  for (let index = 0; index < feature.points.length; index += step) {
    const point = transformPoint(feature.points[index]);
    ctx.beginPath();
    ctx.arc(point.x, point.y, index === selectedVertexIndex ? 5 : 3, 0, Math.PI * 2);
    ctx.fillStyle = index === selectedVertexIndex ? "#f3c969" : "#49c6b5";
    ctx.fill();
    ctx.strokeStyle = "#091015";
    ctx.lineWidth = 1;
    ctx.stroke();
  }
  if (selectedVertexIndex >= 0 && selectedVertexIndex % step !== 0) {
    const point = transformPoint(feature.points[selectedVertexIndex]);
    ctx.beginPath();
    ctx.arc(point.x, point.y, 5, 0, Math.PI * 2);
    ctx.fillStyle = "#f3c969";
    ctx.fill();
    ctx.strokeStyle = "#091015";
    ctx.stroke();
  }
  ctx.restore();
}

function drawNorthMark() {
  ctx.save();
  ctx.translate(window.innerWidth - 74, 94);
  ctx.rotate((-view.yaw * Math.PI) / 180);
  ctx.fillStyle = "#e7b957";
  ctx.beginPath();
  ctx.moveTo(0, -22);
  ctx.lineTo(8, 12);
  ctx.lineTo(0, 7);
  ctx.lineTo(-8, 12);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
  ctx.fillStyle = "rgba(238, 243, 248, 0.8)";
  ctx.font = "12px 'Segoe UI', sans-serif";
  ctx.fillText("N", window.innerWidth - 78, 130);
}

function drawSatelliteAttribution() {
  if (!els.satelliteToggle.checked || !satellite.loaded) return;
  ctx.save();
  ctx.font = "11px 'Segoe UI', 'Malgun Gothic', sans-serif";
  ctx.textAlign = "right";
  ctx.fillStyle = "rgba(8, 13, 17, 0.72)";
  ctx.fillRect(window.innerWidth - 122, window.innerHeight - 65, 108, 18);
  ctx.fillStyle = "rgba(238, 243, 248, 0.86)";
  ctx.fillText("영상 © VWorld", window.innerWidth - 20, window.innerHeight - 52);
  ctx.restore();
}

function draw() {
  ctx.clearRect(0, 0, window.innerWidth, window.innerHeight);
  drawBackground();
  if (!project || !bounds) return;
  drawSlopeAnalysis();
  elevationLabelCells = new Set();
  const drawables = project.features.slice().sort((first, second) => {
    if (first.kind === "image" && second.kind !== "image") return -1;
    if (first.kind !== "image" && second.kind === "image") return 1;
    if (first.kind === "hatch" && second.kind !== "hatch") return -1;
    if (first.kind !== "hatch" && second.kind === "hatch") return 1;
    const firstY = first.points.reduce((sum, point) => sum + point[1], 0) / first.points.length;
    const secondY = second.points.reduce((sum, point) => sum + point[1], 0) / second.points.length;
    return secondY - firstY;
  });
  if (els.cadToggle.checked) {
    for (const feature of drawables) drawFeature(feature);
    for (const feature of project.features) {
      if (visibleLayers.has(feature.layer) && feature.kind === "polyline") {
        drawElevationLabel(feature, feature.points.map(transformPoint));
      }
    }
    drawSelectionHandles();
  }
  drawNorthMark();
  drawSatelliteAttribution();
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function currentExportPayload(visibleOnly = false) {
  let features = visibleOnly
    ? project.features.filter((feature) => visibleLayers.has(feature.layer))
    : project.features;
  if (slopeAnalysis.cells.length) {
    features = [...features, ...slopeExportFeatures()];
  }
  return {
    name: project.fileName,
    sourceId: currentSourceId,
    epsg: project.epsg,
    center: project.center.wgs84,
    range: metersWidth() * 1.8,
    view: { ...view },
    slope: {
      enabled: slopeAnalysis.cells.length > 0,
      breaks: parseSlopeBreaks(),
      cellSize: Number(els.slopeCellRange.value || 2),
      opacity: Number(els.slopeOpacityRange.value || 0.58),
    },
    features: features.map((feature) => ({
      id: feature.id,
      kind: feature.kind,
      layer: feature.layer,
      color: feature.color,
      text: feature.text || "",
      closed: Boolean(feature.closed),
      pattern: feature.pattern || "",
      opacity: feature.opacity,
      image: feature.image ? { ...feature.image } : undefined,
      points: clonePoints(feature.points),
      paths: feature.paths
        ? feature.paths.map((path) => clonePoints(path))
        : undefined,
    })),
  };
}

async function postDownload(path, payload, filename) {
  els.statusText.textContent = `${filename} 만드는 중`;
  const response = await fetch(path, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    const error = await response.json().catch(() => ({ error: response.statusText }));
    throw new Error(error.error || response.statusText);
  }
  downloadBlob(await response.blob(), filename);
  els.statusText.textContent = `${filename} 저장을 시작했습니다.`;
}

function exportEditedDxf() {
  const filename = `${project.fileName.replace(/\.dxf$/i, "")}-수정본.dxf`;
  postDownload("/api/export/dxf", currentExportPayload(false), filename).catch(showError);
}

async function openBlender() {
  els.blenderBtn.disabled = true;
  els.statusText.textContent = "Blender용 지형 모델 만드는 중";
  try {
    const response = await fetch("/api/blender/open", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(currentExportPayload(true)),
    });
    const result = await response.json().catch(() => ({ error: response.statusText }));
    if (!response.ok) throw new Error(result.error || response.statusText);
    els.statusText.textContent = result.message;
  } catch (error) {
    showError(error);
  } finally {
    els.blenderBtn.disabled = false;
  }
}

async function openGoogleEarth() {
  els.googleEarthBtn.disabled = true;
  const payload = currentExportPayload(true);
  const slopeCount = payload.features.filter((feature) => String(feature.layer).startsWith("경사분석")).length;
  els.statusText.textContent = slopeCount
    ? `Google Earth용 KMZ 만드는 중 · 경사분석 ${slopeCount.toLocaleString("ko-KR")}개 면 포함`
    : "Google Earth용 KMZ 만드는 중 · 경사분석 면 없음";
  try {
    const response = await fetch("/api/google-earth/open", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const result = await response.json().catch(() => ({ error: response.statusText }));
    if (!response.ok) throw new Error(result.error || response.statusText);
    els.statusText.textContent = result.message;
  } catch (error) {
    showError(error);
  } finally {
    els.googleEarthBtn.disabled = false;
  }
}

function distanceToSegmentSquared(px, py, ax, ay, bx, by) {
  const dx = bx - ax;
  const dy = by - ay;
  if (dx === 0 && dy === 0) return (px - ax) ** 2 + (py - ay) ** 2;
  const t = Math.max(0, Math.min(1, ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)));
  const x = ax + t * dx;
  const y = ay + t * dy;
  return (px - x) ** 2 + (py - y) ** 2;
}

function pointInPolygon(screenX, screenY, points) {
  let inside = false;
  for (let index = 0, previous = points.length - 1; index < points.length; previous = index++) {
    const currentPoint = points[index];
    const previousPoint = points[previous];
    const crosses =
      (currentPoint.y > screenY) !== (previousPoint.y > screenY) &&
      screenX <
        ((previousPoint.x - currentPoint.x) * (screenY - currentPoint.y)) /
          (previousPoint.y - currentPoint.y || Number.EPSILON) +
          currentPoint.x;
    if (crosses) inside = !inside;
  }
  return inside;
}

function nearestFeatureAt(screenX, screenY) {
  let best = null;
  let bestDistance = 12 ** 2;
  for (const feature of project.features) {
    if (
      !["polyline", "hatch"].includes(feature.kind) ||
      !visibleLayers.has(feature.layer) ||
      feature.points.length < 2
    ) continue;
    const sourcePaths = feature.kind === "hatch" && feature.paths
      ? feature.paths
      : [feature.points];
    let offset = 0;
    for (const sourcePath of sourcePaths) {
      const points = sourcePath.map(transformPoint);
      if (feature.kind === "hatch" && points.length >= 3 && pointInPolygon(screenX, screenY, points)) {
        let nearestIndex = 0;
        let nearestDistance = Infinity;
        for (let index = 0; index < points.length; index += 1) {
          const distance = (screenX - points[index].x) ** 2 + (screenY - points[index].y) ** 2;
          if (distance < nearestDistance) {
            nearestDistance = distance;
            nearestIndex = index;
          }
        }
        bestDistance = 0;
        best = { feature, vertexIndex: offset + nearestIndex };
        break;
      }
      for (let index = 0; index < points.length - 1; index += 1) {
        const start = points[index];
        const end = points[index + 1];
        const distance = distanceToSegmentSquared(screenX, screenY, start.x, start.y, end.x, end.y);
        if (distance < bestDistance) {
          const startDistance = (screenX - start.x) ** 2 + (screenY - start.y) ** 2;
          const endDistance = (screenX - end.x) ** 2 + (screenY - end.y) ** 2;
          bestDistance = distance;
          best = {
            feature,
            vertexIndex: offset + (startDistance <= endDistance ? index : index + 1),
          };
        }
      }
      offset += sourcePath.length;
    }
    if (bestDistance === 0) break;
  }
  return best;
}

function nearestSelectedVertex(screenX, screenY) {
  const feature = selectedFeature();
  if (!feature) return -1;
  let bestIndex = -1;
  let bestDistance = 10 ** 2;
  for (let index = 0; index < feature.points.length; index += 1) {
    const point = transformPoint(feature.points[index]);
    const distance = (screenX - point.x) ** 2 + (screenY - point.y) ** 2;
    if (distance < bestDistance) {
      bestDistance = distance;
      bestIndex = index;
    }
  }
  return bestIndex;
}

function rememberOriginal(feature) {
  if (!originalGeometry.has(feature.id)) {
    originalGeometry.set(feature.id, featureSnapshot(feature));
  }
}

function snapshotsEqual(first, second) {
  if (!first || !second) return false;
  return JSON.stringify(first) === JSON.stringify(second);
}

function recordGeometryEdit(feature) {
  const original = originalGeometry.get(feature.id);
  const current = featureSnapshot(feature);
  if (!addedFeatureIds.has(feature.id) && original && snapshotsEqual(current, original)) {
    geometryEdits.delete(feature.id);
    originalGeometry.delete(feature.id);
    return;
  }
  geometryEdits.set(feature.id, { id: feature.id, ...current });
}

function changeCount() {
  return new Set([
    ...geometryEdits.keys(),
    ...addedFeatureIds,
    ...deletedFeatureIds,
  ]).size;
}

function cloneFeature(feature) {
  return {
    ...feature,
    points: clonePoints(feature.points),
    paths: feature.paths
      ? feature.paths.map((path) => clonePoints(path))
      : undefined,
  };
}

function updateUndoButton() {
  if (!els.undoBtn) return;
  els.undoBtn.disabled = undoStack.length === 0;
  els.undoBtn.title = undoStack.length
    ? `마지막 작업: ${undoStack[undoStack.length - 1].label}`
    : "되돌릴 작업이 없습니다.";
}

function pushGeometryUndo(feature, label) {
  undoStack.push({
    type: "geometry",
    label,
    id: feature.id,
    snapshot: featureSnapshot(feature),
  });
  updateUndoButton();
}

function refreshFeatureCount() {
  els.featureCount.textContent = project.features.length.toLocaleString("ko-KR");
}

function undoLastAction() {
  const action = undoStack.pop();
  if (!action) return;
  if (action.type === "geometry") {
    const feature = project.features.find((item) => item.id === action.id);
    if (feature) {
      restoreFeatureSnapshot(feature, action.snapshot);
      recordGeometryEdit(feature);
      selectedFeatureId = feature.id;
      selectedVertexIndex = Math.min(
        Math.max(0, selectedVertexIndex),
        feature.points.length - 1,
      );
    }
  } else if (action.type === "add") {
    project.features = project.features.filter((feature) => feature.id !== action.id);
    addedFeatureIds.delete(action.id);
    geometryEdits.delete(action.id);
    if (selectedFeatureId === action.id) {
      selectedFeatureId = null;
      selectedVertexIndex = -1;
    }
  } else if (action.type === "delete") {
    project.features.splice(action.index, 0, cloneFeature(action.feature));
    if (action.wasAdded) addedFeatureIds.add(action.feature.id);
    else deletedFeatureIds.delete(action.feature.id);
    if (action.edit) geometryEdits.set(action.feature.id, action.edit);
    selectedFeatureId = action.feature.id;
    selectedVertexIndex = 0;
  }
  refreshFeatureCount();
  updateSelectionPanel();
  updateUndoButton();
  scheduleDraw();
  els.statusText.textContent = `${action.label} 작업을 되돌렸습니다.`;
}

function applyFeatureTransform() {
  const feature = selectedFeature();
  if (!feature) {
    showError(new Error("먼저 수정할 선을 선택해 주세요."));
    return;
  }
  const moveX = Number(els.moveXInput.value || 0);
  const moveY = Number(els.moveYInput.value || 0);
  const moveZ = Number(els.moveZInput.value || 0);
  const angle = Number(els.rotateInput.value || 0);
  const scale = Number(els.scaleInput.value || 1);
  if (![moveX, moveY, moveZ, angle, scale].every(Number.isFinite) || scale <= 0) {
    showError(new Error("이동·회전·크기 값을 확인해 주세요."));
    return;
  }
  pushGeometryUndo(feature, "객체 변형");
  rememberOriginal(feature);
  const center = feature.points.reduce(
    (sum, point) => [sum[0] + point[0], sum[1] + point[1], sum[2] + point[2]],
    [0, 0, 0],
  ).map((value) => value / feature.points.length);
  const radians = (angle * Math.PI) / 180;
  const cosine = Math.cos(radians);
  const sine = Math.sin(radians);
  mapFeaturePoints(feature, (point) => {
    const localX = (point[0] - center[0]) * scale;
    const localY = (point[1] - center[1]) * scale;
    return [
      center[0] + localX * cosine - localY * sine + moveX,
      center[1] + localX * sine + localY * cosine + moveY,
      point[2] + moveZ,
    ];
  });
  recordGeometryEdit(feature);
  for (const input of [els.moveXInput, els.moveYInput, els.moveZInput, els.rotateInput]) {
    input.value = "0";
  }
  els.scaleInput.value = "1";
  updateSelectionPanel();
  scheduleDraw();
  els.statusText.textContent = "선택 객체의 이동·회전·크기를 적용했습니다.";
}

function duplicateSelectedFeature() {
  const feature = selectedFeature();
  if (!feature) {
    showError(new Error("복사할 선을 선택해 주세요."));
    return;
  }
  const copy = cloneFeature(feature);
  copy.id = nextFeatureId++;
  mapFeaturePoints(copy, (point) => [point[0] + 2, point[1] + 2, point[2]]);
  project.features.push(copy);
  addedFeatureIds.add(copy.id);
  geometryEdits.set(copy.id, { id: copy.id, ...featureSnapshot(copy) });
  undoStack.push({ type: "add", label: "객체 복사", id: copy.id });
  selectedFeatureId = copy.id;
  selectedVertexIndex = 0;
  refreshFeatureCount();
  updateSelectionPanel();
  updateUndoButton();
  scheduleDraw();
  els.statusText.textContent = "선택 객체를 2m 옆에 복사했습니다.";
}

function deleteSelectedFeature() {
  const feature = selectedFeature();
  if (!feature) {
    showError(new Error("삭제할 선을 선택해 주세요."));
    return;
  }
  const index = project.features.findIndex((item) => item.id === feature.id);
  const wasAdded = addedFeatureIds.has(feature.id);
  undoStack.push({
    type: "delete",
    label: "객체 삭제",
    feature: cloneFeature(feature),
    index,
    wasAdded,
    edit: geometryEdits.get(feature.id),
  });
  project.features.splice(index, 1);
  if (wasAdded) addedFeatureIds.delete(feature.id);
  else deletedFeatureIds.add(feature.id);
  geometryEdits.delete(feature.id);
  selectedFeatureId = null;
  selectedVertexIndex = -1;
  refreshFeatureCount();
  updateSelectionPanel();
  updateUndoButton();
  scheduleDraw();
  els.statusText.textContent = "선택 객체를 삭제했습니다.";
}

function updateSelectionPanel() {
  const feature = selectedFeature();
  const vertex = feature?.points?.[selectedVertexIndex];
  const hasFeature = Boolean(feature);
  const isHatch = feature?.kind === "hatch";
  for (const button of [
    els.applyTransformBtn,
    els.duplicateBtn,
    els.deleteBtn,
    els.revertFeatureBtn,
  ]) {
    button.disabled = !hasFeature;
  }
  els.hatchStyleControls.hidden = !isHatch;
  els.applyHatchStyleBtn.disabled = !isHatch;
  if (!feature || !vertex) {
    els.selectionText.textContent = editMode ? "수정할 선 또는 해치 면을 클릭하세요." : "간단 편집을 켜면 사용할 수 있습니다.";
    for (const input of [els.vertexXInput, els.vertexYInput, els.vertexZInput, els.allZInput]) input.value = "";
    return;
  }
  els.selectionText.textContent =
    `${feature.layer} · ${isHatch ? "해치 면" : "선"} ${feature.id + 1} · 꼭짓점 ${selectedVertexIndex + 1}/${feature.points.length}`;
  if (isHatch) {
    els.hatchColorInput.value = /^#[0-9a-f]{6}$/i.test(feature.color)
      ? feature.color
      : "#45c7d8";
    const opacity = Number.isFinite(Number(feature.opacity))
      ? Number(feature.opacity)
      : (feature.pattern || "").toUpperCase() === "SOLID" ? 0.24 : 0.55;
    els.hatchOpacityRange.value = String(opacity);
    els.hatchOpacityValue.textContent = `${Math.round(opacity * 100)}%`;
  }
  els.vertexXInput.value = vertex[0].toFixed(3);
  els.vertexYInput.value = vertex[1].toFixed(3);
  els.vertexZInput.value = vertex[2].toFixed(3);
  const averageZ = feature.points.reduce((sum, point) => sum + point[2], 0) / feature.points.length;
  els.allZInput.value = averageZ.toFixed(3);
}

function selectFeatureAt(screenX, screenY) {
  const result = nearestFeatureAt(screenX, screenY);
  selectedFeatureId = result?.feature.id ?? null;
  selectedVertexIndex = result?.vertexIndex ?? -1;
  updateSelectionPanel();
  els.statusText.textContent = result
    ? `${result.feature.layer} ${result.feature.kind === "hatch" ? "해치 면" : "선"}을 선택했습니다.`
    : "객체가 선택되지 않았습니다. 선 가까이 또는 해치 내부를 클릭하세요.";
  scheduleDraw();
}

function toggleEditMode() {
  editMode = !editMode;
  els.editBtn.classList.toggle("active", editMode);
  els.editBtn.setAttribute("aria-pressed", String(editMode));
  els.editPanel.hidden = !editMode;
  canvas.classList.toggle("editing", editMode);
  selectedFeatureId = null;
  selectedVertexIndex = -1;
  if (editMode) {
    viewBeforeEdit = { ...view };
    view.yaw = 0;
    view.pitch = 0;
    fitView();
    els.statusText.textContent = "객체 편집 모드: 선 또는 해치 면을 클릭해 선택하세요.";
  } else {
    if (viewBeforeEdit) Object.assign(view, viewBeforeEdit);
    viewBeforeEdit = null;
    syncControls();
    scheduleDraw();
    els.statusText.textContent = `변경 사항 ${changeCount().toLocaleString("ko-KR")}개가 작업 상태에 보관됩니다.`;
  }
  updateSelectionPanel();
}

function applySelectedVertex() {
  const feature = selectedFeature();
  if (!feature || selectedVertexIndex < 0) return;
  const values = [els.vertexXInput.value, els.vertexYInput.value, els.vertexZInput.value].map(Number);
  if (values.some((value) => !Number.isFinite(value))) {
    showError(new Error("X, Y, Z 좌표를 숫자로 입력해 주세요."));
    return;
  }
  pushGeometryUndo(feature, "꼭짓점 수정");
  rememberOriginal(feature);
  setFeaturePoint(feature, selectedVertexIndex, values);
  recordGeometryEdit(feature);
  updateSelectionPanel();
  scheduleDraw();
  els.statusText.textContent = "선택 꼭짓점 좌표를 적용했습니다.";
}

function applyAllZ() {
  const feature = selectedFeature();
  const z = Number(els.allZInput.value);
  if (!feature || !Number.isFinite(z)) {
    showError(new Error("수정할 객체와 전체 Z 값을 확인해 주세요."));
    return;
  }
  pushGeometryUndo(feature, "전체 높이 수정");
  rememberOriginal(feature);
  mapFeaturePoints(feature, (point) => [point[0], point[1], z]);
  recordGeometryEdit(feature);
  updateSelectionPanel();
  scheduleDraw();
  els.statusText.textContent = `${feature.layer} 객체 전체 고도를 ${formatNum(z, 3)}m로 적용했습니다.`;
}

function revertSelectedFeature() {
  const feature = selectedFeature();
  const original = feature ? originalGeometry.get(feature.id) : null;
  if (!feature || !original) return;
  pushGeometryUndo(feature, "원본 복원");
  restoreFeatureSnapshot(feature, original);
  originalGeometry.delete(feature.id);
  geometryEdits.delete(feature.id);
  selectedVertexIndex = Math.min(selectedVertexIndex, feature.points.length - 1);
  updateSelectionPanel();
  scheduleDraw();
  els.statusText.textContent = "객체 수정 내용을 원본으로 되돌렸습니다.";
}

function applyHatchStyle() {
  const feature = selectedFeature();
  if (!feature || feature.kind !== "hatch") {
    showError(new Error("색상을 변경할 해치 면을 선택해 주세요."));
    return;
  }
  const color = els.hatchColorInput.value;
  const opacity = Number(els.hatchOpacityRange.value);
  if (!/^#[0-9a-f]{6}$/i.test(color) || !Number.isFinite(opacity)) return;
  pushGeometryUndo(feature, "해치 색상 변경");
  rememberOriginal(feature);
  feature.color = color;
  feature.opacity = Math.max(0.05, Math.min(0.9, opacity));
  recordGeometryEdit(feature);
  updateSelectionPanel();
  scheduleDraw();
  els.statusText.textContent = `${feature.layer} 해치 색상과 투명도를 변경했습니다.`;
}

function saveProjectState() {
  const state = {
    format: "large-project-aerial-viewer",
    version: 1,
    savedAt: new Date().toISOString(),
    sourceId: currentSourceId,
    sourceName: project.fileName,
    epsg: project.epsg,
    view: { ...view },
    visibleLayers: [...visibleLayers],
    geometryEdits: [...geometryEdits.values()].map((edit) => ({
      ...edit,
      points: clonePoints(edit.points),
      paths: edit.paths
        ? edit.paths.map((path) => clonePoints(path))
        : undefined,
    })),
    addedFeatures: project.features
      .filter((feature) => addedFeatureIds.has(feature.id))
      .map(cloneFeature),
    deletedFeatureIds: [...deletedFeatureIds],
    toggles: {
      grid: els.gridToggle.checked,
      cad: els.cadToggle.checked,
      labels: els.labelToggle.checked,
      elevations: els.elevationToggle.checked,
      weight: els.weightToggle.checked,
      hatch: els.hatchToggle.checked,
      satellite: false,
      terrain: false,
    },
    satellite: {
      opacity: Number(els.satelliteOpacityRange.value),
      zoom: Number(els.satelliteZoomSelect.value),
    },
    slope: {
      breaks: els.slopeBreaksInput.value,
      cellSize: Number(els.slopeCellRange.value || 2),
      opacity: Number(els.slopeOpacityRange.value || 0.58),
    },
  };
  const filename = `${project.fileName.replace(/\.dxf$/i, "")}-작업.aerial.json`;
  downloadBlob(new Blob([JSON.stringify(state, null, 2)], { type: "application/json" }), filename);
  els.statusText.textContent = "현재 카메라와 레이어 상태를 저장했습니다.";
}

async function loadProjectState(file) {
  const state = JSON.parse(await file.text());
  if (state.format !== "large-project-aerial-viewer" || !state.sourceId) {
    throw new Error("이 프로그램에서 저장한 작업 파일이 아닙니다.");
  }
  await fetchProject({
    sourceId: state.sourceId,
    epsg: state.epsg || 5187,
    resetView: false,
    savedState: state,
  });
  els.statusText.textContent = `${state.sourceName || project.fileName} 작업 상태를 복원했습니다.`;
}

async function uploadDxf(file) {
  els.statusText.textContent = `${file.name} 불러오는 중`;
  const response = await fetch(`/api/upload?epsg=${encodeURIComponent(els.epsgSelect.value || "auto")}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/octet-stream",
      "X-File-Name": encodeURIComponent(file.name),
    },
    body: file,
  });
  if (!response.ok) {
    const error = await response.json().catch(() => ({ error: response.statusText }));
    throw new Error(error.error || response.statusText);
  }
  const data = await response.json();
  applyProject(data, { resetView: true });
  els.statusText.textContent =
    `${file.name} 불러오기 완료 · ${data.crsDetection?.reason || `EPSG:${data.epsg}`}`;
}

function setViewMode() {
  if (editMode) toggleEditMode();
  if (view.pitch > 5) {
    view.pitch = 0;
    view.yaw = 0;
    els.viewBtn.textContent = "2D";
  } else {
    view.pitch = 62;
    view.yaw = -28;
    els.viewBtn.textContent = "3D";
  }
  syncControls();
  scheduleDraw();
}

canvas.addEventListener("pointerdown", (event) => {
  if (event.button > 2) return;
  canvas.setPointerCapture(event.pointerId);
  if (editMode && event.button === 0) {
    const advancedOpen = Boolean(document.querySelector(".advanced-edit")?.open);
    const vertexIndex = advancedOpen
      ? nearestSelectedVertex(event.clientX, event.clientY)
      : -1;
    if (vertexIndex >= 0) {
      const feature = selectedFeature();
      pushGeometryUndo(feature, "꼭짓점 이동");
      rememberOriginal(feature);
      selectedVertexIndex = vertexIndex;
      drag = {
        type: "vertex",
        pointerId: event.pointerId,
        feature,
        vertexIndex,
        moved: false,
      };
      updateSelectionPanel();
    } else {
      const hit = nearestFeatureAt(event.clientX, event.clientY);
      if (hit && hit.feature.id === selectedFeatureId) {
        const feature = hit.feature;
        pushGeometryUndo(feature, "객체 이동");
        rememberOriginal(feature);
        drag = {
          type: "feature",
          pointerId: event.pointerId,
          feature,
          start: screenToProjected(event.clientX, event.clientY),
          snapshot: featureSnapshot(feature),
          moved: false,
        };
      } else {
        selectedFeatureId = hit?.feature.id ?? null;
        selectedVertexIndex = hit?.vertexIndex ?? -1;
        updateSelectionPanel();
        els.statusText.textContent = hit
          ? "선을 선택했습니다. 같은 선을 끌면 통째로 이동합니다."
          : "선 가까이를 클릭해 선택해 주세요.";
        scheduleDraw();
      }
    }
    return;
  }
  canvas.classList.add("dragging");
  drag = {
    type: "camera",
    pointerId: event.pointerId,
    x: event.clientX,
    y: event.clientY,
    panX: view.panX,
    panY: view.panY,
    yaw: view.yaw,
    pitch: view.pitch,
    rotate: event.button === 1 || event.button === 2 || event.shiftKey || event.ctrlKey,
  };
});

canvas.addEventListener("contextmenu", (event) => event.preventDefault());

canvas.addEventListener("pointermove", (event) => {
  if (bounds) {
    const point = screenToProjected(event.clientX, event.clientY);
    els.cursorText.textContent = `X ${formatNum(point.x, 2)}, Y ${formatNum(point.y, 2)}`;
  }
  if (!drag || drag.pointerId !== event.pointerId) return;
  if (drag.type === "vertex") {
    drag.moved = true;
    const projected = screenToProjected(event.clientX, event.clientY);
    const point = drag.feature.points[drag.vertexIndex];
    setFeaturePoint(
      drag.feature,
      drag.vertexIndex,
      [projected.x, projected.y, point[2]],
    );
    recordGeometryEdit(drag.feature);
    updateSelectionPanel();
    scheduleDraw();
    return;
  }
  if (drag.type === "feature") {
    drag.moved = true;
    const projected = screenToProjected(event.clientX, event.clientY);
    const deltaX = projected.x - drag.start.x;
    const deltaY = projected.y - drag.start.y;
    restoreFeatureSnapshot(drag.feature, drag.snapshot);
    mapFeaturePoints(drag.feature, (point) => [
      point[0] + deltaX,
      point[1] + deltaY,
      point[2],
    ]);
    recordGeometryEdit(drag.feature);
    updateSelectionPanel();
    scheduleDraw();
    return;
  }
  const deltaX = event.clientX - drag.x;
  const deltaY = event.clientY - drag.y;
  if (drag.rotate) {
    view.yaw = drag.yaw + deltaX * 0.4;
    view.pitch = Math.max(0, Math.min(78, drag.pitch - deltaY * 0.3));
  } else {
    const yaw = (view.yaw * Math.PI) / 180;
    const pitch = (view.pitch * Math.PI) / 180;
    const moveX = deltaX / view.zoom;
    const moveY = -deltaY / Math.max(0.2, Math.cos(pitch)) / view.zoom;
    view.panX = drag.panX + moveX * Math.cos(yaw) + moveY * Math.sin(yaw);
    view.panY = drag.panY - moveX * Math.sin(yaw) + moveY * Math.cos(yaw);
  }
  syncControls();
  scheduleDraw();
});

function endDrag(event) {
  if (!drag || (event && drag.pointerId !== event.pointerId)) return;
  if ((drag.type === "vertex" || drag.type === "feature") && !drag.moved) {
    undoStack.pop();
    updateUndoButton();
  } else if (drag.type === "vertex") {
    recordGeometryEdit(drag.feature);
    els.statusText.textContent = "꼭짓점 위치를 수정했습니다.";
  } else if (drag.type === "feature") {
    recordGeometryEdit(drag.feature);
    els.statusText.textContent = "선택 객체를 이동했습니다.";
  }
  canvas.classList.remove("dragging");
  drag = null;
}

canvas.addEventListener("pointerup", endDrag);
canvas.addEventListener("pointercancel", endDrag);

canvas.addEventListener(
  "wheel",
  (event) => {
    event.preventDefault();
    if (!bounds) return;
    const before = screenToProjected(event.clientX, event.clientY);
    const factor = Math.exp(-event.deltaY * 0.0012);
    view.zoom = Math.max(0.05, Math.min(12, view.zoom * factor));
    const after = screenToProjected(event.clientX, event.clientY);
    view.panX += after.x - before.x;
    view.panY += after.y - before.y;
    syncControls();
    scheduleDraw();
  },
  { passive: false },
);

canvas.addEventListener("dblclick", (event) => {
  if (!bounds) return;
  const before = screenToProjected(event.clientX, event.clientY);
  view.zoom = Math.min(12, view.zoom * 1.7);
  const after = screenToProjected(event.clientX, event.clientY);
  view.panX += after.x - before.x;
  view.panY += after.y - before.y;
  syncControls();
  scheduleDraw();
});

els.openBtn.addEventListener("click", () => els.dxfInput.click());
els.projectSaveBtn.addEventListener("click", saveProjectState);
els.projectLoadBtn.addEventListener("click", () => els.stateInput.click());
els.editBtn.addEventListener("click", toggleEditMode);
els.undoBtn.addEventListener("click", undoLastAction);
els.exportDxfBtn.addEventListener("click", exportEditedDxf);
els.blenderBtn.addEventListener("click", openBlender);
els.googleEarthBtn.addEventListener("click", openGoogleEarth);
els.applyTransformBtn.addEventListener("click", applyFeatureTransform);
els.duplicateBtn.addEventListener("click", duplicateSelectedFeature);
els.deleteBtn.addEventListener("click", deleteSelectedFeature);
els.applyVertexBtn.addEventListener("click", applySelectedVertex);
els.applyAllZBtn.addEventListener("click", applyAllZ);
els.revertFeatureBtn.addEventListener("click", revertSelectedFeature);
els.applyHatchStyleBtn.addEventListener("click", applyHatchStyle);
els.hatchOpacityRange.addEventListener("input", () => {
  els.hatchOpacityValue.textContent = `${Math.round(Number(els.hatchOpacityRange.value) * 100)}%`;
});
els.slopeAnalyzeBtn.addEventListener("click", () => analyzeSlope().catch(showError));
els.slopeToggleBtn.addEventListener("click", toggleSlopeAnalysis);
els.slopeClearBtn.addEventListener("click", clearSlopeAnalysis);
els.slopeCellRange.addEventListener("input", syncSlopeControls);
els.slopeOpacityRange.addEventListener("input", () => {
  syncSlopeControls();
  scheduleDraw();
});
els.regenBtn.addEventListener("click", () => regenerateProject().catch(showError));
els.fitBtn.addEventListener("click", fitView);
els.viewBtn.addEventListener("click", setViewMode);
els.dxfInput.addEventListener("change", () => {
  const file = els.dxfInput.files?.[0];
  if (file) uploadDxf(file).catch(showError);
  els.dxfInput.value = "";
});

els.stateInput.addEventListener("change", () => {
  const file = els.stateInput.files?.[0];
  if (file) loadProjectState(file).catch(showError);
  els.stateInput.value = "";
});

els.epsgSelect.addEventListener("change", () => {
  if (!project) return;
  fetchProject({ epsg: els.epsgSelect.value, resetView: false }).catch(showError);
});

els.satelliteToggle.addEventListener("change", () => {
  if (els.satelliteToggle.checked) loadSatellite().catch(showSatelliteError);
  else {
    satellite.generation += 1;
    updateSatelliteStatus();
    scheduleDraw();
  }
});

els.satelliteOpacityRange.addEventListener("input", () => {
  syncSatelliteControls();
  scheduleDraw();
});

els.terrainToggle.addEventListener("change", () => {
  updateSatelliteStatus();
  scheduleDraw();
});

els.satelliteZoomSelect.addEventListener("change", () => {
  loadSatellite().catch(showSatelliteError);
});

for (const [control, key] of [
  [els.zoomRange, "zoom"],
  [els.yawRange, "yaw"],
  [els.pitchRange, "pitch"],
  [els.heightRange, "heightScale"],
]) {
  control.addEventListener("input", () => {
    view[key] = Number(control.value);
    syncControls();
    scheduleDraw();
  });
}

for (const toggle of [
  els.cadToggle,
  els.gridToggle,
  els.labelToggle,
  els.elevationToggle,
  els.weightToggle,
  els.hatchToggle,
]) {
  toggle.addEventListener("change", scheduleDraw);
}

function showError(error) {
  els.statusText.textContent = error.message;
  console.error(error);
}

window.addEventListener("keydown", (event) => {
  const tagName = document.activeElement?.tagName;
  const isTyping = tagName === "INPUT" || tagName === "SELECT" || tagName === "TEXTAREA";
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "z" && !isTyping) {
    event.preventDefault();
    undoLastAction();
  } else if (event.key === "Delete" && editMode && !isTyping) {
    event.preventDefault();
    deleteSelectedFeature();
  } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "d" && editMode && !isTyping) {
    event.preventDefault();
    duplicateSelectedFeature();
  } else if (event.key === "Escape" && editMode) {
    selectedFeatureId = null;
    selectedVertexIndex = -1;
    updateSelectionPanel();
    scheduleDraw();
  }
});

window.addEventListener("resize", resize);
resize();
syncSlopeControls();
syncControls();
const startupSourceId = new URLSearchParams(window.location.search).get("source");
if (startupSourceId) {
  fetchProject({ sourceId: startupSourceId }).catch(showError);
} else {
  initializeEmptyProject();
}
