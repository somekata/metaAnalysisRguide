/* ============================================================
   MetaVis — Meta-Analysis Visualizer
   Forest Plot / Funnel Plot / Sensitivity Analysis
   Input: study, year, evente, ne, eventc, nc
   ============================================================ */

// ─── STATE ────────────────────────────────────────────────────
const state = {
  data: [],          // raw rows from CSV
  computed: [],      // rows with effect_size, se, lower_ci, upper_ci computed
  activeTab: 'forest',
  model: 'random',
  ciMult: 1.96,
  showDiamond: true,
  seThreshold: 0.30,
  measure: 'OR',     // 'OR' | 'RR' | 'RD'
};

// ─── SAMPLE RAW DATA (evente/ne/eventc/nc) ────────────────────
const SAMPLE_RAW = [
  { study:'Smith et al.',     year:2018, evente:34, ne:120, eventc:22, nc:115 },
  { study:'Johnson & Lee',    year:2019, evente:18, ne: 85, eventc:24, nc: 90 },
  { study:'Wang et al.',      year:2019, evente:72, ne:200, eventc:45, nc:195 },
  { study:'Garcia et al.',    year:2020, evente:14, ne: 60, eventc:12, nc: 65 },
  { study:'Nakamura et al.',  year:2020, evente:48, ne:145, eventc:30, nc:140 },
  { study:'Brown & Davis',    year:2021, evente:22, ne: 95, eventc:20, nc:100 },
  { study:'Patel et al.',     year:2021, evente:89, ne:230, eventc:52, nc:225 },
  { study:'Kim & Park',       year:2022, evente:36, ne:110, eventc:24, nc:105 },
  { study:'Chen et al.',      year:2022, evente:61, ne:175, eventc:38, nc:170 },
  { study:'Martinez et al.',  year:2023, evente:26, ne: 80, eventc:18, nc: 78 },
  { study:'Thompson et al.',  year:2023, evente:95, ne:260, eventc:58, nc:255 },
  { study:'Liu & Zhang',      year:2023, evente:19, ne: 70, eventc:16, nc: 72 },
];

// ─── COLORS (read from CSS variables so light/dark both work) ──
function getColors() {
  const s = getComputedStyle(document.documentElement);
  const v = k => s.getPropertyValue(k).trim();
  return {
    bg:      v('--bg'),
    bg2:     v('--bg2'),
    bg3:     v('--bg3'),
    border:  v('--border'),
    border2: v('--border2'),
    text:    v('--text'),
    text2:   v('--text2'),
    text3:   v('--text3'),
    accent:  v('--accent'),
    accent2: v('--accent2'),
    gold:    v('--gold'),
    danger:  v('--danger'),
    success: v('--success'),
    grid:    v('--canvas-grid'),
    rowAlt:  v('--canvas-row-alt'),
    zero:    v('--canvas-zero'),
    funnel:  v('--canvas-funnel'),
  };
}
// Live alias — always fresh on each render
let C = getColors();

// ─── 2×2 TABLE → effect_size / se ─────────────────────────────
// Haldane–Anscombe continuity correction (add 0.5 when any cell = 0)
function computeFromRaw(raw, measure) {
  return raw.map(d => {
    let { evente, ne, eventc, nc } = d;
    const nonevente = ne - evente;
    const noneventc = nc - eventc;

    // continuity correction if any cell is zero
    const cc = (evente === 0 || nonevente === 0 || eventc === 0 || noneventc === 0) ? 0.5 : 0;
    const a = evente    + cc;
    const b = nonevente + cc;
    const c = eventc    + cc;
    const dd = noneventc + cc;
    const n1 = ne + 2 * cc;
    const n2 = nc + 2 * cc;

    let effect_size, se, lower_ci, upper_ci;
    const z = 1.96;

    if (measure === 'OR') {
      const logOR = Math.log((a * dd) / (b * c));
      se = Math.sqrt(1/a + 1/b + 1/c + 1/dd);
      effect_size = logOR;
      lower_ci = logOR - z * se;
      upper_ci = logOR + z * se;
    } else if (measure === 'RR') {
      const p1 = a / n1, p2 = c / n2;
      const logRR = Math.log(p1 / p2);
      se = Math.sqrt((1 - p1) / (a) + (1 - p2) / (c));
      effect_size = logRR;
      lower_ci = logRR - z * se;
      upper_ci = logRR + z * se;
    } else { // RD
      const p1 = a / n1, p2 = c / n2;
      effect_size = p1 - p2;
      se = Math.sqrt(p1 * (1 - p1) / n1 + p2 * (1 - p2) / n2);
      lower_ci = effect_size - z * se;
      upper_ci = effect_size + z * se;
    }

    return {
      ...d,
      n_treatment: ne,
      n_control: nc,
      effect_size,
      se,
      lower_ci,
      upper_ci,
      // Original scale values for display
      or:  cc === 0 ? (evente * noneventc) / (nonevente * eventc) : Math.exp(Math.log((a * dd) / (b * c))),
      rr:  (evente / ne) / (eventc / nc),
      rd:  evente / ne - eventc / nc,
      p_e: evente / ne,
      p_c: eventc / nc,
    };
  });
}

// ─── LABEL HELPERS ────────────────────────────────────────────
function measureLabel(measure) {
  return { OR: 'log OR', RR: 'log RR', RD: 'Risk Difference' }[measure];
}
function measureLabelNatural(measure) {
  return { OR: 'Odds Ratio', RR: 'Risk Ratio', RD: 'Risk Difference' }[measure];
}
function formatNatural(d, measure) {
  if (measure === 'OR') return `OR = ${d.or.toFixed(3)}`;
  if (measure === 'RR') return `RR = ${d.rr.toFixed(3)}`;
  return `RD = ${d.rd.toFixed(3)}`;
}

// ─── STATS ENGINE ─────────────────────────────────────────────
function computeMeta(data, model = 'random') {
  if (!data || data.length === 0) return null;

  const yi = data.map(d => d.effect_size);
  const vi = data.map(d => d.se * d.se);
  const wi_fixed = vi.map(v => 1 / v);

  // Fixed effect pooled
  const sumW  = wi_fixed.reduce((a, b) => a + b, 0);
  const sumWY = wi_fixed.reduce((s, w, i) => s + w * yi[i], 0);
  const fe_pooled = sumWY / sumW;

  // Cochran's Q & I²
  const Q = wi_fixed.reduce((s, w, i) => s + w * (yi[i] - fe_pooled) ** 2, 0);
  const k = data.length;
  const df = k - 1;
  const I2 = Math.max(0, (Q - df) / Q * 100);

  // τ² (DerSimonian-Laird)
  const sumW2 = wi_fixed.reduce((s, w) => s + w * w, 0);
  const c  = sumW - sumW2 / sumW;
  const tau2 = Math.max(0, (Q - df) / c);

  let pooled, se_pooled, wi;
  if (model === 'random') {
    wi = vi.map(v => 1 / (v + tau2));
    const sumWr  = wi.reduce((a, b) => a + b, 0);
    const sumWrY = wi.reduce((s, w, i) => s + w * yi[i], 0);
    pooled    = sumWrY / sumWr;
    se_pooled = Math.sqrt(1 / sumWr);
  } else {
    wi = wi_fixed;
    pooled    = fe_pooled;
    se_pooled = Math.sqrt(1 / sumW);
  }

  const z = 1.96;
  const ci_lower = pooled - z * se_pooled;
  const ci_upper = pooled + z * se_pooled;

  // p-value for Q (chi-squared approx)
  const pQ = 1 - chiSquaredCDF(Q, df);

  // weights in percent
  const totalW = wi.reduce((a, b) => a + b, 0);
  const weights = wi.map(w => (w / totalW) * 100);

  return { pooled, ci_lower, ci_upper, se_pooled, Q, I2, tau2, pQ, weights, k, df, wi };
}

// Simple chi-squared CDF approximation
function chiSquaredCDF(x, k) {
  if (x <= 0) return 0;
  return regularizedGammaP(k / 2, x / 2);
}

function regularizedGammaP(a, x) {
  if (x < 0) return 0;
  if (x === 0) return 0;
  if (x > 1e8) return 1;
  let sum = 1 / a, term = 1 / a;
  for (let n = 1; n <= 200; n++) {
    term *= x / (a + n);
    sum += term;
    if (Math.abs(term) < 1e-10 * Math.abs(sum)) break;
  }
  return Math.min(1, sum * Math.exp(-x + a * Math.log(x) - lgamma(a)));
}

function lgamma(x) {
  const c = [76.18009172947146,-86.50532032941677,24.01409824083091,
             -1.231739572450155,0.1208650973866179e-2,-0.5395239384953e-5];
  let y = x, tmp = x + 5.5;
  tmp -= (x + 0.5) * Math.log(tmp);
  let ser = 1.000000000190015;
  for (let j = 0; j < 6; j++) { y++; ser += c[j] / y; }
  return -tmp + Math.log(2.5066282746310005 * ser / x);
}

// ─── CANVAS UTILS ─────────────────────────────────────────────
function getCanvas() {
  const canvas = document.getElementById('plotCanvas');
  const wrapper = canvas.parentElement;
  const dpr = window.devicePixelRatio || 1;
  const w = wrapper.clientWidth;
  const h = wrapper.clientHeight;
  canvas.width  = w * dpr;
  canvas.height = h * dpr;
  canvas.style.width  = w + 'px';
  canvas.style.height = h + 'px';
  const ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);
  return { ctx, w, h };
}

function drawGrid(ctx, xScale, yTicks, margin, w, h, xLabel = 'Effect Size') {
  // Vertical grid lines
  ctx.strokeStyle = C.grid;
  ctx.lineWidth = 0.5;
  ctx.setLineDash([4, 4]);
  xTicks(xScale).forEach(v => {
    const x = xScale(v);
    ctx.beginPath(); ctx.moveTo(x, margin.top); ctx.lineTo(x, h - margin.bottom);
    ctx.stroke();
  });
  ctx.setLineDash([]);

  // Zero line
  const x0 = xScale(0);
  if (x0 >= margin.left && x0 <= w - margin.right) {
    ctx.strokeStyle = 'rgba(232,197,106,0.25)';
    ctx.lineWidth = 1.2;
    ctx.setLineDash([6, 3]);
    ctx.beginPath(); ctx.moveTo(x0, margin.top); ctx.lineTo(x0, h - margin.bottom);
    ctx.stroke();
    ctx.setLineDash([]);
  }
}

function xTicks(xScale) {
  // derive from domain
  const [dMin, dMax] = xScale.domain;
  const range = dMax - dMin;
  let step = 0.1;
  if (range > 4) step = 1;
  else if (range > 2) step = 0.5;
  else if (range > 1) step = 0.2;
  const ticks = [];
  for (let v = Math.ceil(dMin / step) * step; v <= dMax + 1e-9; v += step) {
    ticks.push(parseFloat(v.toFixed(4)));
  }
  return ticks;
}

function makeXScale(min, max, left, right) {
  const fn = v => left + (v - min) / (max - min) * (right - left);
  fn.domain = [min, max];
  fn.invert = x => min + (x - left) / (right - left) * (max - min);
  return fn;
}

function drawAxes(ctx, xScale, margin, w, h) {
  ctx.strokeStyle = C.border2;
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(margin.left, h - margin.bottom);
  ctx.lineTo(w - margin.right, h - margin.bottom);
  ctx.stroke();

  ctx.fillStyle = C.text2;
  ctx.font = '11px "DM Mono", monospace';
  ctx.textAlign = 'center';
  xTicks(xScale).forEach(v => {
    const x = xScale(v);
    ctx.fillText(v.toFixed(1), x, h - margin.bottom + 16);
    // tick
    ctx.beginPath();
    ctx.strokeStyle = C.border2;
    ctx.lineWidth = 0.8;
    ctx.moveTo(x, h - margin.bottom);
    ctx.lineTo(x, h - margin.bottom + 5);
    ctx.stroke();
  });
}

// ─── FOREST PLOT ──────────────────────────────────────────────
function drawForest(data, meta) {
  const { ctx, w, h } = getCanvas();
  ctx.clearRect(0, 0, w, h);

  const n = data.length;
  const margin = { top: 30, bottom: 50, left: 160, right: 160 };
  const plotH = h - margin.top - margin.bottom;
  const rowH  = Math.min(plotH / (n + 2), 32);

  // x scale
  const xmin = Math.min(...data.map(d => d.lower_ci)) - 0.3;
  const xmax = Math.max(...data.map(d => d.upper_ci)) + 0.3;
  const xScale = makeXScale(
    Math.min(xmin, -0.2), Math.max(xmax, 1.2),
    margin.left, w - margin.right
  );

  // background
  ctx.fillStyle = C.bg;
  ctx.fillRect(0, 0, w, h);

  // grid
  ctx.strokeStyle = C.grid;
  ctx.lineWidth = 0.5;
  ctx.setLineDash([4, 4]);
  xTicks(xScale).forEach(v => {
    const x = xScale(v);
    ctx.beginPath(); ctx.moveTo(x, margin.top - 10); ctx.lineTo(x, h - margin.bottom);
    ctx.stroke();
  });
  ctx.setLineDash([]);

  // zero line
  const x0 = xScale(0);
  ctx.strokeStyle = C.zero;
  ctx.lineWidth = 1;
  ctx.setLineDash([5, 3]);
  ctx.beginPath(); ctx.moveTo(x0, margin.top - 10); ctx.lineTo(x0, h - margin.bottom - 5);
  ctx.stroke();
  ctx.setLineDash([]);

  // header row
  ctx.fillStyle = C.text3;
  ctx.font = '10px "DM Mono", monospace';
  ctx.textAlign = 'left';
  ctx.fillText('Study', 8, margin.top - 6);
  ctx.textAlign = 'right';
  ctx.fillText('ES [95% CI]', w - margin.right + 80, margin.top - 6);

  // rows
  data.forEach((d, i) => {
    const y = margin.top + (i + 0.5) * rowH;
    const xL = xScale(d.lower_ci);
    const xR = xScale(d.upper_ci);
    const xE = xScale(d.effect_size);
    const wPct = meta ? meta.weights[i] : (d.weight || 5);
    const sq = Math.max(3, Math.min(10, wPct * 0.6));

    // alternating row bg
    if (i % 2 === 0) {
      ctx.fillStyle = C.rowAlt;
      ctx.fillRect(0, margin.top + i * rowH, w, rowH);
    }

    // CI line
    ctx.strokeStyle = C.accent;
    ctx.lineWidth = 1.5;
    ctx.beginPath(); ctx.moveTo(xL, y); ctx.lineTo(xR, y); ctx.stroke();

    // CI caps
    ctx.beginPath();
    ctx.moveTo(xL, y - 5); ctx.lineTo(xL, y + 5);
    ctx.moveTo(xR, y - 5); ctx.lineTo(xR, y + 5);
    ctx.stroke();

    // effect square
    const inCI = d.lower_ci > 0 || d.upper_ci < 0;
    ctx.fillStyle = inCI ? C.accent : C.accent2;
    ctx.fillRect(xE - sq/2, y - sq/2, sq, sq);

    // study label
    ctx.fillStyle = C.text2;
    ctx.font = '11px "DM Mono", monospace';
    ctx.textAlign = 'right';
    ctx.fillText(d.study + ' ' + d.year, margin.left - 8, y + 4);

    // ES value — natural scale for OR/RR, raw for RD
    ctx.fillStyle = C.text;
    ctx.textAlign = 'left';
    ctx.font = '11px "DM Mono", monospace';
    const m = state.measure;
    let ciText;
    if (m === 'OR' || m === 'RR') {
      const nat  = (m === 'OR' ? d.or  : d.rr).toFixed(2);
      const natL = Math.exp(d.lower_ci).toFixed(2);
      const natR = Math.exp(d.upper_ci).toFixed(2);
      ciText = `${nat} [${natL}, ${natR}]`;
    } else {
      ciText = `${d.effect_size.toFixed(2)} [${d.lower_ci.toFixed(2)}, ${d.upper_ci.toFixed(2)}]`;
    }
    ctx.fillText(ciText, w - margin.right + 8, y + 4);
  });

  // pooled diamond
  if (state.showDiamond && meta) {
    const yD = margin.top + (n + 0.8) * rowH;
    const xPool = xScale(meta.pooled);
    const xPoolL = xScale(meta.ci_lower);
    const xPoolR = xScale(meta.ci_upper);
    const dh = rowH * 0.35;

    // separator
    ctx.strokeStyle = C.border;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(margin.left, margin.top + n * rowH + 4);
    ctx.lineTo(w - margin.right, margin.top + n * rowH + 4);
    ctx.stroke();

    // diamond
    ctx.beginPath();
    ctx.moveTo(xPool, yD - dh);
    ctx.lineTo(xPoolR, yD);
    ctx.lineTo(xPool, yD + dh);
    ctx.lineTo(xPoolL, yD);
    ctx.closePath();
    ctx.fillStyle = C.gold;
    ctx.fill();
    ctx.strokeStyle = 'rgba(232,197,106,0.5)';
    ctx.lineWidth = 1;
    ctx.stroke();

    // label
    ctx.fillStyle = C.gold;
    ctx.font = '500 11px "DM Mono", monospace';
    ctx.textAlign = 'right';
    ctx.fillText(`Pooled (${state.model === 'random' ? 'RE' : 'FE'})`, margin.left - 8, yD + 4);

    ctx.fillStyle = C.text;
    ctx.textAlign = 'left';
    const pm = state.measure;
    let pText;
    if (pm === 'OR' || pm === 'RR') {
      const expP = Math.exp(meta.pooled).toFixed(2);
      const expL = Math.exp(meta.ci_lower).toFixed(2);
      const expU = Math.exp(meta.ci_upper).toFixed(2);
      pText = `${expP} [${expL}, ${expU}]`;
    } else {
      pText = `${meta.pooled.toFixed(2)} [${meta.ci_lower.toFixed(2)}, ${meta.ci_upper.toFixed(2)}]`;
    }
    ctx.fillText(pText, w - margin.right + 8, yD + 4);
  }

  // x-axis
  drawAxes(ctx, xScale, margin, w, h);
  ctx.fillStyle = C.text3;
  ctx.font = '11px "DM Mono", monospace';
  ctx.textAlign = 'center';
  ctx.fillText(measureLabel(state.measure), w / 2, h - 8);
}

// ─── FUNNEL PLOT ──────────────────────────────────────────────
function drawFunnel(data, meta) {
  const { ctx, w, h } = getCanvas();
  ctx.clearRect(0, 0, w, h);

  ctx.fillStyle = C.bg;
  ctx.fillRect(0, 0, w, h);

  const margin = { top: 40, bottom: 55, left: 70, right: 40 };
  const plotW = w - margin.left - margin.right;
  const plotH = h - margin.top - margin.bottom;

  const pooled = meta ? meta.pooled : data.reduce((s, d) => s + d.effect_size, 0) / data.length;
  const maxSE = Math.max(...data.map(d => d.se)) * 1.15;

  const yScale = v => margin.top + (v / maxSE) * plotH;
  const allX = data.map(d => d.effect_size);
  const xpad = Math.max(Math.abs(pooled - Math.min(...allX)), Math.abs(Math.max(...allX) - pooled)) * 1.4 + 0.2;
  const xmin = pooled - xpad;
  const xmax = pooled + xpad;
  const xScale = makeXScale(xmin, xmax, margin.left, w - margin.right);

  // grid
  [0.1, 0.2, 0.3, 0.4, 0.5].filter(v => v <= maxSE).forEach(se => {
    const y = yScale(se);
    ctx.strokeStyle = C.grid;
    ctx.lineWidth = 0.5;
    ctx.setLineDash([4, 4]);
    ctx.beginPath(); ctx.moveTo(margin.left, y); ctx.lineTo(w - margin.right, y); ctx.stroke();
    ctx.setLineDash([]);
    ctx.fillStyle = C.text3;
    ctx.font = '10px "DM Mono"';
    ctx.textAlign = 'right';
    ctx.fillText(se.toFixed(2), margin.left - 6, y + 4);
  });

  // funnel pseudo-CI lines (95%)
  const z = 1.96;
  ctx.strokeStyle = C.accent + '30';
  ctx.lineWidth = 1;
  ctx.setLineDash([8, 4]);
  ctx.beginPath();
  ctx.moveTo(xScale(pooled), margin.top);
  ctx.lineTo(xScale(pooled - z * maxSE), h - margin.bottom);
  ctx.stroke();
  ctx.beginPath();
  ctx.moveTo(xScale(pooled), margin.top);
  ctx.lineTo(xScale(pooled + z * maxSE), h - margin.bottom);
  ctx.stroke();
  ctx.setLineDash([]);

  // funnel fill
  ctx.beginPath();
  ctx.moveTo(xScale(pooled), margin.top);
  ctx.lineTo(xScale(pooled + z * maxSE), h - margin.bottom);
  ctx.lineTo(xScale(pooled - z * maxSE), h - margin.bottom);
  ctx.closePath();
  ctx.fillStyle = C.funnel;
  ctx.fill();

  // pooled vertical
  ctx.strokeStyle = C.zero;
  ctx.lineWidth = 1.2;
  ctx.setLineDash([5, 3]);
  const xp = xScale(pooled);
  ctx.beginPath(); ctx.moveTo(xp, margin.top); ctx.lineTo(xp, h - margin.bottom); ctx.stroke();
  ctx.setLineDash([]);

  // dots
  data.forEach(d => {
    const x = xScale(d.effect_size);
    const y = yScale(d.se);
    const inside = Math.abs(d.effect_size - pooled) <= z * d.se;
    ctx.beginPath();
    ctx.arc(x, y, 5.5, 0, Math.PI * 2);
    ctx.fillStyle = inside ? C.accent : C.danger;
    ctx.fill();
    ctx.strokeStyle = inside ? 'rgba(79,142,247,0.5)' : 'rgba(240,112,112,0.5)';
    ctx.lineWidth = 1;
    ctx.stroke();
  });

  // axes
  ctx.strokeStyle = C.border2;
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(margin.left, margin.top); ctx.lineTo(margin.left, h - margin.bottom);
  ctx.moveTo(margin.left, h - margin.bottom); ctx.lineTo(w - margin.right, h - margin.bottom);
  ctx.stroke();

  // x ticks
  ctx.fillStyle = C.text2;
  ctx.font = '11px "DM Mono"';
  ctx.textAlign = 'center';
  xTicks(xScale).forEach(v => {
    const x = xScale(v);
    ctx.fillText(v.toFixed(1), x, h - margin.bottom + 16);
    ctx.strokeStyle = C.border2;
    ctx.lineWidth = 0.8;
    ctx.beginPath();
    ctx.moveTo(x, h - margin.bottom); ctx.lineTo(x, h - margin.bottom + 4);
    ctx.stroke();
  });

  ctx.fillStyle = C.text3;
  ctx.font = '11px "DM Mono"';
  ctx.textAlign = 'center';
  ctx.fillText(measureLabel(state.measure), w / 2, h - 8);

  // y label
  ctx.save();
  ctx.translate(14, h / 2);
  ctx.rotate(-Math.PI / 2);
  ctx.fillText('Standard Error', 0, 0);
  ctx.restore();

  // y axis label at top (inverted)
  ctx.fillStyle = C.text3;
  ctx.font = '10px "DM Mono"';
  ctx.textAlign = 'center';
  ctx.fillText('← More precise', margin.left + 60, margin.top - 16);
}

// ─── SENSITIVITY ANALYSIS ─────────────────────────────────────
function drawSensitivity(data) {
  const { ctx, w, h } = getCanvas();
  ctx.clearRect(0, 0, w, h);
  ctx.fillStyle = C.bg;
  ctx.fillRect(0, 0, w, h);

  const n = data.length;
  const margin = { top: 40, bottom: 55, left: 170, right: 160 };
  const rowH = Math.min((h - margin.top - margin.bottom) / (n + 1), 32);

  // leave-one-out pooled values
  const results = data.map((_, omit) => {
    const subset = data.filter((__, j) => j !== omit);
    const m = computeMeta(subset, state.model);
    return m;
  });

  const allVals = results.flatMap(r => r ? [r.ci_lower, r.ci_upper] : []);
  const xmin = Math.min(...allVals) - 0.15;
  const xmax = Math.max(...allVals) + 0.15;
  const xScale = makeXScale(xmin, xmax, margin.left, w - margin.right);

  // full meta
  const fullMeta = computeMeta(data, state.model);
  const fullPooled = fullMeta ? fullMeta.pooled : 0;

  // grid
  ctx.strokeStyle = C.grid;
  ctx.lineWidth = 0.5;
  ctx.setLineDash([4, 4]);
  xTicks(xScale).forEach(v => {
    const x = xScale(v);
    ctx.beginPath(); ctx.moveTo(x, margin.top - 10); ctx.lineTo(x, h - margin.bottom);
    ctx.stroke();
  });
  ctx.setLineDash([]);

  // full pooled reference line
  const xFull = xScale(fullPooled);
  ctx.strokeStyle = C.zero;
  ctx.lineWidth = 1.5;
  ctx.setLineDash([5, 3]);
  ctx.beginPath(); ctx.moveTo(xFull, margin.top - 10); ctx.lineTo(xFull, h - margin.bottom); ctx.stroke();
  ctx.setLineDash([]);

  // header
  ctx.fillStyle = C.text3;
  ctx.font = '10px "DM Mono"';
  ctx.textAlign = 'left';
  ctx.fillText('Omitted Study', 8, margin.top - 10);
  ctx.textAlign = 'right';
  ctx.fillText('Pooled ES [95% CI]', w - margin.right + 42, margin.top - 10);

  results.forEach((r, i) => {
    if (!r) return;
    const d = data[i];
    const y = margin.top + (i + 0.5) * rowH;
    const xL = xScale(r.ci_lower);
    const xR = xScale(r.ci_upper);
    const xP = xScale(r.pooled);

    // row bg
    if (i % 2 === 0) {
      ctx.fillStyle = C.rowAlt;
      ctx.fillRect(0, margin.top + i * rowH, w, rowH);
    }

    // deviation from full — compare on the internal (log) scale
    const dev = Math.abs(r.pooled - fullPooled);
    const isOutlier = dev > state.seThreshold;
    const barColor = isOutlier ? C.danger : C.success;

    // CI line
    ctx.strokeStyle = isOutlier ? 'rgba(240,112,112,0.7)' : 'rgba(93,216,154,0.7)';
    ctx.lineWidth = 1.5;
    ctx.beginPath(); ctx.moveTo(xL, y); ctx.lineTo(xR, y); ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(xL, y-5); ctx.lineTo(xL, y+5);
    ctx.moveTo(xR, y-5); ctx.lineTo(xR, y+5);
    ctx.stroke();

    // diamond
    const dh = rowH * 0.3;
    ctx.beginPath();
    ctx.moveTo(xP, y - dh); ctx.lineTo(xR > xL ? Math.min(xP + 6, xR) : xP + 6, y);
    ctx.lineTo(xP, y + dh); ctx.lineTo(xR > xL ? Math.max(xP - 6, xL) : xP - 6, y);
    ctx.closePath();
    ctx.fillStyle = barColor;
    ctx.fill();

    // study name
    ctx.fillStyle = isOutlier ? C.danger : C.text2;
    ctx.font = '11px "DM Mono"';
    ctx.textAlign = 'right';
    ctx.fillText(d.study + ' ' + d.year, margin.left - 8, y + 4);

    // value label — natural scale for OR/RR
    ctx.fillStyle = isOutlier ? C.danger : C.text;
    ctx.font = '11px "DM Mono"';
    ctx.textAlign = 'left';
    const sm = state.measure;
    let sLabel;
    if (sm === 'OR' || sm === 'RR') {
      sLabel = `${Math.exp(r.pooled).toFixed(2)} [${Math.exp(r.ci_lower).toFixed(2)}, ${Math.exp(r.ci_upper).toFixed(2)}]`;
    } else {
      sLabel = `${r.pooled.toFixed(2)} [${r.ci_lower.toFixed(2)}, ${r.ci_upper.toFixed(2)}]`;
    }
    ctx.fillText(sLabel, w - margin.right + 8, y + 4);
  });

  // x-axis
  drawAxes(ctx, xScale, margin, w, h);
  ctx.fillStyle = C.text3;
  ctx.font = '11px "DM Mono"';
  ctx.textAlign = 'center';
  ctx.fillText(`Pooled ${measureLabelNatural(state.measure)} (leave-one-out)`, w / 2, h - 8);

  // legend
  const lx = margin.left + 10, ly = h - margin.bottom + 28;
  ctx.fillStyle = C.success; ctx.fillRect(lx, ly - 8, 12, 10);
  ctx.fillStyle = C.text3; ctx.font = '10px "DM Mono"'; ctx.textAlign = 'left';
  ctx.fillText(`Stable (Δ ≤ ${state.seThreshold.toFixed(2)})`, lx + 16, ly);
  ctx.fillStyle = C.danger; ctx.fillRect(lx + 140, ly - 8, 12, 10);
  ctx.fillText(`Influential (Δ > ${state.seThreshold.toFixed(2)})`, lx + 156, ly);
}

// ─── RENDER DISPATCHER ────────────────────────────────────────
function render() {
  C = getColors(); // refresh on every render (theme may have changed)
  const raw = state.data.length > 0 ? state.data : SAMPLE_RAW;
  const data = computeFromRaw(raw, state.measure);
  state.computed = data;
  const meta = computeMeta(data, state.model);

  if (meta) updateStats(meta, data, data.length);

  if (state.activeTab === 'forest')           drawForest(data, meta);
  else if (state.activeTab === 'funnel')      drawFunnel(data, meta);
  else if (state.activeTab === 'sensitivity') drawSensitivity(data);
}

// ─── STATS UPDATE ─────────────────────────────────────────────
function updateStats(meta, data, n) {
  const m = state.measure;
  document.getElementById('statMeasure').textContent  = measureLabelNatural(m);

  // Pooled on natural scale for OR/RR, raw for RD
  let pooledDisplay = meta.pooled.toFixed(3);
  let ciDisplay = `[${meta.ci_lower.toFixed(3)}, ${meta.ci_upper.toFixed(3)}]`;
  if (m === 'OR' || m === 'RR') {
    const expP  = Math.exp(meta.pooled).toFixed(3);
    const expL  = Math.exp(meta.ci_lower).toFixed(3);
    const expU  = Math.exp(meta.ci_upper).toFixed(3);
    pooledDisplay = `${expP} (log: ${meta.pooled.toFixed(3)})`;
    ciDisplay = `[${expL}, ${expU}]`;
  }
  document.getElementById('statPooled').textContent = pooledDisplay;
  document.getElementById('statCI').textContent     = ciDisplay;
  document.getElementById('statI2').textContent     = meta.I2.toFixed(1) + '%';
  document.getElementById('statTau2').textContent   = meta.tau2.toFixed(4);
  document.getElementById('statQ').textContent      = meta.Q.toFixed(2);
  document.getElementById('statPQ').textContent     = meta.pQ < 0.001 ? '<0.001' : meta.pQ.toFixed(3);
  document.getElementById('statN').textContent      = n;
}

// ─── CSV PARSER ───────────────────────────────────────────────
function parseCSV(text) {
  const lines = text.trim().split(/\r?\n/);
  const headers = lines[0].split(',').map(h => h.trim().toLowerCase());
  const required = ['study', 'evente', 'ne', 'eventc', 'nc'];
  const missing = required.filter(f => !headers.includes(f));
  if (missing.length > 0) {
    alert(`CSVに必要な列がありません: ${missing.join(', ')}\n必要: study, year, evente, ne, eventc, nc`);
    return [];
  }
  return lines.slice(1)
    .filter(l => l.trim())
    .map(line => {
      const vals = line.split(',');
      const obj = {};
      headers.forEach((h, i) => {
        const v = vals[i]?.trim() ?? '';
        obj[h] = (h === 'study') ? v : parseFloat(v);
      });
      return obj;
    })
    .filter(d =>
      d.study &&
      !isNaN(d.evente) && !isNaN(d.ne) &&
      !isNaN(d.eventc) && !isNaN(d.nc) &&
      d.ne > 0 && d.nc > 0
    );
}

// ─── TOOLTIP ─────────────────────────────────────────────────
function setupTooltip() {
  const canvas = document.getElementById('plotCanvas');
  const tooltip = document.getElementById('tooltip');

  canvas.addEventListener('mousemove', e => {
    const raw = state.data.length > 0 ? state.data : SAMPLE_RAW;
    const data = state.computed.length > 0 ? state.computed : computeFromRaw(raw, state.measure);
    const meta = computeMeta(data, state.model);
    const rect = canvas.getBoundingClientRect();
    const my = e.clientY - rect.top;

    if (state.activeTab !== 'forest' && state.activeTab !== 'funnel') {
      tooltip.classList.remove('visible');
      return;
    }

    const n = data.length;
    const margin = { top: 30, bottom: 50, left: 160, right: 160 };
    const rowH = Math.min((canvas.clientHeight - margin.top - margin.bottom) / (n + 2), 32);

    let found = null, foundIdx = -1;
    data.forEach((d, i) => {
      const y = margin.top + (i + 0.5) * rowH;
      if (Math.abs(my - y) < rowH / 2 + 2) { found = d; foundIdx = i; }
    });

    if (found) {
      const m = state.measure;
      const naturalVal = m === 'OR' ? `OR = ${found.or.toFixed(3)}`
                       : m === 'RR' ? `RR = ${found.rr.toFixed(3)}`
                       : `RD = ${found.rd.toFixed(3)}`;
      const ciNatural = (m === 'OR' || m === 'RR')
        ? `[${Math.exp(found.lower_ci).toFixed(3)}, ${Math.exp(found.upper_ci).toFixed(3)}]`
        : `[${found.lower_ci.toFixed(3)}, ${found.upper_ci.toFixed(3)}]`;

      tooltip.innerHTML = `
        <div class="tooltip-title">${found.study}${found.year ? ' (' + found.year + ')' : ''}</div>
        <div class="tooltip-row"><span>Events (T/C)</span><span class="tooltip-val">${found.evente} / ${found.eventc}</span></div>
        <div class="tooltip-row"><span>N (T/C)</span><span class="tooltip-val">${found.n_treatment} / ${found.n_control}</span></div>
        <div class="tooltip-row"><span>Rate (T/C)</span><span class="tooltip-val">${(found.p_e*100).toFixed(1)}% / ${(found.p_c*100).toFixed(1)}%</span></div>
        <div class="tooltip-row"><span>${measureLabelNatural(m)}</span><span class="tooltip-val">${naturalVal.split('= ')[1]}</span></div>
        <div class="tooltip-row"><span>95% CI</span><span class="tooltip-val">${ciNatural}</span></div>
        <div class="tooltip-row"><span>SE</span><span class="tooltip-val">${found.se.toFixed(4)}</span></div>
        <div class="tooltip-row"><span>Weight</span><span class="tooltip-val">${(meta?.weights[foundIdx] || 0).toFixed(1)}%</span></div>
      `;
      let tx = e.clientX - rect.left + 14;
      let ty = e.clientY - rect.top - 60;
      if (tx + 240 > canvas.clientWidth) tx -= 254;
      if (ty < 0) ty = 4;
      tooltip.style.left = tx + 'px';
      tooltip.style.top  = ty + 'px';
      tooltip.classList.add('visible');
    } else {
      tooltip.classList.remove('visible');
    }
  });

  canvas.addEventListener('mouseleave', () => tooltip.classList.remove('visible'));
}

// ─── EXPORT ───────────────────────────────────────────────────
document.getElementById('exportBtn').addEventListener('click', () => {
  const canvas = document.getElementById('plotCanvas');
  const link = document.createElement('a');
  link.download = `metavis_${state.activeTab}_${Date.now()}.png`;
  link.href = canvas.toDataURL('image/png');
  link.click();
});

// ─── EVENTS ───────────────────────────────────────────────────
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    state.activeTab = btn.dataset.tab;

    const titles = { forest: 'Forest Plot', funnel: 'Funnel Plot', sensitivity: 'Sensitivity Analysis' };
    const notes = {
      forest: '効果量と95%信頼区間。■のサイズは研究のウェイトに比例。◆はプール効果量を示す。',
      funnel: '効果量 vs 標準誤差。点線は95%擬似信頼区間（ファネル）。非対称は出版バイアスを示唆。',
      sensitivity: 'Leave-one-out解析。各研究を除外したときのプール効果量の変化。赤は影響の大きい研究。'
    };
    document.getElementById('plotTitle').textContent = titles[state.activeTab];
    document.getElementById('footnote').textContent = notes[state.activeTab];
    document.getElementById('sensitivitySettings').style.display =
      state.activeTab === 'sensitivity' ? 'block' : 'none';

    render();
  });
});

document.querySelectorAll('input[name="model"]').forEach(r => {
  r.addEventListener('change', e => { state.model = e.target.value; render(); });
});

document.getElementById('ciLevel').addEventListener('change', e => {
  state.ciMult = parseFloat(e.target.value); render();
});

document.getElementById('showDiamond').addEventListener('change', e => {
  state.showDiamond = e.target.checked; render();
});

document.getElementById('seThreshold').addEventListener('input', e => {
  state.seThreshold = parseFloat(e.target.value);
  document.getElementById('seThresholdVal').textContent = state.seThreshold.toFixed(2);
  render();
});

document.getElementById('effectMeasure').addEventListener('change', e => {
  state.measure = e.target.value; render();
});

document.getElementById('csvInput').addEventListener('change', e => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    const parsed = parseCSV(ev.target.result);
    if (parsed.length > 0) {
      state.data = parsed;
      document.getElementById('dataBadge').textContent = file.name;
      render();
    }
  };
  reader.readAsText(file);
});

window.addEventListener('resize', () => {
  clearTimeout(window._resizeTimer);
  window._resizeTimer = setTimeout(render, 80);
});

// ─── THEME TOGGLE ─────────────────────────────────────────────
(function () {
  const html   = document.documentElement;
  const btn    = document.getElementById('themeToggle');
  const moon   = document.getElementById('iconMoon');
  const sun    = document.getElementById('iconSun');

  // Restore saved preference
  const saved = localStorage.getItem('metavis-theme');
  if (saved === 'light') {
    html.setAttribute('data-theme', 'light');
    moon.style.display = 'none';
    sun.style.display  = '';
  }

  btn.addEventListener('click', () => {
    const isLight = html.getAttribute('data-theme') === 'light';
    const next = isLight ? 'dark' : 'light';
    html.setAttribute('data-theme', next);
    localStorage.setItem('metavis-theme', next);
    moon.style.display = next === 'dark' ? '' : 'none';
    sun.style.display  = next === 'light' ? '' : 'none';
    render(); // redraw canvas with new colors
  });
})();

// ─── INIT ─────────────────────────────────────────────────────
setupTooltip();
render();
