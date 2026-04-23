'use strict';

/* ════════════════════════════════════════
   Constants & State
════════════════════════════════════════ */
const STORAGE_KEY = 'studycsv_v3';
const COL_KEY     = 'studycsv_cols_v3';
const WIDTH_KEY   = 'studycsv_widths';
const THEME_KEY   = 'studycsv_theme';
const HIST_LIMIT  = 50;

// Fixed columns (always present, order preserved in CSV)
const FIXED_FIELDS = ['year','study','region','url','notes','evente','ne','eventc','nc'];
const DEFAULT_COLS = { evente:'evente', ne:'ne', eventc:'eventc', nc:'nc' };

// Default column pixel widths  (key = field id)
const DEFAULT_WIDTHS = {
  _rownum: 36, _inc: 40,
  year: 68, study: 180, region: 100, url: 160, notes: 200,
  evente: 80, ne: 80, eventc: 80, nc: 80,
  _order: 64, _act: 68,
};

let rows      = [];
let colNames  = { ...DEFAULT_COLS };  // rename map for fixed var-cols
let customCols= [];  // [{ id, label, type }]  user-added extra columns
let colWidths = {};  // { fieldId: px }
let filter    = 'all';
let history   = [];
let future    = [];
let toastTimer= null;

// Record panel state
let recIdx    = -1;   // index in visibleRows() currently shown
let recDirty  = false;

/* ════════════════════════════════════════
   DOM
════════════════════════════════════════ */
const tableBody  = document.getElementById('tableBody');
const emptyState = document.getElementById('emptyState');
const statusDot  = document.getElementById('statusDot');
const statusText = document.getElementById('statusText');
const colgroup   = document.getElementById('colgroup');
const headerRow  = document.getElementById('headerRow');

const undoBtn = document.getElementById('undoBtn');
const redoBtn = document.getElementById('redoBtn');
const confirmDlg = document.getElementById('confirmDialog');
const addColDlg  = document.getElementById('addColDialog');
const recOverlay = document.getElementById('recPanel');
const toast      = document.getElementById('toast');

/* ════════════════════════════════════════
   Utilities
════════════════════════════════════════ */
function uid()    { return Date.now().toString(36) + Math.random().toString(36).slice(2,5); }
function nowIso() { return new Date().toISOString(); }
function cloneRows()  { return JSON.parse(JSON.stringify(rows)); }
function cloneState() { return { rows: cloneRows(), cols: { ...colNames }, custom: JSON.parse(JSON.stringify(customCols)) }; }

function fmtDate(iso) {
  if (!iso) return '';
  return new Date(iso).toLocaleString('ja-JP', {
    year:'2-digit', month:'2-digit', day:'2-digit', hour:'2-digit', minute:'2-digit'
  }).replace(/\//g,'-');
}

function esc(v) {
  return (v??'').toString()
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function csvQ(v) {
  // Normalize line endings inside the value to \r\n (RFC 4180)
  const s = String(v??'').replace(/\r\n/g,'\n').replace(/\r/g,'\n').replace(/\n/g,'\r\n');
  return (s.includes(',') || s.includes('"') || s.includes('\r') || s.includes('\n'))
    ? '"'+s.replace(/"/g,'""')+'"' : s;
}

function showToast(msg, type='') {
  toast.textContent = msg;
  toast.className   = 'toast' + (type ? ' '+type : '');
  toast.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { toast.hidden = true; }, 2600);
}

/* ── Confirm ── */
function openConfirm(title, msg) {
  return new Promise(resolve => {
    document.getElementById('confirmTitle').textContent   = title;
    document.getElementById('confirmMessage').textContent = msg;
    confirmDlg.hidden = false;
    const ok  = document.getElementById('confirmOk');
    const cnl = document.getElementById('confirmCancel');
    const done = r => { confirmDlg.hidden = true; ok.onclick = cnl.onclick = null; resolve(r); };
    ok.onclick  = () => done(true);
    cnl.onclick = () => done(false);
  });
}

/* ════════════════════════════════════════
   History
════════════════════════════════════════ */
function snapshot() {
  history.push(cloneState());
  if (history.length > HIST_LIMIT) history.shift();
  future = [];
  updateUndoRedo();
}

function restoreSnap(s) {
  rows      = JSON.parse(JSON.stringify(s.rows));
  colNames  = { ...s.cols };
  customCols= JSON.parse(JSON.stringify(s.custom));
  save(); applyColNames(); rebuildHeader(); render(); updateUndoRedo();
}

function undo() { if (!history.length) return; future.push(cloneState()); restoreSnap(history.pop()); }
function redo() { if (!future.length)  return; history.push(cloneState()); restoreSnap(future.pop()); }

function updateUndoRedo() {
  undoBtn.disabled = !history.length;
  redoBtn.disabled = !future.length;
}

/* ── Clear storage button ── */
document.getElementById('clearStorageBtn').addEventListener('click', async () => {
  if (!rows.length) { showToast('消去するデータがありません', 'err'); return; }
  const ok = await openConfirm(
    'ストレージを消去',
    `${rows.length} 件のデータをすべて削除します。この操作は取り消せません。続行しますか？`
  );
  if (!ok) return;
  rows = []; history = []; future = [];
  save(); render(); updateUndoRedo();
  showToast('ストレージを消去しました');
});

/* ── Import mode dialog (3-choice: append / replace / cancel) ── */
function openImportMode(count) {
  return new Promise(resolve => {
    document.getElementById('importModeMsg').textContent =
      `${count} 件のデータが見つかりました。既存データ（${rows.length} 件）への反映方法を選んでください。`;
    const dlg     = document.getElementById('importModeDialog');
    const btnApp  = document.getElementById('importModeAppend');
    const btnRep  = document.getElementById('importModeReplace');
    const btnCnl  = document.getElementById('importModeCancel');
    dlg.hidden = false;
    const done = r => {
      dlg.hidden = true;
      btnApp.onclick = btnRep.onclick = btnCnl.onclick = null;
      resolve(r);  // 'append' | 'replace' | null
    };
    btnApp.onclick = () => done('append');
    btnRep.onclick = () => done('replace');
    btnCnl.onclick = () => done(null);
  });
}
document.addEventListener('keydown', e => {
  if (e.key === 'Escape' && !document.getElementById('importModeDialog').hidden) {
    document.getElementById('importModeDialog').hidden = true;
  }
});

/* ════════════════════════════════════════
   Storage
════════════════════════════════════════ */
function save() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
  localStorage.setItem(COL_KEY,     JSON.stringify({ names: colNames, custom: customCols }));
  localStorage.setItem(WIDTH_KEY,   JSON.stringify(colWidths));
  updateStatus();
}

function load() {
  try { rows = JSON.parse(localStorage.getItem(STORAGE_KEY)) || []; } catch { rows = []; }
  try {
    const cd = JSON.parse(localStorage.getItem(COL_KEY)) || {};
    colNames   = { ...DEFAULT_COLS, ...(cd.names  || {}) };
    customCols = cd.custom || [];
  } catch { colNames = { ...DEFAULT_COLS }; customCols = []; }
  try { colWidths = JSON.parse(localStorage.getItem(WIDTH_KEY)) || {}; } catch { colWidths = {}; }
  const theme = localStorage.getItem(THEME_KEY) || 'dark';
  document.documentElement.setAttribute('data-theme', theme);
  updateThemeIcon(theme);
}

function updateStatus() {
  const n = rows.length;
  statusText.textContent = n + ' 件';
  statusDot.classList.toggle('active', n > 0);
}

/* ════════════════════════════════════════
   Theme
════════════════════════════════════════ */
function updateThemeIcon(t) {
  document.getElementById('iconSun').style.display  = t==='dark'  ? 'block' : 'none';
  document.getElementById('iconMoon').style.display = t==='light' ? 'block' : 'none';
}
document.getElementById('themeBtn').addEventListener('click', () => {
  const cur  = document.documentElement.getAttribute('data-theme');
  const next = cur==='dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', next);
  localStorage.setItem(THEME_KEY, next);
  updateThemeIcon(next);
});

/* ════════════════════════════════════════
   Column width helper
════════════════════════════════════════ */
function getWidth(id) { return colWidths[id] ?? DEFAULT_WIDTHS[id] ?? 100; }

/* ════════════════════════════════════════
   Column Name Bar (fixed 4)
════════════════════════════════════════ */
const colInputs = {
  evente: document.getElementById('colEventE'),
  ne:     document.getElementById('colNe'),
  eventc: document.getElementById('colEventC'),
  nc:     document.getElementById('colNc'),
};

function applyColNames() {
  Object.keys(colInputs).forEach(k => {
    colInputs[k].value = colNames[k];
  });
  // Update record panel labels
  document.getElementById('rfLabelEE').textContent = colNames.evente;
  document.getElementById('rfLabelNe').textContent = colNames.ne;
  document.getElementById('rfLabelEC').textContent = colNames.eventc;
  document.getElementById('rfLabelNc').textContent = colNames.nc;
  // custom col bar items already rebuilt in rebuildColNameBar
}

Object.keys(colInputs).forEach(k => {
  colInputs[k].addEventListener('change', () => {
    snapshot();
    colNames[k] = colInputs[k].value.trim() || DEFAULT_COLS[k];
    applyColNames(); rebuildHeader(); save();
  });
});

document.getElementById('resetColBtn').addEventListener('click', () => {
  snapshot();
  colNames = { ...DEFAULT_COLS };
  applyColNames(); rebuildHeader(); save();
  showToast('列名をリセットしました');
});

/* ════════════════════════════════════════
   Custom Columns
════════════════════════════════════════ */
document.getElementById('addColBtn').addEventListener('click', () => {
  document.getElementById('newColName').value = '';
  document.getElementById('newColType').value = 'text';
  addColDlg.hidden = false;
  setTimeout(() => document.getElementById('newColName').focus(), 50);
});

document.getElementById('addColCancel').addEventListener('click', () => { addColDlg.hidden = true; });

document.getElementById('addColOk').addEventListener('click', () => {
  const label = document.getElementById('newColName').value.trim();
  if (!label) { showToast('列名を入力してください', 'err'); return; }
  // check duplicate
  const allKeys = [...FIXED_FIELDS, ...customCols.map(c=>c.id)];
  if (allKeys.includes(label)) { showToast('同じ名前の列がすでにあります', 'err'); return; }
  const type  = document.getElementById('newColType').value;
  const id    = label; // use label as id (csv header)
  snapshot();
  customCols.push({ id, label, type });
  // add field to all existing rows
  rows.forEach(r => { if (!(id in r)) r[id] = ''; });
  addColDlg.hidden = true;
  rebuildColNameBar(); rebuildHeader(); render(); save();
  showToast(`列「${label}」を追加しました`, 'ok');
});

function deleteCustomCol(id) {
  snapshot();
  customCols = customCols.filter(c => c.id !== id);
  rows.forEach(r => { delete r[id]; });
  rebuildColNameBar(); rebuildHeader(); render(); save();
  showToast('列を削除しました');
}

function rebuildColNameBar() {
  // Remove existing custom col items
  document.querySelectorAll('.col-name-item[data-custom]').forEach(el => el.remove());
  const container = document.getElementById('colNameInputs');
  customCols.forEach(col => {
    const div = document.createElement('div');
    div.className = 'col-name-item';
    div.dataset.custom = col.id;
    div.innerHTML = `<span class="col-orig">${esc(col.id)}</span><button class="del-col-btn" title="この列を削除">✕</button>`;
    div.querySelector('.del-col-btn').addEventListener('click', async () => {
      const ok = await openConfirm('列を削除', `列「${col.label}」とその全データを削除しますか？`);
      if (ok) deleteCustomCol(col.id);
    });
    container.appendChild(div);
  });
}

/* ════════════════════════════════════════
   Header rebuild
════════════════════════════════════════ */
function rebuildHeader() {
  // Build colgroup
  colgroup.innerHTML = '';
  headerRow.innerHTML = '';

  const cols = buildColDef();
  cols.forEach(c => {
    const col = document.createElement('col');
    col.style.width = getWidth(c.id) + 'px';
    col.dataset.colId = c.id;
    colgroup.appendChild(col);

    const th = document.createElement('th');
    th.dataset.colId = c.id;
    if (c.cls) th.className = c.cls;
    th.innerHTML = c.html;

    // Resize handle (skip action/order/rownum/inc cols)
    if (!['_rownum','_inc','_order','_act'].includes(c.id)) {
      const handle = document.createElement('div');
      handle.className = 'col-resize';
      handle.dataset.colId = c.id;
      th.appendChild(handle);
      attachResizeHandle(handle, c.id);
    }
    headerRow.appendChild(th);
  });
}

// Column definition list
function buildColDef() {
  return [
    { id:'_rownum', cls:'', html:'<span style="font-size:10px;color:var(--text3)">#</span>' },
    { id:'_inc',    cls:'', html:'<span title="Include">✓</span>' },
    { id:'year',    cls:'', html:'Year' },
    { id:'study',   cls:'', html:'Study' },
    { id:'region',  cls:'', html:'Region' },
    { id:'url',     cls:'', html:'URL' },
    { id:'notes',   cls:'', html:'Notes' },
    { id:'evente',  cls:'', html: esc(colNames.evente) },
    { id:'ne',      cls:'', html: esc(colNames.ne) },
    { id:'eventc',  cls:'', html: esc(colNames.eventc) },
    { id:'nc',      cls:'', html: esc(colNames.nc) },
    ...customCols.map(c => ({ id: c.id, cls:'', html: esc(c.label) })),
    { id:'_order',  cls:'', html:'' },
    { id:'_act',    cls:'', html:'' },
  ];
}

/* ════════════════════════════════════════
   Column resize
════════════════════════════════════════ */
function attachResizeHandle(handle, colId) {
  handle.addEventListener('mousedown', e => {
    e.preventDefault();
    handle.classList.add('active');
    const startX  = e.clientX;
    const colEl   = colgroup.querySelector(`col[data-col-id="${colId}"]`);
    const startW  = parseInt(colEl?.style.width) || getWidth(colId);

    const onMove = e => {
      const newW = Math.max(40, startX - e.clientX + startW);  // min 40px
      colWidths[colId] = newW;
      // update col element width live
      colgroup.querySelectorAll(`col[data-col-id="${colId}"]`).forEach(c => c.style.width = newW+'px');
    };
    const onUp = () => {
      handle.classList.remove('active');
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      save();
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  });
}

/* ════════════════════════════════════════
   Row factory
════════════════════════════════════════ */
function makeRow(o = {}) {
  const now = nowIso();
  const base = { id:uid(), year:'', study:'', region:'', url:'', notes:'',
    evente:'', ne:'', eventc:'', nc:'', include:true, created_at:now, updated_at:now };
  customCols.forEach(c => { base[c.id] = ''; });
  return Object.assign(base, o);
}

/* ════════════════════════════════════════
   Filter
════════════════════════════════════════ */
document.querySelectorAll('.filter-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    filter = btn.dataset.filter;
    render();
  });
});

function visibleRows() {
  if (filter==='include') return rows.filter(r=>r.include);
  if (filter==='exclude') return rows.filter(r=>!r.include);
  return rows;
}

/* ════════════════════════════════════════
   Render table
════════════════════════════════════════ */
function render() {
  const vis = visibleRows();
  tableBody.innerHTML = '';
  emptyState.classList.toggle('visible', vis.length===0);

  vis.forEach((row, visIdx) => {
    const realIdx = rows.indexOf(row);
    const tr = document.createElement('tr');
    tr.dataset.id = row.id;
    if (!row.include) tr.classList.add('excluded');

    // Build cells
    const urlVisible = row.url ? 'visible' : '';

    let customCells = customCols.map(c => `
      <td>
        <input class="cell-input ${c.type==='number'?'cell-mono':''}"
          type="${c.type}" data-field="${esc(c.id)}"
          value="${esc(row[c.id]??'')}" placeholder="—">
      </td>`).join('');

    tr.innerHTML = `
      <td class="td-rownum" title="クリックで詳細表示">${visIdx+1}</td>
      <td><div class="inc-wrap"><input type="checkbox" class="inc-check" ${row.include?'checked':''}></div></td>
      <td><input class="cell-input" type="number" data-field="year" value="${esc(row.year)}" placeholder="year" min="1900" max="2100"></td>
      <td><input class="cell-input" data-field="study"  value="${esc(row.study)}"  placeholder="study名"></td>
      <td><input class="cell-input" data-field="region" value="${esc(row.region)}" placeholder="region"></td>
      <td><div class="url-cell">
        <input class="cell-input" data-field="url" value="${esc(row.url)}" placeholder="https://...">
        <a class="url-open ${urlVisible}" href="${esc(row.url)}" target="_blank" rel="noopener">↗</a>
      </div></td>
      <td><input class="cell-input" data-field="notes"  value="${esc(row.notes)}"  placeholder="notes"></td>
      <td><input class="cell-input cell-mono" type="number" data-field="evente" value="${esc(row.evente)}" placeholder="—"></td>
      <td><input class="cell-input cell-mono" type="number" data-field="ne"     value="${esc(row.ne)}"     placeholder="—"></td>
      <td><input class="cell-input cell-mono" type="number" data-field="eventc" value="${esc(row.eventc)}" placeholder="—"></td>
      <td><input class="cell-input cell-mono" type="number" data-field="nc"     value="${esc(row.nc)}"     placeholder="—"></td>
      ${customCells}
      <td><div class="order-btns">
        <button class="order-btn up-btn"   ${realIdx===0             ?'disabled':''}>▲</button>
        <button class="order-btn down-btn" ${realIdx===rows.length-1 ?'disabled':''}>▼</button>
      </div></td>
      <td><div class="act-btns">
        <button class="act-btn clone" title="複製">⧉</button>
        <button class="act-btn del"   title="削除">✕</button>
      </div></td>`;

    /* ── Cell edit ── */
    tr.querySelectorAll('.cell-input').forEach(inp => {
      inp.addEventListener('focus', () => { inp._snap = false; });
      inp.addEventListener('input', () => {
        if (!inp._snap) { snapshot(); inp._snap = true; }
        const f = inp.dataset.field;
        if (!f) return;
        row[f] = inp.value;
        row.updated_at = nowIso();
        if (f==='url') {
          const a = tr.querySelector('.url-open');
          a.href = inp.value;
          a.classList.toggle('visible', !!inp.value.trim());
        }
        save();
      });
    });

    /* ── Include ── */
    tr.querySelector('.inc-check').addEventListener('change', e => {
      snapshot();
      row.include = e.target.checked;
      row.updated_at = nowIso();
      tr.classList.toggle('excluded', !row.include);
      save();
    });

    /* ── Row number → open record panel ── */
    tr.querySelector('.td-rownum').addEventListener('click', () => openRecPanel(visIdx));

    /* ── Order ── */
    tr.querySelector('.up-btn').addEventListener('click',   () => moveRow(realIdx,-1));
    tr.querySelector('.down-btn').addEventListener('click', () => moveRow(realIdx, 1));

    /* ── Clone ── */
    tr.querySelector('.clone').addEventListener('click', () => {
      snapshot();
      const clone = { ...row, id:uid(), created_at:nowIso(), updated_at:nowIso() };
      rows.splice(realIdx+1, 0, clone);
      save(); render();
      showToast('行を複製しました','ok');
    });

    /* ── Delete ── */
    tr.querySelector('.del').addEventListener('click', async () => {
      const ok = await openConfirm('行を削除', `「${row.study||'(未入力)'}」を削除しますか？`);
      if (!ok) return;
      snapshot(); rows.splice(realIdx,1); save(); render();
      showToast('削除しました');
    });

    tableBody.appendChild(tr);
  });
}

function moveRow(idx, dir) {
  const n = idx+dir;
  if (n<0||n>=rows.length) return;
  snapshot();
  [rows[idx],rows[n]] = [rows[n],rows[idx]];
  save(); render();
}

/* ════════════════════════════════════════
   Add Row
════════════════════════════════════════ */
document.getElementById('addRowBtn').addEventListener('click', () => {
  snapshot();
  rows.push(makeRow());
  save(); render();
  const wrap = document.querySelector('.table-scroll');
  wrap.scrollTop = wrap.scrollHeight;
  setTimeout(() => tableBody.lastElementChild?.querySelector('.cell-input')?.focus(), 30);
});

/* ════════════════════════════════════════
   Record Panel
════════════════════════════════════════ */
function openRecPanel(visIdx) {
  const vis = visibleRows();
  if (!vis.length) return;
  recIdx = Math.max(0, Math.min(visIdx, vis.length-1));
  loadRecPanel(vis[recIdx]);
  recOverlay.hidden = false;
  setTimeout(() => document.getElementById('rfStudy').focus(), 60);
}

function loadRecPanel(row) {
  recDirty = false;
  document.getElementById('recDirty').hidden = true;

  document.getElementById('rfInclude').checked = !!row.include;
  document.getElementById('rfIncLabel').textContent = row.include ? 'Include' : 'Exclude';
  document.getElementById('rfYear').value   = row.year    ?? '';
  document.getElementById('rfStudy').value  = row.study   ?? '';
  document.getElementById('rfRegion').value = row.region  ?? '';
  document.getElementById('rfUrl').value    = row.url     ?? '';
  document.getElementById('rfNotes').value  = row.notes   ?? '';
  document.getElementById('rfEventE').value = row.evente  ?? '';
  document.getElementById('rfNe').value     = row.ne      ?? '';
  document.getElementById('rfEventC').value = row.eventc  ?? '';
  document.getElementById('rfNc').value     = row.nc      ?? '';

  // URL open btn
  const urlOpen = document.getElementById('rfUrlOpen');
  urlOpen.href = row.url || '#';
  urlOpen.style.pointerEvents = row.url ? '' : 'none';
  urlOpen.style.opacity = row.url ? '1' : '0.4';

  // Custom cols
  const customRow = document.getElementById('rfCustomRow');
  customRow.innerHTML = '';
  customCols.forEach(col => {
    const div = document.createElement('div');
    div.className = 'rec-field';
    div.innerHTML = `<label>${esc(col.label)}</label>
      <input type="${col.type}" class="rf-input ${col.type==='number'?'rf-mono':''}" data-custom-field="${esc(col.id)}" value="${esc(row[col.id]??'')}">`;
    div.querySelector('input').addEventListener('input', markDirty);
    customRow.appendChild(div);
  });

  // Nav position
  const vis = visibleRows();
  document.getElementById('recPos').textContent   = `${recIdx+1} / ${vis.length}`;
  document.getElementById('recPrev').disabled = recIdx === 0;
  document.getElementById('recNext').disabled = recIdx === vis.length-1;

  // Title & meta
  document.getElementById('recTitle').textContent =
    row.study ? row.study : '(タイトル未設定)';
  document.getElementById('recMeta').innerHTML =
    `作成: ${fmtDate(row.created_at)}<br>更新: ${fmtDate(row.updated_at)}`;

  // Attach dirty listeners
  ['rfYear','rfStudy','rfRegion','rfUrl','rfNotes','rfEventE','rfNe','rfEventC','rfNc'].forEach(id => {
    const el = document.getElementById(id);
    el.oninput = null;
    el.addEventListener('input', markDirty);
  });
  document.getElementById('rfInclude').onchange = markDirty;

  // URL input → update open btn
  document.getElementById('rfUrl').addEventListener('input', function() {
    urlOpen.href = this.value || '#';
    urlOpen.style.pointerEvents = this.value ? '' : 'none';
    urlOpen.style.opacity = this.value ? '1' : '0.4';
  });
}

function markDirty() {
  recDirty = true;
  document.getElementById('recDirty').hidden = false;
  // live update title
  const study = document.getElementById('rfStudy').value.trim();
  document.getElementById('recTitle').textContent = study || '(タイトル未設定)';
  const incChecked = document.getElementById('rfInclude').checked;
  document.getElementById('rfIncLabel').textContent = incChecked ? 'Include' : 'Exclude';
}

function saveRecPanel() {
  const vis = visibleRows();
  const row = vis[recIdx];
  if (!row) return;
  snapshot();
  row.year    = document.getElementById('rfYear').value;
  row.study   = document.getElementById('rfStudy').value.trim();
  row.region  = document.getElementById('rfRegion').value.trim();
  row.url     = document.getElementById('rfUrl').value.trim();
  row.notes   = document.getElementById('rfNotes').value;
  row.evente  = document.getElementById('rfEventE').value;
  row.ne      = document.getElementById('rfNe').value;
  row.eventc  = document.getElementById('rfEventC').value;
  row.nc      = document.getElementById('rfNc').value;
  row.include = document.getElementById('rfInclude').checked;
  // custom cols
  document.querySelectorAll('#rfCustomRow input[data-custom-field]').forEach(inp => {
    row[inp.dataset.customField] = inp.value;
  });
  row.updated_at = nowIso();
  recDirty = false;
  document.getElementById('recDirty').hidden = true;
  document.getElementById('recMeta').innerHTML =
    `作成: ${fmtDate(row.created_at)}<br>更新: ${fmtDate(row.updated_at)}`;
  save(); render();
  showToast('保存しました','ok');
}

async function closeRecPanel() {
  if (recDirty) {
    const ok = await openConfirm('変更を破棄', '保存していない変更があります。閉じますか？');
    if (!ok) return;
  }
  recOverlay.hidden = true;
  recIdx = -1; recDirty = false;
}

document.getElementById('recSaveBtn').addEventListener('click', saveRecPanel);
document.getElementById('recClose').addEventListener('click', closeRecPanel);
recOverlay.addEventListener('click', e => { if (e.target===recOverlay) closeRecPanel(); });

document.getElementById('recPrev').addEventListener('click', () => {
  if (recIdx > 0) { if (recDirty) saveRecPanel(); recIdx--; loadRecPanel(visibleRows()[recIdx]); }
});
document.getElementById('recNext').addEventListener('click', () => {
  const vis = visibleRows();
  if (recIdx < vis.length-1) { if (recDirty) saveRecPanel(); recIdx++; loadRecPanel(vis[recIdx]); }
});

/* ════════════════════════════════════════
   Keyboard shortcuts
════════════════════════════════════════ */
document.addEventListener('keydown', e => {
  // Esc
  if (e.key==='Escape') {
    if (!confirmDlg.hidden)  { confirmDlg.hidden = true; return; }
    if (!addColDlg.hidden)   { addColDlg.hidden  = true; return; }
    if (!recOverlay.hidden)  { closeRecPanel(); return; }
  }
  // Ctrl+Enter → save record panel
  if ((e.ctrlKey||e.metaKey) && e.key==='Enter' && !recOverlay.hidden) {
    e.preventDefault(); saveRecPanel(); return;
  }
  // Arrow keys in record panel
  if (!recOverlay.hidden && !e.target.matches('input,textarea,select')) {
    if (e.key==='ArrowLeft')  { document.getElementById('recPrev').click(); return; }
    if (e.key==='ArrowRight') { document.getElementById('recNext').click(); return; }
  }
  // Undo/Redo (only when modals closed)
  if (confirmDlg.hidden && addColDlg.hidden) {
    if ((e.ctrlKey||e.metaKey) && !e.shiftKey && e.key==='z') { e.preventDefault(); undo(); }
    if ((e.ctrlKey||e.metaKey) && (e.key==='y'||(e.shiftKey&&e.key==='z'))) { e.preventDefault(); redo(); }
  }
});

// Add col dialog Enter key
document.getElementById('newColName').addEventListener('keydown', e => {
  if (e.key==='Enter') document.getElementById('addColOk').click();
});

/* ════════════════════════════════════════
   Export CSV
════════════════════════════════════════ */
async function exportRows(subset, label) {
  if (!subset.length) { showToast('エクスポートするデータがありません','err'); return; }
  const ok = await openConfirm(`Export（${label}）`,
    'エクスポート後、ローカルストレージのデータは削除されます。続行しますか？');
  if (!ok) return;

  const varHeaders = [colNames.evente, colNames.ne, colNames.eventc, colNames.nc];
  const customHeaders = customCols.map(c=>c.label);
  const headers = ['year','study','region','url','notes', ...varHeaders,
    ...customHeaders, 'include','created_at','updated_at'];

  const lines = [
    headers.join(','),
    ...subset.map(r => [
      r.year, r.study, r.region, r.url, r.notes,
      r.evente, r.ne, r.eventc, r.nc,
      ...customCols.map(c => r[c.id]??''),
      r.include ? '1' : '0',
      r.created_at, r.updated_at,
    ].map(csvQ).join(','))
  ];

  const blob = new Blob(['\uFEFF'+lines.join('\r\n')], { type:'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url; a.download = `study_${new Date().toISOString().slice(0,10)}.csv`;
  a.click(); URL.revokeObjectURL(url);

  rows=[]; history=[]; future=[];
  save(); render(); updateUndoRedo();
  showToast(`${subset.length} 件をエクスポートしました。ストレージをクリアしました`,'ok');
}

document.getElementById('exportAllBtn').addEventListener('click', () => exportRows(rows,'全件'));
document.getElementById('exportIncBtn').addEventListener('click', () => exportRows(rows.filter(r=>r.include),'Include のみ'));

/* ════════════════════════════════════════
   Import CSV
════════════════════════════════════════ */
document.getElementById('importFile').addEventListener('change', async function() {
  const file = this.files[0];
  if (!file) return;
  this.value = '';
  let text;
  try { text = await file.text(); }
  catch { showToast('読み込み失敗','err'); return; }

  const parsed = parseCSV(text);
  if (!parsed||parsed.length<2) { showToast('CSV解析失敗','err'); return; }

  const headers = parsed[0].map(h=>h.trim().replace(/^\uFEFF/,''));
  const dataRows = parsed.slice(1).filter(r=>r.some(c=>c.trim()));

  // Identify fixed var cols vs custom cols
  const systemCols = new Set(['year','study','region','url','notes','include','created_at','updated_at']);
  // First 4 non-system cols → fixed var cols; rest → custom
  const nonSystem = headers.filter(h=>!systemCols.has(h));
  const impFixed = { evente: nonSystem[0]||'evente', ne: nonSystem[1]||'ne',
                     eventc: nonSystem[2]||'eventc', nc: nonSystem[3]||'nc' };
  const impCustomIds = nonSystem.slice(4);

  const hi = k => headers.indexOf(k);
  const g  = (cells, key) => (cells[hi(key)]??'').trim();

  const newRows = dataRows.map(cells => {
    const base = makeRow({
      year:   g(cells,'year'), study:  g(cells,'study'), region: g(cells,'region'),
      url:    g(cells,'url'),  notes:  g(cells,'notes'),
      evente: g(cells,impFixed.evente)||g(cells,'evente'),
      ne:     g(cells,impFixed.ne)    ||g(cells,'ne'),
      eventc: g(cells,impFixed.eventc)||g(cells,'eventc'),
      nc:     g(cells,impFixed.nc)    ||g(cells,'nc'),
      include: g(cells,'include')!=='0',
      created_at: g(cells,'created_at')||nowIso(),
      updated_at: g(cells,'updated_at')||nowIso(),
    });
    impCustomIds.forEach(id => { base[id] = g(cells,id); });
    return base;
  });

  // Ask user: append / replace / cancel
  let mode;
  if (rows.length > 0) {
    mode = await openImportMode(newRows.length);
    if (!mode) return;  // cancelled
  } else {
    mode = 'replace';
  }

  snapshot();
  if (mode === 'append') {
    // Merge custom cols
    impCustomIds.forEach(id => {
      if (!customCols.find(c=>c.id===id)) {
        customCols.push({ id, label:id, type:'text' });
        rows.forEach(r => { if (!(id in r)) r[id]=''; });
      }
    });
    rows = [...rows, ...newRows];
  } else {
    rows = newRows;
    colNames = { ...DEFAULT_COLS, ...impFixed };
    customCols = impCustomIds.map(id => ({ id, label:id, type:'text' }));
  }

  applyColNames(); rebuildColNameBar(); rebuildHeader(); render(); save();
  showToast(`${newRows.length} 件をインポートしました`,'ok');
});

/* ─ CSV parser (RFC 4180 compliant — handles embedded newlines/commas in quoted fields) ─ */
function parseCSV(text) {
  // Normalize all line endings to \n, then parse character-by-character
  const src = text.replace(/\r\n/g,'\n').replace(/\r/g,'\n');
  const records = [];
  let fields = [];
  let cur    = '';
  let inQ    = false;
  let i      = 0;

  while (i < src.length) {
    const c = src[i];

    if (inQ) {
      if (c === '"') {
        if (src[i+1] === '"') {          // escaped quote ""
          cur += '"'; i += 2;
        } else {                          // closing quote
          inQ = false; i++;
        }
      } else {
        cur += c; i++;                    // any char inside quotes, including \n
      }
    } else {
      if (c === '"') {
        inQ = true; i++;                  // opening quote
      } else if (c === ',') {
        fields.push(cur); cur = ''; i++;  // field separator
      } else if (c === '\n') {
        fields.push(cur); cur = '';       // end of record
        if (fields.some(f => f.trim())) records.push(fields);
        fields = []; i++;
      } else {
        cur += c; i++;
      }
    }
  }

  // Last field / record (file may not end with newline)
  fields.push(cur);
  if (fields.some(f => f.trim())) records.push(fields);

  return records;
}

/* ════════════════════════════════════════
   HTML Export / Print
════════════════════════════════════════ */

/** Build a self-contained HTML string from current rows */
function buildHtmlReport(subset) {
  const varHeaders = [colNames.evente, colNames.ne, colNames.eventc, colNames.nc];
  const customHeaders = customCols.map(c => c.label);
  const allHeaders = ['#', 'Inc', 'Year', 'Study', 'Region', 'Notes',
    ...varHeaders, ...customHeaders, 'URL'];

  // Compute column count for colgroup
  const colCount = allHeaders.length;

  const escH = v => (v??'').toString()
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

  const rows_html = subset.map((r, i) => {
    const inc = r.include;
    const urlCell = r.url
      ? `<a href="${escH(r.url)}" target="_blank" rel="noopener">${escH(r.url.replace(/^https?:\/\//,'').substring(0,40))}${r.url.length>50?'…':''}</a>`
      : '';
    const varCells = [r.evente, r.ne, r.eventc, r.nc].map(v =>
      `<td class="num">${escH(v)}</td>`).join('');
    const customCells = customCols.map(c =>
      `<td>${escH(r[c.id]??'')}</td>`).join('');
    // Notes: preserve line breaks
    const notesEsc = escH(r.notes||'').replace(/\n/g,'<br>');
    return `<tr class="${inc?'inc':'exc'}">
      <td class="num">${i+1}</td>
      <td class="inc-cell">${inc?'<span class="badge-inc">✓</span>':'<span class="badge-exc">✕</span>'}</td>
      <td class="num">${escH(r.year)}</td>
      <td class="study-cell">${escH(r.study)}</td>
      <td>${escH(r.region)}</td>
      <td class="notes-cell">${notesEsc}</td>
      ${varCells}
      ${customCells}
      <td class="url-cell">${urlCell}</td>
    </tr>`;
  }).join('\n');

  const theadCells = allHeaders.map(h => `<th>${escH(h)}</th>`).join('');

  const incCount = subset.filter(r=>r.include).length;
  const excCount = subset.length - incCount;
  const ts = new Date().toLocaleString('ja-JP');

  return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>MetaCSVBuilder Export — ${escH(ts)}</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=DM+Sans:wght@400;500;600&display=swap');

:root {
  --acc:  #3a72e8;
  --acc2: #0fb885;
  --bg:   #f7f9fc;
  --surf: #ffffff;
  --bdr:  #d8dff0;
  --txt:  #1a2035;
  --txt2: #5a6582;
  --txt3: #9aa3ba;
  --inc-bg: #f0faf6;
  --exc-bg: #fafafa;
  --exc-c:  #b0b8cc;
}

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

body {
  font-family: 'DM Sans', sans-serif;
  background: var(--bg);
  color: var(--txt);
  font-size: 13px;
  line-height: 1.55;
  padding: 24px 28px 48px;
}

/* ── Report header ── */
.report-header {
  margin-bottom: 20px;
  padding-bottom: 14px;
  border-bottom: 2px solid var(--acc);
  display: flex;
  justify-content: space-between;
  align-items: flex-end;
  flex-wrap: wrap;
  gap: 8px;
}
.report-title {
  font-size: 18px;
  font-weight: 700;
  color: var(--txt);
  letter-spacing: -.3px;
}
.report-title span { color: var(--acc); }
.report-meta {
  font-size: 11px;
  color: var(--txt3);
  font-family: 'IBM Plex Mono', monospace;
  text-align: right;
  line-height: 1.8;
}

/* ── Summary chips ── */
.summary-row {
  display: flex;
  gap: 10px;
  margin-bottom: 18px;
  flex-wrap: wrap;
}
.chip {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  padding: 5px 12px;
  border-radius: 100px;
  font-size: 12px;
  font-weight: 600;
  border: 1px solid;
}
.chip-total { background: #eef2ff; border-color: #c5d0f5; color: var(--acc); }
.chip-inc   { background: #e8faf4; border-color: #9de8cc; color: #0a7a58; }
.chip-exc   { background: #f5f5f5; border-color: #d0d0d0; color: #888; }

/* ── Table ── */
.table-wrap {
  overflow-x: auto;
  border-radius: 8px;
  border: 1px solid var(--bdr);
  box-shadow: 0 2px 12px rgba(0,0,0,.06);
}

table {
  width: 100%;
  border-collapse: collapse;
  background: var(--surf);
  font-size: 12.5px;
}

thead {
  position: sticky;
  top: 0;
  z-index: 2;
}

th {
  background: #1a2035;
  color: #c8d0e8;
  font-size: 10.5px;
  font-weight: 700;
  letter-spacing: .7px;
  text-transform: uppercase;
  padding: 10px 11px;
  text-align: left;
  border-right: 1px solid #2c3654;
  white-space: nowrap;
}
th:last-child { border-right: none; }

td {
  padding: 8px 11px;
  border-bottom: 1px solid var(--bdr);
  border-right: 1px solid #edf0f8;
  vertical-align: top;
}
td:last-child { border-right: none; }

/* Row styles */
tr.inc { background: var(--surf); }
tr.inc:hover { background: #f5f8ff; }
tr.exc { background: var(--exc-bg); }
tr.exc td { color: var(--exc-c); }
tr.exc:hover { background: #f2f2f2; }
tr:last-child td { border-bottom: none; }

/* Specific cells */
.num   { font-family: 'IBM Plex Mono', monospace; font-size: 12px; text-align: right; white-space: nowrap; }
.study-cell { font-weight: 600; min-width: 160px; }
.notes-cell { min-width: 200px; max-width: 340px; font-size: 12px; color: var(--txt2); word-break: break-word; }
.url-cell a { color: var(--acc); text-decoration: none; font-size: 11.5px; word-break: break-all; }
.url-cell a:hover { text-decoration: underline; }

.inc-cell { text-align: center; }
.badge-inc { display: inline-block; background: #e8faf4; color: #0a7a58; border: 1px solid #9de8cc; border-radius: 4px; padding: 1px 6px; font-size: 11px; font-weight: 700; }
.badge-exc { display: inline-block; background: #f5f5f5; color: #aaa; border: 1px solid #ddd; border-radius: 4px; padding: 1px 6px; font-size: 11px; }

/* ── Footer ── */
footer {
  margin-top: 28px;
  padding-top: 14px;
  border-top: 1px solid var(--bdr);
  font-size: 11px;
  color: var(--txt3);
  line-height: 1.7;
}

/* ════════════════════════════════════
   PRINT STYLES
   - A4横向き想定 / 見切れ防止
════════════════════════════════════ */
@media print {
  @page {
    size: A4 landscape;
    margin: 12mm 10mm 15mm;
  }

  body {
    background: #fff;
    font-size: 11px;
    padding: 0;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
    color-adjust: exact;
  }

  .report-header { margin-bottom: 10px; padding-bottom: 8px; }
  .report-title  { font-size: 14px; }
  .summary-row   { margin-bottom: 10px; gap: 6px; }
  .chip          { padding: 3px 9px; font-size: 11px; }

  .table-wrap {
    border-radius: 0;
    box-shadow: none;
    overflow: visible;
  }

  table {
    font-size: 10px;
    table-layout: auto;
    width: 100%;
  }

  th {
    font-size: 9px;
    padding: 6px 8px;
    background: #1a2035 !important;
    color: #c8d0e8 !important;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }

  td {
    padding: 5px 8px;
    font-size: 10px;
    word-break: break-word;
  }

  .notes-cell { max-width: 200px; font-size: 9.5px; }
  .num        { font-size: 10px; }
  .url-cell   { max-width: 120px; font-size: 9px; }
  .url-cell a { word-break: break-all; }

  tr.inc { background: #fff !important; }
  tr.exc { background: #f8f8f8 !important; }
  tr.exc td { color: #aaa !important; }

  /* avoid row split across pages */
  tr { page-break-inside: avoid; break-inside: avoid; }
  thead { display: table-header-group; }
  tfoot { display: table-footer-group; }

  footer { margin-top: 12px; font-size: 9px; }
}
</style>
</head>
<body>
<div class="report-header">
  <div>
    <div class="report-title">MetaCSVBuilder — <span>データエクスポート</span></div>
  </div>
  <div class="report-meta">
    生成日時: ${escH(ts)}<br>
    総行数: ${subset.length} 件 ／ Include: ${incCount} 件 ／ Exclude: ${excCount} 件
  </div>
</div>

<div class="summary-row">
  <span class="chip chip-total">📊 全 ${subset.length} 件</span>
  <span class="chip chip-inc">✓ Include ${incCount} 件</span>
  ${excCount > 0 ? `<span class="chip chip-exc">✕ Exclude ${excCount} 件</span>` : ''}
</div>

<div class="table-wrap">
<table>
  <thead><tr>${theadCells}</tr></thead>
  <tbody>${rows_html}</tbody>
</table>
</div>

<footer>
  <strong>注意事項：</strong>本ファイルはAI（Claude, Anthropic）を用いて作成されたMetaCSVBuilderで生成されました。内容には誤りが含まれる可能性があります。医療上の判断には必ず公式の教科書・文献・専門家の指導を参照してください。二次利用・再配布は自己責任のもとで可能です。その際、本注記の保持を推奨します。
</footer>
</body>
</html>`;
}

/** Show preview dialog */
function openHtmlPreview() {
  if (!rows.length) { showToast('エクスポートするデータがありません','err'); return; }
  const dlg = document.getElementById('htmlPreviewDialog');
  const cb  = document.getElementById('htmlPreviewIncOnly');
  const frame = document.getElementById('htmlPreviewFrame');

  function refreshPreview() {
    const subset = cb.checked ? rows.filter(r=>r.include) : rows;
    if (!subset.length) { frame.srcdoc = '<p style="padding:20px;font-family:sans-serif;color:#888">対象データがありません</p>'; return; }
    frame.srcdoc = buildHtmlReport(subset);
  }

  cb.checked = false;
  cb.onchange = refreshPreview;
  refreshPreview();
  dlg.hidden = false;

  document.getElementById('htmlPrintBtn').onclick = () => {
    frame.contentWindow.print();
  };

  document.getElementById('htmlDlBtn').onclick = () => {
    const subset = cb.checked ? rows.filter(r=>r.include) : rows;
    if (!subset.length) { showToast('対象データがありません','err'); return; }
    const html = buildHtmlReport(subset);
    const blob = new Blob([html], { type:'text/html;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `study_${new Date().toISOString().slice(0,10)}.html`;
    a.click();
    URL.revokeObjectURL(a.href);
    showToast('HTMLを保存しました','ok');
  };
}

document.getElementById('htmlPreviewClose').addEventListener('click', () => {
  document.getElementById('htmlPreviewDialog').hidden = true;
});
document.getElementById('exportHtmlBtn').addEventListener('click', openHtmlPreview);

/* ════════════════════════════════════════
   Init
════════════════════════════════════════ */
load();
applyColNames();
rebuildColNameBar();
rebuildHeader();
render();
updateStatus();
updateUndoRedo();
