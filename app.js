/**
 * app.js — Offline Edition
 * Orchestrates: file upload → parse → preview → export locally
 * No Google Auth, no network calls. 100% offline.
 */

// ─── State ────────────────────────────────────────────────
const state = {
  workbook:    null,
  activeSheet: 0,
  exporting:   false,
};

// ─── DOM ──────────────────────────────────────────────────
const $ = (id) => document.getElementById(id);

const dom = {
  dropZone:         $('dropZone'),
  fileInput:        $('fileInput'),
  fileInfoCard:     $('fileInfoCard'),
  fileName:         $('fileName'),
  fileSize:         $('fileSize'),
  btnClearFile:     $('btnClearFile'),

  sheetTabs:        $('sheetTabs'),
  tableContainer:   $('tableContainer'),
  previewPanel:     $('previewPanel'),
  previewRowCount:  $('previewRowCount'),
  previewColCount:  $('previewColCount'),
  previewSheetCount:$('previewSheetCount'),

  exportPanel:      $('exportPanel'),
  sheetExportList:  $('sheetExportList'),

  btnExportXLSX:    $('btnExportXLSX'),
  btnExportAllCSV:  $('btnExportAllCSV'),
  btnExportHTML:    $('btnExportHTML'),
  btnExportJSON:    $('btnExportJSON'),
  btnExportPDF:     $('btnExportPDF'),

  progressWrap:     $('progressWrap'),
  progressFill:     $('progressFill'),
  progressStatus:   $('progressStatus'),
  progressPct:      $('progressPct'),
};

// ─── Init ─────────────────────────────────────────────────

function init() {
  bindEvents();
}

// ─── Events ───────────────────────────────────────────────

function bindEvents() {
  dom.dropZone.addEventListener('click',     () => dom.fileInput.click());
  dom.dropZone.addEventListener('dragover',  onDragOver);
  dom.dropZone.addEventListener('dragleave', onDragLeave);
  dom.dropZone.addEventListener('drop',      onDrop);
  dom.fileInput.addEventListener('change',   onFileInput);
  dom.btnClearFile.addEventListener('click', (e) => { e.stopPropagation(); clearFile(); });

  // Export buttons
  dom.btnExportXLSX?.addEventListener('click',   handleExportXLSX);
  dom.btnExportAllCSV?.addEventListener('click',  handleExportAllCSV);
  dom.btnExportHTML?.addEventListener('click',    handleExportHTML);
  dom.btnExportJSON?.addEventListener('click',    handleExportJSON);
  dom.btnExportPDF?.addEventListener('click',     handleExportPDF);
}

// ─── Drag & Drop ──────────────────────────────────────────

function onDragOver(e) {
  e.preventDefault();
  dom.dropZone.classList.add('dragover');
}

function onDragLeave(e) {
  if (!dom.dropZone.contains(e.relatedTarget))
    dom.dropZone.classList.remove('dragover');
}

function onDrop(e) {
  e.preventDefault();
  dom.dropZone.classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if (file) processFile(file);
}

function onFileInput(e) {
  const file = e.target.files[0];
  if (file) processFile(file);
  e.target.value = '';
}

// ─── File Processing ──────────────────────────────────────

async function processFile(file) {
  if (!isValidExcel(file)) {
    showToast('⚠️ Please upload a valid .xlsx or .xls file.', 'error');
    return;
  }

  showToast('⏳ Parsing Excel file…', 'info');
  setProgress(10, 'Reading file…');
  dom.progressWrap.classList.add('visible');

  try {
    const workbook = await ExcelParser.parseExcel(file);
    state.workbook    = workbook;
    state.activeSheet = 0;

    setProgress(80, 'Rendering preview…');

    dom.fileName.textContent = file.name;
    dom.fileSize.textContent = ExcelParser.formatFileSize(file.size);
    dom.fileInfoCard.style.display = 'flex';

    renderPreview();
    renderExportPanel();

    setProgress(100, 'Done!');
    setTimeout(() => dom.progressWrap.classList.remove('visible'), 600);

    showToast(`✅ Parsed ${workbook.sheets.length} sheet(s)!`, 'success');
  } catch (err) {
    dom.progressWrap.classList.remove('visible');
    showToast('❌ ' + err.message, 'error');
  }
}

function isValidExcel(file) {
  return /\.(xlsx?|ods)$/i.test(file.name) ||
    file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
    file.type === 'application/vnd.ms-excel';
}

function clearFile() {
  state.workbook    = null;
  state.activeSheet = 0;
  dom.fileInfoCard.style.display = 'none';
  dom.previewPanel.classList.remove('visible');
  dom.exportPanel.classList.remove('visible');
}

// ─── Preview ──────────────────────────────────────────

// Toolbar state
const tableState = { wrapMode: false, colWidth: 'col-md' };

function renderPreview() {
  const wb = state.workbook;
  if (!wb) return;

  dom.previewPanel.classList.add('visible');
  dom.previewSheetCount.textContent = `${wb.sheets.length} sheet${wb.sheets.length !== 1 ? 's' : ''}`;

  dom.sheetTabs.innerHTML = wb.sheets.map((s, i) => `
    <button class="sheet-tab ${i === 0 ? 'active' : ''}"
            onclick="setActiveSheet(${i})"
            title="${esc(s.name)}">
      📋 ${esc(s.name)}
    </button>
  `).join('');

  renderToolbar();
  renderTable();
}

function renderToolbar() {
  const toolbar = document.getElementById('tableToolbar');
  if (!toolbar) return;
  toolbar.innerHTML = `
    <span class="table-toolbar-label">View:</span>
    <button class="toggle-pill ${tableState.wrapMode ? 'active' : ''}" id="btnWrapToggle" onclick="toggleWrapMode()" title="Toggle text wrapping in cells">
      <span class="pill-dot"></span>
      ${tableState.wrapMode ? 'Wrap ON' : 'Wrap OFF'}
    </button>
    <span class="table-toolbar-label" style="margin-left:4px;">Col width:</span>
    <select class="col-width-select" id="colWidthSelect" onchange="setColWidth(this.value)" title="Set maximum column width">
      <option value="col-sm" ${tableState.colWidth === 'col-sm' ? 'selected' : ''}>Narrow (120px)</option>
      <option value="col-md" ${tableState.colWidth === 'col-md' ? 'selected' : ''}>Medium (240px)</option>
      <option value="col-lg" ${tableState.colWidth === 'col-lg' ? 'selected' : ''}>Wide (380px)</option>
      <option value="col-xl" ${tableState.colWidth === 'col-xl' ? 'selected' : ''}>Extra Wide (560px)</option>
    </select>
  `;
}

function toggleWrapMode() {
  tableState.wrapMode = !tableState.wrapMode;
  const wrapper = document.getElementById('tableContainer');
  if (wrapper) wrapper.classList.toggle('wrap-mode', tableState.wrapMode);
  renderToolbar(); // update button label
}
window.toggleWrapMode = toggleWrapMode;

function setColWidth(cls) {
  tableState.colWidth = cls;
  const wrapper = document.getElementById('tableContainer');
  if (wrapper) {
    wrapper.classList.remove('col-sm', 'col-md', 'col-lg', 'col-xl');
    wrapper.classList.add(cls);
  }
}
window.setColWidth = setColWidth;

const TRUNCATE_THRESHOLD = 60; // chars before we consider a cell "long"

function renderTable() {
  const sheet = state.workbook?.sheets[state.activeSheet];
  if (!sheet) return;

  const MAX_ROWS = 50;
  const preview  = sheet.data.slice(0, MAX_ROWS + 1); // +1 to include header
  const hasMore  = sheet.rows > MAX_ROWS;

  if (preview.length === 0) {
    dom.tableContainer.innerHTML = '<p style="padding:20px;color:var(--text-muted);text-align:center;">Empty sheet</p>';
    dom.previewRowCount.textContent = '0 rows';
    dom.previewColCount.textContent = '0 cols';
    return;
  }

  const [header, ...body] = preview;
  const colTypes = detectColumnTypes(body, header.length);

  // Build header — with row-number gutter
  const thCells = `<th class="row-num" title="Row">#</th>` +
    header.map(h => {
      const txt = String(h ?? '');
      return `<th title="${esc(txt)}">${esc(txt || '(empty)')}</th>`;
    }).join('');

  // Build body rows
  const trRows = body.map((row, ri) => {
    const rowNum = ri + 2; // 1-based, accounting for header being row 1
    const cells  = row.map((cell, ci) => {
      const txt   = String(cell ?? '');
      const empty = txt === '';
      const long  = txt.length > TRUNCATE_THRESHOLD;

      // Determine CSS classes
      const classes = [];
      if (colTypes[ci] === 'num')  classes.push('num');
      if (colTypes[ci] === 'date') classes.push('date');
      if (colTypes[ci] === 'time') classes.push('time');
      if (empty)                   classes.push('empty');
      if (long)                    classes.push('truncated');

      const cls   = classes.length ? ` class="${classes.join(' ')}"` : '';
      const titleAttr = txt ? ` title="${esc(txt)}"` : '';

      const display = empty ? '<em>—</em>' : esc(txt);
      return `<td${cls}${titleAttr}>${display}</td>`;
    }).join('');

    return `<tr><td class="row-num">${rowNum}</td>${cells}</tr>`;
  }).join('');

  const colSpan = header.length + 1; // +1 for row-num column
  const moreRow = hasMore
    ? `<tr><td colspan="${colSpan}" style="text-align:center;color:var(--text-muted);padding:12px;font-size:0.8rem;">… and ${(sheet.rows - MAX_ROWS).toLocaleString()} more rows not shown in preview</td></tr>`
    : '';

  // Apply current col width + wrap mode classes
  const wrapCls  = tableState.wrapMode ? ' wrap-mode' : '';
  const widthCls = ' ' + tableState.colWidth;

  dom.tableContainer.className = `table-wrapper${wrapCls}${widthCls}`;
  dom.tableContainer.innerHTML = `
    <table class="data-table" role="grid">
      <thead><tr>${thCells}</tr></thead>
      <tbody>${trRows}${moreRow}</tbody>
    </table>`;

  dom.previewRowCount.textContent = `${sheet.rows.toLocaleString()} rows`;
  dom.previewColCount.textContent = `${sheet.cols} cols`;
}

/** Detect column data types from sample rows */
function detectColumnTypes(rows, numCols) {
  const types  = Array(numCols).fill('text');
  const sample = rows.slice(0, 15); // sample first 15 data rows

  for (let c = 0; c < numCols; c++) {
    let nums = 0, dates = 0, times = 0, total = 0;
    for (const row of sample) {
      const v = String(row[c] ?? '').trim();
      if (!v) continue;
      total++;
      if (!isNaN(Number(v)) && v !== '') nums++;
      if (/\d{1,4}[\/\-]\d{1,2}[\/\-]\d{2,4}/.test(v)) dates++;
    }
    if (total > 0) {
      if (dates / total > 0.5)      types[c] = 'date';
      else if (nums / total > 0.7)  types[c] = 'num';
    }
  }
  return types;
}

function setActiveSheet(index) {
  state.activeSheet = index;
  document.querySelectorAll('.sheet-tab').forEach((btn, i) =>
    btn.classList.toggle('active', i === index)
  );
  renderTable();
}
window.setActiveSheet = setActiveSheet;

// ─── Export Panel ─────────────────────────────────────────

function renderExportPanel() {
  const wb = state.workbook;
  if (!wb) return;

  dom.exportPanel.classList.add('visible');

  // Render per-sheet CSV list
  dom.sheetExportList.innerHTML = wb.sheets.map((sheet, i) => `
    <div class="sheet-export-row">
      <span style="font-size:1.2rem">📋</span>
      <div class="sheet-export-name">${esc(sheet.name)}</div>
      <span class="sheet-export-meta">${sheet.rows.toLocaleString()} rows</span>
      <button class="btn-dl" onclick="downloadSingleCSV(${i})">⬇ CSV</button>
    </div>
  `).join('');
}

// ─── Export Handlers ──────────────────────────────────────

function handleExportXLSX() {
  if (!state.workbook) return;
  try {
    ExportEngine.downloadXLSX(state.workbook);
    showToast('✅ XLSX downloaded!', 'success');
  } catch (e) {
    showToast('❌ Export failed: ' + e.message, 'error');
  }
}

function handleExportAllCSV() {
  if (!state.workbook) return;
  ExportEngine.downloadAllCSV(state.workbook);
  showToast(`✅ Downloading ${state.workbook.sheets.length} CSV file(s)…`, 'success');
}

function handleExportHTML() {
  if (!state.workbook) return;
  try {
    ExportEngine.downloadHTML(state.workbook);
    showToast('✅ HTML viewer downloaded!', 'success');
  } catch (e) {
    showToast('❌ Export failed: ' + e.message, 'error');
  }
}

function handleExportJSON() {
  if (!state.workbook) return;
  try {
    ExportEngine.downloadJSON(state.workbook);
    showToast('✅ JSON downloaded!', 'success');
  } catch (e) {
    showToast('❌ Export failed: ' + e.message, 'error');
  }
}

async function handleExportPDF() {
  if (!state.workbook) return;
  if (state.exporting) return;
  state.exporting = true;

  const btn = dom.btnExportPDF;
  if (btn) { btn.disabled = true; btn.style.opacity = '0.6'; }

  dom.progressWrap.classList.add('visible');
  setProgress(0, 'Building PDF…');

  try {
    await ExportEngine.downloadPDF(state.workbook, (pct) => {
      setProgress(pct, pct < 15  ? 'Creating cover page…'
                     : pct < 95  ? `Rendering sheets… ${pct}%`
                     :             'Saving PDF…');
    });
    setProgress(100, 'PDF ready!');
    showToast('✅ PDF downloaded!', 'success');
  } catch (e) {
    showToast('❌ PDF export failed: ' + e.message, 'error');
  } finally {
    state.exporting = false;
    if (btn) { btn.disabled = false; btn.style.opacity = ''; }
    setTimeout(() => dom.progressWrap.classList.remove('visible'), 800);
  }
}

function downloadSingleCSV(sheetIndex) {
  if (!state.workbook) return;
  const sheet = state.workbook.sheets[sheetIndex];
  const base  = ExcelParser.cleanTitle(state.workbook.fileName);
  try {
    ExportEngine.downloadCSV(sheet, base);
    showToast(`✅ "${sheet.name}" CSV downloaded!`, 'success');
  } catch (e) {
    showToast('❌ Export failed: ' + e.message, 'error');
  }
}
window.downloadSingleCSV = downloadSingleCSV;

// ─── Progress ─────────────────────────────────────────────

function setProgress(pct, label) {
  dom.progressFill.style.width     = pct + '%';
  dom.progressStatus.textContent   = label || '';
  dom.progressPct.textContent      = Math.round(pct) + '%';
}

// ─── Toast ────────────────────────────────────────────────

function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  const toast     = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.textContent = message;
  container.appendChild(toast);
  setTimeout(() => {
    toast.classList.add('out');
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

// ─── Helpers ─────────────────────────────────────────────

function esc(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ─── Start ────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', init);
