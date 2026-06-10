/**
 * export-engine.js — Offline Export Engine
 * Converts parsed workbook data to various downloadable formats:
 *   - CSV (per sheet)
 *   - XLSX (re-exported multi-sheet)
 *   - HTML (styled, viewable in browser)
 *   - JSON
 *   - PDF (jsPDF + AutoTable, fully offline)
 *
 * Requires SheetJS (XLSX) and jsPDF loaded before this script.
 * No network access needed — 100% offline.
 */

// ─── CSV Export ───────────────────────────────────────────

/**
 * Download a single sheet as a CSV file.
 * @param {import('./excel-parser').ParsedSheet} sheet
 * @param {string} fileBaseName  base filename (no extension)
 */
function downloadCSV(sheet, fileBaseName) {
  const csvContent = sheet.data.map(row =>
    row.map(cell => {
      const str = String(cell ?? '');
      // Quote cells that contain commas, quotes, or newlines
      if (str.includes(',') || str.includes('"') || str.includes('\n')) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    }).join(',')
  ).join('\r\n');

  const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
  const safe = sanitizeFileName(fileBaseName + '_' + sheet.name);
  triggerDownload(blob, safe + '.csv');
}

/**
 * Download all sheets as individual CSV files (one by one).
 * @param {import('./excel-parser').ParsedWorkbook} workbook
 */
function downloadAllCSV(workbook) {
  const base = ExcelParser.cleanTitle(workbook.fileName);
  workbook.sheets.forEach((sheet, i) => {
    setTimeout(() => downloadCSV(sheet, base), i * 200);
  });
}

// ─── XLSX Re-export ───────────────────────────────────────

/**
 * Re-export the workbook as a fresh XLSX file.
 * Preserves all sheets and data. Can be opened directly in Google Sheets.
 * @param {import('./excel-parser').ParsedWorkbook} workbook
 */
function downloadXLSX(workbook) {
  const wb = XLSX.utils.book_new();

  workbook.sheets.forEach(sheet => {
    const ws = XLSX.utils.aoa_to_sheet(sheet.data);

    // Apply bold to first row
    if (sheet.data.length > 0) {
      const numCols = sheet.data[0].length;
      for (let c = 0; c < numCols; c++) {
        const cellAddr = XLSX.utils.encode_cell({ r: 0, c });
        if (ws[cellAddr]) {
          ws[cellAddr].s = {
            font: { bold: true },
            fill: { patternType: 'solid', fgColor: { rgb: '1E3A5F' } },
          };
        }
      }
    }

    XLSX.utils.book_append_sheet(wb, ws, sheet.name.slice(0, 31)); // Excel 31-char limit
  });

  const safe = sanitizeFileName(ExcelParser.cleanTitle(workbook.fileName));
  XLSX.writeFile(wb, safe + '_converted.xlsx');
}

// ─── HTML Export ──────────────────────────────────────────

/**
 * Export all sheets as a self-contained styled HTML file.
 * Opens in any browser, looks professional.
 * @param {import('./excel-parser').ParsedWorkbook} workbook
 */
function downloadHTML(workbook) {
  const title = ExcelParser.cleanTitle(workbook.fileName);

  const tabsHtml = workbook.sheets.map((s, i) =>
    `<button onclick="showSheet(${i})" id="tab-${i}" class="tab${i===0?' active':''}">${escHtml(s.name)}</button>`
  ).join('');

  const sheetsHtml = workbook.sheets.map((sheet, i) => {
    if (sheet.data.length === 0) return `<div id="sheet-${i}" class="sheet-content" style="display:${i===0?'block':'none'}"><p class="empty">Empty sheet</p></div>`;

    const [header, ...rows] = sheet.data;
    const thead = '<tr>' + header.map(h => `<th>${escHtml(h)}</th>`).join('') + '</tr>';
    const tbody = rows.map(row =>
      '<tr>' + row.map(cell => `<td>${escHtml(cell)}</td>`).join('') + '</tr>'
    ).join('');

    return `
      <div id="sheet-${i}" class="sheet-content" style="display:${i===0?'block':'none'}">
        <p class="meta">${sheet.rows} rows &times; ${sheet.cols} columns</p>
        <div class="table-wrap">
          <table><thead>${thead}</thead><tbody>${tbody}</tbody></table>
        </div>
      </div>`;
  }).join('');

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>${escHtml(title)}</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;background:#0f172a;color:#e2e8f0;padding:32px 24px}
h1{font-size:1.6rem;font-weight:800;margin-bottom:6px;color:#f8fafc}
.meta-bar{font-size:0.8rem;color:#64748b;margin-bottom:24px}
.tabs{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:16px}
.tab{padding:7px 16px;border-radius:20px;border:1px solid #334155;background:#1e293b;color:#94a3b8;font-size:0.82rem;font-weight:600;cursor:pointer;transition:all .2s}
.tab.active,.tab:hover{border-color:#60a5fa;background:rgba(96,165,250,.12);color:#60a5fa}
.sheet-content{}
.meta{font-size:0.78rem;color:#64748b;margin-bottom:10px}
.table-wrap{overflow-x:auto;border-radius:10px;border:1px solid #1e293b}
table{width:100%;border-collapse:collapse;font-size:0.82rem}
th{padding:10px 14px;text-align:left;font-weight:700;font-size:0.72rem;letter-spacing:.06em;text-transform:uppercase;color:#64748b;background:#1e293b;border-bottom:1px solid #334155;white-space:nowrap;position:sticky;top:0}
td{padding:8px 14px;color:#94a3b8;border-bottom:1px solid #0f172a;white-space:nowrap;max-width:240px;overflow:hidden;text-overflow:ellipsis}
tr:hover td{background:#1e293b;color:#e2e8f0}
.empty{color:#64748b;padding:20px;text-align:center}
footer{margin-top:40px;font-size:0.75rem;color:#334155;text-align:center}
</style>
</head>
<body>
<h1>📊 ${escHtml(title)}</h1>
<div class="meta-bar">Exported ${workbook.sheets.length} sheet(s) &bull; ${new Date().toLocaleString()}</div>
<div class="tabs">${tabsHtml}</div>
${sheetsHtml}
<footer>Generated by Excel → Google Sheets Converter (Offline Edition)</footer>
<script>
function showSheet(i){
  document.querySelectorAll('.sheet-content').forEach((el,j)=>el.style.display=j===i?'block':'none');
  document.querySelectorAll('.tab').forEach((t,j)=>t.classList.toggle('active',j===i));
}
<\/script>
</body>
</html>`;

  const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
  const safe  = sanitizeFileName(title);
  triggerDownload(blob, safe + '_viewer.html');
}

// ─── JSON Export ──────────────────────────────────────────

/**
 * Export all sheets as JSON (array of {sheetName, headers, rows}).
 * @param {import('./excel-parser').ParsedWorkbook} workbook
 */
function downloadJSON(workbook) {
  const output = workbook.sheets.map(sheet => {
    const [headers, ...rows] = sheet.data;
    return {
      sheetName: sheet.name,
      headers:   headers || [],
      rows:      rows.map(row => {
        const obj = {};
        (headers || []).forEach((h, i) => { obj[h || `col${i+1}`] = row[i] ?? ''; });
        return obj;
      }),
      totalRows: sheet.rows,
      totalCols: sheet.cols,
    };
  });

  const json = JSON.stringify(output, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const safe = sanitizeFileName(ExcelParser.cleanTitle(workbook.fileName));
  triggerDownload(blob, safe + '.json');
}

// ─── PDF Export ─────────────────────────────────────────

/** Colour palette for the PDF */
const PDF_COLORS = {
  headerBg:    [30,  58,  95],   // dark navy  — header row fill
  headerText:  [255, 255, 255],  // white      — header text
  oddRow:      [245, 248, 252],  // near-white — odd data rows
  evenRow:     [255, 255, 255],  // white      — even data rows
  borderColor: [200, 210, 225],  // light grey — cell borders
  titleText:   [15,  23,  42],   // near-black — document title
  sheetTitle:  [30,  58,  95],   // navy       — sheet section title
  metaText:    [100, 116, 139],  // slate      — metadata / footers
  accentBar:   [74,  222, 128],  // green      — cover accent line
};

/**
 * Download all sheets as a single, beautifully formatted PDF.
 * Features:
 *  - Cover page with title, sheet list and export date
 *  - One section per sheet with sheet name header
 *  - Auto landscape/portrait based on column count
 *  - Frozen header row repeated on every page
 *  - Alternating row colours
 *  - Page numbers ("Sheet X of Y — Page P of N")
 *  - Long cell text wraps within cells (no data lost)
 *
 * @param {import('./excel-parser').ParsedWorkbook} workbook
 * @param {function} [onProgress]  progress callback 0–100
 */
async function downloadPDF(workbook, onProgress) {
  const progress = (p) => onProgress?.(p);

  // jsPDF is exposed as window.jspdf.jsPDF in UMD build
  const { jsPDF } = window.jspdf;
  if (!jsPDF) throw new Error('jsPDF library not loaded.');

  progress(5);

  const title    = ExcelParser.cleanTitle(workbook.fileName);
  const sheets   = workbook.sheets;
  const now      = new Date();
  const dateStr  = now.toLocaleDateString('en-GB', { day:'2-digit', month:'short', year:'numeric' });
  const timeStr  = now.toLocaleTimeString('en-GB', { hour:'2-digit', minute:'2-digit' });

  // ── Determine page orientation for each sheet ────────────
  // Use landscape if the sheet has more than 7 columns
  const LANDSCAPE_THRESHOLD = 7;

  // We'll create one PDF doc. We detect orientation per sheet and
  // add pages accordingly. Start with the first sheet's orientation.
  const firstOrientation = (sheets[0]?.cols ?? 0) > LANDSCAPE_THRESHOLD ? 'landscape' : 'portrait';

  const doc = new jsPDF({
    orientation: firstOrientation,
    unit: 'mm',
    format: 'a4',
    compress: true,
  });

  // ── Helper: page dimensions ───────────────────────────────
  const getPageSize = (orientation) => orientation === 'landscape'
    ? { w: 297, h: 210 }
    : { w: 210, h: 297 };

  // ── Cover Page ────────────────────────────────────────────
  const coverSize = getPageSize(firstOrientation);
  const cx = coverSize.w;
  const cy = coverSize.h;

  // Green accent bar at top
  doc.setFillColor(...PDF_COLORS.accentBar);
  doc.rect(0, 0, cx, 8, 'F');

  // Spreadsheet emoji area (coloured rectangle)
  doc.setFillColor(...PDF_COLORS.headerBg);
  doc.roundedRect(cx / 2 - 16, 32, 32, 32, 4, 4, 'F');
  doc.setFontSize(22);
  doc.setTextColor(255, 255, 255);
  doc.text('XLS', cx / 2, 52, { align: 'center' });

  // Title
  doc.setFontSize(22);
  doc.setTextColor(...PDF_COLORS.titleText);
  doc.setFont('helvetica', 'bold');
  const titleLines = doc.splitTextToSize(title, cx - 40);
  doc.text(titleLines, cx / 2, 80, { align: 'center' });

  // Subtitle bar
  doc.setFillColor(...PDF_COLORS.headerBg);
  doc.rect(20, 95, cx - 40, 0.5, 'F');

  // Meta info
  doc.setFontSize(10);
  doc.setFont('helvetica', 'normal');
  doc.setTextColor(...PDF_COLORS.metaText);
  doc.text(`Exported on ${dateStr} at ${timeStr}`, cx / 2, 106, { align: 'center' });
  doc.text(`Source file: ${workbook.fileName}`, cx / 2, 113, { align: 'center' });
  doc.text(`${sheets.length} sheet${sheets.length !== 1 ? 's' : ''} — ${sheets.reduce((s, sh) => s + sh.rows, 0).toLocaleString()} total rows`, cx / 2, 120, { align: 'center' });

  // Sheet list
  if (sheets.length > 0) {
    doc.setFontSize(9);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(...PDF_COLORS.sheetTitle);
    doc.text('CONTENTS', cx / 2, 138, { align: 'center' });

    doc.setFont('helvetica', 'normal');
    doc.setTextColor(...PDF_COLORS.metaText);
    sheets.forEach((sh, i) => {
      const y = 146 + i * 8;
      if (y > cy - 20) return; // don't overflow cover page
      doc.text(`${i + 1}.  ${sh.name}`, cx / 2 - 40, y);
      doc.text(`${sh.rows.toLocaleString()} rows × ${sh.cols} cols`, cx - 25, y, { align: 'right' });
    });
  }

  // Footer on cover
  doc.setFontSize(8);
  doc.setTextColor(...PDF_COLORS.metaText);
  doc.text('Generated by Excel → Google Sheets Converter (Offline Edition)', cx / 2, cy - 10, { align: 'center' });

  progress(15);

  // ── Sheet Pages ───────────────────────────────────────────
  for (let si = 0; si < sheets.length; si++) {
    const sheet = sheets[si];
    const orientation = sheet.cols > LANDSCAPE_THRESHOLD ? 'landscape' : 'portrait';
    const pageSize    = getPageSize(orientation);
    const pw = pageSize.w;
    const ph = pageSize.h;

    // Add a new page (with correct orientation) for each sheet
    doc.addPage([pw, ph], orientation);

    // ── Sheet title banner ───────────────────────────────────
    doc.setFillColor(...PDF_COLORS.headerBg);
    doc.rect(0, 0, pw, 14, 'F');

    doc.setFontSize(11);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(255, 255, 255);
    doc.text(`✐ ${sheet.name}`, 14, 9.5);

    // Sheet metadata (right-aligned)
    doc.setFontSize(8);
    doc.setFont('helvetica', 'normal');
    doc.text(`${sheet.rows.toLocaleString()} rows × ${sheet.cols} columns`, pw - 14, 9.5, { align: 'right' });

    if (sheet.data.length === 0) {
      doc.setFontSize(10);
      doc.setTextColor(...PDF_COLORS.metaText);
      doc.text('(This sheet is empty)', pw / 2, 40, { align: 'center' });
      continue;
    }

    const [header, ...rows] = sheet.data;

    // ── Build AutoTable column definitions ───────────────────
    // Calculate optimal column widths based on content
    const usableWidth = pw - 28; // 14mm margin each side
    const numCols     = header.length;

    // Sample content widths (chars) to determine proportional col widths
    const sampleRows  = rows.slice(0, 30);
    const charWidths  = header.map((h, ci) => {
      const headerLen  = String(h ?? '').length;
      const maxDataLen = sampleRows.reduce((mx, row) =>
        Math.max(mx, String(row[ci] ?? '').length), 0);
      return Math.max(headerLen, maxDataLen, 4); // minimum 4 chars
    });

    const totalChars  = charWidths.reduce((s, w) => s + w, 0) || 1;
    const MIN_COL_MM  = 12;
    const MAX_COL_MM  = 70;

    const colWidths = charWidths.map(cw => {
      const proportional = (cw / totalChars) * usableWidth;
      return Math.min(MAX_COL_MM, Math.max(MIN_COL_MM, proportional));
    });

    // Scale back if total exceeds usable width
    const totalW = colWidths.reduce((s, w) => s + w, 0);
    const scale  = totalW > usableWidth ? usableWidth / totalW : 1;
    const finalWidths = colWidths.map(w => +(w * scale).toFixed(2));

    // ── AutoTable ────────────────────────────────────────────
    doc.autoTable({
      startY:      18,
      margin:      { left: 14, right: 14, bottom: 16 },

      head:        [header.map(h => String(h ?? ''))],
      body:        rows.map(row => row.map(cell => String(cell ?? ''))),

      columnStyles: Object.fromEntries(
        finalWidths.map((w, i) => [i, { cellWidth: w }])
      ),

      // Header row styling
      headStyles: {
        fillColor:       PDF_COLORS.headerBg,
        textColor:       PDF_COLORS.headerText,
        fontStyle:       'bold',
        fontSize:        8,
        cellPadding:     { top: 3, right: 4, bottom: 3, left: 4 },
        lineColor:       PDF_COLORS.borderColor,
        lineWidth:       0.1,
        halign:          'left',
        valign:          'middle',
        overflow:        'linebreak',
      },

      // Data cell styling
      bodyStyles: {
        fontSize:    7.5,
        textColor:   [30, 41, 59],
        cellPadding: { top: 2.5, right: 4, bottom: 2.5, left: 4 },
        lineColor:   PDF_COLORS.borderColor,
        lineWidth:   0.1,
        overflow:    'linebreak',    // wrap text — NO data lost
        valign:      'top',
      },

      // Alternating row colours
      alternateRowStyles: {
        fillColor: PDF_COLORS.oddRow,
      },

      // Repeat header on every page
      showHead: 'everyPage',

      // Page break hook: add sheet banner on continuation pages
      didDrawPage: (hookData) => {
        const pageNum  = doc.internal.getNumberOfPages();
        const sheetNum = si + 1;

        // Footer: page info + sheet name
        doc.setFontSize(7.5);
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(...PDF_COLORS.metaText);
        doc.text(
          `${title}  —  Sheet ${sheetNum} of ${sheets.length}: ${sheet.name}`,
          14, ph - 6
        );
        doc.text(
          `Page ${pageNum}`,
          pw - 14, ph - 6,
          { align: 'right' }
        );

        // Thin bottom rule
        doc.setDrawColor(...PDF_COLORS.borderColor);
        doc.setLineWidth(0.3);
        doc.line(14, ph - 9, pw - 14, ph - 9);
      },
    });

    progress(15 + Math.round((si + 1) / sheets.length * 80));
  }

  progress(98);

  // ── Save ──────────────────────────────────────────────────
  const safe = sanitizeFileName(title);
  doc.save(safe + '.pdf');

  progress(100);
}

// ─── Helpers ─────────────────────────────────────────────

function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a   = document.createElement('a');
  a.href     = url;
  a.download = filename;
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);
}

function sanitizeFileName(name) {
  return name.replace(/[<>:"/\\|?*]/g, '_').replace(/\s+/g, '_').slice(0, 100);
}

function escHtml(str) {
  return String(str ?? '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

window.ExportEngine = {
  downloadCSV,
  downloadAllCSV,
  downloadXLSX,
  downloadHTML,
  downloadJSON,
  downloadPDF,
};
