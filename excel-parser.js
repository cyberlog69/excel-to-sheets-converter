/**
 * excel-parser.js — SheetJS wrapper
 * Parses an Excel .xlsx or .xls File object into an array of sheet objects.
 *
 * Requires SheetJS (XLSX) to be loaded via CDN before this script.
 */

/**
 * Parse an Excel File and return structured data.
 *
 * @param {File} file  The Excel file to parse
 * @returns {Promise<ParsedWorkbook>}
 *
 * @typedef {Object} ParsedSheet
 * @property {string}   name     Sheet/tab name
 * @property {string[][]} data   2D array of cell values (row-major, header first)
 * @property {number}   rows     Number of data rows (including header)
 * @property {number}   cols     Number of columns
 *
 * @typedef {Object} ParsedWorkbook
 * @property {string}        fileName  Original file name
 * @property {number}        fileSize  File size in bytes
 * @property {ParsedSheet[]} sheets    Array of parsed sheets
 */
async function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data     = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {
          type:      'array',
          cellDates: true,   // Parse date/time serials as JS Date objects
          cellNF:    true,   // Keep number format strings (needed for time detection)
          cellText:  false,
        });

        const sheets = workbook.SheetNames.map(name => {
          const ws = workbook.Sheets[name];

          // Use raw:true so we get actual JS Date objects and numbers
          const raw = XLSX.utils.sheet_to_json(ws, {
            header:    1,
            defval:    '',
            blankrows: false,
            raw:       true,   // Get real values (Date objects, numbers, strings)
          });

          // Build a map of cell address → format string for time detection
          const fmtMap = buildFormatMap(ws);

          // Normalize: ensure all rows are the same length
          const maxCols = raw.reduce((m, r) => Math.max(m, r.length), 0);

          const normalized = raw.map((row, rowIdx) => {
            const padded = [...row];
            while (padded.length < maxCols) padded.push('');

            return padded.map((cell, colIdx) => {
              if (cell === null || cell === undefined || cell === '') return '';

              // ── Handle JS Date objects (from cellDates:true) ──────────────
              if (cell instanceof Date) {
                if (isNaN(cell.getTime())) return '';
                return formatDateTimeCell(cell, fmtMap, rowIdx, colIdx);
              }

              // ── Handle numbers that Excel stored as time/date fractions ───
              // (happens when cellDates:true doesn't convert — e.g. some formats)
              if (typeof cell === 'number') {
                const fmt = fmtMap[`${rowIdx}:${colIdx}`] || '';
                if (isTimeFmt(fmt)) {
                  const asDate = XLSX.SSF.parse_date_code(cell);
                  if (asDate) {
                    const d = new Date(
                      asDate.y, asDate.m - 1, asDate.d,
                      asDate.H, asDate.M, Math.floor(asDate.S)
                    );
                    return formatDateTimeCell(d, fmtMap, rowIdx, colIdx);
                  }
                }
                // Regular number — preserve as-is
                return String(cell);
              }

              return String(cell);
            });
          });

          return {
            name,
            data:  normalized,
            rows:  normalized.length,
            cols:  maxCols,
          };
        }).filter(s => s.rows > 0);

        if (sheets.length === 0) {
          reject(new Error('The Excel file appears to be empty or has no readable sheets.'));
          return;
        }

        resolve({ fileName: file.name, fileSize: file.size, sheets });

      } catch (err) {
        reject(new Error('Failed to parse Excel file: ' + (err.message || err)));
      }
    };

    reader.onerror = () => reject(new Error('Failed to read the file. Please try again.'));
    reader.readAsArrayBuffer(file);
  });
}

// ─── Date / Time Formatting Helpers ──────────────────────

/**
 * Build a map of "rowIndex:colIndex" → Excel format string
 * by scanning all cells in the worksheet.
 * Used to detect time-only vs date-only vs datetime cells.
 */
function buildFormatMap(ws) {
  const map = {};
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');

  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = ws[addr];
      if (cell && cell.z) {
        map[`${R}:${C}`] = cell.z; // z = number format string
      }
    }
  }
  return map;
}

/**
 * Return true if the Excel format string represents a time format
 * (contains h, H, AM/PM markers, or looks like HH:MM).
 */
function isTimeFmt(fmt) {
  if (!fmt) return false;
  const f = fmt.toLowerCase();
  return /[h]\s*:/.test(f) || /am\/pm/i.test(fmt) || /a\/p/i.test(fmt);
}

/**
 * Return true if the Excel format string is date-only (no time component).
 */
function isDateOnlyFmt(fmt) {
  if (!fmt) return false;
  const f = fmt.toLowerCase();
  return /[ymd]/.test(f) && !/[h]/.test(f) && !/am\/pm/i.test(fmt);
}

/**
 * Return true if the format string has a date component.
 */
function hasDateFmt(fmt) {
  if (!fmt) return false;
  return /[yYdDmM]/.test(fmt);
}

/**
 * Format a JS Date object into a human-readable string,
 * correctly handling date-only, time-only, and datetime cells.
 *
 * @param {Date}   date
 * @param {Object} fmtMap   map of "row:col" → format string
 * @param {number} row      0-based row index
 * @param {number} col      0-based col index
 * @returns {string}
 */
function formatDateTimeCell(date, fmtMap, row, col) {
  const fmt = fmtMap[`${row}:${col}`] || '';

  const hasTime = date.getHours() !== 0 || date.getMinutes() !== 0 || date.getSeconds() !== 0;

  // Excel stores time-only as a date with year=1899 or year=1900, day 0 or 1
  const isTimeOnly =
    (date.getFullYear() === 1899 || date.getFullYear() === 1900) &&
    date.getMonth() === 0 &&
    date.getDate() <= 1;

  if (isTimeOnly || isTimeFmt(fmt)) {
    // Time-only cell
    return formatTime(date);
  }

  if (isDateOnlyFmt(fmt) || (!hasTime && hasDateFmt(fmt))) {
    // Date-only cell
    return formatDate(date);
  }

  if (hasTime) {
    // DateTime cell — show both
    return formatDate(date) + ' ' + formatTime(date);
  }

  // Fallback: date only
  return formatDate(date);
}

/**
 * Format a Date as DD/MM/YYYY.
 */
function formatDate(date) {
  const d  = String(date.getDate()).padStart(2, '0');
  const m  = String(date.getMonth() + 1).padStart(2, '0');
  const y  = date.getFullYear();
  return `${d}/${m}/${y}`;
}

/**
 * Format a Date as 12-hour time with AM/PM.
 * e.g. 9:05 AM, 2:30 PM, 11:59:45 PM
 */
function formatTime(date) {
  let   hours   = date.getHours();
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = date.getSeconds();
  const ampm    = hours >= 12 ? 'PM' : 'AM';

  hours = hours % 12;
  if (hours === 0) hours = 12; // midnight/noon → 12

  const hStr = String(hours); // no leading zero for hours (standard 12-hr)
  const base = `${hStr}:${minutes} ${ampm}`;

  // Include seconds only if non-zero
  if (seconds > 0) {
    const sStr = String(seconds).padStart(2, '0');
    return `${hStr}:${minutes}:${sStr} ${ampm}`;
  }
  return base;
}

// ─── Utilities ────────────────────────────────────────────

/**
 * Format file size bytes to human-readable string.
 */
function formatFileSize(bytes) {
  if (bytes < 1024)           return `${bytes} B`;
  if (bytes < 1024 * 1024)    return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/**
 * Get a clean spreadsheet title from a filename.
 */
function cleanTitle(filename) {
  return filename
    .replace(/\.(xlsx?|csv|ods)$/i, '')
    .replace(/[^\w\s\-().]/g, '')
    .trim() || 'Imported Spreadsheet';
}

window.ExcelParser = { parseExcel, formatFileSize, cleanTitle };
