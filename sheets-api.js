/**
 * sheets-api.js — Google Sheets REST API Client
 * Handles creating spreadsheets and writing data via Google Sheets v4 API.
 */

const SHEETS_BASE = 'https://sheets.googleapis.com/v4/spreadsheets';
const DRIVE_BASE  = 'https://www.googleapis.com/drive/v3/files';

// ─── Helpers ─────────────────────────────────────────────

function authHeaders(token) {
  return {
    'Authorization': `Bearer ${token}`,
    'Content-Type':  'application/json',
  };
}

async function apiRequest(url, options) {
  const res = await fetch(url, options);
  if (!res.ok) {
    let msg = `HTTP ${res.status}`;
    try {
      const body = await res.json();
      msg = body?.error?.message || msg;
    } catch (_) {}
    throw new Error(msg);
  }
  return res.json();
}

/**
 * Extract spreadsheet ID from a Google Sheets URL or plain ID.
 * @param {string} urlOrId
 * @returns {string|null}
 */
function extractSheetId(urlOrId) {
  if (!urlOrId) return null;
  const match = urlOrId.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  // Might be a plain ID already
  if (/^[a-zA-Z0-9_-]{20,}$/.test(urlOrId.trim())) return urlOrId.trim();
  return null;
}

// ─── Core API Functions ───────────────────────────────────

/**
 * Create a new Google Spreadsheet with the given title and sheet names.
 * Returns the full spreadsheet resource.
 *
 * @param {string}   token       OAuth2 access token
 * @param {string}   title       Spreadsheet title
 * @param {string[]} sheetNames  Names for each sheet/tab
 * @returns {Promise<{id:string, url:string}>}
 */
async function createSpreadsheet(token, title, sheetNames) {
  const sheets = sheetNames.map((name, index) => ({
    properties: {
      title:     name,
      index,
      gridProperties: { rowCount: 1000, columnCount: 26 },
    },
  }));

  const body = {
    properties: { title },
    sheets,
  };

  const data = await apiRequest(SHEETS_BASE, {
    method:  'POST',
    headers: authHeaders(token),
    body:    JSON.stringify(body),
  });

  return {
    id:  data.spreadsheetId,
    url: data.spreadsheetUrl,
  };
}

/**
 * Write data to a sheet by clearing it first, then writing all values.
 *
 * @param {string}     token          OAuth2 access token
 * @param {string}     spreadsheetId  Google Sheet ID
 * @param {string}     sheetName      Tab name to write to
 * @param {string[][]} data           2D array of values
 * @param {function}   [onProgress]   Optional progress callback (0-1)
 */
async function writeSheetData(token, spreadsheetId, sheetName, data, onProgress) {
  if (!data || data.length === 0) return;

  // 1. Clear existing data
  await apiRequest(
    `${SHEETS_BASE}/${spreadsheetId}/values/${encodeURIComponent(sheetName)}!A1:ZZ999999:clear`,
    { method: 'POST', headers: authHeaders(token), body: JSON.stringify({}) }
  );

  onProgress?.(0.3);

  // 2. Write data in chunks to avoid 10MB payload limit
  const CHUNK_ROWS = 1000;
  const totalChunks = Math.ceil(data.length / CHUNK_ROWS);

  for (let i = 0; i < totalChunks; i++) {
    const chunk      = data.slice(i * CHUNK_ROWS, (i + 1) * CHUNK_ROWS);
    const startRow   = i * CHUNK_ROWS + 1;
    const range      = `${sheetName}!A${startRow}`;

    await apiRequest(
      `${SHEETS_BASE}/${spreadsheetId}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,
      {
        method:  'PUT',
        headers: authHeaders(token),
        body:    JSON.stringify({ range, majorDimension: 'ROWS', values: chunk }),
      }
    );

    onProgress?.(0.3 + 0.65 * ((i + 1) / totalChunks));
  }

  onProgress?.(0.95);
}

/**
 * Apply basic formatting to the header row (bold + background color).
 *
 * @param {string} token
 * @param {string} spreadsheetId
 * @param {number} sheetId        Numeric sheet ID (from spreadsheet resource)
 * @param {number} numCols        Number of columns in the header
 */
async function formatHeaderRow(token, spreadsheetId, sheetId, numCols) {
  const requests = [
    // Bold + background for row 0
    {
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: 0, endRowIndex: 1,
          startColumnIndex: 0, endColumnIndex: numCols,
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 0.17, green: 0.24, blue: 0.31 },
            textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    },
    // Freeze first row
    {
      updateSheetProperties: {
        properties: {
          sheetId,
          gridProperties: { frozenRowCount: 1 },
        },
        fields: 'gridProperties.frozenRowCount',
      },
    },
    // Auto-resize columns
    {
      autoResizeDimensions: {
        dimensions: {
          sheetId,
          dimension: 'COLUMNS',
          startIndex: 0,
          endIndex: numCols,
        },
      },
    },
  ];

  await apiRequest(`${SHEETS_BASE}/${spreadsheetId}:batchUpdate`, {
    method:  'POST',
    headers: authHeaders(token),
    body:    JSON.stringify({ requests }),
  });
}

/**
 * Get the spreadsheet metadata (title, sheets list).
 *
 * @param {string} token
 * @param {string} spreadsheetId
 * @returns {Promise<{title:string, sheets:Array<{id:number, name:string}>}>}
 */
async function getSpreadsheetMeta(token, spreadsheetId) {
  const data = await apiRequest(
    `${SHEETS_BASE}/${spreadsheetId}?fields=properties,sheets.properties`,
    { headers: authHeaders(token) }
  );

  return {
    title: data.properties.title,
    sheets: (data.sheets || []).map(s => ({
      id:   s.properties.sheetId,
      name: s.properties.title,
    })),
  };
}

/**
 * Ensure all required sheet tabs exist in the spreadsheet (add if missing).
 *
 * @param {string}   token
 * @param {string}   spreadsheetId
 * @param {string[]} requiredNames
 * @param {Array<{id:number, name:string}>} existingSheets
 */
async function ensureSheets(token, spreadsheetId, requiredNames, existingSheets) {
  const existingNames = new Set(existingSheets.map(s => s.name));
  const toAdd = requiredNames.filter(n => !existingNames.has(n));

  if (toAdd.length === 0) return;

  const requests = toAdd.map(title => ({
    addSheet: { properties: { title } },
  }));

  await apiRequest(`${SHEETS_BASE}/${spreadsheetId}:batchUpdate`, {
    method:  'POST',
    headers: authHeaders(token),
    body:    JSON.stringify({ requests }),
  });
}

// ─── High-Level Converter ─────────────────────────────────

/**
 * Full convert pipeline:
 *  1. Create a new spreadsheet (or use existing)
 *  2. Write each sheet's data
 *  3. Format header rows
 *
 * @param {Object}   opts
 * @param {string}   opts.token          OAuth2 access token
 * @param {string}   opts.title          Spreadsheet title
 * @param {Array}    opts.sheets         ParsedSheet[] from excel-parser
 * @param {string}   [opts.targetId]     Existing sheet ID (optional)
 * @param {function} [opts.onProgress]   Progress callback (0-100)
 * @returns {Promise<{id:string, url:string}>}
 */
async function convertToGoogleSheets(opts) {
  const { token, title, sheets, targetId, onProgress } = opts;
  const progress = (pct) => onProgress?.(pct);

  progress(5);

  let spreadsheetId, spreadsheetUrl;
  let sheetMeta; // { id, name }[]

  if (targetId) {
    // Use existing sheet — ensure tabs exist
    const meta = await getSpreadsheetMeta(token, targetId);
    spreadsheetId  = targetId;
    spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${targetId}/edit`;
    sheetMeta      = meta.sheets;
    await ensureSheets(token, targetId, sheets.map(s => s.name), sheetMeta);
    // Refresh meta after adding sheets
    const updated  = await getSpreadsheetMeta(token, targetId);
    sheetMeta      = updated.sheets;
  } else {
    // Create brand-new spreadsheet
    const created  = await createSpreadsheet(token, title, sheets.map(s => s.name));
    spreadsheetId  = created.id;
    spreadsheetUrl = created.url;
    const meta     = await getSpreadsheetMeta(token, spreadsheetId);
    sheetMeta      = meta.sheets;
  }

  progress(15);

  // Write each sheet
  for (let i = 0; i < sheets.length; i++) {
    const sheet     = sheets[i];
    const baseStart = 15 + (i / sheets.length) * 75;
    const baseEnd   = 15 + ((i + 1) / sheets.length) * 75;

    await writeSheetData(
      token,
      spreadsheetId,
      sheet.name,
      sheet.data,
      (p) => progress(baseStart + p * (baseEnd - baseStart))
    );

    // Apply header formatting
    const sheetInfo = sheetMeta.find(s => s.name === sheet.name);
    if (sheetInfo && sheet.data.length > 0) {
      try {
        await formatHeaderRow(token, spreadsheetId, sheetInfo.id, sheet.cols);
      } catch (_) {
        // Non-critical, skip if formatting fails
      }
    }
  }

  progress(100);

  return { id: spreadsheetId, url: spreadsheetUrl };
}

window.SheetsAPI = {
  createSpreadsheet,
  writeSheetData,
  formatHeaderRow,
  getSpreadsheetMeta,
  ensureSheets,
  convertToGoogleSheets,
  extractSheetId,
};
