/**************************************************************
 * Shared Utilities (one file)
 * Project: Ping "Sauron" Sheet Pipeline
 *
 * Drop this in a single Apps Script file, e.g. "00_utils.gs"
 *
 * Notes:
 * - Header-based: readHeaderMap() builds {headerLower: colIndex1Based}
 * - Email key: normalizeEmail() -> lower(trim(email))
 * - Batch writes: batchSetValues() writes in chunks to avoid limits
 * - Upserts: buildIndexByKey() builds {key: rowIndex0BasedInArray}
 * - Logging: writeSyncLog() appends to a "sync_log" tab
 * - Locking: lockWrap() prevents overlapping runs
 **************************************************************/

const UTIL_CFG = {
  SYNC_LOG_SHEET: 'sync_log',
  DEFAULT_BATCH_ROWS: 5000,
  LOCK_TIMEOUT_MS: 5 * 60 * 1000, // 5 minutes
};

/**
 * Get or create a sheet by name.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet(ss, name) {
  if (!ss) throw new Error('getOrCreateSheet: ss is required');
  if (!name) throw new Error('getOrCreateSheet: name is required');
  const existing = ss.getSheetByName(name);
  return existing || ss.insertSheet(name);
}

/**
 * Read header row and return a map of header -> 1-based col index.
 * Keys are normalized to lowercase trimmed strings.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} headerRow
 * @returns {{ map: Object<string, number>, headers: string[] }}
 */
function readHeaderMap(sheet, headerRow) {
  if (!sheet) throw new Error('readHeaderMap: sheet is required');
  if (!headerRow || headerRow < 1) throw new Error('readHeaderMap: headerRow must be >= 1');

  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return { map: {}, headers: [] };

  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || '').trim());

  const map = {};
  headers.forEach((h, i) => {
    const key = String(h || '').trim().toLowerCase();
    if (!key) return;
    // If duplicate headers exist, keep the first one (predictable)
    if (map[key] == null) map[key] = i + 1;
  });

  return { map, headers };
}

/**
 * Normalize an email address into a stable key.
 * @param {string} email
 * @returns {string} email_key
 */
function normalizeEmail(email) {
  if (email == null) return '';
  return String(email).trim().toLowerCase();
}

/**
 * Batch write a 2D array to a sheet in chunks.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} startRow 1-based
 * @param {number} startCol 1-based
 * @param {any[][]} values 2D array
 * @param {number=} batchRows optional override
 */
function batchSetValues(sheet, startRow, startCol, values, batchRows) {
  if (!sheet) throw new Error('batchSetValues: sheet is required');
  if (!startRow || startRow < 1) throw new Error('batchSetValues: startRow must be >= 1');
  if (!startCol || startCol < 1) throw new Error('batchSetValues: startCol must be >= 1');
  if (!Array.isArray(values) || values.length === 0) return;

  const rowsPerBatch = batchRows || UTIL_CFG.DEFAULT_BATCH_ROWS;
  const numCols = values[0].length;

  for (let i = 0; i < values.length; i += rowsPerBatch) {
    const chunk = values.slice(i, i + rowsPerBatch);
    sheet.getRange(startRow + i, startCol, chunk.length, numCols).setValues(chunk);
  }
}

/**
 * Build an index map from a 2D array of rows keyed by a specific column index (0-based).
 *
 * @param {any[][]} rows
 * @param {number} keyColIndex0 0-based index within each row
 * @param {Object=} opts
 * @param {boolean=} opts.lowercase
 * @param {boolean=} opts.trim
 * @returns {Object<string, number>} key -> rowIndex0Based
 */
function buildIndexByKey(rows, keyColIndex0, opts) {
  if (!Array.isArray(rows)) throw new Error('buildIndexByKey: rows must be an array');
  if (keyColIndex0 == null || keyColIndex0 < 0) throw new Error('buildIndexByKey: keyColIndex0 must be >= 0');

  const o = Object.assign({ lowercase: true, trim: true }, opts || {});
  const index = {};

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || [];
    let key = r[keyColIndex0];

    if (key == null) continue;
    key = String(key);

    if (o.trim) key = key.trim();
    if (o.lowercase) key = key.toLowerCase();
    if (!key) continue;

    // Keep first occurrence (stable). Later duplicates are ignored.
    if (index[key] == null) index[key] = i;
  }

  return index;
}

/**
 * Append a row to sync_log (creates the tab + header if needed).
 *
 * Columns:
 * timestamp | step | status | rows_in | rows_out | seconds | error
 *
 * @param {string} step
 * @param {string} status e.g. "ok" | "error"
 * @param {number=} rowsIn
 * @param {number=} rowsOut
 * @param {number=} seconds
 * @param {string=} errorMsg
 */
function writeSyncLog(step, status, rowsIn, rowsOut, seconds, errorMsg) {
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet(ss, UTIL_CFG.SYNC_LOG_SHEET);

  if (sh.getLastRow() === 0) {
    sh.appendRow(['timestamp', 'step', 'status', 'rows_in', 'rows_out', 'seconds', 'error']);
    sh.setFrozenRows(1);
  }

  sh.appendRow([
    new Date(),
    step || '',
    status || '',
    rowsIn != null ? rowsIn : '',
    rowsOut != null ? rowsOut : '',
    seconds != null ? seconds : '',
    errorMsg || ''
  ]);
}

/**
 * Wrap a function call in a document lock to prevent overlapping runs.
 *
 * Usage:
 *   lockWrap('daily_pipeline', () => run_daily_pipeline_impl_());
 *
 * @param {string} lockName used only for log labeling
 * @param {Function} fn
 * @param {Object=} opts
 * @param {number=} opts.timeoutMs
 * @returns {any} return value of fn
 */
function lockWrap(lockName, fn, opts) {
  if (typeof fn !== 'function') throw new Error('lockWrap: fn must be a function');

  const timeoutMs = (opts && opts.timeoutMs) || UTIL_CFG.LOCK_TIMEOUT_MS;
  const lock = LockService.getDocumentLock();

  const t0 = new Date();
  const got = lock.tryLock(timeoutMs);

  if (!got) {
    const msg = `lockWrap: could not acquire lock (${lockName || 'job'}) within ${timeoutMs}ms`;
    writeSyncLog(lockName || 'job', 'error', '', '', (new Date() - t0) / 1000, msg);
    throw new Error(msg);
  }

  try {
    return fn();
  } catch (err) {
    const msg = String(err && err.message ? err.message : err);
    writeSyncLog(lockName || 'job', 'error', '', '', (new Date() - t0) / 1000, msg);
    throw err;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}