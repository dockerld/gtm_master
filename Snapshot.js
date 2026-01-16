/**************************************************************
 * write_daily_snapshot()
 *
 * Appends a daily snapshot of your "Sauron" view into an append-only
 * table: "snap_users_daily", with a snapshot_date column.
 *
 * Behavior:
 * - Reads Sauron headers from Row 3 starting at Col B
 * - Reads Sauron data from Row 4 starting at Col B
 * - Writes to snap_users_daily with snapshot_date as the FIRST column
 * - Dedupe: will NOT re-insert rows for the same (snapshot_date + email_key)
 * - Robust to trailing blank columns in Sauron (uses header width, not sheet last col)
 *
 * Shared utils used if present:
 * - normalizeEmail(email)
 * - batchSetValues(sheet, startRow, startCol, values, chunkSize)
 * - lockWrap(lockName, fn)   OR lockWrap(fn) depending on your implementation
 **************************************************************/

const SNAP_CFG = {
  SOURCE_SHEET: 'Sauron',
  SNAP_SHEET: 'snap_users_daily',

  // Sauron layout
  HEADER_ROW: 3,
  START_COL: 2,        // Col B
  DATA_START_ROW: 4,

  // Snapshot metadata
  SNAPSHOT_DATE_HEADER: 'snapshot_date',
  SNAPSHOT_DATE_FMT: 'yyyy-MM-dd',

  // Required key for dedupe
  EMAIL_HEADER: 'Email',

  // Batching
  WRITE_CHUNK: 3000
}

function write_daily_snapshot() {
  lockWrapCompat_('write_daily_snapshot', () => {
    const t0 = new Date()
    const ss = SpreadsheetApp.getActive()

    const src = ss.getSheetByName(SNAP_CFG.SOURCE_SHEET)
    if (!src) throw new Error(`Source sheet not found: ${SNAP_CFG.SOURCE_SHEET}`)

    const snap = getOrCreateSheetCompat_(ss, SNAP_CFG.SNAP_SHEET)

    // Compute today snapshot_date in your script timezone
    const tz = Session.getScriptTimeZone()
    const snapshotDate = Utilities.formatDate(new Date(), tz, SNAP_CFG.SNAPSHOT_DATE_FMT)

    // ---- Read Sauron headers: find contiguous header width starting at B3 ----
    const maxColsFromStart = src.getLastColumn() - SNAP_CFG.START_COL + 1
    if (maxColsFromStart <= 0) throw new Error('Sauron has no columns in the expected region')

    const rawHeaderRow = src
      .getRange(SNAP_CFG.HEADER_ROW, SNAP_CFG.START_COL, 1, maxColsFromStart)
      .getValues()[0]
      .map(h => String(h || '').trim())

    const headerWidth = contiguousHeaderWidth_(rawHeaderRow)
    if (headerWidth <= 0) throw new Error('Sauron header row appears empty starting at B3')

    const srcHeaders = rawHeaderRow.slice(0, headerWidth)

    const emailIdxInSrc = srcHeaders.findIndex(h => h.toLowerCase() === SNAP_CFG.EMAIL_HEADER.toLowerCase())
    if (emailIdxInSrc < 0) throw new Error(`Sauron is missing required header: ${SNAP_CFG.EMAIL_HEADER}`)

    // ---- Read Sauron data (use header width, not sheet last col) ----
    const srcLastRow = src.getLastRow()
    if (srcLastRow < SNAP_CFG.DATA_START_ROW) {
      Logger.log('No data rows in Sauron. Snapshot skipped.')
      return
    }

    const numRows = srcLastRow - SNAP_CFG.DATA_START_ROW + 1
    const srcData = src.getRange(SNAP_CFG.DATA_START_ROW, SNAP_CFG.START_COL, numRows, headerWidth).getValues()

    // Filter out blank email rows
    const rowsWithEmail = []
    for (const r of srcData) {
      const email = String(r[emailIdxInSrc] || '').trim()
      if (normalizeEmailCompat_(email)) rowsWithEmail.push(r)
    }

    if (!rowsWithEmail.length) {
      Logger.log('No rows with Email in Sauron. Snapshot skipped.')
      return
    }

    // ---- Ensure snapshot headers: snapshot_date + Sauron headers ----
    const snapHeaders = [SNAP_CFG.SNAPSHOT_DATE_HEADER].concat(srcHeaders)
    ensureSnapshotHeaders_(snap, snapHeaders)

    // ---- Build dedupe set for today's date: snapshot_date|email_key ----
    const existingKeys = buildExistingSnapshotKeySet_(
      snap,
      snapshotDate,
      SNAP_CFG.SNAPSHOT_DATE_HEADER,
      SNAP_CFG.EMAIL_HEADER
    )

    // ---- Prepare rows to append ----
    const out = []
    let skipped = 0

    for (const r of rowsWithEmail) {
      const email = String(r[emailIdxInSrc] || '').trim()
      const emailKey = normalizeEmailCompat_(email)
      const key = snapshotDate + '|' + emailKey

      if (existingKeys.has(key)) {
        skipped++
        continue
      }
      out.push([snapshotDate].concat(r))
    }

    if (!out.length) {
      Logger.log(`No new rows to snapshot for ${snapshotDate}. Skipped existing: ${skipped}`)
      return
    }

    // ---- Append to snapshot sheet ----
    const startRow = snap.getLastRow() + 1
    batchSetValuesCompat_(snap, startRow, 1, out, SNAP_CFG.WRITE_CHUNK)

    Logger.log(
      `Snapshot ${snapshotDate}: appended ${out.length} rows. ` +
      `Skipped existing: ${skipped}. Took ${((new Date() - t0) / 1000).toFixed(2)}s`
    )
  })
}

/* =========================
 * Internal helpers
 * ========================= */

/**
 * Returns contiguous header width until the first blank cell.
 * This avoids pulling trailing empty columns from Sauron.
 */
function contiguousHeaderWidth_(headerRowArray) {
  let w = 0
  for (let i = 0; i < headerRowArray.length; i++) {
    if (!headerRowArray[i]) break
    w++
  }
  return w
}

function ensureSnapshotHeaders_(sheet, headers) {
  const lastRow = sheet.getLastRow()

  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    sheet.setFrozenRows(1)
    return
  }

  const lastCol = Math.max(sheet.getLastColumn(), headers.length)
  const existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())

  const existingTrimmed = existing.filter(Boolean)
  if (!existingTrimmed.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    sheet.setFrozenRows(1)
    return
  }

  const same =
    existing.length >= headers.length &&
    headers.every((h, i) => String(existing[i] || '').trim() === h)

  if (!same) {
    // Strict reset to desired headers
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    sheet.setFrozenRows(1)
  }
}

function buildExistingSnapshotKeySet_(sheet, snapshotDate, snapDateHeader, emailHeader) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  const set = new Set()
  if (lastRow < 2) return set

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())
  const snapIdx = header.findIndex(h => h.toLowerCase() === snapDateHeader.toLowerCase())
  const emailIdx = header.findIndex(h => h.toLowerCase() === emailHeader.toLowerCase())

  if (snapIdx < 0) throw new Error(`Snapshot sheet missing header: ${snapDateHeader}`)
  if (emailIdx < 0) throw new Error(`Snapshot sheet missing header: ${emailHeader}`)

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()

  for (const r of data) {
    const d = String(r[snapIdx] || '').trim()
    if (d !== snapshotDate) continue

    const email = String(r[emailIdx] || '').trim()
    const emailKey = normalizeEmailCompat_(email)
    if (!emailKey) continue

    set.add(snapshotDate + '|' + emailKey)
  }

  return set
}

function buildExistingSnapshotKeySetGeneric_(sheet, snapshotDate, snapDateHeader, keyHeader, normalizeFn) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  const set = new Set()
  if (lastRow < 2) return set

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())
  const snapIdx = header.findIndex(h => h.toLowerCase() === snapDateHeader.toLowerCase())
  const keyIdx = header.findIndex(h => h.toLowerCase() === keyHeader.toLowerCase())

  if (snapIdx < 0) throw new Error(`Snapshot sheet missing header: ${snapDateHeader}`)
  if (keyIdx < 0) throw new Error(`Snapshot sheet missing header: ${keyHeader}`)

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
  const norm = typeof normalizeFn === 'function' ? normalizeFn : (v => String(v || '').trim())

  for (const r of data) {
    const d = String(r[snapIdx] || '').trim()
    if (d !== snapshotDate) continue

    const key = norm(r[keyIdx])
    if (!key) continue

    set.add(snapshotDate + '|' + key)
  }

  return set
}

/* =========================
 * Shared util compatibility wrappers
 * ========================= */

function getOrCreateSheetCompat_(ss, name) {
  if (!ss) ss = SpreadsheetApp.getActive()
  name = String(name || '').trim()
  if (!name) throw new Error('getOrCreateSheetCompat_: sheet name is required')

  if (typeof getOrCreateSheet === 'function') {
    try { return getOrCreateSheet(ss, name) } catch (e) {}
    try { return getOrCreateSheet(name) } catch (e) {}
  }

  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function normalizeEmailCompat_(email) {
  if (typeof normalizeEmail === 'function') return normalizeEmail(email)
  return String(email || '').trim().toLowerCase()
}

function batchSetValuesCompat_(sheet, startRow, startCol, values, chunkSize) {
  if (typeof batchSetValues === 'function') {
    return batchSetValues(sheet, startRow, startCol, values, chunkSize)
  }

  const size = chunkSize || 2000
  for (let i = 0; i < values.length; i += size) {
    const chunk = values.slice(i, i + size)
    sheet.getRange(startRow + i, startCol, chunk.length, chunk[0].length).setValues(chunk)
  }
}

function lockWrapCompat_(lockName, fn) {
  if (typeof lockWrap === 'function') {
    try {
      // preferred: lockWrap(lockName, fn)
      return lockWrap(lockName, fn)
    } catch (e) {
      // alternate: lockWrap(fn)
      return lockWrap(fn)
    }
  }

  const lock = LockService.getScriptLock()
  const ok = lock.tryLock(300000) // 5 minutes
  if (!ok) throw new Error(`Could not acquire lock: ${lockName}`)
  try {
    return fn()
  } finally {
    lock.releaseLock()
  }
}
