/**************************************************************
 * Shared snapshot helpers
 *
 * (The snap_users_daily producer was removed — no longer generated.)
 * These helpers are still used by the ARR snapshot/waterfall steps:
 *   contiguousHeaderWidth_, ensureSnapshotHeaders_,
 *   buildExistingSnapshotKeySetGeneric_, getOrCreateSheetCompat_,
 *   batchSetValuesCompat_, lockWrapCompat_, normalizeEmailCompat_
 **************************************************************/

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
      // Only fallback for legacy lockWrap(fn) signatures.
      // If the wrapped fn threw, preserve the original error.
      const msg = String(e && e.message ? e.message : e)
      const signatureMismatch =
        msg.indexOf('fn must be a function') >= 0 ||
        msg.indexOf('lockName') >= 0
      if (!signatureMismatch) throw e
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
