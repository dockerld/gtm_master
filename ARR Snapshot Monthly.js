/**************************************************************
 * write_arr_snapshot_monthly()
 *
 * Monthly snapshot for ARR with BOM/EOM + upgrade/downgrade deltas.
 * Appends to "arr_snapshot" with snapshot_date as first column.
 *
 * Depends on shared helpers (defined elsewhere):
 * - contiguousHeaderWidth_
 * - ensureSnapshotHeaders_
 * - buildExistingSnapshotKeySetGeneric_
 * - getOrCreateSheetCompat_
 * - batchSetValuesCompat_
 * - lockWrapCompat_
 **************************************************************/

const ARR_SNAP_CFG = {
  SOURCE_SHEET: 'arr_raw_data',
  SNAP_SHEET: 'arr_snapshot',

  HEADER_ROW: 2,
  START_COL: 1,
  DATA_START_ROW: 3,

  SNAPSHOT_DATE_HEADER: 'snapshot_date',
  SNAPSHOT_DATE_FMT: 'yyyy-MM-dd',

  KEY_HEADER: 'org_id',
  ARR_HEADER: 'total_arr',
  BOM_HEADER: 'bom_arr',
  EOM_HEADER: 'eom_arr',
  UPGRADE_HEADER: 'upgrade_arr',
  DOWNGRADE_HEADER: 'downgrade_arr',
  WRITE_CHUNK: 3000
}

function write_arr_snapshot() {
  return write_arr_snapshot_monthly()
}

function write_arr_snapshot_monthly() {
  lockWrapCompat_('write_arr_snapshot_monthly', () => {
    const t0 = new Date()
    const ss = SpreadsheetApp.getActive()

    const src = ss.getSheetByName(ARR_SNAP_CFG.SOURCE_SHEET)
    if (!src) throw new Error(`Source sheet not found: ${ARR_SNAP_CFG.SOURCE_SHEET}`)

    const snap = getOrCreateSheetCompat_(ss, ARR_SNAP_CFG.SNAP_SHEET)

    const tz = Session.getScriptTimeZone()
    const snapshotDate = Utilities.formatDate(new Date(), tz, ARR_SNAP_CFG.SNAPSHOT_DATE_FMT)

    const maxColsFromStart = src.getLastColumn() - ARR_SNAP_CFG.START_COL + 1
    if (maxColsFromStart <= 0) throw new Error('arr_raw_data has no columns in the expected region')

    const rawHeaderRow = src
      .getRange(ARR_SNAP_CFG.HEADER_ROW, ARR_SNAP_CFG.START_COL, 1, maxColsFromStart)
      .getValues()[0]
      .map(h => String(h || '').trim())

    const headerWidth = contiguousHeaderWidth_(rawHeaderRow)
    if (headerWidth <= 0) throw new Error('arr_raw_data header row appears empty')

    const srcHeaders = rawHeaderRow.slice(0, headerWidth)
    const keyIdxInSrc = srcHeaders.findIndex(h => h.toLowerCase() === ARR_SNAP_CFG.KEY_HEADER.toLowerCase())
    if (keyIdxInSrc < 0) throw new Error(`arr_raw_data missing header: ${ARR_SNAP_CFG.KEY_HEADER}`)

    const arrIdxInSrc = srcHeaders.findIndex(h => h.toLowerCase() === ARR_SNAP_CFG.ARR_HEADER.toLowerCase())
    if (arrIdxInSrc < 0) throw new Error(`arr_raw_data missing header: ${ARR_SNAP_CFG.ARR_HEADER}`)

    const srcLastRow = src.getLastRow()
    if (srcLastRow < ARR_SNAP_CFG.DATA_START_ROW) {
      Logger.log('No data rows in arr_raw_data. Snapshot skipped.')
      return
    }

    const numRows = srcLastRow - ARR_SNAP_CFG.DATA_START_ROW + 1
    const srcData = src.getRange(ARR_SNAP_CFG.DATA_START_ROW, ARR_SNAP_CFG.START_COL, numRows, headerWidth).getValues()

    const rows = []
    for (const r of srcData) {
      const key = String(r[keyIdxInSrc] || '').trim()
      if (!key) continue
      rows.push(r)
    }

    if (!rows.length) {
      Logger.log('No rows to snapshot in arr_raw_data. Snapshot skipped.')
      return
    }

    const snapHeaders = [ARR_SNAP_CFG.SNAPSHOT_DATE_HEADER]
      .concat(srcHeaders)
      .concat([
        ARR_SNAP_CFG.BOM_HEADER,
        ARR_SNAP_CFG.EOM_HEADER,
        ARR_SNAP_CFG.UPGRADE_HEADER,
        ARR_SNAP_CFG.DOWNGRADE_HEADER
      ])
    ensureSnapshotHeaders_(snap, snapHeaders)

    const existingKeys = buildExistingSnapshotKeySetGeneric_(
      snap,
      snapshotDate,
      ARR_SNAP_CFG.SNAPSHOT_DATE_HEADER,
      ARR_SNAP_CFG.KEY_HEADER,
      v => String(v || '').trim()
    )

    const prevMonthEomByOrg = ARR_monthly_buildPrevMonthEomByOrg_(
      snap,
      snapshotDate,
      ARR_SNAP_CFG.SNAPSHOT_DATE_HEADER,
      ARR_SNAP_CFG.KEY_HEADER,
      ARR_SNAP_CFG.EOM_HEADER,
      ARR_SNAP_CFG.ARR_HEADER
    )

    const out = []
    let skipped = 0

    for (const r of rows) {
      const key = String(r[keyIdxInSrc] || '').trim()
      const mapKey = snapshotDate + '|' + key
      if (existingKeys.has(mapKey)) {
        skipped++
        continue
      }
      existingKeys.add(mapKey)

      const eom = ARR_monthly_num_(r[arrIdxInSrc])
      const bom = prevMonthEomByOrg.get(key) || 0
      const delta = eom - bom
      const upgrade = delta > 0 ? delta : 0
      const downgrade = delta < 0 ? Math.abs(delta) : 0

      out.push([snapshotDate].concat(r).concat([bom, eom, upgrade, downgrade]))
    }

    if (!out.length) {
      Logger.log(`No new rows to snapshot for ${snapshotDate}. Skipped existing: ${skipped}`)
      return
    }

    const startRow = snap.getLastRow() + 1
    batchSetValuesCompat_(snap, startRow, 1, out, ARR_SNAP_CFG.WRITE_CHUNK)

    Logger.log(
      `ARR snapshot ${snapshotDate}: appended ${out.length} rows. ` +
      `Skipped existing: ${skipped}. Took ${((new Date() - t0) / 1000).toFixed(2)}s`
    )
  })
}

function ARR_monthly_buildPrevMonthEomByOrg_(sheet, snapshotDateStr, snapDateHeader, keyHeader, eomHeader, arrHeader) {
  const prevMonthKey = ARR_monthly_prevMonthKey_(snapshotDateStr)
  const out = new Map()
  if (!prevMonthKey) return out

  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < 2) return out

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())
  const snapIdx = header.findIndex(h => h.toLowerCase() === String(snapDateHeader).toLowerCase())
  const keyIdx = header.findIndex(h => h.toLowerCase() === String(keyHeader).toLowerCase())
  const eomIdx = header.findIndex(h => h.toLowerCase() === String(eomHeader).toLowerCase())
  const arrIdx = header.findIndex(h => h.toLowerCase() === String(arrHeader).toLowerCase())

  if (snapIdx < 0 || keyIdx < 0 || (eomIdx < 0 && arrIdx < 0)) return out

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
  const bestByOrg = new Map()

  for (const r of data) {
    const dateStr = String(r[snapIdx] || '').trim()
    if (!dateStr) continue
    if (ARR_monthly_monthKey_(dateStr) !== prevMonthKey) continue

    const orgId = String(r[keyIdx] || '').trim()
    if (!orgId) continue

    const ms = ARR_monthly_dateMs_(dateStr)
    if (ms == null) continue

    let val = 0
    if (eomIdx >= 0) {
      const raw = r[eomIdx]
      if (raw != null && raw !== '') {
        val = ARR_monthly_num_(raw)
      } else if (arrIdx >= 0) {
        val = ARR_monthly_num_(r[arrIdx])
      }
    } else if (arrIdx >= 0) {
      val = ARR_monthly_num_(r[arrIdx])
    }
    const existing = bestByOrg.get(orgId)
    if (!existing || ms > existing.ms) bestByOrg.set(orgId, { ms, val })
  }

  bestByOrg.forEach((v, k) => out.set(k, v.val))
  return out
}

function ARR_monthly_monthKey_(dateStr) {
  const s = String(dateStr || '').trim()
  const parts = s.split('-')
  if (parts.length < 2) return ''
  const y = parts[0]
  const m = parts[1]
  if (!y || !m) return ''
  return y + '-' + m
}

function ARR_monthly_prevMonthKey_(dateStr) {
  const s = String(dateStr || '').trim()
  const parts = s.split('-')
  if (parts.length < 2) return ''
  const y = Number(parts[0])
  const m = Number(parts[1])
  if (!isFinite(y) || !isFinite(m) || m < 1 || m > 12) return ''

  const prevY = m === 1 ? y - 1 : y
  const prevM = m === 1 ? 12 : (m - 1)
  return String(prevY) + '-' + ARR_monthly_pad2_(prevM)
}

function ARR_monthly_dateMs_(dateStr) {
  const s = String(dateStr || '').trim()
  const parts = s.split('-')
  if (parts.length < 3) return null
  const y = Number(parts[0])
  const m = Number(parts[1])
  const d = Number(parts[2])
  if (!isFinite(y) || !isFinite(m) || !isFinite(d)) return null
  return new Date(y, m - 1, d).getTime()
}

function ARR_monthly_pad2_(n) {
  const s = String(Math.floor(Math.abs(Number(n) || 0)))
  return s.length === 1 ? '0' + s : s
}

function ARR_monthly_num_(v) {
  const n = Number(v)
  return isFinite(n) ? n : 0
}
