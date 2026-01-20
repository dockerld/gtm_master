/**************************************************************
 * write_arr_snapshot_backfill_last_6_months()
 *
 * One-time backfill of arr_snapshot for the last 6 months.
 * Uses current arr_raw_data values and includes rows where
 * Includes rows where org_creation_date < snapshot_date (date-only, UTC).
 **************************************************************/

const ARR_BACKFILL_CFG = {
  SOURCE_SHEET: 'arr_raw_data',
  SNAP_SHEET: 'arr_snapshot',

  HEADER_ROW: 2,
  START_COL: 1,
  DATA_START_ROW: 3,

  SNAPSHOT_DATE_HEADER: 'snapshot_date',
  SNAPSHOT_DATE_FMT: 'yyyy-MM-dd',

  KEY_HEADER: 'org_id',
  ORG_CREATED_HEADER: 'org_creation_date',
  ARR_HEADER: 'total_arr',
  BOM_HEADER: 'bom_arr',
  EOM_HEADER: 'eom_arr',
  UPGRADE_HEADER: 'upgrade_arr',
  DOWNGRADE_HEADER: 'downgrade_arr',

  MONTHS_BACKFILL: 6,
  WRITE_CHUNK: 3000
}

function write_arr_snapshot_backfill_last_6_months() {
  lockWrapCompat_('write_arr_snapshot_backfill_last_6_months', () => {
    const t0 = new Date()
    const ss = SpreadsheetApp.getActive()

    const src = ss.getSheetByName(ARR_BACKFILL_CFG.SOURCE_SHEET)
    if (!src) throw new Error(`Source sheet not found: ${ARR_BACKFILL_CFG.SOURCE_SHEET}`)

    const snap = getOrCreateSheetCompat_(ss, ARR_BACKFILL_CFG.SNAP_SHEET)
    const snapshotDates = ARR_backfill_monthStartDates_(ARR_BACKFILL_CFG.MONTHS_BACKFILL)
    if (!snapshotDates.length) {
      Logger.log('No snapshot dates to backfill. Skipping.')
      return
    }

    const maxColsFromStart = src.getLastColumn() - ARR_BACKFILL_CFG.START_COL + 1
    if (maxColsFromStart <= 0) throw new Error('arr_raw_data has no columns in the expected region')

    const rawHeaderRow = src
      .getRange(ARR_BACKFILL_CFG.HEADER_ROW, ARR_BACKFILL_CFG.START_COL, 1, maxColsFromStart)
      .getValues()[0]
      .map(h => String(h || '').trim())

    const headerWidth = contiguousHeaderWidth_(rawHeaderRow)
    if (headerWidth <= 0) throw new Error('arr_raw_data header row appears empty')

    const srcHeaders = rawHeaderRow.slice(0, headerWidth)
    const keyIdxInSrc = srcHeaders.findIndex(h => h.toLowerCase() === ARR_BACKFILL_CFG.KEY_HEADER.toLowerCase())
    if (keyIdxInSrc < 0) throw new Error(`arr_raw_data missing header: ${ARR_BACKFILL_CFG.KEY_HEADER}`)

    const arrIdxInSrc = srcHeaders.findIndex(h => h.toLowerCase() === ARR_BACKFILL_CFG.ARR_HEADER.toLowerCase())
    if (arrIdxInSrc < 0) throw new Error(`arr_raw_data missing header: ${ARR_BACKFILL_CFG.ARR_HEADER}`)

    const createdIdxInSrc = srcHeaders.findIndex(
      h => h.toLowerCase() === ARR_BACKFILL_CFG.ORG_CREATED_HEADER.toLowerCase()
    )
    if (createdIdxInSrc < 0) throw new Error(`arr_raw_data missing header: ${ARR_BACKFILL_CFG.ORG_CREATED_HEADER}`)

    const srcLastRow = src.getLastRow()
    if (srcLastRow < ARR_BACKFILL_CFG.DATA_START_ROW) {
      Logger.log('No data rows in arr_raw_data. Snapshot skipped.')
      return
    }

    const numRows = srcLastRow - ARR_BACKFILL_CFG.DATA_START_ROW + 1
    const srcData = src.getRange(
      ARR_BACKFILL_CFG.DATA_START_ROW,
      ARR_BACKFILL_CFG.START_COL,
      numRows,
      headerWidth
    ).getValues()

    const rows = []
    for (const r of srcData) {
      const orgId = String(r[keyIdxInSrc] || '').trim()
      if (!orgId) continue

      const createdAt = ARR_backfill_parseDate_(r[createdIdxInSrc])
      if (!createdAt) continue
      rows.push({ orgId, createdAt, row: r })
    }

    if (!rows.length) {
      Logger.log('No rows with org_id and org_creation_date. Snapshot skipped.')
      return
    }

    const snapHeaders = [ARR_BACKFILL_CFG.SNAPSHOT_DATE_HEADER]
      .concat(srcHeaders)
      .concat([
        ARR_BACKFILL_CFG.BOM_HEADER,
        ARR_BACKFILL_CFG.EOM_HEADER,
        ARR_BACKFILL_CFG.UPGRADE_HEADER,
        ARR_BACKFILL_CFG.DOWNGRADE_HEADER
      ])
    ensureSnapshotHeaders_(snap, snapHeaders)

    const existingKeyMap = ARR_backfill_buildExistingKeyMap_(
      snap,
      snapshotDates.map(d => d.dateStr),
      ARR_BACKFILL_CFG.SNAPSHOT_DATE_HEADER,
      ARR_BACKFILL_CFG.KEY_HEADER
    )

    const out = []
    let skipped = 0

    for (const sd of snapshotDates) {
      const keySet = existingKeyMap.get(sd.dateStr) || new Set()

      for (const r of rows) {
        const createdDateStr = ARR_snap_utcDateStr_(r.createdAt)
        if (!createdDateStr || createdDateStr >= sd.dateStr) continue

        r.row[arrIdxInSrc] = ARR_backfill_num_(r.row[arrIdxInSrc])
        const mapKey = sd.dateStr + '|' + r.orgId
        if (keySet.has(mapKey)) {
          skipped++
          continue
        }
        keySet.add(mapKey)
        const eom = r.row[arrIdxInSrc]
        const bom = eom
        out.push([sd.dateStr].concat(r.row).concat([bom, eom, 0, 0]))
      }

      existingKeyMap.set(sd.dateStr, keySet)
    }

    if (!out.length) {
      Logger.log('No new rows to backfill. All snapshot rows already exist.')
      return
    }

    const startRow = snap.getLastRow() + 1
    batchSetValuesCompat_(snap, startRow, 1, out, ARR_BACKFILL_CFG.WRITE_CHUNK)
    if (typeof ARR_snap_applyCohortFormat_ === 'function') {
      ARR_snap_applyCohortFormat_(snap)
    }

    Logger.log(
      `ARR snapshot backfill: appended ${out.length} rows. ` +
      `Skipped existing: ${skipped}. Took ${((new Date() - t0) / 1000).toFixed(2)}s`
    )
  })
}

/* =========================
 * Internal helpers
 * ========================= */

function ARR_backfill_monthStartDates_(monthsBack) {
  const count = Math.max(0, Number(monthsBack) || 0)
  if (!count) return []

  const now = new Date()
  const y = now.getUTCFullYear()
  const m = now.getUTCMonth()
  const list = []

  for (let i = count - 1; i >= 0; i--) {
    const d = new Date(Date.UTC(y, m - i, 1))
    const dateStr = ARR_snap_utcDateStr_(d)
    list.push({ dateStr })
  }

  return list
}

function ARR_backfill_parseDate_(v) {
  if (!v) return null
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v

  const s = String(v || '').trim()
  if (!s) return null

  if (/^\d+$/.test(s)) {
    const n = Number(s)
    const ms = n > 1e12 ? n : n * 1000
    const d = new Date(ms)
    return isNaN(d.getTime()) ? null : d
  }

  const d = new Date(s)
  return isNaN(d.getTime()) ? null : d
}

function ARR_backfill_num_(v) {
  const n = Number(v)
  return isFinite(n) ? n : 0
}

function ARR_backfill_buildExistingKeyMap_(sheet, snapshotDates, snapDateHeader, keyHeader) {
  const map = new Map()
  snapshotDates.forEach(d => map.set(d, new Set()))

  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < 2) return map

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())
  const snapIdx = header.findIndex(h => h.toLowerCase() === String(snapDateHeader).toLowerCase())
  const keyIdx = header.findIndex(h => h.toLowerCase() === String(keyHeader).toLowerCase())

  if (snapIdx < 0) throw new Error(`Snapshot sheet missing header: ${snapDateHeader}`)
  if (keyIdx < 0) throw new Error(`Snapshot sheet missing header: ${keyHeader}`)

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
  const dateSet = new Set(snapshotDates)

  for (const r of data) {
    const dateStr = String(r[snapIdx] || '').trim()
    if (!dateSet.has(dateStr)) continue

    const key = String(r[keyIdx] || '').trim()
    if (!key) continue

    const mapKey = dateStr + '|' + key
    const set = map.get(dateStr)
    if (set) set.add(mapKey)
  }

  return map
}
