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
  WRITE_CHUNK: 3000,

  COHORT_HEADER: 'sign_up_cohort_month',
  COHORT_FMT: 'MMM yyyy'
}

const ARR_SNAPSHOT_MIGRATE_CFG = {
  SHEET: 'arr_snapshot',
  POSTHOG_ORGS_SHEET: 'raw_posthog_orgs',
  HEADER_ROW: 1,
  DATA_START_ROW: 2,
  OUT_HEADERS: [
    'snapshot_date',
    'org_id',
    'org_name',
    'org_creation_date',
    'first_payment_date',
    'churn_date',
    'sign_up_cohort_month',
    'first_payment_cohort_month',
    'current_status',
    'plan_name',
    'billing_frequency',
    'total_arr',
    'subscription_start_date'
  ],
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
    ARR_snap_pruneLatestIfNotMonthStart_(snap)

    const snapshotDate = ARR_snap_utcDateStr_(new Date())

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

    const snapHeaders = [ARR_SNAP_CFG.SNAPSHOT_DATE_HEADER].concat(srcHeaders)
    ensureSnapshotHeaders_(snap, snapHeaders)

    const existingKeys = buildExistingSnapshotKeySetGeneric_(
      snap,
      snapshotDate,
      ARR_SNAP_CFG.SNAPSHOT_DATE_HEADER,
      ARR_SNAP_CFG.KEY_HEADER,
      v => String(v || '').trim()
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

      r[arrIdxInSrc] = ARR_monthly_num_(r[arrIdxInSrc])
      out.push([snapshotDate].concat(r))
    }

    if (!out.length) {
      Logger.log(`No new rows to snapshot for ${snapshotDate}. Skipped existing: ${skipped}`)
      return
    }

    const startRow = snap.getLastRow() + 1
    batchSetValuesCompat_(snap, startRow, 1, out, ARR_SNAP_CFG.WRITE_CHUNK)
    ARR_snap_applyCohortFormat_(snap)

    Logger.log(
      `ARR snapshot ${snapshotDate}: appended ${out.length} rows. ` +
      `Skipped existing: ${skipped}. Took ${((new Date() - t0) / 1000).toFixed(2)}s`
    )
  })
}

/**
 * One-time migration:
 * - Rewrites arr_snapshot to the new arr_raw_data-aligned schema
 * - Updates org_id to app_org_id when available
 * - Renames/moves legacy fields (e.g. purchase_date -> first_payment_date)
 */
function one_time_migrate_arr_snapshot_to_new_schema() {
  lockWrapCompat_('one_time_migrate_arr_snapshot_to_new_schema', () => {
    const ss = SpreadsheetApp.getActive()
    const sh = ss.getSheetByName(ARR_SNAPSHOT_MIGRATE_CFG.SHEET)
    if (!sh) throw new Error(`Missing sheet: ${ARR_SNAPSHOT_MIGRATE_CFG.SHEET}`)
    const shPosthogOrgs = ss.getSheetByName(ARR_SNAPSHOT_MIGRATE_CFG.POSTHOG_ORGS_SHEET)

    const lastRow = sh.getLastRow()
    const lastCol = sh.getLastColumn()
    if (lastRow < ARR_SNAPSHOT_MIGRATE_CFG.HEADER_ROW || lastCol < 1) {
      throw new Error('arr_snapshot appears empty.')
    }

    const rawHeader = sh
      .getRange(ARR_SNAPSHOT_MIGRATE_CFG.HEADER_ROW, 1, 1, lastCol)
      .getValues()[0]
      .map(h => String(h || '').trim())
    const width = contiguousHeaderWidth_(rawHeader)
    if (width <= 0) throw new Error('arr_snapshot header row appears empty')

    const headers = rawHeader.slice(0, width)
    const idx = ARR_snap_migrate_headerMap_(headers)
    const orgIdByName = ARR_snap_buildOrgIdByNameMapFromPosthog_(shPosthogOrgs)

    const numRows = Math.max(0, lastRow - ARR_SNAPSHOT_MIGRATE_CFG.HEADER_ROW)
    const data = numRows
      ? sh.getRange(ARR_SNAPSHOT_MIGRATE_CFG.DATA_START_ROW, 1, numRows, width).getValues()
      : []

    const out = data.map(r => {
      const oldOrgId = ARR_snap_migrate_pick_(r, idx, ['org_id'])
      const orgName = String(ARR_snap_migrate_pick_(r, idx, ['org_name']) || '').trim()
      const mappedOrgId = ARR_snap_mapOrgIdFromName_(orgName, orgIdByName)
      const resolvedOrgId = String(mappedOrgId || oldOrgId || '').trim()

      const orgCreation = ARR_snap_migrate_pick_(r, idx, ['org_creation_date'])
      const signUpCohortMonth =
        ARR_snap_migrate_pick_(r, idx, ['sign_up_cohort_month']) ||
        ARR_snap_migrate_pick_(r, idx, ['cohort_month']) ||
        ARR_snap_migrate_pick_(r, idx, ['trial_cohort_month'])
      const firstPaymentCohort =
        ARR_snap_migrate_pick_(r, idx, ['first_payment_cohort_month']) ||
        ARR_snap_monthKeyFromDateLike_(ARR_snap_migrate_pick_(r, idx, ['first_payment_date', 'purchase_date']))

      return [
        ARR_snap_migrate_pick_(r, idx, ['snapshot_date']),
        resolvedOrgId,
        orgName,
        orgCreation,
        ARR_snap_migrate_pick_(r, idx, ['first_payment_date', 'purchase_date']),
        ARR_snap_migrate_pick_(r, idx, ['churn_date']),
        signUpCohortMonth,
        firstPaymentCohort,
        ARR_snap_migrate_pick_(r, idx, ['current_status']),
        ARR_snap_migrate_pick_(r, idx, ['plan_name']),
        ARR_snap_migrate_pick_(r, idx, ['billing_frequency']),
        ARR_snap_migrate_pick_(r, idx, ['total_arr', 'eom_arr']),
        ARR_snap_migrate_pick_(r, idx, ['subscription_start_date'])
      ]
    })

    sh.clearContents()
    sh.getRange(ARR_SNAPSHOT_MIGRATE_CFG.HEADER_ROW, 1, 1, ARR_SNAPSHOT_MIGRATE_CFG.OUT_HEADERS.length)
      .setValues([ARR_SNAPSHOT_MIGRATE_CFG.OUT_HEADERS])

    if (out.length) {
      batchSetValuesCompat_(
        sh,
        ARR_SNAPSHOT_MIGRATE_CFG.DATA_START_ROW,
        1,
        out,
        ARR_SNAPSHOT_MIGRATE_CFG.WRITE_CHUNK
      )
    }

    ARR_snap_applyCohortFormat_(sh)
    sh.setFrozenRows(1)
    sh.autoResizeColumns(1, ARR_SNAPSHOT_MIGRATE_CFG.OUT_HEADERS.length)

    Logger.log(`one_time_migrate_arr_snapshot_to_new_schema: migrated ${out.length} rows`)
  })
}

/**
 * One-time additive migration:
 * - Adds first_payment_cohort_month to arr_snapshot if missing
 * - Backfills from first_payment_date (format: YYYY-MM)
 */
function one_time_add_first_payment_cohort_to_arr_snapshot() {
  lockWrapCompat_('one_time_add_first_payment_cohort_to_arr_snapshot', () => {
    const ss = SpreadsheetApp.getActive()
    const sh = ss.getSheetByName(ARR_SNAPSHOT_MIGRATE_CFG.SHEET)
    if (!sh) throw new Error(`Missing sheet: ${ARR_SNAPSHOT_MIGRATE_CFG.SHEET}`)

    const lastRow = sh.getLastRow()
    const lastCol = sh.getLastColumn()
    if (lastRow < 1 || lastCol < 1) throw new Error('arr_snapshot appears empty.')

    const header = sh.getRange(ARR_SNAPSHOT_MIGRATE_CFG.HEADER_ROW, 1, 1, lastCol).getValues()[0]
      .map(h => String(h || '').trim())

    const fpCohortHeader = 'first_payment_cohort_month'
    let fpCohortIdx = header.findIndex(h => h.toLowerCase() === fpCohortHeader)
    const fpDateIdx = header.findIndex(h => h.toLowerCase() === 'first_payment_date')
    if (fpDateIdx < 0) throw new Error('arr_snapshot missing header: first_payment_date')

    if (fpCohortIdx < 0) {
      fpCohortIdx = header.length
      header.push(fpCohortHeader)
      sh.getRange(ARR_SNAPSHOT_MIGRATE_CFG.HEADER_ROW, 1, 1, header.length).setValues([header])
    }

    if (lastRow < ARR_SNAPSHOT_MIGRATE_CFG.DATA_START_ROW) {
      Logger.log('one_time_add_first_payment_cohort_to_arr_snapshot: no data rows to backfill')
      return
    }

    const numRows = lastRow - ARR_SNAPSHOT_MIGRATE_CFG.DATA_START_ROW + 1
    const data = sh.getRange(ARR_SNAPSHOT_MIGRATE_CFG.DATA_START_ROW, 1, numRows, header.length).getValues()
    const out = data.map(r => {
      const row = r.slice()
      row[fpCohortIdx] = ARR_snap_monthKeyFromDateLike_(row[fpDateIdx])
      return row
    })

    sh.getRange(ARR_SNAPSHOT_MIGRATE_CFG.DATA_START_ROW, 1, out.length, header.length).setValues(out)
    sh.getRange(ARR_SNAPSHOT_MIGRATE_CFG.DATA_START_ROW, fpCohortIdx + 1, out.length, 1).setNumberFormat('mmm yyyy')
    sh.autoResizeColumn(fpCohortIdx + 1)

    Logger.log(`one_time_add_first_payment_cohort_to_arr_snapshot: updated ${out.length} rows`)
  })
}

/**
 * One-time backfill:
 * - Fills missing org_creation_date and sign_up_cohort_month in arr_snapshot
 * - Uses raw_posthog_orgs by org_id, fallback by org_name
 */
function one_time_backfill_arr_snapshot_org_creation_and_cohort() {
  lockWrapCompat_('one_time_backfill_arr_snapshot_org_creation_and_cohort', () => {
    const ss = SpreadsheetApp.getActive()
    const shSnap = ss.getSheetByName(ARR_SNAPSHOT_MIGRATE_CFG.SHEET)
    const shOrgs = ss.getSheetByName(ARR_SNAPSHOT_MIGRATE_CFG.POSTHOG_ORGS_SHEET)
    if (!shSnap) throw new Error(`Missing sheet: ${ARR_SNAPSHOT_MIGRATE_CFG.SHEET}`)
    if (!shOrgs) throw new Error(`Missing sheet: ${ARR_SNAPSHOT_MIGRATE_CFG.POSTHOG_ORGS_SHEET}`)

    const snapLastRow = shSnap.getLastRow()
    const snapLastCol = shSnap.getLastColumn()
    if (snapLastRow < 2 || snapLastCol < 1) {
      Logger.log('one_time_backfill_arr_snapshot_org_creation_and_cohort: no snapshot rows')
      return
    }

    const snapHeader = shSnap.getRange(1, 1, 1, snapLastCol).getValues()[0].map(h => String(h || '').trim().toLowerCase())
    const idxOrgId = snapHeader.indexOf('org_id')
    const idxOrgName = snapHeader.indexOf('org_name')
    const idxOrgCreated = snapHeader.indexOf('org_creation_date')
    const idxCohort = snapHeader.indexOf('sign_up_cohort_month') >= 0
      ? snapHeader.indexOf('sign_up_cohort_month')
      : snapHeader.indexOf('cohort_month')
    if (idxOrgId < 0 || idxOrgName < 0 || idxOrgCreated < 0 || idxCohort < 0) {
      throw new Error('arr_snapshot missing one of required headers: org_id, org_name, org_creation_date, sign_up_cohort_month')
    }

    const orgLookup = ARR_snap_buildOrgCreatedLookupFromPosthog_(shOrgs)
    const numRows = snapLastRow - 1
    const vals = shSnap.getRange(2, 1, numRows, snapLastCol).getValues()
    let changed = 0

    for (let i = 0; i < vals.length; i++) {
      const row = vals[i]
      const curCreated = row[idxOrgCreated]
      const curCohort = row[idxCohort]
      const needsCreated = (curCreated == null || curCreated === '')
      const needsCohort = (curCohort == null || curCohort === '')
      if (!needsCreated && !needsCohort) continue

      const orgId = String(row[idxOrgId] || '').trim()
      const orgName = String(row[idxOrgName] || '').trim()
      const created = ARR_snap_lookupOrgCreated_(orgLookup, orgId, orgName)
      if (!created) continue

      if (needsCreated) row[idxOrgCreated] = created
      if (needsCohort) row[idxCohort] = ARR_snap_monthKeyFromDateLike_(created)
      changed++
    }

    if (changed > 0) {
      shSnap.getRange(2, 1, vals.length, snapLastCol).setValues(vals)
      shSnap.getRange(2, idxOrgCreated + 1, vals.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss')
      shSnap.getRange(2, idxCohort + 1, vals.length, 1).setNumberFormat('mmm yyyy')
    }

    Logger.log(`one_time_backfill_arr_snapshot_org_creation_and_cohort: updated ${changed} rows`)
  })
}

function ARR_snap_applyCohortFormat_(sheet) {
  const lastCol = sheet.getLastColumn()
  const lastRow = sheet.getLastRow()
  if (lastCol < 1 || lastRow < 2) return

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())
  const numRows = Math.max(1, sheet.getMaxRows() - 1)

  const applyFmt = (headerName, fmt) => {
    const idx = header.findIndex(h => h.toLowerCase() === String(headerName).toLowerCase())
    if (idx < 0) return
    sheet.getRange(2, idx + 1, numRows, 1).setNumberFormat(fmt)
  }

  applyFmt(ARR_SNAP_CFG.SNAPSHOT_DATE_HEADER, ARR_SNAP_CFG.SNAPSHOT_DATE_FMT)
  applyFmt(ARR_SNAP_CFG.COHORT_HEADER, ARR_SNAP_CFG.COHORT_FMT)
  applyFmt('first_payment_cohort_month', ARR_SNAP_CFG.COHORT_FMT)
  applyFmt(ARR_SNAP_CFG.ARR_HEADER, '0')
}

function ARR_snap_pruneLatestIfNotMonthStart_(sheet) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < 2 || lastCol < 1) return

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())
  const snapIdx = header.findIndex(h => h.toLowerCase() === ARR_SNAP_CFG.SNAPSHOT_DATE_HEADER.toLowerCase())
  if (snapIdx < 0) return

  const latestKey = ARR_snap_latestSnapshotKey_(sheet, snapIdx)
  if (!latestKey) return

  const day = latestKey.split('-')[2]
  if (day === '01') return

  ARR_snap_deleteSnapshotKey_(sheet, snapIdx, latestKey)
}

function ARR_snap_latestSnapshotKey_(sheet, snapIdx) {
  const lastRow = sheet.getLastRow()
  if (lastRow < 2) return ''

  const vals = sheet.getRange(2, snapIdx + 1, lastRow - 1, 1).getValues()
  let best = ''

  for (const r of vals) {
    const key = ARR_snap_normSnapshotKey_(r[0])
    if (!key) continue
    if (!best || key > best) best = key
  }

  return best
}

function ARR_snap_deleteSnapshotKey_(sheet, snapIdx, snapshotKey) {
  const lastRow = sheet.getLastRow()
  if (lastRow < 2) return

  const vals = sheet.getRange(2, snapIdx + 1, lastRow - 1, 1).getValues()
  const rows = []

  for (let i = 0; i < vals.length; i++) {
    const key = ARR_snap_normSnapshotKey_(vals[i][0])
    if (key === snapshotKey) rows.push(i + 2)
  }

  if (!rows.length) return

  rows.sort((a, b) => a - b)
  const ranges = []
  let start = rows[0]
  let prev = rows[0]

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i]
    if (r === prev + 1) {
      prev = r
      continue
    }
    ranges.push([start, prev])
    start = r
    prev = r
  }
  ranges.push([start, prev])

  for (let i = ranges.length - 1; i >= 0; i--) {
    const range = ranges[i]
    sheet.deleteRows(range[0], range[1] - range[0] + 1)
  }
}

function ARR_snap_normSnapshotKey_(v) {
  if (!v) return ''
  if (v instanceof Date) return ARR_snap_utcDateStr_(v)

  const s = String(v || '').trim()
  if (!s) return ''
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s

  const d = new Date(s)
  return isNaN(d.getTime()) ? '' : ARR_snap_utcDateStr_(d)
}

function ARR_snap_migrate_headerMap_(headers) {
  const out = {}
  ;(headers || []).forEach((h, i) => {
    const key = String(h || '').trim().toLowerCase()
    if (!key) return
    out[key] = i
  })
  return out
}

function ARR_snap_migrate_pick_(row, idx, keys) {
  for (const k of (keys || [])) {
    const key = String(k || '').trim().toLowerCase()
    if (!key) continue
    const i = idx[key]
    if (i == null || i < 0) continue
    const v = row[i]
    if (v !== '' && v != null) return v
  }
  return ''
}

function ARR_snap_normOrgName_(v) {
  const s = String(v || '').trim().toLowerCase()
  if (!s) return ''
  return s.replace(/\s+/g, ' ')
}

function ARR_snap_buildOrgIdByNameMapFromPosthog_(shPosthogOrgs) {
  const byName = new Map()
  if (!shPosthogOrgs) return byName

  const rows = ARR_snap_readObjects_(shPosthogOrgs, 1)
  ;(rows || []).forEach(r => {
    const name = String(r.org_name || r.name || '').trim()
    const orgId = String(r.org_id || r.app_org_id || r.id || '').trim()
    const key = ARR_snap_normOrgName_(name)
    if (!key || !orgId) return

    const hasLatestSubscriptionId = !!String(r.latest_subscription_id || '').trim()
    const ts = ARR_snap_pickRowTimestampMs_(r)
    const prev = byName.get(key)
    if (!prev) {
      byName.set(key, { orgId, ts, hasLatestSubscriptionId })
      return
    }

    // Prefer rows that have latest_subscription_id populated.
    if (!prev.hasLatestSubscriptionId && hasLatestSubscriptionId) {
      byName.set(key, { orgId, ts, hasLatestSubscriptionId })
      return
    }
    if (prev.hasLatestSubscriptionId && !hasLatestSubscriptionId) {
      return
    }

    // Tie-breaker: most recent timestamp.
    if (ts >= prev.ts) {
      byName.set(key, { orgId, ts, hasLatestSubscriptionId })
    }
  })

  const out = new Map()
  byName.forEach((v, k) => out.set(k, String(v.orgId || '').trim()))
  return out
}

function ARR_snap_buildOrgCreatedLookupFromPosthog_(shPosthogOrgs) {
  const byId = new Map()
  const byName = new Map()
  if (!shPosthogOrgs) return { byId, byName }

  const rows = ARR_snap_readObjects_(shPosthogOrgs, 1)
  ;(rows || []).forEach(r => {
    const orgId = String(r.org_id || r.app_org_id || r.id || '').trim()
    const orgName = String(r.org_name || r.name || '').trim()
    const createdAt = r.created_at || ''
    const createdIso = ARR_snap_toIsoOrBlank_(createdAt)
    if (!createdIso) return
    if (orgId && !byId.has(orgId)) byId.set(orgId, createdIso)
    const key = ARR_snap_normOrgName_(orgName)
    if (key && !byName.has(key)) byName.set(key, createdIso)
  })
  return { byId, byName }
}

function ARR_snap_lookupOrgCreated_(lookup, orgId, orgName) {
  const byId = lookup && lookup.byId ? lookup.byId : new Map()
  const byName = lookup && lookup.byName ? lookup.byName : new Map()
  const id = String(orgId || '').trim()
  if (id && byId.has(id)) return byId.get(id)
  const key = ARR_snap_normOrgName_(orgName)
  if (key && byName.has(key)) return byName.get(key)
  return ''
}

function ARR_snap_mapOrgIdFromName_(orgName, orgIdByName) {
  if (!orgIdByName || !(orgIdByName instanceof Map)) return ''
  const key = ARR_snap_normOrgName_(orgName)
  if (!key) return ''
  return String(orgIdByName.get(key) || '').trim()
}

function ARR_snap_pickRowTimestampMs_(r) {
  const vals = [r && r.updated_at, r && r.pulled_at, r && r.created_at]
  for (const v of vals) {
    const ms = ARR_snap_toMs_(v)
    if (ms > 0) return ms
  }
  return 0
}

function ARR_snap_toMs_(v) {
  if (!v) return 0
  if (v instanceof Date) return isNaN(v.getTime()) ? 0 : v.getTime()
  const s = String(v || '').trim()
  if (!s) return 0
  const d = new Date(s)
  return isNaN(d.getTime()) ? 0 : d.getTime()
}

function ARR_snap_toIsoOrBlank_(v) {
  if (!v) return ''
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : v.toISOString()
  const s = String(v || '').trim()
  if (!s) return ''
  const d = new Date(s)
  return isNaN(d.getTime()) ? '' : d.toISOString()
}

function ARR_snap_readObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1 || lastCol < 1) return []
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || '').trim().toLowerCase())
  const vals = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()
  return vals.map(r => {
    const o = {}
    headers.forEach((h, i) => { if (h) o[h] = r[i] })
    return o
  })
}

function ARR_snap_monthKeyFromDateLike_(v) {
  if (!v) return ''
  let d = null
  if (v instanceof Date) {
    d = isNaN(v.getTime()) ? null : v
  } else {
    const s = String(v || '').trim()
    if (!s) return ''
    if (/^\d{4}-\d{2}$/.test(s)) return s
    const parsed = new Date(s)
    d = isNaN(parsed.getTime()) ? null : parsed
  }
  if (!d) return ''
  return String(d.getUTCFullYear()) + '-' + ARR_monthly_pad2_(d.getUTCMonth() + 1)
}

function ARR_snap_utcDateStr_(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return ''
  const y = d.getUTCFullYear()
  const m = ARR_monthly_pad2_(d.getUTCMonth() + 1)
  const day = ARR_monthly_pad2_(d.getUTCDate())
  return String(y) + '-' + m + '-' + day
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
