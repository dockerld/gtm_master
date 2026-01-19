/**************************************************************
 * render_arr_waterfall_facts()
 *
 * Builds the "arr_waterfall_facts" table from arr_snapshot.
 * Output columns:
 *   snapshot_date | cohort_month_trial | cohort_month_subscription | cohort_month_paid |
 *   org_id | org_name | subscription_start_date | trial_start_date | purchase_date |
 *   metric | amount
 *
 * Metrics (rows):
 *   SOM, Upgrade, Downgrade, Churn, EOM
 **************************************************************/

const ARR_WATERFALL_CFG = {
  SOURCE_SHEET: 'arr_snapshot',
  OUT_SHEET: 'arr_waterfall_facts',

  HEADER_ROW: 1,
  DATA_START_ROW: 2,

  SNAPSHOT_DATE_HEADER: 'snapshot_date',
  TRIAL_COHORT_HEADER: 'trial_cohort_month',
  ORG_ID_HEADER: 'org_id',
  ORG_NAME_HEADER: 'org_name',
  SUB_START_HEADER: 'subscription_start_date',
  TRIAL_START_HEADER: 'trial_start_date',
  PURCHASE_DATE_HEADER: 'purchase_date',
  BOM_HEADER: 'bom_arr',
  EOM_HEADER: 'eom_arr',
  ARR_HEADER: 'total_arr',
  COHORT_FMT: 'MMM yyyy',
  SNAPSHOT_FMT: 'MMM dd yyyy',

  WRITE_CHUNK: 4000
}

function render_arr_waterfall_facts() {
  lockWrapCompat_('render_arr_waterfall_facts', () => {
    const t0 = new Date()
    const ss = SpreadsheetApp.getActive()

    const src = ss.getSheetByName(ARR_WATERFALL_CFG.SOURCE_SHEET)
    if (!src) throw new Error(`Source sheet not found: ${ARR_WATERFALL_CFG.SOURCE_SHEET}`)

    const outSheet = getOrCreateSheetCompat_(ss, ARR_WATERFALL_CFG.OUT_SHEET)

    const lastCol = src.getLastColumn()
    if (lastCol < 1) throw new Error('arr_snapshot has no columns')
    const tz = Session.getScriptTimeZone()

    const rawHeader = src
      .getRange(ARR_WATERFALL_CFG.HEADER_ROW, 1, 1, lastCol)
      .getValues()[0]
      .map(h => String(h || '').trim())

    const headerWidth = contiguousHeaderWidth_(rawHeader)
    if (headerWidth <= 0) throw new Error('arr_snapshot header row appears empty')

    const headers = rawHeader.slice(0, headerWidth)
    const snapIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.SNAPSHOT_DATE_HEADER.toLowerCase())
    const cohortIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.TRIAL_COHORT_HEADER.toLowerCase())
    const orgIdIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.ORG_ID_HEADER.toLowerCase())
    const orgNameIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.ORG_NAME_HEADER.toLowerCase())
    const subStartIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.SUB_START_HEADER.toLowerCase())
    const trialStartIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.TRIAL_START_HEADER.toLowerCase())
    const purchaseIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.PURCHASE_DATE_HEADER.toLowerCase())
    const bomIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.BOM_HEADER.toLowerCase())
    const eomIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.EOM_HEADER.toLowerCase())
    const arrIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.ARR_HEADER.toLowerCase())

    if (snapIdx < 0) throw new Error(`arr_snapshot missing header: ${ARR_WATERFALL_CFG.SNAPSHOT_DATE_HEADER}`)
    if (cohortIdx < 0) throw new Error(`arr_snapshot missing header: ${ARR_WATERFALL_CFG.TRIAL_COHORT_HEADER}`)
    if (subStartIdx < 0) throw new Error(`arr_snapshot missing header: ${ARR_WATERFALL_CFG.SUB_START_HEADER}`)
    if (trialStartIdx < 0) throw new Error(`arr_snapshot missing header: ${ARR_WATERFALL_CFG.TRIAL_START_HEADER}`)
    if (purchaseIdx < 0) throw new Error(`arr_snapshot missing header: ${ARR_WATERFALL_CFG.PURCHASE_DATE_HEADER}`)
    if (eomIdx < 0 && arrIdx < 0) {
      throw new Error(`arr_snapshot missing header: ${ARR_WATERFALL_CFG.EOM_HEADER} or ${ARR_WATERFALL_CFG.ARR_HEADER}`)
    }

    const lastRow = src.getLastRow()
    const numRows = Math.max(0, lastRow - ARR_WATERFALL_CFG.HEADER_ROW)
    const data = numRows
      ? src.getRange(ARR_WATERFALL_CFG.DATA_START_ROW, 1, numRows, headerWidth).getValues()
      : []

    const out = []
    for (const r of data) {
      const snapshotDate = ARR_waterfall_formatSnapshot_(r[snapIdx], tz)
      if (!snapshotDate) continue

      const cohortTrial = ARR_waterfall_formatCohort_(r[cohortIdx], tz)
      const orgId = (orgIdIdx >= 0) ? String(r[orgIdIdx] || '').trim() : ''
      const orgName = (orgNameIdx >= 0) ? String(r[orgNameIdx] || '').trim() : ''
      const subscriptionStartRaw = r[subStartIdx]
      const trialStartRaw = r[trialStartIdx]
      const purchaseDateRaw = r[purchaseIdx]
      const cohortSubscription = ARR_waterfall_formatCohort_(subscriptionStartRaw, tz)
      const cohortPaid = ARR_waterfall_formatCohort_(purchaseDateRaw, tz)
      const subscriptionStart = ARR_waterfall_str_(subscriptionStartRaw)
      const trialStart = ARR_waterfall_str_(trialStartRaw)
      const purchaseDate = ARR_waterfall_str_(purchaseDateRaw)
      const eom = ARR_waterfall_num_(ARR_waterfall_pick_(r, eomIdx, arrIdx))
      const bomRaw = (bomIdx >= 0) ? r[bomIdx] : ''
      const bom = ARR_waterfall_num_(bomRaw !== '' && bomRaw != null ? bomRaw : eom)

      const som = bom
      const upgrade = eom > bom ? (eom - bom) : 0
      const churn = (bom > 0 && eom === 0) ? bom : 0
      const downgrade = (eom < bom && eom > 0) ? (bom - eom) : 0

      out.push([
        snapshotDate, cohortTrial, cohortSubscription, cohortPaid,
        orgId, orgName, subscriptionStart, trialStart, purchaseDate,
        'SOM', som
      ])
      out.push([
        snapshotDate, cohortTrial, cohortSubscription, cohortPaid,
        orgId, orgName, subscriptionStart, trialStart, purchaseDate,
        'Upgrade', upgrade
      ])
      out.push([
        snapshotDate, cohortTrial, cohortSubscription, cohortPaid,
        orgId, orgName, subscriptionStart, trialStart, purchaseDate,
        'Downgrade', downgrade
      ])
      out.push([
        snapshotDate, cohortTrial, cohortSubscription, cohortPaid,
        orgId, orgName, subscriptionStart, trialStart, purchaseDate,
        'Churn', churn
      ])
      out.push([
        snapshotDate, cohortTrial, cohortSubscription, cohortPaid,
        orgId, orgName, subscriptionStart, trialStart, purchaseDate,
        'EOM', eom
      ])
    }

    outSheet.clearContents()
    outSheet.getRange(1, 1, 1, 11).setValues([[
      'snapshot_date',
      'cohort_month_trial',
      'cohort_month_subscription',
      'cohort_month_paid',
      'org_id',
      'org_name',
      'subscription_start_date',
      'trial_start_date',
      'purchase_date',
      'metric',
      'amount'
    ]])

    if (out.length) {
      batchSetValuesCompat_(outSheet, 2, 1, out, ARR_WATERFALL_CFG.WRITE_CHUNK)
    }

    outSheet.setFrozenRows(1)
    outSheet.autoResizeColumns(1, 11)

    const seconds = (new Date() - t0) / 1000
    if (typeof writeSyncLog === 'function') {
      writeSyncLog('render_arr_waterfall_facts', 'ok', data.length, out.length, seconds, '')
    } else {
      Logger.log(`[render_arr_waterfall_facts] ok rows_in=${data.length} rows_out=${out.length} seconds=${seconds}`)
    }
  })
}

function ARR_waterfall_pick_(row, primaryIdx, fallbackIdx) {
  if (primaryIdx >= 0) {
    const v = row[primaryIdx]
    if (v != null && v !== '') return v
  }
  if (fallbackIdx >= 0) return row[fallbackIdx]
  return ''
}

function ARR_waterfall_num_(v) {
  const n = Number(v)
  return isFinite(n) ? n : 0
}

function ARR_waterfall_str_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : v.toISOString()
  return String(v || '').trim()
}

function ARR_waterfall_formatCohort_(v, tz) {
  if (!v) return ''
  let d = null

  if (v instanceof Date) {
    d = isNaN(v.getTime()) ? null : v
  } else {
    const s = String(v || '').trim()
    if (!s) return ''

    if (/^\d{4}-\d{2}$/.test(s)) {
      const y = Number(s.slice(0, 4))
      const m = Number(s.slice(5, 7))
      if (isFinite(y) && isFinite(m)) d = new Date(y, m - 1, 1)
    } else if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
      const parsed = new Date(s)
      d = isNaN(parsed.getTime()) ? null : parsed
    }
  }

  if (!d) return String(v || '').trim()
  return Utilities.formatDate(d, tz, ARR_WATERFALL_CFG.COHORT_FMT)
}

function ARR_waterfall_formatSnapshot_(v, tz) {
  if (!v) return ''
  let d = null

  if (v instanceof Date) {
    d = isNaN(v.getTime()) ? null : v
  } else {
    const s = String(v || '').trim()
    if (!s) return ''

    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      const y = Number(s.slice(0, 4))
      const m = Number(s.slice(5, 7))
      const day = Number(s.slice(8, 10))
      if (isFinite(y) && isFinite(m) && isFinite(day)) d = new Date(y, m - 1, day)
    } else if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
      const parsed = new Date(s)
      d = isNaN(parsed.getTime()) ? null : parsed
    }
  }

  if (!d) return String(v || '').trim()
  return Utilities.formatDate(d, tz, ARR_WATERFALL_CFG.SNAPSHOT_FMT)
}
