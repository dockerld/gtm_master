/**************************************************************
 * render_arr_waterfall_facts()
 *
 * Builds the "arr_waterfall_facts" table from arr_snapshot.
 * Output columns:
 *   snapshot_date | org_id | org_name | org_creation_date | first_payment_date |
 *   churn_date | sign_up_cohort_month | paid_cohort_month | current_status |
 *   ring_bucket | plan_name | billing_frequency | subscription_start_date |
 *   metric | amount
 *
 * Metrics (rows):
 *   SOM, New, Upgrade, Downgrade, Churn, EOM
 **************************************************************/

const ARR_WATERFALL_CFG = {
  SOURCE_SHEET: 'arr_snapshot',
  OUT_SHEET: 'arr_waterfall_facts',

  HEADER_ROW: 1,
  DATA_START_ROW: 2,

  SNAPSHOT_DATE_HEADER: 'snapshot_date',
  ORG_ID_HEADER: 'org_id',
  ORG_NAME_HEADER: 'org_name',
  ORG_CREATED_HEADER: 'org_creation_date',
  FIRST_PAYMENT_HEADER: 'first_payment_date',
  CHURN_DATE_HEADER: 'churn_date',
  COHORT_HEADER: 'sign_up_cohort_month',
  PAID_COHORT_HEADER: 'paid_cohort_month',
  STATUS_HEADER: 'current_status',
  RING_BUCKET_HEADER: 'ring_bucket',
  PLAN_HEADER: 'plan_name',
  BILLING_FREQ_HEADER: 'billing_frequency',
  SUB_START_HEADER: 'subscription_start_date',
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
    const orgIdIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.ORG_ID_HEADER.toLowerCase())
    const orgNameIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.ORG_NAME_HEADER.toLowerCase())
    const orgCreatedIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.ORG_CREATED_HEADER.toLowerCase())
    const firstPaymentIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.FIRST_PAYMENT_HEADER.toLowerCase())
    const churnDateIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.CHURN_DATE_HEADER.toLowerCase())
    const cohortIdx = headers.findIndex(h =>
      h.toLowerCase() === ARR_WATERFALL_CFG.COHORT_HEADER.toLowerCase() ||
      h.toLowerCase() === 'cohort_month' ||
      h.toLowerCase() === 'trial_cohort_month'
    )
    const paidCohortIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.PAID_COHORT_HEADER.toLowerCase())
    const statusIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.STATUS_HEADER.toLowerCase())
    const ringBucketIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.RING_BUCKET_HEADER.toLowerCase())
    const planIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.PLAN_HEADER.toLowerCase())
    const billingFreqIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.BILLING_FREQ_HEADER.toLowerCase())
    const subStartIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.SUB_START_HEADER.toLowerCase())
    const arrIdx = headers.findIndex(h => h.toLowerCase() === ARR_WATERFALL_CFG.ARR_HEADER.toLowerCase())

    if (snapIdx < 0) throw new Error(`arr_snapshot missing header: ${ARR_WATERFALL_CFG.SNAPSHOT_DATE_HEADER}`)
    if (cohortIdx < 0) throw new Error(`arr_snapshot missing header: ${ARR_WATERFALL_CFG.COHORT_HEADER}`)
    if (subStartIdx < 0) throw new Error(`arr_snapshot missing header: ${ARR_WATERFALL_CFG.SUB_START_HEADER}`)
    if (orgCreatedIdx < 0) throw new Error(`arr_snapshot missing header: ${ARR_WATERFALL_CFG.ORG_CREATED_HEADER}`)
    if (firstPaymentIdx < 0) throw new Error(`arr_snapshot missing header: ${ARR_WATERFALL_CFG.FIRST_PAYMENT_HEADER}`)
    if (arrIdx < 0) throw new Error(`arr_snapshot missing header: ${ARR_WATERFALL_CFG.ARR_HEADER}`)

    const lastRow = src.getLastRow()
    const numRows = Math.max(0, lastRow - ARR_WATERFALL_CFG.HEADER_ROW)
    const data = numRows
      ? src.getRange(ARR_WATERFALL_CFG.DATA_START_ROW, 1, numRows, headerWidth).getValues()
      : []

    const records = []
    for (const r of data) {
      const snapshotDate = ARR_waterfall_formatSnapshot_(r[snapIdx], tz)
      if (!snapshotDate) continue

      const orgId = (orgIdIdx >= 0) ? String(r[orgIdIdx] || '').trim() : ''
      const orgName = (orgNameIdx >= 0) ? String(r[orgNameIdx] || '').trim() : ''
      const orgCreated = ARR_waterfall_str_(r[orgCreatedIdx])
      const firstPayment = ARR_waterfall_str_(r[firstPaymentIdx])
      const churnDate = (churnDateIdx >= 0) ? ARR_waterfall_str_(r[churnDateIdx]) : ''
      const cohortMonth = ARR_waterfall_formatCohort_(r[cohortIdx], tz)
      const paidCohort = (paidCohortIdx >= 0)
        ? ARR_waterfall_formatCohort_(r[paidCohortIdx], tz)
        : ''
      const currentStatus = (statusIdx >= 0) ? ARR_waterfall_str_(r[statusIdx]) : ''
      const ringBucket = (ringBucketIdx >= 0) ? ARR_waterfall_str_(r[ringBucketIdx]) : ''
      const planName = (planIdx >= 0) ? ARR_waterfall_str_(r[planIdx]) : ''
      const billingFreq = (billingFreqIdx >= 0) ? ARR_waterfall_str_(r[billingFreqIdx]) : ''
      const subscriptionStart = ARR_waterfall_str_(r[subStartIdx])
      const eom = ARR_waterfall_num_(r[arrIdx])
      const snapshotMs = ARR_waterfall_snapshotMs_(r[snapIdx])
      if (!snapshotMs) continue
      const orgKey = orgId || ('name:' + ARR_waterfall_normName_(orgName))
      if (!orgKey || orgKey === 'name:') continue

      const monthKey = ARR_waterfall_monthKey_(snapshotMs)
      const dayOfMonth = ARR_waterfall_dayOfMonth_(snapshotMs)
      records.push({
        snapshotDate,
        snapshotMs,
        orgKey,
        monthKey,
        dayOfMonth,
        orgId,
        orgName,
        orgCreated,
        firstPayment,
        churnDate,
        cohortMonth,
        paidCohort,
        currentStatus,
        ringBucket,
        planName,
        billingFreq,
        subscriptionStart,
        eom
      })
    }

    records.sort((a, b) => {
      if (a.orgKey !== b.orgKey) return String(a.orgKey).localeCompare(String(b.orgKey))
      return a.snapshotMs - b.snapshotMs
    })

    // SOM baseline per org + month comes strictly from month-start snapshot (day 1).
    const somByOrgMonth = new Map()
    for (const rec of records) {
      if (rec.dayOfMonth !== 1) continue
      const key = rec.orgKey + '|' + rec.monthKey
      somByOrgMonth.set(key, rec.eom)
    }

    // Build a map of the first month each org had ARR > 0 (for New vs Upgrade)
    const orgFirstSeenMonth = new Map()
    for (const rec of records) {
      if (rec.dayOfMonth !== 1) continue
      if (ARR_waterfall_num_(rec.eom) <= 0) continue
      const existing = orgFirstSeenMonth.get(rec.orgKey)
      if (!existing || rec.monthKey < existing) {
        orgFirstSeenMonth.set(rec.orgKey, rec.monthKey)
      }
    }

    const out = []
    for (const rec of records) {
      const somKey = rec.orgKey + '|' + rec.monthKey
      const som = ARR_waterfall_num_(somByOrgMonth.get(somKey) || 0)
      const eom = ARR_waterfall_num_(rec.eom)
      let newCustomer = 0
      let upgrade = 0
      let downgrade = 0
      let churn = 0

      if (som > 0 && eom === 0) {
        churn = som
      } else if (som === 0 && eom > 0) {
        // First time this org appears with ARR — it's a new customer
        const firstMonth = orgFirstSeenMonth.get(rec.orgKey) || rec.monthKey
        if (rec.monthKey === firstMonth) {
          newCustomer = eom
        } else {
          // Was previously seen but SOM is 0 this month — reactivation, treat as upgrade
          upgrade = eom
        }
      } else if (eom > som) {
        upgrade = eom - som
      } else if (eom < som && eom > 0) {
        downgrade = som - eom
      }

      const base = [
        rec.snapshotDate, rec.orgId, rec.orgName, rec.orgCreated, rec.firstPayment, rec.churnDate,
        rec.cohortMonth, rec.paidCohort, rec.currentStatus, rec.ringBucket, rec.planName, rec.billingFreq,
        rec.subscriptionStart
      ]
      out.push([...base, 'SOM', som])
      out.push([...base, 'New', newCustomer])
      out.push([...base, 'Upgrade', upgrade])
      out.push([...base, 'Downgrade', downgrade])
      out.push([...base, 'Churn', churn])
      out.push([...base, 'EOM', eom])
    }

    outSheet.clearContents()
    const outHeaders = [
      'snapshot_date',
      'org_id',
      'org_name',
      'org_creation_date',
      'first_payment_date',
      'churn_date',
      'sign_up_cohort_month',
      'paid_cohort_month',
      'current_status',
      'ring_bucket',
      'plan_name',
      'billing_frequency',
      'subscription_start_date',
      'metric',
      'amount'
    ]
    outSheet.getRange(1, 1, 1, outHeaders.length).setValues([outHeaders])

    if (out.length) {
      batchSetValuesCompat_(outSheet, 2, 1, out, ARR_WATERFALL_CFG.WRITE_CHUNK)
    }

    outSheet.setFrozenRows(1)
    outSheet.autoResizeColumns(1, outHeaders.length)

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

function ARR_waterfall_snapshotMs_(v) {
  if (!v) return 0
  if (v instanceof Date) return isNaN(v.getTime()) ? 0 : v.getTime()
  const s = String(v || '').trim()
  if (!s) return 0
  const d = new Date(s)
  return isNaN(d.getTime()) ? 0 : d.getTime()
}

function ARR_waterfall_monthKey_(ms) {
  const d = new Date(ms)
  if (isNaN(d.getTime())) return ''
  const y = d.getUTCFullYear()
  const m = String(d.getUTCMonth() + 1).padStart(2, '0')
  return `${y}-${m}`
}

function ARR_waterfall_dayOfMonth_(ms) {
  const d = new Date(ms)
  if (isNaN(d.getTime())) return 0
  return d.getUTCDate()
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
