/**************************************************************
 * render_arr_kpi_vs_column_audit()
 *
 * One-time audit: compares the two ARR calculation paths used
 * by "The Ring" to surface per-subscription differences.
 *
 * Path A — "KPI ARR" (top of The Ring):
 *   Source: raw_stripe_subscriptions
 *   Logic:  Apply real-time active discounts to Stripe amount,
 *           then computeMrrArr_. Used for paidArr accumulator.
 *
 * Path B — "Column ARR" (table rows):
 *   Source: org_subscription_info
 *   Logic:  Use amount_yearly if > 0, else computeMrrArr_ from
 *           org_subscription_info amount/interval. = infoArr.
 *
 * Output sheet: "ARR KPI vs Column audit"
 *   Section 1: Summary totals + delta
 *   Section 2: Per-subscription rows showing both values + diff
 *
 * Only includes subscriptions that pass Ring filters (active/trialing,
 * not excluded by Manual Stripe Changes, has org_subscription_info match).
 **************************************************************/

const KPIAUDIT_CFG = {
  OUT_SHEET: 'ARR KPI vs Column audit',
  STRIPE_SHEET: 'raw_stripe_subscriptions',
  ORG_SUB_INFO_SHEET: 'org_subscription_info',
  MANUAL_CHANGES_SHEET: 'Manual Stripe Changes',

  CURRENCY_FMT: '$#,##0.00',
  INT_FMT: '0',
  DELTA_FMT: '+$#,##0.00;-$#,##0.00;$0.00'
}

function render_arr_kpi_vs_column_audit() {
  const t0 = new Date()
  const ss = SpreadsheetApp.getActive()

  const shStripe = ss.getSheetByName(KPIAUDIT_CFG.STRIPE_SHEET)
  const shOrgSub = ss.getSheetByName(KPIAUDIT_CFG.ORG_SUB_INFO_SHEET)
  const shManual = ss.getSheetByName(KPIAUDIT_CFG.MANUAL_CHANGES_SHEET)

  if (!shStripe) throw new Error('Missing sheet: ' + KPIAUDIT_CFG.STRIPE_SHEET)
  if (!shOrgSub) throw new Error('Missing sheet: ' + KPIAUDIT_CFG.ORG_SUB_INFO_SHEET)

  const stripeRows = KPIAUDIT_readSheetObjects_(shStripe, 1)
  const orgSubRows = KPIAUDIT_readSheetObjects_(shOrgSub, 1)
  const manualChangesBySubId = KPIAUDIT_buildManualChanges_(shManual)

  // Index org_subscription_info by stripe_subscription_id
  const orgSubBySubId = new Map()
  for (const r of orgSubRows) {
    const subId = KPIAUDIT_str_(r.stripe_subscription_id)
    if (subId && !orgSubBySubId.has(subId)) orgSubBySubId.set(subId, r)
  }

  const asOfNow = new Date()
  const detailRows = []
  let totalKpiArr = 0
  let totalColumnArr = 0
  let paidKpiArr = 0
  let paidColumnArr = 0
  let intentKpiArr = 0
  let intentColumnArr = 0
  let freeTrialKpiArr = 0
  let freeTrialColumnArr = 0

  for (const r of stripeRows) {
    const stripeSubId =
      KPIAUDIT_str_(r.stripe_subscription_id) ||
      KPIAUDIT_str_(r.subscription_id) ||
      KPIAUDIT_str_(r.subscription) ||
      KPIAUDIT_str_(r.id)
    if (!stripeSubId) continue

    // Must have org_subscription_info match (same filter as Ring)
    const orgSubInfo = orgSubBySubId.get(stripeSubId)
    if (!orgSubInfo) continue

    // Exclude internal manual changes (same filter as Ring)
    const manualChange = manualChangesBySubId.get(stripeSubId)
    if (manualChange && manualChange.excludeInternal) continue

    // Determine ring bucket (same logic as Ring)
    const statusRaw = KPIAUDIT_str_(orgSubInfo.status).toLowerCase()
    const firstPaymentAtIso = KPIAUDIT_str_(orgSubInfo.first_payment_at)
    const hasPaymentMethod = KPIAUDIT_toBool_(orgSubInfo.has_payment_method)
    const ringBucket = KPIAUDIT_ringBucket_(statusRaw, firstPaymentAtIso, hasPaymentMethod)
    if (!ringBucket) continue

    // ── Path A: KPI ARR (real-time discount on Stripe amount) ──
    const interval = KPIAUDIT_str_(r.interval).toLowerCase()
    const intervalCount = Math.max(1, KPIAUDIT_num_(r.interval_count) || 1)
    const amountRaw = KPIAUDIT_moneyAmount_(r.amount)

    const discountCtx = KPIAUDIT_buildDiscountContextNow_(r, {
      amountRaw: amountRaw,
      interval: interval,
      intervalCount: intervalCount,
      asOfDate: asOfNow
    })
    const kpiAmount = discountCtx.amount
    const kpiArr = KPIAUDIT_computeMrrArr_(kpiAmount, interval, intervalCount).arr

    // ── Path B: Column ARR (org_subscription_info amount) ──
    const infoAmount = KPIAUDIT_moneyAmount_(orgSubInfo.amount)
    const infoAmountYearly = KPIAUDIT_num_(orgSubInfo.amount_yearly)
    const infoInterval = KPIAUDIT_str_(orgSubInfo.interval).toLowerCase() || interval
    const infoIntervalCount = Math.max(1, KPIAUDIT_num_(orgSubInfo.interval_count) || intervalCount || 1)
    const columnArr = (infoAmountYearly > 0)
      ? infoAmountYearly
      : KPIAUDIT_computeMrrArr_(infoAmount, infoInterval, infoIntervalCount).arr

    // ── Delta ──
    const delta = kpiArr - columnArr
    const bucketLabel = ringBucket === 'paid' ? 'Paid'
      : ringBucket === 'intent_to_pay' ? 'Intent to Pay'
      : 'Free Trial'

    // Accumulate totals
    totalKpiArr += kpiArr
    totalColumnArr += columnArr
    if (ringBucket === 'paid') {
      paidKpiArr += kpiArr
      paidColumnArr += columnArr
    } else if (ringBucket === 'intent_to_pay') {
      // Ring KPI uses infoArr for intent_to_pay, not kpiArr
      intentKpiArr += columnArr
      intentColumnArr += columnArr
    } else {
      freeTrialKpiArr += columnArr
      freeTrialColumnArr += columnArr
    }

    detailRows.push([
      stripeSubId,
      KPIAUDIT_str_(r.customer_email || r.email),
      KPIAUDIT_str_(orgSubInfo.org_name),
      KPIAUDIT_str_(orgSubInfo.app_org_id),
      bucketLabel,
      statusRaw,
      interval,
      intervalCount,
      amountRaw,
      kpiAmount,
      kpiArr,
      infoAmount,
      infoAmountYearly || '',
      infoInterval,
      infoIntervalCount,
      columnArr,
      delta,
      discountCtx.discountPct > 0 ? (discountCtx.discountPct / 100) : '',
      discountCtx.promoCodes.join(', '),
      discountCtx.durationLabel,
      Math.abs(delta) < 0.01 ? 'MATCH' : 'DIFF'
    ])
  }

  // Sort: largest absolute delta first
  detailRows.sort((a, b) => Math.abs(b[16]) - Math.abs(a[16]))

  const matchCount = detailRows.filter(r => r[20] === 'MATCH').length
  const diffCount = detailRows.filter(r => r[20] === 'DIFF').length

  // ── Write output ──
  const shOut = KPIAUDIT_getOrCreateSheet_(ss, KPIAUDIT_CFG.OUT_SHEET)
  shOut.clear()

  // Section 1: Summary
  const summary = [
    ['ARR KPI vs Column Audit', '', 'Run at:', asOfNow],
    ['', '', '', ''],
    ['Bucket', 'KPI ARR (Path A)', 'Column ARR (Path B)', 'Delta'],
    ['Paid', paidKpiArr, paidColumnArr, paidKpiArr - paidColumnArr],
    ['Intent to Pay', intentKpiArr, intentColumnArr, intentKpiArr - intentColumnArr],
    ['Free Trial', freeTrialKpiArr, freeTrialColumnArr, freeTrialKpiArr - freeTrialColumnArr],
    ['TOTAL', totalKpiArr, totalColumnArr, totalKpiArr - totalColumnArr],
    ['', '', '', ''],
    ['Subscriptions compared:', detailRows.length, '', ''],
    ['Matching:', matchCount, '', ''],
    ['Different:', diffCount, '', ''],
    ['', '', '', '']
  ]

  shOut.getRange(1, 1, summary.length, 4).setValues(summary)

  // Section 2: Detail headers
  const detailHeaders = [
    'subscription_id',
    'customer_email',
    'org_name',
    'app_org_id',
    'ring_bucket',
    'status',
    'stripe_interval',
    'stripe_interval_count',
    'stripe_amount_raw',
    'stripe_amount_after_discount',
    'kpi_arr (Path A)',
    'info_amount',
    'info_amount_yearly',
    'info_interval',
    'info_interval_count',
    'column_arr (Path B)',
    'delta (A - B)',
    'active_discount_pct',
    'promo_codes',
    'discount_duration',
    'result'
  ]

  const detailStartRow = summary.length + 1
  shOut.getRange(detailStartRow, 1, 1, detailHeaders.length).setValues([detailHeaders])

  if (detailRows.length) {
    shOut.getRange(detailStartRow + 1, 1, detailRows.length, detailHeaders.length).setValues(detailRows)
  }

  // ── Formatting ──
  // Summary
  shOut.getRange(1, 1).setFontWeight('bold').setFontSize(14)
  shOut.getRange(3, 1, 1, 4).setFontWeight('bold').setBackground('#F3F3F3')
  shOut.getRange(7, 1, 1, 4).setFontWeight('bold')

  // Summary currency columns
  shOut.getRange(4, 2, 4, 3).setNumberFormat(KPIAUDIT_CFG.CURRENCY_FMT)

  // Detail header
  shOut.getRange(detailStartRow, 1, 1, detailHeaders.length)
    .setFontWeight('bold')
    .setBackground('#F3F3F3')
  shOut.setFrozenRows(detailStartRow)

  if (detailRows.length) {
    const dataStart = detailStartRow + 1
    const n = detailRows.length

    // Currency columns: stripe_amount_raw(9), stripe_amount_after_discount(10),
    // kpi_arr(11), info_amount(12), info_amount_yearly(13), column_arr(16), delta(17)
    ;[9, 10, 11, 12, 13, 16].forEach(col => {
      shOut.getRange(dataStart, col, n, 1).setNumberFormat(KPIAUDIT_CFG.CURRENCY_FMT)
    })
    shOut.getRange(dataStart, 17, n, 1).setNumberFormat(KPIAUDIT_CFG.DELTA_FMT)

    // Int columns: interval_count(8), info_interval_count(15)
    ;[8, 15].forEach(col => {
      shOut.getRange(dataStart, col, n, 1).setNumberFormat(KPIAUDIT_CFG.INT_FMT)
    })

    // Percent column: active_discount_pct(18)
    shOut.getRange(dataStart, 18, n, 1).setNumberFormat('0.##%')

    // Conditional color on result column
    for (let i = 0; i < n; i++) {
      const cell = shOut.getRange(dataStart + i, 21)
      if (detailRows[i][20] === 'DIFF') {
        cell.setBackground('#FCE4EC').setFontColor('#C62828')
      } else {
        cell.setBackground('#E8F5E9').setFontColor('#2E7D32')
      }
    }

    // Conditional color on delta column
    for (let i = 0; i < n; i++) {
      const deltaVal = detailRows[i][16]
      if (Math.abs(deltaVal) >= 0.01) {
        shOut.getRange(dataStart + i, 17).setBackground('#FFF3E0')
      }
    }
  }

  shOut.autoResizeColumns(1, detailHeaders.length)

  const seconds = (new Date() - t0) / 1000
  if (typeof writeSyncLog === 'function') {
    writeSyncLog('render_arr_kpi_vs_column_audit', 'ok', stripeRows.length, detailRows.length, seconds, '')
  }

  return { rows_in: stripeRows.length, rows_out: detailRows.length, matches: matchCount, diffs: diffCount }
}

/* =========================
 * Discount logic (mirrors Render The Ring.js exactly)
 * ========================= */

function KPIAUDIT_buildDiscountContextNow_(row, opts) {
  const cfg = opts || {}
  const amountRaw = KPIAUDIT_moneyAmount_(cfg.amountRaw)
  const asOfDate = (cfg.asOfDate instanceof Date && !isNaN(cfg.asOfDate.getTime()))
    ? cfg.asOfDate
    : new Date()

  const details = KPIAUDIT_parseDiscountDetails_(row)
  const active = details.filter(d => KPIAUDIT_isDiscountActiveNow_(d, asOfDate))

  let amount = amountRaw
  const promoCodes = []
  const durationLabels = []
  for (const d of active) {
    const pct = KPIAUDIT_num_(d.percent_off)
    if (pct > 0) {
      const bounded = Math.max(0, Math.min(100, pct))
      amount *= (1 - bounded / 100)
    }
    const amountOff = KPIAUDIT_num_(d.amount_off)
    if (amountOff > 0) amount -= amountOff

    if (KPIAUDIT_str_(d.promotion_code)) promoCodes.push(KPIAUDIT_str_(d.promotion_code))
    const label = KPIAUDIT_formatDiscountDuration_(d.duration, d.duration_in_months)
    if (label) durationLabels.push(label)
  }

  amount = Math.max(0, amount)
  const discountPct = amountRaw > 0 ? ((amountRaw - amount) / amountRaw) * 100 : 0

  return {
    amount,
    discountPct: Math.max(0, Math.min(100, discountPct)),
    promoCodes: Array.from(new Set(promoCodes)),
    durationLabel: Array.from(new Set(durationLabels)).join(', ')
  }
}

function KPIAUDIT_parseDiscountDetails_(row) {
  const out = []
  const r = row || {}

  const detailsJson = KPIAUDIT_str_(r.discount_details_json)
  if (detailsJson) {
    try {
      const parsed = JSON.parse(detailsJson)
      if (Array.isArray(parsed)) {
        parsed.forEach((d, i) => {
          if (!d || typeof d !== 'object') return
          out.push({
            index: i + 1,
            percent_off: KPIAUDIT_num_(d.percent_off),
            amount_off: KPIAUDIT_num_(d.amount_off),
            duration: KPIAUDIT_str_(d.duration).toLowerCase(),
            duration_in_months: KPIAUDIT_num_(d.duration_in_months),
            start_at: KPIAUDIT_str_(d.start_at),
            end_at: KPIAUDIT_str_(d.end_at),
            promotion_code: KPIAUDIT_str_(d.promotion_code)
          })
        })
      }
    } catch (e) {}
  }
  if (out.length) return out

  const pctAll = KPIAUDIT_csvList_(r.discount_percent_all)
  const amtAll = KPIAUDIT_csvList_(r.discount_amount_off_all)
  const durAll = KPIAUDIT_csvList_(r.discount_duration_all)
  const durMonthsAll = KPIAUDIT_csvList_(r.discount_duration_months_all)
  const startAll = KPIAUDIT_csvList_(r.discount_start_at_all)
  const endAll = KPIAUDIT_csvList_(r.discount_end_at_all)
  const promoAll = KPIAUDIT_csvList_(r.promo_code_all)
  const n = Math.max(
    pctAll.length, amtAll.length, durAll.length,
    durMonthsAll.length, startAll.length, endAll.length, promoAll.length
  )
  for (let i = 0; i < n; i++) {
    out.push({
      index: i + 1,
      percent_off: KPIAUDIT_num_(pctAll[i]),
      amount_off: KPIAUDIT_num_(amtAll[i]),
      duration: KPIAUDIT_str_(durAll[i]).toLowerCase(),
      duration_in_months: KPIAUDIT_num_(durMonthsAll[i]),
      start_at: KPIAUDIT_str_(startAll[i]),
      end_at: KPIAUDIT_str_(endAll[i]),
      promotion_code: KPIAUDIT_str_(promoAll[i])
    })
  }
  if (out.length) return out

  const pct = KPIAUDIT_num_(r.discount_percent)
  const amt = KPIAUDIT_num_(r.discount_amount_off)
  const duration = KPIAUDIT_str_(r.discount_duration).toLowerCase()
  const durationMonths = KPIAUDIT_num_(r.discount_duration_months)
  if (pct <= 0 && amt <= 0) return []

  return [{
    index: 1,
    percent_off: pct,
    amount_off: amt,
    duration,
    duration_in_months: durationMonths,
    start_at: KPIAUDIT_str_(r.discount_start_at || r.first_payment_at || r.created_at),
    end_at: KPIAUDIT_str_(r.discount_end_at),
    promotion_code: KPIAUDIT_str_(r.promo_code)
  }]
}

function KPIAUDIT_isDiscountActiveNow_(d, asOfDate) {
  const asOf = (asOfDate instanceof Date && !isNaN(asOfDate.getTime())) ? asOfDate : new Date()
  const start = KPIAUDIT_isoToDate_(d && d.start_at)
  const end = KPIAUDIT_isoToDate_(d && d.end_at)
  const duration = KPIAUDIT_str_(d && d.duration).toLowerCase()
  const hasValue = KPIAUDIT_num_(d && d.percent_off) > 0 || KPIAUDIT_num_(d && d.amount_off) > 0
  if (!hasValue) return false

  if (start && asOf < start) return false
  if (end) return asOf < end
  if (duration === 'forever') return true
  if (duration === 'repeating' || duration === 'once') {
    if (!start) return false
    const months = Math.max(1, KPIAUDIT_num_(d && d.duration_in_months) || 1)
    const until = new Date(start.getTime())
    until.setUTCMonth(until.getUTCMonth() + months)
    return asOf < until
  }
  return false
}

function KPIAUDIT_formatDiscountDuration_(duration, durationMonths) {
  const d = KPIAUDIT_str_(duration).toLowerCase().trim()
  if (d === 'forever') return 'forever'
  if (d === 'once') return 'once'
  if (d === 'repeating') {
    const n = Number(durationMonths)
    if (isFinite(n) && n > 0) return 'repeating ' + Math.floor(n) + ' mo'
    return 'repeating'
  }
  return ''
}

/* =========================
 * Ring bucket logic (mirrors Render The Ring.js)
 * ========================= */

function KPIAUDIT_ringBucket_(statusRaw, firstPaymentAt, hasPaymentMethod) {
  const status = KPIAUDIT_str_(statusRaw).toLowerCase()
  const hasFirstPayment = !!KPIAUDIT_str_(firstPaymentAt)
  const hasPM = !!hasPaymentMethod

  if (status === 'active' && hasFirstPayment) return 'paid'
  if (status === 'active' && !hasFirstPayment) return 'intent_to_pay'
  if (status === 'trialing' && hasPM) return 'intent_to_pay'
  if (status === 'trialing' && !hasPM) return 'free_trial'
  return ''
}

function KPIAUDIT_computeMrrArr_(amount, interval, intervalCount) {
  const amt = Number(amount || 0) || 0
  const intv = String(interval || '').toLowerCase().trim()
  const count = Math.max(1, Number(intervalCount || 1) || 1)
  if (intv === 'year' || intv === 'annual' || intv === 'yr') {
    const arr = amt / count
    return { arr: arr, mrr: arr / 12 }
  }
  if (intv === 'month' || intv === 'mo') {
    const arr = amt * (12 / count)
    return { mrr: arr / 12, arr: arr }
  }
  return { mrr: amt, arr: amt * 12 }
}

/* =========================
 * Manual changes
 * ========================= */

function KPIAUDIT_buildManualChanges_(sheet) {
  const out = new Map()
  if (!sheet) return out
  const rows = KPIAUDIT_readSheetObjects_(sheet, 1)
  for (const r of rows) {
    const subId =
      KPIAUDIT_str_(r.subscription_id) ||
      KPIAUDIT_str_(r.stripe_subscription_id) ||
      KPIAUDIT_str_(r.subscription)
    if (!subId) continue
    const reason = KPIAUDIT_str_(r.exclude_reason).toLowerCase()
    if (reason !== 'internal') continue
    out.set(subId, { excludeInternal: true })
  }
  return out
}

/* =========================
 * Helpers
 * ========================= */

function KPIAUDIT_readSheetObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1 || lastCol < 1) return []
  const header = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || '').trim() })
  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()
  return data.map(function(r) {
    var obj = {}
    header.forEach(function(h, i) {
      if (!h) return
      obj[KPIAUDIT_key_(h)] = r[i]
    })
    return obj
  })
}

function KPIAUDIT_key_(h) {
  return String(h || '').trim().toLowerCase().replace(/\s+/g, '_')
}

function KPIAUDIT_getOrCreateSheet_(ss, name) {
  if (typeof getOrCreateSheet === 'function') {
    try { return getOrCreateSheet(ss, name) } catch (e) {}
  }
  var sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function KPIAUDIT_str_(v) {
  if (v === null || v === undefined) return ''
  return String(v).trim()
}

function KPIAUDIT_num_(v) {
  if (v === null || v === undefined || v === '') return 0
  if (typeof v === 'number') return v
  var s = String(v).replace(/[^0-9.\-]/g, '').trim()
  var n = Number(s)
  return isNaN(n) ? 0 : n
}

function KPIAUDIT_moneyAmount_(raw) {
  if (raw === null || raw === undefined || raw === '') return 0
  var n = KPIAUDIT_num_(raw)
  if (!isFinite(n)) return 0
  return Math.round(n * 100) / 100
}

function KPIAUDIT_toBool_(v) {
  if (v === true) return true
  if (typeof v === 'number') return v === 1
  var s = String(v || '').trim().toLowerCase()
  return s === 'true' || s === '1' || s === 'yes' || s === 'y'
}

function KPIAUDIT_isoToDate_(iso) {
  var s = String(iso || '').trim()
  if (!s) return null
  var d = new Date(s)
  if (isNaN(d.getTime())) return null
  return d
}

function KPIAUDIT_csvList_(v) {
  var s = KPIAUDIT_str_(v)
  if (!s) return []
  return s.split(',').map(function(x) { return KPIAUDIT_str_(x) })
}
