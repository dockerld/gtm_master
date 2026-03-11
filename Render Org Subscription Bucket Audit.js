/**************************************************************
 * render_org_subscription_bucket_audit()
 *
 * Audits org_subscription_info bucket classification and writes:
 * - Summary counts for Paid / Promo Trial / Free Trial
 * - Full list table for each bucket
 * - Repeats table: entities appearing in more than one bucket
 **************************************************************/

const ORG_SUB_BUCKET_AUDIT_CFG = {
  SHEET_NAME: 'org_subscription_bucket_audit',
  INPUT_SHEET: 'org_subscription_info',
  STRIPE_SHEET: 'raw_stripe_subscriptions',
  MANUAL_CHANGES_SHEET: 'Manual Stripe Changes',
  DATE_FMT: 'yyyy-mm-dd hh:mm:ss',
  INT_FMT: '0',
  CURRENCY_FMT: '$#,##0.00'
}

function render_org_subscription_bucket_audit() {
  return ORGSUBA_lockWrap_('render_org_subscription_bucket_audit', () => {
    const t0 = new Date()
    try {
      const ss = SpreadsheetApp.getActive()
      const shIn = ss.getSheetByName(ORG_SUB_BUCKET_AUDIT_CFG.INPUT_SHEET)
      if (!shIn) throw new Error(`Missing input sheet: ${ORG_SUB_BUCKET_AUDIT_CFG.INPUT_SHEET}`)
      const shStripe = ss.getSheetByName(ORG_SUB_BUCKET_AUDIT_CFG.STRIPE_SHEET)
      if (!shStripe) throw new Error(`Missing input sheet: ${ORG_SUB_BUCKET_AUDIT_CFG.STRIPE_SHEET}`)
      const shManual = ss.getSheetByName(ORG_SUB_BUCKET_AUDIT_CFG.MANUAL_CHANGES_SHEET)

      const rows = ORGSUBA_readSheetObjects_(shIn, 1)
      const stripeRows = ORGSUBA_readSheetObjects_(shStripe, 1)
      const outSh = ORGSUBA_getOrCreateSheet_(ss, ORG_SUB_BUCKET_AUDIT_CFG.SHEET_NAME)
      const stripeBySubId = ORGSUBA_buildStripeBySubId_(stripeRows)
      const excludedSubIds = ORGSUBA_buildInternalExcludedSubIdSet_(shManual)
      const asOfNow = new Date()

      const buckets = {
        paid: [],
        intent_to_pay: [],
        free_trial: []
      }
      const orgGroups = new Map() // org_id -> Set(bucket)
      const subGroups = new Map() // sub_id -> Set(bucket)
      const orgNamesByOrg = new Map()

      for (const r of rows) {
        const subId = ORGSUBA_str_(r.latest_subscription_id || r.stripe_subscription_id)
        if (subId && excludedSubIds.has(subId)) continue

        const status = ORGSUBA_str_(r.status).toLowerCase()
        const hasFirstPayment = !!ORGSUBA_str_(r.first_payment_at)
        const hasPaymentMethod = ORGSUBA_toBool_(r.has_payment_method)
        const bucket = ORGSUBA_bucket_(status, hasFirstPayment, hasPaymentMethod)
        if (!bucket) continue

        const orgId = ORGSUBA_str_(r.app_org_id || r.org_id)
        const orgName = ORGSUBA_str_(r.org_name)
        const infoArr = ORGSUBA_num_(r.amount_yearly) > 0
          ? ORGSUBA_num_(r.amount_yearly)
          : ORGSUBA_annualize_(ORGSUBA_num_(r.amount), ORGSUBA_str_(r.interval), ORGSUBA_num_(r.interval_count))
        const stripeRow = subId ? (stripeBySubId.get(subId) || null) : null
        const paidArr = stripeRow
          ? ORGSUBA_computeDiscountedArrFromStripe_(stripeRow, asOfNow)
          : infoArr
        const arr = bucket === 'paid' ? paidArr : infoArr

        const reason = ORGSUBA_reason_(status, hasFirstPayment, hasPaymentMethod)

        buckets[bucket].push([
          orgName,
          orgId,
          subId,
          status,
          ORGSUBA_toDateOrBlank_(r.subscription_created_at_date || r.created_at),
          ORGSUBA_toDateOrBlank_(r.first_payment_at),
          hasPaymentMethod,
          ORGSUBA_numOrBlank_(r.trial_days_remaining),
          ORGSUBA_numOrBlank_(r.quantity_total),
          arr,
          reason
        ])

        if (orgId) {
          if (!orgGroups.has(orgId)) orgGroups.set(orgId, new Set())
          orgGroups.get(orgId).add(bucket)
          if (orgName && !orgNamesByOrg.has(orgId)) orgNamesByOrg.set(orgId, orgName)
        }
        if (subId) {
          if (!subGroups.has(subId)) subGroups.set(subId, new Set())
          subGroups.get(subId).add(bucket)
        }
      }

      const repeats = []
      orgGroups.forEach((set, orgId) => {
        const groups = Array.from(set)
        if (groups.length <= 1) return
        repeats.push([
          'org_id',
          orgId,
          orgNamesByOrg.get(orgId) || '',
          groups.join(', '),
          groups.length
        ])
      })
      subGroups.forEach((set, subId) => {
        const groups = Array.from(set)
        if (groups.length <= 1) return
        repeats.push([
          'subscription_id',
          subId,
          '',
          groups.join(', '),
          groups.length
        ])
      })
      repeats.sort((a, b) => String(a[0] || '').localeCompare(String(b[0] || '')) || String(a[1] || '').localeCompare(String(b[1] || '')))

      ORGSUBA_render_(outSh, buckets, repeats)

      const seconds = (new Date() - t0) / 1000
      ORGSUBA_writeSyncLog_(
        'render_org_subscription_bucket_audit',
        'ok',
        rows.length,
        buckets.paid.length + buckets.intent_to_pay.length + buckets.free_trial.length,
        seconds,
        ''
      )
      return {
        rows_in: rows.length,
        rows_out: buckets.paid.length + buckets.intent_to_pay.length + buckets.free_trial.length
      }
    } catch (err) {
      const seconds = (new Date() - t0) / 1000
      ORGSUBA_writeSyncLog_(
        'render_org_subscription_bucket_audit',
        'error',
        '',
        '',
        seconds,
        String(err && err.message ? err.message : err)
      )
      throw err
    }
  })
}

function ORGSUBA_render_(sheet, buckets, repeats) {
  sheet.clear()

  const headers = [
    'Org Name',
    'app_org_id',
    'latest_subscription_id',
    'status',
    'created_at',
    'first_payment_at',
    'has_payment_method',
    'trial_days_remaining',
    'quantity_total',
    'arr',
    'why_in_bucket'
  ]

  let row = 1
  sheet.getRange(row, 1, 1, 4).setValues([['Bucket', 'Subscriptions', 'Unique Orgs', 'ARR Total']]).setFontWeight('bold').setBackground('#F3F4F6')
  row += 1

  const sumRow = (name, list) => {
    const orgSet = new Set(list.map(r => ORGSUBA_str_(r[1])).filter(Boolean))
    const arrTotal = list.reduce((n, r) => n + (ORGSUBA_num_(r[9]) || 0), 0)
    return [name, list.length, orgSet.size, arrTotal]
  }
  const summaryRows = [
    sumRow('Paid', buckets.paid),
    sumRow('Intent to Pay', buckets.intent_to_pay),
    sumRow('Trialing', buckets.free_trial)
  ]
  sheet.getRange(row, 1, summaryRows.length, 4).setValues(summaryRows)
  sheet.getRange(row, 2, summaryRows.length, 2).setNumberFormat(ORG_SUB_BUCKET_AUDIT_CFG.INT_FMT)
  sheet.getRange(row, 4, summaryRows.length, 1).setNumberFormat(ORG_SUB_BUCKET_AUDIT_CFG.CURRENCY_FMT)
  row += summaryRows.length + 2

  row = ORGSUBA_writeSection_(sheet, row, 'Paid List')
  row = ORGSUBA_writeTable_(sheet, row, headers, buckets.paid)
  row += 1
  row = ORGSUBA_writeSection_(sheet, row, 'Intent to Pay List')
  row = ORGSUBA_writeTable_(sheet, row, headers, buckets.intent_to_pay)
  row += 1
  row = ORGSUBA_writeSection_(sheet, row, 'Trialing List')
  row = ORGSUBA_writeTable_(sheet, row, headers, buckets.free_trial)
  row += 1

  row = ORGSUBA_writeSection_(sheet, row, 'Repeated Across Lists')
  const repeatHeaders = ['entity_type', 'entity_id', 'org_name', 'groups', 'group_count']
  ORGSUBA_writeTable_(sheet, row, repeatHeaders, repeats)
}

function ORGSUBA_writeSection_(sheet, row, title) {
  sheet.getRange(row, 1, 1, 11).merge()
  sheet.getRange(row, 1).setValue(title).setFontWeight('bold').setBackground('#EEF2FF')
  return row + 1
}

function ORGSUBA_writeTable_(sheet, row, headers, rows) {
  const safeRows = (rows && rows.length) ? rows : [headers.map((_, i) => i === 0 ? '(none)' : '')]
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#F3F4F6')
  sheet.getRange(row + 1, 1, safeRows.length, headers.length).setValues(safeRows)

  if (headers.indexOf('created_at') >= 0) {
    sheet.getRange(row + 1, 5, safeRows.length, 2).setNumberFormat(ORG_SUB_BUCKET_AUDIT_CFG.DATE_FMT)
    sheet.getRange(row + 1, 8, safeRows.length, 2).setNumberFormat(ORG_SUB_BUCKET_AUDIT_CFG.INT_FMT)
    sheet.getRange(row + 1, 10, safeRows.length, 1).setNumberFormat(ORG_SUB_BUCKET_AUDIT_CFG.CURRENCY_FMT)
  }
  sheet.autoResizeColumns(1, headers.length)
  return row + 1 + safeRows.length
}

function ORGSUBA_bucket_(status, hasFirstPayment, hasPaymentMethod) {
  if (status === 'active' && hasFirstPayment) return 'paid'
  if (status === 'active' && !hasFirstPayment) return 'intent_to_pay'
  if (status === 'trialing' && hasPaymentMethod) return 'intent_to_pay'
  if (status === 'trialing' && !hasPaymentMethod) return 'free_trial'
  return ''
}

function ORGSUBA_reason_(status, hasFirstPayment, hasPaymentMethod) {
  if (status === 'active' && hasFirstPayment) return 'status=active and first_payment_at has value'
  if (status === 'active' && !hasFirstPayment) return 'status=active and first_payment_at is empty'
  if (status === 'trialing' && hasPaymentMethod) return 'status=trialing and has_payment_method=true'
  if (status === 'trialing' && !hasPaymentMethod) return 'status=trialing and has_payment_method=false'
  return ''
}

function ORGSUBA_annualize_(amount, interval, intervalCount) {
  const amt = ORGSUBA_num_(amount)
  const intv = ORGSUBA_str_(interval).toLowerCase()
  const count = Math.max(1, ORGSUBA_num_(intervalCount) || 1)
  if (intv === 'year' || intv === 'annual' || intv === 'yr') return amt / count
  if (intv === 'month' || intv === 'mo') return amt * (12 / count)
  return amt * 12
}

function ORGSUBA_buildStripeBySubId_(rows) {
  const out = new Map()
  ;(rows || []).forEach(r => {
    const subId =
      ORGSUBA_str_(r.stripe_subscription_id) ||
      ORGSUBA_str_(r.subscription_id) ||
      ORGSUBA_str_(r.subscription) ||
      ORGSUBA_str_(r.id)
    if (!subId || out.has(subId)) return
    out.set(subId, r)
  })
  return out
}

function ORGSUBA_buildInternalExcludedSubIdSet_(sheet) {
  const out = new Set()
  if (!sheet) return out
  const rows = ORGSUBA_readSheetObjects_(sheet, 1)
  ;(rows || []).forEach(r => {
    const reason = ORGSUBA_str_(r.exclude_reason).toLowerCase()
    if (reason !== 'internal') return
    const subId =
      ORGSUBA_str_(r.subscription_id) ||
      ORGSUBA_str_(r.stripe_subscription_id) ||
      ORGSUBA_str_(r.subscription)
    if (subId) out.add(subId)
  })
  return out
}

function ORGSUBA_computeDiscountedArrFromStripe_(row, asOfDate) {
  const interval = ORGSUBA_str_(row && row.interval).toLowerCase()
  const intervalCount = Math.max(1, ORGSUBA_num_(row && row.interval_count) || 1)
  const amountYearly = ORGSUBA_num_(row && row.amount_yearly)
  const amount = ORGSUBA_num_(row && row.amount)
  const baseAmount = amount > 0 ? amount : 0

  const baseArr = amountYearly > 0
    ? amountYearly
    : ORGSUBA_annualize_(baseAmount, interval, intervalCount)
  if (baseArr <= 0) return 0

  if (baseAmount > 0) {
    const effectiveAmount = ORGSUBA_applyActiveDiscountsToAmount_(baseAmount, row, asOfDate)
    return Math.max(0, ORGSUBA_annualize_(effectiveAmount, interval, intervalCount))
  }

  let out = baseArr
  const details = ORGSUBA_activeDiscountsFromRow_(row, asOfDate)
  for (const d of details) {
    const pct = ORGSUBA_num_(d.percent_off)
    if (pct > 0) out *= (1 - Math.max(0, Math.min(100, pct)) / 100)
    const amountOff = ORGSUBA_num_(d.amount_off)
    if (amountOff > 0) out -= ORGSUBA_annualize_(amountOff, interval, intervalCount)
  }
  return Math.max(0, out)
}

function ORGSUBA_applyActiveDiscountsToAmount_(baseAmount, row, asOfDate) {
  let out = ORGSUBA_num_(baseAmount)
  const details = ORGSUBA_activeDiscountsFromRow_(row, asOfDate)
  for (const d of details) {
    const pct = ORGSUBA_num_(d.percent_off)
    if (pct > 0) out *= (1 - Math.max(0, Math.min(100, pct)) / 100)
    const amountOff = ORGSUBA_num_(d.amount_off)
    if (amountOff > 0) out -= amountOff
    if (out <= 0) return 0
  }
  return Math.max(0, out)
}

function ORGSUBA_activeDiscountsFromRow_(row, asOfDate) {
  return ORGSUBA_parseDiscountDetailsFromRow_(row).filter(d => ORGSUBA_isDiscountActiveOnDate_(d, asOfDate))
}

function ORGSUBA_parseDiscountDetailsFromRow_(row) {
  const out = []
  if (!row) return out

  const jsonRaw = ORGSUBA_str_(row.discount_details_json)
  if (jsonRaw) {
    try {
      const arr = JSON.parse(jsonRaw)
      if (Array.isArray(arr)) {
        arr.forEach((d, i) => {
          if (!d || typeof d !== 'object') return
          out.push({
            index: i + 1,
            percent_off: ORGSUBA_num_(d.percent_off),
            amount_off: ORGSUBA_num_(d.amount_off),
            duration: ORGSUBA_str_(d.duration).toLowerCase(),
            duration_in_months: ORGSUBA_num_(d.duration_in_months),
            start_at: ORGSUBA_toIsoOrBlank_(d.start_at),
            end_at: ORGSUBA_toIsoOrBlank_(d.end_at)
          })
        })
      }
    } catch (_) {}
  }
  if (out.length) return out

  const pctAll = ORGSUBA_csvList_(row.discount_percent_all)
  const amtAll = ORGSUBA_csvList_(row.discount_amount_off_all)
  const durAll = ORGSUBA_csvList_(row.discount_duration_all)
  const durMonthsAll = ORGSUBA_csvList_(row.discount_duration_months_all)
  const startAll = ORGSUBA_csvList_(row.discount_start_at_all)
  const endAll = ORGSUBA_csvList_(row.discount_end_at_all)
  const n = Math.max(pctAll.length, amtAll.length, durAll.length, durMonthsAll.length, startAll.length, endAll.length)

  for (let i = 0; i < n; i++) {
    out.push({
      index: i + 1,
      percent_off: ORGSUBA_num_(pctAll[i]),
      amount_off: ORGSUBA_num_(amtAll[i]),
      duration: ORGSUBA_str_(durAll[i]).toLowerCase(),
      duration_in_months: ORGSUBA_num_(durMonthsAll[i]),
      start_at: ORGSUBA_toIsoOrBlank_(startAll[i]),
      end_at: ORGSUBA_toIsoOrBlank_(endAll[i])
    })
  }
  if (out.length) return out

  out.push({
    index: 1,
    percent_off: ORGSUBA_num_(row.discount_percent),
    amount_off: ORGSUBA_num_(row.discount_amount_off),
    duration: ORGSUBA_str_(row.discount_duration).toLowerCase(),
    duration_in_months: ORGSUBA_num_(row.discount_duration_months),
    start_at: ORGSUBA_toIsoOrBlank_(row.discount_start_at || row.first_payment_at || row.created_at),
    end_at: ORGSUBA_toIsoOrBlank_(row.discount_end_at)
  })
  return out
}

function ORGSUBA_isDiscountActiveOnDate_(detail, asOfDate) {
  const asOf = (asOfDate instanceof Date && !isNaN(asOfDate.getTime())) ? asOfDate : new Date()
  const d = detail || {}
  const duration = ORGSUBA_str_(d.duration).toLowerCase()
  const start = ORGSUBA_toDateOrBlank_(d.start_at)
  const end = ORGSUBA_toDateOrBlank_(d.end_at)

  if (start && asOf < start) return false
  if (end) return asOf < end
  if (duration === 'forever') return true
  if (duration === 'repeating' || duration === 'once') {
    if (!start) return false
    const months = Math.max(1, ORGSUBA_num_(d.duration_in_months) || 1)
    const until = new Date(start.getTime())
    until.setUTCMonth(until.getUTCMonth() + months)
    return asOf < until
  }
  return false
}

function ORGSUBA_csvList_(v) {
  const s = ORGSUBA_str_(v)
  if (!s) return []
  return s.split(',').map(x => ORGSUBA_str_(x))
}

function ORGSUBA_numOrBlank_(v) {
  if (v == null || v === '') return ''
  const n = Number(v)
  return isFinite(n) ? n : ''
}

function ORGSUBA_toBool_(v) {
  if (v == null || v === '') return false
  if (v === true || v === false) return v
  const s = ORGSUBA_str_(v).toLowerCase()
  if (s === 'true' || s === '1' || s === 'yes' || s === 'y') return true
  if (s === 'false' || s === '0' || s === 'no' || s === 'n') return false
  return false
}

function ORGSUBA_toDateOrBlank_(v) {
  if (v == null || v === '') return ''
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : v
  const d = new Date(String(v))
  return isNaN(d.getTime()) ? '' : d
}

function ORGSUBA_toIsoOrBlank_(v) {
  if (!v) return ''
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : v.toISOString()
  const s = String(v || '').trim()
  if (!s) return ''
  if (s.includes('T') && s.endsWith('Z')) return s
  if (/^\d+$/.test(s)) {
    const n = Number(s)
    const ms = n > 1e12 ? n : n * 1000
    const d = new Date(ms)
    return isNaN(d.getTime()) ? '' : d.toISOString()
  }
  const d = new Date(s)
  return isNaN(d.getTime()) ? '' : d.toISOString()
}

function ORGSUBA_readSheetObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1 || lastCol < 1) return []
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => ORGSUBA_str_(h))
  const rows = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()
  return rows.map(r => {
    const obj = {}
    headers.forEach((h, i) => { obj[h] = r[i] })
    return obj
  })
}

function ORGSUBA_getOrCreateSheet_(ss, name) {
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function ORGSUBA_lockWrap_(name, fn) {
  if (typeof lockWrap === 'function') {
    try { return lockWrap(name, fn) } catch (_) { return lockWrap(fn) }
  }
  return fn()
}

function ORGSUBA_writeSyncLog_(step, status, rowsIn, rowsOut, seconds, error) {
  if (typeof writeSyncLog === 'function') {
    return writeSyncLog(step, status, rowsIn, rowsOut, seconds, error || '')
  }
  Logger.log(`[SYNCLOG missing] ${step} ${status} rows_in=${rowsIn} rows_out=${rowsOut} seconds=${seconds} error=${error || ''}`)
}

function ORGSUBA_num_(v) {
  const n = Number(v)
  return isFinite(n) ? n : 0
}

function ORGSUBA_str_(v) {
  if (v == null) return ''
  return String(v).trim()
}
