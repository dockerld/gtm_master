/**************************************************************
 * render_org_subscription_info()
 *
 * Builds/overwrites "org_subscription_info" from:
 * - raw_posthog_orgs (base: all orgs)
 * - raw_stripe_subscriptions (joined by latest_subscription_id)
 *
 * Rule:
 * - If latest_subscription_id is blank, write:
 *   "null_id_invalid_org"
 **************************************************************/

const ORG_SUB_INFO_CFG = {
  SHEET_NAME: 'org_subscription_info',
  INPUTS: {
    POSTHOG_ORGS: 'raw_posthog_orgs',
    STRIPE_SUBS: 'raw_stripe_subscriptions',
    PROMO_REDEMPTIONS: 'promo_redemptions',
    POSTHOG_ORG_SUBS: 'raw_posthog_org_subscriptions'
  },
  MISSING_SUB_ID_LABEL: 'null_id_invalid_org',
  DATE_FMT: 'yyyy-mm-dd hh:mm:ss',
  INT_FMT: '0',
  CURRENCY_FMT: '$#,##0.00'
}

function render_org_subscription_info() {
  return ORGSUBINFO_lockWrap_('render_org_subscription_info', () => {
    const t0 = new Date()
    try {
      const ss = SpreadsheetApp.getActive()
      const shPosthogOrgs = ss.getSheetByName(ORG_SUB_INFO_CFG.INPUTS.POSTHOG_ORGS)
      const shStripe = ss.getSheetByName(ORG_SUB_INFO_CFG.INPUTS.STRIPE_SUBS)
      const shPromo = ss.getSheetByName(ORG_SUB_INFO_CFG.INPUTS.PROMO_REDEMPTIONS)
      const shPosthogOrgSubs = ss.getSheetByName(ORG_SUB_INFO_CFG.INPUTS.POSTHOG_ORG_SUBS)

      if (!shPosthogOrgs) throw new Error(`Missing input sheet: ${ORG_SUB_INFO_CFG.INPUTS.POSTHOG_ORGS}`)
      if (!shStripe) throw new Error(`Missing input sheet: ${ORG_SUB_INFO_CFG.INPUTS.STRIPE_SUBS}`)
      if (!shPromo) throw new Error(`Missing input sheet: ${ORG_SUB_INFO_CFG.INPUTS.PROMO_REDEMPTIONS}`)
      if (!shPosthogOrgSubs) throw new Error(`Missing input sheet: ${ORG_SUB_INFO_CFG.INPUTS.POSTHOG_ORG_SUBS}`)

      const posthogOrgs = ORGSUBINFO_readSheetObjects_(shPosthogOrgs, 1)
      const stripeRows = ORGSUBINFO_readSheetObjects_(shStripe, 1)
      const promoRows = ORGSUBINFO_readSheetObjects_(shPromo, 1)
      const posthogOrgSubsRows = ORGSUBINFO_readSheetObjects_(shPosthogOrgSubs, 1)

      const stripeBySubId = ORGSUBINFO_buildStripeBySubId_(stripeRows)
      const latestPromoBySubId = ORGSUBINFO_buildLatestPromoBySubId_(promoRows)
      const latestPromoByOrgId = ORGSUBINFO_buildLatestPromoByOrgId_(promoRows)
      const posthogOrgSubBySubId = ORGSUBINFO_buildPosthogOrgSubBySubId_(posthogOrgSubsRows)
      const orgs = ORGSUBINFO_dedupePosthogOrgs_(posthogOrgs)

      const headers = [
        'app_org_id',
        'org_name',
        'org_created_at',
        'latest_subscription_id',
        'stripe_subscription_id',
        'status',
        'subscription_created_at_date',
        'first_payment_at',
        'churn_date',
        'plan_name',
        'stripe_customer_id',
        'customer_email',
        'has_payment_method',
        'payment_method_created_at',
        'interval',
        'interval_count',
        'quantity_total',
        'unit_price',
        'amount',
        'unit_price_monthly',
        'unit_price_yearly',
        'amount_monthly',
        'amount_yearly',
        'trial_ends_at',
        'current_period_end',
        'trial_days_remaining',
        'promo_used',
        'last_promo_used',
        'redemption_location',
        'redeemed_at',
        'promo_code',
        'promo_name',
        'trial_days',
        'promo_type',
        'discount_percent',
        'discount_duration',
        'discount_duration_months',
        'discount_start_at',
        'discount_end_at',
        // ── Discount-adjusted amounts from Stripe sync ──
        'amount_after_discounts',
        'amount_after_discounts_monthly',
        'amount_after_discounts_yearly',
        'cancel_at_period_end',
        'churn_reason'
      ]

      const rowsOut = orgs.map(org => {
        const appOrgId = ORGSUBINFO_str_(org.app_org_id || org.org_id)
        const orgName = ORGSUBINFO_str_(org.org_name)
        const latestSubIdRaw = ORGSUBINFO_str_(org.latest_subscription_id)
        const latestSubId = latestSubIdRaw || ORG_SUB_INFO_CFG.MISSING_SUB_ID_LABEL
        const sub = latestSubIdRaw ? (stripeBySubId.get(latestSubIdRaw) || null) : null
        const orgSub = latestSubIdRaw ? (posthogOrgSubBySubId.get(latestSubIdRaw) || null) : null
        const promoBySub = latestSubIdRaw ? (latestPromoBySubId.get(latestSubIdRaw) || null) : null
        const promoByOrg = latestPromoByOrgId.get(appOrgId) || null
        const promo = promoBySub || promoByOrg || null
        const trialEndForRemaining =
          ORGSUBINFO_val_(orgSub, 'trial_ends_at') ||
          ORGSUBINFO_val_(orgSub, 'current_period_end') ||
          ''
        const trialDaysRemaining = ORGSUBINFO_computeTrialDaysRemaining_(trialEndForRemaining, new Date())
        const promoUsed = !!promo
        const subStatus = ORGSUBINFO_str_(ORGSUBINFO_val_(sub, 'status')).toLowerCase()
        const cancelAtPeriodEnd = ORGSUBINFO_toBoolOrBlank_(ORGSUBINFO_val_(sub, 'cancel_at_period_end'))
        const currentPeriodEndRaw =
          ORGSUBINFO_val_(sub, 'current_period_end') ||
          ORGSUBINFO_val_(orgSub, 'current_period_end') ||
          ''
        const churnDateRaw = (subStatus === 'canceled')
          ? (ORGSUBINFO_val_(sub, 'canceled_at') || currentPeriodEndRaw)
          : (cancelAtPeriodEnd ? currentPeriodEndRaw : '')
        const lastPromoUsed =
          ORGSUBINFO_str_(promo && promo.promo_code) ||
          ORGSUBINFO_str_(promo && promo.promo_name) ||
          ''

        return [
          appOrgId,
          orgName,
          ORGSUBINFO_toDateOrBlank_(
            ORGSUBINFO_val_(org, 'created_at') ||
            ORGSUBINFO_val_(org, 'org_created_at')
          ),
          latestSubId,
          latestSubIdRaw,
          ORGSUBINFO_val_(sub, 'status'),
          ORGSUBINFO_toDateOrBlank_(ORGSUBINFO_val_(sub, 'created_at')),
          ORGSUBINFO_toDateOrBlank_(ORGSUBINFO_val_(sub, 'first_payment_at')),
          ORGSUBINFO_toDateOrBlank_(churnDateRaw),
          ORGSUBINFO_pickPlanName_(sub),
          ORGSUBINFO_val_(sub, 'stripe_customer_id'),
          ORGSUBINFO_val_(sub, 'customer_email'),
          ORGSUBINFO_toBoolOrBlank_(ORGSUBINFO_val_(sub, 'has_payment_method')),
          ORGSUBINFO_toDateOrBlank_(ORGSUBINFO_val_(sub, 'payment_method_created_at')),
          ORGSUBINFO_val_(sub, 'interval'),
          ORGSUBINFO_numOrBlank_(ORGSUBINFO_val_(sub, 'interval_count')),
          ORGSUBINFO_numOrBlank_(ORGSUBINFO_val_(sub, 'quantity_total')),
          ORGSUBINFO_numOrBlank_(ORGSUBINFO_val_(sub, 'unit_price')),
          ORGSUBINFO_numOrBlank_(ORGSUBINFO_val_(sub, 'amount')),
          ORGSUBINFO_numOrBlank_(ORGSUBINFO_val_(sub, 'unit_price_monthly')),
          ORGSUBINFO_numOrBlank_(ORGSUBINFO_val_(sub, 'unit_price_yearly')),
          ORGSUBINFO_numOrBlank_(ORGSUBINFO_val_(sub, 'amount_monthly')),
          ORGSUBINFO_numOrBlank_(ORGSUBINFO_val_(sub, 'amount_yearly')),
          ORGSUBINFO_toDateOrBlank_(ORGSUBINFO_val_(orgSub, 'trial_ends_at')),
          ORGSUBINFO_toDateOrBlank_(currentPeriodEndRaw),
          trialDaysRemaining,
          promoUsed,
          lastPromoUsed,
          ORGSUBINFO_str_(promo && promo.redemption_location),
          ORGSUBINFO_toDateOrBlank_(promo && promo.redeemed_at),
          ORGSUBINFO_str_(promo && promo.promo_code),
          ORGSUBINFO_str_(promo && promo.promo_name),
          ORGSUBINFO_numOrBlank_(promo && promo.trial_days),
          ORGSUBINFO_str_(promo && promo.promo_type),
          ORGSUBINFO_numOrBlank_(promo && promo.discount_percent),
          ORGSUBINFO_str_(promo && promo.discount_duration),
          ORGSUBINFO_numOrBlank_(promo && promo.discount_duration_months),
          ORGSUBINFO_toDateOrBlank_(promo && promo.discount_start_at),
          ORGSUBINFO_toDateOrBlank_(promo && promo.discount_end_at),
          // ── Discount-adjusted amounts from Stripe sync ──
          ORGSUBINFO_numOrBlank_(ORGSUBINFO_val_(sub, 'amount_after_discounts')),
          ORGSUBINFO_numOrBlank_(ORGSUBINFO_val_(sub, 'amount_after_discounts_monthly')),
          ORGSUBINFO_numOrBlank_(ORGSUBINFO_val_(sub, 'amount_after_discounts_yearly')),
          cancelAtPeriodEnd,
          ORGSUBINFO_str_(ORGSUBINFO_val_(sub, 'cancellation_reason'))
        ]
      })

      const shOut = ORGSUBINFO_getOrCreateSheet_(ss, ORG_SUB_INFO_CFG.SHEET_NAME)
      shOut.clear()
      shOut.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#F3F4F6')

      if (rowsOut.length) {
        ORGSUBINFO_batchSetValues_(shOut, 2, 1, rowsOut, 2000)
      }

      ORGSUBINFO_applyFormats_(shOut, rowsOut.length)
      shOut.setFrozenRows(1)
      shOut.autoResizeColumns(1, headers.length)

      const seconds = (new Date() - t0) / 1000
      ORGSUBINFO_writeSyncLog_('render_org_subscription_info', 'ok', posthogOrgs.length, rowsOut.length, seconds, '')
      return { rows_in: posthogOrgs.length, rows_out: rowsOut.length }
    } catch (err) {
      const seconds = (new Date() - t0) / 1000
      ORGSUBINFO_writeSyncLog_('render_org_subscription_info', 'error', '', '', seconds, String(err && err.message ? err.message : err))
      throw err
    }
  })
}

function ORGSUBINFO_applyFormats_(sheet, rowCount) {
  if (!rowCount) return
  const startRow = 2

  // date columns
  ;[3, 7, 8, 9, 14, 24, 25, 30, 38, 39].forEach(col => {
    sheet.getRange(startRow, col, rowCount, 1).setNumberFormat(ORG_SUB_INFO_CFG.DATE_FMT)
  })

  // int columns
  ;[16, 17, 26, 33, 35, 37].forEach(col => {
    sheet.getRange(startRow, col, rowCount, 1).setNumberFormat(ORG_SUB_INFO_CFG.INT_FMT)
  })

  // currency columns (18-23 = gross amounts, 40-42 = discount-adjusted)
  ;[18, 19, 20, 21, 22, 23, 40, 41, 42].forEach(col => {
    sheet.getRange(startRow, col, rowCount, 1).setNumberFormat(ORG_SUB_INFO_CFG.CURRENCY_FMT)
  })
}

function ORGSUBINFO_dedupePosthogOrgs_(rows) {
  const byOrgId = new Map()
  for (const r of (rows || [])) {
    const appOrgId = ORGSUBINFO_str_(r.app_org_id || r.org_id)
    if (!appOrgId) continue
    const prev = byOrgId.get(appOrgId)
    if (!prev) {
      byOrgId.set(appOrgId, r)
      continue
    }

    const prevTs = ORGSUBINFO_ms_(prev.updated_at || prev.pulled_at || prev.created_at)
    const nextTs = ORGSUBINFO_ms_(r.updated_at || r.pulled_at || r.created_at)
    if (nextTs >= prevTs) byOrgId.set(appOrgId, r)
  }

  return Array.from(byOrgId.values()).sort((a, b) => {
    const aName = ORGSUBINFO_str_(a.org_name)
    const bName = ORGSUBINFO_str_(b.org_name)
    return aName.localeCompare(bName)
  })
}

function ORGSUBINFO_buildStripeBySubId_(rows) {
  const out = new Map()
  for (const r of (rows || [])) {
    const subId =
      ORGSUBINFO_str_(r.stripe_subscription_id) ||
      ORGSUBINFO_str_(r.subscription_id) ||
      ORGSUBINFO_str_(r.subscription) ||
      ORGSUBINFO_str_(r.id)
    if (!subId || out.has(subId)) continue
    out.set(subId, r)
  }
  return out
}

function ORGSUBINFO_buildLatestPromoBySubId_(rows) {
  const out = new Map()
  for (const r of (rows || [])) {
    const subId = ORGSUBINFO_str_(r.stripe_subscription_id)
    if (!subId) continue

    const prev = out.get(subId)
    if (!prev) {
      out.set(subId, r)
      continue
    }

    const prevTs = ORGSUBINFO_promoSortMs_(prev)
    const curTs = ORGSUBINFO_promoSortMs_(r)
    if (curTs >= prevTs) out.set(subId, r)
  }
  return out
}

function ORGSUBINFO_buildPosthogOrgSubBySubId_(rows) {
  const out = new Map()
  for (const r of (rows || [])) {
    const subId =
      ORGSUBINFO_str_(r.stripe_subscription_id) ||
      ORGSUBINFO_str_(r.subscription_id) ||
      ORGSUBINFO_str_(r.id)
    if (!subId || out.has(subId)) continue
    out.set(subId, r)
  }
  return out
}

function ORGSUBINFO_buildLatestPromoByOrgId_(rows) {
  const out = new Map()
  for (const r of (rows || [])) {
    const orgId = ORGSUBINFO_str_(r.app_org_id || r.org_id)
    if (!orgId) continue

    const prev = out.get(orgId)
    if (!prev) {
      out.set(orgId, r)
      continue
    }

    const prevTs = ORGSUBINFO_promoSortMs_(prev)
    const curTs = ORGSUBINFO_promoSortMs_(r)
    if (curTs >= prevTs) out.set(orgId, r)
  }
  return out
}

function ORGSUBINFO_promoSortMs_(r) {
  const redeemedTs = ORGSUBINFO_ms_(r && r.redeemed_at)
  if (redeemedTs > 0) return redeemedTs
  const discountStartTs = ORGSUBINFO_ms_(r && r.discount_start_at)
  if (discountStartTs > 0) return discountStartTs
  const pulledTs = ORGSUBINFO_ms_(r && r.pulled_at)
  return pulledTs > 0 ? pulledTs : 0
}

function ORGSUBINFO_pickPlanName_(sub) {
  const s = sub || {}
  return (
    ORGSUBINFO_str_(s.current_plan) ||
    ORGSUBINFO_str_(s.subscription_tier) ||
    ORGSUBINFO_str_(s.plan_name) ||
    ORGSUBINFO_str_(s.product_name) ||
    ''
  )
}

function ORGSUBINFO_computeTrialDaysRemaining_(trialEndRaw, asOfDate) {
  const end = ORGSUBINFO_toDateOrBlank_(trialEndRaw)
  if (!(end instanceof Date) || isNaN(end.getTime())) return ''
  const now = (asOfDate instanceof Date && !isNaN(asOfDate.getTime())) ? asOfDate : new Date()
  const msPerDay = 24 * 60 * 60 * 1000
  const days = Math.ceil((end.getTime() - now.getTime()) / msPerDay)
  return days > 0 ? days : 0
}

function ORGSUBINFO_val_(obj, key) {
  if (!obj) return ''
  return obj[key]
}

function ORGSUBINFO_numOrBlank_(v) {
  if (v == null || v === '') return ''
  const n = Number(v)
  return isFinite(n) ? n : ''
}

function ORGSUBINFO_toBoolOrBlank_(v) {
  if (v == null || v === '') return ''
  if (v === true || v === false) return v
  const s = ORGSUBINFO_str_(v).toLowerCase()
  if (s === 'true' || s === '1' || s === 'yes' || s === 'y') return true
  if (s === 'false' || s === '0' || s === 'no' || s === 'n') return false
  return ''
}

function ORGSUBINFO_toDateOrBlank_(v) {
  if (v == null || v === '') return ''
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : v
  const d = new Date(String(v))
  return isNaN(d.getTime()) ? '' : d
}

function ORGSUBINFO_ms_(v) {
  const d = ORGSUBINFO_toDateOrBlank_(v)
  return d instanceof Date ? d.getTime() : 0
}

function ORGSUBINFO_readSheetObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1 || lastCol < 1) return []
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => ORGSUBINFO_str_(h))
  const rows = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()
  return rows.map(r => {
    const obj = {}
    headers.forEach((h, i) => { obj[h] = r[i] })
    return obj
  })
}

function ORGSUBINFO_batchSetValues_(sheet, startRow, startCol, values, chunkSize) {
  const size = Math.max(1, Number(chunkSize || 1000))
  for (let i = 0; i < values.length; i += size) {
    const chunk = values.slice(i, i + size)
    sheet.getRange(startRow + i, startCol, chunk.length, chunk[0].length).setValues(chunk)
  }
}

function ORGSUBINFO_getOrCreateSheet_(ss, name) {
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function ORGSUBINFO_lockWrap_(name, fn) {
  if (typeof lockWrap === 'function') {
    try { return lockWrap(name, fn) } catch (_) { return lockWrap(fn) }
  }
  return fn()
}

function ORGSUBINFO_writeSyncLog_(step, status, rowsIn, rowsOut, seconds, error) {
  if (typeof writeSyncLog === 'function') {
    return writeSyncLog(step, status, rowsIn, rowsOut, seconds, error || '')
  }
  Logger.log(`[SYNCLOG missing] ${step} ${status} rows_in=${rowsIn} rows_out=${rowsOut} seconds=${seconds} error=${error || ''}`)
}

function ORGSUBINFO_str_(v) {
  if (v == null) return ''
  return String(v).trim()
}
