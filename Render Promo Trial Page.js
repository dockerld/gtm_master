/**************************************************************
 * render_promo_trial_page()
 *
 * Builds one sheet with:
 * - Promo Trial subscriptions + reason
 * - Free Trial subscriptions + reason
 * - Counts for each list
 **************************************************************/

const PROMO_TRIAL_PAGE_CFG = {
  SHEET_NAME: 'Promo Trial Page',
  INPUTS: {
    STRIPE_SUBS: 'raw_stripe_subscriptions',
    CANON_ORGS: 'canon_orgs',
    POSTHOG_ORG_SUBS: 'raw_posthog_org_subscriptions',
    PROMO_REDEMPTIONS: 'promo_redemptions',
    MANUAL_CHANGES: 'Manual Stripe Changes'
  },
  CURRENCY_FMT: '$#,##0.00',
  INT_FMT: '0',
  DATETIME_FMT: 'yyyy-mm-dd hh:mm:ss'
}

function render_promo_trial_page() {
  return PROMOTRIAL_lockWrap_('render_promo_trial_page', () => {
    const t0 = new Date()
    try {
      const ss = SpreadsheetApp.getActive()
      const shOut = PROMOTRIAL_getOrCreateSheet_(ss, PROMO_TRIAL_PAGE_CFG.SHEET_NAME)

      const shStripe = ss.getSheetByName(PROMO_TRIAL_PAGE_CFG.INPUTS.STRIPE_SUBS)
      if (!shStripe) throw new Error(`Missing input sheet: ${PROMO_TRIAL_PAGE_CFG.INPUTS.STRIPE_SUBS}`)

      const shCanon = ss.getSheetByName(PROMO_TRIAL_PAGE_CFG.INPUTS.CANON_ORGS)
      const shOrgSubs = ss.getSheetByName(PROMO_TRIAL_PAGE_CFG.INPUTS.POSTHOG_ORG_SUBS)
      const shPromo = ss.getSheetByName(PROMO_TRIAL_PAGE_CFG.INPUTS.PROMO_REDEMPTIONS)
      const shManual = ss.getSheetByName(PROMO_TRIAL_PAGE_CFG.INPUTS.MANUAL_CHANGES)

      const stripeRows = PROMOTRIAL_readSheetObjects_(shStripe, 1)
      const canonRows = shCanon ? PROMOTRIAL_readSheetObjects_(shCanon, 1) : []
      const orgSubRows = shOrgSubs ? PROMOTRIAL_readSheetObjects_(shOrgSubs, 1) : []
      const promoRows = shPromo ? PROMOTRIAL_readSheetObjects_(shPromo, 1) : []

      const excludedSubIds = PROMOTRIAL_buildInternalExcludeSubIdSet_(shManual)
      const canonMap = PROMOTRIAL_buildCanonMaps_(canonRows)
      const orgSubMap = PROMOTRIAL_buildOrgSubByStripeSubId_(orgSubRows)
      const appPromoByOrgId = PROMOTRIAL_buildAppPromoByOrgId_(promoRows)

      const promoTrial = []
      const freeTrial = []
      const now = new Date()

      for (const r of stripeRows) {
        const subId =
          PROMOTRIAL_str_(r.stripe_subscription_id) ||
          PROMOTRIAL_str_(r.subscription_id) ||
          PROMOTRIAL_str_(r.subscription) ||
          PROMOTRIAL_str_(r.id)
        if (!subId) continue
        if (excludedSubIds.has(subId)) continue

        const status = PROMOTRIAL_str_(r.status).toLowerCase()
        if (status !== 'active' && status !== 'trialing') continue

        const firstPaymentAt = PROMOTRIAL_str_(r.first_payment_at)
        const hasFirstPayment = !!firstPaymentAt

        const orgSub = orgSubMap.get(subId) || null
        const orgId =
          (canonMap.orgIdBySubId.get(subId) || '') ||
          PROMOTRIAL_str_(orgSub && (orgSub.app_org_id || orgSub.org_id)) ||
          PROMOTRIAL_str_(r.org_id || r.organization_id)
        const orgName =
          (canonMap.orgNameByOrgId.get(orgId) || '') ||
          PROMOTRIAL_str_(r.org_name || r.organization_name || r.org)
        const orgSignUpDate = PROMOTRIAL_toDateOrBlank_(canonMap.orgCreatedAtByOrgId.get(orgId))

        const stripePromoCodes = PROMOTRIAL_collectStripePromoCodes_(r)
        const appPromoCodes = orgId ? (appPromoByOrgId.get(orgId) || []) : []
        const hasPromoUsage = stripePromoCodes.length > 0 || appPromoCodes.length > 0

        const trialDaysRemaining = PROMOTRIAL_computeTrialDaysRemaining_(orgSub, now)
        const seats = PROMOTRIAL_safeInt_(r.quantity_total)
        const arr = PROMOTRIAL_computeArrFromRow_(r)

        const base = [
          orgName,
          orgId,
          orgSignUpDate,
          subId,
          status,
          firstPaymentAt ? PROMOTRIAL_toDateOrBlank_(firstPaymentAt) : '',
          trialDaysRemaining,
          seats,
          arr,
          stripePromoCodes.join(', '),
          appPromoCodes.join(', ')
        ]

        if (status === 'active' && !hasFirstPayment) {
          promoTrial.push(base.concat(['Active with no first payment date (treated as Promo Trial)']))
          continue
        }

        if (status === 'trialing' && !hasFirstPayment && hasPromoUsage) {
          promoTrial.push(base.concat(['Trialing, no first payment, and promo usage (Stripe and/or in-app)']))
          continue
        }

        if (status === 'trialing' && !hasFirstPayment && !hasPromoUsage) {
          freeTrial.push(base.concat(['Trialing, no first payment, and no promo usage']))
        }
      }

      promoTrial.sort((a, b) => String(a[0] || '').localeCompare(String(b[0] || '')))
      freeTrial.sort((a, b) => String(a[0] || '').localeCompare(String(b[0] || '')))

      PROMOTRIAL_renderSheet_(shOut, promoTrial, freeTrial)

      const seconds = (new Date() - t0) / 1000
      PROMOTRIAL_writeSyncLog_('render_promo_trial_page', 'ok', stripeRows.length, promoTrial.length + freeTrial.length, seconds, '')
      return { rows_in: stripeRows.length, rows_out: promoTrial.length + freeTrial.length }
    } catch (err) {
      const seconds = (new Date() - t0) / 1000
      PROMOTRIAL_writeSyncLog_('render_promo_trial_page', 'error', '', '', seconds, String(err && err.message ? err.message : err))
      throw err
    }
  })
}

function PROMOTRIAL_renderSheet_(sheet, promoRows, freeRows) {
  sheet.clear()

  let row = 1
  sheet.getRange(row, 1, 1, 3).setValues([['List', 'Subscriptions', 'Unique Orgs']]).setFontWeight('bold').setBackground('#F3F4F6')
  row += 1
  const promoUniqueOrgs = new Set((promoRows || []).map(r => PROMOTRIAL_str_(r[1])).filter(Boolean)).size
  const freeUniqueOrgs = new Set((freeRows || []).map(r => PROMOTRIAL_str_(r[1])).filter(Boolean)).size
  sheet.getRange(row, 1, 2, 2).setValues([
    ['Promo Trial subscriptions', promoRows.length],
    ['Free Trial subscriptions', freeRows.length]
  ])
  sheet.getRange(row, 3, 2, 1).setValues([[promoUniqueOrgs], [freeUniqueOrgs]])
  sheet.getRange(row, 2, 2, 2).setNumberFormat(PROMO_TRIAL_PAGE_CFG.INT_FMT)
  row += 3

  const headers = [
    'Org Name',
    'Org ID',
    'Org Sign Up Date',
    'Subscription ID',
    'Status',
    'First Payment At',
    'Trial Days Remaining',
    'Seats',
    'ARR',
    'Stripe Promo Codes',
    'In-App Promo Codes',
    'Why On List'
  ]

  row = PROMOTRIAL_writeSection_(sheet, row, 'Promo Trial Subscriptions')
  row = PROMOTRIAL_writeTable_(sheet, row, headers, promoRows)
  row += 1
  row = PROMOTRIAL_writeSection_(sheet, row, 'Free Trial Subscriptions')
  PROMOTRIAL_writeTable_(sheet, row, headers, freeRows)
}

function PROMOTRIAL_writeSection_(sheet, row, title) {
  sheet.getRange(row, 1, 1, 11).merge()
  sheet.getRange(row, 1).setValue(title).setFontWeight('bold').setBackground('#EEF2FF')
  return row + 1
}

function PROMOTRIAL_writeTable_(sheet, row, headers, rows) {
  const safeRows = rows && rows.length ? rows : [headers.map((_, i) => i === 0 ? '(none)' : '')]
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#F3F4F6')
  sheet.getRange(row + 1, 1, safeRows.length, headers.length).setValues(safeRows)

  sheet.getRange(row + 1, 3, safeRows.length, 1).setNumberFormat(PROMO_TRIAL_PAGE_CFG.DATETIME_FMT)
  sheet.getRange(row + 1, 6, safeRows.length, 1).setNumberFormat(PROMO_TRIAL_PAGE_CFG.DATETIME_FMT)
  sheet.getRange(row + 1, 7, safeRows.length, 1).setNumberFormat(PROMO_TRIAL_PAGE_CFG.INT_FMT)
  sheet.getRange(row + 1, 8, safeRows.length, 1).setNumberFormat(PROMO_TRIAL_PAGE_CFG.INT_FMT)
  sheet.getRange(row + 1, 9, safeRows.length, 1).setNumberFormat(PROMO_TRIAL_PAGE_CFG.CURRENCY_FMT)

  sheet.autoResizeColumns(1, headers.length)
  return row + 1 + safeRows.length
}

function PROMOTRIAL_buildCanonMaps_(rows) {
  const orgIdBySubId = new Map()
  const orgNameByOrgId = new Map()
  const orgCreatedAtByOrgId = new Map()

  for (const r of (rows || [])) {
    const orgId = PROMOTRIAL_str_(r.app_org_id || r.org_id || r.clerk_org_id)
    const orgName = PROMOTRIAL_str_(r.org_name || r.posthog_org_name || r.org_slug)
    const orgCreatedAt = PROMOTRIAL_str_(r.org_created_at || r.posthog_org_created_at || r.created_at)
    if (orgId && orgName && !orgNameByOrgId.has(orgId)) orgNameByOrgId.set(orgId, orgName)
    if (orgId && orgCreatedAt && !orgCreatedAtByOrgId.has(orgId)) orgCreatedAtByOrgId.set(orgId, orgCreatedAt)

    const subIds = PROMOTRIAL_csvList_(r.stripe_subscription_ids)
    subIds.forEach(subId => {
      if (!subId || !orgId || orgIdBySubId.has(subId)) return
      orgIdBySubId.set(subId, orgId)
    })
  }

  return { orgIdBySubId, orgNameByOrgId, orgCreatedAtByOrgId }
}

function PROMOTRIAL_buildOrgSubByStripeSubId_(rows) {
  const out = new Map()
  for (const r of (rows || [])) {
    const subId =
      PROMOTRIAL_str_(r.stripe_subscription_id) ||
      PROMOTRIAL_str_(r.subscription_id) ||
      PROMOTRIAL_str_(r.subscription) ||
      PROMOTRIAL_str_(r.id)
    if (!subId || out.has(subId)) continue
    out.set(subId, r)
  }
  return out
}

function PROMOTRIAL_buildAppPromoByOrgId_(rows) {
  const sets = new Map()
  for (const r of (rows || [])) {
    const orgId = PROMOTRIAL_str_(r.app_org_id || r.org_id)
    if (!orgId) continue
    const code =
      PROMOTRIAL_str_(r.promo_code) ||
      PROMOTRIAL_str_(r.code) ||
      PROMOTRIAL_str_(r.promo_name) ||
      PROMOTRIAL_str_(r.name)
    if (!code) continue
    if (!sets.has(orgId)) sets.set(orgId, new Set())
    sets.get(orgId).add(code)
  }

  const out = new Map()
  sets.forEach((set, orgId) => out.set(orgId, Array.from(set)))
  return out
}

function PROMOTRIAL_collectStripePromoCodes_(r) {
  const out = new Set()
  PROMOTRIAL_csvList_(r.promo_code_all).forEach(c => out.add(c))
  const single = PROMOTRIAL_str_(r.promo_code)
  if (single) out.add(single)
  return Array.from(out)
}

function PROMOTRIAL_computeTrialDaysRemaining_(orgSubRow, now) {
  if (!orgSubRow) return ''
  const trialEndRaw =
    PROMOTRIAL_str_(orgSubRow.trial_ends_at) ||
    PROMOTRIAL_str_(orgSubRow.current_period_end)
  const end = PROMOTRIAL_toDateOrNull_(trialEndRaw)
  if (!end) return ''
  const msPerDay = 24 * 60 * 60 * 1000
  const days = Math.ceil((end.getTime() - now.getTime()) / msPerDay)
  return days > 0 ? days : 0
}

function PROMOTRIAL_computeArrFromRow_(r) {
  const amount = Number(r.amount || 0) || 0
  const interval = PROMOTRIAL_str_(r.interval).toLowerCase()
  const intervalCount = Math.max(1, Number(r.interval_count || 1) || 1)
  if (interval === 'year' || interval === 'annual' || interval === 'yr') return amount / intervalCount
  if (interval === 'month' || interval === 'mo') return amount * (12 / intervalCount)
  return amount * 12
}

function PROMOTRIAL_buildInternalExcludeSubIdSet_(sheetMaybe) {
  const out = new Set()
  if (!sheetMaybe) return out
  const rows = PROMOTRIAL_readSheetObjects_(sheetMaybe, 1)
  rows.forEach(r => {
    const reason = PROMOTRIAL_str_(r.exclude_reason).toLowerCase()
    if (reason !== 'internal') return
    const subId =
      PROMOTRIAL_str_(r.subscription_id) ||
      PROMOTRIAL_str_(r.stripe_subscription_id) ||
      PROMOTRIAL_str_(r.subscription)
    if (subId) out.add(subId)
  })
  return out
}

function PROMOTRIAL_readSheetObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1 || lastCol < 1) return []
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => PROMOTRIAL_str_(h))
  const rows = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()
  return rows.map(r => {
    const obj = {}
    headers.forEach((h, i) => { obj[h] = r[i] })
    return obj
  })
}

function PROMOTRIAL_getOrCreateSheet_(ss, name) {
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function PROMOTRIAL_lockWrap_(name, fn) {
  if (typeof lockWrap === 'function') {
    try { return lockWrap(name, fn) } catch (_) { return lockWrap(fn) }
  }
  return fn()
}

function PROMOTRIAL_writeSyncLog_(step, status, rowsIn, rowsOut, seconds, error) {
  if (typeof writeSyncLog === 'function') {
    return writeSyncLog(step, status, rowsIn, rowsOut, seconds, error || '')
  }
  Logger.log(`[SYNCLOG missing] ${step} ${status} rows_in=${rowsIn} rows_out=${rowsOut} seconds=${seconds} error=${error || ''}`)
}

function PROMOTRIAL_safeInt_(v) {
  const n = Number(v)
  return isFinite(n) ? Math.max(0, Math.round(n)) : 0
}

function PROMOTRIAL_csvList_(v) {
  const s = PROMOTRIAL_str_(v)
  if (!s) return []
  return s.split(',').map(x => PROMOTRIAL_str_(x)).filter(Boolean)
}

function PROMOTRIAL_toDateOrNull_(v) {
  if (!v && v !== 0) return null
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v
  const s = PROMOTRIAL_str_(v)
  if (!s) return null
  const d = new Date(s)
  return isNaN(d.getTime()) ? null : d
}

function PROMOTRIAL_toDateOrBlank_(v) {
  const d = PROMOTRIAL_toDateOrNull_(v)
  return d || ''
}

function PROMOTRIAL_str_(v) {
  if (v == null) return ''
  return String(v).trim()
}
