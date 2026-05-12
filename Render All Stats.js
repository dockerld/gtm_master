/**************************************************************
 * render_all_stats_view()
 *
 * Builds/overwrites the "All the Stats" sheet with org-level
 * metrics and lists derived from Stripe + Clerk + ARR snapshots.
 **************************************************************/

const ALL_STATS_CFG = {
  SHEET_NAME: 'All the Stats',
  INPUTS: {
    STRIPE_SUBS: 'raw_stripe_subscriptions',
    ORG_SUBSCRIPTIONS: 'org_subscription_info',
    POSTHOG_PROMO_REDEMPTIONS: 'promo_redemptions',
    MANUAL_CHANGES: 'Manual Stripe Changes',
    CANON_ORGS: 'canon_orgs',
    CLERK_USERS: 'raw_clerk_users',
    CLERK_MEMBERSHIPS: 'raw_clerk_memberships',
    CLERK_ORGS: 'raw_clerk_orgs',
    POSTHOG_USERS: 'raw_posthog_user_metrics',
    POSTHOG_ORG_SUBSCRIPTIONS: 'raw_posthog_org_subscriptions',
    ARR_SNAPSHOT: 'arr_snapshot',
    ARR_WATERFALL_FACTS: 'arr_waterfall_facts'
  },
  CURRENCY_FMT: '$#,##0.00',
  INT_FMT: '0',
  DECIMAL_FMT: '0.00',
  PCT_FMT: '0.0%',
  DATETIME_FMT: 'yyyy-mm-dd hh:mm:ss',
  DATE_FMT: 'yyyy-mm-dd',
  TITLE_BG: '#1F2937',
  TITLE_FG: '#FFFFFF',
  SECTION_BG: '#EEF2FF',
  HEADER_BG: '#F3F4F6'
}

function render_all_stats_view() {
  return ALLSTATS_lockWrap_('render_all_stats_view', () => {
    const t0 = new Date()

    const ss = SpreadsheetApp.getActive()
    const out = ALLSTATS_getOrCreateSheet_(ss, ALL_STATS_CFG.SHEET_NAME)

    const shStripe = ss.getSheetByName(ALL_STATS_CFG.INPUTS.STRIPE_SUBS)
    if (!shStripe) throw new Error('Missing input sheet: raw_stripe_subscriptions')
    const shOrgSubscriptions = ss.getSheetByName(ALL_STATS_CFG.INPUTS.ORG_SUBSCRIPTIONS)
    if (!shOrgSubscriptions) throw new Error('Missing input sheet: org_subscription_info')

    const shPromoRedemptions = ss.getSheetByName(ALL_STATS_CFG.INPUTS.POSTHOG_PROMO_REDEMPTIONS)
    const shManual = ss.getSheetByName(ALL_STATS_CFG.INPUTS.MANUAL_CHANGES)
    const shCanonOrgs = ss.getSheetByName(ALL_STATS_CFG.INPUTS.CANON_ORGS)
    if (!shCanonOrgs) throw new Error('Missing input sheet: canon_orgs')
    const shUsers = ss.getSheetByName(ALL_STATS_CFG.INPUTS.CLERK_USERS)
    const shMems = ss.getSheetByName(ALL_STATS_CFG.INPUTS.CLERK_MEMBERSHIPS)
    const shOrgs = ss.getSheetByName(ALL_STATS_CFG.INPUTS.CLERK_ORGS)
    const shPosthog = ss.getSheetByName(ALL_STATS_CFG.INPUTS.POSTHOG_USERS)
    const shPosthogOrgSubs = ss.getSheetByName(ALL_STATS_CFG.INPUTS.POSTHOG_ORG_SUBSCRIPTIONS)
    const shSnap = ss.getSheetByName(ALL_STATS_CFG.INPUTS.ARR_SNAPSHOT)
    const shWaterfall = ss.getSheetByName(ALL_STATS_CFG.INPUTS.ARR_WATERFALL_FACTS)

    const stripeRows = ALLSTATS_readSheetObjects_(shStripe, 1)
    const orgSubscriptions = ALLSTATS_loadOrgSubscriptions_(shOrgSubscriptions)
    const orgSubsIndex = ALLSTATS_buildOrgSubscriptionsIndex_(orgSubscriptions)
    const stripeBySubId = ALLSTATS_buildStripeBySubId_(stripeRows)
    const manualBySubId = ALLSTATS_buildManualStripeChangesBySubId_(shManual)
    const promoRedemptions = ALLSTATS_loadPromoRedemptions_(shPromoRedemptions)
    const canonOrgs = ALLSTATS_readSheetObjects_(shCanonOrgs, 1)

    const clerkUsers = shUsers ? ALLSTATS_readSheetObjects_(shUsers, 1) : []
    const clerkMems = shMems ? ALLSTATS_readSheetObjects_(shMems, 1) : []
    const clerkOrgs = shOrgs ? ALLSTATS_readSheetObjects_(shOrgs, 1) : []
    const posthogUsers = shPosthog ? ALLSTATS_readSheetObjects_(shPosthog, 1) : []
    const posthogOrgSubsRows = shPosthogOrgSubs ? ALLSTATS_readSheetObjects_(shPosthogOrgSubs, 1) : []

    const indexes = ALLSTATS_buildIndexes_(clerkUsers, clerkMems, clerkOrgs, posthogUsers, canonOrgs)
    const orgAggByKey = new Map()

    for (const row of stripeRows) {
      const sub = ALLSTATS_normalizeSubscription_(row, manualBySubId, indexes, orgSubsIndex)
      if (!sub.include) continue

      const orgKey = ALLSTATS_orgKey_(sub)
      if (!orgAggByKey.has(orgKey)) {
        orgAggByKey.set(orgKey, ALLSTATS_newOrgAgg_(sub, orgKey))
      }
      ALLSTATS_addSubToOrgAgg_(orgAggByKey.get(orgKey), sub)
    }

    const orgs = Array.from(orgAggByKey.values()).map(ALLSTATS_finalizeOrgAgg_)
    const promoRedemptionsByOrgId = ALLSTATS_buildPromoRedemptionByOrgId_(promoRedemptions)
    ALLSTATS_applyPromoEligibilityFromRedemptions_(orgs, promoRedemptionsByOrgId)
    ALLSTATS_applyTrialExtensionFromCanon_(orgs, canonOrgs)
    orgs.sort((a, b) => {
      if (b.arr_total !== a.arr_total) return b.arr_total - a.arr_total
      return String(a.org_name || '').localeCompare(String(b.org_name || ''))
    })

    const managedSets = ALLSTATS_buildManagedExclusionSets_(shManual)
    const allMetrics = ALLSTATS_buildStageMetricsFromOrgSubscriptionInfo_(orgSubscriptions, manualBySubId, stripeBySubId, managedSets)
    const conversion = ALLSTATS_buildConversionMetrics_(ss)

    const snapRows = shSnap ? ALLSTATS_readSheetObjects_(shSnap, 1) : []
    const waterfallRows = shWaterfall ? ALLSTATS_readSheetObjects_(shWaterfall, 1) : []
    const netNew = ALLSTATS_buildNetNewByMonth_(waterfallRows)
    const retention = ALLSTATS_buildRetentionMetrics_(snapRows, waterfallRows, orgSubscriptions)
    const subscriptionChurn = ALLSTATS_buildSubscriptionChurnMetrics_(orgSubscriptions, manualBySubId, posthogOrgSubsRows)
    const waterfall = ALLSTATS_buildWaterfallTables_(waterfallRows)

    const freeTrialList = ALLSTATS_buildFreeTrialList_(orgs)
    const promoTrialList = ALLSTATS_buildPromoTrialList_(orgs)
    const paidList = ALLSTATS_buildPaidList_(orgs)
    const seatTypes = ALLSTATS_buildSeatTypeSummary_(canonOrgs, orgSubscriptions, manualBySubId, stripeBySubId)

    ALLSTATS_renderAllStatsSheet_(out, {
      generatedAt: new Date(),
      stripeRowsCount: stripeRows.length,
      orgs,
      metrics: allMetrics,
      conversion,
      netNew,
      retention,
      subscriptionChurn,
      waterfall,
      freeTrialList,
      promoTrialList,
      paidList,
      seatTypes
    })

    // Write Promo Audit sheet
    ALLSTATS_renderPromoAudit_(ss, conversion)

    // Clean up deprecated sheets
    const deprecatedSheets = ['Seat Audit', 'Conversion Audit', 'arr_subscription_mapping_audit', 'No Login again']
    for (const name of deprecatedSheets) {
      const old = ss.getSheetByName(name)
      if (old) { try { ss.deleteSheet(old) } catch (e) {} }
    }

    const seconds = (new Date() - t0) / 1000
    ALLSTATS_writeSyncLog_('render_all_stats_view', 'ok', stripeRows.length, orgs.length, seconds, '')

    return { rows_in: stripeRows.length, rows_out: orgs.length }
  })
}

function ALLSTATS_normalizeSubscription_(r, manualBySubId, indexes, orgSubsIndex) {
  const statusRaw = ALLSTATS_str_(r.status).toLowerCase()
  const hasPaymentMethod = ALLSTATS_toBool_(r.has_payment_method)

  const subId =
    ALLSTATS_str_(r.stripe_subscription_id) ||
    ALLSTATS_str_(r.subscription_id) ||
    ALLSTATS_str_(r.subscription) ||
    ALLSTATS_str_(r.id)

  const manual = subId ? (manualBySubId.get(subId) || null) : null
  if (manual && manual.excludeInternal) {
    return { include: false }
  }

  const stage = ALLSTATS_classifyStage_(statusRaw, hasPaymentMethod)
  if (!stage) return { include: false }

  const discountPercent = ALLSTATS_num_(r.discount_percent)
  const discountDuration = ALLSTATS_str_(r.discount_duration).toLowerCase()
  const discountDurationMonths = ALLSTATS_num_(r.discount_duration_months)

  if (discountPercent === 100 && discountDuration === 'forever') {
    return { include: false }
  }

  const interval = ALLSTATS_str_(r.interval).toLowerCase()
  const intervalCount = Math.max(1, ALLSTATS_num_(r.interval_count) || 1)
  const amountRaw = ALLSTATS_moneyAmount_(r.amount)
  const discountCtx = ALLSTATS_buildDiscountContextNow_(r, {
    amountRaw,
    interval,
    intervalCount,
    asOfDate: new Date()
  })
  const amount = discountCtx.amount
  const mrrArr = ALLSTATS_computeMrrArr_(amount, interval, intervalCount)

  const seats = ALLSTATS_safeInt_(r.quantity_total)

  const firstPaymentAtIso = ALLSTATS_str_(r.first_payment_at)
  const firstPaymentAtDate = ALLSTATS_isoToDateOrNull_(firstPaymentAtIso)

  const createdAtIso = ALLSTATS_str_(r.created_at)
  const createdAtDate = ALLSTATS_isoToDateOrNull_(createdAtIso)

  const stripeEmail = ALLSTATS_str_(r.customer_email || r.email || r.billing_email)
  const stripeEmailKey = ALLSTATS_normalizeEmail_(stripeEmail)

  const resolved = ALLSTATS_resolveCustomer_(stripeEmailKey, subId, indexes)
  const canonMatch = (subId && indexes && indexes.canonBySubId)
    ? (indexes.canonBySubId.get(subId) || null)
    : null
  const orgSub = (orgSubsIndex && orgSubsIndex.bySubId) ? (orgSubsIndex.bySubId.get(subId) || null) : null

  const customerEmail = resolved.email || stripeEmail || (canonMatch ? canonMatch.billingEmail : '')
  const customerName = resolved.customerName || ALLSTATS_str_(r.customer_name || r.name)
  const orgId = ALLSTATS_str_(
    (canonMatch && (canonMatch.appOrgId || canonMatch.orgId || canonMatch.clerkOrgId)) ||
    (orgSub && (orgSub.app_org_id || orgSub.org_id)) ||
    (resolved.appOrgId || resolved.orgId) ||
    ''
  )
  const orgName =
    (canonMatch && canonMatch.orgName) ||
    (orgId && indexes && indexes.orgNameByOrgId && indexes.orgNameByOrgId.get(orgId)) ||
    resolved.orgName ||
    ALLSTATS_str_(r.org_name || r.organization_name || r.org)

  const promoCode = discountCtx.promoCodes.join(', ')
  const hasPromoCode = !!promoCode
  const promoEvidence = hasPromoCode || discountCtx.activeDiscountCount > 0

  return {
    include: true,
    sub_id: subId,
    org_id: orgId,
    org_name: orgName,
    customer_email: customerEmail,
    customer_email_key: ALLSTATS_normalizeEmail_(customerEmail),
    customer_name: customerName,

    stage,
    interval,
    amount,
    mrr: mrrArr.mrr,
    arr: mrrArr.arr,
    seats,

    discount_percent: discountPercent,
    discount_duration: discountDuration,
    discount_duration_months: discountDurationMonths,
    promo_code: promoCode,
    promo_evidence: promoEvidence,
    trial_extended_days: 0,

    created_at: createdAtDate,
    first_payment_at: firstPaymentAtDate,

    has_payment_method: hasPaymentMethod
  }
}

function ALLSTATS_orgKey_(sub) {
  if (sub.org_id) return 'org:' + sub.org_id
  if (sub.sub_id) return 'sub:' + sub.sub_id
  return 'sub:unknown'
}

function ALLSTATS_newOrgAgg_(sub, orgKey) {
  return {
    org_key: orgKey,
    org_id: sub.org_id || '',
    org_name: sub.org_name || '',
    customer_name: sub.customer_name || '',
    customer_email: sub.customer_email || '',

    has_paid: false,
    has_promo: false,
    has_free: false,
    has_paid_us: false,

    has_promo_evidence: false,
    has_free_promo_evidence: false,
    has_promo_conversion_eligible: false,
    has_manual_trial_extension: false,

    paid: { arr: 0, mrr: 0, seats: 0, subscriptions: 0 },
    paid_true: { arr: 0, mrr: 0, seats: 0, subscriptions: 0 },
    promo: { arr: 0, mrr: 0, seats: 0, subscriptions: 0 },
    free: { arr: 0, mrr: 0, seats: 0, subscriptions: 0 },

    sub_ids: new Set(),
    promo_codes: new Set(),
    promo_redemption_code_ids: new Set(),
    promo_redemption_codes: new Set(),

    earliest_first_payment_at: null,
    earliest_free_trial_start: null,
    earliest_promo_trial_start: null
  }
}

function ALLSTATS_addSubToOrgAgg_(org, sub) {
  if (sub.org_id && !org.org_id) org.org_id = sub.org_id
  if (sub.org_name && !org.org_name) org.org_name = sub.org_name
  if (sub.customer_name && !org.customer_name) org.customer_name = sub.customer_name
  if (sub.customer_email && !org.customer_email) org.customer_email = sub.customer_email

  if (sub.sub_id) org.sub_ids.add(sub.sub_id)
  if (sub.promo_code) org.promo_codes.add(sub.promo_code)

  if (sub.promo_evidence) org.has_promo_evidence = true
  if (Number(sub.trial_extended_days || 0) > 0) {
    org.has_manual_trial_extension = true
    org.has_promo_conversion_eligible = true
  }

  if (sub.stage === 'Paid') {
    org.has_paid = true
    org.paid.arr += sub.arr
    org.paid.mrr += sub.mrr
    org.paid.seats += sub.seats
    org.paid.subscriptions += 1

    if (sub.first_payment_at) {
      org.has_paid_us = true
      org.paid_true.arr += sub.arr
      org.paid_true.mrr += sub.mrr
      org.paid_true.seats += sub.seats
      org.paid_true.subscriptions += 1
      if (!org.earliest_first_payment_at || sub.first_payment_at < org.earliest_first_payment_at) {
        org.earliest_first_payment_at = sub.first_payment_at
      }
    }
    return
  }

  if (sub.stage === 'Promo Trial') {
    org.has_promo = true
    org.promo.arr += sub.arr
    org.promo.mrr += sub.mrr
    org.promo.seats += sub.seats
    org.promo.subscriptions += 1

    if (sub.created_at) {
      if (!org.earliest_promo_trial_start || sub.created_at < org.earliest_promo_trial_start) {
        org.earliest_promo_trial_start = sub.created_at
      }
    }
    return
  }

  if (sub.stage === 'Free Trial') {
    org.has_free = true
    org.free.arr += sub.arr
    org.free.mrr += sub.mrr
    org.free.seats += sub.seats
    org.free.subscriptions += 1

    if (sub.created_at) {
      if (!org.earliest_free_trial_start || sub.created_at < org.earliest_free_trial_start) {
        org.earliest_free_trial_start = sub.created_at
      }
    }

    if (sub.promo_evidence) org.has_free_promo_evidence = true
  }
}

function ALLSTATS_finalizeOrgAgg_(org) {
  const stage = org.has_paid ? 'Paid' : (org.has_promo ? 'Promo Trial' : (org.has_free ? 'Free Trial' : 'Other'))
  const now = new Date()
  const freeDays = org.earliest_free_trial_start
    ? Math.max(0, Math.floor((now.getTime() - org.earliest_free_trial_start.getTime()) / 86400000) + 1)
    : ''

  const subIds = Array.from(org.sub_ids)
  const promoCodes = Array.from(org.promo_codes)
  const promoRedemptionCodeIds = Array.from(org.promo_redemption_code_ids || [])
  const promoRedemptionCodes = Array.from(org.promo_redemption_codes || [])

  return {
    org_key: org.org_key,
    org_id: org.org_id,
    org_name: org.org_name,
    customer_name: org.customer_name,
    customer_email: org.customer_email,

    stage,

    has_paid: org.has_paid,
    has_promo: org.has_promo,
    has_free: org.has_free,
    has_paid_us: org.has_paid_us,
    has_promo_evidence: org.has_promo_evidence,
    has_free_promo_evidence: org.has_free_promo_evidence,
    has_promo_conversion_eligible: org.has_promo_conversion_eligible,
    has_manual_trial_extension: org.has_manual_trial_extension,

    paid: org.paid,
    paid_true: org.paid_true,
    promo: org.promo,
    free: org.free,

    arr_total: org.paid.arr + org.promo.arr + org.free.arr,
    mrr_total: org.paid.mrr + org.promo.mrr + org.free.mrr,
    seats_total: org.paid.seats + org.promo.seats + org.free.seats,

    subscription_count: subIds.length,
    subscription_ids: subIds,
    subscription_ids_text: subIds.join(', '),

    promo_codes: promoCodes,
    promo_codes_text: promoCodes.join(', '),
    promo_redemption_code_ids: promoRedemptionCodeIds,
    promo_redemption_code_ids_text: promoRedemptionCodeIds.join(', '),
    promo_redemption_codes: promoRedemptionCodes,
    promo_redemption_codes_text: promoRedemptionCodes.join(', '),
    promo_redemption_count: 0,
    promo_last_redeemed_at: null,

    earliest_first_payment_at: org.earliest_first_payment_at,
    earliest_free_trial_start: org.earliest_free_trial_start,
    earliest_promo_trial_start: org.earliest_promo_trial_start,
    free_trial_days: freeDays
  }
}

function ALLSTATS_buildSeatTypeSummary_(canonOrgs, orgSubscriptions, manualBySubId, stripeBySubId) {
  const LITE_PRODUCT_IDS = new Set([
    'prod_THnR7uM4okazF7',
    'prod_TPWj79k7DB190P',
    'prod_U1vVdZlnpqyrk8',
    'prod_U1vVKlFs5EGSfo'
  ])

  const byStage = {
    paid: { full: 0, lite: 0, fullOrgs: 0, liteOrgs: 0 },
    intent: { full: 0, lite: 0, fullOrgs: 0, liteOrgs: 0 },
    trialing: { full: 0, lite: 0, fullOrgs: 0, liteOrgs: 0 }
  }
  let fullSeats = 0, liteSeats = 0, fullOrgs = 0, liteOrgs = 0
  const orgsWithSeats = new Set()

  // Track per-org totals to count orgs correctly
  const orgFullSeats = new Map()
  const orgLiteSeats = new Map()
  const orgStageBucket = new Map()

  ;(orgSubscriptions || []).forEach(r => {
    const subId =
      ALLSTATS_str_(r.latest_subscription_id) ||
      ALLSTATS_str_(r.stripe_subscription_id) ||
      ALLSTATS_str_(r.subscription_id) ||
      ALLSTATS_str_(r.id)
    const manual = subId ? (manualBySubId && manualBySubId.get(subId)) : null
    if (manual && manual.excludeInternal) return

    const oid = ALLSTATS_str_(r.app_org_id || r.org_id)
    if (!oid) return

    const status = ALLSTATS_str_(r.status).toLowerCase()
    const hasFirstPayment = !!ALLSTATS_str_(r.first_payment_at)
    const hasPaymentMethod = ALLSTATS_toBool_(r.has_payment_method)
    let bucket = ALLSTATS_bucketFromOrgSubscriptionInfo_(status, hasFirstPayment, hasPaymentMethod)
    if (!bucket) return

    // Same 100% discount override as ARR summary
    if (bucket === 'paid') {
      const discountedYearly = ALLSTATS_num_(r.amount_after_discounts_yearly)
      const discountedAmount = ALLSTATS_num_(r.amount_after_discounts)
      const arr = discountedYearly > 0 ? discountedYearly : discountedAmount > 0 ? discountedAmount : 0
      if (arr < 0.01) {
        bucket = hasPaymentMethod ? 'intent_to_pay' : 'free_trial'
      }
    }

    const seats = ALLSTATS_safeInt_(r.quantity_total)
    if (seats <= 0) return

    // Determine lite vs full from stripe items_json product_ids
    const stripeRow = subId ? (stripeBySubId && stripeBySubId.get(subId)) : null
    let isLite = false
    let liteSeatsInSub = 0
    let fullSeatsInSub = 0
    if (stripeRow) {
      let items = []
      try { items = JSON.parse(ALLSTATS_str_(stripeRow.items_json)) || [] } catch (e) { items = [] }
      if (items.length > 0) {
        for (const it of items) {
          const pid = ALLSTATS_str_(it.product_id)
          const qty = ALLSTATS_safeInt_(it.quantity)
          if (LITE_PRODUCT_IDS.has(pid)) {
            liteSeatsInSub += qty
          } else {
            fullSeatsInSub += qty
          }
        }
      } else {
        // No items_json, treat all seats as full
        fullSeatsInSub = seats
      }
    } else {
      fullSeatsInSub = seats
    }

    if (!orgFullSeats.has(oid)) orgFullSeats.set(oid, 0)
    if (!orgLiteSeats.has(oid)) orgLiteSeats.set(oid, 0)

    orgFullSeats.set(oid, orgFullSeats.get(oid) + fullSeatsInSub)
    orgLiteSeats.set(oid, orgLiteSeats.get(oid) + liteSeatsInSub)

    // Keep highest priority bucket per org
    const rank = { paid: 3, intent_to_pay: 2, free_trial: 1 }
    const existing = orgStageBucket.get(oid)
    if (!existing || (rank[bucket] || 0) > (rank[existing] || 0)) {
      orgStageBucket.set(oid, bucket)
    }
  })

  // Aggregate per-org totals into stage breakdowns
  const allOrgIds = new Set([...orgFullSeats.keys(), ...orgLiteSeats.keys()])
  for (const oid of allOrgIds) {
    const full = orgFullSeats.get(oid) || 0
    const lite = orgLiteSeats.get(oid) || 0
    const stage = orgStageBucket.get(oid)

    if (full > 0) { fullSeats += full; fullOrgs += 1; orgsWithSeats.add(oid) }
    if (lite > 0) { liteSeats += lite; liteOrgs += 1; orgsWithSeats.add(oid) }

    const stageBucket = stage === 'paid' ? byStage.paid
      : stage === 'intent_to_pay' ? byStage.intent
      : (stage === 'free_trial' ? byStage.trialing : null)
    if (stageBucket) {
      if (full > 0) { stageBucket.full += full; stageBucket.fullOrgs += 1 }
      if (lite > 0) { stageBucket.lite += lite; stageBucket.liteOrgs += 1 }
    }
  }

  return { fullSeats, liteSeats, fullOrgs, liteOrgs, totalOrgsWithSeats: orgsWithSeats.size, byStage }
}

function ALLSTATS_buildStageMetrics_(orgs) {
  const out = {
    paidDefined: ALLSTATS_emptyMetric_(),
    promoTrialDefined: ALLSTATS_emptyMetric_(),
    freeTrialDefined: ALLSTATS_emptyMetric_(),

    paidPromo: ALLSTATS_emptyMetric_(),
    onlyPaidTrue: ALLSTATS_emptyMetric_(),
    onlyPromo: ALLSTATS_emptyMetric_(),
    onlyPaid: ALLSTATS_emptyMetric_(),
    onlyFree: ALLSTATS_emptyMetric_(),

    stagePaid: { firms: 0, seats: 0 },
    stagePromo: { firms: 0, seats: 0 },
    stageFree: { firms: 0, seats: 0 },

    paidStageTotals: { arr: 0, seats: 0, firms: 0 }
  }

  for (const o of orgs) {
    const hasFirstPayment = !!o.has_paid_us
    const isActiveNoFirstPayment = !!(o.has_paid && !hasFirstPayment)
    const hasAnyPromoUsage =
      !!o.has_promo_conversion_eligible ||
      !!o.has_promo_evidence ||
      !!o.has_free_promo_evidence ||
      !!ALLSTATS_str_(o.promo_codes_text) ||
      !!ALLSTATS_str_(o.promo_redemption_codes_text) ||
      !!ALLSTATS_str_(o.promo_redemption_code_ids_text) ||
      Number(o.promo_redemption_count || 0) > 0
    const trialArr = (Number(o.promo.arr || 0) + Number(o.free.arr || 0))
    const trialMrr = (Number(o.promo.mrr || 0) + Number(o.free.mrr || 0))
    const trialSeats = (Number(o.promo.seats || 0) + Number(o.free.seats || 0))

    if (hasFirstPayment) {
      out.paidDefined.arr += o.paid_true.arr
      out.paidDefined.mrr += o.paid_true.mrr
      out.paidDefined.firms += 1
      out.paidDefined.seats += o.paid_true.seats
    } else if (isActiveNoFirstPayment || hasAnyPromoUsage) {
      out.promoTrialDefined.arr += trialArr
      out.promoTrialDefined.mrr += trialMrr
      out.promoTrialDefined.firms += 1
      out.promoTrialDefined.seats += trialSeats
    } else {
      out.freeTrialDefined.arr += trialArr
      out.freeTrialDefined.mrr += trialMrr
      out.freeTrialDefined.firms += 1
      out.freeTrialDefined.seats += trialSeats
    }

    const onlyPaid = o.has_paid && !o.has_promo
    const onlyPaidTrue = onlyPaid && o.has_paid_us
    const onlyPromo = o.has_promo && !o.has_paid
    const onlyFree = o.has_free && !o.has_paid && !o.has_promo
    const paidPromo = o.has_paid || o.has_promo

    if (paidPromo) {
      out.paidPromo.arr += (o.paid.arr + o.promo.arr)
      out.paidPromo.mrr += (o.paid.mrr + o.promo.mrr)
      out.paidPromo.firms += 1
      out.paidPromo.seats += (o.paid.seats + o.promo.seats)
    }

    if (onlyPromo) {
      out.onlyPromo.arr += o.promo.arr
      out.onlyPromo.mrr += o.promo.mrr
      out.onlyPromo.firms += 1
      out.onlyPromo.seats += o.promo.seats
    }

    if (onlyPaid) {
      out.onlyPaid.arr += o.paid.arr
      out.onlyPaid.mrr += o.paid.mrr
      out.onlyPaid.firms += 1
      out.onlyPaid.seats += o.paid.seats
    }

    if (onlyPaidTrue) {
      out.onlyPaidTrue.arr += o.paid_true.arr
      out.onlyPaidTrue.mrr += o.paid_true.mrr
      out.onlyPaidTrue.firms += 1
      out.onlyPaidTrue.seats += o.paid_true.seats
    }

    if (onlyFree) {
      out.onlyFree.arr += o.free.arr
      out.onlyFree.mrr += o.free.mrr
      out.onlyFree.firms += 1
      out.onlyFree.seats += o.free.seats
    }

    if (o.has_paid) {
      out.stagePaid.firms += 1
      out.stagePaid.seats += (o.paid.seats + o.promo.seats)
      out.paidStageTotals.arr += o.paid.arr
      out.paidStageTotals.seats += o.paid.seats
      out.paidStageTotals.firms += 1
    } else if (o.has_promo) {
      out.stagePromo.firms += 1
      out.stagePromo.seats += o.promo.seats
    } else if (o.has_free) {
      out.stageFree.firms += 1
      out.stageFree.seats += o.free.seats
    }
  }

  out.avgRevenuePerFirmPaid = out.paidStageTotals.firms
    ? (out.paidStageTotals.arr / out.paidStageTotals.firms)
    : 0

  out.avgRevenuePerSeat = out.paidStageTotals.seats
    ? (out.paidStageTotals.arr / out.paidStageTotals.seats)
    : 0

  out.avgSeatCountPerFirmPaid = out.paidStageTotals.firms
    ? (out.paidStageTotals.seats / out.paidStageTotals.firms)
    : 0

  return out
}

function ALLSTATS_buildStageMetricsFromOrgSubscriptionInfo_(rows, manualBySubId, stripeBySubId, managedSets) {
  const managed = managedSets || { subIds: new Set(), customerIds: new Set() }
  const out = {
    paidDefined: ALLSTATS_emptyMetric_(),
    promoTrialDefined: ALLSTATS_emptyMetric_(),
    freeTrialDefined: ALLSTATS_emptyMetric_(),

    paidPromo: ALLSTATS_emptyMetric_(),
    onlyPaidTrue: ALLSTATS_emptyMetric_(),
    onlyPromo: ALLSTATS_emptyMetric_(),
    onlyPaid: ALLSTATS_emptyMetric_(),
    onlyFree: ALLSTATS_emptyMetric_(),
    stagePaid: { firms: 0, seats: 0 },
    stagePromo: { firms: 0, seats: 0 },
    stageFree: { firms: 0, seats: 0 },
    paidStageTotals: { arr: 0, seats: 0, firms: 0 },
    paidStageTotalsExclManaged: { arr: 0, seats: 0, firms: 0 },
    avgRevenuePerFirmPaid: 0,
    avgRevenuePerFirmPaidExclManaged: 0,
    avgRevenuePerSeat: 0,
    avgSeatCountPerFirmPaid: 0
  }

  ;(rows || []).forEach(r => {
    const subId =
      ALLSTATS_str_(r.latest_subscription_id) ||
      ALLSTATS_str_(r.stripe_subscription_id) ||
      ALLSTATS_str_(r.subscription_id) ||
      ALLSTATS_str_(r.id)
    const manual = subId ? (manualBySubId && manualBySubId.get(subId)) : null
    if (manual && manual.excludeInternal) return

    const custId = ALLSTATS_str_(r.stripe_customer_id)
    const isManaged = (subId && managed.subIds.has(subId)) ||
                      (custId && managed.customerIds.has(custId))

    const status = ALLSTATS_str_(r.status).toLowerCase()
    const hasFirstPayment = !!ALLSTATS_str_(r.first_payment_at)
    const hasPaymentMethod = ALLSTATS_toBool_(r.has_payment_method)
    let bucket = ALLSTATS_bucketFromOrgSubscriptionInfo_(status, hasFirstPayment, hasPaymentMethod)
    if (!bucket) return

    const interval = ALLSTATS_str_(r.interval).toLowerCase()
    const intervalCount = Math.max(1, ALLSTATS_num_(r.interval_count) || 1)
    const seats = ALLSTATS_safeInt_(r.quantity_total)

    // ── Single source of truth: discount-adjusted ARR from org_subscription_info (same as Ring) ──
    const discountedYearly = ALLSTATS_num_(r.amount_after_discounts_yearly)
    const discountedAmount = ALLSTATS_num_(r.amount_after_discounts)
    const arr = discountedYearly > 0
      ? discountedYearly
      : discountedAmount > 0
        ? ALLSTATS_computeMrrArr_(discountedAmount, interval, intervalCount).arr
        : 0

    // Gross (pre-discount) ARR for 100%-off reclassification
    const grossYearly = ALLSTATS_num_(r.amount_yearly)
    const grossAmount = ALLSTATS_num_(r.amount)
    const grossArr = grossYearly > 0
      ? grossYearly
      : ALLSTATS_computeMrrArr_(grossAmount, interval, intervalCount).arr

    // ── 100% discount override (same as Ring) ──
    let bucketArr = arr
    if (bucket === 'paid' && arr < 0.01) {
      bucket = hasPaymentMethod ? 'intent_to_pay' : 'free_trial'
      bucketArr = grossArr
    }

    const mrr = bucketArr / 12

    if (bucket === 'paid') {
      out.paidDefined.arr += bucketArr
      out.paidDefined.mrr += mrr
      out.paidDefined.firms += 1
      out.paidDefined.seats += seats
      if (!isManaged) {
        out.paidStageTotalsExclManaged.arr += bucketArr
        out.paidStageTotalsExclManaged.seats += seats
        out.paidStageTotalsExclManaged.firms += 1
      }
    } else if (bucket === 'intent_to_pay') {
      out.promoTrialDefined.arr += bucketArr
      out.promoTrialDefined.mrr += mrr
      out.promoTrialDefined.firms += 1
      out.promoTrialDefined.seats += seats
    } else if (bucket === 'free_trial') {
      out.freeTrialDefined.arr += bucketArr
      out.freeTrialDefined.mrr += mrr
      out.freeTrialDefined.firms += 1
      out.freeTrialDefined.seats += seats
    }
  })

  out.stagePaid.firms = out.paidDefined.firms
  out.stagePaid.seats = out.paidDefined.seats
  out.stagePromo.firms = out.promoTrialDefined.firms
  out.stagePromo.seats = out.promoTrialDefined.seats
  out.stageFree.firms = out.freeTrialDefined.firms
  out.stageFree.seats = out.freeTrialDefined.seats
  out.paidStageTotals.arr = out.paidDefined.arr
  out.paidStageTotals.seats = out.paidDefined.seats
  out.paidStageTotals.firms = out.paidDefined.firms

  out.avgRevenuePerFirmPaid = out.paidStageTotals.firms
    ? (out.paidStageTotals.arr / out.paidStageTotals.firms)
    : 0

  out.avgRevenuePerSeat = out.paidStageTotals.seats
    ? (out.paidStageTotals.arr / out.paidStageTotals.seats)
    : 0

  out.avgRevenuePerFirmPaidExclManaged = out.paidStageTotalsExclManaged.firms
    ? (out.paidStageTotalsExclManaged.arr / out.paidStageTotalsExclManaged.firms)
    : 0

  out.avgSeatCountPerFirmPaid = out.paidStageTotals.firms
    ? (out.paidStageTotals.seats / out.paidStageTotals.firms)
    : 0

  return out
}

function ALLSTATS_buildManagedExclusionSets_(shManual) {
  const subIds = new Set()
  const customerIds = new Set()
  if (!shManual) return { subIds, customerIds }
  const rows = ALLSTATS_readSheetObjects_(shManual, 1)
  for (const r of rows) {
    const reason = ALLSTATS_str_(r.exclude_reason).toLowerCase()
    if (reason !== 'managed') continue
    const subId = ALLSTATS_str_(r.subscription_id || r.stripe_subscription_id || r.subscription)
    if (subId) subIds.add(subId)
    const custId = ALLSTATS_str_(r.customer_id || r.stripe_customer_id)
    if (custId) customerIds.add(custId)
  }
  return { subIds, customerIds }
}

function ALLSTATS_bucketFromOrgSubscriptionInfo_(status, hasFirstPayment, hasPaymentMethod) {
  // Same logic as Ring: ringBucketFromOrgSubscriptionInfo_
  if (status === 'active' && hasFirstPayment) return 'paid'
  if (status === 'active' && !hasFirstPayment && hasPaymentMethod) return 'intent_to_pay'
  if (status === 'trialing' && hasPaymentMethod) return 'intent_to_pay'
  if (status === 'active' && !hasFirstPayment && !hasPaymentMethod) return 'free_trial'
  if (status === 'trialing' && !hasPaymentMethod) return 'free_trial'
  return ''
}

function ALLSTATS_buildStripeBySubId_(rows) {
  const out = new Map()
  ;(rows || []).forEach(r => {
    const subId =
      ALLSTATS_str_(r.stripe_subscription_id) ||
      ALLSTATS_str_(r.subscription_id) ||
      ALLSTATS_str_(r.subscription) ||
      ALLSTATS_str_(r.id)
    if (!subId || out.has(subId)) return
    out.set(subId, r)
  })
  return out
}

function ALLSTATS_emptyMetric_() {
  return { arr: 0, mrr: 0, firms: 0, seats: 0 }
}

function ALLSTATS_buildConversionMetrics_(ss) {
  // Uses shared canon_orgs-based org list (filtered, enriched with first_payment_at)
  const convOrgs = buildConversionOrgList_(ss)

  let totalSignedUp = 0
  let paidUs = 0
  let signupTrialPotential = 0
  let promoPool = 0
  let promoToPaid = 0
  let promoTrialPotential = 0

  for (const org of convOrgs) {
    totalSignedUp += 1

    if (org.first_payment_at) {
      paidUs += 1
    }
    if (org.is_currently_trialing) {
      signupTrialPotential += 1
    }

    if (org.has_promo) {
      promoPool += 1
      if (org.first_payment_at) {
        promoToPaid += 1
      } else if (org.is_currently_trialing) {
        promoTrialPotential += 1
      }
    }
  }

  const signupPotentialMaxRate = totalSignedUp
    ? ((paidUs + signupTrialPotential) / totalSignedUp)
    : 0

  const promoPotentialMaxRate = promoPool
    ? ((promoToPaid + promoTrialPotential) / promoPool)
    : 0

  // Denominator excludes orgs still trialing — only count orgs whose trial resolved
  const signupResolved = totalSignedUp - signupTrialPotential
  const promoResolved = promoPool - promoTrialPotential

  return {
    totalSignedUp,
    signupResolved,
    paidUs,
    signupToPaidRate: signupResolved ? (paidUs / signupResolved) : 0,
    signupTrialPotential,
    signupPotentialMaxRate,
    promoPool,
    promoResolved,
    promoToPaid,
    promoToPaidRate: promoResolved ? (promoToPaid / promoResolved) : 0,
    promoTrialPotential,
    promoPotentialMaxRate
  }
}

function ALLSTATS_loadPromoRedemptions_(sheetMaybe) {
  if (!sheetMaybe) return []
  return ALLSTATS_readSheetObjects_(sheetMaybe, 1)
}

function ALLSTATS_loadOrgSubscriptions_(sheetMaybe) {
  if (!sheetMaybe) return []
  return ALLSTATS_readSheetObjects_(sheetMaybe, 1)
}

function ALLSTATS_buildOrgSubscriptionsIndex_(rows) {
  const bySubId = new Map()
  const byOrgId = new Map()
  const warnings = []

  for (const r of (rows || [])) {
    const subId = ALLSTATS_str_(r.stripe_subscription_id || r.subscription_id || r.id)
    const orgId = ALLSTATS_str_(r.app_org_id || r.org_id)
    if (!subId || !orgId) continue

    if (!bySubId.has(subId)) bySubId.set(subId, r)
    else warnings.push(`duplicate stripe_subscription_id in org_subscriptions: ${subId}`)

    const prev = byOrgId.get(orgId)
    if (!prev) byOrgId.set(orgId, r)
    else {
      const pick = ALLSTATS_pickLatestOrgSub_(prev, r)
      byOrgId.set(orgId, pick)
      warnings.push(`multiple subscriptions for org_id in org_subscriptions: ${orgId}`)
    }
  }

  if (warnings.length) {
    Logger.log('[ALLSTATS] org_subscriptions warnings: ' + warnings.slice(0, 20).join(' | '))
  }

  return { bySubId, byOrgId }
}

function ALLSTATS_pickLatestOrgSub_(a, b) {
  const aDate = ALLSTATS_toDateOrNull_(a.updated_at || a.created_at || a.trial_started_at || '')
  const bDate = ALLSTATS_toDateOrNull_(b.updated_at || b.created_at || b.trial_started_at || '')
  const aTs = aDate ? aDate.getTime() : 0
  const bTs = bDate ? bDate.getTime() : 0
  return bTs >= aTs ? b : a
}

function ALLSTATS_buildPromoRedemptionByOrgId_(promoRedemptions, promoCodeById) {
  const promoLookup = promoCodeById instanceof Map ? promoCodeById : new Map()
  const byOrgId = new Map()
  for (const r of (promoRedemptions || [])) {
    const orgId = ALLSTATS_str_(r.app_org_id || r.org_id)
    if (!orgId) continue

    const promoCodeId = ALLSTATS_str_(r.promo_code_id || r.promocode_id || r.code_id)
    const promoLabel =
      ALLSTATS_str_(r.promo_code) ||
      ALLSTATS_str_(r.code) ||
      ALLSTATS_str_(r.promo_name) ||
      (promoCodeId ? (promoLookup.get(promoCodeId) || promoCodeId) : '')

    const redeemedAt = ALLSTATS_toDateOrNull_(r.redeemed_at)

    if (!byOrgId.has(orgId)) {
      byOrgId.set(orgId, {
        count: 0,
        promoCodeIds: new Set(),
        promoCodes: new Set(),
        firstRedeemedAt: null,
        lastRedeemedAt: null
      })
    }

    const bucket = byOrgId.get(orgId)
    bucket.count += 1
    if (promoCodeId) bucket.promoCodeIds.add(promoCodeId)
    if (promoLabel) bucket.promoCodes.add(promoLabel)

    if (redeemedAt) {
      if (!bucket.firstRedeemedAt || redeemedAt < bucket.firstRedeemedAt) bucket.firstRedeemedAt = redeemedAt
      if (!bucket.lastRedeemedAt || redeemedAt > bucket.lastRedeemedAt) bucket.lastRedeemedAt = redeemedAt
    }
  }
  return byOrgId
}

function ALLSTATS_applyPromoEligibilityFromRedemptions_(orgs, promoRedemptionsByOrgId) {
  for (const o of (orgs || [])) {
    const orgId = ALLSTATS_str_(o.org_id)
    const hit = orgId ? promoRedemptionsByOrgId.get(orgId) : null
    if (!hit) continue

    o.has_promo_conversion_eligible = true
    o.promo_redemption_count = Number(hit.count || 0)
    o.promo_redemption_code_ids_text = Array.from(hit.promoCodeIds || []).join(', ')
    o.promo_redemption_codes_text = Array.from(hit.promoCodes || []).join(', ')
    o.promo_last_redeemed_at = hit.lastRedeemedAt || null
  }
}

function ALLSTATS_applyTrialExtensionFromCanon_(orgs, canonOrgs) {
  const canonByOrgId = new Map()
  for (const c of (canonOrgs || [])) {
    const oid = ALLSTATS_str_(c.app_org_id || c.org_id)
    if (oid) canonByOrgId.set(oid, c)
  }
  for (const org of (orgs || [])) {
    const orgId = ALLSTATS_str_(org.org_id)
    if (!orgId) continue
    const canon = canonByOrgId.get(orgId)
    if (!canon) continue
    const extended = ALLSTATS_str_(canon.trial_extended).toLowerCase() === 'true'
    if (extended) {
      org.has_promo = true
      org.has_promo_conversion_eligible = true
      const src = ALLSTATS_str_(canon.trial_extension_source)
      if (src.includes('stripe_manual')) org.has_manual_trial_extension = true
    }
  }
}

function ALLSTATS_buildNetNewByMonth_(waterfallRows) {
  // Aggregate arr_waterfall_facts (per-org-per-month metrics) into monthly totals
  const byMonth = new Map()

  for (const r of (waterfallRows || [])) {
    const monthKey = ALLSTATS_str_(r.month)
    const metric = ALLSTATS_str_(r.metric)
    const amount = ALLSTATS_num_(r.amount)
    const orgId = ALLSTATS_str_(r.org_id)
    if (!monthKey || !metric) continue

    if (!byMonth.has(monthKey)) {
      byMonth.set(monthKey, {
        month: monthKey,
        bom_arr: 0,
        eom_arr: 0,
        net_new_arr: 0,
        new_orgs: 0,
        upgrades: 0,
        downgrades: 0,
        churned_orgs: 0,
        churn_arr: 0
      })
    }

    const b = byMonth.get(monthKey)
    if (metric === 'SOM') b.bom_arr += amount
    else if (metric === 'EOM') b.eom_arr += amount
    else if (metric === 'New') { b.upgrades += amount; b.new_orgs += 1 }
    else if (metric === 'Upgrade') b.upgrades += amount
    else if (metric === 'Downgrade') b.downgrades += amount
    else if (metric === 'Churn') { b.churn_arr += Math.abs(amount); b.churned_orgs += 1 }
  }

  for (const b of byMonth.values()) {
    b.net_new_arr = b.eom_arr - b.bom_arr
  }

  const rows = Array.from(byMonth.values()).sort((a, b) => String(a.month).localeCompare(String(b.month)))
  return { rows }
}

function ALLSTATS_buildRetentionMetrics_(snapshotRows, waterfallRows, orgSubscriptions) {
  const emptyResult = {
    latest_snapshot: '',
    base_bom_arr: 0,
    base_orgs: 0,
    nrr: 0,
    grr: 0,
    logo_churn_rate: 0,
    gross_arr_churn_rate: 0,
    full_arr_churn_rate: 0,
    churned_orgs: 0,
    churned_arr: 0,
    churned_orgs_list: [],
    churned_orgs_list_all: []
  }

  // Use arr_waterfall_facts which already has per-org SOM/EOM/Churn per month
  if (!waterfallRows || !waterfallRows.length) return emptyResult

  // Group waterfall rows by month + org
  const byMonthOrg = new Map()
  const allMonths = new Set()

  for (const r of waterfallRows) {
    const month = ALLSTATS_str_(r.month)
    const orgId = ALLSTATS_str_(r.org_id)
    const metric = ALLSTATS_str_(r.metric)
    const amount = ALLSTATS_num_(r.amount)
    if (!month || !orgId || !metric) continue

    allMonths.add(month)
    const key = month + '|' + orgId
    if (!byMonthOrg.has(key)) byMonthOrg.set(key, {})
    byMonthOrg.get(key)[metric] = amount
  }

  const monthsSorted = [...allMonths].sort()
  if (!monthsSorted.length) return emptyResult

  // Use the latest completed month (second to last) if available, otherwise latest
  // The latest month is likely the current (incomplete) month
  const latestMonth = monthsSorted.length >= 2
    ? monthsSorted[monthsSorted.length - 2]
    : monthsSorted[monthsSorted.length - 1]

  // Also collect org metadata from waterfall rows for the audit
  const orgMeta = new Map()
  for (const r of waterfallRows) {
    const orgId = ALLSTATS_str_(r.org_id)
    if (!orgId || orgMeta.has(orgId)) continue
    orgMeta.set(orgId, {
      org_id: orgId,
      org_name: ALLSTATS_str_(r.org_name),
      current_status: ALLSTATS_str_(r.current_status),
      ring_bucket: ALLSTATS_str_(r.ring_bucket),
      first_payment_date: ALLSTATS_str_(r.first_payment_date),
      churn_date: ALLSTATS_str_(r.churn_date),
      plan_name: ALLSTATS_str_(r.plan_name)
    })
  }

  // Stripe lookup metadata by org for audit debugging fields
  const stripeMetaByOrg = new Map()
  for (const r of (orgSubscriptions || [])) {
    const orgId = ALLSTATS_str_(r.app_org_id || r.org_id)
    if (!orgId || stripeMetaByOrg.has(orgId)) continue
    stripeMetaByOrg.set(orgId, {
      stripe_customer_email: ALLSTATS_str_(r.customer_email || r.billing_email || r.email),
      stripe_subscription_id: ALLSTATS_str_(r.stripe_subscription_id || r.latest_subscription_id || r.subscription_id)
    })
  }

  let baseBom = 0
  let baseEom = 0
  let retainedEomNoExpansion = 0
  let grossLoss = 0
  let fullChurnArr = 0
  let baseOrgs = 0
  let churnedOrgs = 0
  const churnedOrgsList = []
  const churnedOrgsAllByOrg = new Map()

  for (const [key, metrics] of byMonthOrg) {
    const month = key.split('|')[0]
    const orgId = key.split('|')[1]
    const som = ALLSTATS_num_(metrics.SOM)
    const eom = ALLSTATS_num_(metrics.EOM)
    if (som <= 0) continue  // only count orgs that had ARR at start of month
    const meta = orgMeta.get(orgId) || {}
    const status = ALLSTATS_str_(meta.current_status).toLowerCase()
    const hasCancelSignal = !!ALLSTATS_str_(meta.churn_date) || status === 'canceled' || status === 'cancelled'

    // All-time churn audit list (unique orgs, keep most recent churn month)
    if (eom <= 0 && hasCancelSignal) {
      const existing = churnedOrgsAllByOrg.get(orgId)
      if (!existing || String(month) > String(existing.churn_month || '')) {
        const stripeMeta = stripeMetaByOrg.get(orgId) || {}
        churnedOrgsAllByOrg.set(orgId, {
          org_id: orgId,
          org_name: meta.org_name || '',
          stripe_customer_email: stripeMeta.stripe_customer_email || '',
          stripe_subscription_id: stripeMeta.stripe_subscription_id || '',
          churn_month: month || '',
          som_arr: som,
          eom_arr: eom,
          current_status: meta.current_status || '',
          ring_bucket: meta.ring_bucket || '',
          first_payment_date: meta.first_payment_date || '',
          churn_date: meta.churn_date || '',
          plan_name: meta.plan_name || ''
        })
      }
    }

    // Latest month churn metrics used in Conversion + Retention section
    if (month !== latestMonth) continue

    baseOrgs += 1
    baseBom += som
    baseEom += eom
    retainedEomNoExpansion += Math.min(som, eom)

    const loss = Math.max(0, som - eom)
    grossLoss += loss

    // Churned org definition: paid at SOM, then canceled to zero ARR by EOM.
    if (eom <= 0 && hasCancelSignal) {
      churnedOrgs += 1
      fullChurnArr += som
      const stripeMeta = stripeMetaByOrg.get(orgId) || {}
      churnedOrgsList.push({
        org_id: orgId,
        org_name: meta.org_name || '',
        stripe_customer_email: stripeMeta.stripe_customer_email || '',
        stripe_subscription_id: stripeMeta.stripe_subscription_id || '',
        som_arr: som,
        eom_arr: eom,
        current_status: meta.current_status || '',
        ring_bucket: meta.ring_bucket || '',
        first_payment_date: meta.first_payment_date || '',
        churn_date: meta.churn_date || '',
        plan_name: meta.plan_name || ''
      })
    }
  }

  const nrr = baseBom > 0 ? (baseEom / baseBom) : 0
  const grr = baseBom > 0 ? (retainedEomNoExpansion / baseBom) : 0
  const logoChurnRate = baseOrgs > 0 ? (churnedOrgs / baseOrgs) : 0
  const grossArrChurnRate = baseBom > 0 ? (grossLoss / baseBom) : 0
  const fullArrChurnRate = baseBom > 0 ? (fullChurnArr / baseBom) : 0
  const churnedOrgsListAll = Array.from(churnedOrgsAllByOrg.values()).sort((a, b) => {
    const monthCmp = String(b.churn_month || '').localeCompare(String(a.churn_month || ''))
    if (monthCmp !== 0) return monthCmp
    return String(a.org_name || '').localeCompare(String(b.org_name || ''))
  })

  return {
    latest_snapshot: latestMonth,
    base_bom_arr: baseBom,
    base_orgs: baseOrgs,
    nrr,
    grr,
    logo_churn_rate: logoChurnRate,
    gross_arr_churn_rate: grossArrChurnRate,
    full_arr_churn_rate: fullArrChurnRate,
    churned_orgs: churnedOrgs,
    churned_arr: fullChurnArr,
    churned_orgs_list: churnedOrgsList,
    churned_orgs_list_all: churnedOrgsListAll
  }
}

function ALLSTATS_buildWaterfallTables_(rows) {
  const parsed = []

  for (const r of (rows || [])) {
    const d = ALLSTATS_toDateOrNull_(r.snapshot_date)
    if (!d) continue

    const dateKey = Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd')
    const metric = ALLSTATS_str_(r.metric).toLowerCase()
    const amount = ALLSTATS_num_(r.amount)

    parsed.push({
      date_key: dateKey,
      cohort: ALLSTATS_str_(r.sign_up_cohort_month) || ALLSTATS_str_(r.paid_cohort_month) || '(blank)',
      org_id: ALLSTATS_str_(r.org_id),
      org_name: ALLSTATS_str_(r.org_name),
      metric,
      amount
    })
  }

  if (!parsed.length) {
    return { latest_snapshot: '', byCohort: [], byOrg: [] }
  }

  const latest = parsed.map(p => p.date_key).sort().slice(-1)[0]
  const latestRows = parsed.filter(p => p.date_key === latest)

  const byCohortMap = new Map()
  const byOrgMap = new Map()

  for (const r of latestRows) {
    const cKey = r.cohort
    if (!byCohortMap.has(cKey)) byCohortMap.set(cKey, ALLSTATS_emptyWaterfallBucket_({ cohort: cKey }))
    ALLSTATS_addMetricToWaterfallBucket_(byCohortMap.get(cKey), r.metric, r.amount)

    const orgKey = (r.org_id || '') + '|' + (r.org_name || '')
    if (!byOrgMap.has(orgKey)) {
      byOrgMap.set(orgKey, ALLSTATS_emptyWaterfallBucket_({ org_id: r.org_id, org_name: r.org_name }))
    }
    ALLSTATS_addMetricToWaterfallBucket_(byOrgMap.get(orgKey), r.metric, r.amount)
  }

  const byCohort = Array.from(byCohortMap.values()).sort((a, b) => String(a.cohort || '').localeCompare(String(b.cohort || '')))
  const byOrg = Array.from(byOrgMap.values()).sort((a, b) => {
    if (b.eom !== a.eom) return b.eom - a.eom
    return String(a.org_name || '').localeCompare(String(b.org_name || ''))
  })

  return { latest_snapshot: latest, byCohort, byOrg }
}

function ALLSTATS_buildSubscriptionChurnMetrics_(orgSubscriptions, manualBySubId, posthogOrgSubsRows) {
  const ACTIVE_STATUSES = new Set(['active', 'trialing', 'past_due', 'unpaid'])
  const byOrg = new Map()
  const activeSubCountByOrg = new Map()
  const activeSubIdsByOrg = new Map()
  const latestPosthogSubById = new Map()
  const nowMs = Date.now()
  const staleGraceMs = 2 * 24 * 60 * 60 * 1000

  for (const r of (posthogOrgSubsRows || [])) {
    const orgId = ALLSTATS_str_(r.app_org_id || r.org_id)
    if (!orgId) continue

    const subId =
      ALLSTATS_str_(r.stripe_subscription_id) ||
      ALLSTATS_str_(r.subscription_id) ||
      ALLSTATS_str_(r.id)
    if (!subId) continue
    const status = ALLSTATS_str_(r.status).toLowerCase()
    const currentPeriodEnd = ALLSTATS_str_(r.current_period_end)
    const updatedTs = ALLSTATS_toDateOrNull_(r.updated_at || r.pulled_at || r.created_at || r.current_period_end)
    const ts = updatedTs ? updatedTs.getTime() : 0

    const prev = latestPosthogSubById.get(subId)
    if (!prev || ts >= prev.ts) {
      latestPosthogSubById.set(subId, {
        org_id: orgId,
        sub_id: subId,
        status,
        current_period_end: currentPeriodEnd,
        ts
      })
    }
  }

  for (const rec of latestPosthogSubById.values()) {
    const manual = rec.sub_id ? (manualBySubId && manualBySubId.get(rec.sub_id)) : null
    if (manual && manual.excludeInternal) continue
    if (!ACTIVE_STATUSES.has(rec.status)) continue

    // Ignore stale "active" rows whose period end is already in the past.
    const cpe = ALLSTATS_toDateOrNull_(rec.current_period_end)
    if (cpe && cpe.getTime() < (nowMs - staleGraceMs)) continue

    activeSubCountByOrg.set(rec.org_id, (activeSubCountByOrg.get(rec.org_id) || 0) + 1)
    if (!activeSubIdsByOrg.has(rec.org_id)) activeSubIdsByOrg.set(rec.org_id, [])
    activeSubIdsByOrg.get(rec.org_id).push(rec.sub_id)
  }

  for (const r of (orgSubscriptions || [])) {
    const orgId = ALLSTATS_str_(r.app_org_id || r.org_id)
    if (!orgId) continue

    const subId =
      ALLSTATS_str_(r.stripe_subscription_id) ||
      ALLSTATS_str_(r.latest_subscription_id) ||
      ALLSTATS_str_(r.subscription_id)
    const manual = subId ? (manualBySubId && manualBySubId.get(subId)) : null
    if (manual && manual.excludeInternal) continue

    const row = {
      org_id: orgId,
      org_name: ALLSTATS_str_(r.org_name),
      stripe_customer_email: ALLSTATS_str_(r.customer_email || r.billing_email || r.email),
      stripe_subscription_id: subId,
      status: ALLSTATS_str_(r.status).toLowerCase(),
      first_payment_date: ALLSTATS_str_(r.first_payment_at),
      churn_date: ALLSTATS_str_(r.churn_date),
      churn_reason: ALLSTATS_str_(r.churn_reason || r.cancellation_reason),
      arr_yearly: (() => {
        const discounted = ALLSTATS_num_(r.amount_after_discounts_yearly)
        const gross = ALLSTATS_num_(r.amount_yearly)
        return discounted > 0 ? discounted : gross
      })(),
      seat_count: ALLSTATS_safeInt_(r.quantity_total),
      current_period_end: ALLSTATS_str_(r.current_period_end),
      cancel_at_period_end: ALLSTATS_toBool_(r.cancel_at_period_end),
      plan_name: ALLSTATS_str_(r.plan_name),
      subscription_created_at: ALLSTATS_str_(r.subscription_created_at_date)
    }

    const existing = byOrg.get(orgId)
    if (!existing) {
      byOrg.set(orgId, row)
      continue
    }

    const existingTs = ALLSTATS_toDateOrNull_(existing.subscription_created_at)
    const rowTs = ALLSTATS_toDateOrNull_(row.subscription_created_at)
    if (rowTs && (!existingTs || rowTs.getTime() > existingTs.getTime())) byOrg.set(orgId, row)
  }

  let paidBaseOrgs = 0
  let churnedOrgs = 0
  let canceledOrgs = 0
  let scheduledOrgs = 0
  let churnedArr = 0
  let churnedSeats = 0
  const rows = []

  for (const r of byOrg.values()) {
    const hasPaid = !!r.first_payment_date
    if (!hasPaid) continue
    paidBaseOrgs += 1

    const activeSubCount = activeSubCountByOrg.get(r.org_id) || 0
    const hasAnyActiveSub = activeSubCount > 0
    const isCanceledNow = r.status === 'canceled' || r.status === 'cancelled'
    const churnEffectiveDate = r.churn_date || r.current_period_end || ''
    const isScheduledCancel = !isCanceledNow && r.cancel_at_period_end && !!churnEffectiveDate
    const isTrulyChurned = !hasAnyActiveSub && (isCanceledNow || (!!r.churn_date && !r.cancel_at_period_end))
    const isChurnRisk = isTrulyChurned || isScheduledCancel
    if (!isChurnRisk) continue

    churnedOrgs += 1
    if (isTrulyChurned) canceledOrgs += 1
    if (isScheduledCancel) scheduledOrgs += 1
    churnedArr += ALLSTATS_num_(r.arr_yearly)
    churnedSeats += ALLSTATS_safeInt_(r.seat_count)

    rows.push({
      org_id: r.org_id,
      org_name: r.org_name,
      stripe_customer_email: r.stripe_customer_email,
      stripe_subscription_id: r.stripe_subscription_id,
      status: r.status,
      active_subscription_count: activeSubCount,
      active_subscription_ids: (activeSubIdsByOrg.get(r.org_id) || []).join(', '),
      churn_type: isTrulyChurned ? 'canceled' : 'scheduled_cancel',
      churn_reason: r.churn_reason || (isScheduledCancel ? 'cancel_at_period_end' : (isTrulyChurned ? 'canceled' : '')),
      churn_effective_date: churnEffectiveDate,
      subscription_end_date: isTrulyChurned ? churnEffectiveDate : (r.current_period_end || churnEffectiveDate || ''),
      current_period_end: r.current_period_end || '',
      first_payment_date: r.first_payment_date,
      seat_count: ALLSTATS_safeInt_(r.seat_count),
      arr_yearly: ALLSTATS_num_(r.arr_yearly),
      plan_name: r.plan_name
    })
  }

  rows.sort((a, b) => {
    const da = ALLSTATS_toDateOrNull_(a.churn_effective_date)
    const db = ALLSTATS_toDateOrNull_(b.churn_effective_date)
    const ta = da ? da.getTime() : 0
    const tb = db ? db.getTime() : 0
    if (tb !== ta) return tb - ta
    return String(a.org_name || '').localeCompare(String(b.org_name || ''))
  })

  return {
    paid_base_orgs: paidBaseOrgs,
    churned_orgs: churnedOrgs,
    canceled_orgs: canceledOrgs,
    scheduled_churn_orgs: scheduledOrgs,
    churned_arr: churnedArr,
    churned_seats: churnedSeats,
    logo_churn_rate: paidBaseOrgs > 0 ? (churnedOrgs / paidBaseOrgs) : 0,
    rows
  }
}

function ALLSTATS_emptyWaterfallBucket_(base) {
  return Object.assign({}, base || {}, {
    som: 0,
    new_customer: 0,
    upgrade: 0,
    downgrade: 0,
    churn: 0,
    eom: 0
  })
}

function ALLSTATS_addMetricToWaterfallBucket_(bucket, metric, amount) {
  if (!bucket) return
  const m = String(metric || '').toLowerCase()
  if (m === 'som') bucket.som += amount
  else if (m === 'new') bucket.new_customer += amount
  else if (m === 'upgrade') bucket.upgrade += amount
  else if (m === 'downgrade') bucket.downgrade += amount
  else if (m === 'churn') bucket.churn += amount
  else if (m === 'eom') bucket.eom += amount
}

function ALLSTATS_buildFreeTrialList_(orgs) {
  return orgs
    .filter(o => o.stage === 'Free Trial')
    .filter(o => !o.has_promo_conversion_eligible)
    .filter(o => o.free_trial_days === '' || o.free_trial_days <= 14)
    .map(o => [
      o.org_id,
      o.org_name,
      o.customer_name,
      o.customer_email,
      o.free_trial_days,
      o.free.seats,
      o.free.mrr,
      o.free.arr,
      o.subscription_count,
      o.subscription_ids_text
    ])
}

function ALLSTATS_buildPromoTrialList_(orgs) {
  return orgs
    .filter(o => !o.has_paid && o.has_promo)
    .filter(o => o.has_promo_conversion_eligible)
    .map(o => [
      o.org_id,
      o.org_name,
      o.customer_name,
      o.customer_email,
      o.promo.seats,
      o.promo.mrr,
      o.promo.arr,
      ALLSTATS_pickPromoSourceLabel_(o),
      o.subscription_count,
      o.subscription_ids_text
    ])
}

function ALLSTATS_buildPaidList_(orgs) {
  return orgs
    .filter(o => o.has_paid_us)
    .map(o => [
      o.org_id,
      o.org_name,
      o.customer_name,
      o.customer_email,
      o.paid.seats,
      o.paid.mrr,
      o.paid.arr,
      o.earliest_first_payment_at || '',
      o.subscription_count,
      o.subscription_ids_text
    ])
}

function ALLSTATS_pickPromoSourceLabel_(org) {
  const labels = []
  const redemptionCodes = ALLSTATS_str_(org.promo_redemption_codes_text)
  const redemptionCodeIds = ALLSTATS_str_(org.promo_redemption_code_ids_text)
  if (redemptionCodes) labels.push(redemptionCodes)
  else if (redemptionCodeIds) labels.push(redemptionCodeIds)
  if (org.has_manual_trial_extension) labels.push('manual trial_extended')
  return labels.join(' | ')
}

function ALLSTATS_renderPromoAudit_(ss, conversion) {
  const orgs = buildConversionOrgList_(ss)
  const promoOrgs = orgs.filter(o => o.has_promo)

  // Sort: paid first, then trialing, then expired; within each group sort by org name
  const statusRank = o => {
    if (o.first_payment_at) return 0   // converted
    if (o.is_currently_trialing) return 1  // still trialing
    return 2  // expired
  }
  promoOrgs.sort((a, b) => {
    const ra = statusRank(a), rb = statusRank(b)
    if (ra !== rb) return ra - rb
    return String(a.org_name || '').localeCompare(String(b.org_name || ''))
  })

  const headers = [
    'org_name', 'org_id', 'owner_email', 'org_status',
    'promo_codes', 'extension_source',
    'org_created_at', 'trial_ends_at', 'first_payment_at',
    'converted', 'still_trialing'
  ]

  const rows = promoOrgs.map(o => [
    o.org_name,
    o.org_id,
    o.owner_email,
    o.org_status,
    o.promo_codes,
    o.trial_extension_source || '',
    o.org_created_at || '',
    o.trial_ends_at || '',
    o.first_payment_at || '',
    o.first_payment_at ? 'Yes' : 'No',
    o.is_currently_trialing ? 'Yes' : 'No'
  ])

  const sh = ALLSTATS_getOrCreateSheet_(ss, 'Promo Audit')
  sh.clear()
  try { sh.clearNotes() } catch (e) {}

  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#F3F4F6')

  if (rows.length) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows)
    // Format date columns: org_created_at=7, trial_ends_at=8, first_payment_at=9
    sh.getRange(2, 7, rows.length, 3).setNumberFormat('yyyy-mm-dd')
  }

  sh.autoResizeColumns(1, headers.length)
}

function ALLSTATS_renderAllStatsSheet_(sheet, data) {
  sheet.clear()

  let row = 1

  row = ALLSTATS_writeTitle_(sheet, row, 'All the Stats', data.generatedAt)

  row = ALLSTATS_writeSection_(sheet, row, 'ARR / MRR Summary (Org Level)')
  const arrSummaryRow = row
  row = ALLSTATS_writeTable_(sheet, row, 1,
    ['Metric', 'Paid', 'Intent to Pay', 'Trialing'],
    [
      ['ARR', data.metrics.paidDefined.arr, data.metrics.promoTrialDefined.arr, data.metrics.freeTrialDefined.arr],
      ['MRR', data.metrics.paidDefined.mrr, data.metrics.promoTrialDefined.mrr, data.metrics.freeTrialDefined.mrr],
      ['Firms', data.metrics.paidDefined.firms, data.metrics.promoTrialDefined.firms, data.metrics.freeTrialDefined.firms],
      ['Seats', data.metrics.paidDefined.seats, data.metrics.promoTrialDefined.seats, data.metrics.freeTrialDefined.seats]
    ],
    {
      currencyRowsAcross: [1, 2],
      intRowsAcross: [3, 4]
    }
  )

  // Seat breakdown by type and stage — placed to the right of ARR summary
  const seatCol = 6
  const bs = data.seatTypes.byStage
  const paidTotal = bs.paid.full + bs.paid.lite
  const intentTotal = bs.intent.full + bs.intent.lite
  const trialingTotal = bs.trialing.full + bs.trialing.lite
  ALLSTATS_writeTable_(sheet, arrSummaryRow, seatCol,
    ['Seat Type', 'Paid', 'Intent to Pay', 'Trialing'],
    [
      ['Full Seats', bs.paid.full, bs.intent.full, bs.trialing.full],
      ['Lite Seats', bs.paid.lite, bs.intent.lite, bs.trialing.lite],
      ['Total', paidTotal, intentTotal, trialingTotal],
      ['% Full', paidTotal ? bs.paid.full / paidTotal : 0, intentTotal ? bs.intent.full / intentTotal : 0, trialingTotal ? bs.trialing.full / trialingTotal : 0],
      ['% Lite', paidTotal ? bs.paid.lite / paidTotal : 0, intentTotal ? bs.intent.lite / intentTotal : 0, trialingTotal ? bs.trialing.lite / trialingTotal : 0]
    ],
    { intCols: [2, 3, 4], pctRowsAcross: [4, 5] }
  )

  row = ALLSTATS_writeSection_(sheet, row, '# of Paying Firms / Seats by Stage')
  row = ALLSTATS_writeTable_(sheet, row, 1,
    ['Stage', 'Firms', 'Seats', 'Avg Seats / Firm'],
    [
      [
        'Paid',
        data.metrics.paidDefined.firms,
        data.metrics.paidDefined.seats,
        data.metrics.paidDefined.firms ? (data.metrics.paidDefined.seats / data.metrics.paidDefined.firms) : 0
      ],
      [
        'Intent to Pay',
        data.metrics.promoTrialDefined.firms,
        data.metrics.promoTrialDefined.seats,
        data.metrics.promoTrialDefined.firms ? (data.metrics.promoTrialDefined.seats / data.metrics.promoTrialDefined.firms) : 0
      ],
      [
        'Trialing',
        data.metrics.freeTrialDefined.firms,
        data.metrics.freeTrialDefined.seats,
        data.metrics.freeTrialDefined.firms ? (data.metrics.freeTrialDefined.seats / data.metrics.freeTrialDefined.firms) : 0
      ]
    ],
    { intCols: [2, 3], decimalCols: [4] }
  )

  row = ALLSTATS_writeSection_(sheet, row, 'Avg Revenue Metrics')
  row = ALLSTATS_writeTable_(sheet, row, 1,
    ['Metric', 'Value'],
    [
      ['Avg ARR per firm (Paid, excl. managed)', data.metrics.avgRevenuePerFirmPaidExclManaged],
      ['Avg ARR per firm (Paid)', data.metrics.avgRevenuePerFirmPaid],
      ['Avg ARR per seat (Paid)', data.metrics.avgRevenuePerSeat],
      ['Avg seat count per firm (Paid)', data.metrics.avgSeatCountPerFirmPaid]
    ],
    { currencyRows: [1, 2, 3], intRows: [4] }
  )

  row = ALLSTATS_writeSection_(sheet, row, 'Conversion + Retention')
  row = ALLSTATS_writeTable_(sheet, row, 1,
    ['Metric', 'Value', 'Detail', 'Potential'],
    [
      [
        'Sign up to paid conversion',
        data.conversion.signupToPaidRate,
        data.conversion.paidUs + ' / ' + data.conversion.signupResolved,
        data.conversion.signupTrialPotential + ' still trialing (' + ALLSTATS_fmtPct_(data.conversion.signupPotentialMaxRate) + ' max) | ' + data.conversion.totalSignedUp + ' total signups'
      ],
      [
        'Promo trial to paid conversion',
        data.conversion.promoToPaidRate,
        data.conversion.promoToPaid + ' / ' + data.conversion.promoResolved,
        data.conversion.promoTrialPotential + ' still trialing (' + ALLSTATS_fmtPct_(data.conversion.promoPotentialMaxRate) + ' max) | ' + data.conversion.promoPool + ' total promo signups'
      ],
      ['NRR (latest snapshot)', data.retention.nrr, data.retention.latest_snapshot || '', ''],
      ['GRR (latest snapshot)', data.retention.grr, data.retention.latest_snapshot || '', ''],
      ['Gross ARR churn rate (snapshot)', data.retention.gross_arr_churn_rate, 'Base ARR ' + ALLSTATS_fmtMoney_(data.retention.base_bom_arr), ''],
      ['Full ARR churn rate (snapshot)', data.retention.full_arr_churn_rate, 'Churned ARR ' + ALLSTATS_fmtMoney_(data.retention.churned_arr), ''],
      ['Churned orgs (subscription-based)', data.subscriptionChurn.churned_orgs, data.subscriptionChurn.canceled_orgs + ' canceled | ' + data.subscriptionChurn.scheduled_churn_orgs + ' scheduled cancel', ''],
      ['Churned ARR (subscription-based)', data.subscriptionChurn.churned_arr, data.subscriptionChurn.churned_orgs ? ('Avg ' + ALLSTATS_fmtMoney_(data.subscriptionChurn.churned_arr / data.subscriptionChurn.churned_orgs) + ' per churn org') : '', ''],
      ['Churned seats (subscription-based)', data.subscriptionChurn.churned_seats, data.subscriptionChurn.churned_orgs ? ('Avg ' + (data.subscriptionChurn.churned_seats / data.subscriptionChurn.churned_orgs).toFixed(1) + ' seats per churn org') : '', '']
    ],
    { pctRows: [1, 2, 3, 4, 5, 6], intRows: [7, 9], currencyRows: [8] }
  )

  const churnAuditRows = (data.subscriptionChurn && data.subscriptionChurn.rows) ? data.subscriptionChurn.rows : []

  if (churnAuditRows.length) {
    row = ALLSTATS_writeSection_(sheet, row, 'Churned Orgs Audit (Subscription-Based)')
    row = ALLSTATS_writeTable_(sheet, row, 1,
      ['Org Name', 'Stripe Customer Email', 'Stripe Subscription ID', 'Seats', 'ARR', 'Status', 'Churn Type', 'Churn Reason', 'Subscription End Date', 'First Payment', 'Plan'],
      churnAuditRows.map(o => [
        o.org_name, o.stripe_customer_email || '', o.stripe_subscription_id || '', o.seat_count || 0, o.arr_yearly || 0, o.status || '', o.churn_type || '', o.churn_reason || '',
        o.subscription_end_date || '', o.first_payment_date || '', o.plan_name || ''
      ]),
      { intCols: [4], currencyCols: [5], dateCols: [9, 10] }
    )
  }

  row = ALLSTATS_writeSection_(sheet, row, 'Net New ARR by Month')
  row = ALLSTATS_writeTable_(sheet, row, 1,
    ['Month', 'BOM ARR', 'EOM ARR', 'Net New ARR', 'New Orgs', 'Upgrades', 'Downgrades', 'Churned Orgs', 'Churn ARR'],
    (data.netNew.rows || []).map(r => [
      r.month,
      r.bom_arr,
      r.eom_arr,
      r.net_new_arr,
      r.new_orgs,
      r.upgrades,
      r.downgrades,
      r.churned_orgs,
      r.churn_arr
    ]),
    {
      currencyCols: [2, 3, 4, 6, 7, 9],
      intCols: [5, 8]
    }
  )

  row = ALLSTATS_writeSection_(sheet, row, 'Waterfall (Latest Snapshot by Cohort)')
  row = ALLSTATS_writeTable_(sheet, row, 1,
    ['Snapshot', 'Cohort', 'SOM', 'New', 'Upgrade', 'Downgrade', 'Churn', 'EOM'],
    (data.waterfall.byCohort || []).map(r => [
      data.waterfall.latest_snapshot || '',
      r.cohort,
      r.som,
      r.new_customer,
      r.upgrade,
      r.downgrade,
      r.churn,
      r.eom
    ]),
    { currencyCols: [3, 4, 5, 6, 7, 8] }
  )

  row = ALLSTATS_writeSection_(sheet, row, 'Waterfall (Latest Snapshot by Org)')
  row = ALLSTATS_writeTable_(sheet, row, 1,
    ['Snapshot', 'Org ID', 'Org Name', 'SOM', 'New', 'Upgrade', 'Downgrade', 'Churn', 'EOM'],
    (data.waterfall.byOrg || []).map(r => [
      data.waterfall.latest_snapshot || '',
      r.org_id,
      r.org_name,
      r.som,
      r.new_customer,
      r.upgrade,
      r.downgrade,
      r.churn,
      r.eom
    ]),
    { currencyCols: [4, 5, 6, 7, 8, 9] }
  )

  row = ALLSTATS_writeSection_(sheet, row, 'Trialing (No Payment Method)')
  row = ALLSTATS_writeTable_(sheet, row, 1,
    ['Org ID', 'Org Name', 'Customer Name', 'Customer Email', 'Days in Trial', 'Seats', 'MRR', 'ARR', 'Subscriptions', 'Subscription IDs'],
    data.freeTrialList,
    { intCols: [5, 6, 9], currencyCols: [7, 8] }
  )

  row = ALLSTATS_writeSection_(sheet, row, 'Promo Trial (Used Promo Code)')
  row = ALLSTATS_writeTable_(sheet, row, 1,
    ['Org ID', 'Org Name', 'Customer Name', 'Customer Email', 'Seats', 'MRR', 'ARR', 'Promo Codes', 'Subscriptions', 'Subscription IDs'],
    data.promoTrialList,
    { intCols: [5, 9], currencyCols: [6, 7] }
  )

  row = ALLSTATS_writeSection_(sheet, row, 'Paid (Paid Us)')
  row = ALLSTATS_writeTable_(sheet, row, 1,
    ['Org ID', 'Org Name', 'Customer Name', 'Customer Email', 'Seats', 'MRR', 'ARR', 'First Payment At', 'Subscriptions', 'Subscription IDs'],
    data.paidList,
    { intCols: [5, 9], currencyCols: [6, 7], datetimeCols: [8] }
  )

  sheet.setFrozenRows(2)
  sheet.autoResizeColumns(1, 10)
}

function ALLSTATS_writeTitle_(sheet, startRow, title, generatedAt) {
  const tz = Session.getScriptTimeZone()
  const generated = Utilities.formatDate(generatedAt || new Date(), tz, 'yyyy-MM-dd HH:mm:ss')

  sheet.getRange(startRow, 1, 1, 10).merge()
  sheet.getRange(startRow, 1).setValue(title)
    .setFontWeight('bold')
    .setFontSize(18)
    .setBackground(ALL_STATS_CFG.TITLE_BG)
    .setFontColor(ALL_STATS_CFG.TITLE_FG)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')

  sheet.getRange(startRow + 1, 1, 1, 10).merge()
  sheet.getRange(startRow + 1, 1).setValue('Generated at: ' + generated)
    .setFontSize(10)
    .setBackground('#F9FAFB')

  return startRow + 3
}

function ALLSTATS_writeSection_(sheet, startRow, title) {
  sheet.getRange(startRow, 1, 1, 10).merge()
  sheet.getRange(startRow, 1).setValue(title)
    .setFontWeight('bold')
    .setBackground(ALL_STATS_CFG.SECTION_BG)
    .setFontColor('#111827')
  return startRow + 1
}

function ALLSTATS_writeTable_(sheet, startRow, startCol, headers, rows, opts) {
  const options = opts || {}
  const safeRows = (rows && rows.length)
    ? rows
    : [headers.map((_, i) => i === 0 ? '(none)' : '')]

  sheet.getRange(startRow, startCol, 1, headers.length).setValues([headers])
  sheet.getRange(startRow, startCol, 1, headers.length)
    .setFontWeight('bold')
    .setBackground(ALL_STATS_CFG.HEADER_BG)

  sheet.getRange(startRow + 1, startCol, safeRows.length, headers.length).setValues(safeRows)

  const dataRange = sheet.getRange(startRow + 1, startCol, safeRows.length, headers.length)
  dataRange.setVerticalAlignment('middle')

  ALLSTATS_applyTableFormats_(sheet, startRow + 1, startCol, safeRows.length, headers.length, options)

  return startRow + 1 + safeRows.length + 2
}

function ALLSTATS_applyTableFormats_(sheet, row, col, numRows, numCols, opts) {
  const options = opts || {}

  const applyCols = (cols, fmt) => {
    ;(cols || []).forEach(c => {
      if (c < 1 || c > numCols) return
      sheet.getRange(row, col + c - 1, numRows, 1).setNumberFormat(fmt)
    })
  }

  const applyRows = (rows, fmt, colIdx) => {
    ;(rows || []).forEach(r => {
      if (r < 1 || r > numRows) return
      const targetCol = colIdx || 2
      if (targetCol < 1 || targetCol > numCols) return
      sheet.getRange(row + r - 1, col + targetCol - 1, 1, 1).setNumberFormat(fmt)
    })
  }

  const applyRowsAcross = (rows, fmt, startColIdx) => {
    const startIdx = startColIdx || 2
    if (startIdx < 1 || startIdx > numCols) return
    const width = numCols - startIdx + 1
    if (width <= 0) return
    ;(rows || []).forEach(r => {
      if (r < 1 || r > numRows) return
      sheet.getRange(row + r - 1, col + startIdx - 1, 1, width).setNumberFormat(fmt)
    })
  }

  applyCols(options.currencyCols, ALL_STATS_CFG.CURRENCY_FMT)
  applyCols(options.intCols, ALL_STATS_CFG.INT_FMT)
  applyCols(options.decimalCols, ALL_STATS_CFG.DECIMAL_FMT)
  applyCols(options.pctCols, ALL_STATS_CFG.PCT_FMT)
  applyCols(options.datetimeCols, ALL_STATS_CFG.DATETIME_FMT)
  applyCols(options.dateCols, ALL_STATS_CFG.DATE_FMT)

  applyRows(options.currencyRows, ALL_STATS_CFG.CURRENCY_FMT, 2)
  applyRows(options.intRows, ALL_STATS_CFG.INT_FMT, 2)
  applyRows(options.pctRows, ALL_STATS_CFG.PCT_FMT, 2)
  applyRowsAcross(options.currencyRowsAcross, ALL_STATS_CFG.CURRENCY_FMT, 2)
  applyRowsAcross(options.intRowsAcross, ALL_STATS_CFG.INT_FMT, 2)
  applyRowsAcross(options.pctRowsAcross, ALL_STATS_CFG.PCT_FMT, 2)
}

function ALLSTATS_classifyStage_(statusRaw, hasPaymentMethod) {
  const s = String(statusRaw || '').toLowerCase()
  if (s === 'active') return 'Paid'
  if (s === 'trialing') return hasPaymentMethod ? 'Promo Trial' : 'Free Trial'
  return ''
}

function ALLSTATS_buildIndexes_(clerkUsers, clerkMems, clerkOrgs, posthogUsers, canonOrgs) {
  const orgNameByOrgId = new Map()
  for (const o of (clerkOrgs || [])) {
    const orgId = ALLSTATS_str_(o.org_id)
    if (!orgId) continue
    const name = ALLSTATS_str_(o.org_name) || ALLSTATS_str_(o.org_slug)
    if (name) orgNameByOrgId.set(orgId, name)
  }

  const canonBySubId = new Map()
  const canonByOrgId = new Map()
  const canonByClerkOrgId = new Map()
  for (const c of (canonOrgs || [])) {
    const appOrgId = ALLSTATS_str_(c.app_org_id)
    const clerkOrgId = ALLSTATS_str_(c.org_id)
    const orgId = appOrgId || clerkOrgId
    if (!orgId && !clerkOrgId) continue

    const orgName = ALLSTATS_str_(c.org_name || c.org_slug)
    const billingEmail = ALLSTATS_str_(c.billing_email)
    if (orgId) canonByOrgId.set(orgId, { orgId, appOrgId, clerkOrgId, orgName, billingEmail })
    if (clerkOrgId) canonByClerkOrgId.set(clerkOrgId, { orgId, appOrgId, clerkOrgId, orgName, billingEmail })
    if (orgName) {
      if (orgId) orgNameByOrgId.set(orgId, orgName)
      if (clerkOrgId) orgNameByOrgId.set(clerkOrgId, orgName)
      if (appOrgId) orgNameByOrgId.set(appOrgId, orgName)
    }

    const subIds = ALLSTATS_csvList_(c.stripe_subscription_ids)
    for (const subIdRaw of subIds) {
      const subId = ALLSTATS_str_(subIdRaw)
      if (!subId || canonBySubId.has(subId)) continue
      canonBySubId.set(subId, { orgId, appOrgId, clerkOrgId, orgName, billingEmail })
    }
  }

  const membershipsByEmailKey = new Map()
  for (const m of (clerkMems || [])) {
    const emailKey =
      ALLSTATS_str_(m.email_key) ||
      ALLSTATS_normalizeEmail_(ALLSTATS_str_(m.email))

    const orgId = ALLSTATS_str_(m.org_id)
    if (!emailKey || !orgId) continue

    const role = ALLSTATS_str_(m.role).toLowerCase()
    const isOwnerish =
      role.indexOf('owner') >= 0 ||
      role.indexOf('admin') >= 0 ||
      role === 'org:admin' ||
      role === 'admin' ||
      role === 'owner'

    if (!membershipsByEmailKey.has(emailKey)) membershipsByEmailKey.set(emailKey, [])
    membershipsByEmailKey.get(emailKey).push({ orgId, role, isOwnerish })
  }

  const usersByStripeSubId = new Map()
  const usersByEmailKey = new Map()

  for (const u of (clerkUsers || [])) {
    const subId = ALLSTATS_str_(u.stripe_subscription_id || u.stripeSubscriptionId)

    const email = ALLSTATS_str_(u.email)
    const emailKey = ALLSTATS_str_(u.email_key) || ALLSTATS_normalizeEmail_(email)
    if (!emailKey) continue

    const name = ALLSTATS_str_(u.name)
    const orgId = ALLSTATS_str_(u.org_id)
    const userObj = { email, emailKey, name, orgId }

    if (!usersByEmailKey.has(emailKey)) usersByEmailKey.set(emailKey, userObj)

    if (!subId) continue
    if (!usersByStripeSubId.has(subId)) usersByStripeSubId.set(subId, [])
    usersByStripeSubId.get(subId).push(userObj)
  }

  const posthogEmailKeysBySubId = new Map()
  for (const p of (posthogUsers || [])) {
    const subId = ALLSTATS_str_(p.stripe_subscription_id || p.subscription_id || p.subscription)
    if (!subId) continue

    const emailKey = ALLSTATS_str_(p.email_key) || ALLSTATS_normalizeEmail_(ALLSTATS_str_(p.email))
    if (!emailKey) continue

    if (!posthogEmailKeysBySubId.has(subId)) posthogEmailKeysBySubId.set(subId, new Set())
    posthogEmailKeysBySubId.get(subId).add(emailKey)
  }

  return {
    canonBySubId,
    canonByOrgId,
    canonByClerkOrgId,
    orgNameByOrgId,
    membershipsByEmailKey,
    usersByStripeSubId,
    usersByEmailKey,
    posthogEmailKeysBySubId
  }
}

function ALLSTATS_resolveCustomer_(stripeEmailKey, stripeSubscriptionId, idx) {
  const subId = ALLSTATS_str_(stripeSubscriptionId)

  let candidates = (subId && idx.usersByStripeSubId.has(subId))
    ? idx.usersByStripeSubId.get(subId).slice()
    : []

  if (!candidates.length && subId && idx.posthogEmailKeysBySubId.has(subId)) {
    const fromPosthog = []
    const emailKeys = Array.from(idx.posthogEmailKeysBySubId.get(subId) || [])
    for (const k of emailKeys) {
      const hit = idx.usersByEmailKey.get(k)
      if (hit) fromPosthog.push(hit)
    }
    candidates = fromPosthog
  }

  if (!candidates.length && stripeEmailKey && idx.usersByEmailKey.has(stripeEmailKey)) {
    candidates = [idx.usersByEmailKey.get(stripeEmailKey)]
  }

  if (!candidates.length) return { email: '', customerName: '', orgName: '', orgId: '', appOrgId: '' }

  let filtered = candidates
  if (stripeEmailKey) {
    const exact = candidates.filter(c => c.emailKey === stripeEmailKey)
    if (exact.length) filtered = exact
  }

  const scored = filtered.map(c => {
    const mems = idx.membershipsByEmailKey.get(c.emailKey) || []
    const hasOwnerish = mems.some(m => m.isOwnerish)
    return { user: c, hasOwnerish }
  })

  scored.sort((a, b) => {
    if (a.hasOwnerish !== b.hasOwnerish) return a.hasOwnerish ? -1 : 1
    return String(a.user.email || '').localeCompare(String(b.user.email || ''))
  })

  const picked = scored[0].user
  const mems = idx.membershipsByEmailKey.get(picked.emailKey) || []
  let orgId = ''

  if (mems.length) {
    const ownerish = mems.find(m => m.isOwnerish)
    orgId = (ownerish ? ownerish.orgId : mems[0].orgId) || ''
  }

  if (!orgId && picked.orgId) orgId = picked.orgId

  let appOrgId = ''
  if (orgId && idx.canonByOrgId && idx.canonByOrgId.has(orgId)) {
    const canon = idx.canonByOrgId.get(orgId)
    appOrgId = ALLSTATS_str_(canon && canon.appOrgId)
  } else if (orgId && idx.canonByClerkOrgId && idx.canonByClerkOrgId.has(orgId)) {
    const canon = idx.canonByClerkOrgId.get(orgId)
    appOrgId = ALLSTATS_str_(canon && canon.appOrgId)
  }
  const orgName =
    (appOrgId ? (idx.orgNameByOrgId.get(appOrgId) || '') : '') ||
    (orgId ? (idx.orgNameByOrgId.get(orgId) || '') : '')

  return {
    email: picked.email || '',
    customerName: picked.name || '',
    orgName,
    orgId,
    appOrgId
  }
}

function ALLSTATS_buildManualStripeChangesBySubId_(sheet) {
  const out = new Map()
  if (!sheet) return out

  const rows = ALLSTATS_readSheetObjects_(sheet, 1)
  for (const r of rows) {
    const subId =
      ALLSTATS_str_(r.subscription_id) ||
      ALLSTATS_str_(r.stripe_subscription_id) ||
      ALLSTATS_str_(r.subscription)

    if (!subId) continue

    const excludeReason = ALLSTATS_str_(r.exclude_reason).toLowerCase()
    const EXCLUDE_REASONS = new Set(['internal', 'partner', 'free subscription'])
    if (!EXCLUDE_REASONS.has(excludeReason)) continue
    out.set(subId, { excludeInternal: true })
  }

  return out
}

function ALLSTATS_isInternalOrg_(r) {
  const name = ALLSTATS_str_(r.org_name || r.org_slug || r.org).toLowerCase()
  if (name.includes('ping') || name.includes('test')) return true
  const email = ALLSTATS_str_(r.customer_email || r.billing_email || r.email).toLowerCase()
  if (email.endsWith('@pingassistant.com')) return true
  return false
}

function ALLSTATS_computeMrrArr_(amount, interval, intervalCount) {
  const amt = Number(amount || 0) || 0
  const intv = String(interval || '').toLowerCase().trim()
  const count = Math.max(1, Number(intervalCount || 1) || 1)
  if (intv === 'year' || intv === 'annual' || intv === 'yr') {
    const arr = amt / count
    return { arr, mrr: arr / 12 }
  }
  if (intv === 'month' || intv === 'mo') {
    const arr = amt * (12 / count)
    return { mrr: arr / 12, arr }
  }
  return { mrr: amt, arr: amt * 12 }
}

function ALLSTATS_moneyAmount_(v) {
  if (v === null || v === undefined || v === '') return 0
  const n = ALLSTATS_num_(v)
  if (!isFinite(n)) return 0
  return Math.round(n * 100) / 100
}

function ALLSTATS_buildDiscountContextNow_(row, opts) {
  const cfg = opts || {}
  const amountRaw = ALLSTATS_moneyAmount_(cfg.amountRaw)
  const asOfDate = (cfg.asOfDate instanceof Date && !isNaN(cfg.asOfDate.getTime()))
    ? cfg.asOfDate
    : new Date()

  const details = ALLSTATS_parseDiscountDetails_(row)
  const active = details.filter(d => ALLSTATS_isDiscountActiveNow_(d, asOfDate))

  let amount = amountRaw
  const promoCodes = []
  for (const d of active) {
    const pct = ALLSTATS_num_(d.percent_off)
    if (pct > 0) {
      const bounded = Math.max(0, Math.min(100, pct))
      amount *= (1 - bounded / 100)
    }
    const amountOff = ALLSTATS_num_(d.amount_off)
    if (amountOff > 0) amount -= amountOff
    const promo = ALLSTATS_str_(d.promotion_code)
    if (promo) promoCodes.push(promo)
  }

  return {
    amount: Math.max(0, amount),
    promoCodes: Array.from(new Set(promoCodes)),
    activeDiscountCount: active.length
  }
}

function ALLSTATS_parseDiscountDetails_(row) {
  const out = []
  const r = row || {}

  const detailsJson = ALLSTATS_str_(r.discount_details_json)
  if (detailsJson) {
    try {
      const parsed = JSON.parse(detailsJson)
      if (Array.isArray(parsed)) {
        parsed.forEach((d, i) => {
          if (!d || typeof d !== 'object') return
          out.push({
            index: i + 1,
            percent_off: ALLSTATS_num_(d.percent_off),
            amount_off: ALLSTATS_num_(d.amount_off),
            duration: ALLSTATS_str_(d.duration).toLowerCase(),
            duration_in_months: ALLSTATS_num_(d.duration_in_months),
            start_at: ALLSTATS_str_(d.start_at),
            end_at: ALLSTATS_str_(d.end_at),
            promotion_code: ALLSTATS_str_(d.promotion_code)
          })
        })
      }
    } catch (e) {}
  }
  if (out.length) return out

  const pctAll = ALLSTATS_csvList_(r.discount_percent_all)
  const amtAll = ALLSTATS_csvList_(r.discount_amount_off_all)
  const durAll = ALLSTATS_csvList_(r.discount_duration_all)
  const durMonthsAll = ALLSTATS_csvList_(r.discount_duration_months_all)
  const startAll = ALLSTATS_csvList_(r.discount_start_at_all)
  const endAll = ALLSTATS_csvList_(r.discount_end_at_all)
  const promoAll = ALLSTATS_csvList_(r.promo_code_all)
  const n = Math.max(
    pctAll.length,
    amtAll.length,
    durAll.length,
    durMonthsAll.length,
    startAll.length,
    endAll.length,
    promoAll.length
  )
  for (let i = 0; i < n; i++) {
    out.push({
      index: i + 1,
      percent_off: ALLSTATS_num_(pctAll[i]),
      amount_off: ALLSTATS_num_(amtAll[i]),
      duration: ALLSTATS_str_(durAll[i]).toLowerCase(),
      duration_in_months: ALLSTATS_num_(durMonthsAll[i]),
      start_at: ALLSTATS_str_(startAll[i]),
      end_at: ALLSTATS_str_(endAll[i]),
      promotion_code: ALLSTATS_str_(promoAll[i])
    })
  }
  if (out.length) return out

  const pct = ALLSTATS_num_(r.discount_percent)
  const amt = ALLSTATS_num_(r.discount_amount_off)
  const duration = ALLSTATS_str_(r.discount_duration).toLowerCase()
  const durationMonths = ALLSTATS_num_(r.discount_duration_months)
  if (pct <= 0 && amt <= 0) return []

  return [{
    index: 1,
    percent_off: pct,
    amount_off: amt,
    duration,
    duration_in_months: durationMonths,
    start_at: ALLSTATS_str_(r.discount_start_at || r.first_payment_at || r.created_at),
    end_at: ALLSTATS_str_(r.discount_end_at),
    promotion_code: ALLSTATS_str_(r.promo_code)
  }]
}

function ALLSTATS_isDiscountActiveNow_(detail, asOfDate) {
  const asOf = (asOfDate instanceof Date && !isNaN(asOfDate.getTime()))
    ? asOfDate
    : new Date()
  const d = detail || {}
  const start = ALLSTATS_toDateOrNull_(d.start_at)
  const end = ALLSTATS_toDateOrNull_(d.end_at)
  const duration = ALLSTATS_str_(d.duration).toLowerCase()
  const hasValue = ALLSTATS_num_(d.percent_off) > 0 || ALLSTATS_num_(d.amount_off) > 0
  if (!hasValue) return false

  if (start && asOf < start) return false
  if (end) return asOf < end
  if (duration === 'forever') return true
  if (duration === 'repeating' || duration === 'once') {
    if (!start) return false
    const months = Math.max(1, ALLSTATS_num_(d.duration_in_months) || 1)
    const until = new Date(start.getTime())
    until.setUTCMonth(until.getUTCMonth() + months)
    return asOf < until
  }
  return false
}

function ALLSTATS_csvList_(v) {
  const s = ALLSTATS_str_(v)
  if (!s) return []
  return s.split(',').map(x => ALLSTATS_str_(x))
}

function ALLSTATS_toDateOrNull_(v) {
  if (!v) return null
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v
  const d = new Date(String(v || '').trim())
  return isNaN(d.getTime()) ? null : d
}

function ALLSTATS_isoToDateOrNull_(iso) {
  const s = String(iso || '').trim()
  if (!s) return null
  const d = new Date(s)
  return isNaN(d.getTime()) ? null : d
}

function ALLSTATS_readSheetObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1) return []

  const header = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())
  const range = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol)
  const data = range.getValues()
  const display = range.getDisplayValues()

  return data.map((r, ri) => {
    const obj = {}
    header.forEach((h, i) => {
      if (!h) return
      const v = r[i]
      obj[ALLSTATS_key_(h)] = (v instanceof Date) ? display[ri][i] : v
    })
    return obj
  })
}

function ALLSTATS_key_(h) {
  return String(h || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_')
}

function ALLSTATS_str_(v) {
  if (v === null || v === undefined) return ''
  return String(v).trim()
}

function ALLSTATS_num_(v) {
  if (v === null || v === undefined || v === '') return 0
  if (typeof v === 'number') return v
  const s = String(v).replace(/[^0-9.\-]/g, '').trim()
  const n = Number(s)
  return isNaN(n) ? 0 : n
}

function ALLSTATS_safeInt_(v) {
  const n = Number(v)
  if (isNaN(n) || !isFinite(n)) return 0
  return Math.max(0, Math.floor(n))
}

function ALLSTATS_toBool_(v) {
  if (v === true) return true
  if (typeof v === 'number') return v === 1
  const s = String(v || '').trim().toLowerCase()
  return s === 'true' || s === '1' || s === 'yes' || s === 'y'
}

function ALLSTATS_normalizeEmail_(email) {
  if (typeof normalizeEmail === 'function') return normalizeEmail(email)
  return String(email || '').trim().toLowerCase()
}

function ALLSTATS_fmtMoney_(n) {
  const x = Number(n || 0)
  if (!isFinite(x)) return '$0.00'
  return '$' + x.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
}

function ALLSTATS_fmtPct_(n) {
  const x = Number(n || 0)
  if (!isFinite(x)) return '0.0%'
  return (x * 100).toFixed(1) + '%'
}

function ALLSTATS_getOrCreateSheet_(ss, name) {
  if (typeof getOrCreateSheetCompat_ === 'function') return getOrCreateSheetCompat_(ss, name)
  if (typeof getOrCreateSheet === 'function') {
    try { return getOrCreateSheet(ss, name) } catch (e1) {}
    try { return getOrCreateSheet(name) } catch (e2) {}
  }
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function ALLSTATS_lockWrap_(name, fn) {
  if (typeof fn !== 'function') throw new Error('ALLSTATS_lockWrap_: fn must be a function')
  const lock = LockService.getScriptLock()
  const ok = lock.tryLock(300000)
  if (!ok) throw new Error('Could not acquire lock: ' + name)
  try { return fn() } finally { lock.releaseLock() }
}

function ALLSTATS_writeSyncLog_(step, status, rowsIn, rowsOut, seconds, error) {
  if (typeof writeSyncLogCompat_ === 'function') return writeSyncLogCompat_(step, status, rowsIn, rowsOut, seconds, error)
  if (typeof writeSyncLog === 'function') return writeSyncLog(step, status, rowsIn, rowsOut, seconds, error || '')
  Logger.log('[SYNCLOG missing] ' + step + ' ' + status)
}
