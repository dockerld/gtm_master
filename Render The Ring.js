/**************************************************************
 * render_ring_view()
 *
 * Builds a money-focused view sheet: "The Ring"
 *
 * Layout:
 * - Big KPIs on top:
 *    - Paid: ARR, Subscriptions, Total Seats
 *    - Intent to Pay: ARR, Subscriptions, Total Seats
 *    - Trialing: ARR, Subscriptions, Total Seats
 * - Table headers on Row 3 starting Col B
 * - Table data starts Row 4 starting Col B
 *
 * Source:
 * - raw_stripe_subscriptions (header row = 1)
 * - Manual Stripe Changes (optional manual overrides)
 * - raw_posthog_user_metrics (fallback subscription->email mapping)
 * - org_subscription_info (subscription->org + trial timing mapping source)
 * - promo_redemptions (app+stripe promo redemptions by org/location)
 *
 * Rules:
 * - KPI classification considers active + trialing subscriptions
 * - Exclude subscriptions listed in Manual Stripe Changes only when
 *   exclude_reason == "internal" (case-insensitive exact match)
 * - Display status labels:
 *     paid bucket -> "Paid"
 *     intent bucket -> "Intent to Pay"
 * - Exclude subscriptions where:
 *     discount_percent == 100 AND discount_duration == 'forever'
 * - Total Seats comes from quantity_total
 *
 * ARR/MRR calculation:
 * - If interval == "year":  ARR = amount, MRR = amount / 12
 * - If interval == "month": MRR = amount, ARR = amount * 12
 *
 * Discount display:
 * - Show only discounts currently active "in the moment"
 * - Show all active Stripe promo codes on the subscription
 * - Also append app promo codes redeemed by the mapped org
 *
 * Enrichment (Customer Name + Org Name):
 * - Match Stripe stripe_subscription_id -> raw_clerk_users.stripe_subscription_id
 * - Pick a “best” user:
 *     1) If Stripe customer email matches a candidate user email, prefer it
 *     2) Else prefer a candidate who is an owner/admin in memberships
 *     3) Else first candidate
 * - Org Name:
 *     - From raw_clerk_memberships by email_key -> org_id (prefer owner/admin)
 *     - Then raw_clerk_orgs org_id -> org_name
 *
 * NEW:
 * - Adds "First Payment At", "Sign Up Date", and "Trial Days Remaining"
 **************************************************************/

const RING_CFG = {
  SHEET_NAME: 'The Ring',
  INPUT_SHEET: 'raw_stripe_subscriptions',
  MANUAL_CHANGES_SHEET: 'Manual Stripe Changes',

  // Clerk enrichment sources
  CANON_ORGS_SHEET: 'canon_orgs',
  CLERK_USERS_SHEET: 'raw_clerk_users',
  CLERK_MEMBERSHIPS_SHEET: 'raw_clerk_memberships',
  CLERK_ORGS_SHEET: 'raw_clerk_orgs',
  POSTHOG_USERS_SHEET: 'raw_posthog_user_metrics',
  POSTHOG_ORG_SUBS_SHEET: 'org_subscription_info',
  POSTHOG_PROMO_REDEMPTIONS_SHEET: 'promo_redemptions',

  // Ring layout
  KPI_ROW_LABEL: 1,
  KPI_ROW_VALUE: 2,

  HEADER_ROW: 3,
  START_COL: 2,        // Col B
  DATA_START_ROW: 4,

  // KPI blocks
  // Left -> right order:
  // 1) Paid
  // 2) Intent to Pay
  // 3) Trialing
  KPI_COLS: {
    PAID: {
      ARR: 2,            // B
      SUBSCRIPTIONS: 3,  // C
      TOTAL_SEATS: 4     // D
    },
    PROMO_TRIAL: {
      ARR: 6,            // F
      SUBSCRIPTIONS: 7,  // G
      TOTAL_SEATS: 8     // H
    },
    FREE_TRIAL: {
      ARR: 10,           // J
      SUBSCRIPTIONS: 11, // K
      TOTAL_SEATS: 12    // L
    }
  },

  // Table headers (Row 3, starting col B)
  HEADERS: [
    'Customer Email',
    'Customer Name',
    'Org Name',
    'Status',
    'Sign Up Date',
    'First Payment At',     // ✅ NEW
    'Trial Days Remaining',
    'Interval',
    'ARR',
    'Seats'
  ],

  // Formatting
  CURRENCY_FMT: '$#,##0.00',
  INT_FMT: '0',
  PERCENT_FMT: '0.##%',
  TEXT_FMT: '@',
  DATETIME_FMT: 'yyyy-mm-dd hh:mm:ss',
  DATE_FMT: 'yyyy-mm-dd'
}

const RING_AUTO_PUBLISH_GOOD_STUFF = true

function render_ring_view() {
  lockWrapCompat_('render_ring_view', () => {
    const t0 = new Date()
    try {
      const ss = SpreadsheetApp.getActive()
      const sh = getOrCreateSheetCompat_(ss, RING_CFG.SHEET_NAME)

      const src = ss.getSheetByName(RING_CFG.INPUT_SHEET)
      if (!src) throw new Error(`Missing input sheet: ${RING_CFG.INPUT_SHEET}`)
      const manualChangesSrc = ss.getSheetByName(RING_CFG.MANUAL_CHANGES_SHEET)

      // Load Clerk sources for enrichment
      const canonOrgsSh = ss.getSheetByName(RING_CFG.CANON_ORGS_SHEET)
      if (!canonOrgsSh) throw new Error(`Missing input sheet: ${RING_CFG.CANON_ORGS_SHEET}`)
      const clerkUsersSh = ss.getSheetByName(RING_CFG.CLERK_USERS_SHEET)
      const clerkMemsSh  = ss.getSheetByName(RING_CFG.CLERK_MEMBERSHIPS_SHEET)
      const clerkOrgsSh  = ss.getSheetByName(RING_CFG.CLERK_ORGS_SHEET)
      const posthogUsersSh = ss.getSheetByName(RING_CFG.POSTHOG_USERS_SHEET)
      const posthogOrgSubsSh = ss.getSheetByName(RING_CFG.POSTHOG_ORG_SUBS_SHEET)
      const posthogPromoRedemptionsSh = ss.getSheetByName(RING_CFG.POSTHOG_PROMO_REDEMPTIONS_SHEET)

      const canonOrgs = readSheetObjects_(canonOrgsSh, 1)
      const clerkUsers = clerkUsersSh ? readSheetObjects_(clerkUsersSh, 1) : []
      const clerkMems  = clerkMemsSh  ? readSheetObjects_(clerkMemsSh, 1)  : []
      const clerkOrgs  = clerkOrgsSh  ? readSheetObjects_(clerkOrgsSh, 1)  : []
      const posthogUsers = posthogUsersSh ? readSheetObjects_(posthogUsersSh, 1) : []
      const posthogOrgSubs = posthogOrgSubsSh ? readSheetObjects_(posthogOrgSubsSh, 1) : []
      const posthogPromoRedemptions = posthogPromoRedemptionsSh ? readSheetObjects_(posthogPromoRedemptionsSh, 1) : []

      const ringCanon = buildRingCanonIndex_(canonOrgs)
      const ringIndexes = buildRingIndexes_(clerkUsers, clerkMems, clerkOrgs, posthogUsers, ringCanon)
      const orgIdByStripeSubId = buildRingOrgIdByStripeSubId_(posthogOrgSubs)
      const posthogSubByStripeSubId = buildRingPosthogSubByStripeSubId_(posthogOrgSubs)
      const appPromoCodesByOrgId = buildRingAppPromoCodesByOrgId_(posthogPromoRedemptions)

      // Stripe subscriptions
      const rows = readSheetObjects_(src, 1)
      const manualChangesBySubId = buildManualStripeChangesBySubId_(manualChangesSrc)

      const out = []
      let paidArr = 0
      let paidSeats = 0
      let paidSubs = 0
      let intentToPayArr = 0
      let intentToPaySeats = 0
      let intentToPaySubs = 0
      let freeTrialArr = 0
      let freeTrialSeats = 0
      let freeTrialSubs = 0

      for (const r of rows) {
        const stripeSubscriptionId =
          str_(r.stripe_subscription_id) ||
          str_(r.subscription_id) ||
          str_(r.subscription) ||
          str_(r.id) ||
          ''

        const orgSubInfo = stripeSubscriptionId ? (posthogSubByStripeSubId.get(stripeSubscriptionId) || null) : null
        if (!orgSubInfo) continue

        const manualChange = stripeSubscriptionId ? (manualChangesBySubId.get(stripeSubscriptionId) || null) : null
        if (manualChange && manualChange.excludeInternal) continue

        const statusRaw = str_(orgSubInfo.status).toLowerCase()
        const firstPaymentAtIso = str_(orgSubInfo.first_payment_at)
        const firstPaymentAtDate = isoToDateOrBlank_(firstPaymentAtIso)
        const hasPaymentMethod = toBool_(orgSubInfo.has_payment_method)

        // ── Single source of truth: discount-adjusted ARR from Stripe sync ──
        const interval = str_(orgSubInfo.interval).toLowerCase() || str_(r.interval).toLowerCase()
        const intervalCount = Math.max(1, num_(orgSubInfo.interval_count) || num_(r.interval_count) || 1)

        // Discount-adjusted ARR (computed at item level in Stripe sync)
        const discountedYearly = num_(orgSubInfo.amount_after_discounts_yearly)
        const discountedAmount = num_(orgSubInfo.amount_after_discounts)
        const arr = discountedYearly > 0
          ? discountedYearly
          : discountedAmount > 0
            ? computeMrrArr_(discountedAmount, interval, intervalCount).arr
            : 0

        // Gross (pre-discount) ARR for 100%-off reclassification
        const grossYearly = num_(orgSubInfo.amount_yearly)
        const grossAmount = moneyAmount_(orgSubInfo.amount)
        const grossArr = grossYearly > 0
          ? grossYearly
          : computeMrrArr_(grossAmount, interval, intervalCount).arr

        // ── Bucket logic with 100%-off override ──
        let ringBucket = ringBucketFromOrgSubscriptionInfo_(statusRaw, firstPaymentAtIso, hasPaymentMethod)
        if (!ringBucket) continue

        // Override: if sub would be "paid" but net ARR is $0 (100% discount),
        // reclassify based on payment method presence
        let bucketArr = arr
        if (ringBucket === 'paid' && arr < 0.01) {
          ringBucket = hasPaymentMethod ? 'intent_to_pay' : 'free_trial'
          bucketArr = grossArr  // Intent/Free Trial shows full pre-discount ARR
        }

        // Seat count from org_subscription_info when available.
        const seats = Math.max(0, safeInt_(orgSubInfo.quantity_total) || safeInt_(r.quantity_total))

        // Stripe identifiers for enrichment
        const stripeEmailRaw = str_(r.customer_email || r.email || r.billing_email)
        const stripeEmailKey = normalizeEmailCompat_(stripeEmailRaw)

        const resolved = resolveRingCustomer_(stripeEmailKey, stripeSubscriptionId, ringIndexes)
        const canonMatch = (stripeSubscriptionId && ringIndexes.canonBySubId)
          ? (ringIndexes.canonBySubId.get(stripeSubscriptionId) || null)
          : null
        const resolvedOrgId =
          (canonMatch && (canonMatch.appOrgId || canonMatch.orgId || canonMatch.clerkOrgId)) ||
          (stripeSubscriptionId ? (orgIdByStripeSubId.get(stripeSubscriptionId) || '') : '') ||
          (resolved.appOrgId || resolved.orgId) ||
          str_(r.org_id || r.organization_id)

        const email = resolved.email || stripeEmailRaw || (canonMatch ? canonMatch.billingEmail : '')
        const customerName = resolved.customerName || str_(r.customer_name || r.name)
        const orgName =
          (canonMatch && canonMatch.orgName) ||
          (resolvedOrgId ? (ringIndexes.orgNameByOrgId.get(resolvedOrgId) || '') : '') ||
          resolved.orgName ||
          str_(r.org_name || r.organization_name || r.org)

        const appPromoCodes = resolvedOrgId ? (appPromoCodesByOrgId.get(resolvedOrgId) || []) : []

        if (ringBucket === 'paid') {
          paidSubs += 1
          paidArr += bucketArr
          paidSeats += seats
        } else if (ringBucket === 'intent_to_pay') {
          intentToPaySubs += 1
          intentToPayArr += bucketArr
          intentToPaySeats += seats
        } else {
          freeTrialSubs += 1
          freeTrialArr += bucketArr
          freeTrialSeats += seats
        }

        // Keep free-trial orgs out of the Ring detail table.
        if (ringBucket === 'free_trial') continue
        const displayStatus = ringBucket === 'paid' ? 'Paid' : 'Intent to Pay'
        const asOfNow = new Date()

        const canonByResolvedOrg = resolvedOrgId ? (ringIndexes.canonByOrgId.get(resolvedOrgId) || null) : null
        const signUpIso =
          str_((canonMatch && canonMatch.orgCreatedAt) || '') ||
          str_((canonByResolvedOrg && canonByResolvedOrg.orgCreatedAt) || '') ||
          str_(r.created_at)
        const signUpDate = isoToDateOrBlank_(signUpIso)
        const signUpMs = signUpDate instanceof Date ? signUpDate.getTime() : 0

        const trialEndIso =
          str_(orgSubInfo && (orgSubInfo.trial_ends_at || orgSubInfo.current_period_end)) ||
          str_(r.current_period_end) ||
          ''
        const trialDaysRemaining = ringComputeTrialDaysRemaining_(trialEndIso, asOfNow)

        out.push({
          sortGroup: ringBucket === 'intent_to_pay' ? 0 : 1,
          signUpMs,
          row: [
            email,
            customerName,
            orgName,
            displayStatus,
            signUpDate,
            firstPaymentAtDate,
            ringBucket === 'intent_to_pay' ? trialDaysRemaining : '',
            interval || '',
            bucketArr,
            seats
          ]
        })
      }

      out.sort((a, b) => {
        if (a.sortGroup !== b.sortGroup) return a.sortGroup - b.sortGroup
        return (b.signUpMs || 0) - (a.signUpMs || 0)
      })
      const outRows = out.map(x => x.row)

      // Clean rebuild
      sh.clear()

      // KPIs
      writeKpis_(sh, {
        paid: { arr: paidArr, subscriptions: paidSubs, totalSeats: paidSeats },
        promoTrial: { arr: intentToPayArr, subscriptions: intentToPaySubs, totalSeats: intentToPaySeats },
        freeTrial: { arr: freeTrialArr, subscriptions: freeTrialSubs, totalSeats: freeTrialSeats }
      })

      // Headers
      sh.getRange(RING_CFG.HEADER_ROW, RING_CFG.START_COL, 1, RING_CFG.HEADERS.length).setValues([RING_CFG.HEADERS])
      sh.setFrozenRows(RING_CFG.HEADER_ROW)

      // Data
      if (outRows.length) {
        batchSetValuesCompat_(sh, RING_CFG.DATA_START_ROW, RING_CFG.START_COL, outRows, 3000)
      }

      // Formatting
      applyRingFormats_(sh, outRows.length)

      // Resize
      sh.autoResizeColumns(RING_CFG.START_COL, RING_CFG.HEADERS.length)

      writeSyncLogCompat_(
        'render_ring_view',
        'ok',
        rows.length,
        outRows.length,
        (new Date() - t0) / 1000,
        ''
      )

      // Best effort: publish external dashboard after a successful Ring render.
      // Do not fail Ring if publish encounters an external permission/network issue.
      if (RING_AUTO_PUBLISH_GOOD_STUFF && typeof publish_the_good_stuff === 'function') {
        try {
          publish_the_good_stuff()
        } catch (pubErr) {
          writeSyncLogCompat_(
            'publish_the_good_stuff (auto)',
            'error',
            '',
            '',
            '',
            String(pubErr && pubErr.message ? pubErr.message : pubErr)
          )
        }
      }

      return { rows_in: rows.length, rows_out: outRows.length }
    } catch (err) {
      writeSyncLogCompat_(
        'render_ring_view',
        'error',
        '',
        '',
        '',
        String(err && err.message ? err.message : err)
      )
      throw err
    }
  })
}

/* =========================
 * Clerk enrichment indexes
 * ========================= */

function buildRingCanonIndex_(canonRows) {
  const bySubId = new Map()
  const byOrgId = new Map()
  const byClerkOrgId = new Map()

  for (const r of (canonRows || [])) {
    const appOrgId = str_(r.app_org_id)
    const clerkOrgId = str_(r.clerk_org_id || r.org_id)
    const orgId = appOrgId || clerkOrgId
    if (!orgId && !clerkOrgId) continue

    const orgName = str_(r.org_name || r.org_slug)
    const billingEmail = str_(r.billing_email)
    const orgCreatedAt = str_(r.org_created_at || r.created_at)
    if (orgId) byOrgId.set(orgId, { orgId, appOrgId, clerkOrgId, orgName, billingEmail, orgCreatedAt })
    if (clerkOrgId) byClerkOrgId.set(clerkOrgId, { orgId, appOrgId, clerkOrgId, orgName, billingEmail, orgCreatedAt })

    const subIds = ringCsvList_(r.stripe_subscription_ids)
    subIds.forEach(subId => {
      const s = str_(subId)
      if (!s || bySubId.has(s)) return
      bySubId.set(s, { orgId, appOrgId, clerkOrgId, orgName, billingEmail, orgCreatedAt })
    })
  }

  return { bySubId, byOrgId, byClerkOrgId }
}

function buildRingIndexes_(clerkUsers, clerkMems, clerkOrgs, posthogUsers, ringCanon) {
  // org_id -> org_name
  const orgNameByOrgId = new Map()
  for (const o of (clerkOrgs || [])) {
    const orgId = str_(o.org_id)
    if (!orgId) continue
    const name = str_(o.org_name) || str_(o.org_slug)
    if (name) orgNameByOrgId.set(orgId, name)
  }

  if (ringCanon && ringCanon.byOrgId instanceof Map) {
    ringCanon.byOrgId.forEach((v, orgId) => {
      const name = str_(v && v.orgName)
      if (orgId && name) orgNameByOrgId.set(orgId, name)
    })
  }
  if (ringCanon && ringCanon.byClerkOrgId instanceof Map) {
    ringCanon.byClerkOrgId.forEach((v, clerkOrgId) => {
      const appOrgId = str_(v && v.appOrgId)
      const name = str_(v && v.orgName)
      if (clerkOrgId && name) orgNameByOrgId.set(clerkOrgId, name)
      if (appOrgId && name) orgNameByOrgId.set(appOrgId, name)
    })
  }

  // email_key -> memberships [{orgId, role, isOwnerish}]
  const membershipsByEmailKey = new Map()
  for (const m of (clerkMems || [])) {
    const emailKey =
      str_(m.email_key) ||
      normalizeEmailCompat_(str_(m.email))

    const orgId = str_(m.org_id)
    if (!emailKey || !orgId) continue

    const role = str_(m.role).toLowerCase()
    const isOwnerish =
      role.includes('owner') ||
      role.includes('admin') ||
      role === 'org:admin' ||
      role === 'admin' ||
      role === 'owner'

    if (!membershipsByEmailKey.has(emailKey)) membershipsByEmailKey.set(emailKey, [])
    membershipsByEmailKey.get(emailKey).push({ orgId, role, isOwnerish })
  }

  // stripe_subscription_id -> list of users
  const usersByStripeSubId = new Map()
  const usersByEmailKey = new Map()
  for (const u of (clerkUsers || [])) {
    const subId = str_(u.stripe_subscription_id || u.stripeSubscriptionId)

    const email = str_(u.email)
    const emailKey = str_(u.email_key) || normalizeEmailCompat_(email)
    if (!emailKey) continue

    const name = str_(u.name)
    const orgId = str_(u.org_id)
    const userObj = { email, emailKey, name, orgId }

    if (!usersByEmailKey.has(emailKey)) usersByEmailKey.set(emailKey, userObj)

    if (!subId) continue

    if (!usersByStripeSubId.has(subId)) usersByStripeSubId.set(subId, [])
    usersByStripeSubId.get(subId).push(userObj)
  }

  // Fallback map from PostHog: stripe_subscription_id -> Set(email_key)
  const posthogEmailKeysByStripeSubId = new Map()
  for (const p of (posthogUsers || [])) {
    const subId = str_(p.stripe_subscription_id || p.subscription_id || p.subscription)
    if (!subId) continue

    const emailKey = str_(p.email_key) || normalizeEmailCompat_(str_(p.email))
    if (!emailKey) continue

    if (!posthogEmailKeysByStripeSubId.has(subId)) posthogEmailKeysByStripeSubId.set(subId, new Set())
    posthogEmailKeysByStripeSubId.get(subId).add(emailKey)
  }

  return {
    canonBySubId: (ringCanon && ringCanon.bySubId) || new Map(),
    canonByOrgId: (ringCanon && ringCanon.byOrgId) || new Map(),
    orgNameByOrgId,
    membershipsByEmailKey,
    usersByStripeSubId,
    usersByEmailKey,
    posthogEmailKeysByStripeSubId
  }
}

function ringCsvList_(v) {
  const s = str_(v)
  if (!s) return []
  return s.split(',').map(x => str_(x))
}

function buildRingOrgIdByStripeSubId_(rows) {
  const out = new Map()
  for (const r of (rows || [])) {
    const subId =
      str_(r.stripe_subscription_id) ||
      str_(r.subscription_id) ||
      str_(r.subscription) ||
      str_(r.id)
    const orgId = str_(r.app_org_id || r.org_id)
    if (!subId || !orgId) continue
    if (!out.has(subId)) out.set(subId, orgId)
  }
  return out
}

function buildRingPosthogSubByStripeSubId_(rows) {
  const out = new Map()
  for (const r of (rows || [])) {
    const subId =
      str_(r.stripe_subscription_id) ||
      str_(r.subscription_id) ||
      str_(r.subscription) ||
      str_(r.id)
    if (!subId) continue
    if (!out.has(subId)) out.set(subId, r)
  }
  return out
}

function buildRingAppPromoCodesByOrgId_(rows) {
  const setsByOrgId = new Map()
  for (const r of (rows || [])) {
    const location = str_(r.redemption_location).toLowerCase()
    if (location && location !== 'in_app') continue
    const orgId = str_(r.app_org_id || r.org_id)
    if (!orgId) continue
    const code =
      str_(r.promo_code) ||
      str_(r.code) ||
      str_(r.promo_name) ||
      str_(r.name)
    if (!code) continue

    if (!setsByOrgId.has(orgId)) setsByOrgId.set(orgId, new Set())
    setsByOrgId.get(orgId).add(code)
  }

  const out = new Map()
  setsByOrgId.forEach((set, orgId) => out.set(orgId, Array.from(set)))
  return out
}

function resolveRingCustomer_(stripeEmailKey, stripeSubscriptionId, idx) {
  const subId = str_(stripeSubscriptionId)
  let candidates = (subId && idx.usersByStripeSubId.has(subId))
    ? idx.usersByStripeSubId.get(subId).slice()
    : []

  // Fallback: if Clerk users are not directly keyed by sub id, use PostHog mapping
  // (sub id -> email_key) then map those email keys back to Clerk users.
  if (!candidates.length && subId && idx.posthogEmailKeysByStripeSubId && idx.posthogEmailKeysByStripeSubId.has(subId)) {
    const emailKeys = Array.from(idx.posthogEmailKeysByStripeSubId.get(subId) || [])
    const fromPosthog = []
    emailKeys.forEach(emailKey => {
      const hit = idx.usersByEmailKey && idx.usersByEmailKey.get(emailKey)
      if (hit) fromPosthog.push(hit)
    })
    candidates = fromPosthog
  }

  // Final fallback: Stripe customer email -> Clerk user email
  if (!candidates.length && stripeEmailKey && idx.usersByEmailKey && idx.usersByEmailKey.has(stripeEmailKey)) {
    candidates = [idx.usersByEmailKey.get(stripeEmailKey)]
  }

  if (!candidates.length) return { email: '', customerName: '', orgName: '', orgId: '', appOrgId: '' }

  // 1) Prefer exact Stripe email match if present
  let filtered = candidates
  if (stripeEmailKey) {
    const exact = candidates.filter(c => c.emailKey === stripeEmailKey)
    if (exact.length) filtered = exact
  }

  // 2) Prefer owner/admin (based on memberships)
  const scored = filtered.map(c => {
    const mems = idx.membershipsByEmailKey.get(c.emailKey) || []
    const hasOwnerish = mems.some(m => m.isOwnerish)
    return { ...c, _hasOwnerish: hasOwnerish }
  })

  scored.sort((a, b) => {
    if (a._hasOwnerish !== b._hasOwnerish) return a._hasOwnerish ? -1 : 1
    return String(a.email || '').localeCompare(String(b.email || ''))
  })

  const picked = scored[0]

  // Resolve org for this picked user from memberships, preferring owner/admin
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
    appOrgId = str_(canon && canon.appOrgId)
  }
  const orgName = (appOrgId && idx.orgNameByOrgId.get(appOrgId)) || (orgId ? (idx.orgNameByOrgId.get(orgId) || '') : '')

  return {
    email: picked.email || '',
    customerName: picked.name || '',
    orgName,
    orgId,
    appOrgId
  }
}

/* =========================
 * KPI + Formatting helpers
 * ========================= */

function writeKpis_(sheet, { paid, promoTrial, freeTrial }) {
  const paidCols = RING_CFG.KPI_COLS.PAID
  const promoCols = RING_CFG.KPI_COLS.PROMO_TRIAL
  const freeCols = RING_CFG.KPI_COLS.FREE_TRIAL

  sheet.getRange(RING_CFG.KPI_ROW_LABEL, paidCols.ARR).setValue('ARR')
  sheet.getRange(RING_CFG.KPI_ROW_LABEL, paidCols.SUBSCRIPTIONS).setValue('Subscriptions')
  sheet.getRange(RING_CFG.KPI_ROW_LABEL, paidCols.TOTAL_SEATS).setValue('Total Seats')
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, paidCols.ARR).setValue((paid && paid.arr) || 0)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, paidCols.SUBSCRIPTIONS).setValue((paid && paid.subscriptions) || 0)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, paidCols.TOTAL_SEATS).setValue((paid && paid.totalSeats) || 0)

  sheet.getRange(RING_CFG.KPI_ROW_LABEL, promoCols.ARR).setValue('Intent to Pay')
  sheet.getRange(RING_CFG.KPI_ROW_LABEL, promoCols.SUBSCRIPTIONS).setValue('Subscriptions')
  sheet.getRange(RING_CFG.KPI_ROW_LABEL, promoCols.TOTAL_SEATS).setValue('Total Seats')
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, promoCols.ARR).setValue((promoTrial && promoTrial.arr) || 0)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, promoCols.SUBSCRIPTIONS).setValue((promoTrial && promoTrial.subscriptions) || 0)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, promoCols.TOTAL_SEATS).setValue((promoTrial && promoTrial.totalSeats) || 0)

  sheet.getRange(RING_CFG.KPI_ROW_LABEL, freeCols.ARR).setValue('Trialing')
  sheet.getRange(RING_CFG.KPI_ROW_LABEL, freeCols.SUBSCRIPTIONS).setValue('Subscriptions')
  sheet.getRange(RING_CFG.KPI_ROW_LABEL, freeCols.TOTAL_SEATS).setValue('Total Seats')
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, freeCols.ARR).setValue((freeTrial && freeTrial.arr) || 0)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, freeCols.SUBSCRIPTIONS).setValue((freeTrial && freeTrial.subscriptions) || 0)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, freeCols.TOTAL_SEATS).setValue((freeTrial && freeTrial.totalSeats) || 0)

  formatKpiGroup_(sheet, paidCols)
  formatKpiGroup_(sheet, promoCols)
  formatKpiGroup_(sheet, freeCols)

  sheet.getRange(1, 1, 2, Math.max(sheet.getLastColumn(), 16)).setVerticalAlignment('middle')
}

function formatKpiGroup_(sheet, cols) {
  const startCol = Math.min(cols.ARR, cols.SUBSCRIPTIONS, cols.TOTAL_SEATS)
  const labelRange = sheet.getRange(RING_CFG.KPI_ROW_LABEL, startCol, 1, 3)
  labelRange.setFontWeight('bold').setHorizontalAlignment('center')

  const valueRange = sheet.getRange(RING_CFG.KPI_ROW_VALUE, startCol, 1, 3)
  valueRange.setFontWeight('bold').setFontSize(22).setHorizontalAlignment('center')

  sheet.getRange(RING_CFG.KPI_ROW_VALUE, cols.ARR).setNumberFormat(RING_CFG.CURRENCY_FMT)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, cols.SUBSCRIPTIONS).setNumberFormat(RING_CFG.INT_FMT)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, cols.TOTAL_SEATS).setNumberFormat(RING_CFG.INT_FMT)
}

function applyRingFormats_(sheet, numDataRows) {
  const headerRange = sheet.getRange(RING_CFG.HEADER_ROW, RING_CFG.START_COL, 1, RING_CFG.HEADERS.length)
  headerRange.setFontWeight('bold').setBackground('#f3f3f3')

  if (!numDataRows) return

  const startRow = RING_CFG.DATA_START_ROW
  const startCol = RING_CFG.START_COL
  const nRows = numDataRows

  const colSignUp = colByHeader_(startCol, 'Sign Up Date')
  const colFirstPay = colByHeader_(startCol, 'First Payment At')
  const colTrialDays = colByHeader_(startCol, 'Trial Days Remaining')
  const colArr = colByHeader_(startCol, 'ARR')
  const colSeats = colByHeader_(startCol, 'Seats')

  const full = sheet.getRange(startRow, startCol, nRows, RING_CFG.HEADERS.length)
  full.setNumberFormat('@')
  full.setVerticalAlignment('middle')

  sheet.getRange(startRow, colSignUp, nRows, 1).setNumberFormat(RING_CFG.DATE_FMT)
  sheet.getRange(startRow, colFirstPay, nRows, 1).setNumberFormat(RING_CFG.DATETIME_FMT)
  sheet.getRange(startRow, colTrialDays, nRows, 1).setNumberFormat(RING_CFG.INT_FMT)
  sheet.getRange(startRow, colArr, nRows, 1).setNumberFormat(RING_CFG.CURRENCY_FMT)
  sheet.getRange(startRow, colSeats, nRows, 1).setNumberFormat(RING_CFG.INT_FMT)
}

function colByHeader_(startCol, headerName) {
  const i = RING_CFG.HEADERS.indexOf(headerName)
  if (i < 0) throw new Error(`RING_CFG.HEADERS missing: ${headerName}`)
  return startCol + i
}

/* =========================
 * Business logic helpers
 * ========================= */

function computeMrrArr_(amount, interval, intervalCount) {
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

function moneyAmount_(raw) {
  if (raw === null || raw === undefined || raw === '') return 0
  const n = num_(raw)
  if (!isFinite(n)) return 0
  return Math.round(n * 100) / 100
}

function ringBuildDiscountContextNow_(row, opts) {
  const cfg = opts || {}
  const amountRaw = moneyAmount_(cfg.amountRaw)
  const asOfDate = (cfg.asOfDate instanceof Date && !isNaN(cfg.asOfDate.getTime()))
    ? cfg.asOfDate
    : new Date()

  const details = ringParseDiscountDetails_(row)
  const active = details.filter(d => ringIsDiscountActiveNow_(d, asOfDate))

  let amount = amountRaw
  const promoCodes = []
  const durationLabels = []
  for (const d of active) {
    const pct = num_(d.percent_off)
    if (pct > 0) {
      const bounded = Math.max(0, Math.min(100, pct))
      amount *= (1 - bounded / 100)
    }

    const amountOff = num_(d.amount_off)
    if (amountOff > 0) amount -= amountOff

    if (str_(d.promotion_code)) promoCodes.push(str_(d.promotion_code))
    const label = ringFormatDiscountDuration_(d.duration, d.duration_in_months)
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

function ringBuildAmountIgnoringPercentDiscountsNow_(row, opts) {
  const cfg = opts || {}
  const amountRaw = moneyAmount_(cfg.amountRaw)
  const asOfDate = (cfg.asOfDate instanceof Date && !isNaN(cfg.asOfDate.getTime()))
    ? cfg.asOfDate
    : new Date()

  const details = ringParseDiscountDetails_(row)
  const active = details.filter(d => ringIsDiscountActiveNow_(d, asOfDate))

  let amount = amountRaw
  for (const d of active) {
    const amountOff = num_(d.amount_off)
    if (amountOff > 0) amount -= amountOff
  }

  return Math.max(0, amount)
}

function ringComputeTrialDaysRemaining_(trialEndIso, asOfDate) {
  const end = isoToDateOrBlank_(trialEndIso)
  if (!(end instanceof Date) || isNaN(end.getTime())) return ''
  const now = (asOfDate instanceof Date && !isNaN(asOfDate.getTime())) ? asOfDate : new Date()
  const msPerDay = 24 * 60 * 60 * 1000
  const days = Math.ceil((end.getTime() - now.getTime()) / msPerDay)
  return days > 0 ? days : 0
}

function ringBucketFromOrgSubscriptionInfo_(statusRaw, firstPaymentAt, hasPaymentMethodRaw) {
  const status = str_(statusRaw).toLowerCase()
  const hasFirstPayment = !!str_(firstPaymentAt)
  const hasPaymentMethod = !!toBool_(hasPaymentMethodRaw)

  // Group 1: Paid
  if (status === 'active' && hasFirstPayment) return 'paid'

  // Group 2: Intent to Pay (must have payment method)
  if (status === 'active' && !hasFirstPayment && hasPaymentMethod) return 'intent_to_pay'
  if (status === 'trialing' && hasPaymentMethod) return 'intent_to_pay'

  // Group 3: Free Trial (no payment method)
  if (status === 'active' && !hasFirstPayment && !hasPaymentMethod) return 'free_trial'
  if (status === 'trialing' && !hasPaymentMethod) return 'free_trial'

  return ''
}

function ringParseDiscountDetails_(row) {
  const out = []
  const r = row || {}

  const detailsJson = str_(r.discount_details_json)
  if (detailsJson) {
    try {
      const parsed = JSON.parse(detailsJson)
      if (Array.isArray(parsed)) {
        parsed.forEach((d, i) => {
          if (!d || typeof d !== 'object') return
          out.push({
            index: i + 1,
            percent_off: num_(d.percent_off),
            amount_off: num_(d.amount_off),
            duration: str_(d.duration).toLowerCase(),
            duration_in_months: num_(d.duration_in_months),
            start_at: str_(d.start_at),
            end_at: str_(d.end_at),
            promotion_code: str_(d.promotion_code)
          })
        })
      }
    } catch (e) {}
  }
  if (out.length) return out

  const pctAll = ringCsvList_(r.discount_percent_all)
  const amtAll = ringCsvList_(r.discount_amount_off_all)
  const durAll = ringCsvList_(r.discount_duration_all)
  const durMonthsAll = ringCsvList_(r.discount_duration_months_all)
  const startAll = ringCsvList_(r.discount_start_at_all)
  const endAll = ringCsvList_(r.discount_end_at_all)
  const promoAll = ringCsvList_(r.promo_code_all)
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
      percent_off: num_(pctAll[i]),
      amount_off: num_(amtAll[i]),
      duration: str_(durAll[i]).toLowerCase(),
      duration_in_months: num_(durMonthsAll[i]),
      start_at: str_(startAll[i]),
      end_at: str_(endAll[i]),
      promotion_code: str_(promoAll[i])
    })
  }
  if (out.length) return out

  const pct = num_(r.discount_percent)
  const amt = num_(r.discount_amount_off)
  const duration = str_(r.discount_duration).toLowerCase()
  const durationMonths = num_(r.discount_duration_months)
  if (pct <= 0 && amt <= 0) return []

  return [{
    index: 1,
    percent_off: pct,
    amount_off: amt,
    duration,
    duration_in_months: durationMonths,
    start_at: str_(r.discount_start_at || r.first_payment_at || r.created_at),
    end_at: str_(r.discount_end_at),
    promotion_code: str_(r.promo_code)
  }]
}

function ringIsDiscountActiveNow_(d, asOfDate) {
  const asOf = (asOfDate instanceof Date && !isNaN(asOfDate.getTime())) ? asOfDate : new Date()
  const start = isoToDateOrBlank_(d && d.start_at)
  const end = isoToDateOrBlank_(d && d.end_at)
  const duration = str_(d && d.duration).toLowerCase()
  const hasValue = num_(d && d.percent_off) > 0 || num_(d && d.amount_off) > 0
  if (!hasValue) return false

  if (start && asOf < start) return false
  if (end) return asOf < end
  if (duration === 'forever') return true
  if (duration === 'repeating' || duration === 'once') {
    if (!start) return false
    const months = Math.max(1, num_(d && d.duration_in_months) || 1)
    const until = new Date(start.getTime())
    until.setUTCMonth(until.getUTCMonth() + months)
    return asOf < until
  }
  return false
}

function ringFormatDiscountDuration_(duration, durationMonths) {
  const d = str_(duration).toLowerCase().trim()
  if (d === 'forever') return 'forever'
  if (d === 'once') return 'once'
  if (d === 'repeating') {
    const n = Number(durationMonths)
    if (isFinite(n) && n > 0) return `repeating ${Math.floor(n)} mo`
    return 'repeating'
  }
  return ''
}

function ringCombinePromoCodes_(stripePromoCodes, appPromoCodes) {
  const all = []
  ;(stripePromoCodes || []).forEach(c => {
    const v = str_(c)
    if (v) all.push(v)
  })
  ;(appPromoCodes || []).forEach(c => {
    const v = str_(c)
    if (v) all.push(v)
  })
  return Array.from(new Set(all)).join(', ')
}

function clamp01_(n) {
  const x = Number(n)
  if (!isFinite(x)) return 0
  return Math.max(0, Math.min(1, x))
}

function isoToDateOrBlank_(iso) {
  const s = String(iso || '').trim()
  if (!s) return ''
  const d = new Date(s)
  if (isNaN(d.getTime())) return ''
  return d
}

function buildManualStripeChangesBySubId_(sheet) {
  const out = new Map()
  if (!sheet) return out

  const rows = readSheetObjects_(sheet, 1)
  for (const r of rows) {
    const subId =
      str_(r.subscription_id) ||
      str_(r.stripe_subscription_id) ||
      str_(r.subscription) ||
      ''
    if (!subId) continue

    const excludeReason = str_(r.exclude_reason).toLowerCase()
    if (excludeReason !== 'internal') continue
    out.set(subId, { excludeInternal: true })
  }

  return out
}

/* =========================
 * Sheet reading helpers
 * ========================= */

function readSheetObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1) return []

  const header = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())
  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()

  return data.map(r => {
    const obj = {}
    header.forEach((h, i) => {
      if (!h) return
      obj[key_(h)] = r[i]
    })
    return obj
  })
}

function key_(h) {
  return String(h || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_')
}

/* =========================
 * Tiny helpers
 * ========================= */

function str_(v) {
  if (v === null || v === undefined) return ''
  return String(v).trim()
}

function num_(v) {
  if (v === null || v === undefined || v === '') return 0
  if (typeof v === 'number') return v
  const s = String(v).replace(/[^0-9.\-]/g, '').trim()
  const n = Number(s)
  return isNaN(n) ? 0 : n
}

function safeInt_(v) {
  const n = Number(v)
  if (isNaN(n) || !isFinite(n)) return 0
  return Math.max(0, Math.floor(n))
}

function toBool_(v) {
  if (v === true) return true
  if (typeof v === 'number') return v === 1
  const s = String(v || '').trim().toLowerCase()
  return s === 'true' || s === '1' || s === 'yes' || s === 'y'
}

function normalizeEmailCompat_(email) {
  if (typeof normalizeEmail === 'function') return normalizeEmail(email)
  return String(email || '').trim().toLowerCase()
}

/* =========================
 * Compatibility wrappers
 * ========================= */

function getOrCreateSheetCompat_(ss, name) {
  if (typeof getOrCreateSheet === 'function') {
    try { return getOrCreateSheet(ss, name) } catch (e) {}
    try { return getOrCreateSheet(name) } catch (e) {}
  }
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function batchSetValuesCompat_(sheet, startRow, startCol, values, chunkSize) {
  if (typeof batchSetValues === 'function') return batchSetValues(sheet, startRow, startCol, values, chunkSize)
  const size = chunkSize || 2000
  for (let i = 0; i < values.length; i += size) {
    const chunk = values.slice(i, i + size)
    sheet.getRange(startRow + i, startCol, chunk.length, chunk[0].length).setValues(chunk)
  }
}

function lockWrapCompat_(lockName, fn) {
  if (typeof lockWrap === 'function') {
    try { return lockWrap(lockName, fn) } catch (e) { return lockWrap(fn) }
  }
  const lock = LockService.getScriptLock()
  const ok = lock.tryLock(300000)
  if (!ok) throw new Error(`Could not acquire lock: ${lockName}`)
  try { return fn() } finally { lock.releaseLock() }
}

function writeSyncLogCompat_(step, status, rowsIn, rowsOut, seconds, error) {
  if (typeof writeSyncLog === 'function') return writeSyncLog(step, status, rowsIn, rowsOut, seconds, error || '')
  Logger.log(`[SYNCLOG missing] ${step} ${status} rows_in=${rowsIn} rows_out=${rowsOut} seconds=${seconds} error=${error || ''}`)
}
