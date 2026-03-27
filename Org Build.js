/**************************************************************
 * Canonical Orgs Builder (upsert + rules) — UPDATED
 *
 * Builds/Overwrites:
 *  - canon_orgs
 *
 * Inputs:
 *  - raw_clerk_orgs
 *  - raw_clerk_memberships
 *  - raw_clerk_users              (used for Stripe linkage from Clerk metadata)
 *  - raw_stripe_subscriptions
 *  - raw_posthog_org_subscriptions (optional: org-level subscription mapping/seat/status context)
 *  - promo_redemptions (optional: app/stripe promo redemption context)
 * Optional (fallback only):
 *  - org_billing_map              (manual mapping if needed for edge cases)
 *
 * Stripe → Org mapping precedence:
 *  - Stripe metadata org_id (from raw_stripe_subscriptions.metadata_json)
 *  - PostHog org_subscriptions(org_id -> stripe_subscription_id)
 *  - Clerk members -> Clerk users -> stripe_subscription_id / stripe_customer_id
 *  - Optional org_billing_map
 *  - Controlled fallback: member email_key -> Stripe customer_email (only when no strong link)
 *
 * Behavior:
 *  - Overwrites computed fields each run (name, billing, Stripe rollups, PostHog rollups, promo rollups)
 *  - Preserves manual fields (service, white_glove, in_onboarding, onboarding_note)
 *
 * Requires shared utils:
 *  - getOrCreateSheet, readHeaderMap, normalizeEmail, batchSetValues,
 *    writeSyncLog, lockWrap
 **************************************************************/

const CANON_ORGS_CFG = {
  SHEETS: {
    CLERK_ORGS: 'raw_clerk_orgs',
    CLERK_MEMBERSHIPS: 'raw_clerk_memberships',
    CLERK_USERS: 'raw_clerk_users',
    STRIPE_SUBS: 'raw_stripe_subscriptions',
    POSTHOG_ORGS: 'raw_posthog_orgs',
    POSTHOG_ORG_SUBS: 'raw_posthog_org_subscriptions',
    POSTHOG_PROMO_REDEMPTIONS: 'promo_redemptions',
    ORG_BILLING_MAP: 'org_billing_map', // optional fallback
    ORG_SUBSCRIPTION_INFO: 'org_subscription_info',
    CANON_ORGS: 'canon_orgs'
  },

  CANON_HEADERS: [
    'org_id',
    'app_org_id',
    'org_name',
    'org_slug',
    'org_created_at',
    'posthog_org_name',
    'posthog_org_created_at',

    'owner_user_id',
    'owner_email',
    'owner_name',

    'is_paying',
    'seats',
    'promo_code',

    // derived convenience/debugging
    'billing_email',
    'billing_customer_id',

    // Stripe rollups
    'stripe_subscription_ids',
    'stripe_subscription_count',
    'stripe_active_subscription_count',
    'stripe_trialing_subscription_count',
    'stripe_canceled_subscription_count',
    'stripe_customer_ids',
    'stripe_customer_count',
    'stripe_statuses',
    'stripe_seats_paying_sum',
    'stripe_seats_max',
    'stripe_has_discount_active',
    'stripe_discount_active_count',
    'stripe_discount_subscription_count',
    'stripe_discount_percent_all',
    'stripe_discount_amount_off_all',
    'stripe_discount_duration_all',
    'stripe_promo_codes',
    'stripe_promo_code_count',
    'stripe_arr_discounted',
    'stripe_arr_full_price',
    'stripe_mrr_discounted',
    'stripe_mrr_full_price',

    // Active subscription snapshot (single selected active/paying subscription)
    'active_subscription_id',
    'active_subscription_status',
    'active_subscription_interval',
    'active_subscription_interval_count',
    'active_subscription_amount',
    'active_subscription_seats',
    'active_subscription_arr_discounted',
    'active_subscription_arr_full_price',
    'active_subscription_promo_codes',

    // CRM-derived status
    'org_status',
    'trial_ends_at',
    'health_score',

    // App promo redemption rollups (PostHog promo_redemptions)
    'app_promo_codes',
    'app_promo_code_count',
    'app_promo_redemption_count',
    'app_promo_trial_days_total',
    'app_promo_latest_redeemed_at',
    'combined_promo_codes',

    // Mapping/debug columns to expose cross-source joins
    'mapping_member_user_count',
    'mapping_member_with_stripe_sub_count',
    'mapping_member_with_customer_id_count',
    'mapping_linked_via',
    'mapping_notes',

    'updated_at'
  ],

  MANUAL_FIELDS: new Set([]),

  PAYING_STATUSES: new Set(['active', 'trialing'])
}

/**
 * Build canon_orgs from raw sources
 */
function build_canon_orgs() {
  lockWrap('build_canon_orgs', () => {
    const t0 = new Date()
    try {
      const ss = SpreadsheetApp.getActive()

      const shOrgs = ss.getSheetByName(CANON_ORGS_CFG.SHEETS.CLERK_ORGS)
      const shMems = ss.getSheetByName(CANON_ORGS_CFG.SHEETS.CLERK_MEMBERSHIPS)
      const shUsers = ss.getSheetByName(CANON_ORGS_CFG.SHEETS.CLERK_USERS)
      const shStripe = ss.getSheetByName(CANON_ORGS_CFG.SHEETS.STRIPE_SUBS)
      const shPosthogOrgs = ss.getSheetByName(CANON_ORGS_CFG.SHEETS.POSTHOG_ORGS)
      const shPosthogOrgSubs = ss.getSheetByName(CANON_ORGS_CFG.SHEETS.POSTHOG_ORG_SUBS)
      const shPosthogPromoRedemptions = ss.getSheetByName(CANON_ORGS_CFG.SHEETS.POSTHOG_PROMO_REDEMPTIONS)

      if (!shOrgs) throw new Error(`Missing input sheet: ${CANON_ORGS_CFG.SHEETS.CLERK_ORGS}`)
      if (!shMems) throw new Error(`Missing input sheet: ${CANON_ORGS_CFG.SHEETS.CLERK_MEMBERSHIPS}`)
      if (!shUsers) throw new Error(`Missing input sheet: ${CANON_ORGS_CFG.SHEETS.CLERK_USERS}`)
      if (!shStripe) throw new Error(`Missing input sheet: ${CANON_ORGS_CFG.SHEETS.STRIPE_SUBS}`)

      const orgRows = readRaw_(shOrgs, 1)
      const memRows = readRaw_(shMems, 1)
      const userRows = readRaw_(shUsers, 1)
      const stripeRows = readRaw_(shStripe, 1)
      const posthogOrgsRows = shPosthogOrgs ? readRaw_(shPosthogOrgs, 1) : { rows: [], has: () => false, col: () => -1 }
      const posthogOrgSubsRows = shPosthogOrgSubs ? readRaw_(shPosthogOrgSubs, 1) : { rows: [], has: () => false, col: () => -1 }
      const posthogPromoRows = shPosthogPromoRedemptions ? readRaw_(shPosthogPromoRedemptions, 1) : { rows: [], has: () => false, col: () => -1 }
      const shOrgSubInfo = ss.getSheetByName(CANON_ORGS_CFG.SHEETS.ORG_SUBSCRIPTION_INFO)
      const orgSubInfoRows = shOrgSubInfo ? readRaw_(shOrgSubInfo, 1) : { rows: [], has: () => false, col: () => -1 }
      const asOfNow = new Date()

      // Health scores (keyed by email domain / account)
      const shHealthScores = ss.getSheetByName('raw_posthog_health_scores')
      const healthByDomain = new Map()
      if (shHealthScores) {
        const hsRows = readRaw_(shHealthScores, 1)
        hsRows.rows.forEach(r => {
          const account = orgBuildStr_(orgBuildRawGet_(hsRows, r, ['account'])).toLowerCase()
          const score = Number(orgBuildRawGet_(hsRows, r, ['health_score']) || 0)
          if (account) healthByDomain.set(account, score)
        })
      }

      // Optional fallback maps
      const billingMap = readOrgBillingMap_()

      // ---- 1) Build: org_id -> Set(clerk_user_id) ----
      const orgToMemberUserIds = new Map()
      const orgToMemberEmailKeys = new Map()
      memRows.rows.forEach(r => {
        const orgId = orgBuildStr_(orgBuildRawGet_(memRows, r, ['org_id']))
        if (!orgId) return

        const uid = orgBuildStr_(orgBuildRawGet_(memRows, r, ['clerk_user_id', 'user_id']))
        if (uid) {
          if (!orgToMemberUserIds.has(orgId)) orgToMemberUserIds.set(orgId, new Set())
          orgToMemberUserIds.get(orgId).add(uid)
        }

        const memEmailKeyRaw = orgBuildStr_(orgBuildRawGet_(memRows, r, ['email_key', 'email']))
        const memEmailKey = orgBuildNormalizeEmail_(memEmailKeyRaw)
        if (memEmailKey) {
          if (!orgToMemberEmailKeys.has(orgId)) orgToMemberEmailKeys.set(orgId, new Set())
          orgToMemberEmailKeys.get(orgId).add(memEmailKey)
        }
      })

      // ---- 2) Build: clerk_user_id -> Stripe linkage from Clerk metadata ----
      const clerkUserToStripe = new Map()
      const clerkUserById = new Map()
      userRows.rows.forEach(r => {
        const clerkUserId = orgBuildStr_(orgBuildRawGet_(userRows, r, ['clerk_user_id', 'user_id', 'id']))
        if (!clerkUserId) return

        const email = orgBuildStr_(orgBuildRawGet_(userRows, r, ['email']))
        const name = orgBuildStr_(orgBuildRawGet_(userRows, r, ['name', 'first_name']))
        clerkUserById.set(clerkUserId, { email, name })
        let stripeSubId = orgBuildStr_(orgBuildRawGet_(userRows, r, [
          'stripe_subscription_id',
          'stripesubscriptionid',
          'stripe_subscription'
        ]))
        let stripeCustomerId = orgBuildStr_(orgBuildRawGet_(userRows, r, [
          'stripe_customer_id',
          'stripecustomerid',
          'stripe_customer'
        ]))

        if (!stripeSubId || !stripeCustomerId) {
          const metaStr = orgBuildStr_(
            orgBuildRawGet_(userRows, r, ['private_metadata', 'public_metadata', 'unsafe_metadata', 'metadata'])
          )
          if (metaStr) {
            const meta = tryParseJson_(metaStr)
            if (meta && typeof meta === 'object') {
              stripeSubId = stripeSubId || orgBuildStr_(meta.stripeSubscriptionId || meta.stripe_subscription_id)
              stripeCustomerId = stripeCustomerId || orgBuildStr_(meta.stripeCustomerId || meta.stripe_customer_id)
            }
          }
        }

        clerkUserToStripe.set(clerkUserId, {
          clerk_user_id: clerkUserId,
          email: email || '',
          stripe_subscription_id: stripeSubId,
          stripe_customer_id: stripeCustomerId
        })
      })

      // ---- 3) Build: subscription_id -> Stripe aggregates ----
      const stripeBySubId = new Map()
      const stripeSubIdsByMetadataOrgId = new Map()
      stripeRows.rows.forEach(r => {
        const subId = orgBuildStr_(orgBuildRawGet_(stripeRows, r, [
          'stripe_subscription_id',
          'subscription_id',
          'stripe_subscription',
          'subscription id',
          'id'
        ]))
        if (!subId) return

        const status = orgBuildStr_(orgBuildRawGet_(stripeRows, r, ['status'])).toLowerCase()
        const isPaying = CANON_ORGS_CFG.PAYING_STATUSES.has(status)
        const seats = orgBuildNum_(orgBuildRawGet_(stripeRows, r, ['quantity_total', 'quantity total']))
        const customerId = orgBuildStr_(orgBuildRawGet_(stripeRows, r, [
          'stripe_customer_id',
          'customer_id',
          'customer id'
        ]))
        const customerEmail = orgBuildStr_(orgBuildRawGet_(stripeRows, r, ['customer_email', 'customer email']))
        const customerEmailKey = orgBuildNormalizeEmail_(customerEmail)
        const metadataOrgId = orgBuildExtractOrgIdFromStripeRow_(stripeRows, r)
        const interval = orgBuildStr_(orgBuildRawGet_(stripeRows, r, ['interval'])).toLowerCase()
        const intervalCount = Math.max(1, orgBuildNum_(orgBuildRawGet_(stripeRows, r, ['interval_count'])) || 1)
        const amount = orgBuildNum_(orgBuildRawGet_(stripeRows, r, ['amount']))
        const createdAt = orgBuildStr_(orgBuildRawGet_(stripeRows, r, ['created_at']))
        const currentPeriodEnd = orgBuildStr_(orgBuildRawGet_(stripeRows, r, ['current_period_end']))
        const trialEnd = orgBuildStr_(orgBuildRawGet_(stripeRows, r, ['trial_end']))
        const discountSummary = orgBuildStripeDiscountSummaryFromRaw_(stripeRows, r, asOfNow, amount)
        const arrFullPrice = orgBuildAnnualizeAmount_(amount, interval, intervalCount)
        const arrDiscounted = orgBuildAnnualizeAmount_(discountSummary.discounted_amount, interval, intervalCount)
        const promoCodes = orgBuildStripePromoCodesFromRaw_(stripeRows, r)

        stripeBySubId.set(subId, {
          subscription_id: subId,
          status,
          is_paying: isPaying,
          seats,
          billing_customer_id: customerId,
          billing_email: customerEmail,
          interval,
          interval_count: intervalCount,
          amount,
          created_at: createdAt,
          current_period_end: currentPeriodEnd,
          arr_full_price: arrFullPrice,
          arr_discounted: arrDiscounted,
          mrr_full_price: arrFullPrice / 12,
          mrr_discounted: arrDiscounted / 12,
          promo_codes: promoCodes,
          has_active_discount: discountSummary.active_count > 0,
          active_discount_count: discountSummary.active_count,
          active_discount_percents: discountSummary.active_percent_off,
          active_discount_amount_off: discountSummary.active_amount_off,
          active_discount_durations: discountSummary.active_duration,
          trial_end: trialEnd,
          metadata_org_id: metadataOrgId,
          customer_email_key: customerEmailKey
        })

        if (metadataOrgId) {
          if (!stripeSubIdsByMetadataOrgId.has(metadataOrgId)) stripeSubIdsByMetadataOrgId.set(metadataOrgId, new Set())
          stripeSubIdsByMetadataOrgId.get(metadataOrgId).add(subId)
        }
      })

      const stripeSubIdsByCustomerId = new Map()
      stripeBySubId.forEach(s => {
        const customerId = orgBuildStr_(s.billing_customer_id)
        if (!customerId) return
        if (!stripeSubIdsByCustomerId.has(customerId)) stripeSubIdsByCustomerId.set(customerId, new Set())
        stripeSubIdsByCustomerId.get(customerId).add(s.subscription_id)
      })

      // ---- 4) Build PostHog org_subscriptions aggregates ----
      // Build direct clerk_org_id → app_org_id map from raw_posthog_orgs
      const appOrgIdByClerkOrgId = new Map()
      const posthogOrgById = new Map()
      posthogOrgsRows.rows.forEach(r => {
        const id = orgBuildStr_(orgBuildRawGet_(posthogOrgsRows, r, ['app_org_id', 'org_id']))
        if (!id || posthogOrgById.has(id)) return
        const clerkId = orgBuildStr_(orgBuildRawGet_(posthogOrgsRows, r, ['clerk_org_id']))
        if (clerkId && !appOrgIdByClerkOrgId.has(clerkId)) {
          appOrgIdByClerkOrgId.set(clerkId, id)
        }
        posthogOrgById.set(id, {
          org_name: orgBuildStr_(orgBuildRawGet_(posthogOrgsRows, r, ['org_name', 'name'])),
          org_slug: orgBuildStr_(orgBuildRawGet_(posthogOrgsRows, r, ['org_slug', 'slug'])),
          org_status: orgBuildStr_(orgBuildRawGet_(posthogOrgsRows, r, ['org_status', 'status'])),
          billing_email: orgBuildStr_(orgBuildRawGet_(posthogOrgsRows, r, ['billing_email'])),
          owner_user_id: orgBuildStr_(orgBuildRawGet_(posthogOrgsRows, r, ['owner_user_id'])),
          created_at: orgBuildStr_(orgBuildRawGet_(posthogOrgsRows, r, ['created_at'])),
          updated_at: orgBuildStr_(orgBuildRawGet_(posthogOrgsRows, r, ['updated_at'])),
          latest_subscription_id: orgBuildStr_(orgBuildRawGet_(posthogOrgsRows, r, ['latest_subscription_id'])),
          latest_subscription_status: orgBuildStr_(orgBuildRawGet_(posthogOrgsRows, r, ['latest_subscription_status'])),
          latest_active_subscription_id: orgBuildStr_(orgBuildRawGet_(posthogOrgsRows, r, ['latest_active_subscription_id']))
        })
      })

      const posthogOrgAggByOrgId = new Map()
      const posthogOrgIdBySubId = new Map()
      const posthogOrgIdByCustomerId = new Map()
      posthogOrgSubsRows.rows.forEach(r => {
        const orgId = orgBuildStr_(orgBuildRawGet_(posthogOrgSubsRows, r, ['app_org_id', 'org_id']))
        if (!orgId) return

        if (!posthogOrgAggByOrgId.has(orgId)) {
          posthogOrgAggByOrgId.set(orgId, {
            sub_ids: new Set(),
            statuses: new Set(),
            owner_user_ids: new Set(),
            full_seat_count: 0,
            lite_seat_count: 0,
            active_count: 0,
            trialing_count: 0,
            latest_updated_at: ''
          })
        }

        const agg = posthogOrgAggByOrgId.get(orgId)
        const subId = orgBuildStr_(orgBuildRawGet_(posthogOrgSubsRows, r, ['stripe_subscription_id', 'subscription_id', 'id']))
        const customerId = orgBuildStr_(orgBuildRawGet_(posthogOrgSubsRows, r, ['stripe_customer_id', 'customer_id']))
        const status = orgBuildStr_(orgBuildRawGet_(posthogOrgSubsRows, r, ['status'])).toLowerCase()
        const ownerUserId = orgBuildStr_(orgBuildRawGet_(posthogOrgSubsRows, r, ['owner_user_id']))
        const fullSeats = orgBuildNum_(orgBuildRawGet_(posthogOrgSubsRows, r, ['full_seat_count']))
        const liteSeats = orgBuildNum_(orgBuildRawGet_(posthogOrgSubsRows, r, ['lite_seat_count']))
        const updatedAt = orgBuildStr_(orgBuildRawGet_(posthogOrgSubsRows, r, ['updated_at']))

        if (subId) agg.sub_ids.add(subId)
        if (subId && !posthogOrgIdBySubId.has(subId)) posthogOrgIdBySubId.set(subId, orgId)
        if (customerId && !posthogOrgIdByCustomerId.has(customerId)) posthogOrgIdByCustomerId.set(customerId, orgId)
        if (status) agg.statuses.add(status)
        if (ownerUserId) agg.owner_user_ids.add(ownerUserId)
        agg.full_seat_count += fullSeats
        agg.lite_seat_count += liteSeats
        if (status === 'active') agg.active_count += 1
        if (status === 'trialing') agg.trialing_count += 1
        agg.latest_updated_at = orgBuildMaxIso_(agg.latest_updated_at, updatedAt)
      })

      // ---- 5) Build app promo redemption aggregates (PostHog promo_redemptions) ----
      const appPromoByOrgId = new Map()
      posthogPromoRows.rows.forEach(r => {
        const location = orgBuildStr_(orgBuildRawGet_(posthogPromoRows, r, ['redemption_location'])).toLowerCase()
        if (location && location !== 'in_app') return
        const orgId = orgBuildStr_(orgBuildRawGet_(posthogPromoRows, r, ['app_org_id', 'org_id']))
        if (!orgId) return

        if (!appPromoByOrgId.has(orgId)) {
          appPromoByOrgId.set(orgId, {
            promo_codes: new Set(),
            redemption_count: 0,
            trial_days_total: 0,
            latest_redeemed_at: ''
          })
        }

        const agg = appPromoByOrgId.get(orgId)
        const promoCode = orgBuildStr_(orgBuildRawGet_(posthogPromoRows, r, [
          'promo_code',
          'code',
          'promo_name',
          'name',
          'promo_code_id'
        ]))
        const trialDays = orgBuildNum_(orgBuildRawGet_(posthogPromoRows, r, ['trial_days']))
        const redeemedAt = orgBuildStr_(orgBuildRawGet_(posthogPromoRows, r, ['redeemed_at', 'created_at']))

        if (promoCode) agg.promo_codes.add(promoCode)
        agg.redemption_count += 1
        agg.trial_days_total += trialDays
        agg.latest_redeemed_at = orgBuildMaxIso_(agg.latest_redeemed_at, redeemedAt)
      })

      // ---- 5b) Build org_subscription_info index for org_status ----
      // Pick best subscription per org: active > trialing > everything else
      const ORG_SUB_STATUS_RANK = { 'active': 3, 'trialing': 2 }
      const orgSubInfoByOrgId = new Map()
      orgSubInfoRows.rows.forEach(r => {
        const oid = orgBuildStr_(orgBuildRawGet_(orgSubInfoRows, r, ['app_org_id', 'org_id']))
        if (!oid) return
        const status = orgBuildStr_(orgBuildRawGet_(orgSubInfoRows, r, ['status'])).toLowerCase()
        const entry = {
          status,
          first_payment_at: orgBuildStr_(orgBuildRawGet_(orgSubInfoRows, r, ['first_payment_at'])),
          has_payment_method: orgBuildStr_(orgBuildRawGet_(orgSubInfoRows, r, ['has_payment_method']))
        }
        const existing = orgSubInfoByOrgId.get(oid)
        if (!existing || (ORG_SUB_STATUS_RANK[status] || 0) > (ORG_SUB_STATUS_RANK[existing.status] || 0)) {
          orgSubInfoByOrgId.set(oid, entry)
        }
      })

      const canonSheet = getOrCreateSheet(ss, CANON_ORGS_CFG.SHEETS.CANON_ORGS)

      // ---- 7) Build output rows ----
      const updatedAt = new Date()
      const outRows = []

      orgRows.rows.forEach(o => {
        const orgId = orgBuildStr_(orgBuildRawGet_(orgRows, o, ['org_id']))
        if (!orgId) return

        const orgName = orgBuildStr_(orgBuildRawGet_(orgRows, o, ['org_name']))
        const orgSlug = orgBuildStr_(orgBuildRawGet_(orgRows, o, ['org_slug', 'slug']))
        const orgCreatedAt = orgBuildStr_(orgBuildRawGet_(orgRows, o, ['created_at', 'org_created_at']))
        const ownerUserId = orgBuildStr_(orgBuildRawGet_(orgRows, o, ['org_owner_user_id']))
        const ownerInfo = ownerUserId ? clerkUserById.get(ownerUserId) : null

        const mapped = billingMap[orgId] || {}
        const mappedSubId = orgBuildStr_(mapped.stripe_subscription_id)
        const mappedCustId = orgBuildStr_(mapped.stripe_customer_id)
        const mappedEmail = orgBuildStr_(mapped.billing_email)
        const mappedEmailKey = orgBuildNormalizeEmail_(mappedEmail)

        const memberIds = orgToMemberUserIds.get(orgId) || new Set()
        const memberEmailKeys = orgToMemberEmailKeys.get(orgId) || new Set()

        // Direct lookup: clerk_org_id → app_org_id from postgres.orgs.external_id
        const appOrgId = appOrgIdByClerkOrgId.get(orgId) || ''

        const posthogAgg = (appOrgId && posthogOrgAggByOrgId.get(appOrgId))
          || posthogOrgAggByOrgId.get(orgId)
          || null

        const linkedSubIds = new Set()
        const linkedCustomerIds = new Set()
        const linkedVia = new Set()
        let memberEmailFallback = ''

        // Primary path: postgres.org_subscriptions (org_id → stripe_subscription_id)
        if (posthogAgg && posthogAgg.sub_ids && posthogAgg.sub_ids.size) {
          posthogAgg.sub_ids.forEach(subId => linkedSubIds.add(subId))
          linkedVia.add('posthog_org_subscriptions')
        }

        // Secondary: Stripe metadata org_id (catches any subs not yet in org_subscriptions)
        const metadataSubSet = stripeSubIdsByMetadataOrgId.get(orgId)
        if (metadataSubSet && metadataSubSet.size) {
          metadataSubSet.forEach(subId => linkedSubIds.add(subId))
          linkedVia.add('stripe_metadata_org_id')
        }

        // Tertiary: manual org_billing_map
        if (mappedSubId) {
          linkedSubIds.add(mappedSubId)
          linkedVia.add('org_billing_map_sub_id')
        }
        if (mappedCustId) {
          linkedCustomerIds.add(mappedCustId)
          linkedVia.add('org_billing_map_customer_id')
          const subSet = stripeSubIdsByCustomerId.get(mappedCustId)
          if (subSet && subSet.size) {
            subSet.forEach(subId => linkedSubIds.add(subId))
            linkedVia.add('org_billing_map_customer_to_subscriptions')
          }
        }

        // Resolve customer IDs from linked subscriptions
        linkedSubIds.forEach(subId => {
          const s = stripeBySubId.get(subId)
          if (s && s.billing_customer_id) linkedCustomerIds.add(s.billing_customer_id)
        })

        // Member email fallback (for owner info, not linkage)
        memberIds.forEach(uid => {
          const link = clerkUserToStripe.get(uid)
          if (link && link.email && !memberEmailFallback) memberEmailFallback = orgBuildStr_(link.email)
        })
        if (!memberEmailFallback && memberEmailKeys.size) {
          memberEmailFallback = orgBuildFirstFromSet_(memberEmailKeys)
        }

        const posthogOrgInfo = (appOrgId && posthogOrgById.get(appOrgId))
          || posthogOrgById.get(orgId)
          || null

        const stripeSubIdsFound = new Set()
        const stripeStatuses = new Set()
        const stripeCustomerIds = new Set()
        const stripePromoCodes = new Set()
        const stripeDiscountPercents = new Set()
        const stripeDiscountAmountOff = new Set()
        const stripeDiscountDurations = new Set()

        let stripeActiveSubCount = 0
        let stripeTrialingSubCount = 0
        let stripeCanceledSubCount = 0
        let stripeSeatsPayingSum = 0
        let stripeSeatsMax = 0
        let stripeHasDiscountActive = false
        let stripeDiscountActiveCount = 0
        let stripeDiscountSubscriptionCount = 0
        let stripeArrDiscounted = 0
        let stripeArrFullPrice = 0
        let billingEmailPaying = ''
        let billingEmailAny = ''
        let billingCustomerPaying = ''
        let billingCustomerAny = ''

        linkedSubIds.forEach(subId => {
          const s = stripeBySubId.get(subId)
          if (!s) return

          stripeSubIdsFound.add(subId)
          if (s.status) stripeStatuses.add(s.status)
          if (s.status === 'active') stripeActiveSubCount += 1
          if (s.status === 'trialing') stripeTrialingSubCount += 1
          if (s.status === 'canceled') stripeCanceledSubCount += 1

          stripeSeatsMax = Math.max(stripeSeatsMax, orgBuildNum_(s.seats))
          if (s.is_paying) {
            stripeSeatsPayingSum += orgBuildNum_(s.seats)
            stripeArrDiscounted += orgBuildNum_(s.arr_discounted)
            stripeArrFullPrice += orgBuildNum_(s.arr_full_price)
          }

          if (s.billing_customer_id) {
            stripeCustomerIds.add(s.billing_customer_id)
            if (!billingCustomerAny) billingCustomerAny = s.billing_customer_id
            if (s.is_paying && !billingCustomerPaying) billingCustomerPaying = s.billing_customer_id
          }
          if (s.billing_email) {
            if (!billingEmailAny) billingEmailAny = s.billing_email
            if (s.is_paying && !billingEmailPaying) billingEmailPaying = s.billing_email
          }

          ;(s.promo_codes || []).forEach(code => {
            const val = orgBuildStr_(code)
            if (val) stripePromoCodes.add(val)
          })

          if (s.has_active_discount) {
            stripeHasDiscountActive = true
            stripeDiscountSubscriptionCount += 1
            stripeDiscountActiveCount += orgBuildNum_(s.active_discount_count)
            ;(s.active_discount_percents || []).forEach(v => {
              const val = orgBuildNum_(v)
              if (val > 0) stripeDiscountPercents.add(String(val))
            })
            ;(s.active_discount_amount_off || []).forEach(v => {
              const val = orgBuildNum_(v)
              if (val > 0) stripeDiscountAmountOff.add(String(val))
            })
            ;(s.active_discount_durations || []).forEach(v => {
              const val = orgBuildStr_(v)
              if (val) stripeDiscountDurations.add(val)
            })
          }
        })

        linkedCustomerIds.forEach(custId => {
          if (custId) stripeCustomerIds.add(custId)
        })
        if (mappedCustId) stripeCustomerIds.add(mappedCustId)

        const appPromo = (appOrgId && appPromoByOrgId.get(appOrgId)) || appPromoByOrgId.get(orgId) || null
        const appPromoCodesSet = (appPromo && appPromo.promo_codes) ? appPromo.promo_codes : new Set()
        const appPromoCodes = Array.from(appPromoCodesSet)
        const combinedPromoCodesSet = new Set()
        Array.from(stripePromoCodes).forEach(code => combinedPromoCodesSet.add(code))
        appPromoCodes.forEach(code => combinedPromoCodesSet.add(code))

        const isPaying =
          (stripeActiveSubCount + stripeTrialingSubCount) > 0 ||
          (!!posthogAgg && (posthogAgg.active_count + posthogAgg.trialing_count) > 0)
        const activeStripe = orgBuildPickPrimaryActiveStripeSubscription_(linkedSubIds, stripeBySubId)

        // Compute org_status using Ring-style bucket logic
        const orgStatus = orgBuildComputeOrgStatus_(appOrgId, isPaying, activeStripe, orgSubInfoByOrgId, stripeStatuses)
        const trialEndsAt = activeStripe ? activeStripe.trial_end : ''

        const seatsOut =
          stripeSeatsPayingSum > 0
            ? stripeSeatsPayingSum
            : ((posthogAgg && posthogAgg.full_seat_count > 0) ? posthogAgg.full_seat_count : '')

        const stripePromoCodesText = orgBuildJoinSet_(stripePromoCodes)
        const appPromoCodesText = orgBuildJoinSet_(appPromoCodesSet)
        const combinedPromoCodesText = orgBuildJoinSet_(combinedPromoCodesSet)
        const promoCodePrimary = stripePromoCodesText
          ? orgBuildFirstFromSet_(stripePromoCodes)
          : (appPromoCodesText ? orgBuildFirstFromSet_(appPromoCodesSet) : '')

        const billingEmail =
          billingEmailPaying ||
          billingEmailAny ||
          memberEmailFallback ||
          mappedEmail ||
          ''

        const billingCustomerId =
          billingCustomerPaying ||
          billingCustomerAny ||
          mappedCustId ||
          orgBuildFirstFromSet_(linkedCustomerIds) ||
          ''

        const mappingNotes = []
        if (memberIds.size === 0) mappingNotes.push('no_clerk_memberships')
        if (linkedSubIds.size === 0) mappingNotes.push('no_linked_stripe_subscription_ids')
        if (linkedSubIds.size > 1) mappingNotes.push('multiple_linked_stripe_subscription_ids')
        if (linkedSubIds.size > 0 && stripeSubIdsFound.size === 0) {
          mappingNotes.push('linked_subscription_ids_missing_in_raw_stripe_subscriptions')
        }
        if (posthogAgg && posthogAgg.sub_ids.size > 0 && stripeSubIdsFound.size === 0) {
          mappingNotes.push('posthog_subscriptions_present_but_not_in_raw_stripe_subscriptions')
        }

        outRows.push([
          orgId,
          appOrgId,
          orgName,
          orgSlug,
          orgCreatedAt,
          posthogOrgInfo ? posthogOrgInfo.org_name : '',
          posthogOrgInfo ? posthogOrgInfo.created_at : '',

          ownerUserId,
          ownerInfo ? ownerInfo.email : '',
          ownerInfo ? ownerInfo.name : '',

          isPaying === true,
          seatsOut,
          promoCodePrimary,

          billingEmail,
          billingCustomerId,

          orgBuildJoinSet_(stripeSubIdsFound),
          stripeSubIdsFound.size,
          stripeActiveSubCount,
          stripeTrialingSubCount,
          stripeCanceledSubCount,
          orgBuildJoinSet_(stripeCustomerIds),
          stripeCustomerIds.size,
          orgBuildJoinSet_(stripeStatuses),
          stripeSeatsPayingSum,
          stripeSeatsMax,
          stripeHasDiscountActive === true,
          stripeDiscountActiveCount,
          stripeDiscountSubscriptionCount,
          orgBuildJoinSet_(stripeDiscountPercents),
          orgBuildJoinSet_(stripeDiscountAmountOff),
          orgBuildJoinSet_(stripeDiscountDurations),
          stripePromoCodesText,
          stripePromoCodes.size,
          stripeArrDiscounted,
          stripeArrFullPrice,
          stripeArrDiscounted / 12,
          stripeArrFullPrice / 12,

          activeStripe ? activeStripe.subscription_id : '',
          activeStripe ? activeStripe.status : '',
          activeStripe ? activeStripe.interval : '',
          activeStripe ? activeStripe.interval_count : '',
          activeStripe ? activeStripe.amount : '',
          activeStripe ? activeStripe.seats : '',
          activeStripe ? activeStripe.arr_discounted : '',
          activeStripe ? activeStripe.arr_full_price : '',
          activeStripe ? orgBuildJoinSet_(activeStripe.promo_codes || []) : '',

          orgStatus,
          trialEndsAt,
          (() => {
            const email = ownerInfo ? (ownerInfo.email || '') : ''
            const domain = email.includes('@') ? email.split('@')[1].toLowerCase() : ''
            return domain ? (healthByDomain.get(domain) || 0) : 0
          })(),

          appPromoCodesText,
          appPromo ? appPromo.promo_codes.size : 0,
          appPromo ? appPromo.redemption_count : 0,
          appPromo ? appPromo.trial_days_total : 0,
          appPromo ? appPromo.latest_redeemed_at : '',
          combinedPromoCodesText,

          memberIds.size,
          0,
          0,
          orgBuildJoinSet_(linkedVia),
          mappingNotes.join('; '),

          updatedAt
        ])
      })

      if (outRows.length && outRows[0].length !== CANON_ORGS_CFG.CANON_HEADERS.length) {
        throw new Error(
          `canon_orgs column mismatch: headers=${CANON_ORGS_CFG.CANON_HEADERS.length} row=${outRows[0].length}`
        )
      }

      // Overwrite canon table (manual fields are carried forward via prior values above)
      writeCanonOverwrite_(canonSheet, CANON_ORGS_CFG.CANON_HEADERS, outRows)

      writeSyncLog(
        'build_canon_orgs',
        'ok',
        orgRows.rows.length,
        outRows.length,
        (new Date() - t0) / 1000,
        ''
      )
    } catch (err) {
      writeSyncLog(
        'build_canon_orgs',
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
 * Optional manual map
 * =========================
 * org_billing_map columns (header row 1), any subset is fine:
 * - org_id
 * - billing_email
 * - stripe_customer_id
 * - stripe_subscription_id
 */
function readOrgBillingMap_() {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(CANON_ORGS_CFG.SHEETS.ORG_BILLING_MAP)
  if (!sh || sh.getLastRow() < 2) return {}

  const { map } = readHeaderMap(sh, 1)
  const cOrg = map['org_id']
  if (!cOrg) return {}

  const cEmail = map['billing_email']
  const cCust = map['stripe_customer_id']
  const cSub = map['stripe_subscription_id']

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
  const out = {}

  data.forEach(r => {
    const orgId = String(r[cOrg - 1] || '').trim()
    if (!orgId) return

    out[orgId] = {
      billing_email: cEmail ? String(r[cEmail - 1] || '').trim() : '',
      stripe_customer_id: cCust ? String(r[cCust - 1] || '').trim() : '',
      stripe_subscription_id: cSub ? String(r[cSub - 1] || '').trim() : ''
    }
  })

  return out
}

/* =========================
 * Reading helpers
 * ========================= */

function readRaw_(sheet, headerRow) {
  const { map } = readHeaderMap(sheet, headerRow)
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()

  if (lastRow < headerRow + 1) {
    return {
      rows: [],
      has: () => false,
      col: () => { throw new Error(`No rows on sheet "${sheet.getName()}"`) }
    }
  }

  const rows = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()

  return {
    rows: rows,
    has: (h) => map[String(h).toLowerCase()] != null,
    col: (h) => {
      const idx = map[String(h).toLowerCase()]
      if (!idx) throw new Error(`Missing header "${h}" on sheet "${sheet.getName()}"`)
      return idx - 1 // 0-based index for array rows
    }
  }
}


function writeCanonOverwrite_(sheet, headers, rows) {
  sheet.clearContents()
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
  sheet.setFrozenRows(1)
  if (rows && rows.length) batchSetValues(sheet, 2, 1, rows, 5000)
  sheet.autoResizeColumns(1, headers.length)
}

function tryParseJson_(s) {
  try {
    return JSON.parse(String(s))
  } catch (e) {
    return null
  }
}

function orgBuildRawGet_(tbl, row, headers) {
  const candidates = Array.isArray(headers) ? headers : [headers]
  for (const h of candidates) {
    const key = String(h || '').toLowerCase()
    if (!key) continue
    if (tbl && typeof tbl.has === 'function' && tbl.has(key)) {
      return row[tbl.col(key)]
    }
  }
  return ''
}

function orgBuildStr_(v) {
  return String(v == null ? '' : v).trim()
}

function orgBuildNormalizeEmail_(email) {
  if (typeof normalizeEmail === 'function') return normalizeEmail(email)
  return orgBuildStr_(email).toLowerCase()
}

function orgBuildNum_(v) {
  const n = Number(v)
  return isFinite(n) ? n : 0
}

function orgBuildCsvList_(v) {
  const s = orgBuildStr_(v)
  if (!s) return []
  return s.split(',').map(x => orgBuildStr_(x)).filter(Boolean)
}

function orgBuildJoinSet_(setLike) {
  const arr = Array.isArray(setLike)
    ? setLike
    : Array.from(setLike || [])
  return arr
    .map(v => orgBuildStr_(v))
    .filter(Boolean)
    .filter((v, i, a) => a.indexOf(v) === i)
    .sort((a, b) => a.localeCompare(b))
    .join(', ')
}

function orgBuildFirstFromSet_(setLike) {
  const arr = Array.isArray(setLike) ? setLike.slice() : Array.from(setLike || [])
  const clean = arr.map(v => orgBuildStr_(v)).filter(Boolean).sort((a, b) => a.localeCompare(b))
  return clean.length ? clean[0] : ''
}


function orgBuildPickPrimaryActiveStripeSubscription_(subIdSet, stripeBySubId) {
  let best = null
  ;(subIdSet || new Set()).forEach(subId => {
    const s = stripeBySubId.get(subId)
    if (!s || !s.is_paying) return
    if (!best) {
      best = s
      return
    }

    const bestStatus = orgBuildStr_(best.status)
    const status = orgBuildStr_(s.status)
    if (bestStatus !== 'active' && status === 'active') {
      best = s
      return
    }

    const bestTs = Math.max(orgBuildToMs_(best.current_period_end), orgBuildToMs_(best.created_at))
    const ts = Math.max(orgBuildToMs_(s.current_period_end), orgBuildToMs_(s.created_at))
    if (ts >= bestTs) best = s
  })
  return best
}

function orgBuildToMs_(isoLike) {
  const s = orgBuildStr_(isoLike)
  if (!s) return 0
  const d = new Date(s)
  return isNaN(d.getTime()) ? 0 : d.getTime()
}

function orgBuildMaxIso_(a, b) {
  const ams = orgBuildToMs_(a)
  const bms = orgBuildToMs_(b)
  if (!ams && !bms) return ''
  if (!ams) return orgBuildStr_(b)
  if (!bms) return orgBuildStr_(a)
  return ams >= bms ? orgBuildStr_(a) : orgBuildStr_(b)
}

function orgBuildAnnualizeAmount_(amount, interval, intervalCount) {
  const amt = orgBuildNum_(amount)
  const intv = orgBuildStr_(interval).toLowerCase()
  const count = Math.max(1, orgBuildNum_(intervalCount) || 1)
  if (!amt) return 0
  if (intv === 'year' || intv === 'annual' || intv === 'yr') return amt / count
  if (intv === 'month' || intv === 'mo') return amt * (12 / count)
  return amt * 12
}

function orgBuildExtractOrgIdFromStripeRow_(stripeRows, row) {
  const direct = orgBuildStr_(orgBuildRawGet_(stripeRows, row, [
    'org_id',
    'organization_id',
    'orgid',
    'organizationid',
    'metadata_org_id',
    'metadata_orgid',
    'metadata.organization_id',
    'metadata.organizationid',
    'metadata.org_id',
    'metadata.orgid',
    'metadata[organization_id]',
    'metadata[organizationid]',
    'metadata[org_id]',
    'metadata[orgid]'
  ]))
  if (direct) return direct

  const metadataRaw = orgBuildRawGet_(stripeRows, row, ['metadata_json', 'metadata'])
  const fromMetadata = orgBuildExtractOrgIdFromMetadata_(metadataRaw)
  if (fromMetadata) return fromMetadata

  return ''
}

function orgBuildExtractOrgIdFromMetadata_(metadataRaw) {
  let obj = null
  if (metadataRaw && typeof metadataRaw === 'object') {
    obj = metadataRaw
  } else {
    const s = orgBuildStr_(metadataRaw)
    if (!s) return ''
    obj = tryParseJson_(s)
  }
  if (!obj || typeof obj !== 'object') return ''

  const candidateKeys = new Set(['orgid', 'organizationid', 'clerkorgid'])
  const stack = [obj]
  let found = ''
  while (stack.length) {
    const node = stack.pop()
    if (!node || typeof node !== 'object') continue
    if (Array.isArray(node)) {
      node.forEach(v => {
        if (v && typeof v === 'object') stack.push(v)
      })
      continue
    }

    const keys = Object.keys(node)
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i]
      const rawVal = node[k]
      const keyNorm = String(k || '').toLowerCase().replace(/[^a-z0-9]/g, '')
      if (candidateKeys.has(keyNorm)) {
        const val = orgBuildStr_(rawVal)
        if (val) {
          found = val
          break
        }
      }
      if (rawVal && typeof rawVal === 'object') stack.push(rawVal)
    }
    if (found) break
  }
  return found
}


function orgBuildStripePromoCodesFromRaw_(stripeRows, row) {
  const out = new Set()
  orgBuildCsvList_(orgBuildRawGet_(stripeRows, row, ['promo_code_all'])).forEach(c => out.add(c))
  const one = orgBuildStr_(orgBuildRawGet_(stripeRows, row, ['promo_code']))
  if (one) out.add(one)
  return Array.from(out)
}

function orgBuildStripeDiscountSummaryFromRaw_(stripeRows, row, asOfDate, baseAmount) {
  const details = orgBuildParseStripeDiscountDetailsFromRaw_(stripeRows, row)
  const asOf = (asOfDate instanceof Date && !isNaN(asOfDate.getTime())) ? asOfDate : new Date()
  const active = details.filter(d => orgBuildIsDiscountActiveNow_(d, asOf))
  const discountedAmount = orgBuildApplyActiveDiscountsToAmount_(orgBuildNum_(baseAmount), active)

  const activePercentOff = []
  const activeAmountOff = []
  const activeDuration = []

  active.forEach(d => {
    const pct = orgBuildNum_(d.percent_off)
    const amt = orgBuildNum_(d.amount_off)
    const dur = orgBuildStr_(d.duration).toLowerCase()
    if (pct > 0) activePercentOff.push(pct)
    if (amt > 0) activeAmountOff.push(amt)
    if (dur) activeDuration.push(dur)
  })

  return {
    active_count: active.length,
    active_percent_off: Array.from(new Set(activePercentOff)),
    active_amount_off: Array.from(new Set(activeAmountOff)),
    active_duration: Array.from(new Set(activeDuration)),
    discounted_amount: discountedAmount
  }
}

function orgBuildApplyActiveDiscountsToAmount_(baseAmount, activeDetails) {
  let out = orgBuildNum_(baseAmount)
  ;(activeDetails || []).forEach(d => {
    const pct = orgBuildNum_(d.percent_off)
    if (pct > 0) {
      const bounded = Math.max(0, Math.min(100, pct))
      out *= (1 - bounded / 100)
    }
    const amt = orgBuildNum_(d.amount_off)
    if (amt > 0) out -= amt
  })
  return Math.max(0, out)
}

function orgBuildParseStripeDiscountDetailsFromRaw_(stripeRows, row) {
  const out = []

  const jsonRaw = orgBuildStr_(orgBuildRawGet_(stripeRows, row, ['discount_details_json']))
  if (jsonRaw) {
    const parsed = tryParseJson_(jsonRaw)
    if (Array.isArray(parsed)) {
      parsed.forEach(d => {
        if (!d || typeof d !== 'object') return
        out.push({
          percent_off: orgBuildNum_(d.percent_off),
          amount_off: orgBuildNum_(d.amount_off),
          duration: orgBuildStr_(d.duration).toLowerCase(),
          duration_in_months: orgBuildNum_(d.duration_in_months),
          start_at: orgBuildStr_(d.start_at),
          end_at: orgBuildStr_(d.end_at)
        })
      })
    }
  }
  if (out.length) return out

  const pctAll = orgBuildCsvList_(orgBuildRawGet_(stripeRows, row, ['discount_percent_all']))
  const amtAll = orgBuildCsvList_(orgBuildRawGet_(stripeRows, row, ['discount_amount_off_all']))
  const durAll = orgBuildCsvList_(orgBuildRawGet_(stripeRows, row, ['discount_duration_all']))
  const durMonthsAll = orgBuildCsvList_(orgBuildRawGet_(stripeRows, row, ['discount_duration_months_all']))
  const startAll = orgBuildCsvList_(orgBuildRawGet_(stripeRows, row, ['discount_start_at_all']))
  const endAll = orgBuildCsvList_(orgBuildRawGet_(stripeRows, row, ['discount_end_at_all']))

  const n = Math.max(
    pctAll.length,
    amtAll.length,
    durAll.length,
    durMonthsAll.length,
    startAll.length,
    endAll.length
  )
  for (let i = 0; i < n; i++) {
    out.push({
      percent_off: orgBuildNum_(pctAll[i]),
      amount_off: orgBuildNum_(amtAll[i]),
      duration: orgBuildStr_(durAll[i]).toLowerCase(),
      duration_in_months: orgBuildNum_(durMonthsAll[i]),
      start_at: orgBuildStr_(startAll[i]),
      end_at: orgBuildStr_(endAll[i])
    })
  }
  if (out.length) return out

  const pct = orgBuildNum_(orgBuildRawGet_(stripeRows, row, ['discount_percent']))
  const amt = orgBuildNum_(orgBuildRawGet_(stripeRows, row, ['discount_amount_off']))
  const dur = orgBuildStr_(orgBuildRawGet_(stripeRows, row, ['discount_duration'])).toLowerCase()
  const durMonths = orgBuildNum_(orgBuildRawGet_(stripeRows, row, ['discount_duration_months']))
  const startAt = orgBuildStr_(orgBuildRawGet_(stripeRows, row, ['discount_start_at', 'first_payment_at', 'created_at']))
  const endAt = orgBuildStr_(orgBuildRawGet_(stripeRows, row, ['discount_end_at']))

  if (pct <= 0 && amt <= 0) return []
  return [{
    percent_off: pct,
    amount_off: amt,
    duration: dur,
    duration_in_months: durMonths,
    start_at: startAt,
    end_at: endAt
  }]
}

function orgBuildIsDiscountActiveNow_(detail, asOfDate) {
  const d = detail || {}
  const asOf = (asOfDate instanceof Date && !isNaN(asOfDate.getTime())) ? asOfDate : new Date()
  const hasValue = orgBuildNum_(d.percent_off) > 0 || orgBuildNum_(d.amount_off) > 0
  if (!hasValue) return false

  const duration = orgBuildStr_(d.duration).toLowerCase()
  const start = orgBuildStr_(d.start_at) ? new Date(orgBuildStr_(d.start_at)) : null
  const end = orgBuildStr_(d.end_at) ? new Date(orgBuildStr_(d.end_at)) : null

  if (start && !isNaN(start.getTime()) && asOf < start) return false
  if (end && !isNaN(end.getTime())) return asOf < end
  if (duration === 'forever') return true

  if (duration === 'repeating' || duration === 'once') {
    if (!start || isNaN(start.getTime())) return false
    const months = Math.max(1, orgBuildNum_(d.duration_in_months) || 1)
    const until = new Date(start.getTime())
    until.setUTCMonth(until.getUTCMonth() + months)
    return asOf < until
  }
  return false
}

/**
 * Compute org_status for CRM sync (mirrors Ring bucket logic):
 *   Paid            = active + has first payment
 *   Intent to Pay   = (active no payment yet OR trialing) + has payment method
 *   Trialing        = active/trialing without payment method or first payment
 *   Expired         = canceled, expired, or anything else
 */
function orgBuildComputeOrgStatus_(appOrgId, isPaying, activeStripe, orgSubInfoByOrgId, stripeStatuses) {
  const subInfo = appOrgId ? orgSubInfoByOrgId.get(appOrgId) : null

  if (subInfo) {
    const status = subInfo.status
    const hasFirstPayment = !!subInfo.first_payment_at
    const hasPaymentMethod = orgBuildStr_(subInfo.has_payment_method).toLowerCase()
    const hasPM = hasPaymentMethod === 'true' || hasPaymentMethod === '1' || hasPaymentMethod === 'yes'

    if (status === 'active' && hasFirstPayment) return 'Paid'
    if (status === 'active' && !hasFirstPayment && hasPM) return 'Intent to Pay'
    if (status === 'trialing' && hasPM) return 'Intent to Pay'
    if (status === 'active' && !hasFirstPayment && !hasPM) return 'Trialing'
    if (status === 'trialing' && !hasPM) return 'Trialing'
    if (status) return 'Expired'
  }

  // No org_subscription_info record → Expired
  return 'Expired'
}
