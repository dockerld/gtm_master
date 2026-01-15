/**************************************************************
 * Canonical Orgs Builder (upsert + rules) — UPDATED
 *
 * Builds/Overwrites:
 *  - canon_orgs   (key = org_id)
 *
 * Inputs:
 *  - raw_clerk_orgs
 *  - raw_clerk_memberships
 *  - raw_clerk_users              (NEW: used for Stripe subscription IDs from Clerk metadata)
 *  - raw_stripe_subscriptions
 * Optional (fallback only):
 *  - org_billing_map              (manual mapping if needed for edge cases)
 *
 * Stripe → Org mapping (NEW PRIMARY PATH):
 *  - For each org_id, look at its member clerk_user_id(s)
 *  - From raw_clerk_users, read stripe_subscription_id (or parse from metadata)
 *  - Join to raw_stripe_subscriptions by subscription_id
 *
 * Behavior:
 *  - Overwrites computed fields each run (name, created_at, is_paying, seats, promo_code, billing fields)
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
    ORG_BILLING_MAP: 'org_billing_map', // optional fallback
    CANON_ORGS: 'canon_orgs'
  },

  CANON_HEADERS: [
    'org_id',
    'org_name',
    'org_slug',
    'org_created_at',

    'is_paying',
    'seats',
    'promo_code',

    // manual fields (preserved)
    'service',
    'white_glove',
    'in_onboarding',
    'onboarding_note',

    // derived convenience/debugging
    'billing_email',
    'billing_customer_id',

    'updated_at'
  ],

  MANUAL_FIELDS: new Set(['service', 'white_glove', 'in_onboarding', 'onboarding_note']),

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

      if (!shOrgs) throw new Error(`Missing input sheet: ${CANON_ORGS_CFG.SHEETS.CLERK_ORGS}`)
      if (!shMems) throw new Error(`Missing input sheet: ${CANON_ORGS_CFG.SHEETS.CLERK_MEMBERSHIPS}`)
      if (!shUsers) throw new Error(`Missing input sheet: ${CANON_ORGS_CFG.SHEETS.CLERK_USERS}`)
      if (!shStripe) throw new Error(`Missing input sheet: ${CANON_ORGS_CFG.SHEETS.STRIPE_SUBS}`)

      const orgRows = readRaw_(shOrgs, 1)
      const memRows = readRaw_(shMems, 1)
      const userRows = readRaw_(shUsers, 1)
      const stripeRows = readRaw_(shStripe, 1)

      // Optional fallback maps
      const billingMap = readOrgBillingMap_()

      // ---- 1) Build: org_id -> Set(clerk_user_id) ----
      const orgToMemberUserIds = new Map()
      memRows.rows.forEach(r => {
        const orgId = String(r[memRows.col('org_id')] || '').trim()
        if (!orgId) return

        // membership sheet can store either "clerk_user_id" or "user_id"
        const uid =
          (memRows.has('clerk_user_id') ? String(r[memRows.col('clerk_user_id')] || '').trim() : '') ||
          (memRows.has('user_id') ? String(r[memRows.col('user_id')] || '').trim() : '') ||
          ''

        if (!uid) return

        if (!orgToMemberUserIds.has(orgId)) orgToMemberUserIds.set(orgId, new Set())
        orgToMemberUserIds.get(orgId).add(uid)
      })

      // ---- 2) Build: clerk_user_id -> Stripe linkage from Clerk metadata ----
      // Prefer explicit columns if you add them to raw_clerk_users:
      //  - stripe_subscription_id
      //  - stripe_customer_id
      // Otherwise attempt to parse from JSON metadata columns if present
      const clerkUserToStripe = new Map()

      userRows.rows.forEach(r => {
        const clerkUserId = String(r[userRows.col('clerk_user_id')] || '').trim()
        if (!clerkUserId) return

        const email = userRows.has('email') ? String(r[userRows.col('email')] || '').trim() : ''

        let stripeSubId = ''
        let stripeCustomerId = ''
        let subStatus = ''
        let tier = ''
        let plan = ''
        let trialStart = ''
        let trialEnds = ''

        if (userRows.has('stripe_subscription_id')) {
          stripeSubId = String(r[userRows.col('stripe_subscription_id')] || '').trim()
        }
        if (userRows.has('stripe_customer_id')) {
          stripeCustomerId = String(r[userRows.col('stripe_customer_id')] || '').trim()
        }
        if (userRows.has('subscription_status')) {
          subStatus = String(r[userRows.col('subscription_status')] || '').trim()
        }
        if (userRows.has('subscription_tier')) {
          tier = String(r[userRows.col('subscription_tier')] || '').trim()
        }
        if (userRows.has('current_plan')) {
          plan = String(r[userRows.col('current_plan')] || '').trim()
        }
        if (userRows.has('trial_start_date')) {
          trialStart = String(r[userRows.col('trial_start_date')] || '').trim()
        }
        if (userRows.has('trial_ends_at')) {
          trialEnds = String(r[userRows.col('trial_ends_at')] || '').trim()
        }

        // Fallback: parse from a metadata JSON column if present
        // Common names you might have: private_metadata, public_metadata, unsafe_metadata, metadata
        if (!stripeSubId || !stripeCustomerId) {
          const metaStr =
            (userRows.has('private_metadata') ? String(r[userRows.col('private_metadata')] || '') : '') ||
            (userRows.has('public_metadata') ? String(r[userRows.col('public_metadata')] || '') : '') ||
            (userRows.has('unsafe_metadata') ? String(r[userRows.col('unsafe_metadata')] || '') : '') ||
            (userRows.has('metadata') ? String(r[userRows.col('metadata')] || '') : '')

          if (metaStr) {
            const meta = tryParseJson_(metaStr)
            if (meta && typeof meta === 'object') {
              stripeSubId = stripeSubId || String(meta.stripeSubscriptionId || meta.stripe_subscription_id || '').trim()
              stripeCustomerId = stripeCustomerId || String(meta.stripeCustomerId || meta.stripe_customer_id || '').trim()
              subStatus = subStatus || String(meta.subscriptionStatus || meta.subscription_status || '').trim()
              tier = tier || String(meta.subscriptionTier || meta.subscription_tier || '').trim()
              plan = plan || String(meta.currentPlan || meta.current_plan || '').trim()
              trialStart = trialStart || String(meta.trialStartDate || meta.trial_start_date || '').trim()
              trialEnds = trialEnds || String(meta.trialEndsAt || meta.trial_ends_at || '').trim()
            }
          }
        }

        clerkUserToStripe.set(clerkUserId, {
          clerk_user_id: clerkUserId,
          email: email,
          stripe_subscription_id: stripeSubId,
          stripe_customer_id: stripeCustomerId,
          subscription_status: subStatus,
          subscription_tier: tier,
          current_plan: plan,
          trial_start_date: trialStart,
          trial_ends_at: trialEnds
        })
      })

      // ---- 3) Build: subscription_id -> Stripe aggregates ----
      // Expect raw_stripe_subscriptions columns (case-insensitive header mapping via readRaw_)
      // Recommended raw columns:
      //  - subscription_id
      //  - customer_id
      //  - customer_email
      //  - status
      //  - quantity_total
      //  - discount_promo_code (or promo_code)
      const stripeBySubId = new Map()

      stripeRows.rows.forEach(r => {
        const subId =
          (stripeRows.has('subscription_id') ? String(r[stripeRows.col('subscription_id')] || '').trim() : '') ||
          (stripeRows.has('Subscription ID') ? String(r[stripeRows.col('Subscription ID')] || '').trim() : '')

        if (!subId) return

        const status = stripeRows.has('status') ? String(r[stripeRows.col('status')] || '').toLowerCase().trim() : ''
        const isPaying = CANON_ORGS_CFG.PAYING_STATUSES.has(status)

        let seats = 0
        if (stripeRows.has('quantity_total')) seats = Number(r[stripeRows.col('quantity_total')] ?? 0) || 0
        else if (stripeRows.has('Quantity Total')) seats = Number(r[stripeRows.col('Quantity Total')] ?? 0) || 0

        const customerId =
          (stripeRows.has('customer_id') ? String(r[stripeRows.col('customer_id')] || '').trim() : '') ||
          (stripeRows.has('Customer ID') ? String(r[stripeRows.col('Customer ID')] || '').trim() : '')

        const customerEmail =
          (stripeRows.has('customer_email') ? String(r[stripeRows.col('customer_email')] || '').trim() : '') ||
          (stripeRows.has('Customer Email') ? String(r[stripeRows.col('Customer Email')] || '').trim() : '')

        const promo =
          (stripeRows.has('discount_promo_code') ? String(r[stripeRows.col('discount_promo_code')] || '').trim() : '') ||
          (stripeRows.has('promo_code') ? String(r[stripeRows.col('promo_code')] || '').trim() : '') ||
          (stripeRows.has('Promo Code') ? String(r[stripeRows.col('Promo Code')] || '').trim() : '')

        stripeBySubId.set(subId, {
          subscription_id: subId,
          status: status,
          is_paying: isPaying,
          seats: seats,
          promo_code: promo,
          billing_customer_id: customerId,
          billing_email: customerEmail
        })
      })

      // ---- 4) Derive org-wide metrics using Clerk user → subscription_id join ----
      const orgDerived = new Map()

      orgRows.rows.forEach(o => {
        const orgId = String(o[orgRows.col('org_id')] || '').trim()
        if (!orgId) return

        let isPaying = false
        let seats = 0
        let promo = ''
        let billingCustomerId = ''
        let billingEmail = ''

        // Optional fallback: if org_billing_map provides stripe_subscription_id or stripe_customer_id
        const mapped = billingMap[orgId] || {}
        let mappedSubId = mapped.stripe_subscription_id ? String(mapped.stripe_subscription_id).trim() : ''
        let mappedCustId = mapped.stripe_customer_id ? String(mapped.stripe_customer_id).trim() : ''
        let mappedEmail = mapped.billing_email ? String(mapped.billing_email).trim() : ''

        // Primary path: walk members
        const memberIds = orgToMemberUserIds.get(orgId) || new Set()
        memberIds.forEach(clerkUserId => {
          const link = clerkUserToStripe.get(clerkUserId)
          if (!link) return

          const subId = String(link.stripe_subscription_id || '').trim()
          const custId = String(link.stripe_customer_id || '').trim()

          // If we have subscription id, that is best
          if (subId && stripeBySubId.has(subId)) {
            const s = stripeBySubId.get(subId)
            if (s.is_paying) isPaying = true
            seats = Math.max(seats, Number(s.seats || 0) || 0)

            if (!promo && s.promo_code) promo = s.promo_code
            if (!billingCustomerId && s.billing_customer_id) billingCustomerId = s.billing_customer_id
            if (!billingEmail && s.billing_email) billingEmail = s.billing_email
          } else {
            // If no subscription join, still capture customer id/email from Clerk if present
            if (!billingCustomerId && custId) billingCustomerId = custId
            if (!billingEmail && link.email) billingEmail = link.email
          }
        })

        // Fallback: mapped subscription id
        if (!isPaying && mappedSubId && stripeBySubId.has(mappedSubId)) {
          const s = stripeBySubId.get(mappedSubId)
          if (s.is_paying) isPaying = true
          seats = Math.max(seats, Number(s.seats || 0) || 0)
          if (!promo && s.promo_code) promo = s.promo_code
          if (!billingCustomerId && s.billing_customer_id) billingCustomerId = s.billing_customer_id
          if (!billingEmail && s.billing_email) billingEmail = s.billing_email
        }

        // Fallback: mapped email if provided (we do not use it to join by default anymore)
        if (!billingEmail && mappedEmail) billingEmail = mappedEmail

        // Fallback: mapped customer id if provided
        if (!billingCustomerId && mappedCustId) billingCustomerId = mappedCustId

        orgDerived.set(orgId, {
          is_paying: isPaying,
          seats: seats || '',
          promo_code: promo || '',
          billing_email: billingEmail || '',
          billing_customer_id: billingCustomerId || ''
        })
      })

      // ---- 5) Preserve manual fields from existing canon_orgs ----
      const canonSheet = getOrCreateSheet(ss, CANON_ORGS_CFG.SHEETS.CANON_ORGS)
      const existing = readMaybeCanon_(canonSheet)

      // ---- 6) Build output rows ----
      const updatedAt = new Date()
      const outRows = []

      orgRows.rows.forEach(o => {
        const orgId = String(o[orgRows.col('org_id')] || '').trim()
        if (!orgId) return

        const orgName = orgRows.has('org_name') ? String(o[orgRows.col('org_name')] || '').trim() : ''
        const orgSlug =
          orgRows.has('org_slug') ? String(o[orgRows.col('org_slug')] || '').trim() :
          (orgRows.has('slug') ? String(o[orgRows.col('slug')] || '').trim() : '')

        const orgCreatedAt =
          orgRows.has('created_at') ? String(o[orgRows.col('created_at')] || '').trim() :
          (orgRows.has('org_created_at') ? String(o[orgRows.col('org_created_at')] || '').trim() : '')

        const derived = orgDerived.get(orgId) || {
          is_paying: false,
          seats: '',
          promo_code: '',
          billing_email: '',
          billing_customer_id: ''
        }

        const prior = existing.byOrgId[orgId] || {}
        const service = prior.service || ''
        const whiteGlove = prior.white_glove === true
        const inOnboarding = prior.in_onboarding === true
        const onboardingNote = prior.onboarding_note || ''

        outRows.push([
          orgId,
          orgName,
          orgSlug,
          orgCreatedAt,

          derived.is_paying === true,
          derived.seats,
          derived.promo_code,

          service,
          whiteGlove,
          inOnboarding,
          onboardingNote,

          derived.billing_email,
          derived.billing_customer_id,

          updatedAt
        ])
      })

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

/**
 * Read existing canon_orgs so we can preserve manual fields
 * Returns:
 *  - byOrgId: { [org_id]: {service, white_glove, in_onboarding, onboarding_note} }
 */
function readMaybeCanon_(sheet) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < 2) return { byOrgId: {} }

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())
  const headerMap = {}
  header.forEach((h, i) => {
    if (h) headerMap[h.toLowerCase()] = i + 1
  })

  if (!headerMap['org_id']) return { byOrgId: {} }

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()

  const cOrg = headerMap['org_id'] - 1
  const cService = headerMap['service'] ? headerMap['service'] - 1 : null
  const cWG = headerMap['white_glove'] ? headerMap['white_glove'] - 1 : null
  const cOnb = headerMap['in_onboarding'] ? headerMap['in_onboarding'] - 1 : null
  const cNote = headerMap['onboarding_note'] ? headerMap['onboarding_note'] - 1 : null

  const byOrgId = {}

  data.forEach(r => {
    const orgId = String(r[cOrg] || '').trim()
    if (!orgId) return

    byOrgId[orgId] = {
      service: cService != null ? String(r[cService] || '') : '',
      white_glove: cWG != null ? r[cWG] === true : false,
      in_onboarding: cOnb != null ? r[cOnb] === true : false,
      onboarding_note: cNote != null ? String(r[cNote] || '') : ''
    }
  })

  return { byOrgId: byOrgId }
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