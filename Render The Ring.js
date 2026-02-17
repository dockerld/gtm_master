/**************************************************************
 * render_ring_view()
 *
 * Builds a money-focused view sheet: "The Ring"
 *
 * Layout:
 * - Big KPIs on top:
 *    - Paid + Promo Trial: ARR, Subscriptions, Total Seats
 *    - Paid Only: ARR, Subscriptions, Total Seats
 *    - Paid + First Payment At: ARR, Subscriptions, Total Seats
 * - Table headers on Row 3 starting Col B
 * - Table data starts Row 4 starting Col B
 *
 * Source:
 * - raw_stripe_subscriptions (header row = 1)
 * - Manual Stripe Changes (optional manual overrides)
 * - raw_posthog_user_metrics (fallback subscription->email mapping)
 *
 * Rules:
 * - Include active subscriptions
 * - Include trialing subscriptions ONLY when has_payment_method is true
 * - Exclude subscriptions listed in Manual Stripe Changes when effective reason
 *   (cancel_reason first, else exclude_reason) contains:
 *     "internal", "testing", or "duplicate" (case-insensitive)
 * - If effective reason contains "free seat":
 *     monthly subscription amount -= 30 * quantity
 *     yearly subscription amount -= 288 * quantity
 * - Display status labels:
 *     active -> "Paid"
 *     trialing(+payment method) -> "Promo Trial"
 * - Exclude subscriptions where:
 *     discount_percent == 100 AND discount_duration == 'forever'
 * - Total Seats comes from quantity_total
 *
 * ARR/MRR calculation:
 * - If interval == "year":  ARR = amount, MRR = amount / 12
 * - If interval == "month": MRR = amount, ARR = amount * 12
 *
 * Discount Duration display:
 * - 'forever' OR the number in discount_duration_months
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
 * - Adds "First Payment At" column (from raw_stripe_subscriptions.first_payment_at)
 **************************************************************/

const RING_CFG = {
  SHEET_NAME: 'The Ring',
  INPUT_SHEET: 'raw_stripe_subscriptions',
  MANUAL_CHANGES_SHEET: 'Manual Stripe Changes',

  // Clerk enrichment sources
  CLERK_USERS_SHEET: 'raw_clerk_users',
  CLERK_MEMBERSHIPS_SHEET: 'raw_clerk_memberships',
  CLERK_ORGS_SHEET: 'raw_clerk_orgs',
  POSTHOG_USERS_SHEET: 'raw_posthog_user_metrics',

  // Ring layout
  KPI_ROW_LABEL: 1,
  KPI_ROW_VALUE: 2,

  HEADER_ROW: 3,
  START_COL: 2,        // Col B
  DATA_START_ROW: 4,

  // KPI blocks
  // Combined block keeps legacy B/C/D cells used by weekly email.
  KPI_COLS: {
    COMBINED: {
      ARR: 2,            // B
      SUBSCRIPTIONS: 3,  // C
      TOTAL_SEATS: 4     // D
    },
    PAID_ONLY: {
      ARR: 6,            // F
      SUBSCRIPTIONS: 7,  // G
      TOTAL_SEATS: 8     // H
    },
    PAID_WITH_FIRST_PAYMENT: {
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
    'First Payment At',     // ✅ NEW
    'Interval',
    'Amount',
    'MRR',
    'ARR',
    'Discount %',
    'Duration',
    'Promo Code',
    'Seats'
  ],

  // Formatting
  CURRENCY_FMT: '$#,##0.00',
  INT_FMT: '0',
  PERCENT_FMT: '0.##%',
  TEXT_FMT: '@',
  DATETIME_FMT: 'yyyy-mm-dd hh:mm:ss' // ✅ NEW
}

const RING_EXCLUDED_REASON_TERMS = ['internal', 'testing', 'duplicate']
const RING_FREE_SEAT_MONTHLY_DISCOUNT = 30
const RING_FREE_SEAT_YEARLY_DISCOUNT = 288
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
      const clerkUsersSh = ss.getSheetByName(RING_CFG.CLERK_USERS_SHEET)
      const clerkMemsSh  = ss.getSheetByName(RING_CFG.CLERK_MEMBERSHIPS_SHEET)
      const clerkOrgsSh  = ss.getSheetByName(RING_CFG.CLERK_ORGS_SHEET)
      const posthogUsersSh = ss.getSheetByName(RING_CFG.POSTHOG_USERS_SHEET)

      const clerkUsers = clerkUsersSh ? readSheetObjects_(clerkUsersSh, 1) : []
      const clerkMems  = clerkMemsSh  ? readSheetObjects_(clerkMemsSh, 1)  : []
      const clerkOrgs  = clerkOrgsSh  ? readSheetObjects_(clerkOrgsSh, 1)  : []
      const posthogUsers = posthogUsersSh ? readSheetObjects_(posthogUsersSh, 1) : []

      const ringIndexes = buildRingIndexes_(clerkUsers, clerkMems, clerkOrgs, posthogUsers)

      // Stripe subscriptions
      const rows = readSheetObjects_(src, 1)
      const manualChangesBySubId = buildManualStripeChangesBySubId_(manualChangesSrc)

      const out = []
      let combinedARR = 0
      let combinedSeats = 0
      let combinedSubs = 0
      let paidOnlyARR = 0
      let paidOnlySeats = 0
      let paidOnlySubs = 0
      let paidWithFirstPaymentARR = 0
      let paidWithFirstPaymentSeats = 0
      let paidWithFirstPaymentSubs = 0

      for (const r of rows) {
        const statusRaw = str_(r.status).toLowerCase()
        const hasPaymentMethod = toBool_(r.has_payment_method)

        const stripeSubscriptionId =
          str_(r.stripe_subscription_id) ||
          str_(r.subscription_id) ||
          str_(r.subscription) ||
          str_(r.id) ||
          ''

        const manualChange = stripeSubscriptionId ? (manualChangesBySubId.get(stripeSubscriptionId) || null) : null
        const manualReason = manualChange ? (manualChange.reason || '') : ''
        const manualQuantity = manualChange ? manualChange.quantity : 0
        if (manualReason && RING_EXCLUDED_REASON_TERMS.some(term => manualReason.includes(term))) continue

        let displayStatus = ''
        if (statusRaw === 'active') displayStatus = 'Paid'
        else if (statusRaw === 'trialing' && hasPaymentMethod) displayStatus = 'Promo Trial'
        else continue

        const discountPercentRaw = num_(r.discount_percent) // 0-100
        const discountDuration = str_(r.discount_duration).toLowerCase()
        const discountDurationMonths = num_(r.discount_duration_months)

        // Exclude 100% forever discounts
        if (discountPercentRaw === 100 && discountDuration === 'forever') continue

        const interval = str_(r.interval).toLowerCase()

        // Treat raw amount as whole dollars always (1800 => $1,800.00)
        const amountRaw = moneyAmount_(r.amount)
        const amount = applyManualAmountOverride_(amountRaw, interval, manualReason, manualQuantity)

        const { mrr, arr } = computeMrrArr_(amount, interval)

        // Seat count from quantity_total
        const seats = safeInt_(r.quantity_total)

        // ✅ NEW: first payment at (ISO string from raw)
        const firstPaymentAtIso = str_(r.first_payment_at)
        const firstPaymentAtDate = isoToDateOrBlank_(firstPaymentAtIso) // Date object or ''

        // Stripe identifiers for enrichment
        const stripeEmailRaw = str_(r.customer_email || r.email || r.billing_email)
        const stripeEmailKey = normalizeEmailCompat_(stripeEmailRaw)

        const resolved = resolveRingCustomer_(stripeEmailKey, stripeSubscriptionId, ringIndexes)

        const email = resolved.email || stripeEmailRaw
        const customerName = resolved.customerName || str_(r.customer_name || r.name)
        const orgName = resolved.orgName || str_(r.org_name || r.organization_name || r.org)

        const promoCode = str_(r.promo_code)
        const durationDisplay = formatDiscountDuration_(discountDuration, discountDurationMonths)

        // Convert percent to decimal for Sheets percent format (25 -> 0.25)
        const discountPctDecimal = clamp01_(discountPercentRaw / 100)

        combinedSubs += 1
        combinedARR += arr
        combinedSeats += seats

        if (displayStatus === 'Paid') {
          paidOnlySubs += 1
          paidOnlyARR += arr
          paidOnlySeats += seats

          if (firstPaymentAtIso) {
            paidWithFirstPaymentSubs += 1
            paidWithFirstPaymentARR += arr
            paidWithFirstPaymentSeats += seats
          }
        }

        out.push([
          email,
          customerName,
          orgName,
          displayStatus,
          firstPaymentAtDate,   // ✅ NEW column value
          interval || '',
          amount,
          mrr,
          arr,
          discountPctDecimal,
          durationDisplay,
          promoCode,
          seats
        ])
      }

      // Clean rebuild
      sh.clear()

      // KPIs
      writeKpis_(sh, {
        combined: {
          arr: combinedARR,
          subscriptions: combinedSubs,
          totalSeats: combinedSeats
        },
        paidOnly: {
          arr: paidOnlyARR,
          subscriptions: paidOnlySubs,
          totalSeats: paidOnlySeats
        },
        paidWithFirstPayment: {
          arr: paidWithFirstPaymentARR,
          subscriptions: paidWithFirstPaymentSubs,
          totalSeats: paidWithFirstPaymentSeats
        }
      })

      // Headers
      sh.getRange(RING_CFG.HEADER_ROW, RING_CFG.START_COL, 1, RING_CFG.HEADERS.length).setValues([RING_CFG.HEADERS])
      sh.setFrozenRows(RING_CFG.HEADER_ROW)

      // Data
      if (out.length) {
        batchSetValuesCompat_(sh, RING_CFG.DATA_START_ROW, RING_CFG.START_COL, out, 3000)
      }

      // Formatting
      applyRingFormats_(sh, out.length)

      // Resize
      sh.autoResizeColumns(RING_CFG.START_COL, RING_CFG.HEADERS.length)

      writeSyncLogCompat_(
        'render_ring_view',
        'ok',
        rows.length,
        out.length,
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

      return { rows_in: rows.length, rows_out: out.length }
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

function buildRingIndexes_(clerkUsers, clerkMems, clerkOrgs, posthogUsers) {
  // org_id -> org_name
  const orgNameByOrgId = new Map()
  for (const o of (clerkOrgs || [])) {
    const orgId = str_(o.org_id)
    if (!orgId) continue
    const name = str_(o.org_name) || str_(o.org_slug)
    if (name) orgNameByOrgId.set(orgId, name)
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
    orgNameByOrgId,
    membershipsByEmailKey,
    usersByStripeSubId,
    usersByEmailKey,
    posthogEmailKeysByStripeSubId
  }
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

  if (!candidates.length) return { email: '', customerName: '', orgName: '' }

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

  const orgName = orgId ? (idx.orgNameByOrgId.get(orgId) || '') : ''

  return {
    email: picked.email || '',
    customerName: picked.name || '',
    orgName
  }
}

/* =========================
 * KPI + Formatting helpers
 * ========================= */

function writeKpis_(sheet, { combined, paidOnly, paidWithFirstPayment }) {
  const combinedCols = RING_CFG.KPI_COLS.COMBINED
  const paidCols = RING_CFG.KPI_COLS.PAID_ONLY
  const paidWithFirstPaymentCols = RING_CFG.KPI_COLS.PAID_WITH_FIRST_PAYMENT

  sheet.getRange(RING_CFG.KPI_ROW_LABEL, combinedCols.ARR).setValue('ARR (Paid + Promo Trial)')
  sheet.getRange(RING_CFG.KPI_ROW_LABEL, combinedCols.SUBSCRIPTIONS).setValue('Subscriptions (Paid + Promo Trial)')
  sheet.getRange(RING_CFG.KPI_ROW_LABEL, combinedCols.TOTAL_SEATS).setValue('Total Seats (Paid + Promo Trial)')

  sheet.getRange(RING_CFG.KPI_ROW_VALUE, combinedCols.ARR).setValue((combined && combined.arr) || 0)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, combinedCols.SUBSCRIPTIONS).setValue((combined && combined.subscriptions) || 0)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, combinedCols.TOTAL_SEATS).setValue((combined && combined.totalSeats) || 0)

  sheet.getRange(RING_CFG.KPI_ROW_LABEL, paidCols.ARR).setValue('ARR (Paid Only)')
  sheet.getRange(RING_CFG.KPI_ROW_LABEL, paidCols.SUBSCRIPTIONS).setValue('Subscriptions (Paid Only)')
  sheet.getRange(RING_CFG.KPI_ROW_LABEL, paidCols.TOTAL_SEATS).setValue('Total Seats (Paid Only)')

  sheet.getRange(RING_CFG.KPI_ROW_VALUE, paidCols.ARR).setValue((paidOnly && paidOnly.arr) || 0)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, paidCols.SUBSCRIPTIONS).setValue((paidOnly && paidOnly.subscriptions) || 0)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, paidCols.TOTAL_SEATS).setValue((paidOnly && paidOnly.totalSeats) || 0)

  sheet.getRange(RING_CFG.KPI_ROW_LABEL, paidWithFirstPaymentCols.ARR).setValue('ARR (Paid + First Payment At)')
  sheet.getRange(RING_CFG.KPI_ROW_LABEL, paidWithFirstPaymentCols.SUBSCRIPTIONS).setValue('Subscriptions (Paid + First Payment At)')
  sheet.getRange(RING_CFG.KPI_ROW_LABEL, paidWithFirstPaymentCols.TOTAL_SEATS).setValue('Total Seats (Paid + First Payment At)')

  sheet.getRange(RING_CFG.KPI_ROW_VALUE, paidWithFirstPaymentCols.ARR).setValue((paidWithFirstPayment && paidWithFirstPayment.arr) || 0)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, paidWithFirstPaymentCols.SUBSCRIPTIONS).setValue((paidWithFirstPayment && paidWithFirstPayment.subscriptions) || 0)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, paidWithFirstPaymentCols.TOTAL_SEATS).setValue((paidWithFirstPayment && paidWithFirstPayment.totalSeats) || 0)

  formatKpiGroup_(sheet, combinedCols)
  formatKpiGroup_(sheet, paidCols)
  formatKpiGroup_(sheet, paidWithFirstPaymentCols)

  sheet.getRange(1, 1, 2, Math.max(sheet.getLastColumn(), 14)).setVerticalAlignment('middle')
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

  const colFirstPay = colByHeader_(startCol, 'First Payment At') // ✅ NEW
  const colAmount = colByHeader_(startCol, 'Amount')
  const colMrr = colByHeader_(startCol, 'MRR')
  const colArr = colByHeader_(startCol, 'ARR')
  const colDiscPct = colByHeader_(startCol, 'Discount %')
  const colDuration = colByHeader_(startCol, 'Duration')
  const colSeats = colByHeader_(startCol, 'Seats')

  const full = sheet.getRange(startRow, startCol, nRows, RING_CFG.HEADERS.length)
  full.setNumberFormat('@')
  full.setVerticalAlignment('middle')

  // ✅ date format
  sheet.getRange(startRow, colFirstPay, nRows, 1).setNumberFormat(RING_CFG.DATETIME_FMT)

  sheet.getRange(startRow, colAmount, nRows, 1).setNumberFormat(RING_CFG.CURRENCY_FMT)
  sheet.getRange(startRow, colMrr, nRows, 1).setNumberFormat(RING_CFG.CURRENCY_FMT)
  sheet.getRange(startRow, colArr, nRows, 1).setNumberFormat(RING_CFG.CURRENCY_FMT)

  sheet.getRange(startRow, colDiscPct, nRows, 1).setNumberFormat(RING_CFG.PERCENT_FMT)
  sheet.getRange(startRow, colDuration, nRows, 1).setNumberFormat(RING_CFG.TEXT_FMT)
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

function computeMrrArr_(amount, interval) {
  const amt = Number(amount || 0) || 0
  const intv = String(interval || '').toLowerCase().trim()
  if (intv === 'year' || intv === 'annual' || intv === 'yr') return { arr: amt, mrr: amt / 12 }
  return { mrr: amt, arr: amt * 12 }
}

function moneyAmount_(raw) {
  if (raw === null || raw === undefined || raw === '') return 0
  const n = num_(raw)
  if (!isFinite(n)) return 0
  return Math.round(n * 100) / 100
}

function formatDiscountDuration_(duration, durationMonths) {
  const d = String(duration || '').toLowerCase().trim()
  if (d === 'forever') return 'forever'
  const n = Number(durationMonths)
  if (!isNaN(n) && isFinite(n) && n > 0) return String(Math.floor(n))
  return ''
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

    const reason = pickManualReason_(r)
    if (!reason) continue

    const quantityRaw =
      (r.quantity != null && r.quantity !== '') ? r.quantity :
      (r.free_seats_quantity != null && r.free_seats_quantity !== '') ? r.free_seats_quantity :
      ''
    const quantityParsed = Number(quantityRaw)
    const quantity = (isFinite(quantityParsed) && quantityParsed > 0)
      ? Math.floor(quantityParsed)
      : 0

    if (!out.has(subId)) {
      out.set(subId, { reason, quantity })
    }
  }

  return out
}

function pickManualReason_(row) {
  const cancelReason = str_(row.cancel_reason).toLowerCase()
  if (cancelReason) return cancelReason
  const excludeReason = str_(row.exclude_reason).toLowerCase()
  if (excludeReason) return excludeReason
  return str_(row.free_seats || row.free_seat).toLowerCase()
}

function applyManualAmountOverride_(amount, interval, reason, quantity) {
  const amt = Number(amount || 0) || 0
  const intv = String(interval || '').toLowerCase().trim()
  const why = String(reason || '').toLowerCase().trim()
  const qtyNum = Number(quantity)
  const qty = (isFinite(qtyNum) && qtyNum > 0) ? Math.floor(qtyNum) : 0

  if (!why.includes('free seat')) return amt
  if (!qty) return amt
  if (intv === 'month') return Math.max(0, amt - (RING_FREE_SEAT_MONTHLY_DISCOUNT * qty))
  if (intv === 'year') return Math.max(0, amt - (RING_FREE_SEAT_YEARLY_DISCOUNT * qty))
  return amt
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
