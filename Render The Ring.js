/**************************************************************
 * render_ring_view()
 *
 * Builds a money-focused view sheet: "The Ring"
 *
 * Layout:
 * - Big KPIs on top:
 *    ARR, Subscriptions, Total Seats
 * - Table headers on Row 3 starting Col B
 * - Table data starts Row 4 starting Col B
 *
 * Source:
 * - raw_stripe_subscriptions (header row = 1)
 *
 * Rules:
 * - Only include active subscriptions
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

  // Clerk enrichment sources
  CLERK_USERS_SHEET: 'raw_clerk_users',
  CLERK_MEMBERSHIPS_SHEET: 'raw_clerk_memberships',
  CLERK_ORGS_SHEET: 'raw_clerk_orgs',

  // Ring layout
  KPI_ROW_LABEL: 1,
  KPI_ROW_VALUE: 2,

  HEADER_ROW: 3,
  START_COL: 2,        // Col B
  DATA_START_ROW: 4,

  // KPIs positions (B,C,D)
  KPI_COLS: {
    ARR: 2,            // B
    SUBSCRIPTIONS: 3,  // C
    TOTAL_SEATS: 4     // D
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

function render_ring_view() {
  lockWrapCompat_('render_ring_view', () => {
    const t0 = new Date()
    try {
      const ss = SpreadsheetApp.getActive()
      const sh = getOrCreateSheetCompat_(ss, RING_CFG.SHEET_NAME)

      const src = ss.getSheetByName(RING_CFG.INPUT_SHEET)
      if (!src) throw new Error(`Missing input sheet: ${RING_CFG.INPUT_SHEET}`)

      // Load Clerk sources for enrichment
      const clerkUsersSh = ss.getSheetByName(RING_CFG.CLERK_USERS_SHEET)
      const clerkMemsSh  = ss.getSheetByName(RING_CFG.CLERK_MEMBERSHIPS_SHEET)
      const clerkOrgsSh  = ss.getSheetByName(RING_CFG.CLERK_ORGS_SHEET)

      const clerkUsers = clerkUsersSh ? readSheetObjects_(clerkUsersSh, 1) : []
      const clerkMems  = clerkMemsSh  ? readSheetObjects_(clerkMemsSh, 1)  : []
      const clerkOrgs  = clerkOrgsSh  ? readSheetObjects_(clerkOrgsSh, 1)  : []

      const ringIndexes = buildRingIndexes_(clerkUsers, clerkMems, clerkOrgs)

      // Stripe subscriptions
      const rows = readSheetObjects_(src, 1)

      const out = []
      let totalARR = 0
      let totalSeats = 0
      let subsCount = 0

      for (const r of rows) {
        const status = str_(r.status).toLowerCase()
        if (status !== 'active') continue

        const discountPercentRaw = num_(r.discount_percent) // 0-100
        const discountDuration = str_(r.discount_duration).toLowerCase()
        const discountDurationMonths = num_(r.discount_duration_months)

        // Exclude 100% forever discounts
        if (discountPercentRaw === 100 && discountDuration === 'forever') continue

        const interval = str_(r.interval).toLowerCase()

        // Treat raw amount as whole dollars always (1800 => $1,800.00)
        const amount = moneyAmount_(r.amount)

        const { mrr, arr } = computeMrrArr_(amount, interval)

        // Seat count from quantity_total
        const seats = safeInt_(r.quantity_total)

        // ✅ NEW: first payment at (ISO string from raw)
        const firstPaymentAtIso = str_(r.first_payment_at)
        const firstPaymentAtDate = isoToDateOrBlank_(firstPaymentAtIso) // Date object or ''

        // Stripe identifiers for enrichment
        const stripeEmailRaw = str_(r.customer_email || r.email || r.billing_email)
        const stripeEmailKey = normalizeEmailCompat_(stripeEmailRaw)

        const stripeSubscriptionId =
          str_(r.stripe_subscription_id) ||
          str_(r.subscription_id) ||
          str_(r.subscription) ||
          str_(r.id) ||
          ''

        const resolved = resolveRingCustomer_(stripeEmailKey, stripeSubscriptionId, ringIndexes)

        const email = resolved.email || stripeEmailRaw
        const customerName = resolved.customerName || str_(r.customer_name || r.name)
        const orgName = resolved.orgName || str_(r.org_name || r.organization_name || r.org)

        const promoCode = str_(r.promo_code)
        const durationDisplay = formatDiscountDuration_(discountDuration, discountDurationMonths)

        // Convert percent to decimal for Sheets percent format (25 -> 0.25)
        const discountPctDecimal = clamp01_(discountPercentRaw / 100)

        subsCount += 1
        totalARR += arr
        totalSeats += seats

        out.push([
          email,
          customerName,
          orgName,
          'active',
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
        arr: totalARR,
        subscriptions: subsCount,
        totalSeats
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

function buildRingIndexes_(clerkUsers, clerkMems, clerkOrgs) {
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
  for (const u of (clerkUsers || [])) {
    const subId = str_(u.stripe_subscription_id || u.stripeSubscriptionId)
    if (!subId) continue

    const email = str_(u.email)
    const emailKey = str_(u.email_key) || normalizeEmailCompat_(email)
    if (!emailKey) continue

    const name = str_(u.name)
    const orgId = str_(u.org_id)

    if (!usersByStripeSubId.has(subId)) usersByStripeSubId.set(subId, [])
    usersByStripeSubId.get(subId).push({ email, emailKey, name, orgId })
  }

  return {
    orgNameByOrgId,
    membershipsByEmailKey,
    usersByStripeSubId
  }
}

function resolveRingCustomer_(stripeEmailKey, stripeSubscriptionId, idx) {
  const subId = str_(stripeSubscriptionId)
  const candidates = (subId && idx.usersByStripeSubId.has(subId))
    ? idx.usersByStripeSubId.get(subId).slice()
    : []

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

function writeKpis_(sheet, { arr, subscriptions, totalSeats }) {
  sheet.getRange(RING_CFG.KPI_ROW_LABEL, RING_CFG.KPI_COLS.ARR).setValue('ARR')
  sheet.getRange(RING_CFG.KPI_ROW_LABEL, RING_CFG.KPI_COLS.SUBSCRIPTIONS).setValue('Subscriptions')
  sheet.getRange(RING_CFG.KPI_ROW_LABEL, RING_CFG.KPI_COLS.TOTAL_SEATS).setValue('Total Seats')

  sheet.getRange(RING_CFG.KPI_ROW_VALUE, RING_CFG.KPI_COLS.ARR).setValue(arr || 0)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, RING_CFG.KPI_COLS.SUBSCRIPTIONS).setValue(subscriptions || 0)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, RING_CFG.KPI_COLS.TOTAL_SEATS).setValue(totalSeats || 0)

  const labelRange = sheet.getRange(RING_CFG.KPI_ROW_LABEL, RING_CFG.KPI_COLS.ARR, 1, 3)
  labelRange.setFontWeight('bold').setHorizontalAlignment('center')

  const valueRange = sheet.getRange(RING_CFG.KPI_ROW_VALUE, RING_CFG.KPI_COLS.ARR, 1, 3)
  valueRange.setFontWeight('bold').setFontSize(22).setHorizontalAlignment('center')

  sheet.getRange(RING_CFG.KPI_ROW_VALUE, RING_CFG.KPI_COLS.ARR).setNumberFormat(RING_CFG.CURRENCY_FMT)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, RING_CFG.KPI_COLS.SUBSCRIPTIONS).setNumberFormat(RING_CFG.INT_FMT)
  sheet.getRange(RING_CFG.KPI_ROW_VALUE, RING_CFG.KPI_COLS.TOTAL_SEATS).setNumberFormat(RING_CFG.INT_FMT)

  sheet.getRange(1, 1, 2, Math.max(sheet.getLastColumn(), 10)).setVerticalAlignment('middle')
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
