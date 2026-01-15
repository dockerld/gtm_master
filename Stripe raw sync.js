/**************************************************************
 * Stripe Raw Sync (overwrite-only) — IMPROVED (FIXED first_payment_at)
 *
 * Creates/overwrites:
 *  - raw_stripe_subscriptions
 *
 * Uses Script Property:
 *  - STRIPE_KEY
 *
 * Key fixes:
 * 1) Skips subscriptions with metadata.exclude_from_ring = true (also 1/yes)
 * 2) first_payment_at computed WITHOUT per-subscription invoice calls
 * 3) ✅ FIX: invoice->subscription linkage now uses robust extraction:
 *    - invoice.subscription
 *    - invoice.parent.subscription_details.subscription
 *    - invoice.lines.data[].parent.*.subscription
 *
 * Notes:
 * - If your paid-invoice history is huge, tune lookback + caps.
 **************************************************************/

const STRIPE_RAW_CFG = {
  API_BASE: 'https://api.stripe.com/v1',
  PAGE_LIMIT: 100,
  PAUSE_MS: 200,
  WRITE_CHUNK: 2000,

  SHEETS: {
    SUBSCRIPTIONS: 'raw_stripe_subscriptions'
  }
}

// ===== Metadata filter config =====
const STRIPE_EXCLUDE_META_KEY = 'exclude_from_ring'
const STRIPE_EXCLUDE_REASON_KEY = 'exclude_reason' // optional

// ===== First payment settings (bulk invoice scan) =====
const STRIPE_INVOICE_LOOKBACK_DAYS = 540          // ~18 months
const STRIPE_INVOICE_PAGE_LIMIT = 100             // Stripe max 100
const STRIPE_MAX_INVOICE_PAGES_TOTAL = 120        // safety cap
const STRIPE_MAX_INVOICES_TOTAL = 12000           // safety cap

// ✅ If you only want "money moved", keep this true.
// If you want to treat $0 invoices (e.g. credits) as "paid", set false.
const STRIPE_REQUIRE_AMOUNT_PAID_POSITIVE = true

function stripe_pull_subscriptions_to_raw() {
  const t0 = new Date()
  const apiKey = stripeGetSecretKey_()

  const ss = SpreadsheetApp.getActive()
  const sh = getOrCreateSheetSafe_(ss, STRIPE_RAW_CFG.SHEETS.SUBSCRIPTIONS)

  const headers = [
    'stripe_subscription_id',
    'status',
    'created_at',

    'first_payment_at',

    'stripe_customer_id',
    'customer_email',

    'currency',
    'interval',
    'interval_count',

    'quantity_total',
    'unit_price',
    'amount',

    'unit_price_monthly',
    'unit_price_yearly',
    'amount_monthly',
    'amount_yearly',

    'discount_percent',
    'discount_duration',
    'discount_duration_months',

    'promo_code',

    'cancel_at_period_end',
    'canceled_at',
    'current_period_start',
    'current_period_end',

    'metadata_exclude_from_ring',
    'metadata_exclude_reason'
  ]

  // 1) Fetch subscriptions (expanded customer + discounts)
  const subsAll = stripeFetchAllSubscriptionsExpanded_(apiKey)

  // 2) Filter excluded via metadata
  const subs = subsAll.filter(sub => !stripeShouldExcludeSub_(sub))

  Logger.log(
    `Fetched ${subsAll.length} subscriptions total; keeping ${subs.length}; excluded ${subsAll.length - subs.length} via metadata`
  )

  // 3) Build first_payment_at map from PAID invoices (bulk scan)
  const firstPaidBySubId = stripeBuildFirstPaidAtBySubscription_(apiKey, {
    lookbackDays: STRIPE_INVOICE_LOOKBACK_DAYS,
    pageLimit: STRIPE_INVOICE_PAGE_LIMIT,
    maxPages: STRIPE_MAX_INVOICE_PAGES_TOTAL,
    maxInvoices: STRIPE_MAX_INVOICES_TOTAL,
    requireAmountPaidPositive: STRIPE_REQUIRE_AMOUNT_PAID_POSITIVE
  })

  // 4) Collect discount coupon ids + promotion code ids for lookup
  const couponIds = new Set()
  const promoIds = new Set()

  subs.forEach(sub => {
    const discountsArr = stripeNormalizeDiscounts_(sub)
    discountsArr.forEach(d => {
      const couponId =
        (d.source && d.source.coupon) ||
        (d.coupon && (typeof d.coupon === 'string' ? d.coupon : d.coupon.id))

      if (couponId) couponIds.add(String(couponId))
      if (d.promotion_code) promoIds.add(String(d.promotion_code))
    })
  })

  const couponMap = stripeFetchCouponsMap_(apiKey, Array.from(couponIds))
  const promoMap = stripeFetchPromotionCodesMap_(apiKey, Array.from(promoIds))

  // 5) Build rows
  const rows = subs.map(sub => {
    const subId = strOrBlank_(sub.id)
    const firstPaymentAt = firstPaidBySubId.get(subId) || ''

    // amounts / quantities (sum of items)
    let totalCents = 0
    let totalQty = 0
    let currency = ''
    let interval = ''
    let intervalCount = 1
    let unitCents = null

    const items = sub.items && sub.items.data ? sub.items.data : []
    items.forEach((it, idx) => {
      const qty = it.quantity != null ? Number(it.quantity) : 1
      totalQty += qty

      const price = it.price
      if (!price) return

      if (!currency && price.currency) currency = String(price.currency).toUpperCase()

      if (!interval && price.recurring && price.recurring.interval) {
        interval = String(price.recurring.interval)
        intervalCount = Number(price.recurring.interval_count || 1)
      }

      if (price.unit_amount != null) {
        if (idx === 0) unitCents = Number(price.unit_amount)
        totalCents += Number(price.unit_amount) * qty
      }
    })

    const unitPrice = unitCents != null ? unitCents / 100 : ''
    const amount = totalCents ? totalCents / 100 : ''

    // normalize monthly/yearly
    let months = null
    if (interval === 'month') months = intervalCount || 1
    if (interval === 'year') months = (intervalCount || 1) * 12

    let unitMonthly = ''
    let unitYearly = ''
    let amountMonthly = ''
    let amountYearly = ''

    if (months && unitPrice !== '') {
      unitMonthly = unitPrice / months
      unitYearly = unitPrice * (12 / months)
    }
    if (months && amount !== '') {
      amountMonthly = amount / months
      amountYearly = amount * (12 / months)
    }

    // customer
    const customerId = sub.customer && sub.customer.id ? String(sub.customer.id) : strOrBlank_(sub.customer)
    const email =
      (sub.customer && sub.customer.email) ||
      sub.customer_email ||
      ''

    // discount fields (use first discount)
    let discountPercent = ''
    let discountDuration = ''
    let discountDurationMonths = ''
    let promoCode = ''

    const discountsArr = stripeNormalizeDiscounts_(sub)
    if (discountsArr.length > 0) {
      const d = discountsArr[0]

      const couponId =
        (d.source && d.source.coupon) ||
        (d.coupon && (typeof d.coupon === 'string' ? d.coupon : d.coupon.id))

      if (couponId && couponMap[String(couponId)]) {
        const c = couponMap[String(couponId)]
        if (c.percent_off != null) discountPercent = c.percent_off
        if (c.duration) discountDuration = c.duration
        if (c.duration_in_months != null) discountDurationMonths = c.duration_in_months
      }

      if (d.promotion_code) {
        const promoId = String(d.promotion_code)
        const promoObj = promoMap[promoId]
        if (promoObj && promoObj.code) promoCode = promoObj.code
        else promoCode = promoId
      }
    }

    // metadata trace
    const md = sub.metadata || {}
    const mdExclude = strOrBlank_(md[STRIPE_EXCLUDE_META_KEY])
    const mdReason = strOrBlank_(md[STRIPE_EXCLUDE_REASON_KEY])

    return [
      subId,
      strOrBlank_(sub.status),
      stripeUnixToIso_(sub.created),

      firstPaymentAt,

      strOrBlank_(customerId),
      strOrBlank_(email),

      strOrBlank_(currency),
      strOrBlank_(interval),
      intervalCount || '',

      totalQty || '',
      unitPrice,
      amount,

      unitMonthly,
      unitYearly,
      amountMonthly,
      amountYearly,

      discountPercent,
      discountDuration,
      discountDurationMonths,

      promoCode,

      sub.cancel_at_period_end === true,
      stripeUnixToIso_(sub.canceled_at),
      stripeUnixToIso_(sub.current_period_start),
      stripeUnixToIso_(sub.current_period_end),

      mdExclude,
      mdReason
    ]
  })

  stripeOverwriteSheet_(sh, headers, rows)

  const seconds = (new Date() - t0) / 1000
  writeSyncLogSafe_('stripe_pull_subscriptions_to_raw', 'ok', subsAll.length, rows.length, seconds, '')
  return { rows_in: subsAll.length, rows_out: rows.length, excluded: subsAll.length - subs.length }
}

function stripe_pull_all_raw() {
  lockWrapSafe_('stripe_pull_all_raw', () => {
    stripe_pull_subscriptions_to_raw()
  })
}

/* =========================
 * Stripe helpers
 * ========================= */

function stripeGetSecretKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('STRIPE_KEY')
  if (!key) throw new Error('Script Property STRIPE_KEY not set')
  return key
}

function stripeShouldExcludeSub_(sub) {
  const md = (sub && sub.metadata) ? sub.metadata : {}
  const raw = md ? md[STRIPE_EXCLUDE_META_KEY] : ''
  const v = String(raw || '').trim().toLowerCase()
  return v === 'true' || v === '1' || v === 'yes'
}

/**
 * Robustly extract subscription id from an invoice object.
 * Some invoices do NOT populate invoice.subscription, but do populate:
 * - invoice.parent.subscription_details.subscription  [oai_citation:1‡Stripe Docs](https://docs.stripe.com/api/invoices/object)
 */
function stripeExtractSubscriptionIdFromInvoice_(inv) {
  if (!inv) return ''

  // A) top-level
  if (inv.subscription) return String(inv.subscription).trim()

  // B) parent.subscription_details.subscription
  try {
    const sub = inv.parent && inv.parent.subscription_details && inv.parent.subscription_details.subscription
    if (sub) return String(sub).trim()
  } catch (e) {}

  // C) line parents can contain subscription references (best-effort)
  try {
    const lines = inv.lines && inv.lines.data ? inv.lines.data : []
    for (const line of lines) {
      const p = line && line.parent ? line.parent : null
      const s1 = p && p.invoice_item_details && p.invoice_item_details.subscription
      if (s1) return String(s1).trim()
      const s2 = p && p.subscription_item_details && p.subscription_item_details.subscription
      if (s2) return String(s2).trim()
    }
  } catch (e) {}

  return ''
}

/**
 * Bulk invoice scan:
 * subscription_id -> earliest paid_at (ISO)
 */
function stripeBuildFirstPaidAtBySubscription_(apiKey, opts) {
  const lookbackDays = Number(opts && opts.lookbackDays) || STRIPE_INVOICE_LOOKBACK_DAYS
  const pageLimit = Math.min(100, Math.max(1, Number(opts && opts.pageLimit) || 100))
  const maxPages = Math.max(1, Number(opts && opts.maxPages) || STRIPE_MAX_INVOICE_PAGES_TOTAL)
  const maxInvoices = Math.max(100, Number(opts && opts.maxInvoices) || STRIPE_MAX_INVOICES_TOTAL)
  const requirePaidPositive = (opts && opts.requireAmountPaidPositive) === true

  const nowSec = Math.floor(Date.now() / 1000)
  const gteSec = nowSec - Math.floor(lookbackDays * 24 * 60 * 60)

  const out = new Map() // subId -> minPaidAtSec

  let startingAfter = null
  let pages = 0
  let seen = 0

  while (true) {
    pages += 1
    if (pages > maxPages) {
      Logger.log(`Invoice scan: hit maxPages=${maxPages}, stopping early.`)
      break
    }

    let url =
      `${STRIPE_RAW_CFG.API_BASE}/invoices` +
      `?status=paid` +
      `&limit=${pageLimit}` +
      `&created[gte]=${encodeURIComponent(String(gteSec))}`

    if (startingAfter) url += `&starting_after=${encodeURIComponent(startingAfter)}`

    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: `Bearer ${apiKey}` },
      muteHttpExceptions: true
    })

    const code = res.getResponseCode()
    if (code >= 300) {
      Logger.log(`Warning: Stripe API error ${code} while listing paid invoices: ${res.getContentText()}`)
      break
    }

    const body = JSON.parse(res.getContentText())
    const data = body.data || []
    if (!data.length) break

    for (const inv of data) {
      seen += 1
      if (seen > maxInvoices) {
        Logger.log(`Invoice scan: hit maxInvoices=${maxInvoices}, stopping early.`)
        break
      }

      // ✅ subscription id (robust)
      const subId = stripeExtractSubscriptionIdFromInvoice_(inv)
      if (!subId) continue

      // paid_at
      const paidAt = inv && inv.status_transitions && inv.status_transitions.paid_at
      const paidAtSec = Number(paidAt)
      if (!isFinite(paidAtSec) || paidAtSec <= 0) continue

      // optional: require money moved
      if (requirePaidPositive) {
        const amtPaid = Number(inv && inv.amount_paid)
        if (!isFinite(amtPaid) || amtPaid <= 0) continue
      }

      const prev = out.get(subId)
      if (!prev || paidAtSec < prev) out.set(subId, paidAtSec)
    }

    if (seen > maxInvoices) break
    if (!body.has_more) break

    startingAfter = data[data.length - 1].id
    Utilities.sleep(STRIPE_RAW_CFG.PAUSE_MS)
  }

  const isoMap = new Map()
  out.forEach((sec, subId) => {
    isoMap.set(subId, stripeUnixToIso_(sec))
  })

  Logger.log(`Invoice scan done. subscriptions_with_paid_invoices=${isoMap.size}, invoices_scanned≈${seen}`)
  return isoMap
}

function stripeFetchAllSubscriptionsExpanded_(apiKey) {
  const all = []
  let startingAfter = null

  while (true) {
    let url =
      `${STRIPE_RAW_CFG.API_BASE}/subscriptions` +
      `?limit=${STRIPE_RAW_CFG.PAGE_LIMIT}` +
      `&status=all` +
      `&expand[]=data.customer` +
      `&expand[]=data.discounts`

    if (startingAfter) url += `&starting_after=${encodeURIComponent(startingAfter)}`

    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: `Bearer ${apiKey}` },
      muteHttpExceptions: true
    })

    const code = res.getResponseCode()
    if (code >= 300) throw new Error(`Stripe API error ${code} while listing subscriptions: ${res.getContentText()}`)

    const body = JSON.parse(res.getContentText())
    const data = body.data || []
    all.push(...data)

    if (!body.has_more) break
    startingAfter = data[data.length - 1].id
    Utilities.sleep(STRIPE_RAW_CFG.PAUSE_MS)
  }

  Logger.log(`Fetched ${all.length} subscriptions from Stripe`)
  return all
}

function stripeNormalizeDiscounts_(sub) {
  const out = []
  if (Array.isArray(sub.discounts)) out.push(...sub.discounts)
  if (sub.discount) out.push(sub.discount)
  return out
}

function stripeFetchCouponsMap_(apiKey, couponIds) {
  const map = {}
  if (!couponIds || !couponIds.length) return map

  couponIds.forEach(id => {
    const url = `${STRIPE_RAW_CFG.API_BASE}/coupons/${encodeURIComponent(id)}`
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: `Bearer ${apiKey}` },
      muteHttpExceptions: true
    })

    const code = res.getResponseCode()
    if (code >= 300) {
      Logger.log(`Warning: failed to fetch coupon ${id}: ${res.getContentText()}`)
      return
    }

    map[id] = JSON.parse(res.getContentText())
    Utilities.sleep(100)
  })

  Logger.log(`Fetched ${Object.keys(map).length} coupons`)
  return map
}

function stripeFetchPromotionCodesMap_(apiKey, promoIds) {
  const map = {}
  if (!promoIds || !promoIds.length) return map

  promoIds.forEach(id => {
    const url = `${STRIPE_RAW_CFG.API_BASE}/promotion_codes/${encodeURIComponent(id)}`
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: `Bearer ${apiKey}` },
      muteHttpExceptions: true
    })

    const code = res.getResponseCode()
    if (code >= 300) {
      Logger.log(`Warning: failed to fetch promotion_code ${id}: ${res.getContentText()}`)
      return
    }

    map[id] = JSON.parse(res.getContentText())
    Utilities.sleep(100)
  })

  Logger.log(`Fetched ${Object.keys(map).length} promotion codes`)
  return map
}

function stripeUnixToIso_(sec) {
  if (!sec) return ''
  const d = new Date(Number(sec) * 1000)
  if (isNaN(d.getTime())) return ''
  return d.toISOString()
}

function stripeOverwriteSheet_(sheet, headers, rows) {
  sheet.clearContents()
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
  sheet.setFrozenRows(1)

  if (rows && rows.length) {
    batchSetValuesSafe_(sheet, 2, 1, rows, STRIPE_RAW_CFG.WRITE_CHUNK)
  }

  sheet.autoResizeColumns(1, headers.length)
}

/* =========================
 * Minimal shared utilities (fallbacks)
 * ========================= */

function getOrCreateSheetSafe_(ss, name) {
  if (typeof getOrCreateSheet === 'function') {
    try { return getOrCreateSheet(ss, name) } catch (e) {}
    try { return getOrCreateSheet(name) } catch (e) {}
  }
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function batchSetValuesSafe_(sheet, startRow, startCol, values, chunkSize) {
  if (typeof batchSetValues === 'function') return batchSetValues(sheet, startRow, startCol, values, chunkSize)
  const size = chunkSize || 2000
  for (let i = 0; i < values.length; i += size) {
    const chunk = values.slice(i, i + size)
    sheet.getRange(startRow + i, startCol, chunk.length, chunk[0].length).setValues(chunk)
  }
}

function writeSyncLogSafe_(step, status, rowsIn, rowsOut, seconds, error) {
  if (typeof writeSyncLog === 'function') {
    return writeSyncLog(step, status, rowsIn, rowsOut, seconds, error || '')
  }
  Logger.log(`[SYNCLOG missing] ${step} ${status} rows_in=${rowsIn} rows_out=${rowsOut} seconds=${seconds} error=${error || ''}`)
}

function lockWrapSafe_(name, fn) {
  if (typeof lockWrap === 'function') return lockWrap(fn)
  const lock = LockService.getScriptLock()
  if (!lock.tryLock(30000)) throw new Error(`Could not obtain lock: ${name}`)
  try { return fn() } finally { lock.releaseLock() }
}

function strOrBlank_(v) {
  if (v === null || v === undefined) return ''
  return String(v).trim()
}