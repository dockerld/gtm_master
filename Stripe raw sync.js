/**************************************************************
 * Stripe Raw Sync (overwrite-only) — IMPROVED (FIXED first_payment_at)
 *
 * Creates/overwrites:
 *  - raw_stripe_subscriptions
 *  - raw_stripe_subscription_items
 *
 * Uses Script Property:
 *  - STRIPE_KEY
 *
 * Key fixes:
 * 1) Pulls ALL subscriptions (does not exclude metadata.exclude_from_ring)
 * 2) first_payment_at computed WITHOUT per-subscription invoice calls
 * 3) Pulls payment-method presence + best-effort payment method created timestamp
 * 4) ✅ FIX: invoice->subscription linkage now uses robust extraction:
 *    - invoice.subscription
 *    - invoice.parent.subscription_details.subscription
 *    - invoice.lines.data[].parent.*.subscription
 * 5) ✅ Item-level detail: pulls per-item discounts, computes
 *    amount_after_discounts at the item level then sums
 * 6) ✅ Writes raw_stripe_subscription_items (one row per item)
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
    SUBSCRIPTIONS: 'raw_stripe_subscriptions',
    ITEMS: 'raw_stripe_subscription_items'
  }
}

// ===== Metadata field config =====
const STRIPE_EXCLUDE_META_KEY = 'exclude_from_ring'

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
    'has_payment_method',
    'payment_method_source',
    'payment_method_id',
    'payment_method_created_at',

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
    'discount_amount_off',
    'discount_duration',
    'discount_duration_months',
    'discount_start_at',
    'discount_end_at',
    'discount_percent_all',
    'discount_amount_off_all',
    'discount_duration_all',
    'discount_duration_months_all',
    'discount_start_at_all',
    'discount_end_at_all',
    'discount_count',
    'discount_details_json',
    'discounts_json',

    'promo_code',
    'promo_code_all',

    'trial_start',
    'trial_end',

    'cancel_at_period_end',
    'canceled_at',
    'current_period_end',
    'cancellation_reason',

    'metadata_json',
    'metadata_exclude_from_ring',

    // ── Item-level + discount-adjusted columns (Phase 1) ──
    'item_count',
    'amount_after_discounts',
    'amount_after_discounts_monthly',
    'amount_after_discounts_yearly',
    'items_json'
  ]

  const itemHeaders = [
    'stripe_subscription_id',
    'stripe_subscription_item_id',
    'product_id',
    'product_name',
    'price_id',
    'currency',
    'interval',
    'interval_count',
    'quantity',
    'unit_amount',
    'item_amount',
    'item_discount_count',
    'item_discount_details_json',
    'item_amount_after_discounts',
    'item_amount_after_all_discounts',
    'item_arr',
    'item_mrr'
  ]

  // 1) Fetch subscriptions (expanded customer + discounts)
  const subs = stripeFetchAllSubscriptionsExpanded_(apiKey)
  Logger.log(`Fetched ${subs.length} subscriptions total`)

  // 2) Build first_payment_at map from PAID invoices (bulk scan)
  const firstPaidBySubId = stripeBuildFirstPaidAtBySubscription_(apiKey, {
    lookbackDays: STRIPE_INVOICE_LOOKBACK_DAYS,
    pageLimit: STRIPE_INVOICE_PAGE_LIMIT,
    maxPages: STRIPE_MAX_INVOICE_PAGES_TOTAL,
    maxInvoices: STRIPE_MAX_INVOICES_TOTAL,
    requireAmountPaidPositive: STRIPE_REQUIRE_AMOUNT_PAID_POSITIVE
  })

  // 3) Collect discount coupon ids + promotion code ids for lookup
  //    Scans BOTH subscription-level AND item-level discounts
  const couponIds = new Set()
  const promoIds = new Set()
  const paymentMethodIds = new Set()
  const productIds = new Set()

  subs.forEach(sub => {
    const customerObj = (sub && sub.customer && typeof sub.customer === 'object') ? sub.customer : null

    const subDefaultPmId = stripeExtractId_(sub.default_payment_method)
    const customerDefaultPmId = stripeExtractId_(customerObj && customerObj.invoice_settings && customerObj.invoice_settings.default_payment_method)

    if (subDefaultPmId && /^pm_/.test(subDefaultPmId)) paymentMethodIds.add(subDefaultPmId)
    if (customerDefaultPmId && /^pm_/.test(customerDefaultPmId)) paymentMethodIds.add(customerDefaultPmId)

    // Subscription-level discounts
    const discountsArr = stripeNormalizeDiscounts_(sub)
    discountsArr.forEach(d => {
      const couponId = stripeExtractCouponIdFromDiscount_(d)
      const promoId = stripeExtractPromotionCodeIdFromDiscount_(d)
      if (couponId) couponIds.add(String(couponId))
      if (promoId) promoIds.add(String(promoId))
    })

    // Item-level discounts + product IDs
    const items = sub.items && sub.items.data ? sub.items.data : []
    items.forEach(it => {
      const pid = it.price && stripeExtractId_(it.price.product)
      if (pid) productIds.add(String(pid))

      const itemDiscounts = stripeNormalizeDiscounts_(it)
      itemDiscounts.forEach(d => {
        const couponId = stripeExtractCouponIdFromDiscount_(d)
        const promoId = stripeExtractPromotionCodeIdFromDiscount_(d)
        if (couponId) couponIds.add(String(couponId))
        if (promoId) promoIds.add(String(promoId))
      })
    })
  })

  const couponMap = stripeFetchCouponsMap_(apiKey, Array.from(couponIds))
  const promoMap = stripeFetchPromotionCodesMap_(apiKey, Array.from(promoIds))
  const paymentMethodMap = stripeFetchPaymentMethodsMap_(apiKey, Array.from(paymentMethodIds))
  const productMap = stripeFetchProductsMap_(apiKey, Array.from(productIds))

  // 4) Build rows (subscription + item detail)
  const allItemRows = []
  const asOfNow = new Date()

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

    // Item-level detail for items_json + raw_stripe_subscription_items
    const itemDetails = []

    // Subscription-level active discounts (applied after item discounts)
    const subDiscountsArr = stripeNormalizeDiscounts_(sub)
    const subActiveDiscounts = stripeFilterActiveDiscounts_(subDiscountsArr, couponMap, currency, asOfNow)

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

      const itemUnitCents = price.unit_amount != null ? Number(price.unit_amount) : 0
      if (idx === 0) unitCents = itemUnitCents
      totalCents += itemUnitCents * qty

      const itemAmount = itemUnitCents * qty / 100
      const itemUnitAmount = itemUnitCents / 100

      // Item-level discounts
      const itemDiscountsArr = stripeNormalizeDiscounts_(it)
      const itemActiveDiscounts = stripeFilterActiveDiscounts_(itemDiscountsArr, couponMap, currency, asOfNow)
      const itemDiscountDetails = stripeSerializeDiscountDetails_(itemDiscountsArr, couponMap, promoMap, currency)

      // Apply item-level discounts to get item_amount_after_item_discounts
      const itemAmountAfterItemDiscounts = stripeApplyDiscounts_(itemAmount, itemActiveDiscounts)

      // Apply subscription-level discounts on top to get final item amount
      const itemAmountAfterAll = stripeApplyDiscounts_(itemAmountAfterItemDiscounts, subActiveDiscounts)

      // Per-item interval for ARR/MRR
      const itemInterval = (price.recurring && price.recurring.interval)
        ? String(price.recurring.interval) : interval
      const itemIntervalCount = (price.recurring && price.recurring.interval_count)
        ? Number(price.recurring.interval_count) : intervalCount
      const itemMonths = itemInterval === 'month' ? (itemIntervalCount || 1)
        : itemInterval === 'year' ? (itemIntervalCount || 1) * 12
        : null
      const itemArr = itemMonths ? itemAmountAfterAll * (12 / itemMonths) : itemAmountAfterAll * 12
      const itemMrr = itemArr / 12

      const productId = stripeExtractId_(price.product)
      const productName = productId ? strOrBlank_((productMap[productId] || {}).name) : ''

      itemDetails.push({
        item_id: strOrBlank_(it.id),
        product_id: productId,
        product_name: productName,
        price_id: stripeExtractId_(price.id || price),
        quantity: qty,
        unit_amount: itemUnitAmount,
        item_amount: itemAmount,
        item_discount_count: itemDiscountsArr.length,
        item_amount_after_discounts: itemAmountAfterItemDiscounts,
        item_amount_after_all_discounts: itemAmountAfterAll,
        item_arr: itemArr,
        item_mrr: itemMrr
      })

      // Build row for raw_stripe_subscription_items sheet
      allItemRows.push([
        subId,
        strOrBlank_(it.id),
        productId,
        productName,
        stripeExtractId_(price.id || price),
        String(price.currency || currency || '').toUpperCase(),
        itemInterval,
        itemIntervalCount || 1,
        qty,
        itemUnitAmount,
        itemAmount,
        itemDiscountsArr.length,
        stripeSafeJson_(itemDiscountDetails),
        itemAmountAfterItemDiscounts,
        itemAmountAfterAll,
        itemArr,
        itemMrr
      ])
    })

    const unitPrice = unitCents != null ? unitCents / 100 : ''
    const amount = totalCents ? totalCents / 100 : ''

    // Discount-adjusted total: sum of per-item post-discount amounts
    const amountAfterDiscounts = itemDetails.reduce((sum, it) => sum + it.item_amount_after_all_discounts, 0)

    // normalize monthly/yearly
    let months = null
    if (interval === 'month') months = intervalCount || 1
    if (interval === 'year') months = (intervalCount || 1) * 12

    let unitMonthly = ''
    let unitYearly = ''
    let amountMonthly = ''
    let amountYearly = ''
    let amountAfterDiscountsMonthly = ''
    let amountAfterDiscountsYearly = ''

    if (months && unitPrice !== '') {
      unitMonthly = unitPrice / months
      unitYearly = unitPrice * (12 / months)
    }
    if (months && amount !== '') {
      amountMonthly = amount / months
      amountYearly = amount * (12 / months)
    }
    if (months) {
      amountAfterDiscountsMonthly = amountAfterDiscounts / months
      amountAfterDiscountsYearly = amountAfterDiscounts * (12 / months)
    } else {
      amountAfterDiscountsMonthly = amountAfterDiscounts
      amountAfterDiscountsYearly = amountAfterDiscounts * 12
    }

    // customer
    const customerObj = (sub && sub.customer && typeof sub.customer === 'object') ? sub.customer : null
    const customerId = customerObj && customerObj.id ? String(customerObj.id) : strOrBlank_(sub.customer)
    const email =
      (customerObj && customerObj.email) ||
      sub.customer_email ||
      ''

    // payment method presence (subscription-level first, then customer-level)
    const subDefaultPmId = stripeExtractId_(sub.default_payment_method)
    const customerDefaultPmId = stripeExtractId_(customerObj && customerObj.invoice_settings && customerObj.invoice_settings.default_payment_method)
    const subDefaultSourceId = stripeExtractId_(sub.default_source)
    const customerDefaultSourceId = stripeExtractId_(customerObj && customerObj.default_source)

    let hasPaymentMethod = false
    let paymentMethodSource = ''
    let paymentMethodId = ''
    if (subDefaultPmId) {
      hasPaymentMethod = true
      paymentMethodSource = 'subscription.default_payment_method'
      paymentMethodId = subDefaultPmId
    } else if (customerDefaultPmId) {
      hasPaymentMethod = true
      paymentMethodSource = 'customer.invoice_settings.default_payment_method'
      paymentMethodId = customerDefaultPmId
    } else if (subDefaultSourceId) {
      hasPaymentMethod = true
      paymentMethodSource = 'subscription.default_source'
      paymentMethodId = subDefaultSourceId
    } else if (customerDefaultSourceId) {
      hasPaymentMethod = true
      paymentMethodSource = 'customer.default_source'
      paymentMethodId = customerDefaultSourceId
    }

    // NOTE: Stripe does not expose an "attached_at" on payment_method objects.
    // This is payment_method.created (best-effort proxy), not exact attach time.
    let paymentMethodCreatedAt = ''
    if (paymentMethodId && /^pm_/.test(paymentMethodId)) {
      const pmObj = paymentMethodMap[paymentMethodId]
      if (pmObj && pmObj.created) paymentMethodCreatedAt = stripeUnixToIso_(pmObj.created)
    }

    // discount fields:
    // - keep single-value columns for backward compatibility (first discount)
    // - also persist full multi-discount detail in *_all + json columns
    let discountPercent = ''
    let discountAmountOff = ''
    let discountDuration = ''
    let discountDurationMonths = ''
    let discountStartAt = ''
    let discountEndAt = ''
    const discountPercentAll = []
    const discountAmountOffAll = []
    const discountDurationAll = []
    const discountDurationMonthsAll = []
    const discountStartAtAll = []
    const discountEndAtAll = []
    const discountDetails = []
    let discountCount = 0
    let discountDetailsJson = '[]'
    let discountsJson = '[]'
    let promoCode = ''
    const promoCodeAll = []

    const discountsArr = stripeNormalizeDiscounts_(sub)
    if (discountsArr.length > 0) {
      discountCount = discountsArr.length

      discountsArr.forEach((d, i) => {
        const couponId = stripeExtractCouponIdFromDiscount_(d)
        const couponObj = stripeExtractCouponObjectFromDiscount_(d)

        let pct = ''
        let amountOff = ''
        let dur = ''
        let durMonths = ''
        const c = (couponId && couponMap[String(couponId)]) || couponObj || null
        if (c) {
          if (c.percent_off != null) pct = c.percent_off
          if (c.amount_off != null) {
            amountOff = stripeMinorUnitsToMajor_(c.amount_off, c.currency || currency)
          }
          if (c.duration) dur = c.duration
          if (c.duration_in_months != null) durMonths = c.duration_in_months
        }

        const startIso = stripeDiscountBoundaryToIso_(d.start)
        const endIso = stripeDiscountBoundaryToIso_(d.end)

        discountPercentAll.push(pct !== '' ? String(pct) : '')
        discountAmountOffAll.push(amountOff !== '' ? String(amountOff) : '')
        discountDurationAll.push(dur ? String(dur) : '')
        discountDurationMonthsAll.push(durMonths !== '' ? String(durMonths) : '')
        discountStartAtAll.push(startIso || '')
        discountEndAtAll.push(endIso || '')

        if (i === 0) {
          discountPercent = pct
          discountAmountOff = amountOff
          discountDuration = dur
          discountDurationMonths = durMonths
          discountStartAt = startIso
          discountEndAt = endIso
        }

        let promoLabel = ''
        const promoId = stripeExtractPromotionCodeIdFromDiscount_(d)
        if (promoId) {
          const promoObjInline = stripeExtractPromotionCodeObjectFromDiscount_(d)
          const promoObj = promoObjInline || promoMap[promoId]
          promoLabel = (promoObj && promoObj.code) ? String(promoObj.code) : String(promoId)
          if (promoLabel) promoCodeAll.push(String(promoLabel))
          if (i === 0) promoCode = promoLabel
        }

        discountDetails.push({
          index: i + 1,
          coupon_id: couponId ? String(couponId) : '',
          percent_off: pct === '' ? '' : Number(pct),
          amount_off: amountOff === '' ? '' : Number(amountOff),
          duration: dur || '',
          duration_in_months: durMonths === '' ? '' : Number(durMonths),
          start_at: startIso || '',
          end_at: endIso || '',
          promotion_code: promoLabel || ''
        })
      })

      discountDetailsJson = stripeSafeJson_(discountDetails)
      discountsJson = stripeSafeJson_(discountsArr)
    }

    // metadata trace
    const md = sub.metadata || {}
    const mdExclude = strOrBlank_(md[STRIPE_EXCLUDE_META_KEY])
    const metadataJson = stripeSafeJson_(md)
    const cancelDetails = (sub && sub.cancellation_details) ? sub.cancellation_details : {}
    const cancelFeedback = strOrBlank_(cancelDetails.feedback).toLowerCase()
    const cancelReasonRaw = strOrBlank_(cancelDetails.reason).toLowerCase()
    const cancelComment = strOrBlank_(cancelDetails.comment)
    const cancelReasonMap = {
      switched_service: 'I found an alternative',
      too_expensive: 'Too expensive',
      unused: 'I no longer need it',
      missing_features: 'Missing features',
      customer_service: 'Customer service',
      low_quality: 'Low quality',
      too_complex: 'Too complex',
      other: 'Other'
    }
    let cancellationReason = cancelReasonMap[cancelFeedback] || cancelReasonMap[cancelReasonRaw] || ''
    if (!cancellationReason && cancelFeedback) cancellationReason = cancelFeedback.replace(/_/g, ' ')
    if (!cancellationReason && cancelReasonRaw) cancellationReason = cancelReasonRaw.replace(/_/g, ' ')
    if (!cancellationReason && cancelComment) cancellationReason = cancelComment

    return [
      subId,
      strOrBlank_(sub.status),
      stripeUnixToIso_(sub.created),

      firstPaymentAt,

      strOrBlank_(customerId),
      strOrBlank_(email),
      hasPaymentMethod,
      paymentMethodSource,
      paymentMethodId,
      paymentMethodCreatedAt,

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
      discountAmountOff,
      discountDuration,
      discountDurationMonths,
      discountStartAt,
      discountEndAt,
      discountPercentAll.join(', '),
      discountAmountOffAll.join(', '),
      discountDurationAll.join(', '),
      discountDurationMonthsAll.join(', '),
      discountStartAtAll.join(', '),
      discountEndAtAll.join(', '),
      discountCount || 0,
      discountDetailsJson,
      discountsJson,

      promoCode,
      promoCodeAll.join(', '),

      stripeUnixToIso_(sub.trial_start),
      stripeUnixToIso_(sub.trial_end),

      sub.cancel_at_period_end === true,
      stripeUnixToIso_(sub.canceled_at),
      stripeUnixToIso_(sub.current_period_end),
      cancellationReason,

      metadataJson,
      mdExclude,

      // ── Item-level + discount-adjusted columns ──
      items.length,
      amountAfterDiscounts,
      amountAfterDiscountsMonthly,
      amountAfterDiscountsYearly,
      stripeSafeJson_(itemDetails)
    ]
  })

  stripeOverwriteSheet_(sh, headers, rows)

  // Write items sheet
  const shItems = getOrCreateSheetSafe_(ss, STRIPE_RAW_CFG.SHEETS.ITEMS)
  stripeOverwriteSheet_(shItems, itemHeaders, allItemRows)

  const seconds = (new Date() - t0) / 1000
  writeSyncLogSafe_('stripe_pull_subscriptions_to_raw', 'ok', subs.length, rows.length, seconds,
    'items_rows=' + allItemRows.length)
  return { rows_in: subs.length, rows_out: rows.length, item_rows: allItemRows.length }
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

function stripeExtractId_(valueOrObj) {
  if (!valueOrObj) return ''
  if (typeof valueOrObj === 'string') return valueOrObj.trim()
  if (typeof valueOrObj === 'object' && valueOrObj.id) return String(valueOrObj.id).trim()
  return ''
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
      `&expand[]=data.discounts` +
      `&expand[]=data.items.data.discounts`

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
  if (!sub) return out

  const discounts = sub.discounts
  if (Array.isArray(discounts)) out.push(...discounts)
  else if (discounts && Array.isArray(discounts.data)) out.push(...discounts.data)
  else if (discounts && typeof discounts === 'object') {
    const maybeId = stripeExtractId_(discounts)
    if (maybeId || discounts.coupon || discounts.source) out.push(discounts)
  }

  if (sub.discount) out.push(sub.discount)

  const deduped = []
  const seen = new Set()
  out.forEach(d => {
    if (!d || typeof d !== 'object') return
    const key = stripeDiscountDedupeKey_(d)
    if (seen.has(key)) return
    seen.add(key)
    deduped.push(d)
  })
  return deduped
}

function stripeExtractCouponIdFromDiscount_(discount) {
  const d = discount || {}
  const sourceCoupon = d.source && d.source.coupon
  return stripeExtractId_(sourceCoupon) || stripeExtractId_(d.coupon)
}

function stripeExtractCouponObjectFromDiscount_(discount) {
  const d = discount || {}
  const sourceCoupon = d.source && d.source.coupon
  if (sourceCoupon && typeof sourceCoupon === 'object') return sourceCoupon
  if (d.coupon && typeof d.coupon === 'object') return d.coupon
  return null
}

function stripeExtractPromotionCodeIdFromDiscount_(discount) {
  const d = discount || {}
  return stripeExtractId_(d.promotion_code)
}

function stripeExtractPromotionCodeObjectFromDiscount_(discount) {
  const d = discount || {}
  if (d.promotion_code && typeof d.promotion_code === 'object') return d.promotion_code
  return null
}

function stripeDiscountDedupeKey_(discount) {
  const id = stripeExtractId_(discount)
  if (id) return `id:${id}`
  const couponId = stripeExtractCouponIdFromDiscount_(discount)
  const promoId = stripeExtractPromotionCodeIdFromDiscount_(discount)
  const start = strOrBlank_(discount && discount.start)
  const end = strOrBlank_(discount && discount.end)
  return `coupon:${couponId}|promo:${promoId}|start:${start}|end:${end}`
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

function stripeFetchPaymentMethodsMap_(apiKey, pmIds) {
  const map = {}
  if (!pmIds || !pmIds.length) return map

  pmIds.forEach(id => {
    if (!/^pm_/.test(String(id || ''))) return

    const url = `${STRIPE_RAW_CFG.API_BASE}/payment_methods/${encodeURIComponent(id)}`
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: `Bearer ${apiKey}` },
      muteHttpExceptions: true
    })

    const code = res.getResponseCode()
    if (code >= 300) {
      Logger.log(`Warning: failed to fetch payment_method ${id}: ${res.getContentText()}`)
      return
    }

    map[id] = JSON.parse(res.getContentText())
    Utilities.sleep(100)
  })

  Logger.log(`Fetched ${Object.keys(map).length} payment methods`)
  return map
}

function stripeUnixToIso_(sec) {
  if (!sec) return ''
  const d = new Date(Number(sec) * 1000)
  if (isNaN(d.getTime())) return ''
  return d.toISOString()
}

function stripeDiscountBoundaryToIso_(v) {
  if (v === null || v === undefined || v === '') return ''
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : v.toISOString()

  const n = Number(v)
  if (isFinite(n) && n > 0) return stripeUnixToIso_(n)

  const d = new Date(String(v || '').trim())
  return isNaN(d.getTime()) ? '' : d.toISOString()
}

function stripeMinorUnitsToMajor_(minorUnits, currency) {
  const n = Number(minorUnits)
  if (!isFinite(n)) return ''
  const c = String(currency || '').trim().toLowerCase()
  const zeroDecimal = new Set([
    'bif', 'clp', 'djf', 'gnf', 'jpy', 'kmf', 'krw', 'mga',
    'pyg', 'rwf', 'ugx', 'vnd', 'vuv', 'xaf', 'xof', 'xpf'
  ])
  const value = zeroDecimal.has(c) ? n : (n / 100)
  return Math.round(value * 100) / 100
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
 * Discount computation helpers
 * ========================= */

/**
 * Filter an array of raw Stripe discount objects down to only those
 * that are currently active (based on start/end/duration).
 */
function stripeFilterActiveDiscounts_(discountsArr, couponMap, currency, asOfDate) {
  const asOf = (asOfDate instanceof Date && !isNaN(asOfDate.getTime())) ? asOfDate : new Date()
  const out = []

  for (const d of (discountsArr || [])) {
    const couponId = stripeExtractCouponIdFromDiscount_(d)
    const couponObj = stripeExtractCouponObjectFromDiscount_(d)
    const c = (couponId && couponMap[String(couponId)]) || couponObj || null
    if (!c) continue

    const pct = c.percent_off != null ? Number(c.percent_off) : 0
    const amtOffMinor = c.amount_off != null ? Number(c.amount_off) : 0
    const amtOff = amtOffMinor > 0 ? stripeMinorUnitsToMajor_(amtOffMinor, c.currency || currency) : 0
    if (pct <= 0 && amtOff <= 0) continue

    const dur = String(c.duration || '').toLowerCase()
    const durMonths = Number(c.duration_in_months || 0)

    const startTs = d.start ? Number(d.start) : 0
    const endTs = d.end ? Number(d.end) : 0
    const startDate = startTs > 0 ? new Date(startTs * 1000) : null
    const endDate = endTs > 0 ? new Date(endTs * 1000) : null

    // Check if active now
    if (startDate && asOf < startDate) continue
    if (endDate && asOf >= endDate) continue

    if (dur === 'forever') {
      // always active (no end check beyond the explicit end above)
    } else if (dur === 'repeating' || dur === 'once') {
      if (!startDate) continue
      const until = new Date(startDate.getTime())
      until.setUTCMonth(until.getUTCMonth() + Math.max(1, durMonths || 1))
      if (asOf >= until) continue
    } else {
      // Unknown duration type — skip
      continue
    }

    out.push({ percent_off: pct, amount_off: amtOff, duration: dur })
  }

  return out
}

/**
 * Apply a list of active discounts to a dollar amount.
 * Percent discounts first (multiplicative), then amount-off (subtractive).
 */
function stripeApplyDiscounts_(amount, activeDiscounts) {
  let result = Number(amount || 0)
  for (const d of (activeDiscounts || [])) {
    if (d.percent_off > 0) {
      result *= (1 - Math.min(100, d.percent_off) / 100)
    }
  }
  for (const d of (activeDiscounts || [])) {
    if (d.amount_off > 0) {
      result -= d.amount_off
    }
  }
  return Math.max(0, Math.round(result * 100) / 100)
}

/**
 * Serialize discount details for JSON columns (same format as existing discount_details_json).
 */
function stripeSerializeDiscountDetails_(discountsArr, couponMap, promoMap, currency) {
  const out = []
  for (const d of (discountsArr || [])) {
    const couponId = stripeExtractCouponIdFromDiscount_(d)
    const couponObj = stripeExtractCouponObjectFromDiscount_(d)
    const c = (couponId && couponMap[String(couponId)]) || couponObj || null

    let pct = ''
    let amountOff = ''
    let dur = ''
    let durMonths = ''
    if (c) {
      if (c.percent_off != null) pct = c.percent_off
      if (c.amount_off != null) amountOff = stripeMinorUnitsToMajor_(c.amount_off, c.currency || currency)
      if (c.duration) dur = c.duration
      if (c.duration_in_months != null) durMonths = c.duration_in_months
    }

    const startIso = stripeDiscountBoundaryToIso_(d.start)
    const endIso = stripeDiscountBoundaryToIso_(d.end)

    let promoLabel = ''
    const promoId = stripeExtractPromotionCodeIdFromDiscount_(d)
    if (promoId) {
      const promoObjInline = stripeExtractPromotionCodeObjectFromDiscount_(d)
      const promoObj = promoObjInline || promoMap[promoId]
      promoLabel = (promoObj && promoObj.code) ? String(promoObj.code) : String(promoId)
    }

    out.push({
      coupon_id: couponId ? String(couponId) : '',
      percent_off: pct === '' ? '' : Number(pct),
      amount_off: amountOff === '' ? '' : Number(amountOff),
      duration: dur || '',
      duration_in_months: durMonths === '' ? '' : Number(durMonths),
      start_at: startIso || '',
      end_at: endIso || '',
      promotion_code: promoLabel || ''
    })
  }
  return out
}

/**
 * Batch-fetch product objects by ID (for product names).
 */
function stripeFetchProductsMap_(apiKey, productIds) {
  const map = {}
  if (!productIds || !productIds.length) return map

  productIds.forEach(id => {
    if (!id) return
    const url = `${STRIPE_RAW_CFG.API_BASE}/products/${encodeURIComponent(id)}`
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: `Bearer ${apiKey}` },
      muteHttpExceptions: true
    })

    const code = res.getResponseCode()
    if (code >= 300) {
      Logger.log(`Warning: failed to fetch product ${id}: ${res.getContentText()}`)
      return
    }

    map[id] = JSON.parse(res.getContentText())
    Utilities.sleep(100)
  })

  Logger.log(`Fetched ${Object.keys(map).length} products`)
  return map
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

function stripeSafeJson_(obj) {
  try {
    return JSON.stringify(obj == null ? {} : obj)
  } catch (e) {
    return '{}'
  }
}
