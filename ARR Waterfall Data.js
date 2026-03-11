/**************************************************************
 * render_arr_raw_data_view()
 *
 * Builds/overwrites the dataset sheet: "arr_raw_data"
 *
 * Assumptions:
 * - Header row is ROW 2 (as you described)
 * - One row per org in Clerk (raw_clerk_orgs)
 * - Pulls:
 *   - org_id, org_name, org_creation_date from raw_clerk_orgs
 *   - org_email = earliest OWNER membership email (fallback to earliest admin, then earliest member)
 *   - stripe_email + subscription fields from raw_stripe_subscriptions
 *     joined via raw_clerk_users (stripe_subscription_id) + raw_clerk_memberships (org_id)
 *
 * Trial end date logic:
 * A) Standard: trial_start_date + 14 days
 * B) If a subscription start date is within that 14-day window AND
 *    discount_percent == 100 AND discount_duration_months > 0
 *    => trial_end_date = subscription_start_date + discount_duration_months
 *
 * NEW:
 * - Adds/ensures a column named "status" (VALIDATION status)
 * - If discount_percent == 100 AND discount_duration == "forever"
 *   => status = "Invalid"
 *
 * Notes:
 * - This is overwrite-only for the view columns. If you need to preserve manual notes,
 *   we can add manual-preserve logic later (like Sauron).
 **************************************************************/

const ARR_RAW_CFG = {
  SHEET_NAME: "arr_raw_data",
  HEADER_ROW: 2,
  DATA_START_ROW: 3,

  INPUTS: {
    ORG_SUBS_INFO: "org_subscription_info",
    MANUAL_CHANGES: "Manual Stripe Changes",
    STRIPE_SUBS: "raw_stripe_subscriptions",
  },

  HEADERS: [
    "org_id",
    "org_name",
    "org_creation_date",
    "first_payment_date",
    "churn_date",
    "sign_up_cohort_month",
    "first_payment_cohort_month",
    "current_status",
    "plan_name",
    "billing_frequency",
    "total_arr",
    "subscription_start_date",
  ],
}

function render_arr_raw_data_view() {
  return ARR_lockWrap_("render_arr_raw_data_view", () => {
    const t0 = new Date()
    const ss = SpreadsheetApp.getActive()

    const shOut = ARR_getOrCreateSheet_(ss, ARR_RAW_CFG.SHEET_NAME)
    const shOrgSubsInfo = ss.getSheetByName(ARR_RAW_CFG.INPUTS.ORG_SUBS_INFO)
    if (!shOrgSubsInfo) throw new Error(`Missing sheet: ${ARR_RAW_CFG.INPUTS.ORG_SUBS_INFO}`)
    const shManual = ss.getSheetByName(ARR_RAW_CFG.INPUTS.MANUAL_CHANGES)
    const shStripe = ss.getSheetByName(ARR_RAW_CFG.INPUTS.STRIPE_SUBS)
    if (!shStripe) throw new Error(`Missing sheet: ${ARR_RAW_CFG.INPUTS.STRIPE_SUBS}`)

    const header = ARR_RAW_CFG.HEADERS.slice()
    shOut.getRange(ARR_RAW_CFG.HEADER_ROW, 1, 1, Math.max(shOut.getMaxColumns(), header.length)).clearContent()
    shOut.getRange(ARR_RAW_CFG.HEADER_ROW, 1, 1, header.length).setValues([header])
    const excludedSubIds = ARR_buildInternalExcludeSubIdSet_(shManual)
    const orgSubsAll = ARR_readSheetObjects_(shOrgSubsInfo, 1)
    const stripeRows = ARR_readSheetObjects_(shStripe, 1)
    const stripeBySubId = ARR_buildStripeBySubscriptionId_(stripeRows)
    const orgSubs = (orgSubsAll || []).filter(r => {
      const subId =
        ARR_str_(r.latest_subscription_id) ||
        ARR_str_(r.stripe_subscription_id) ||
        ARR_str_(r.subscription_id) ||
        ARR_str_(r.id)
      if (!subId) return true
      return !excludedSubIds.has(subId)
    })

    // Now build output rows in the existing column order
    const outRows = []

    // Sort for stability
    const sortedOrgs = orgSubs
      .map(o => ({
        org_id: ARR_str_(o.app_org_id || o.org_id),
        org_name: ARR_str_(o.org_name),
      }))
      .filter(o => o.org_id)
      .sort((a, b) => (a.org_name || "").localeCompare(b.org_name || "") || a.org_id.localeCompare(b.org_id))

    const infoByOrgId = new Map()
    ;(orgSubs || []).forEach(r => {
      const orgId = ARR_str_(r.app_org_id || r.org_id)
      if (!orgId || infoByOrgId.has(orgId)) return
      infoByOrgId.set(orgId, r)
    })

    for (const o of sortedOrgs) {
      const orgId = o.org_id
      const info = infoByOrgId.get(orgId) || {}
      const orgName = ARR_str_(info.org_name) || o.org_name

      const statusRaw = ARR_str_(info.status).toLowerCase()
      const orgCreatedAtIso = ARR_toIsoOrBlank_(info.org_created_at)
      const firstPaymentIso = ARR_toIsoOrBlank_(info.first_payment_at)
      const subscriptionStartIso = ARR_toIsoOrBlank_(info.subscription_created_at_date)
      const purchaseDate = firstPaymentIso
      const signUpCohortMonth = ARR_isoToCohortMonth_(orgCreatedAtIso || "")
      const firstPaymentCohortMonth = ARR_isoToCohortMonth_(purchaseDate || "")

      // Current status from org_subscription_info
      const currentStatus = ARR_str_(info.status)
      const interval = ARR_str_(info.interval).toLowerCase()
      const intervalCount = Math.max(1, ARR_num_(info.interval_count) || 1)
      const billingFrequency =
        interval === "year" ? "yearly" :
        interval === "month" ? "monthly" :
        (interval ? interval : "")
      const planName =
        ARR_str_(info.plan_name) ||
        (billingFrequency ? `plan_${billingFrequency}_${intervalCount}` : "")

      // ARR group rule: only active + first_payment_at; discount-aware from org_subscription_info fields.
      const arrEligible = statusRaw === "active" && !!firstPaymentIso
      const subId =
        ARR_str_(info.latest_subscription_id) ||
        ARR_str_(info.stripe_subscription_id) ||
        ARR_str_(info.subscription_id) ||
        ARR_str_(info.id)
      const stripeRow = subId ? (stripeBySubId.get(subId) || null) : null
      const arrSourceRow = stripeRow || info
      const totalArr = arrEligible
        ? ARR_computeEffectiveArrFromSubscriptionRow_(arrSourceRow, new Date())
        : 0
      const churnDate = ARR_toIsoOrBlank_(info.churn_date)

      // Map to your output headers by name (so column order can evolve safely)
      const rowObj = {
        org_id: orgId,
        org_name: orgName,
        org_creation_date: orgCreatedAtIso,
        first_payment_date: purchaseDate || "",
        churn_date: churnDate,
        sign_up_cohort_month: signUpCohortMonth,
        first_payment_cohort_month: firstPaymentCohortMonth,
        current_status: currentStatus,
        plan_name: planName,
        billing_frequency: billingFrequency,
        total_arr: totalArr,
        subscription_start_date: subscriptionStartIso,
      }

      outRows.push(ARR_rowFromHeader_(header, rowObj))
    }

    // Write output (clear only data region, keep header row 2 intact)
    ARR_clearDataRegion_(shOut, ARR_RAW_CFG.DATA_START_ROW, header.length)
    if (outRows.length) {
      ARR_batchSetValues_(shOut, ARR_RAW_CFG.DATA_START_ROW, 1, outRows, 2000)
    }

    ARR_applyArrRawFormats_(shOut, header, outRows.length)

    shOut.setFrozenRows(ARR_RAW_CFG.HEADER_ROW)
    shOut.autoResizeColumns(1, header.length)

    const seconds = (new Date() - t0) / 1000
    if (typeof writeSyncLog === "function") {
      writeSyncLog("render_arr_raw_data_view", "ok", orgSubsAll.length, outRows.length, seconds, "")
    } else {
      Logger.log(`[render_arr_raw_data_view] ok rows_in=${orgSubsAll.length} rows_out=${outRows.length} seconds=${seconds}`)
    }

    return { rows_in: orgSubsAll.length, rows_out: outRows.length }
  })
}

/* ============================================================
 * Trial end logic
 * ============================================================ */

function ARR_computeTrialEndIso_({ trialStartIso, subscriptions, trialDays }) {
  const ts = ARR_parseIsoDate_(trialStartIso)
  if (!ts) return ""

  const standardEnd = new Date(ts.getTime() + Number(trialDays || 14) * 24 * 60 * 60 * 1000)
  const candidate = ARR_pickTrialExtensionSub_(subscriptions, ts, standardEnd)

  if (candidate) {
    const end = ARR_addMonths_(candidate.start, candidate.months)
    return end.toISOString()
  }

  return standardEnd.toISOString()
}

function ARR_pickTrialExtensionSub_(subscriptions, trialStart, standardEnd) {
  const startMs = trialStart.getTime()
  const endMs = standardEnd.getTime()
  let best = null

  ;(subscriptions || []).forEach(sub => {
    const pct = Number(sub.discount_percent)
    const months = Number(sub.discount_duration_months)
    if (pct !== 100 || !isFinite(months) || months <= 0) return

    const startIso = ARR_getSubscriptionStartIso_(sub)
    const start = ARR_parseIsoDate_(startIso)
    if (!start) return

    const t = start.getTime()
    if (t < startMs || t > endMs) return

    if (!best || t < best.start.getTime()) best = { start, months }
  })

  return best
}

function ARR_getSubscriptionStartIso_(sub) {
  if (!sub) return ""
  return ARR_toIsoOrBlank_(sub.created_at)
}

function ARR_addMonths_(dateObj, months) {
  const d = new Date(dateObj.getTime())
  const m = Number(months) || 0
  const day = d.getUTCDate()
  d.setUTCMonth(d.getUTCMonth() + m)

  // Best-effort clamp for month length differences:
  if (d.getUTCDate() !== day) d.setUTCDate(0)
  return d
}

/* ============================================================
 * Subscription rollup per org
 * ============================================================ */

function ARR_emptyRollup_() {
  return {
    has_any: false,
    has_active: false,

    stripe_email: "",
    purchase_date: "",
    churn_date: "",

    plan_name: "",
    billing_frequency: "",
    total_arr: 0,

    discount_percent: "",
    discount_duration: "",
    discount_duration_months: "",
    discount_start_at: "",
    discount_end_at: "",

    current_status: "",
  }
}

function ARR_buildOrgSubscriptionRollup_(subIdsByOrgId, stripeBySubId, opts) {
  const out = new Map()
  const options = opts || {}
  const excludedSubIds = options.excludedSubIds instanceof Set ? options.excludedSubIds : new Set()
  const asOfDate = (options.asOfDate instanceof Date && !isNaN(options.asOfDate.getTime()))
    ? options.asOfDate
    : new Date()

  subIdsByOrgId.forEach((subIdSet, orgId) => {
    const ids = Array.from(subIdSet || [])
      .filter(Boolean)
      .filter(id => !excludedSubIds.has(id))
    if (!ids.length) {
      out.set(orgId, ARR_emptyRollup_())
      return
    }

    const rows = ids
      .map(id => stripeBySubId.get(id))
      .filter(Boolean)

    if (!rows.length) {
      out.set(orgId, ARR_emptyRollup_())
      return
    }

    // Determine lifecycle-active (any active) vs ARR-eligible paid-active rows.
    // ARR group rule: only status=active with first_payment_at populated.
    const activeRows = rows.filter(r => String(r.status || "").toLowerCase() === "active")
    const hasActive = activeRows.length > 0
    const paidActiveRows = activeRows.filter(r => !!ARR_toIsoOrBlank_(r.first_payment_at))
    const hasPaidActive = paidActiveRows.length > 0

    // Purchase date = earliest first_payment_at across all subs (do NOT overwrite later)
    const purchaseIso = ARR_minIso_(rows.map(r => r.first_payment_at).filter(Boolean))

    // Churn date = if no active subs, take max(canceled_at) among canceled rows
    let churnIso = ""
    if (!hasActive) {
      churnIso = ARR_maxIso_(
        rows
          .filter(r => String(r.status || "").toLowerCase() === "canceled")
          .map(r => r.canceled_at)
          .filter(Boolean)
      )
    }

    // Pick the “current” subscription to derive plan/billing/arr:
    // - prefer active with latest created_at; else latest created_at among all
    const bestForPlan = ARR_pickBestSubscriptionRow_(hasActive ? activeRows : rows)

    const interval = String(bestForPlan.interval || "").toLowerCase()
    const intervalCount = Number(bestForPlan.interval_count || 1) || 1

    const billingFrequency =
      interval === "year" ? "yearly" :
      interval === "month" ? "monthly" :
      (interval ? interval : "")

    const planName =
      bestForPlan.current_plan ||
      bestForPlan.subscription_tier ||
      bestForPlan.plan_name ||
      (billingFrequency ? `plan_${billingFrequency}_${intervalCount}` : "")

    // ARR is org-level and follows the same "Paid" bucket rule used elsewhere:
    // sum effective ARR across paid-active subscriptions only (active + first_payment_at).
    // If org has no paid-active subs, ARR is 0.
    const totalArr = hasPaidActive
      ? paidActiveRows.reduce((sum, row) => sum + ARR_computeEffectiveArrFromSubscriptionRow_(row, asOfDate), 0)
      : 0

    // Stripe email
    const stripeEmail = String(bestForPlan.customer_email || "").trim()

    // Discount fields (from raw_stripe_subscriptions)
    const discountPercent = bestForPlan.discount_percent
    const discountDuration = bestForPlan.discount_duration
    const discountDurationMonths = bestForPlan.discount_duration_months
    const discountStartAt =
      ARR_toIsoOrBlank_(bestForPlan.discount_start_at) ||
      ARR_toIsoOrBlank_(bestForPlan.first_payment_at) ||
      ARR_toIsoOrBlank_(bestForPlan.created_at)
    const discountEndAt = ARR_toIsoOrBlank_(bestForPlan.discount_end_at)

    // Current status: org-level lifecycle
    let currentStatus = "active"
    if (!hasActive) {
      const hasCanceled = rows.some(r => String(r.status || "").toLowerCase() === "canceled")
      if (hasCanceled) currentStatus = "canceled"
      else currentStatus = String(bestForPlan.status || "").trim()
    }

    out.set(orgId, {
      has_any: true,
      has_active: hasActive,

      stripe_email: stripeEmail,
      purchase_date: purchaseIso || "",
      churn_date: churnIso || "",

      plan_name: planName,
      billing_frequency: billingFrequency,
      total_arr: totalArr,

      discount_percent: discountPercent,
      discount_duration: discountDuration,
      discount_duration_months: discountDurationMonths,
      discount_start_at: discountStartAt,
      discount_end_at: discountEndAt,

      current_status: currentStatus,
    })
  })

  return out
}

function ARR_pickBestSubscriptionRow_(rows) {
  if (!rows || !rows.length) return {}

  // Prefer latest created_at, else first_payment_at
  const scored = rows.slice().sort((a, b) => {
    const aKey = ARR_toMs_(a.created_at) || ARR_toMs_(a.first_payment_at) || 0
    const bKey = ARR_toMs_(b.created_at) || ARR_toMs_(b.first_payment_at) || 0
    return bKey - aKey
  })

  return scored[0] || rows[0]
}

function ARR_computeEffectiveArrFromSubscriptionRow_(row, asOfDate) {
  const amount = ARR_num_(row && row.amount)
  const interval = ARR_str_(row && row.interval).toLowerCase()
  const intervalCount = Math.max(1, ARR_num_(row && row.interval_count) || 1)
  const amountYearly = ARR_num_(row && row.amount_yearly)

  const baseAmount = amount > 0 ? amount : 0
  const baseArr = (amountYearly > 0)
    ? amountYearly
    : ARR_computeAnnualizedAmount_(baseAmount, interval, intervalCount)

  if (baseArr <= 0) return 0

  const effectiveAmount = (baseAmount > 0)
    ? ARR_applyActiveDiscountsToAmount_(baseAmount, row, asOfDate)
    : baseAmount
  const effectiveArrFromAmount = (effectiveAmount > 0)
    ? ARR_computeAnnualizedAmount_(effectiveAmount, interval, intervalCount)
    : 0

  if (effectiveAmount > 0) return effectiveArrFromAmount

  // Fallback when amount field is unavailable but annualized amount exists:
  // apply percent discounts directly on ARR and amount_off via annualization.
  let arrOut = baseArr
  const details = ARR_activeDiscountsFromRow_(row, asOfDate)
  for (const d of details) {
    const pct = Number(d.percent_off)
    if (isFinite(pct) && pct > 0) {
      const bounded = Math.max(0, Math.min(100, pct))
      arrOut *= (1 - bounded / 100)
    }
    const amtOff = Number(d.amount_off)
    if (isFinite(amtOff) && amtOff > 0) {
      arrOut -= ARR_computeAnnualizedAmount_(amtOff, interval, intervalCount)
    }
  }
  return Math.max(0, arrOut)
}

function ARR_computeAnnualizedAmount_(amount, interval, intervalCount) {
  const amt = ARR_num_(amount)
  if (amt <= 0) return 0
  const months = ARR_intervalMonths_(interval, intervalCount)
  return months > 0 ? (amt * (12 / months)) : (amt * 12)
}

function ARR_intervalMonths_(interval, intervalCount) {
  const intv = ARR_str_(interval).toLowerCase().trim()
  const count = Math.max(1, ARR_num_(intervalCount) || 1)
  if (intv === "year" || intv === "annual" || intv === "yr") return 12 * count
  if (intv === "month" || intv === "mo") return count
  return 1
}

function ARR_applyActiveDiscountsToAmount_(baseAmount, row, asOfDate) {
  let out = ARR_num_(baseAmount)
  const active = ARR_activeDiscountsFromRow_(row, asOfDate)
  for (const d of active) {
    const pct = Number(d.percent_off)
    if (isFinite(pct) && pct > 0) {
      const bounded = Math.max(0, Math.min(100, pct))
      out *= (1 - bounded / 100)
    }
    const amountOff = Number(d.amount_off)
    if (isFinite(amountOff) && amountOff > 0) {
      out -= amountOff
    }
    if (out <= 0) return 0
  }
  return Math.max(0, out)
}

function ARR_activeDiscountsFromRow_(row, asOfDate) {
  const details = ARR_parseDiscountDetailsFromRow_(row)
  return details.filter(d => ARR_isDiscountActiveOnDate_(d, asOfDate))
}

function ARR_parseDiscountDetailsFromRow_(row) {
  const out = []
  if (!row) return out

  const jsonRaw = ARR_str_(row.discount_details_json)
  if (jsonRaw) {
    try {
      const arr = JSON.parse(jsonRaw)
      if (Array.isArray(arr)) {
        arr.forEach((d, i) => {
          if (!d || typeof d !== "object") return
          out.push({
            index: i + 1,
            percent_off: ARR_num_(d.percent_off),
            amount_off: ARR_num_(d.amount_off),
            duration: ARR_str_(d.duration).toLowerCase(),
            duration_in_months: ARR_num_(d.duration_in_months),
            start_at: ARR_toIsoOrBlank_(d.start_at),
            end_at: ARR_toIsoOrBlank_(d.end_at),
            promotion_code: ARR_str_(d.promotion_code),
          })
        })
      }
    } catch (e) {}
  }

  if (out.length) return out

  const pctAll = ARR_csvList_(row.discount_percent_all)
  const amtAll = ARR_csvList_(row.discount_amount_off_all)
  const durAll = ARR_csvList_(row.discount_duration_all)
  const durMonthsAll = ARR_csvList_(row.discount_duration_months_all)
  const startAll = ARR_csvList_(row.discount_start_at_all)
  const endAll = ARR_csvList_(row.discount_end_at_all)
  const promoAll = ARR_csvList_(row.promo_code_all)
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
      percent_off: ARR_num_(pctAll[i]),
      amount_off: ARR_num_(amtAll[i]),
      duration: ARR_str_(durAll[i]).toLowerCase(),
      duration_in_months: ARR_num_(durMonthsAll[i]),
      start_at: ARR_toIsoOrBlank_(startAll[i]),
      end_at: ARR_toIsoOrBlank_(endAll[i]),
      promotion_code: ARR_str_(promoAll[i]),
    })
  }

  if (out.length) return out

  out.push({
    index: 1,
    percent_off: ARR_num_(row.discount_percent),
    amount_off: ARR_num_(row.discount_amount_off),
    duration: ARR_str_(row.discount_duration).toLowerCase(),
    duration_in_months: ARR_num_(row.discount_duration_months),
    start_at: ARR_toIsoOrBlank_(row.discount_start_at || row.first_payment_at || row.created_at),
    end_at: ARR_toIsoOrBlank_(row.discount_end_at),
    promotion_code: ARR_str_(row.promo_code),
  })
  return out
}

function ARR_isDiscountActiveOnDate_(detail, asOfDate) {
  const asOf = (asOfDate instanceof Date && !isNaN(asOfDate.getTime()))
    ? asOfDate
    : new Date()
  const d = detail || {}
  const duration = ARR_str_(d.duration).toLowerCase()
  const start = ARR_parseIsoDate_(d.start_at)
  const explicitEnd = ARR_parseIsoDate_(d.end_at)

  if (start && asOf < start) return false
  if (explicitEnd) return asOf < explicitEnd

  if (duration === "forever") return true
  if (duration === "repeating" || duration === "once") {
    if (!start) return false
    const monthsRaw = ARR_num_(d.duration_in_months)
    const months = (isFinite(monthsRaw) && monthsRaw > 0) ? monthsRaw : 1
    const end = ARR_addMonths_(start, months)
    return asOf < end
  }
  return false
}

function ARR_csvList_(v) {
  const s = ARR_str_(v)
  if (!s) return []
  return s.split(",").map(x => ARR_str_(x))
}

/* ============================================================
 * Build indexes
 * ============================================================ */

function ARR_buildMembershipsByOrgId_(mems) {
  const out = new Map()
  ;(mems || []).forEach(m => {
    const orgId = ARR_str_(m.org_id)
    if (!orgId) return
    const email = ARR_str_(m.email)
    const emailKey = ARR_normEmail_(m.email_key || email)
    const role = ARR_str_(m.role).toLowerCase()
    const createdAt = ARR_toIsoOrBlank_(m.created_at)

    if (!out.has(orgId)) out.set(orgId, [])
    out.get(orgId).push({ email, email_key: emailKey, role, created_at: createdAt })
  })
  return out
}

function ARR_buildUsersByEmailKey_(users) {
  const out = new Map()
  ;(users || []).forEach(u => {
    const email = ARR_str_(u.email)
    const emailKey = ARR_normEmail_(u.email_key || email)
    if (!emailKey) return
    out.set(emailKey, u)
  })
  return out
}

function ARR_buildStripeBySubscriptionId_(subs) {
  const out = new Map()
  ;(subs || []).forEach(s => {
    const id = ARR_str_(s.stripe_subscription_id || s.subscription_id || s.id)
    if (!id) return
    out.set(id, s)
  })
  return out
}

function ARR_buildCanonByClerkOrgId_(canonRows) {
  const out = new Map()
  ;(canonRows || []).forEach(r => {
    const clerkOrgId = ARR_str_(r.clerk_org_id || r.org_id)
    if (!clerkOrgId || out.has(clerkOrgId)) return
    out.set(clerkOrgId, r)
  })
  return out
}

function ARR_rowsFromSubIds_(subIdSet, stripeBySubId) {
  const rows = []
  ;(subIdSet || new Set()).forEach(id => {
    const row = stripeBySubId.get(id)
    if (row) rows.push(row)
  })
  return rows
}

function ARR_buildSubIdsByOrgId_(membershipsByOrgId, userByEmailKey, users) {
  const out = new Map()

  membershipsByOrgId.forEach((members, orgId) => {
    const set = out.get(orgId) || new Set()
    ;(members || []).forEach(m => {
      const key = ARR_normEmail_(m.email_key || m.email)
      if (!key) return
      const u = userByEmailKey.get(key)
      const subId = u ? ARR_str_(u.stripe_subscription_id || u.stripeSubscriptionId) : ""
      if (subId) set.add(subId)
    })
    out.set(orgId, set)
  })

  ;(users || []).forEach(u => {
    const orgId = ARR_str_(u.org_id)
    const subId = ARR_str_(u.stripe_subscription_id || u.stripeSubscriptionId)
    if (!orgId || !subId) return
    const set = out.get(orgId) || new Set()
    set.add(subId)
    out.set(orgId, set)
  })

  return out
}

function ARR_buildInternalExcludeSubIdSet_(sheet) {
  const out = new Set()
  if (!sheet) return out

  const rows = ARR_readSheetObjects_(sheet, 1)
  ;(rows || []).forEach(r => {
    const reason = ARR_str_(r.exclude_reason).toLowerCase()
    if (reason !== "internal") return
    const subId =
      ARR_str_(r.subscription_id) ||
      ARR_str_(r.stripe_subscription_id) ||
      ARR_str_(r.subscription)
    if (subId) out.add(subId)
  })
  return out
}

/* ============================================================
 * Owner email selection
 * ============================================================ */

function ARR_pickOrgOwnerEmail_(members) {
  const arr = (members || []).slice()

  // Sort by created_at ascending (earliest)
  arr.sort((a, b) => {
    const ams = ARR_toMs_(a.created_at) || 0
    const bms = ARR_toMs_(b.created_at) || 0
    return ams - bms
  })

  const owners = arr.filter(m => (m.role || "").includes("owner"))
  if (owners.length && owners[0].email) return owners[0].email

  const admins = arr.filter(m => (m.role || "").includes("admin"))
  if (admins.length && admins[0].email) return admins[0].email

  // fallback any earliest
  if (arr.length && arr[0].email) return arr[0].email
  return ""
}

/* ============================================================
 * Sheet writing helpers (header-based)
 * ============================================================ */

function ARR_ensureHeaderRow_(sheet, headerRow, ensureHeaderName) {
  // Read existing header row
  const lastCol = Math.max(sheet.getLastColumn(), 1)
  let header = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim())

  // If sheet is empty or header row is empty, create a basic header set from row 2 currently in the sheet.
  // (We assume the user already has headers in row 2. We just ensure "status" exists.)
  const anyHeader = header.some(h => h)
  if (!anyHeader) {
    throw new Error(`arr_raw_data header row ${headerRow} is empty. Add your headers to row ${headerRow} first.`)
  }

  const ensureList = Array.isArray(ensureHeaderName)
    ? ensureHeaderName
    : (ensureHeaderName ? [ensureHeaderName] : [])

  if (ensureList.length) {
    ensureList.forEach(name => {
      if (!name) return
      if (!header.includes(name)) header.push(name)
    })
  }

  // Keep header width stable
  sheet.getRange(headerRow, 1, 1, header.length).setValues([header])

  return header
}

function ARR_headerMapFromRow_(headerRowArr) {
  const map = {}
  ;(headerRowArr || []).forEach((h, idx) => {
    const key = String(h || "").trim()
    if (!key) return
    map[key] = idx
  })
  return map
}

function ARR_rowFromHeader_(header, obj) {
  return header.map(h => {
    const key = ARR_key_(h)
    // obj is keyed by snake_case, but headers are snake_case already in your sheet.
    // If your sheet headers ever contain spaces, key_ will normalize.
    return Object.prototype.hasOwnProperty.call(obj, key) ? obj[key] : ""
  })
}

function ARR_key_(h) {
  return String(h || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "_")
}

function ARR_clearDataRegion_(sheet, startRow, numCols) {
  const maxRows = sheet.getMaxRows()
  const numRows = Math.max(0, maxRows - startRow + 1)
  if (!numRows) return
  sheet.getRange(startRow, 1, numRows, numCols).clearContent()
}

function ARR_applyArrRawFormats_(sheet, header, numRows) {
  if (!numRows) return
  const cohortHeaders = ["sign_up_cohort_month", "first_payment_cohort_month"]
  cohortHeaders.forEach(h => {
    const idx = header.findIndex(k => String(k || "").trim().toLowerCase() === h)
    if (idx < 0) return
    sheet.getRange(ARR_RAW_CFG.DATA_START_ROW, idx + 1, numRows, 1).setNumberFormat("mmm yyyy")
  })

  const dateHeaders = ["org_creation_date", "first_payment_date", "churn_date", "subscription_start_date"]
  dateHeaders.forEach(h => {
    const idx = header.findIndex(k => String(k || "").trim().toLowerCase() === h)
    if (idx < 0) return
    sheet.getRange(ARR_RAW_CFG.DATA_START_ROW, idx + 1, numRows, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss")
  })

  const numHeaders = ["total_arr"]
  numHeaders.forEach(h => {
    const idx = header.findIndex(k => String(k || "").trim().toLowerCase() === h)
    if (idx < 0) return
    sheet.getRange(ARR_RAW_CFG.DATA_START_ROW, idx + 1, numRows, 1).setNumberFormat("0")
  })
}

function ARR_batchSetValues_(sheet, startRow, startCol, values, chunkSize) {
  const size = chunkSize || 2000
  for (let i = 0; i < values.length; i += size) {
    const chunk = values.slice(i, i + size)
    sheet.getRange(startRow + i, startCol, chunk.length, chunk[0].length).setValues(chunk)
  }
}

/* ============================================================
 * Generic utils
 * ============================================================ */

function ARR_readSheetObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1) return []

  const header = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim())
  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()

  return data.map(r => {
    const obj = {}
    header.forEach((h, i) => {
      if (!h) return
      obj[ARR_key_(h)] = r[i]
    })
    return obj
  })
}

function ARR_getOrCreateSheet_(ss, name) {
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function ARR_lockWrap_(name, fn) {
  const lock = LockService.getScriptLock()
  if (!lock.tryLock(300000)) throw new Error(`Could not acquire lock: ${name}`)
  try {
    return fn()
  } finally {
    lock.releaseLock()
  }
}

function ARR_str_(v) {
  if (v === null || v === undefined) return ""
  return String(v).trim()
}

function ARR_normEmail_(v) {
  const s = String(v || "").trim().toLowerCase()
  if (!s) return ""
  return s.replace(/\+[^@]+(?=@)/, "")
}

function ARR_num_(v) {
  const n = Number(v)
  return isFinite(n) ? n : 0
}

function ARR_toIsoOrBlank_(v) {
  if (!v) return ""
  if (v instanceof Date) return v.toISOString()

  const s = String(v || "").trim()
  if (!s) return ""

  // ISO string already
  if (s.includes("T") && s.endsWith("Z")) return s

  // unix seconds/millis
  if (/^\d+$/.test(s)) {
    const n = Number(s)
    const ms = n > 1e12 ? n : n * 1000
    const d = new Date(ms)
    return isNaN(d.getTime()) ? "" : d.toISOString()
  }

  const d = new Date(s)
  return isNaN(d.getTime()) ? "" : d.toISOString()
}

function ARR_parseIsoDate_(iso) {
  const s = String(iso || "").trim()
  if (!s) return null
  const d = new Date(s)
  return isNaN(d.getTime()) ? null : d
}

function ARR_toMs_(iso) {
  const d = ARR_parseIsoDate_(iso)
  return d ? d.getTime() : 0
}

function ARR_isoToCohortMonth_(iso) {
  const d = ARR_parseIsoDate_(iso)
  if (!d) return ""
  const y = d.getUTCFullYear()
  const m = String(d.getUTCMonth() + 1).padStart(2, "0")
  return `${y}-${m}`
}

function ARR_isoToCohortMonthDate_(iso) {
  const d = ARR_parseIsoDate_(iso)
  if (!d) return ""
  return new Date(d.getFullYear(), d.getMonth(), 1)
}

function ARR_minIso_(isos) {
  let best = null
  ;(isos || []).forEach(s => {
    const d = ARR_parseIsoDate_(s)
    if (!d) return
    if (!best || d.getTime() < best.getTime()) best = d
  })
  return best ? best.toISOString() : ""
}

function ARR_maxIso_(isos) {
  let best = null
  ;(isos || []).forEach(s => {
    const d = ARR_parseIsoDate_(s)
    if (!d) return
    if (!best || d.getTime() > best.getTime()) best = d
  })
  return best ? best.toISOString() : ""
}

function ARR_firstNonEmpty_() {
  for (let i = 0; i < arguments.length; i++) {
    const v = ARR_str_(arguments[i])
    if (v) return v
  }
  return ""
}
