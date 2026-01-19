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
    CLERK_ORGS: "raw_clerk_orgs",
    CLERK_MEMBERSHIPS: "raw_clerk_memberships",
    CLERK_USERS: "raw_clerk_users",
    STRIPE_SUBS: "raw_stripe_subscriptions",
  },

  // Trial standard length
  TRIAL_DAYS: 14,

  // Adds a validation status column if missing
  VALIDATION_STATUS_HEADER: "status",
  SUBSCRIPTION_START_HEADER: "subscription_start_date",
}

function render_arr_raw_data_view() {
  return ARR_lockWrap_("render_arr_raw_data_view", () => {
    const t0 = new Date()
    const ss = SpreadsheetApp.getActive()

    const shOut = ARR_getOrCreateSheet_(ss, ARR_RAW_CFG.SHEET_NAME)

    const shOrgs = ss.getSheetByName(ARR_RAW_CFG.INPUTS.CLERK_ORGS)
    const shMems = ss.getSheetByName(ARR_RAW_CFG.INPUTS.CLERK_MEMBERSHIPS)
    const shUsers = ss.getSheetByName(ARR_RAW_CFG.INPUTS.CLERK_USERS)
    const shStripe = ss.getSheetByName(ARR_RAW_CFG.INPUTS.STRIPE_SUBS)

    if (!shOrgs) throw new Error(`Missing sheet: ${ARR_RAW_CFG.INPUTS.CLERK_ORGS}`)
    if (!shMems) throw new Error(`Missing sheet: ${ARR_RAW_CFG.INPUTS.CLERK_MEMBERSHIPS}`)
    if (!shUsers) throw new Error(`Missing sheet: ${ARR_RAW_CFG.INPUTS.CLERK_USERS}`)
    if (!shStripe) throw new Error(`Missing sheet: ${ARR_RAW_CFG.INPUTS.STRIPE_SUBS}`)

    // Ensure headers exist; also ensure "status" column exists
    const header = ARR_ensureHeaderRow_(
      shOut,
      ARR_RAW_CFG.HEADER_ROW,
      [ARR_RAW_CFG.VALIDATION_STATUS_HEADER, ARR_RAW_CFG.SUBSCRIPTION_START_HEADER]
    )
    const headerMap = ARR_headerMapFromRow_(header)

    // Read inputs
    const orgs = ARR_readSheetObjects_(shOrgs, 1)
    const mems = ARR_readSheetObjects_(shMems, 1)
    const users = ARR_readSheetObjects_(shUsers, 1)
    const subs = ARR_readSheetObjects_(shStripe, 1)

    // Build indexes
    const membershipsByOrgId = ARR_buildMembershipsByOrgId_(mems) // orgId -> [{email,email_key,role,created_at}]
    const stripeBySubId = ARR_buildStripeBySubscriptionId_(subs)  // subId -> stripe row obj
    const userByEmailKey = ARR_buildUsersByEmailKey_(users)       // email_key -> user obj (incl stripe_subscription_id)

    // orgId -> Set(subIds) from memberships -> users -> stripe_subscription_id
    const subIdsByOrgId = ARR_buildSubIdsByOrgId_(membershipsByOrgId, userByEmailKey)

    // orgId -> derived subscription rollup (earliest purchase, current status, etc.)
    const subRollupByOrgId = ARR_buildOrgSubscriptionRollup_(subIdsByOrgId, stripeBySubId)

    // Now build output rows in the existing column order
    const outRows = []
    const nowIso = new Date().toISOString()

    // Sort for stability
    const sortedOrgs = orgs
      .map(o => ({
        org_id: ARR_str_(o.org_id),
        org_name: ARR_str_(o.org_name) || ARR_str_(o.org_slug),
        org_created_at: ARR_str_(o.created_at || o.org_created_at),
      }))
      .filter(o => o.org_id)
      .sort((a, b) => (a.org_name || "").localeCompare(b.org_name || "") || a.org_id.localeCompare(b.org_id))

    for (const o of sortedOrgs) {
      const orgId = o.org_id
      const orgName = o.org_name
      const orgCreationIso = ARR_toIsoOrBlank_(o.org_created_at)

      // org owner email: earliest owner membership (fallback admin/member)
      const ownerEmail = ARR_pickOrgOwnerEmail_(membershipsByOrgId.get(orgId) || [])

      // subscription rollup for org
      const roll = subRollupByOrgId.get(orgId) || ARR_emptyRollup_()

      // Stripe email (customer email)
      const stripeEmail = roll.stripe_email || ""

      // trial start/end from Clerk private metadata in raw_clerk_users (org owner is best guess for org trial metadata)
      // If you later store org trial meta on orgs instead, swap the source.
      const ownerKey = ARR_normEmail_(ownerEmail)
      const ownerUser = ownerKey ? (userByEmailKey.get(ownerKey) || null) : null

      const trialStart = ARR_toIsoOrBlank_(
        ARR_firstNonEmpty_(
          ownerUser && ownerUser.trial_start_date,
          ownerUser && ownerUser.trialStartDate,
          ownerUser && ownerUser.trial_start,
          ownerUser && ownerUser.trial_start_at
        )
      )

      // Compute trial_end_date with your two-path logic
      const subRows = ARR_rowsFromSubIds_(subIdsByOrgId.get(orgId), stripeBySubId)
      const trialEnd = ARR_computeTrialEndIso_({
        trialStartIso: trialStart,
        subscriptions: subRows,
        trialDays: ARR_RAW_CFG.TRIAL_DAYS
      })

      const subscriptionStart = ARR_minIso_(
        subRows.map(r => ARR_toIsoOrBlank_(r.created_at)).filter(Boolean)
      )

      // Cohorts
      const trialCohortMonth = ARR_isoToCohortMonth_(orgCreationIso)
      const paidCohortMonth = ARR_isoToCohortMonth_(roll.purchase_date || "")

      // Current status (lifecycle; keep your existing column name "current_status")
      const currentStatus = roll.current_status || (roll.has_active ? "active" : (roll.has_any ? "inactive" : ""))

      // Plan name / billing freq from Stripe
      const planName = roll.plan_name || ""
      const billingFrequency = roll.billing_frequency || ""

      // ARR
      const totalArr = roll.total_arr || 0
      const meetingAssArr = totalArr // your “same numbers” rule

      // Churn
      const churnDate = roll.churn_date || ""

      // ✅ Validation status column
      const validationStatus =
        (Number(roll.discount_percent) === 100 && String(roll.discount_duration || "").toLowerCase() === "forever")
          ? "Invalid"
          : ""

      // Map to your output headers by name (so column order can evolve safely)
      const rowObj = {
        org_id: orgId,
        org_name: orgName,
        org_email: ownerEmail,
        stripe_email: stripeEmail,
        org_creation_date: orgCreationIso,

        created_at: nowIso,
        last_updated_at: nowIso,

        trial_start_date: trialStart,
        trial_end_date: trialEnd,
        subscription_start_date: subscriptionStart,

        purchase_date: roll.purchase_date || "",
        churn_date: churnDate,

        trial_cohort_month: trialCohortMonth,
        paid_cohort_month: paidCohortMonth,

        current_status: currentStatus,
        plan_name: planName,
        billing_frequency: billingFrequency,

        acquisition_channel: "",

        total_arr: totalArr,
        meeting_ass_arr: meetingAssArr,
        product_2_arr: 0,
        product_3_arr: 0,

        notes: "",
        status: validationStatus, // ✅ NEW column
      }

      outRows.push(ARR_rowFromHeader_(header, rowObj))
    }

    // Write output (clear only data region, keep header row 2 intact)
    ARR_clearDataRegion_(shOut, ARR_RAW_CFG.DATA_START_ROW, header.length)
    if (outRows.length) {
      ARR_batchSetValues_(shOut, ARR_RAW_CFG.DATA_START_ROW, 1, outRows, 2000)
    }

    shOut.setFrozenRows(ARR_RAW_CFG.HEADER_ROW)
    shOut.autoResizeColumns(1, header.length)

    const seconds = (new Date() - t0) / 1000
    if (typeof writeSyncLog === "function") {
      writeSyncLog("render_arr_raw_data_view", "ok", sortedOrgs.length, outRows.length, seconds, "")
    } else {
      Logger.log(`[render_arr_raw_data_view] ok rows_in=${sortedOrgs.length} rows_out=${outRows.length} seconds=${seconds}`)
    }

    return { rows_in: sortedOrgs.length, rows_out: outRows.length }
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

    current_status: "",
  }
}

function ARR_buildOrgSubscriptionRollup_(subIdsByOrgId, stripeBySubId) {
  const out = new Map()

  subIdsByOrgId.forEach((subIdSet, orgId) => {
    const ids = Array.from(subIdSet || []).filter(Boolean)
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

    // Determine active vs any
    const activeRows = rows.filter(r => String(r.status || "").toLowerCase() === "active")
    const hasActive = activeRows.length > 0

    // Purchase date = earliest first_payment_at across all subs (do NOT overwrite later)
    const purchaseIso = ARR_minIso_(rows.map(r => r.first_payment_at).filter(Boolean))

    // Churn date = if no active subs, take max(current_period_end) or canceled_at among rows
    let churnIso = ""
    if (!hasActive) {
      churnIso = ARR_maxIso_(
        rows
          .map(r => r.current_period_end || r.canceled_at)
          .filter(Boolean)
      )
    }

    // Pick the “current” subscription to derive plan/billing/arr:
    // - prefer active with latest current_period_start; else latest current_period_start among all
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

    // ARR: prefer amount_yearly if present, else compute from amount+interval
    let totalArr = 0
    if (bestForPlan.amount_yearly != null && bestForPlan.amount_yearly !== "") {
      totalArr = Number(bestForPlan.amount_yearly) || 0
    } else {
      const amt = Number(bestForPlan.amount) || 0
      totalArr = (interval === "year") ? amt : (amt * 12)
    }

    // Stripe email
    const stripeEmail = String(bestForPlan.customer_email || "").trim()

    // Discount fields (from raw_stripe_subscriptions)
    const discountPercent = bestForPlan.discount_percent
    const discountDuration = bestForPlan.discount_duration
    const discountDurationMonths = bestForPlan.discount_duration_months

    // Current status: "active" if any active else bestForPlan.status
    const currentStatus = hasActive ? "active" : String(bestForPlan.status || "").trim()

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

      current_status: currentStatus,
    })
  })

  return out
}

function ARR_pickBestSubscriptionRow_(rows) {
  if (!rows || !rows.length) return {}

  // Prefer latest current_period_start, else created_at
  const scored = rows.slice().sort((a, b) => {
    const aKey = ARR_toMs_(a.current_period_start) || ARR_toMs_(a.created_at) || 0
    const bKey = ARR_toMs_(b.current_period_start) || ARR_toMs_(b.created_at) || 0
    return bKey - aKey
  })

  return scored[0] || rows[0]
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

function ARR_rowsFromSubIds_(subIdSet, stripeBySubId) {
  const rows = []
  ;(subIdSet || new Set()).forEach(id => {
    const row = stripeBySubId.get(id)
    if (row) rows.push(row)
  })
  return rows
}

function ARR_buildSubIdsByOrgId_(membershipsByOrgId, userByEmailKey) {
  const out = new Map()

  membershipsByOrgId.forEach((members, orgId) => {
    const set = new Set()
    ;(members || []).forEach(m => {
      const key = ARR_normEmail_(m.email_key || m.email)
      if (!key) return
      const u = userByEmailKey.get(key)
      const subId = u ? ARR_str_(u.stripe_subscription_id || u.stripeSubscriptionId) : ""
      if (subId) set.add(subId)
    })
    out.set(orgId, set)
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
