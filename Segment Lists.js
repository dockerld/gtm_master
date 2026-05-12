/**************************************************************
 * Segment Lists (Google Apps Script)
 *
 * Evaluates hardcoded segment criteria against canon_orgs +
 * org_subscription_info, writes a "crm_lists" sheet, then
 * syncs segment tags into Notion Companies as a "Segments"
 * multi-select property (additive-only / union).
 *
 * Functions:
 * 1. build_crm_lists()          — compute segments, write sheet
 * 2. sync_segments_to_notion()  — push segments into Notion
 **************************************************************/

/** =========================
 * CONFIG
 * ========================= */

const SEG_SHEET_NAME = "crm_lists"
const SEG_SHEET_CANON_ORGS = "canon_orgs"
const SEG_SHEET_ORG_SUB_INFO = "org_subscription_info"

// Notion property name (multi-select on Companies)
const SEG_NOTION_PROP = "Segments"

// "Trial Ending" thresholds
const SEG_TRIAL_ENDING_MIN_SEATS = 10
const SEG_TRIAL_ENDING_MAX_DAYS = 7

/** =========================
 * SEGMENT EVALUATORS
 * ========================= */

const SEGMENT_EVALUATORS = [
  { name: "Trial Ending", fn: segmentTrialEnding_ },
  { name: "Trial Ending This Month", fn: segmentTrialEndingThisMonth_ },
  { name: "Trial Ending Next Month", fn: segmentTrialEndingNextMonth_ },
  { name: "Top 25 Firms", fn: segmentTop25Firms_ }
]

/**
 * Bucket logic — mirrors ringBucketFromOrgSubscriptionInfo_() in Render The Ring.js
 * Returns 'paid', 'intent_to_pay', 'free_trial', or '' (unknown)
 */
function SEG_bucket_(subInfo) {
  if (!subInfo) return ""
  const status = SEG_str_(subInfo.status).toLowerCase()
  const hasFirstPayment = !!SEG_str_(subInfo.first_payment_at)
  const hasPaymentMethod = SEG_str_(subInfo.has_payment_method) === "true"
    || SEG_str_(subInfo.has_payment_method) === "TRUE"
    || subInfo.has_payment_method === true

  if (status === "active" && hasFirstPayment) return "paid"
  if (status === "active" && !hasFirstPayment && hasPaymentMethod) return "intent_to_pay"
  if (status === "trialing" && hasPaymentMethod) return "intent_to_pay"
  if (status === "active" && !hasFirstPayment && !hasPaymentMethod) return "free_trial"
  if (status === "trialing" && !hasPaymentMethod) return "free_trial"
  return ""
}

/**
 * Trial Ending: orgs in free_trial or intent_to_pay bucket,
 * with 10+ seats, whose trial ends within 7 days.
 */
function segmentTrialEnding_(org, subInfo) {
  if (!subInfo) return false

  // Must be in a trial bucket (free_trial or intent_to_pay)
  const bucket = SEG_bucket_(subInfo)
  if (bucket !== "free_trial" && bucket !== "intent_to_pay") return false

  // Must be a meaningful org (10+ seats) — from org_subscription_info
  const seats = Math.max(
    Number(subInfo.quantity_total) || 0,
    Number(org.active_subscription_seats) || 0
  )
  if (seats < SEG_TRIAL_ENDING_MIN_SEATS) return false

  // Trial must end within N days
  const daysRemaining = Number(subInfo.trial_days_remaining)
  if (isNaN(daysRemaining) || daysRemaining < 0 || daysRemaining > SEG_TRIAL_ENDING_MAX_DAYS) return false

  return true
}

/**
 * Trial Ending This Month: orgs with a trial_ends_at in canon_orgs
 * that falls within the current calendar month.
 */
function segmentTrialEndingThisMonth_(org, subInfo) {
  const trialEndsAt = SEG_str_(org.trial_ends_at)
  if (!trialEndsAt) return false

  const trialEnd = new Date(trialEndsAt)
  if (isNaN(trialEnd.getTime())) return false

  const now = new Date()
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1)
  const monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59)

  return trialEnd >= monthStart && trialEnd <= monthEnd
}

/**
 * Trial Ending Next Month: orgs with a trial_ends_at in canon_orgs
 * that falls within next calendar month.
 */
/**
 * Top 25 Firms: paid orgs ranked by seats, top 25 win.
 * Precomputed set is populated in build_crm_lists before evaluation.
 */
let _top25OrgIds = new Set()

function precomputeTop25Firms_(orgs) {
  const paid = orgs
    .filter(o => SEG_str_(o.org_status) === "Paid")
    .map(o => ({ org_id: SEG_str_(o.org_id), seats: Number(o.seats) || 0 }))
    .sort((a, b) => b.seats - a.seats)
    .slice(0, 25)
  _top25OrgIds = new Set(paid.map(o => o.org_id))
}

function segmentTop25Firms_(org, subInfo) {
  return _top25OrgIds.has(SEG_str_(org.org_id))
}

function segmentTrialEndingNextMonth_(org, subInfo) {
  const trialEndsAt = SEG_str_(org.trial_ends_at)
  if (!trialEndsAt) return false

  const trialEnd = new Date(trialEndsAt)
  if (isNaN(trialEnd.getTime())) return false

  const now = new Date()
  const monthStart = new Date(now.getFullYear(), now.getMonth() + 1, 1)
  const monthEnd = new Date(now.getFullYear(), now.getMonth() + 2, 0, 23, 59, 59)

  return trialEnd >= monthStart && trialEnd <= monthEnd
}

/** =========================
 * BUILD CRM LISTS SHEET
 * ========================= */

function build_crm_lists() {
  const ss = SpreadsheetApp.getActive()

  // Read inputs
  const canonOrgsSh = ss.getSheetByName(SEG_SHEET_CANON_ORGS)
  if (!canonOrgsSh) throw new Error(`Missing sheet: ${SEG_SHEET_CANON_ORGS}`)
  const orgSubInfoSh = ss.getSheetByName(SEG_SHEET_ORG_SUB_INFO)
  if (!orgSubInfoSh) throw new Error(`Missing sheet: ${SEG_SHEET_ORG_SUB_INFO}`)

  const orgs = SEG_readSheetObjects_(canonOrgsSh, 1)
  const subInfoRows = SEG_readSheetObjects_(orgSubInfoSh, 1)

  // Index org_subscription_info by app_org_id (its primary key column)
  const subInfoByAppOrgId = new Map()
  for (const row of subInfoRows) {
    const id = SEG_str_(row.app_org_id)
    if (id) subInfoByAppOrgId.set(id, row)
  }

  // Precompute segments that need full-list context
  precomputeTop25Firms_(orgs)

  // Evaluate segments for each org
  const headers = ["org_id", "org_name", "segments"]
  const rowsOut = []

  for (const org of orgs) {
    const orgId = SEG_str_(org.org_id)
    if (!orgId) continue

    // Join to org_subscription_info via app_org_id or org_id
    const appOrgId = SEG_str_(org.app_org_id)
    const subInfo = (appOrgId && subInfoByAppOrgId.get(appOrgId))
      || subInfoByAppOrgId.get(orgId)
      || null
    const matched = []

    for (const evaluator of SEGMENT_EVALUATORS) {
      try {
        if (evaluator.fn(org, subInfo)) matched.push(evaluator.name)
      } catch (e) {
        // Non-fatal: skip this evaluator for this org
      }
    }

    if (matched.length > 0) {
      rowsOut.push([orgId, SEG_str_(org.org_name), matched.join(", ")])
    }
  }

  // Write crm_lists sheet (overwrite)
  const shOut = SEG_getOrCreateSheet_(ss, SEG_SHEET_NAME)
  shOut.clear()
  shOut.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#F3F4F6")

  if (rowsOut.length) {
    shOut.getRange(2, 1, rowsOut.length, headers.length).setValues(rowsOut)
  }

  Logger.log(`build_crm_lists: ${rowsOut.length} orgs matched segments out of ${orgs.length} total`)
  return { matched: rowsOut.length, total: orgs.length }
}

/** =========================
 * SYNC SEGMENTS TO NOTION
 * ========================= */

function sync_segments_to_notion() {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(SEG_SHEET_NAME)
  if (!sh || sh.getLastRow() < 2) {
    Logger.log("sync_segments_to_notion: no crm_lists data, skipping")
    return { synced: 0 }
  }

  const rows = SEG_readSheetObjects_(sh, 1)

  // Build Map<org_id, string[]> of segment names
  const segmentsByOrgId = new Map()
  for (const row of rows) {
    const orgId = SEG_str_(row.org_id)
    const segments = SEG_str_(row.segments)
    if (!orgId || !segments) continue
    segmentsByOrgId.set(orgId, segments.split(",").map(s => s.trim()).filter(Boolean))
  }

  if (segmentsByOrgId.size === 0) {
    Logger.log("sync_segments_to_notion: no segments to sync")
    return { synced: 0 }
  }

  // Query all linked Notion companies
  const notion = notionClient_()
  const props = PropertiesService.getScriptProperties()
  const companiesDbId = props.getProperty(PROP_NOTION_COMPANIES_DB_ID)
  if (!companiesDbId) throw new Error("Missing Script Property: NOTION_COMPANIES_DB_ID")

  const allCompanies = notionQueryAll_(notion, companiesDbId, {
    filter: { property: NOTION_COMPANY_PROP_LINKED, checkbox: { equals: true } },
    page_size: 100
  })

  let synced = 0
  let skipped = 0

  for (const company of allCompanies) {
    const orgId = notionGetRichText_(company, NOTION_COMPANY_PROP_SAURON_ORG_ID)
    if (!orgId) continue

    const newSegments = segmentsByOrgId.get(orgId)
    if (!newSegments || newSegments.length === 0) continue

    // Read existing multi-select values
    const existing = SEG_notionGetMultiSelect_(company, SEG_NOTION_PROP)
    const existingSet = new Set(existing)

    // Union: add new, keep existing
    const union = new Set([...existingSet, ...newSegments])

    // Only patch if there are genuinely new tags
    if (union.size <= existingSet.size) {
      skipped += 1
      continue
    }

    const patch = { properties: {} }
    patch.properties[SEG_NOTION_PROP] = {
      multi_select: Array.from(union).map(name => ({ name }))
    }

    notionPatch_(notion, `/pages/${company.id}`, patch)
    synced += 1
  }

  Logger.log(`sync_segments_to_notion: synced=${synced}, skipped=${skipped}, total_companies=${allCompanies.length}`)
  return { synced, skipped }
}

/** =========================
 * NOTION HELPER
 * ========================= */

function SEG_notionGetMultiSelect_(page, propName) {
  const p = page && page.properties ? page.properties[propName] : null
  if (!p || p.type !== "multi_select" || !Array.isArray(p.multi_select)) return []
  return p.multi_select.map(opt => opt.name).filter(Boolean)
}

/** =========================
 * SHEET HELPERS (namespaced)
 * ========================= */

function SEG_readSheetObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1) return []

  const header = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim())
  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()

  return data.map(r => {
    const obj = {}
    header.forEach((h, i) => { if (h) obj[h] = r[i] })
    return obj
  })
}

function SEG_getOrCreateSheet_(ss, name) {
  if (typeof getOrCreateSheet === "function") {
    try { return getOrCreateSheet(ss, name) } catch (e) {}
    try { return getOrCreateSheet(name) } catch (e) {}
  }
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function SEG_str_(v) {
  if (v === null || v === undefined) return ""
  return String(v).trim()
}
