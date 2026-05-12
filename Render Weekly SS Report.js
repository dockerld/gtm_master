/**************************************************************
 * render_weekly_ss_report()
 *
 * Weekly all-hands metrics sheet for Self-Serve (SS) orgs.
 * SS = everyone EXCEPT orgs flagged `managed` in Manual Stripe Changes
 * (Enterprise/managed customers). Also excludes the standard internal/
 * partner/free-subscription filter used elsewhere, plus pingassistant.com
 * owner emails and ping/test org names.
 *
 * Window: Last completed Mon-Sun (based on run date).
 *
 * Sections:
 *   1) Current State — ARR, New ARR this week, Churn $/orgs/seats/list
 *   2) Pipeline     — Trialing $, Intent to Pay $, Top 5 by seats,
 *                     Signup→Paid %, Intent→Paid %, Weighted Pipeline $
 *   3) Paying Base  — Top 5 by seats, Upsell list (manual),
 *                     Expansion seats (manual)
 **************************************************************/

const WEEKLY_SS_CFG = {
  SHEET_NAME: 'Weekly SS Report',
  MANAGED_REASON: 'managed',

  INPUTS: {
    CANON_ORGS: 'canon_orgs',
    ARR_RAW_DATA: 'arr_raw_data',
    ORG_SUBSCRIPTION_INFO: 'org_subscription_info',
    MANUAL_CHANGES: 'Manual Stripe Changes'
  },

  FMT: {
    MONEY: '$#,##0',
    INT: '#,##0',
    PCT: '0.0%'
  }
}

function render_weekly_ss_report() {
  return WSS_lockWrap_('render_weekly_ss_report', () => {
    const t0 = new Date()
    const ss = SpreadsheetApp.getActive()

    const windowRange = WSS_lastCompletedWeek_(new Date())

    const shCanon = ss.getSheetByName(WEEKLY_SS_CFG.INPUTS.CANON_ORGS)
    if (!shCanon) throw new Error(`Missing sheet: ${WEEKLY_SS_CFG.INPUTS.CANON_ORGS}`)
    const canonOrgs = WSS_readSheetObjects_(shCanon, 1)

    const shArrRaw = ss.getSheetByName(WEEKLY_SS_CFG.INPUTS.ARR_RAW_DATA)
    const arrRaw = shArrRaw ? WSS_readSheetObjects_(shArrRaw, 2) : []

    const managed = WSS_buildManagedExclusionSets_(ss)

    const ssOrgs = canonOrgs.filter(o => {
      const orgName = WSS_str_(o.org_name || o.org_slug).toLowerCase()
      if (orgName.includes('ping') || orgName.includes('test')) return false
      const ownerEmail = WSS_str_(o.owner_email).toLowerCase()
      if (ownerEmail.endsWith('@pingassistant.com')) return false
      const subIds = WSS_csvList_(o.stripe_subscription_ids)
      if (subIds.some(id => managed.subIds.has(id))) return false
      const custIds = WSS_csvList_(o.stripe_customer_ids)
      if (custIds.some(id => managed.customerIds.has(id))) return false
      return true
    })

    const arrByOrgId = new Map()
    for (const r of arrRaw) {
      const subId = WSS_str_(r.subscription_id || r.stripe_subscription_id || r.latest_subscription_id)
      if (subId && managed.subIds.has(subId)) continue
      const orgId = WSS_str_(r.org_id)
      if (!orgId) continue
      if (!arrByOrgId.has(orgId)) arrByOrgId.set(orgId, [])
      arrByOrgId.get(orgId).push(r)
    }

    const convRates = WSS_readConversionResolvedRates_(ss)
    convRates.intentToPaidRate = WSS_readIntentToPaidRate_(ss, managed)
    const bigThree = WSS_readBigThreeMetrics_(ss)
    const churnSummary = WSS_loadChurnsForWindow_(ss, windowRange, managed)
    const newPayingOrgs = WSS_loadNewPaymentsForWindow_(ss, windowRange, managed)
    const report = WSS_computeReport_(ssOrgs, arrByOrgId, windowRange, convRates, churnSummary, newPayingOrgs)

    const outSheet = WSS_getOrCreateSheet_(ss, WEEKLY_SS_CFG.SHEET_NAME)
    WSS_writeSheet_(outSheet, report, windowRange, convRates, bigThree)

    if (typeof writeSyncLog === 'function') {
      writeSyncLog('render_weekly_ss_report', 'ok', ssOrgs.length, 0,
        (new Date() - t0) / 1000, '')
    }

    return { rows_in: ssOrgs.length, rows_out: 0 }
  })
}

/* =========================
 * Compute
 * ========================= */

function WSS_computeReport_(ssOrgs, arrByOrgId, windowRange, convRates, churnSummary, newPayingOrgs) {
  const cs = churnSummary || { churnDollars: 0, churnOrgCount: 0, churnSeatCount: 0, churnOrgs: [] }
  const r = {
    currentArr: 0,
    newArrThisWeek: 0,
    churnDollars: cs.churnDollars,
    churnOrgCount: cs.churnOrgCount,
    churnSeatCount: cs.churnSeatCount,
    churnOrgs: cs.churnOrgs.slice(),
    newPayingOrgs: (newPayingOrgs || []).slice(),
    trialingDollars: 0,
    intentToPayDollars: 0,
    pipelineTopSeats: [],
    topPayingBySeats: []
  }

  const paidOrgs = []
  const pipelineOrgs = []

  for (const o of ssOrgs) {
    const status = WSS_str_(o.org_status)
    const arr = WSS_num_(o.active_subscription_arr_discounted)
    const seats = WSS_num_(o.active_subscription_seats) ||
                  WSS_num_(o.stripe_seats_paying_sum) ||
                  WSS_num_(o.stripe_seats_max)
    const name = WSS_str_(o.org_name) ||
                 WSS_str_(o.owner_email) ||
                 WSS_str_(o.app_org_id) ||
                 WSS_str_(o.org_id)
    const orgId = WSS_str_(o.app_org_id) || WSS_str_(o.org_id)

    if (status === 'Paid') {
      r.currentArr += arr
      paidOrgs.push({ name, seats, arr })

      const arrRows = arrByOrgId.get(orgId) || []
      const fpInWeek = arrRows.some(row => {
        const fp = WSS_parseDate_(row.first_payment_date)
        return fp && WSS_isInRange_(fp, windowRange.startMs, windowRange.endMs)
      })
      if (fpInWeek) r.newArrThisWeek += arr
    } else if (status === 'Trialing') {
      r.trialingDollars += arr
      pipelineOrgs.push({ name, seats, status: 'Trialing' })
    } else if (status === 'Intent to Pay') {
      r.intentToPayDollars += arr
      pipelineOrgs.push({ name, seats, status: 'Intent to Pay' })
    }
  }

  r.pipelineTopSeats = pipelineOrgs.slice().sort((a, b) => b.seats - a.seats).slice(0, 5)
  r.topPayingBySeats = paidOrgs.slice().sort((a, b) => b.seats - a.seats).slice(0, 5)
  r.churnOrgs.sort((a, b) => b.seats - a.seats)

  const trialRate = WSS_num_(convRates.signupToPaidRate)
  const intentRate = WSS_num_(convRates.intentToPaidRate)
  r.weightedPipeline = r.trialingDollars * trialRate + r.intentToPayDollars * intentRate

  return r
}

/* =========================
 * Churn (sourced from org_subscription_info)
 *
 * arr_raw_data skips orgs without an active ring bucket, which excludes
 * almost all churned orgs. org_subscription_info has churn_date and ARR
 * for every org including expired ones — the same source the All Stats
 * Churned Orgs Audit uses.
 * ========================= */

function WSS_loadNewPaymentsForWindow_(ss, windowRange, managed) {
  const out = []
  const sh = ss.getSheetByName(WEEKLY_SS_CFG.INPUTS.ORG_SUBSCRIPTION_INFO)
  if (!sh) return out
  const rows = WSS_readSheetObjects_(sh, 1)

  for (const r of rows) {
    const subId = WSS_str_(r.stripe_subscription_id)
    if (subId && managed.subIds.has(subId)) continue
    const custId = WSS_str_(r.stripe_customer_id)
    if (custId && managed.customerIds.has(custId)) continue

    const orgName = WSS_str_(r.org_name).toLowerCase()
    if (orgName.includes('ping') || orgName.includes('test')) continue
    const ownerEmail = WSS_str_(r.customer_email).toLowerCase()
    if (ownerEmail.endsWith('@pingassistant.com')) continue

    const fp = WSS_parseDate_(r.first_payment_at)
    if (!fp || !WSS_isInRange_(fp, windowRange.startMs, windowRange.endMs)) continue

    const name = WSS_str_(r.org_name) || WSS_str_(r.customer_email) || WSS_str_(r.app_org_id)
    const email = WSS_str_(r.customer_email)
    out.push({ name, email, subId })
  }

  return out
}

function WSS_loadChurnsForWindow_(ss, windowRange, managed) {
  const out = { churnDollars: 0, churnOrgCount: 0, churnSeatCount: 0, churnOrgs: [] }
  const sh = ss.getSheetByName(WEEKLY_SS_CFG.INPUTS.ORG_SUBSCRIPTION_INFO)
  if (!sh) return out
  const rows = WSS_readSheetObjects_(sh, 1)

  for (const r of rows) {
    const subId = WSS_str_(r.stripe_subscription_id)
    if (subId && managed.subIds.has(subId)) continue
    const custId = WSS_str_(r.stripe_customer_id)
    if (custId && managed.customerIds.has(custId)) continue

    const orgName = WSS_str_(r.org_name).toLowerCase()
    if (orgName.includes('ping') || orgName.includes('test')) continue
    const ownerEmail = WSS_str_(r.customer_email).toLowerCase()
    if (ownerEmail.endsWith('@pingassistant.com')) continue

    const cd = WSS_parseDate_(r.churn_date)
    if (!cd || !WSS_isInRange_(cd, windowRange.startMs, windowRange.endMs)) continue

    // Must have actually converted to paid — no first_payment_at means they never paid
    if (!WSS_parseDate_(r.first_payment_at)) continue

    const arr = WSS_num_(r.amount_after_discounts_yearly) || WSS_num_(r.amount_yearly)
    const seats = WSS_num_(r.quantity_total)
    const name = WSS_str_(r.org_name) || WSS_str_(r.customer_email) || WSS_str_(r.app_org_id)

    out.churnDollars += arr
    out.churnOrgCount += 1
    out.churnSeatCount += seats
    out.churnOrgs.push({ name, seats, arr, subId })
  }

  return out
}

/* =========================
 * Resolved conversion rates
 *
 * Signup → Paid:   paid / (signed_up − currently_trialing)
 *                  Matches All stats (excludes orgs whose trial hasn't resolved).
 * Intent → Paid:   paid / (ever_had_sub_and_resolved)
 *                  Resolved = paid OR (ever had a sub AND currently expired).
 *                  Currently trialing or currently intent-to-pay are unresolved.
 * ========================= */

function WSS_readConversionResolvedRates_(ss) {
  const out = { signupToPaidRate: 0, intentToPaidRate: 0 }
  const convOrgs = buildConversionOrgList_(ss)
  if (!convOrgs.length) return out

  let totalSignedUp = 0
  let paid = 0
  let currentlyTrialing = 0
  let intentResolved = 0

  for (const org of convOrgs) {
    totalSignedUp += 1
    if (org.first_payment_at) paid += 1
    if (org.is_currently_trialing) currentlyTrialing += 1

    const status = WSS_str_(org.org_status).toLowerCase()
    if (org.first_payment_at) {
      intentResolved += 1
    } else if (org.ever_had_sub && status === 'expired') {
      intentResolved += 1
    }
  }

  const signupResolved = totalSignedUp - currentlyTrialing
  out.signupToPaidRate = signupResolved ? (paid / signupResolved) : 0
  out.intentToPaidRate = intentResolved ? (paid / intentResolved) : 0
  return out
}

/* =========================
 * Intent → Paid rate (from org_subscription_info)
 *
 * Predicts probability that a "real intent" org converts to paid.
 *
 * Denominator: orgs with payment_method_created_at AND either
 *   - first_payment_at landed > 3 days after payment method creation, OR
 *   - no first_payment_at yet
 *   (excludes immediate-pay orgs that paid within 3 days — they aren't
 *    really "intent" pipeline, they're effectively instant paid signups)
 *
 * Numerator: orgs in that denominator who actually paid.
 * ========================= */

function WSS_readIntentToPaidRate_(ss, managed) {
  const sh = ss.getSheetByName(WEEKLY_SS_CFG.INPUTS.ORG_SUBSCRIPTION_INFO)
  if (!sh) return 0
  const rows = WSS_readSheetObjects_(sh, 1)
  const THREE_DAYS_MS = 3 * 24 * 60 * 60 * 1000

  let denominator = 0
  let numerator = 0

  for (const r of rows) {
    const subId = WSS_str_(r.stripe_subscription_id)
    if (subId && managed.subIds.has(subId)) continue
    const custId = WSS_str_(r.stripe_customer_id)
    if (custId && managed.customerIds.has(custId)) continue

    const orgName = WSS_str_(r.org_name).toLowerCase()
    if (orgName.includes('ping') || orgName.includes('test')) continue
    const ownerEmail = WSS_str_(r.customer_email).toLowerCase()
    if (ownerEmail.endsWith('@pingassistant.com')) continue

    const pmAt = WSS_parseDate_(r.payment_method_created_at)
    if (!pmAt) continue

    const fpAt = WSS_parseDate_(r.first_payment_at)
    if (fpAt) {
      if ((fpAt.getTime() - pmAt.getTime()) <= THREE_DAYS_MS) continue
      denominator += 1
      numerator += 1
    } else {
      denominator += 1
    }
  }

  return denominator ? (numerator / denominator) : 0
}

/* =========================
 * Big 3 — pulled from existing report sheets
 *  - signupToPaidRate: "All the Stats" row labeled "Sign up to paid conversion"
 *  - churnArr:         "All the Stats" row labeled "Churned ARR (subscription-based)"
 *  - arr:              "The Ring" cell B2 (paid ARR KPI)
 * ========================= */

function WSS_readBigThreeMetrics_(ss) {
  const out = { signupToPaidRate: 0, churnArr: 0, arr: 0 }

  const allStats = ss.getSheetByName('All the Stats')
  if (allStats) {
    const lastRow = allStats.getLastRow()
    if (lastRow > 0) {
      const labels = allStats.getRange(1, 1, lastRow, 1).getValues()
      const values = allStats.getRange(1, 2, lastRow, 1).getValues()
      for (let i = 0; i < labels.length; i++) {
        const label = WSS_str_(labels[i][0])
        if (label === 'Sign up to paid conversion') {
          out.signupToPaidRate = WSS_num_(values[i][0])
        } else if (label === 'Churned ARR (subscription-based)') {
          out.churnArr = WSS_num_(values[i][0])
        }
      }
    }
  }

  const ring = ss.getSheetByName('The Ring')
  if (ring) {
    out.arr = WSS_num_(ring.getRange(2, 2).getValue())
  }

  return out
}

/* =========================
 * Managed (Enterprise) exclusion
 * ========================= */

function WSS_buildManagedExclusionSets_(ss) {
  const subIds = new Set()
  const customerIds = new Set()
  const sh = ss.getSheetByName(WEEKLY_SS_CFG.INPUTS.MANUAL_CHANGES)
  if (!sh) return { subIds, customerIds }
  const rows = WSS_readSheetObjects_(sh, 1)
  for (const r of rows) {
    const reason = WSS_str_(r.exclude_reason).toLowerCase()
    if (reason !== WEEKLY_SS_CFG.MANAGED_REASON) continue
    const subId = WSS_str_(r.subscription_id || r.stripe_subscription_id || r.subscription)
    if (subId) subIds.add(subId)
    const custId = WSS_str_(r.customer_id || r.stripe_customer_id)
    if (custId) customerIds.add(custId)
  }
  return { subIds, customerIds }
}

/* =========================
 * Sheet write
 * ========================= */

function WSS_writeSheet_(sheet, r, windowRange, convRates, bigThree) {
  // Preserve rows 1-15 (manual content). Report starts at row 16.
  const START_ROW = 16
  const lastRow = sheet.getMaxRows()
  if (lastRow >= START_ROW) {
    sheet.getRange(START_ROW, 1, lastRow - START_ROW + 1, sheet.getMaxColumns())
      .clearContent()
      .clearFormat()
  }

  const rows = []
  const sectionRows = []        // row numbers of section headers
  const tableHeaderRows = []    // row numbers of "Org | Seats" mini-headers
  const moneyRows = []
  const pctRows = []
  const intRows = []
  const manualRows = []         // for upsell/expansion blank cells

  const push = (a, b) => rows.push([a, b == null ? '' : b])
  const mark = (arr) => arr.push(rows.length) // rows.length = index that WILL be pushed next

  const b3 = bigThree || { signupToPaidRate: 0, churnArr: 0, arr: 0 }

  // --- Title
  mark(sectionRows)
  push(`Weekly SS Report — week of ${windowRange.startStr} to ${windowRange.endStr}`, '')
  push('Generated at', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'))
  push('', '')

  // --- 1) The Big 3
  mark(sectionRows)
  push('1) The Big 3', '')
  mark(pctRows);   push('Conversion rate (Signup → Paid)', b3.signupToPaidRate)
  mark(moneyRows); push('Churn $ (subscription-based)', b3.churnArr)
  mark(moneyRows); push('Revenue (ARR)', b3.arr)
  push('', '')

  // --- 2) Current State
  mark(sectionRows)
  push('2) Current State', '')
  mark(moneyRows); push('Current ARR', r.currentArr)
  mark(moneyRows); push('New ARR this week', r.newArrThisWeek)
  mark(moneyRows); push('Churn $', r.churnDollars)
  mark(intRows);   push('Churn org count', r.churnOrgCount)
  mark(intRows);   push('Churn seat count', r.churnSeatCount)
  push('', '')
  push('Churned orgs', '')
  mark(tableHeaderRows); push('Org', 'Seats')
  if (r.churnOrgs.length) {
    for (const o of r.churnOrgs) {
      mark(intRows); push(WSS_stripeSubLink_(o.name, o.subId), o.seats)
    }
  } else {
    push('(none this week)', '')
  }
  push('', '')

  push('New paying orgs (first payment this week)', '')
  mark(tableHeaderRows); push('Org', 'Email')
  if (r.newPayingOrgs.length) {
    for (const o of r.newPayingOrgs) {
      push(WSS_stripeSubLink_(o.name, o.subId), o.email)
    }
  } else {
    push('(none this week)', '')
  }
  push('', '')

  // --- 3) Pipeline
  mark(sectionRows)
  push('3) Pipeline', '')
  mark(moneyRows); push('Trialing $', r.trialingDollars)
  mark(moneyRows); push('Intent to Pay $', r.intentToPayDollars)
  mark(pctRows);   push('Signup → Paid %', convRates.signupToPaidRate)
  mark(pctRows);   push('Intent to Pay → Paid %', convRates.intentToPaidRate)
  mark(moneyRows); push('Weighted Pipeline $', r.weightedPipeline)
  push('', '')
  push('Top 5 pipeline orgs', '')
  mark(tableHeaderRows); push('Org', 'Seats')
  if (r.pipelineTopSeats.length) {
    for (const o of r.pipelineTopSeats) {
      mark(intRows); push(`${o.name} (${o.status})`, o.seats)
    }
  } else {
    push('(none)', '')
  }
  push('', '')

  // --- 4) Current Paying Base
  mark(sectionRows)
  push('4) Current Paying Base', '')
  push('Top 5 paying clients', '')
  mark(tableHeaderRows); push('Org', 'Seats')
  if (r.topPayingBySeats.length) {
    for (const o of r.topPayingBySeats) {
      mark(intRows); push(o.name, o.seats)
    }
  } else {
    push('(none)', '')
  }
  push('', '')

  push('Top 5 clients to upsell (manual)', '')
  mark(tableHeaderRows); const upsellHeaderIdx = rows.length; push('Org Name', 'Seat Count')
  const upsellDataStartIdx = rows.length
  for (let i = 0; i < 5; i++) {
    mark(manualRows); mark(intRows); push('', '')
  }
  push('', '')

  push('Expansion seats this week (manual)', '')
  mark(intRows); mark(manualRows); push('Total seats added', '')

  // --- Write (starting at row 11, preserving rows 1-10)
  const width = 2
  sheet.getRange(START_ROW, 1, rows.length, width).setValues(rows)

  // Extra columns for upsell table (manual fill)
  sheet.getRange(upsellHeaderIdx + START_ROW, 3, 1, 2)
    .setValues([['Employee Count', 'Upsell Seat Count']])
    .setFontWeight('bold').setBackground('#FAFAFA')
  sheet.getRange(upsellDataStartIdx + START_ROW, 3, 5, 2)
    .setBackground('#FFF8E1')
    .setNumberFormat(WEEKLY_SS_CFG.FMT.INT)

  // --- Format
  for (const idx of sectionRows) {
    sheet.getRange(idx + START_ROW, 1, 1, width).setFontWeight('bold').setBackground('#F3F4F6')
  }
  for (const idx of tableHeaderRows) {
    sheet.getRange(idx + START_ROW, 1, 1, width).setFontWeight('bold').setBackground('#FAFAFA')
  }
  for (const idx of moneyRows) {
    sheet.getRange(idx + START_ROW, 2).setNumberFormat(WEEKLY_SS_CFG.FMT.MONEY)
  }
  for (const idx of pctRows) {
    sheet.getRange(idx + START_ROW, 2).setNumberFormat(WEEKLY_SS_CFG.FMT.PCT)
  }
  for (const idx of intRows) {
    sheet.getRange(idx + START_ROW, 2).setNumberFormat(WEEKLY_SS_CFG.FMT.INT)
  }
  for (const idx of manualRows) {
    sheet.getRange(idx + START_ROW, 1, 1, width).setBackground('#FFF8E1')
  }

  sheet.setColumnWidth(1, 360)
  sheet.setColumnWidth(2, 120)
  sheet.setColumnWidth(3, 140)
  sheet.setColumnWidth(4, 160)
}

/* =========================
 * Week window (last completed Mon-Sun)
 * ========================= */

function WSS_lastCompletedWeek_(now) {
  // Find Monday of the week containing `now`, then:
  //   end   = that Monday - 1 millisecond (Sunday 23:59:59.999)
  //   start = that Monday - 7 days
  const d = new Date(now.getFullYear(), now.getMonth(), now.getDate())
  const dow = d.getDay() // 0=Sun..6=Sat
  const daysBackToMonday = dow === 0 ? 6 : (dow - 1)
  const thisMonday = new Date(d.getFullYear(), d.getMonth(), d.getDate() - daysBackToMonday)
  const start = new Date(thisMonday.getFullYear(), thisMonday.getMonth(), thisMonday.getDate() - 7)
  const endExclusive = new Date(thisMonday.getFullYear(), thisMonday.getMonth(), thisMonday.getDate())
  const endInclusiveMs = endExclusive.getTime() - 1
  const tz = Session.getScriptTimeZone()
  return {
    startMs: start.getTime(),
    endMs: endInclusiveMs,
    startStr: Utilities.formatDate(start, tz, 'yyyy-MM-dd'),
    endStr: Utilities.formatDate(new Date(endInclusiveMs), tz, 'yyyy-MM-dd')
  }
}

function WSS_isInRange_(d, startMs, endMs) {
  if (!(d instanceof Date)) return false
  const t = d.getTime()
  return t >= startMs && t <= endMs
}

function WSS_stripeSubLink_(label, subId) {
  if (!subId) return label || ''
  const safeLabel = String(label || subId).replace(/"/g, '""')
  return `=HYPERLINK("https://dashboard.stripe.com/subscriptions/${subId}","${safeLabel}")`
}

/* =========================
 * Sheet IO helpers
 * ========================= */

function WSS_readSheetObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1 || lastCol < 1) return []
  const header = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || '').trim().toLowerCase().replace(/\s+/g, '_'))
  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()
  return data.map(row => {
    const obj = {}
    for (let i = 0; i < header.length; i++) {
      const h = header[i]
      if (h) obj[h] = row[i]
    }
    return obj
  })
}

function WSS_getOrCreateSheet_(ss, name) {
  if (typeof getOrCreateSheet === 'function') {
    try { return getOrCreateSheet(ss, name) } catch (e) {}
    try { return getOrCreateSheet(name) } catch (e) {}
  }
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

/* =========================
 * Small utils
 * ========================= */

function WSS_str_(v) {
  if (v === null || v === undefined) return ''
  return String(v).trim()
}

function WSS_num_(v) {
  if (v instanceof Date) return 0
  const n = Number(v)
  return isFinite(n) ? n : 0
}

function WSS_csvList_(v) {
  return WSS_str_(v).split(',').map(s => s.trim()).filter(Boolean)
}

function WSS_parseDate_(v) {
  if (!v && v !== 0) return null
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v
  const s = String(v).trim()
  if (!s) return null
  if (/^\d+$/.test(s)) {
    const n = Number(s)
    const ms = n > 1e12 ? n : n * 1000
    const d = new Date(ms)
    return isNaN(d.getTime()) ? null : d
  }
  const d = new Date(s)
  return isNaN(d.getTime()) ? null : d
}

/* =========================
 * Lock wrap
 * ========================= */

function WSS_lockWrap_(name, fn) {
  if (typeof lockWrap === 'function') {
    try { return lockWrap(name, fn) } catch (e) { return lockWrap(fn) }
  }
  const lock = LockService.getScriptLock()
  if (!lock.tryLock(300000)) throw new Error(`Could not acquire lock: ${name}`)
  try { return fn() } finally { lock.releaseLock() }
}
