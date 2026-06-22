/**************************************************************
 * render_org_conversion_stats()
 *
 * Builds "Conversion stats" by org signup month:
 * - orgs_signed_up
 * - orgs_converted (has any Stripe subscription on raw_stripe_subscriptions)
 * - conversion_rate
 * - orgs_converted_within_7d_trial_end (clean conversion)
 * - conversion_rate_within_7d_trial_end
 *
 * Conversion is determined by:
 * arr_raw_data.subscription_start_date (derived from Stripe)
 *
 * Clean conversion is determined by:
 * - trial_start_date + trial_end_date are read from arr_raw_data
 * - clean conversion uses purchase_date from arr_raw_data
 **************************************************************/

const CONV_CFG = {
  SHEET_NAME: 'Conversion stats',
  HEADER_ROW: 1,
  DATA_START_ROW: 2,

  INPUTS: {
    CLERK_ORGS: 'raw_clerk_orgs',
    ARR_RAW_DATA: 'arr_raw_data'
  },
  ARR_RAW_HEADER_ROW: 2,

  MONTH_FMT: 'yyyy-MM',
  DATE_FMT: 'MM-dd-yy',
  HEADERS: [
    'cohort_month',
    'orgs_signed_up',
    'orgs_subscribed',
    'sub_rate',
    'orgs_paid',
    'conv_rate',
    'orgs_paid_7d',
    'paid_7d_rate',
    'orgs_in_7d_window',
    'paid_7d_potential_rate'
  ],
  COUNT_FMT: '0',
  PCT_FMT: '0.0%'
}

function render_org_conversion_stats() {
  return CONV_lockWrapCompat_('render_org_conversion_stats', () => {
    if (typeof COMBINED_renderConversionOnboarding_ !== 'function') {
      throw new Error('Combined stats renderer not available.')
    }
    return COMBINED_renderConversionOnboarding_({ logStepName: 'render_org_conversion_stats' })
  })
}

/* =========================
 * Core logic
 * ========================= */

function CONV_getBucket_(map, monthKey) {
  if (!map.has(monthKey)) {
    map.set(monthKey, {
      month: monthKey,
      total: 0,
      subscribed: 0,
      paid: 0,
      paidWithin7d: 0,
      inWindowUnpaid: 0
    })
  }
  return map.get(monthKey)
}

function CONV_collectStatsByMonth_(shOrgs, shArrRaw, tz, ss) {
  // Uses shared canon_orgs-based org list (filtered, enriched with first_payment_at)
  const convOrgs = ss ? buildConversionOrgList_(ss) : []
  const now = new Date()
  const statsByMonth = new Map()

  for (const org of convOrgs) {
    const createdAt = org.org_created_at
    if (!createdAt) continue

    const cohortMonth = Utilities.formatDate(createdAt, tz, CONV_CFG.MONTH_FMT)
    if (!cohortMonth) continue

    const bucket = CONV_getBucket_(statsByMonth, cohortMonth)
    bucket.total += 1

    // Has a subscription (is_paying or has first_payment)
    if (org.is_paying === 'TRUE' || org.is_paying === 'true' || org.first_payment_at) {
      bucket.subscribed += 1
    }

    if (org.first_payment_at) bucket.paid += 1

    // Paid within 7 days of trial end
    const trialEnd = org.trial_ends_at
    const trialWindowEnd = trialEnd ? CONV_addDays_(trialEnd, 7) : null
    const within7d =
      trialEnd &&
      trialWindowEnd &&
      org.first_payment_at &&
      CONV_isWithinRange_(org.first_payment_at, trialEnd, trialWindowEnd)
    if (within7d) bucket.paidWithin7d += 1

    const inWindowUnpaid =
      !within7d &&
      trialEnd &&
      trialWindowEnd &&
      !org.first_payment_at &&
      CONV_isWithinRange_(now, trialEnd, trialWindowEnd)
    if (inWindowUnpaid) bucket.inWindowUnpaid += 1
  }

  return statsByMonth
}

function CONV_buildRows_(statsByMonth) {
  const keys = Array.from(statsByMonth.keys()).sort()
  const rows = keys.map(k => {
    const s = statsByMonth.get(k)
    const total = s.total || 0
    const subscribed = s.subscribed || 0
    const subRate = total ? subscribed / total : 0
    const paid = s.paid || 0
    const convRate = total ? paid / total : 0
    const paid7d = s.paidWithin7d || 0
    const paid7dRate = total ? paid7d / total : 0
    const inWindow = s.inWindowUnpaid || 0
    const paid7dPotentialRate = total ? (paid7d + inWindow) / total : 0

    return [
      s.month,
      total,
      subscribed,
      subRate,
      paid,
      convRate,
      paid7d,
      paid7dRate,
      inWindow,
      paid7dPotentialRate
    ]
  })

  if (!rows.length) return rows

  let total = 0
  let subscribed = 0
  let paid = 0
  let paid7d = 0
  let inWindow = 0

  rows.forEach(r => {
    total += Number(r[1]) || 0
    subscribed += Number(r[2]) || 0
    paid += Number(r[4]) || 0
    paid7d += Number(r[6]) || 0
    inWindow += Number(r[8]) || 0
  })

  const subRate = total ? subscribed / total : 0
  const convRate = total ? paid / total : 0
  const paid7dRate = total ? paid7d / total : 0
  const paid7dPotentialRate = total ? (paid7d + inWindow) / total : 0

  rows.push([
    'TOTAL',
    total,
    subscribed,
    subRate,
    paid,
    convRate,
    paid7d,
    paid7dRate,
    inWindow,
    paid7dPotentialRate
  ])

  return rows
}

function CONV_buildOrgInfoById_(orgInfoRows) {
  const out = new Map()
  ;(orgInfoRows || []).forEach(r => {
    const appOrgId = CONV_str_(r.app_org_id)
    const legacyOrgId = CONV_str_(r.org_id)
    const payload = {
      appOrgId: appOrgId || legacyOrgId,
      subId: CONV_str_(r.stripe_subscription_id || r.subscription_id || r.latest_subscription_id),
      trialStartDate: CONV_parseDate_(r.trial_start_date),
      trialEndDate: CONV_parseDate_(r.trial_end_date),
      subscriptionStartDate: CONV_parseDate_(r.subscription_start_date),
      purchaseDate: CONV_parseDate_(r.purchase_date)
    }

    if (legacyOrgId) out.set(legacyOrgId, payload)
    if (appOrgId) out.set(appOrgId, payload)
  })
  return out
}

/* =========================
 * Sheet IO
 * ========================= */

function CONV_readSheetObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1) return []

  const header = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || '').trim())

  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()

  return data.map(r => {
    const obj = {}
    header.forEach((h, i) => {
      if (!h) return
      obj[CONV_key_(h)] = r[i]
    })
    return obj
  })
}

function CONV_key_(h) {
  return String(h || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_')
}

/* =========================
 * Formatting
 * ========================= */

function CONV_applyFormats_(sheet, numDataRows) {
  return CONV_applyFormatsAt_(sheet, CONV_CFG.HEADER_ROW, CONV_CFG.DATA_START_ROW, numDataRows)
}

function CONV_applyFormatsAt_(sheet, headerRow, dataStartRow, numDataRows) {
  const headerRange = sheet.getRange(headerRow, 1, 1, CONV_CFG.HEADERS.length)
  headerRange.setFontWeight('bold').setBackground('#F3F3F3')

  if (!numDataRows) return

  const startRow = dataStartRow
  const nRows = numDataRows

  const colTotal = CONV_CFG.HEADERS.indexOf('orgs_signed_up') + 1
  const colSubscribed = CONV_CFG.HEADERS.indexOf('orgs_subscribed') + 1
  const colSubRate = CONV_CFG.HEADERS.indexOf('sub_rate') + 1
  const colPaid = CONV_CFG.HEADERS.indexOf('orgs_paid') + 1
  const colConvRate = CONV_CFG.HEADERS.indexOf('conv_rate') + 1
  const colPaid7d = CONV_CFG.HEADERS.indexOf('orgs_paid_7d') + 1
  const colPaid7dRate = CONV_CFG.HEADERS.indexOf('paid_7d_rate') + 1
  const colInWindow = CONV_CFG.HEADERS.indexOf('orgs_in_7d_window') + 1
  const colPotentialRate = CONV_CFG.HEADERS.indexOf('paid_7d_potential_rate') + 1

  if (colTotal > 0) sheet.getRange(startRow, colTotal, nRows, 1).setNumberFormat(CONV_CFG.COUNT_FMT)
  if (colSubscribed > 0) sheet.getRange(startRow, colSubscribed, nRows, 1).setNumberFormat(CONV_CFG.COUNT_FMT)
  if (colSubRate > 0) sheet.getRange(startRow, colSubRate, nRows, 1).setNumberFormat(CONV_CFG.PCT_FMT)
  if (colPaid > 0) sheet.getRange(startRow, colPaid, nRows, 1).setNumberFormat(CONV_CFG.COUNT_FMT)
  if (colConvRate > 0) sheet.getRange(startRow, colConvRate, nRows, 1).setNumberFormat(CONV_CFG.PCT_FMT)
  if (colPaid7d > 0) sheet.getRange(startRow, colPaid7d, nRows, 1).setNumberFormat(CONV_CFG.COUNT_FMT)
  if (colPaid7dRate > 0) sheet.getRange(startRow, colPaid7dRate, nRows, 1).setNumberFormat(CONV_CFG.PCT_FMT)
  if (colInWindow > 0) sheet.getRange(startRow, colInWindow, nRows, 1).setNumberFormat(CONV_CFG.COUNT_FMT)
  if (colPotentialRate > 0) sheet.getRange(startRow, colPotentialRate, nRows, 1).setNumberFormat(CONV_CFG.PCT_FMT)

  const totalRow = dataStartRow + nRows - 1
  if (totalRow >= dataStartRow) {
    sheet.getRange(totalRow, 1, 1, CONV_CFG.HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#F6F4F0')
  }
}

/* =========================
 * Helpers
 * ========================= */

function CONV_str_(v) {
  if (v === null || v === undefined) return ''
  return String(v).trim()
}

function CONV_normEmail_(v) {
  const s = String(v || '').trim().toLowerCase()
  if (!s) return ''
  return s.replace(/\+[^@]+(?=@)/, '')
}

function CONV_parseDate_(v) {
  if (!v) return null
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v

  const s = String(v || '').trim()
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

function CONV_addDays_(dateObj, days) {
  const d = new Date(dateObj.getTime())
  d.setUTCDate(d.getUTCDate() + Number(days || 0))
  return d
}

function CONV_isWithinRange_(dateObj, startObj, endObj) {
  if (!dateObj || !startObj || !endObj) return false
  const t = dateObj.getTime()
  const s = startObj.getTime()
  const e = endObj.getTime()
  if (!isFinite(t) || !isFinite(s) || !isFinite(e)) return false
  return t >= s && t <= e
}

/* =========================
 * Compatibility wrappers
 * ========================= */

function CONV_getOrCreateSheetCompat_(ss, name) {
  if (typeof getOrCreateSheet === 'function') {
    try { return getOrCreateSheet(ss, name) } catch (e) {}
    try { return getOrCreateSheet(name) } catch (e) {}
  }
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function CONV_batchSetValuesCompat_(sheet, startRow, startCol, values, chunkSize) {
  if (typeof batchSetValues === 'function') return batchSetValues(sheet, startRow, startCol, values, chunkSize)
  const size = chunkSize || 2000
  for (let i = 0; i < values.length; i += size) {
    const chunk = values.slice(i, i + size)
    sheet.getRange(startRow + i, startCol, chunk.length, chunk[0].length).setValues(chunk)
  }
}

function CONV_lockWrapCompat_(lockName, fn) {
  if (typeof lockWrap === 'function') {
    try { return lockWrap(lockName, fn) } catch (e) { return lockWrap(fn) }
  }
  const lock = LockService.getScriptLock()
  const ok = lock.tryLock(300000)
  if (!ok) throw new Error(`Could not acquire lock: ${lockName}`)
  try { return fn() } finally { lock.releaseLock() }
}

function CONV_writeSyncLogCompat_(step, status, rowsIn, rowsOut, seconds, error) {
  if (typeof writeSyncLog === 'function') return writeSyncLog(step, status, rowsIn, rowsOut, seconds, error || '')
  Logger.log(`[SYNCLOG missing] ${step} ${status} rows_in=${rowsIn} rows_out=${rowsOut} seconds=${seconds} error=${error || ''}`)
}

function CONV_buildExcludedSubIds_(ss) {
  const out = new Set()
  const sh = ss.getSheetByName('Manual Stripe Changes')
  if (!sh) return out
  const rows = CONV_readSheetObjects_(sh, 1)
  const EXCLUDE_REASONS = new Set(['internal', 'partner', 'free subscription'])
  for (const r of rows) {
    const reason = CONV_str_(r.exclude_reason).toLowerCase()
    if (!EXCLUDE_REASONS.has(reason)) continue
    const subId = CONV_str_(r.subscription_id || r.stripe_subscription_id || r.subscription)
    if (subId) out.add(subId)
  }
  return out
}

function CONV_buildOwnerEmailByOrgId_(ss) {
  const out = new Map()
  const shOrgs = ss.getSheetByName('raw_clerk_orgs')
  const shUsers = ss.getSheetByName('raw_clerk_users')
  if (!shOrgs || !shUsers) return out

  // Build user_id -> email from raw_clerk_users
  const emailByUserId = new Map()
  const users = CONV_readSheetObjects_(shUsers, 1)
  for (const u of users) {
    const userId = CONV_str_(u.clerk_user_id || u.user_id || u.id)
    const email = CONV_str_(u.email)
    if (userId && email) emailByUserId.set(userId, email)
  }

  // Map org_id -> owner email via org_owner_user_id
  const orgs = CONV_readSheetObjects_(shOrgs, 1)
  for (const o of orgs) {
    const orgId = CONV_str_(o.org_id)
    const ownerId = CONV_str_(o.org_owner_user_id)
    if (orgId && ownerId) {
      const email = emailByUserId.get(ownerId) || ''
      if (email) out.set(orgId, email)
    }
  }
  return out
}
