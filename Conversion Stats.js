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
 * org_info.subscription_start_date (derived from Stripe)
 *
 * Clean conversion is determined by:
 * - trial_start_date + trial_end_date are read from org_info
 * - clean conversion uses purchase_date from org_info
 **************************************************************/

const CONV_CFG = {
  SHEET_NAME: 'Conversion stats',
  HEADER_ROW: 1,
  DATA_START_ROW: 2,

  INPUTS: {
    CLERK_ORGS: 'raw_clerk_orgs',
    ORG_INFO: 'org_info'
  },

  MONTH_FMT: 'yyyy-MM',
  HEADERS: [
    'cohort_month',
    'orgs_signed_up',
    'orgs_converted',
    'conversion_rate',
    'orgs_converted_within_7d_trial_end',
    'conversion_rate_within_7d_trial_end'
  ],
  COUNT_FMT: '0',
  PCT_FMT: '0.0%'
}

function render_org_conversion_stats() {
  CONV_lockWrapCompat_('render_org_conversion_stats', () => {
    const t0 = new Date()
    const ss = SpreadsheetApp.getActive()

    const shOut = CONV_getOrCreateSheetCompat_(ss, CONV_CFG.SHEET_NAME)
    const shOrgs = ss.getSheetByName(CONV_CFG.INPUTS.CLERK_ORGS)
    const shOrgInfo = ss.getSheetByName(CONV_CFG.INPUTS.ORG_INFO)

    if (!shOrgs) throw new Error(`Missing input sheet: ${CONV_CFG.INPUTS.CLERK_ORGS}`)
    if (!shOrgInfo) throw new Error(`Missing input sheet: ${CONV_CFG.INPUTS.ORG_INFO}`)

    const tz = Session.getScriptTimeZone()

    const orgs = CONV_readSheetObjects_(shOrgs, 1)
    const orgInfo = CONV_readSheetObjects_(shOrgInfo, 1)
    const orgInfoById = CONV_buildOrgInfoById_(orgInfo)

    const statsByMonth = new Map()

    orgs.forEach(o => {
      const orgId = CONV_str_(o.org_id)
      if (!orgId) return

      const createdAt = CONV_parseDate_(o.created_at || o.org_created_at)
      if (!createdAt) return

      const cohortMonth = Utilities.formatDate(createdAt, tz, CONV_CFG.MONTH_FMT)
      if (!cohortMonth) return

      const bucket = CONV_getBucket_(statsByMonth, cohortMonth)
      bucket.total += 1

      const info = orgInfoById.get(orgId) || {}
      const hasConversion = !!info.subscriptionStartDate
      if (hasConversion) bucket.converted += 1

      const trialStart = info.trialStartDate || null
      const trialEnd = info.trialEndDate || null
      const firstPaymentDate = info.purchaseDate || null
      const trialWindowEnd = trialEnd ? CONV_addDays_(trialEnd, 7) : null
      const within7d =
        trialStart &&
        trialWindowEnd &&
        firstPaymentDate &&
        CONV_isWithinRange_(firstPaymentDate, trialStart, trialWindowEnd)
      if (within7d) bucket.convertedWithin7d += 1
    })

    const outRows = CONV_buildRows_(statsByMonth)

    shOut.clearContents()
    shOut.getRange(CONV_CFG.HEADER_ROW, 1, 1, CONV_CFG.HEADERS.length).setValues([CONV_CFG.HEADERS])

    if (outRows.length) {
      CONV_batchSetValuesCompat_(shOut, CONV_CFG.DATA_START_ROW, 1, outRows, 2000)
    }

    CONV_applyFormats_(shOut, outRows.length)
    shOut.setFrozenRows(CONV_CFG.HEADER_ROW)
    shOut.autoResizeColumns(1, CONV_CFG.HEADERS.length)

    const seconds = (new Date() - t0) / 1000
    CONV_writeSyncLogCompat_('render_org_conversion_stats', 'ok', outRows.length, outRows.length, seconds, '')
    return { rows_out: outRows.length }
  })
}

/* =========================
 * Core logic
 * ========================= */

function CONV_getBucket_(map, monthKey) {
  if (!map.has(monthKey)) {
    map.set(monthKey, { month: monthKey, total: 0, converted: 0, convertedWithin7d: 0 })
  }
  return map.get(monthKey)
}

function CONV_buildRows_(statsByMonth) {
  const keys = Array.from(statsByMonth.keys()).sort()
  return keys.map(k => {
    const s = statsByMonth.get(k)
    const total = s.total || 0
    const conv = s.converted || 0
    const pct = total ? conv / total : 0
    const conv7d = s.convertedWithin7d || 0
    const pct7d = total ? conv7d / total : 0
    return [s.month, total, conv, pct, conv7d, pct7d]
  })
}

function CONV_buildOrgInfoById_(orgInfoRows) {
  const out = new Map()
  ;(orgInfoRows || []).forEach(r => {
    const orgId = CONV_str_(r.org_id)
    if (!orgId) return
    out.set(orgId, {
      trialStartDate: CONV_parseDate_(r.trial_start_date),
      trialEndDate: CONV_parseDate_(r.trial_end_date),
      subscriptionStartDate: CONV_parseDate_(r.subscription_start_date),
      purchaseDate: CONV_parseDate_(r.purchase_date)
    })
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
  const headerRange = sheet.getRange(1, 1, 1, CONV_CFG.HEADERS.length)
  headerRange.setFontWeight('bold').setBackground('#F3F3F3')

  if (!numDataRows) return

  const startRow = CONV_CFG.DATA_START_ROW
  const nRows = numDataRows

  const colTotal = CONV_CFG.HEADERS.indexOf('orgs_signed_up') + 1
  const colConverted = CONV_CFG.HEADERS.indexOf('orgs_converted') + 1
  const colPct = CONV_CFG.HEADERS.indexOf('conversion_rate') + 1
  const colConv7d = CONV_CFG.HEADERS.indexOf('orgs_converted_within_7d_trial_end') + 1
  const colPct7d = CONV_CFG.HEADERS.indexOf('conversion_rate_within_7d_trial_end') + 1

  if (colTotal > 0) sheet.getRange(startRow, colTotal, nRows, 1).setNumberFormat(CONV_CFG.COUNT_FMT)
  if (colConverted > 0) sheet.getRange(startRow, colConverted, nRows, 1).setNumberFormat(CONV_CFG.COUNT_FMT)
  if (colPct > 0) sheet.getRange(startRow, colPct, nRows, 1).setNumberFormat(CONV_CFG.PCT_FMT)
  if (colConv7d > 0) sheet.getRange(startRow, colConv7d, nRows, 1).setNumberFormat(CONV_CFG.COUNT_FMT)
  if (colPct7d > 0) sheet.getRange(startRow, colPct7d, nRows, 1).setNumberFormat(CONV_CFG.PCT_FMT)
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
