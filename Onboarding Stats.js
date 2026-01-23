/**************************************************************
 * render_onboarding_stats()
 *
 * Builds "Onboarding stats" from raw_posthog_user_metrics,
 * grouped by account creation month (raw_clerk_users.created_at).
 *
 * Outputs counts + percentages for:
 * - Calendar connected
 * - Email connected
 * - PM connected (any provider)
 **************************************************************/

const ONB_CFG = {
  SHEET_NAME: 'Onboarding stats',
  POSTHOG_SHEET: 'raw_posthog_user_metrics',
  CLERK_SHEET: 'raw_clerk_users',

  HEADER_ROW: 1,
  DATA_START_ROW: 2,

  WINDOW_DAYS: 14,
  MONTH_FMT: 'yyyy-MM',
  CUTOFF_DATE: '2026-01-11',

  HEADERS: [
    'cohort_month',
    'users_total',

    'calendar_connected_count',
    'calendar_connected_pct',

    'email_connected_count',
    'email_connected_pct',

    'pm_connected_count',
    'pm_connected_pct'
  ],

  PCT_HEADERS: [
    'calendar_connected_pct',
    'email_connected_pct',
    'pm_connected_pct'
  ],

  COUNT_FMT: '0',
  PCT_FMT: '0.0%'
}

function render_onboarding_stats() {
  return ONB_lockWrapCompat_('render_onboarding_stats', () => {
    if (typeof COMBINED_renderConversionOnboarding_ !== 'function') {
      throw new Error('Combined stats renderer not available.')
    }
    return COMBINED_renderConversionOnboarding_({ logStepName: 'render_onboarding_stats' })
  })
}

/* =========================
 * Core logic
 * ========================= */

function ONB_buildCreatedAtIndex_(sheet) {
  const { map } = readHeaderMap(sheet, 1)
  const emailKeyCol = map['email_key']
  const emailCol = map['email']
  const createdCol = map['created_at']

  if (!createdCol) throw new Error('raw_clerk_users missing header: created_at')
  if (!emailKeyCol && !emailCol) throw new Error('raw_clerk_users missing header: email_key or email')

  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < 2) return new Map()

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
  const out = new Map()

  data.forEach(r => {
    const emailRaw = emailKeyCol ? r[emailKeyCol - 1] : r[emailCol - 1]
    const emailFallback = emailCol ? r[emailCol - 1] : ''
    const emailKey = ONB_normalizeEmailCompat_(emailRaw || emailFallback)
    if (!emailKey) return

    const createdAt = ONB_parseDate_(r[createdCol - 1])
    if (!createdAt) return

    const prev = out.get(emailKey)
    if (!prev || createdAt.getTime() < prev.getTime()) out.set(emailKey, createdAt)
  })

  return out
}

function ONB_buildStats_(sheet, createdByEmailKey, tz) {
  const { map } = readHeaderMap(sheet, 1)
  const emailKeyCol = map['email_key']
  const emailCol = map['email']

  if (!emailKeyCol && !emailCol) {
    throw new Error('raw_posthog_user_metrics missing header: email_key or email')
  }

  const idx = {
    calConnected: map['calendar_connected'],
    calDate: map['first_calendar_connected_date'],
    emailConnected: map['email_connected'],
    emailDate: map['first_email_connected_date'],
    pmKarbonConnected: map['pm_karbon_connected'],
    pmKarbonDate: map['pm_karbon_first_connected_date'],
    pmKeeperConnected: map['pm_keeper_connected'],
    pmKeeperDate: map['pm_keeper_first_connected_date'],
    pmFcConnected: map['pm_financial_cents_connected'],
    pmFcDate: map['pm_financial_cents_first_connected_date']
  }

  const requiredHeaders = [
    'calendar_connected',
    'first_calendar_connected_date',
    'email_connected',
    'first_email_connected_date',
    'pm_karbon_connected',
    'pm_karbon_first_connected_date',
    'pm_keeper_connected',
    'pm_keeper_first_connected_date',
    'pm_financial_cents_connected',
    'pm_financial_cents_first_connected_date'
  ]

  requiredHeaders.forEach(h => {
    if (!map[h]) throw new Error(`raw_posthog_user_metrics missing header: ${h}`)
  })

  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()

  const out = new Map()
  const cutoff = ONB_parseDate_(ONB_CFG.CUTOFF_DATE)
  if (!cutoff) throw new Error(`Invalid cutoff date: ${ONB_CFG.CUTOFF_DATE}`)

  const summary = {
    before: ONB_initBucket_('Before 1/11'),
    after: ONB_initBucket_('On/After 1/11'),
    total: ONB_initBucket_('TOTAL')
  }

  if (lastRow < 2) {
    return {
      byMonth: out,
      summary,
      monthCount: 0,
      summaryCount: 3
    }
  }

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()

  data.forEach(r => {
    const emailRaw = emailKeyCol ? r[emailKeyCol - 1] : r[emailCol - 1]
    const emailFallback = emailCol ? r[emailCol - 1] : ''
    const emailKey = ONB_normalizeEmailCompat_(emailRaw || emailFallback)
    if (!emailKey) return

    const createdAt = createdByEmailKey.get(emailKey)
    if (!createdAt) return

    const cohortMonth = ONB_formatMonth_(createdAt, tz)
    if (!cohortMonth) return

    const calDate = ONB_parseDate_(r[idx.calDate - 1])
    const emailDate = ONB_parseDate_(r[idx.emailDate - 1])
    const pmKarbonDate = ONB_parseDate_(r[idx.pmKarbonDate - 1])
    const pmKeeperDate = ONB_parseDate_(r[idx.pmKeeperDate - 1])
    const pmFcDate = ONB_parseDate_(r[idx.pmFcDate - 1])
    const pmDate = ONB_minDate_([pmKarbonDate, pmKeeperDate, pmFcDate])

    const flags = {
      calConnected: ONB_isTruthy_(r[idx.calConnected - 1]) || !!calDate,
      calWithin2w: calDate && ONB_isWithinWindow_(calDate, createdAt, ONB_CFG.WINDOW_DAYS),

      emailConnected: ONB_isTruthy_(r[idx.emailConnected - 1]) || !!emailDate,
      emailWithin2w: emailDate && ONB_isWithinWindow_(emailDate, createdAt, ONB_CFG.WINDOW_DAYS),

      pmConnected:
        ONB_isTruthy_(r[idx.pmKarbonConnected - 1]) ||
        ONB_isTruthy_(r[idx.pmKeeperConnected - 1]) ||
        ONB_isTruthy_(r[idx.pmFcConnected - 1]) ||
        !!pmDate,
      pmWithin2w: pmDate && ONB_isWithinWindow_(pmDate, createdAt, ONB_CFG.WINDOW_DAYS)
    }

    const bucket = ONB_getBucket_(out, cohortMonth)
    ONB_applyFlagsToBucket_(bucket, flags)

    ONB_applyFlagsToBucket_(summary.total, flags)
    if (createdAt.getTime() < cutoff.getTime()) {
      ONB_applyFlagsToBucket_(summary.before, flags)
    } else {
      ONB_applyFlagsToBucket_(summary.after, flags)
    }
  })

  return {
    byMonth: out,
    summary,
    monthCount: out.size,
    summaryCount: 3
  }
}

function ONB_getBucket_(map, monthKey) {
  if (!map.has(monthKey)) {
    map.set(monthKey, ONB_initBucket_(monthKey))
  }
  return map.get(monthKey)
}

function ONB_initBucket_(label) {
  return {
    month: label,
    total: 0,
    calendar: { connected: 0, within2w: 0 },
    email: { connected: 0, within2w: 0 },
    pm: { connected: 0, within2w: 0 }
  }
}

function ONB_applyFlagsToBucket_(bucket, flags) {
  bucket.total += 1

  if (flags.calConnected) bucket.calendar.connected += 1
  if (flags.calWithin2w) bucket.calendar.within2w += 1

  if (flags.emailConnected) bucket.email.connected += 1
  if (flags.emailWithin2w) bucket.email.within2w += 1

  if (flags.pmConnected) bucket.pm.connected += 1
  if (flags.pmWithin2w) bucket.pm.within2w += 1
}

function ONB_buildRows_(statsByMonth, summary) {
  const keys = Array.from(statsByMonth.keys()).sort()
  const rows = keys.map(k => ONB_rowFromBucket_(statsByMonth.get(k)))

  if (summary && summary.before && summary.after && summary.total) {
    rows.push(ONB_rowFromBucket_(summary.before))
    rows.push(ONB_rowFromBucket_(summary.after))
    rows.push(ONB_rowFromBucket_(summary.total))
  }

  return rows
}

function ONB_rowFromBucket_(bucket) {
  const total = bucket.total || 0

  const calPct = total ? bucket.calendar.connected / total : 0
  const emailPct = total ? bucket.email.connected / total : 0
  const pmPct = total ? bucket.pm.connected / total : 0

  return [
    bucket.month,
    total,

    bucket.calendar.connected,
    calPct,

    bucket.email.connected,
    emailPct,

    bucket.pm.connected,
    pmPct
  ]
}

/* =========================
 * Formatting
 * ========================= */

function ONB_applyFormats_(sheet, numDataRows, monthCount, summaryCount) {
  return ONB_applyFormatsAt_(sheet, ONB_CFG.HEADER_ROW, ONB_CFG.DATA_START_ROW, numDataRows, monthCount, summaryCount)
}

function ONB_applyFormatsAt_(sheet, headerRow, dataStartRow, numDataRows, monthCount, summaryCount) {
  const headerRange = sheet.getRange(headerRow, 1, 1, ONB_CFG.HEADERS.length)
  headerRange.setFontWeight('bold').setBackground('#F3F3F3')

  if (!numDataRows) return

  const startRow = dataStartRow
  const nRows = numDataRows

  const countCols = ONB_CFG.HEADERS
    .map((h, i) => (ONB_CFG.PCT_HEADERS.indexOf(h) >= 0 || h === 'cohort_month') ? -1 : i + 1)
    .filter(i => i > 0)

  countCols.forEach(col => {
    sheet.getRange(startRow, col, nRows, 1).setNumberFormat(ONB_CFG.COUNT_FMT)
  })

  ONB_CFG.PCT_HEADERS.forEach(h => {
    const col = ONB_CFG.HEADERS.indexOf(h) + 1
    if (col <= 0) return
    sheet.getRange(startRow, col, nRows, 1).setNumberFormat(ONB_CFG.PCT_FMT)
  })

  const summaryRows = Number(summaryCount || 0)
  const monthlyRows = Number(monthCount || 0)
  if (summaryRows > 0 && monthlyRows >= 0) {
    const summaryStart = dataStartRow + monthlyRows
    sheet.getRange(summaryStart, 1, summaryRows, ONB_CFG.HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#F6F4F0')
  }
}

/* =========================
 * Helpers
 * ========================= */

function ONB_parseDate_(v) {
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

function ONB_formatMonth_(dateObj, tz) {
  if (!dateObj) return ''
  return Utilities.formatDate(dateObj, tz, ONB_CFG.MONTH_FMT)
}

function ONB_isTruthy_(v) {
  if (v === true) return true
  const s = String(v || '').trim().toLowerCase()
  return s === 'true' || s === 'yes' || s === '1'
}

function ONB_isWithinWindow_(dateObj, createdAt, days) {
  if (!dateObj || !createdAt) return false
  const ms = dateObj.getTime() - createdAt.getTime()
  if (!isFinite(ms)) return false
  return ms >= 0 && ms <= (Number(days || 14) * 24 * 60 * 60 * 1000)
}

function ONB_minDate_(dates) {
  let best = null
  ;(dates || []).forEach(d => {
    if (!d) return
    if (!best || d.getTime() < best.getTime()) best = d
  })
  return best
}

function ONB_normalizeEmailCompat_(email) {
  if (typeof normalizeEmail === 'function') return normalizeEmail(email)
  return String(email || '').trim().toLowerCase()
}

/* =========================
 * Compatibility wrappers
 * ========================= */

function ONB_getOrCreateSheetCompat_(ss, name) {
  if (typeof getOrCreateSheet === 'function') {
    try { return getOrCreateSheet(ss, name) } catch (e) {}
    try { return getOrCreateSheet(name) } catch (e) {}
  }
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function ONB_batchSetValuesCompat_(sheet, startRow, startCol, values, chunkSize) {
  if (typeof batchSetValues === 'function') return batchSetValues(sheet, startRow, startCol, values, chunkSize)
  const size = chunkSize || 2000
  for (let i = 0; i < values.length; i += size) {
    const chunk = values.slice(i, i + size)
    sheet.getRange(startRow + i, startCol, chunk.length, chunk[0].length).setValues(chunk)
  }
}

function ONB_lockWrapCompat_(lockName, fn) {
  if (typeof lockWrap === 'function') {
    try { return lockWrap(lockName, fn) } catch (e) { return lockWrap(fn) }
  }
  const lock = LockService.getScriptLock()
  const ok = lock.tryLock(300000)
  if (!ok) throw new Error(`Could not acquire lock: ${lockName}`)
  try { return fn() } finally { lock.releaseLock() }
}

function ONB_writeSyncLogCompat_(step, status, rowsIn, rowsOut, seconds, error) {
  if (typeof writeSyncLog === 'function') return writeSyncLog(step, status, rowsIn, rowsOut, seconds, error || '')
  Logger.log(`[SYNCLOG missing] ${step} ${status} rows_in=${rowsIn} rows_out=${rowsOut} seconds=${seconds} error=${error || ''}`)
}
