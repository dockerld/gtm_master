/**************************************************************
 * render_middle_earth_view()
 *
 * Org-level operational sheet modeled after Sauron, but focused
 * on canonical org rollups and a cleaner output for org health.
 *
 * Primary source:
 * - canon_orgs
 *
 * Enrichment:
 * - canon_users (meetings/clients/integrations rollups)
 * - raw_stripe_subscriptions (custom subscription status rules)
 *
 * Output sheet:
 * - Middle Earth
 **************************************************************/

const MIDDLE_EARTH_CFG = {
  SHEET_NAME: 'Middle Earth',
  INPUTS: {
    CANON_ORGS: 'canon_orgs',
    CANON_USERS: 'canon_users',
    STRIPE_SUBS: 'raw_stripe_subscriptions'
  },
  HEADER_ROW: 1,
  START_COL: 1,
  DATA_START_ROW: 2,
  HEADERS: [
    'Org Name',
    'In Onboarding',
    'Days with Ping',
    '# of Seats',
    'Meetings Recorded',
    '# of Clients',
    'Subscriptions Status',
    'PM',
    'PM Connected Date',
    'Cals Connected',
    'Emails Connected',
    'Other integration'
  ],
  INT_FMT: '0',
  DATE_FMT: 'yyyy-mm-dd',
  TEXT_FMT: '@'
}

function render_middle_earth_view() {
  return MIDDLEEARTH_lockWrap_('render_middle_earth_view', () => {
    const t0 = new Date()
    try {
      const ss = SpreadsheetApp.getActive()
      const outSh = MIDDLEEARTH_getOrCreateSheet_(ss, MIDDLE_EARTH_CFG.SHEET_NAME)

      const shCanonOrgs = ss.getSheetByName(MIDDLE_EARTH_CFG.INPUTS.CANON_ORGS)
      const shCanonUsers = ss.getSheetByName(MIDDLE_EARTH_CFG.INPUTS.CANON_USERS)
      const shStripe = ss.getSheetByName(MIDDLE_EARTH_CFG.INPUTS.STRIPE_SUBS)

      if (!shCanonOrgs) throw new Error(`Missing input sheet: ${MIDDLE_EARTH_CFG.INPUTS.CANON_ORGS}`)
      if (!shCanonUsers) throw new Error(`Missing input sheet: ${MIDDLE_EARTH_CFG.INPUTS.CANON_USERS}`)
      if (!shStripe) throw new Error(`Missing input sheet: ${MIDDLE_EARTH_CFG.INPUTS.STRIPE_SUBS}`)

      const canonOrgs = MIDDLEEARTH_readSheetObjects_(shCanonOrgs, 1)
      const canonUsers = MIDDLEEARTH_readSheetObjects_(shCanonUsers, 1)
      const stripeRows = MIDDLEEARTH_readSheetObjects_(shStripe, 1)

      const usersByOrgId = MIDDLEEARTH_groupUsersByOrgId_(canonUsers)
      const stripeBySubId = MIDDLEEARTH_buildStripeBySubId_(stripeRows)
      const today = new Date()

      const manualByOrgName = MIDDLEEARTH_readExistingManualByOrgName_(outSh)
      const outRows = []

      for (const org of canonOrgs) {
        const orgId = MIDDLEEARTH_str_(org.app_org_id || org.org_id || org.clerk_org_id)
        if (!orgId) continue

        const orgName = MIDDLEEARTH_str_(org.org_name || org.posthog_org_name || org.org_slug)
        if (!orgName) continue

        const orgCreatedAt = MIDDLEEARTH_toDateOrNull_(org.org_created_at || org.posthog_org_created_at || org.created_at)
        const daysWithPing = orgCreatedAt ? Math.max(0, Math.floor((today.getTime() - orgCreatedAt.getTime()) / 86400000)) : ''

        const seats = MIDDLEEARTH_pickSeats_(org)
        const users = usersByOrgId.get(orgId) || []
        const userAgg = MIDDLEEARTH_aggregateUsers_(users)

        const subIds = MIDDLEEARTH_csvList_(org.stripe_subscription_ids)
        const status = MIDDLEEARTH_computeCustomSubscriptionStatus_(org, subIds, stripeBySubId)

        const priorManual = manualByOrgName.get(orgName) || {}
        const otherIntegration = Object.prototype.hasOwnProperty.call(priorManual, 'other_integration')
          ? priorManual.other_integration
          : ''

        outRows.push([
          orgName,
          MIDDLEEARTH_toBool_(org.in_onboarding),
          daysWithPing,
          seats,
          userAgg.meetingsRecorded,
          userAgg.clientsCount,
          status,
          userAgg.pmSummary,
          userAgg.pmConnectedDate,
          userAgg.calendarConnectedCount,
          userAgg.emailConnectedCount,
          otherIntegration
        ])
      }

      outRows.sort((a, b) => String(a[0] || '').localeCompare(String(b[0] || '')))

      outSh.clear()
      outSh.getRange(MIDDLE_EARTH_CFG.HEADER_ROW, MIDDLE_EARTH_CFG.START_COL, 1, MIDDLE_EARTH_CFG.HEADERS.length)
        .setValues([MIDDLE_EARTH_CFG.HEADERS])
        .setFontWeight('bold')
        .setBackground('#F3F4F6')

      if (outRows.length) {
        MIDDLEEARTH_batchSetValues_(outSh, MIDDLE_EARTH_CFG.DATA_START_ROW, MIDDLE_EARTH_CFG.START_COL, outRows, 1000)
      }

      MIDDLEEARTH_applyFormats_(outSh, outRows.length)
      outSh.setFrozenRows(MIDDLE_EARTH_CFG.HEADER_ROW)
      outSh.autoResizeColumns(MIDDLE_EARTH_CFG.START_COL, MIDDLE_EARTH_CFG.HEADERS.length)

      const seconds = (new Date() - t0) / 1000
      MIDDLEEARTH_writeSyncLog_('render_middle_earth_view', 'ok', canonOrgs.length, outRows.length, seconds, '')
      return { rows_in: canonOrgs.length, rows_out: outRows.length }
    } catch (err) {
      const seconds = (new Date() - t0) / 1000
      MIDDLEEARTH_writeSyncLog_('render_middle_earth_view', 'error', '', '', seconds, String(err && err.message ? err.message : err))
      throw err
    }
  })
}

function MIDDLEEARTH_computeCustomSubscriptionStatus_(org, subIds, stripeBySubId) {
  const appPromoCount = Number(org.app_promo_redemption_count || org.app_promo_code_count || 0) || 0
  const hasAppPromo = appPromoCount > 0 || !!MIDDLEEARTH_str_(org.app_promo_codes)

  let hasActiveWithFirstPayment = false
  let hasActiveNoFirstPayment = false
  let hasTrialingNoFirstWithPromo = false
  let hasTrialingNoFirstNoPromo = false
  let hasAnyCanceled = false
  let hasAnyActive = false
  let hasAnyTrialing = false

  for (const subIdRaw of (subIds || [])) {
    const subId = MIDDLEEARTH_str_(subIdRaw)
    if (!subId) continue
    const r = stripeBySubId.get(subId)
    if (!r) continue

    const status = MIDDLEEARTH_str_(r.status).toLowerCase()
    const hasFirstPayment = !!MIDDLEEARTH_str_(r.first_payment_at)
    const stripePromoUsage =
      !!MIDDLEEARTH_str_(r.promo_code) ||
      !!MIDDLEEARTH_str_(r.promo_code_all)
    const hasPromoUsage = stripePromoUsage || hasAppPromo

    if (status === 'active') {
      hasAnyActive = true
      if (hasFirstPayment) hasActiveWithFirstPayment = true
      else hasActiveNoFirstPayment = true
    } else if (status === 'trialing') {
      hasAnyTrialing = true
      if (!hasFirstPayment && hasPromoUsage) hasTrialingNoFirstWithPromo = true
      if (!hasFirstPayment && !hasPromoUsage) hasTrialingNoFirstNoPromo = true
    } else if (status === 'canceled') {
      hasAnyCanceled = true
    }
  }

  if (hasActiveWithFirstPayment) return 'Paid'
  if (hasActiveNoFirstPayment) return 'Promo Trial'
  if (hasTrialingNoFirstWithPromo) return 'Promo Trial'
  if (hasTrialingNoFirstNoPromo) return 'Free Trial'
  if (hasAnyActive) return 'Paid'
  if (hasAnyTrialing) return 'Trialing'
  if (hasAnyCanceled) return 'Canceled'
  return 'No Subscription'
}

function MIDDLEEARTH_aggregateUsers_(users) {
  const pmProviders = new Set()
  const pmDates = []
  let meetingsRecorded = 0
  let clientsCount = 0
  let calendarConnectedCount = 0
  let emailConnectedCount = 0

  for (const u of (users || [])) {
    meetingsRecorded += (Number(u.meetings_recorded || 0) || 0)
    clientsCount += (Number(u.clients_count || 0) || 0)

    const calConnected = MIDDLEEARTH_toBool_(u.calendar_connected)
    const emailConnected = MIDDLEEARTH_toBool_(u.email_connected)
    if (calConnected) calendarConnectedCount += 1
    if (emailConnected) emailConnectedCount += 1

    if (MIDDLEEARTH_toBool_(u.pm_karbon_connected)) {
      pmProviders.add('KARBON')
      const d = MIDDLEEARTH_toDateOrNull_(u.pm_karbon_first_connected_date)
      if (d) pmDates.push(d)
    }
    if (MIDDLEEARTH_toBool_(u.pm_keeper_connected)) {
      pmProviders.add('KEEPER')
      const d = MIDDLEEARTH_toDateOrNull_(u.pm_keeper_first_connected_date)
      if (d) pmDates.push(d)
    }
    if (MIDDLEEARTH_toBool_(u.pm_financial_cents_connected)) {
      pmProviders.add('FINANCIAL_CENTS')
      const d = MIDDLEEARTH_toDateOrNull_(u.pm_financial_cents_first_connected_date)
      if (d) pmDates.push(d)
    }
  }

  pmDates.sort((a, b) => a.getTime() - b.getTime())
  const firstPmDate = pmDates.length ? pmDates[0] : null

  return {
    meetingsRecorded,
    clientsCount,
    pmSummary: Array.from(pmProviders).join(', '),
    pmConnectedDate: firstPmDate || '',
    calendarConnectedCount,
    emailConnectedCount
  }
}

function MIDDLEEARTH_pickSeats_(org) {
  const seatCandidates = [
    org.seats,
    org.stripe_seats_paying_sum,
    org.stripe_seats_max,
    org.active_subscription_seats
  ]
  for (const v of seatCandidates) {
    const n = Number(v)
    if (isFinite(n) && n > 0) return Math.round(n)
  }
  return 0
}

function MIDDLEEARTH_groupUsersByOrgId_(users) {
  const out = new Map()
  for (const u of (users || [])) {
    const orgId = MIDDLEEARTH_str_(u.org_id)
    if (!orgId) continue
    if (!out.has(orgId)) out.set(orgId, [])
    out.get(orgId).push(u)
  }
  return out
}

function MIDDLEEARTH_buildStripeBySubId_(rows) {
  const out = new Map()
  for (const r of (rows || [])) {
    const subId =
      MIDDLEEARTH_str_(r.stripe_subscription_id) ||
      MIDDLEEARTH_str_(r.subscription_id) ||
      MIDDLEEARTH_str_(r.subscription) ||
      MIDDLEEARTH_str_(r.id)
    if (!subId) continue
    if (!out.has(subId)) out.set(subId, r)
  }
  return out
}

function MIDDLEEARTH_readExistingManualByOrgName_(sheet) {
  const out = new Map()
  if (!sheet) return out
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < 2 || lastCol < 1) return out

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => MIDDLEEARTH_str_(h))
  const orgNameIdx = headers.findIndex(h => h.toLowerCase() === 'org name')
  const otherIdx = headers.findIndex(h => h.toLowerCase() === 'other integration')
  if (orgNameIdx < 0 || otherIdx < 0) return out

  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
  rows.forEach(r => {
    const orgName = MIDDLEEARTH_str_(r[orgNameIdx])
    if (!orgName) return
    out.set(orgName, {
      other_integration: r[otherIdx]
    })
  })
  return new Map(Array.from(out.entries()))
}

function MIDDLEEARTH_applyFormats_(sheet, rowCount) {
  if (!rowCount) return
  const startRow = MIDDLE_EARTH_CFG.DATA_START_ROW

  const colInOnboarding = MIDDLE_EARTH_CFG.HEADERS.indexOf('In Onboarding') + 1
  const colDaysPing = MIDDLE_EARTH_CFG.HEADERS.indexOf('Days with Ping') + 1
  const colSeats = MIDDLE_EARTH_CFG.HEADERS.indexOf('# of Seats') + 1
  const colMeetings = MIDDLE_EARTH_CFG.HEADERS.indexOf('Meetings Recorded') + 1
  const colClients = MIDDLE_EARTH_CFG.HEADERS.indexOf('# of Clients') + 1
  const colPmDate = MIDDLE_EARTH_CFG.HEADERS.indexOf('PM Connected Date') + 1
  const colCals = MIDDLE_EARTH_CFG.HEADERS.indexOf('Cals Connected') + 1
  const colEmails = MIDDLE_EARTH_CFG.HEADERS.indexOf('Emails Connected') + 1

  sheet.getRange(startRow, 1, rowCount, MIDDLE_EARTH_CFG.HEADERS.length).setVerticalAlignment('middle')
  sheet.getRange(startRow, colDaysPing, rowCount, 1).setNumberFormat(MIDDLE_EARTH_CFG.INT_FMT)
  sheet.getRange(startRow, colSeats, rowCount, 1).setNumberFormat(MIDDLE_EARTH_CFG.INT_FMT)
  sheet.getRange(startRow, colMeetings, rowCount, 1).setNumberFormat(MIDDLE_EARTH_CFG.INT_FMT)
  sheet.getRange(startRow, colClients, rowCount, 1).setNumberFormat(MIDDLE_EARTH_CFG.INT_FMT)
  sheet.getRange(startRow, colPmDate, rowCount, 1).setNumberFormat(MIDDLE_EARTH_CFG.DATE_FMT)
  sheet.getRange(startRow, colCals, rowCount, 1).setNumberFormat(MIDDLE_EARTH_CFG.INT_FMT)
  sheet.getRange(startRow, colEmails, rowCount, 1).setNumberFormat(MIDDLE_EARTH_CFG.INT_FMT)

  const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build()
  sheet.getRange(startRow, colInOnboarding, rowCount, 1).setDataValidation(rule)
}

function MIDDLEEARTH_csvList_(v) {
  const s = MIDDLEEARTH_str_(v)
  if (!s) return []
  return s.split(',').map(x => MIDDLEEARTH_str_(x)).filter(Boolean)
}

function MIDDLEEARTH_readSheetObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1 || lastCol < 1) return []

  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => MIDDLEEARTH_str_(h))
  const rows = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()

  return rows.map(r => {
    const o = {}
    headers.forEach((h, i) => { o[h] = r[i] })
    return o
  })
}

function MIDDLEEARTH_batchSetValues_(sheet, startRow, startCol, values, chunkSize) {
  const size = Math.max(1, Number(chunkSize || 1000))
  for (let i = 0; i < values.length; i += size) {
    const chunk = values.slice(i, i + size)
    sheet.getRange(startRow + i, startCol, chunk.length, chunk[0].length).setValues(chunk)
  }
}

function MIDDLEEARTH_getOrCreateSheet_(ss, name) {
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function MIDDLEEARTH_lockWrap_(name, fn) {
  if (typeof lockWrap === 'function') {
    try { return lockWrap(name, fn) } catch (_) { return lockWrap(fn) }
  }
  return fn()
}

function MIDDLEEARTH_writeSyncLog_(step, status, rowsIn, rowsOut, seconds, error) {
  if (typeof writeSyncLog === 'function') {
    return writeSyncLog(step, status, rowsIn, rowsOut, seconds, error || '')
  }
  Logger.log(`[SYNCLOG missing] ${step} ${status} rows_in=${rowsIn} rows_out=${rowsOut} seconds=${seconds} error=${error || ''}`)
}

function MIDDLEEARTH_toBool_(v) {
  if (v === true) return true
  const s = MIDDLEEARTH_str_(v).toLowerCase()
  return s === 'true' || s === '1' || s === 'yes' || s === 'y'
}

function MIDDLEEARTH_toDateOrNull_(v) {
  if (!v && v !== 0) return null
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v
  const s = MIDDLEEARTH_str_(v)
  if (!s) return null
  const d = new Date(s)
  return isNaN(d.getTime()) ? null : d
}

function MIDDLEEARTH_str_(v) {
  if (v == null) return ''
  return String(v).trim()
}
