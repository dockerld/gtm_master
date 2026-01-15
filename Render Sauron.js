/**************************************************************
 * render_sauron_view()
 *
 * Renders your Master view sheet "Sauron" from canonical tables.
 *
 * ✅ Key features:
 * - Pulls Active Days (PostHog) from raw_posthog_user_metrics (email_key -> active_days)
 * - Preserves manual columns (Service, In Onboarding, Task Type, Create Task, etc)
 * - ✅ Service is overridden by org_info.Service (org-level manual) via Org ID
 * - ✅ Days since last Login now PREFERS canon_users.days_since_last_login
 *   (falls back to computing from canon_users.last_login_date if needed)
 * - ✅ DEBUG MODE: writes “seen/skipped” rows to a "Sauron Debug" sheet
 * - Auto-apply checkbox data validation to:
 *   In Onboarding, At-Risk, Email Connected, Cal Connected, Paying
 * - Auto-apply dropdown validations to:
 *   Create Task, Task Type, Service
 * - Clears formatting + data validation for the FULL data region (not just lastRow)
 **************************************************************/

const SAURON_CFG = {
  SHEET_NAME: 'Sauron',

  INPUTS: {
    CANON_USERS: 'canon_users',
    CANON_ORGS: 'canon_orgs',
    CLERK_MEMBERSHIPS: 'raw_clerk_memberships',
    CLERK_USERS_RAW: 'raw_clerk_users',
    POSTHOG_USERS_RAW: 'raw_posthog_user_metrics',
    ORG_INFO: 'org_info'
  },

  // IMPORTANT: this is the row index for the SAURON SHEET header row
  HEADER_ROW: 3,
  START_COL: 2,     // Column B
  DATA_START_ROW: 4,

  HEADERS: [
    'Email',
    'Name',
    'Org Name',
    'Service',
    'In Onboarding',
    'Days with Ping',
    'Days since last Login',
    'Meetings Recorded',
    'Logged in #',
    'Active Days (PostHog)',
    '# of Seats',
    '# of Clients',
    'Task Type',
    'Create Task',
    'Note',
    'Hands On',
    'Paying',
    'Sign Up Date',
    'Ask Meeting',
    'Ask Global',
    'Client Page Views',
    'Tests',
    'Last Log in',
    'PM',
    'PM Connected Date',
    'Cal Connected',
    'Cal Connected Date',
    'Email Connected',
    'Email Connected Date',
    'Hierarchy',
    'Consierge Email',
    'Tags',
    'Promo Code',
    'Status',
    'Org Sign Up Date'
  ],

  MANUAL_HEADERS: new Set([
    'Service',
    'In Onboarding',
    'Task Type',
    'Create Task',
    'Note',
    'Hands On',
    'Tests',
    'Hierarchy',
    'Consierge Email',
    'Tags',
    'Status'
  ]),

  ENABLE_EXTRA_COLUMNS: true,
  EXTRA_HEADERS: [
    'Org Members',
    'Activation Stage',
    'Activation Score',
    'At-Risk',
    'Risk Reason',
    'Activation Missing'
  ],

  CHECKBOX_HEADERS: [
    'In Onboarding',
    'At-Risk',
    'Email Connected',
    'Cal Connected',
    'Paying'
  ],

  DROPDOWN_HEADERS: {
    'Create Task': ['Docker', 'Camden'],
    'Task Type': [
      'onboarding',
      'reachout',
      'send video',
      'get clients',
      'white glove',
      'trial ending',
      'trial expired'
    ],
    'Service': ['White Glove', 'Hands On']
  }
}

// org_info columns (row 1 headers)
const SAURON_ORG_INFO_ORG_ID_HEADER = 'Org ID'
const SAURON_ORG_INFO_SERVICE_HEADER = 'Service'

/**
 * ✅ Debug controls
 * Turn ENABLED on to write debug rows to a sheet.
 * Set TARGET_DOMAIN to the domain you’re debugging.
 */
const SAURON_DEBUG = {
  ENABLED: true,
  TARGET_DOMAIN: 'firstpurposetax.com', // <-- change this when debugging other orgs
  SHEET_NAME: 'Sauron Debug',
  MAX_ROWS: 500
}

function render_sauron_view() {
  lockWrap('render_sauron_view', () => {
    const t0 = new Date()
    try {
      const ss = SpreadsheetApp.getActive()

      // quick visual proof you ran THIS version
      if (SAURON_DEBUG.ENABLED) {
        ss.toast('render_sauron_view v2026-01-02-DEBUG', 'Sauron', 5)
      }

      const sh = SAURON_getOrCreateSheet_(ss, SAURON_CFG.SHEET_NAME)

      const shUsers = ss.getSheetByName(SAURON_CFG.INPUTS.CANON_USERS)
      const shOrgs = ss.getSheetByName(SAURON_CFG.INPUTS.CANON_ORGS)
      const shMems = ss.getSheetByName(SAURON_CFG.INPUTS.CLERK_MEMBERSHIPS)
      const shClerkUsersRaw = ss.getSheetByName(SAURON_CFG.INPUTS.CLERK_USERS_RAW)
      const shPosthogUsersRaw = ss.getSheetByName(SAURON_CFG.INPUTS.POSTHOG_USERS_RAW)
      const shOrgInfo = ss.getSheetByName(SAURON_CFG.INPUTS.ORG_INFO)

      if (!shUsers) throw new Error(`Missing input sheet: ${SAURON_CFG.INPUTS.CANON_USERS}`)
      if (!shOrgs) throw new Error(`Missing input sheet: ${SAURON_CFG.INPUTS.CANON_ORGS}`)
      if (!shClerkUsersRaw) throw new Error(`Missing input sheet: ${SAURON_CFG.INPUTS.CLERK_USERS_RAW}`)
      if (!shPosthogUsersRaw) throw new Error(`Missing input sheet: ${SAURON_CFG.INPUTS.POSTHOG_USERS_RAW}`)
      if (!shOrgInfo) throw new Error(`Missing input sheet: ${SAURON_CFG.INPUTS.ORG_INFO}`)

      // DEBUG SHEET (optional)
      const debugSh = SAURON_debugReset_(ss)

      // Read canonical tables
      const users = SAURON_readSheetObjects_(shUsers, 1)
      const orgs = SAURON_readSheetObjects_(shOrgs, 1)

      const memCounts = shMems ? SAURON_buildOrgMemberCounts_(shMems) : new Map()

      const orgById = new Map()
      orgs.forEach(o => {
        const orgId = String(o.org_id || '').trim()
        if (!orgId) return
        orgById.set(orgId, o)
      })

      const payingByEmailKey = SAURON_buildPayingIndex_(shClerkUsersRaw)
      const activeDaysByEmailKey = SAURON_buildActiveDaysIndex_(shPosthogUsersRaw)

      // ✅ org-level Service index (Org ID -> Service)
      const serviceByOrgId = SAURON_buildOrgInfoServiceByOrgId_(shOrgInfo)

      // Preserve manual values from existing Sauron (per-email)
      const existingManualByEmail = SAURON_readExistingManual_(sh)

      const headers = SAURON_CFG.ENABLE_EXTRA_COLUMNS
        ? SAURON_CFG.HEADERS.concat(SAURON_CFG.EXTRA_HEADERS)
        : SAURON_CFG.HEADERS.slice()

      // basic header sanity
      if (SAURON_DEBUG.ENABLED) {
        const dupes = headers.filter((h, i) => headers.indexOf(h) !== i)
        if (dupes.length) {
          ss.toast('Duplicate headers: ' + dupes.join(', '), 'Sauron', 10)
        }
      }

      SAURON_writeHeaders_(sh, headers)

      const tz = Session.getScriptTimeZone()
      const todayYMD = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')

      const out = []
      const rowsIn = users.length

      users.forEach(u => {
        // Raw fields
        const emailRaw = String(u.email || '')
        const email = emailRaw.trim()
        const emailKeyRaw = String(u.email_key || '').trim()

        // tolerant key: never becomes empty just because normalizeEmail is picky
        const emailKey = SAURON_normEmail_(emailKeyRaw || email)

        // DEBUG: log any matching domain rows as "seen" early
        if (SAURON_matchesDebugDomain_(emailRaw) || SAURON_matchesDebugDomain_(emailKeyRaw)) {
          SAURON_debugLog_(debugSh, {
            reason: 'seen_in_loop',
            email_raw: emailRaw,
            email_key_raw: emailKeyRaw,
            emailKey_final: emailKey,
            org_id: String(u.org_id || '').trim(),
            name: String(u.name || '').trim(),
            row_hint: '',
            notes: 'before skip checks'
          })
        }

        if (!emailKey) {
          SAURON_debugLog_(debugSh, {
            reason: 'SKIP_empty_emailKey',
            email_raw: emailRaw,
            email_key_raw: emailKeyRaw,
            emailKey_final: emailKey,
            org_id: String(u.org_id || '').trim(),
            name: String(u.name || '').trim(),
            row_hint: '',
            notes: 'emailKey empty after normalization'
          })
          return
        }

        const orgId = String(u.org_id || '').trim()
        const org = orgById.get(orgId) || {}

        const priorManual = existingManualByEmail.get(emailKey) || {}

        const userSignUpDate = SAURON_asYMD_(u.created_at, tz) || ''
        const orgSignUpDate = SAURON_asYMD_(org.org_created_at, tz) || ''

        const daysWithPing =
          (u.days_with_ping != null && u.days_with_ping !== '') ? u.days_with_ping :
          (userSignUpDate ? SAURON_daysBetweenYMD_(userSignUpDate, todayYMD) : '')

        const lastLoginYMD = SAURON_asYMD_(u.last_login_date, tz) || ''

        // ✅ Prefer canon_users precomputed number; fallback to computing from last_login_date
        let daysSinceLastLogin = ''
        const canonDsl = (u.days_since_last_login != null && u.days_since_last_login !== '')
          ? Number(u.days_since_last_login)
          : NaN

        if (!isNaN(canonDsl) && isFinite(canonDsl) && canonDsl >= 0) {
          daysSinceLastLogin = canonDsl
        } else if (lastLoginYMD) {
          const diff = SAURON_daysBetweenYMD_(lastLoginYMD, todayYMD)
          if (typeof diff === 'number' && isFinite(diff) && diff >= 0) daysSinceLastLogin = diff
        }

        // ✅ Service precedence:
        // 1) org_info.Service (org-level manual)
        // 2) existing manual Service on Sauron (per-email, if present)
        // 3) canon_orgs.service fallback
        const orgInfoService = orgId ? (serviceByOrgId.get(orgId) || '') : ''
        const service = orgInfoService
          ? orgInfoService
          : SAURON_pickManualOrDefault_(priorManual, 'Service', String(org.service || '').trim())

        const dwp = Number(daysWithPing)
        const autoInOnboarding = (!isNaN(dwp) && isFinite(dwp) && dwp <= 14)
        const inOnboarding = SAURON_pickManualOrDefault_(priorManual, 'In Onboarding', autoInOnboarding)

        const taskType = SAURON_pickManualOrDefault_(priorManual, 'Task Type', '')
        const createTask = SAURON_pickManualOrDefault_(priorManual, 'Create Task', '')
        const note = SAURON_pickManualOrDefault_(priorManual, 'Note', '')
        const handsOn = SAURON_pickManualOrDefault_(priorManual, 'Hands On', '')
        const tests = SAURON_pickManualOrDefault_(priorManual, 'Tests', '')
        const hierarchy = SAURON_pickManualOrDefault_(priorManual, 'Hierarchy', '')
        const consiergeEmail = SAURON_pickManualOrDefault_(priorManual, 'Consierge Email', '')
        const tags = SAURON_pickManualOrDefault_(priorManual, 'Tags', '')
        const status = SAURON_pickManualOrDefault_(priorManual, 'Status', '')

        const paying = SAURON_toBool_(payingByEmailKey.get(emailKey) === true)

        const seats = (org.seats != null && org.seats !== '') ? org.seats : ''
        const promo = String(org.promo_code || '').trim()

        const clientsCount = (u.clients_count != null) ? u.clients_count : ''
        const meetingsRecorded = (u.meetings_recorded != null) ? u.meetings_recorded : ''
        const askMeeting = (u.ask_meeting != null) ? u.ask_meeting : ''
        const askGlobal = (u.ask_global != null) ? u.ask_global : ''
        const clientViews = (u.client_page_views != null) ? u.client_page_views : ''
        const loggedInDays = (u.logged_in_days_count != null) ? u.logged_in_days_count : ''

        const activeDaysPosthog = activeDaysByEmailKey.has(emailKey) ? activeDaysByEmailKey.get(emailKey) : ''

        const calConnected = SAURON_toBool_(u.calendar_connected)
        const calDate = SAURON_asYMD_(u.first_calendar_connected_date, tz) || ''

        const emailConnected = SAURON_toBool_(u.email_connected)
        const emailDate = SAURON_asYMD_(u.first_email_connected_date, tz) || ''

        const pmSummary = SAURON_buildPmSummary_(u)
        const pmDates = SAURON_buildPmDatesSummary_(u, tz)

        const baseRow = {
          'Email': email || emailKey,
          'Name': String(u.name || '').trim(),
          'Org Name': String(org.org_name || '').trim(),
          'Service': service,
          'In Onboarding': SAURON_toBool_(inOnboarding),
          'Days with Ping': daysWithPing,
          'Days since last Login': daysSinceLastLogin,
          'Meetings Recorded': meetingsRecorded,
          'Logged in #': loggedInDays,
          'Active Days (PostHog)': activeDaysPosthog,
          '# of Seats': seats,
          '# of Clients': clientsCount,
          'Task Type': taskType,
          'Create Task': createTask,
          'Note': note,
          'Hands On': handsOn,
          'Paying': paying,
          'Sign Up Date': userSignUpDate,
          'Ask Meeting': askMeeting,
          'Ask Global': askGlobal,
          'Client Page Views': clientViews,
          'Tests': tests,
          'Last Log in': lastLoginYMD,
          'PM': pmSummary,
          'PM Connected Date': pmDates,
          'Cal Connected': calConnected,
          'Cal Connected Date': calDate,
          'Email Connected': emailConnected,
          'Email Connected Date': emailDate,
          'Hierarchy': hierarchy,
          'Consierge Email': consiergeEmail,
          'Tags': tags,
          'Promo Code': promo,
          'Status': status,
          'Org Sign Up Date': orgSignUpDate
        }

        if (SAURON_CFG.ENABLE_EXTRA_COLUMNS) {
          const orgMembers = orgId ? (memCounts.get(orgId) || '') : ''

          const activation = SAURON_computeActivation_(baseRow)
          const atRisk = activation.score < 50
          const riskReason = atRisk ? `Activation Score < 50 (${activation.score})` : ''
          const missing = SAURON_activationMissing_(baseRow)

          baseRow['Org Members'] = orgMembers
          baseRow['Activation Stage'] = activation.stage
          baseRow['Activation Score'] = activation.score
          baseRow['At-Risk'] = atRisk
          baseRow['Risk Reason'] = riskReason
          baseRow['Activation Missing'] = missing
        }

        out.push(headers.map(h => (baseRow[h] !== undefined ? baseRow[h] : '')))
      })

      // Clear old values for full sheet height
      SAURON_clearOldData_(sh, headers.length)

      // Clear formatting + validations for full data region BEFORE writing
      SAURON_resetFormattingAndValidation_(sh, headers)

      // Write new values
      if (out.length) {
        SAURON_batchSetValues_(sh, SAURON_CFG.DATA_START_ROW, SAURON_CFG.START_COL, out, 3000)
      }

      // Apply validations down the full data region
      SAURON_applyCheckboxes_(sh, headers, SAURON_CFG.CHECKBOX_HEADERS)
      SAURON_applyDropdowns_(sh, headers, SAURON_CFG.DROPDOWN_HEADERS)

      // Formats
      SAURON_applyDateFormat_(sh, headers, 'Last Log in', 'yyyy-mm-dd')
      SAURON_applyActivationScoreGradient_(sh, headers, out.length)

      sh.setFrozenRows(SAURON_CFG.HEADER_ROW)
      sh.autoResizeColumns(SAURON_CFG.START_COL, headers.length)

      writeSyncLog(
        'render_sauron_view',
        'ok',
        rowsIn,
        out.length,
        (new Date() - t0) / 1000,
        ''
      )
    } catch (err) {
      writeSyncLog('render_sauron_view', 'error', '', '', '', String(err && err.message ? err.message : err))
      throw err
    }
  })
}

/* =========================
 * DEBUG HELPERS
 * ========================= */

function SAURON_getOrCreateSheetByName_(ss, name) {
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function SAURON_debugReset_(ss) {
  if (!SAURON_DEBUG.ENABLED) return null
  const sh = SAURON_getOrCreateSheetByName_(ss, SAURON_DEBUG.SHEET_NAME)
  sh.clear()
  sh.getRange(1, 1, 1, 8).setValues([[
    'reason',
    'email_raw',
    'email_key_raw',
    'emailKey_final',
    'org_id',
    'name',
    'canon_users_row_hint',
    'notes'
  ]])
  return sh
}

function SAURON_debugLog_(debugSh, rowObj) {
  if (!SAURON_DEBUG.ENABLED || !debugSh) return
  const lastRow = debugSh.getLastRow()
  if (lastRow >= SAURON_DEBUG.MAX_ROWS + 1) return
  debugSh.appendRow([
    rowObj.reason || '',
    rowObj.email_raw || '',
    rowObj.email_key_raw || '',
    rowObj.emailKey_final || '',
    rowObj.org_id || '',
    rowObj.name || '',
    rowObj.row_hint || '',
    rowObj.notes || ''
  ])
}

function SAURON_matchesDebugDomain_(email) {
  if (!SAURON_DEBUG.TARGET_DOMAIN) return false
  const s = String(email || '').toLowerCase()
  return s.includes(SAURON_DEBUG.TARGET_DOMAIN.toLowerCase())
}

/* =========================
 * Helpers
 * ========================= */

function SAURON_getOrCreateSheet_(ss, name) {
  if (typeof getOrCreateSheet === 'function') {
    try { return getOrCreateSheet(ss, name) } catch (e) {}
    try { return getOrCreateSheet(name) } catch (e) {}
  }
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function SAURON_readSheetObjects_(sheet, headerRow) {
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
      obj[SAURON_key_(h)] = r[i]
    })
    return obj
  })
}

function SAURON_key_(h) {
  return String(h || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_')
}

function SAURON_writeHeaders_(sheet, headers) {
  sheet.getRange(SAURON_CFG.HEADER_ROW, SAURON_CFG.START_COL, 1, headers.length).setValues([headers])
}

function SAURON_getDataRowCount_(sheet) {
  const maxRows = sheet.getMaxRows()
  const startRow = SAURON_CFG.DATA_START_ROW
  return Math.max(0, maxRows - startRow + 1)
}

function SAURON_clearOldData_(sheet, numCols) {
  const startRow = SAURON_CFG.DATA_START_ROW
  const startCol = SAURON_CFG.START_COL
  const maxRows = sheet.getMaxRows()
  const numRows = Math.max(0, maxRows - startRow + 1)
  if (!numRows) return
  sheet.getRange(startRow, startCol, numRows, numCols).clearContent()
}

function SAURON_readExistingManual_(sheet) {
  const headerRow = SAURON_CFG.HEADER_ROW
  const startCol = SAURON_CFG.START_COL
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1 || lastCol < startCol) return new Map()

  const header = sheet.getRange(headerRow, startCol, 1, lastCol - startCol + 1).getValues()[0]
    .map(h => String(h || '').trim())

  const trimmedHeader = header.filter(h => String(h || '').trim() !== '')
  const emailIdx = trimmedHeader.findIndex(h => String(h).toLowerCase() === 'email')
  if (emailIdx < 0) return new Map()

  const numCols = trimmedHeader.length
  const data = sheet.getRange(headerRow + 1, startCol, lastRow - headerRow, numCols).getValues()

  const out = new Map()
  data.forEach(row => {
    const email = String(row[emailIdx] || '').trim()
    const key = SAURON_normEmail_(email)
    if (!key) return

    const manual = {}
    trimmedHeader.forEach((h, i) => {
      if (!SAURON_CFG.MANUAL_HEADERS.has(h)) return
      manual[h] = row[i]
    })
    out.set(key, manual)
  })
  return out
}

function SAURON_pickManualOrDefault_(priorManual, headerName, fallback) {
  if (priorManual && Object.prototype.hasOwnProperty.call(priorManual, headerName)) {
    return priorManual[headerName]
  }
  return fallback
}

function SAURON_asYMD_(value, tz) {
  if (!value) return ''
  if (value instanceof Date) return Utilities.formatDate(value, tz, 'yyyy-MM-dd')

  const s = String(value).trim()
  if (!s) return ''

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s

  if (/^\d+$/.test(s)) {
    const n = Number(s)
    const ms = n > 1e12 ? n : n * 1000
    const d = new Date(ms)
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'yyyy-MM-dd')
  }

  const d = new Date(s)
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'yyyy-MM-dd')

  return ''
}

function SAURON_daysBetweenYMD_(ymdA, ymdB) {
  const a = new Date(ymdA + 'T00:00:00Z')
  const b = new Date(ymdB + 'T00:00:00Z')
  return Math.floor((b.getTime() - a.getTime()) / (1000 * 60 * 60 * 24))
}

function SAURON_toBool_(v) {
  if (v === true) return true
  if (v === false) return false
  const s = String(v || '').toLowerCase().trim()
  return s === 'yes' || s === 'true' || s === '1'
}

function SAURON_buildPmSummary_(u) {
  const providers = []
  if (SAURON_toBool_(u.pm_karbon_connected)) providers.push('KARBON')
  if (SAURON_toBool_(u.pm_keeper_connected)) providers.push('KEEPER')
  if (SAURON_toBool_(u.pm_financial_cents_connected)) providers.push('FINANCIAL_CENTS')
  return providers.join(', ')
}

function SAURON_buildPmDatesSummary_(u, tz) {
  const parts = []
  if (SAURON_toBool_(u.pm_karbon_connected)) parts.push(SAURON_asYMD_(u.pm_karbon_first_connected_date, tz))
  if (SAURON_toBool_(u.pm_keeper_connected)) parts.push(SAURON_asYMD_(u.pm_keeper_first_connected_date, tz))
  if (SAURON_toBool_(u.pm_financial_cents_connected)) parts.push(SAURON_asYMD_(u.pm_financial_cents_first_connected_date, tz))
  return parts.filter(Boolean).join(', ')
}

function SAURON_buildOrgMemberCounts_(sheet) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < 2) return new Map()

  const { map } = readHeaderMap(sheet, 1)
  const cOrg = map['org_id']
  const cEmail = map['email']
  if (!cOrg) return new Map()

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
  const orgToSet = new Map()

  data.forEach(r => {
    const orgId = String(r[cOrg - 1] || '').trim()
    if (!orgId) return
    const email = cEmail ? String(r[cEmail - 1] || '').trim() : ''
    const key = SAURON_normEmail_(email) || email
    if (!orgToSet.has(orgId)) orgToSet.set(orgId, new Set())
    if (key) orgToSet.get(orgId).add(key)
  })

  const counts = new Map()
  orgToSet.forEach((set, orgId) => counts.set(orgId, set.size))
  return counts
}

function SAURON_computeActivation_(row) {
  let score = 0
  let stage = 'New'

  const cal = row['Cal Connected'] === true
  const mail = row['Email Connected'] === true
  const pm = String(row['PM'] || '').trim().length > 0
  const meetings = Number(row['Meetings Recorded'] || 0) || 0
  const clients = Number(row['# of Clients'] || 0) || 0

  if (cal) score += 20
  if (mail) score += 20
  if (pm) score += 15
  if (clients > 0) score += 15
  if (meetings > 0) score += 30

  if (meetings > 0) stage = 'Recorded Meeting'
  else if (clients > 0) stage = 'Added Client'
  else if (pm) stage = 'Connected PM'
  else if (cal || mail) stage = 'Connected'
  else stage = 'New'

  return { score, stage }
}

function SAURON_activationMissing_(row) {
  const missing = []
  if (row['Cal Connected'] !== true) missing.push('calendar')
  if (row['Email Connected'] !== true) missing.push('email')
  if (!String(row['PM'] || '').trim()) missing.push('pm')
  if (!(Number(row['# of Clients'] || 0) > 0)) missing.push('clients')
  if (!(Number(row['Meetings Recorded'] || 0) > 0)) missing.push('meeting')
  return missing.join(', ')
}

function SAURON_buildPayingIndex_(rawClerkUsersSheet) {
  const lastRow = rawClerkUsersSheet.getLastRow()
  const lastCol = rawClerkUsersSheet.getLastColumn()
  const out = new Map()
  if (lastRow < 2) return out

  const { map } = readHeaderMap(rawClerkUsersSheet, 1)
  const cEmailKey = map['email_key']
  const cStripeSub = map['stripe_subscription_id']

  if (!cEmailKey || !cStripeSub) {
    throw new Error('raw_clerk_users must have headers: email_key, stripe_subscription_id')
  }

  const data = rawClerkUsersSheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
  data.forEach(r => {
    const emailKey = String(r[cEmailKey - 1] || '').trim()
    const stripeSub = String(r[cStripeSub - 1] || '').trim()
    if (!emailKey) return
    out.set(emailKey, !!stripeSub)
  })

  return out
}

function SAURON_buildActiveDaysIndex_(rawPosthogSheet) {
  const lastRow = rawPosthogSheet.getLastRow()
  const lastCol = rawPosthogSheet.getLastColumn()
  const out = new Map()
  if (lastRow < 2) return out

  const { map } = readHeaderMap(rawPosthogSheet, 1)
  const cEmailKey = map['email_key']
  const cActiveDays = map['active_days']

  if (!cEmailKey || !cActiveDays) {
    throw new Error('raw_posthog_user_metrics must have headers: email_key, active_days')
  }

  const data = rawPosthogSheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
  data.forEach(r => {
    const emailKey = String(r[cEmailKey - 1] || '').trim()
    const activeDays = r[cActiveDays - 1]
    if (!emailKey) return
    out.set(emailKey, activeDays != null && activeDays !== '' ? Number(activeDays) : '')
  })

  return out
}

function SAURON_applyActivationScoreGradient_(sheet, headers, dataRowCount) {
  const n = Number(dataRowCount || 0)
  if (!n) return

  const idx = headers.findIndex(h => String(h).trim() === 'Activation Score')
  if (idx < 0) return

  const col = SAURON_CFG.START_COL + idx
  const range = sheet.getRange(SAURON_CFG.DATA_START_ROW, col, n, 1)

  const rules = sheet.getConditionalFormatRules() || []
  const filtered = rules.filter(rule => {
    try {
      const rs = rule.getRanges()
      if (!rs || !rs.length) return true
      return !rs.some(r => r.getColumn() === col)
    } catch (e) {
      return true
    }
  })

  const gradientRule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([range])
    .setGradientMinpointWithValue('#f4c7c3', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMidpointWithValue('#fff2cc', SpreadsheetApp.InterpolationType.NUMBER, '50')
    .setGradientMaxpointWithValue('#d9ead3', SpreadsheetApp.InterpolationType.NUMBER, '100')
    .build()

  sheet.setConditionalFormatRules(filtered.concat([gradientRule]))
}

function SAURON_applyCheckboxes_(sheet, headers, checkboxHeaders) {
  const startRow = SAURON_CFG.DATA_START_ROW
  const numRows = sheet.getMaxRows() - startRow + 1
  if (numRows <= 0) return

  const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build()

  ;(checkboxHeaders || []).forEach(hName => {
    const idx = headers.findIndex(h => String(h).trim() === String(hName).trim())
    if (idx < 0) return
    const col = SAURON_CFG.START_COL + idx
    sheet.getRange(startRow, col, numRows, 1).setDataValidation(rule)
  })
}

function SAURON_applyDropdowns_(sheet, headers, dropdownHeadersMap) {
  const startRow = SAURON_CFG.DATA_START_ROW
  const numRows = sheet.getMaxRows() - startRow + 1
  if (numRows <= 0) return

  Object.keys(dropdownHeadersMap || {}).forEach(headerName => {
    const options = (dropdownHeadersMap[headerName] || []).map(v => String(v))
    if (!options.length) return

    const idx = headers.findIndex(h => String(h).trim() === String(headerName).trim())
    if (idx < 0) return

    const col = SAURON_CFG.START_COL + idx
    const list = [''].concat(options)

    const rule = SpreadsheetApp
      .newDataValidation()
      .requireValueInList(list, true)
      .setAllowInvalid(true)
      .build()

    sheet.getRange(startRow, col, numRows, 1).setDataValidation(rule)
  })
}

function SAURON_applyDateFormat_(sheet, headers, headerName, numberFormat) {
  const startRow = SAURON_CFG.DATA_START_ROW
  const numRows = sheet.getMaxRows() - startRow + 1
  if (numRows <= 0) return

  const idx = headers.findIndex(h => String(h).trim() === String(headerName).trim())
  if (idx < 0) return

  const col = SAURON_CFG.START_COL + idx
  sheet.getRange(startRow, col, numRows, 1).setNumberFormat(numberFormat)
}

function SAURON_resetFormattingAndValidation_(sheet, headers) {
  const numRows = SAURON_getDataRowCount_(sheet)
  if (!numRows) return

  const numCols = headers.length
  const range = sheet.getRange(SAURON_CFG.DATA_START_ROW, SAURON_CFG.START_COL, numRows, numCols)

  range.clearDataValidations()
  range.clearFormat()
}

function SAURON_batchSetValues_(sheet, startRow, startCol, values, chunkSize) {
  if (!values || !values.length) return

  const size = chunkSize || 2000
  for (let i = 0; i < values.length; i += size) {
    const chunk = values.slice(i, i + size)
    sheet
      .getRange(startRow + i, startCol, chunk.length, chunk[0].length)
      .setValues(chunk)
  }
}

/**
 * Email normalization that will NEVER blank out rows.
 * - Tries shared normalizeEmail if present
 * - Fallback: lowercase raw
 */
function SAURON_normEmail_(v) {
  const raw = String(v || '').trim()
  if (!raw) return ''

  if (typeof normalizeEmail === 'function') {
    try {
      const out = normalizeEmail(raw)
      if (out && String(out).trim()) return String(out).trim().toLowerCase()
    } catch (e) {}
  }

  return raw.toLowerCase()
}

/**
 * ✅ org_info -> Service index
 * Expects org_info headers row 1 including:
 * - "Org ID"
 * - "Service"
 */
function SAURON_buildOrgInfoServiceByOrgId_(sheet) {
  const out = new Map()

  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < 2 || lastCol < 1) return out

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim())

  const orgIdIdx = header.findIndex(h => h.toLowerCase() === String(SAURON_ORG_INFO_ORG_ID_HEADER).toLowerCase())
  const serviceIdx = header.findIndex(h => h.toLowerCase() === String(SAURON_ORG_INFO_SERVICE_HEADER).toLowerCase())

  if (orgIdIdx < 0 || serviceIdx < 0) {
    throw new Error(`org_info must have headers: "${SAURON_ORG_INFO_ORG_ID_HEADER}", "${SAURON_ORG_INFO_SERVICE_HEADER}"`)
  }

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
  data.forEach(r => {
    const orgId = String(r[orgIdIdx] || '').trim()
    if (!orgId) return
    const service = String(r[serviceIdx] || '').trim()
    if (!service) return
    out.set(orgId, service)
  })

  return out
}