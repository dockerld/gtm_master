/**************************************************************
 * Clerk (ONE FILE) — SLIMMED (UPDATED)
 * ------------------------------------------------------------
 * A) Clerk Raw Sync (overwrite-only):
 *   - raw_clerk_users  (now includes: org_id, org_role, last_login_date, login_count)
 *   - raw_clerk_orgs   (now includes: org_owner_user_id)
 *   - raw_clerk_memberships
 *
 * B) KEEP Login History (append-only):
 *   - login_events   (append-only)
 *
 * Script Property required:
 *   - CLERK_SECRET_KEY
 **************************************************************/

/* =========================
 * Config
 * ========================= */

const CLERK_CFG = {
  API_BASE: 'https://api.clerk.com/v1',
  USERS_ENDPOINT: '/users',
  ORGS_ENDPOINT: '/organizations',
  MEMBERSHIPS_ENDPOINT: '/organization_memberships',

  PAGE_SIZE: 100,
  PAUSE_MS: 200,
  WRITE_CHUNK: 2000,

  // Raw tables
  RAW: {
    USERS: 'raw_clerk_users',
    ORGS: 'raw_clerk_orgs',
    MEMBERSHIPS: 'raw_clerk_memberships'
  },

  // Login history tables
  LOGIN: {
    EVENTS_TAB: 'login_events',
    SOURCE: 'clerk_poll',
    TIME_FMT: 'yyyy-MM-dd HH:mm:ss',
    DATE_FMT: 'yyyy-MM-dd'
  }
}

// In-memory cache per execution (so pipeline steps reuse memberships fetch)
let _CLERK_MEMBERSHIPS_CACHE = null

/* =========================
 * Public entrypoints
 * ========================= */

function clerk_pull_all_raw() {
  lockWrapSafe_('clerk_pull_all_raw', () => {
    // We fetch memberships inside users/orgs anyway (cached), but this keeps the sheet fresh too.
    clerk_pull_memberships_to_raw()
    clerk_pull_orgs_to_raw()
    clerk_pull_users_to_raw()
  })
}

/**
 * Append-only: login_events
 * Dedupe per user based on latest recorded timestamp in login_events
 */
function syncClerkUsers() {
  const t0 = new Date()
  const apiKey = clerkGetSecret_()

  const headers = {
    Authorization: `Bearer ${apiKey}`,
    'Content-Type': 'application/json'
  }

  const ss = SpreadsheetApp.getActive()
  const events = getOrCreateSheetSafe_(ss, CLERK_CFG.LOGIN.EVENTS_TAB)

  ensureHeaders_(events, [
    'user_id',
    'email',
    'login_timestamp',
    'login_date',
    'source'
  ])

  // latest recorded login per user_id from login_events
  const latestByUserId = clerkBuildLatestLoginByUserId_(events)

  // Pull all Clerk users
  const users = clerkFetchAllUsersForLogin_(headers)

  const tz = Session.getScriptTimeZone()
  const newEventRows = []

  users.forEach(u => {
    const userId = String(u.id || '').trim()
    if (!userId) return

    const email = clerkPrimaryEmail_(u)
    const lastSignInDate = clerkToDate_(u.last_sign_in_at)
    if (!email || !lastSignInDate) return

    const lastSignInMs = lastSignInDate.getTime()
    const lastRecordedMs = latestByUserId.get(userId) || 0

    if (lastSignInMs > lastRecordedMs) {
      const loginTsStr = Utilities.formatDate(lastSignInDate, tz, CLERK_CFG.LOGIN.TIME_FMT)
      const loginDateStr = Utilities.formatDate(lastSignInDate, tz, CLERK_CFG.LOGIN.DATE_FMT)
      newEventRows.push([userId, email, loginTsStr, loginDateStr, CLERK_CFG.LOGIN.SOURCE])
      latestByUserId.set(userId, lastSignInMs)
    }
  })

  if (newEventRows.length) {
    events
      .getRange(events.getLastRow() + 1, 1, newEventRows.length, newEventRows[0].length)
      .setValues(newEventRows)
  }

  const seconds = ((new Date()) - t0) / 1000
  writeSyncLogSafe_('syncClerkUsers', 'ok', users.length, newEventRows.length, seconds, '')
  return { rows_in: users.length, rows_out: newEventRows.length }
}

/**
 * A) Pull all Clerk users into raw_clerk_users (overwrite)
 * Adds:
 *  - org_id, org_role (derived from memberships, preferring owner/admin)
 *  - last_login_date (yyyy-MM-dd) from Clerk last_sign_in_at
 *  - login_count (count of login_events for that user_id)
 */
function clerk_pull_users_to_raw() {
  const t0 = new Date()
  const apiKey = clerkGetSecret_()

  const ss = SpreadsheetApp.getActive()
  const sh = getOrCreateSheetSafe_(ss, CLERK_CFG.RAW.USERS)

  const headers = [
    'clerk_user_id',
    'email',
    'email_key',
    'name',
    'created_at',

    // ✅ NEW: org + role from memberships
    'org_id',
    'org_role',

    'last_sign_in_at',

    // existing
    'last_login_date',
    'login_count',

    // Stripe bridge + plan info from Clerk private metadata
    'stripe_customer_id',
    'stripe_subscription_id',
    'subscription_status',
    'current_plan',
    'subscription_tier',
    'trial_start_date',
    'trial_ends_at',
    'subscription_ends_at',

    // Helpful raw status flags (optional)
    'last_active_at',
    'banned',
    'locked',
    'two_factor_enabled',

    // Debug: raw private meta blob (optional)
    'private_meta_json'
  ]

  // Build login count map from login_events
  const loginCountsByUserId = clerkBuildLoginCountsByUserId_()

  // ✅ Build membership index (user_id -> {org_id, role})
  const memberships = clerkGetAllMembershipsCached_(apiKey)
  const membershipByUserId = clerkBuildUserMembershipIndex_(memberships)

  const users = clerkFetchAll_(`${CLERK_CFG.API_BASE}${CLERK_CFG.USERS_ENDPOINT}`, apiKey, true)
  const tz = Session.getScriptTimeZone()

  const rows = users.map(u => {
    const userId = strOrBlank_(u.id)
    const email = clerkPrimaryEmail_(u)
    const name = clerkFullName_(u)

    const createdAt = clerkToIso_(u.created_at)

    // ✅ org_id + org_role from membership index
    const mem = membershipByUserId.get(userId) || { org_id: '', role: '' }
    const orgId = strOrBlank_(mem.org_id)
    const orgRole = strOrBlank_(mem.role)

    const lastSignInAt = clerkToIso_(u.last_sign_in_at)

    // last_login_date: yyyy-MM-dd from last_sign_in_at
    let lastLoginDate = ''
    const d = clerkToDate_(u.last_sign_in_at)
    if (d) lastLoginDate = Utilities.formatDate(d, tz, 'yyyy-MM-dd')

    // login_count from login_events by user_id
    const loginCount = loginCountsByUserId.get(userId) || 0

    const priv = clerkPrivateMeta_(u)

    const stripeCustomerId = strOrBlank_(priv.stripeCustomerId)
    const stripeSubscriptionId = strOrBlank_(priv.stripeSubscriptionId)

    const subscriptionStatus = strOrBlank_(priv.subscriptionStatus)
    const currentPlan = strOrBlank_(priv.currentPlan)
    const subscriptionTier = strOrBlank_(priv.subscriptionTier)
    const trialStartDate = strOrBlank_(priv.trialStartDate)
    const trialEndsAt = strOrBlank_(priv.trialEndsAt)
    const subscriptionEndsAt = strOrBlank_(priv.subscriptionEndsAt)

    return [
      userId,
      email,
      normalizeEmailSafe_(email),
      name,
      createdAt,

      orgId,
      orgRole,

      lastSignInAt,

      lastLoginDate,
      loginCount,

      stripeCustomerId,
      stripeSubscriptionId,
      subscriptionStatus,
      currentPlan,
      subscriptionTier,
      trialStartDate,
      trialEndsAt,
      subscriptionEndsAt,

      clerkToIso_(u.last_active_at),
      u.banned === true,
      u.locked === true,
      u.two_factor_enabled === true,

      safeJson_(priv)
    ]
  })

  clerkOverwriteSheet_(sh, headers, rows)

  const seconds = (new Date() - t0) / 1000
  writeSyncLogSafe_('clerk_pull_users_to_raw', 'ok', users.length, rows.length, seconds, '')
  return { rows_in: users.length, rows_out: rows.length }
}

/* =========================
 * Orgs + Memberships
 * ========================= */

function clerk_pull_orgs_to_raw() {
  const t0 = new Date()
  const apiKey = clerkGetSecret_()

  const ss = SpreadsheetApp.getActive()
  const shOrgs = getOrCreateSheetSafe_(ss, CLERK_CFG.RAW.ORGS)

  // We compute members_count from raw_clerk_memberships (since /organizations does not return it)
  const shMems = ss.getSheetByName(CLERK_CFG.RAW.MEMBERSHIPS)

  const headers = [
    'org_id',
    'org_name',
    'org_slug',
    'created_at',
    'updated_at',
    'members_count',
    'org_owner_user_id' // ✅ NEW
  ]

  // 1) Pull orgs from Clerk
  const orgs = clerkFetchAll_(`${CLERK_CFG.API_BASE}${CLERK_CFG.ORGS_ENDPOINT}`, apiKey, true)

  // 2) Build org_id -> members_count using memberships sheet (preferred)
  const membersCountByOrgId = buildOrgMemberCountsFromMembershipsSheet_(shMems)

  // 3) Build org_id -> owner_user_id using memberships (owner > admin > manager > member)
  // Prefer cached memberships if available (fast + consistent with your other steps)
  const memberships = clerkGetAllMembershipsCached_(apiKey)
  const orgOwnerByOrgId = clerkBuildOrgOwnerIndex_(memberships) // org_id -> user_id

  // 4) Write rows
  const rows = orgs.map(o => {
    const orgId = strOrBlank_(o.id)
    const membersCount = membersCountByOrgId.get(orgId) ?? 0
    const ownerUserId = orgOwnerByOrgId.get(orgId) || ''

    return [
      orgId,
      strOrBlank_(o.name),
      strOrBlank_(o.slug),
      clerkToIso_(o.created_at),
      clerkToIso_(o.updated_at),
      Number(membersCount) || 0,
      ownerUserId
    ]
  })

  clerkOverwriteSheet_(shOrgs, headers, rows)

  const seconds = (new Date() - t0) / 1000
  writeSyncLogSafe_(
    'clerk_pull_orgs_to_raw',
    'ok',
    orgs.length,
    rows.length,
    seconds,
    'members_count computed from raw_clerk_memberships; org_owner_user_id computed from memberships'
  )
  return { rows_in: orgs.length, rows_out: rows.length }
}

/**
 * Computes org member counts from raw_clerk_memberships.
 * Counts UNIQUE clerk_user_id per org_id (fallback to email_key if needed).
 */
function buildOrgMemberCountsFromMembershipsSheet_(shMems) {
  const out = new Map()
  if (!shMems || shMems.getLastRow() < 2) return out

  const lastRow = shMems.getLastRow()
  const lastCol = shMems.getLastColumn()

  const header = shMems.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || '').trim().toLowerCase())

  const cOrgId = header.indexOf('org_id') + 1
  const cUserId = header.indexOf('clerk_user_id') + 1
  const cEmailKey = header.indexOf('email_key') + 1

  if (!cOrgId) return out

  const data = shMems.getRange(2, 1, lastRow - 1, lastCol).getValues()
  const orgToSet = new Map() // org_id -> Set(unique members)

  data.forEach(r => {
    const orgId = String(r[cOrgId - 1] || '').trim()
    if (!orgId) return

    // Prefer clerk_user_id; fallback to email_key; fallback to email-ish
    let memberKey = ''
    if (cUserId) memberKey = String(r[cUserId - 1] || '').trim()
    if (!memberKey && cEmailKey) memberKey = String(r[cEmailKey - 1] || '').trim()

    if (!memberKey) return

    if (!orgToSet.has(orgId)) orgToSet.set(orgId, new Set())
    orgToSet.get(orgId).add(memberKey)
  })

  orgToSet.forEach((set, orgId) => out.set(orgId, set.size))
  return out
}

function clerk_pull_memberships_to_raw() {
  const t0 = new Date()
  const apiKey = clerkGetSecret_()

  const ss = SpreadsheetApp.getActive()
  const sh = getOrCreateSheetSafe_(ss, CLERK_CFG.RAW.MEMBERSHIPS)

  const headers = [
    'org_id',
    'org_name',
    'clerk_user_id',
    'email',
    'email_key',
    'role',
    'created_at',
    'updated_at'
  ]

  let memberships = []
  let usedFallback = false

  try {
    memberships = clerkFetchAll_(`${CLERK_CFG.API_BASE}${CLERK_CFG.MEMBERSHIPS_ENDPOINT}`, apiKey, true)
  } catch (e) {
    usedFallback = true
    memberships = clerkFetchMembershipsByOrgFallback_(apiKey)
  }

  // cache for this execution
  _CLERK_MEMBERSHIPS_CACHE = memberships

  const orgNameById = clerkGetOrgNameMap_()

  const rows = memberships.map(m => {
    const orgId =
      (m.organization && m.organization.id) ||
      m.organization_id ||
      ''

    const orgName =
      (m.organization && (m.organization.name || m.organization.slug)) ||
      orgNameById[orgId] ||
      ''

    const userId =
      (m.public_user_data && m.public_user_data.user_id) ||
      m.user_id ||
      (m.user && m.user.id) ||
      ''

    const email =
      (m.public_user_data && m.public_user_data.identifier) ||
      (m.public_user_data && m.public_user_data.email_address) ||
      ''

    const role = m.role || (m.public_user_data && m.public_user_data.role) || ''

    return [
      strOrBlank_(orgId),
      strOrBlank_(orgName),
      strOrBlank_(userId),
      strOrBlank_(email),
      normalizeEmailSafe_(email),
      strOrBlank_(role),
      clerkToIso_(m.created_at),
      clerkToIso_(m.updated_at)
    ]
  })

  clerkOverwriteSheet_(sh, headers, rows)

  const seconds = (new Date() - t0) / 1000
  writeSyncLogSafe_(
    'clerk_pull_memberships_to_raw',
    'ok',
    memberships.length,
    rows.length,
    seconds,
    usedFallback ? 'used fallback: per-org memberships' : ''
  )

  return { rows_in: memberships.length, rows_out: rows.length }
}

/* =========================
 * Membership indices
 * ========================= */

function clerkGetAllMembershipsCached_(apiKey) {
  if (Array.isArray(_CLERK_MEMBERSHIPS_CACHE)) return _CLERK_MEMBERSHIPS_CACHE

  // Prefer reading from sheet if it exists and has rows (fast), else fetch from API
  try {
    const ss = SpreadsheetApp.getActive()
    const sh = ss.getSheetByName(CLERK_CFG.RAW.MEMBERSHIPS)
    if (sh && sh.getLastRow() >= 2) {
      const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
        .map(h => String(h || '').trim().toLowerCase())

      const cOrg = header.indexOf('org_id') + 1
      const cUser = header.indexOf('clerk_user_id') + 1
      const cRole = header.indexOf('role') + 1
      if (cOrg && cUser && cRole) {
        const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
        const pseudo = data.map(r => ({
          organization_id: r[cOrg - 1],
          user_id: r[cUser - 1],
          role: r[cRole - 1]
        }))
        _CLERK_MEMBERSHIPS_CACHE = pseudo
        return pseudo
      }
    }
  } catch (e) {}

  // Fallback: fetch from API
  try {
    _CLERK_MEMBERSHIPS_CACHE = clerkFetchAll_(`${CLERK_CFG.API_BASE}${CLERK_CFG.MEMBERSHIPS_ENDPOINT}`, apiKey, true)
  } catch (e) {
    _CLERK_MEMBERSHIPS_CACHE = clerkFetchMembershipsByOrgFallback_(apiKey)
  }
  return _CLERK_MEMBERSHIPS_CACHE
}

function clerkRolePriority_(role) {
  const r = String(role || '').toLowerCase()
  if (r.includes('owner')) return 1
  if (r.includes('admin')) return 2
  if (r.includes('manager')) return 3
  if (r.includes('member')) return 4
  if (r.includes('user')) return 5
  return 9
}

/**
 * user_id -> best membership {org_id, role}
 * If a user has multiple orgs, this picks deterministically:
 * owner/admin first, else lowest org_id.
 */
function clerkBuildUserMembershipIndex_(memberships) {
  const byUser = new Map() // user_id -> array of {org_id, role}

  ;(memberships || []).forEach(m => {
    const orgId =
      (m.organization && m.organization.id) ||
      m.organization_id ||
      ''
    const userId =
      (m.public_user_data && m.public_user_data.user_id) ||
      m.user_id ||
      (m.user && m.user.id) ||
      ''
    const role = m.role || (m.public_user_data && m.public_user_data.role) || ''

    const o = String(orgId || '').trim()
    const u = String(userId || '').trim()
    if (!o || !u) return

    if (!byUser.has(u)) byUser.set(u, [])
    byUser.get(u).push({ org_id: o, role: String(role || '') })
  })

  const out = new Map()
  byUser.forEach((arr, userId) => {
    const best = arr.slice().sort((a, b) => {
      const pa = clerkRolePriority_(a.role)
      const pb = clerkRolePriority_(b.role)
      if (pa !== pb) return pa - pb
      return String(a.org_id).localeCompare(String(b.org_id))
    })[0]
    out.set(userId, best || { org_id: '', role: '' })
  })

  return out
}

/**
 * org_id -> owner user_id
 * Picks first owner; if none, first admin; else blank.
 */
function clerkBuildOrgOwnerIndex_(memberships) {
  const candidates = new Map() // org_id -> array of {user_id, role}

  ;(memberships || []).forEach(m => {
    const orgId =
      (m.organization && m.organization.id) ||
      m.organization_id ||
      ''
    const userId =
      (m.public_user_data && m.public_user_data.user_id) ||
      m.user_id ||
      (m.user && m.user.id) ||
      ''
    const role = m.role || (m.public_user_data && m.public_user_data.role) || ''

    const o = String(orgId || '').trim()
    const u = String(userId || '').trim()
    if (!o || !u) return

    if (!candidates.has(o)) candidates.set(o, [])
    candidates.get(o).push({ user_id: u, role: String(role || '') })
  })

  const out = new Map()
  candidates.forEach((arr, orgId) => {
    const sorted = arr.slice().sort((a, b) => {
      const pa = clerkRolePriority_(a.role)
      const pb = clerkRolePriority_(b.role)
      if (pa !== pb) return pa - pb
      return String(a.user_id).localeCompare(String(b.user_id))
    })
    const best = sorted[0]
    out.set(orgId, best ? best.user_id : '')
  })

  return out
}

/* =========================
 * Clerk API helpers
 * ========================= */

function clerkGetSecret_() {
  const key = PropertiesService.getScriptProperties().getProperty('CLERK_SECRET_KEY')
  if (!key) throw new Error('Missing CLERK_SECRET_KEY in Script Properties')
  return key
}

function clerkFetchAll_(urlBase, apiKey, supportsOrderBy) {
  const all = []
  let offset = 0

  while (true) {
    const order = supportsOrderBy ? '&order_by=created_at' : ''
    const url = `${urlBase}?limit=${CLERK_CFG.PAGE_SIZE}&offset=${offset}${order}`

    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: `Bearer ${apiKey}` },
      muteHttpExceptions: true
    })

    const code = res.getResponseCode()
    if (code >= 300) throw new Error(`Clerk API error ${code}: ${res.getContentText()}`)

    const json = JSON.parse(res.getContentText())
    const page = Array.isArray(json) ? json : (Array.isArray(json.data) ? json.data : [])

    if (!page.length) break

    all.push(...page)
    if (page.length < CLERK_CFG.PAGE_SIZE) break

    offset += CLERK_CFG.PAGE_SIZE
    Utilities.sleep(CLERK_CFG.PAUSE_MS)
  }

  return all
}

function clerkFetchAllUsersForLogin_(headers) {
  const all = []
  let offset = 0

  while (true) {
    const url =
      `${CLERK_CFG.API_BASE}${CLERK_CFG.USERS_ENDPOINT}` +
      `?limit=${CLERK_CFG.PAGE_SIZE}` +
      `&offset=${offset}` +
      `&order_by=created_at`

    const res = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true })
    const code = res.getResponseCode()
    if (code >= 300) throw new Error(`Clerk API error ${code}: ${res.getContentText()}`)

    const json = JSON.parse(res.getContentText())
    const page = Array.isArray(json) ? json : (Array.isArray(json.data) ? json.data : [])
    if (!page.length) break

    all.push(...page)
    if (page.length < CLERK_CFG.PAGE_SIZE) break

    offset += CLERK_CFG.PAGE_SIZE
    Utilities.sleep(CLERK_CFG.PAUSE_MS)
  }

  return all
}

function clerkPrimaryEmail_(u) {
  const arr = u && u.email_addresses ? u.email_addresses : []
  if (!arr || !arr.length) return ''
  const primaryId = u.primary_email_address_id
  const primary = primaryId ? arr.find(e => e && e.id === primaryId) : null
  const picked = primary || arr[0]
  return strOrBlank_(picked && picked.email_address ? picked.email_address : '')
}

function clerkFullName_(u) {
  const first = strOrBlank_(u && u.first_name ? u.first_name : '')
  const last = strOrBlank_(u && u.last_name ? u.last_name : '')
  const full = `${first} ${last}`.trim()
  if (full) return full
  return strOrBlank_(u && u.username ? u.username : '')
}

function clerkToIso_(ts) {
  if (!ts) return ''
  const d = new Date(typeof ts === 'number' ? ts : String(ts))
  if (isNaN(d.getTime())) return ''
  return d.toISOString()
}

function clerkToDate_(ts) {
  if (!ts) return null
  const d = new Date(typeof ts === 'number' ? ts : String(ts))
  return isNaN(d.getTime()) ? null : d
}

function clerkPrivateMeta_(u) {
  const priv =
    (u && u.private_metadata && typeof u.private_metadata === 'object' ? u.private_metadata : null) ||
    (u && u.privateMetaData && typeof u.privateMetaData === 'object' ? u.privateMetaData : null) ||
    {}
  return priv || {}
}

function clerkFetchMembershipsByOrgFallback_(apiKey) {
  const orgs = clerkFetchAll_(`${CLERK_CFG.API_BASE}${CLERK_CFG.ORGS_ENDPOINT}`, apiKey, true)
  const all = []

  orgs.forEach((o, idx) => {
    const orgId = o.id
    if (!orgId) return

    const urlBase = `${CLERK_CFG.API_BASE}${CLERK_CFG.ORGS_ENDPOINT}/${encodeURIComponent(orgId)}/memberships`
    const members = clerkFetchAll_(urlBase, apiKey, false)
    all.push(...members)

    if ((idx + 1) % 5 === 0) Utilities.sleep(CLERK_CFG.PAUSE_MS)
  })

  return all
}

function clerkGetOrgNameMap_() {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(CLERK_CFG.RAW.ORGS)
  if (!sh || sh.getLastRow() < 2) return {}

  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(h => String(h || '').trim().toLowerCase())

  const cId = hdr.indexOf('org_id') + 1
  const cName = hdr.indexOf('org_name') + 1
  if (!cId || !cName) return {}

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
  const out = {}

  data.forEach(r => {
    const id = String(r[cId - 1] || '').trim()
    const name = String(r[cName - 1] || '').trim()
    if (id && !out[id]) out[id] = name
  })

  return out
}

/* =========================
 * Raw sheet writer
 * ========================= */

function clerkOverwriteSheet_(sheet, headers, rows) {
  sheet.clearContents()
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
  sheet.setFrozenRows(1)

  if (rows && rows.length) {
    batchSetValuesSafe_(sheet, 2, 1, rows, CLERK_CFG.WRITE_CHUNK)
  }

  sheet.autoResizeColumns(1, headers.length)
}

/* =========================
 * Login events helpers
 * ========================= */

function clerkBuildLatestLoginByUserId_(eventsSheet) {
  const map = new Map()
  const lastRow = eventsSheet.getLastRow()
  if (lastRow < 2) return map

  const header = eventsSheet.getRange(1, 1, 1, eventsSheet.getLastColumn()).getValues()[0]
    .map(h => String(h || '').trim().toLowerCase())

  const cUserId = header.indexOf('user_id') + 1
  const cLoginTs = header.indexOf('login_timestamp') + 1
  if (!cUserId || !cLoginTs) return map

  const data = eventsSheet.getRange(2, 1, lastRow - 1, eventsSheet.getLastColumn()).getValues()
  data.forEach(r => {
    const userId = String(r[cUserId - 1] || '').trim()
    const tsRaw = r[cLoginTs - 1]
    if (!userId || !tsRaw) return

    const d = new Date(tsRaw)
    if (isNaN(d.getTime())) return

    const ms = d.getTime()
    const prev = map.get(userId) || 0
    if (ms > prev) map.set(userId, ms)
  })

  return map
}

function clerkBuildLoginCountsByUserId_() {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(CLERK_CFG.LOGIN.EVENTS_TAB)
  const out = new Map()
  if (!sh || sh.getLastRow() < 2) return out

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(h => String(h || '').trim().toLowerCase())

  const cUserId = header.indexOf('user_id') + 1
  if (!cUserId) return out

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
  data.forEach(r => {
    const userId = String(r[cUserId - 1] || '').trim()
    if (!userId) return
    out.set(userId, (out.get(userId) || 0) + 1)
  })

  return out
}

/* =========================
 * Safe shared util wrappers
 * ========================= */

function getOrCreateSheetSafe_(ss, name) {
  if (!ss) ss = SpreadsheetApp.getActive()
  const sheetName = String(name || '').trim()
  if (!sheetName) throw new Error('getOrCreateSheetSafe_: name is required')

  // IMPORTANT: do NOT call shared getOrCreateSheet() here.
  // Your shared version appears to sometimes treat the argument as a sheetId (number),
  // which caused: "Sheet <id> not found".
  // This local version is name-only and safe.
  const sh = ss.getSheetByName(sheetName)
  return sh || ss.insertSheet(sheetName)
}

function normalizeEmailSafe_(email) {
  if (typeof normalizeEmail === 'function') return normalizeEmail(email)
  return String(email || '').trim().toLowerCase()
}

function batchSetValuesSafe_(sheet, startRow, startCol, values, chunkSize) {
  if (typeof batchSetValues === 'function') return batchSetValues(sheet, startRow, startCol, values, chunkSize)
  const size = chunkSize || 2000
  for (let i = 0; i < values.length; i += size) {
    const chunk = values.slice(i, i + size)
    sheet.getRange(startRow + i, startCol, chunk.length, chunk[0].length).setValues(chunk)
  }
}

function lockWrapSafe_(lockName, fn) {
  if (typeof fn !== 'function') throw new Error('lockWrap: fn must be a function')

  if (typeof lockWrap === 'function') {
    try {
      return lockWrap(lockName, fn)
    } catch (e) {
      return lockWrap(fn)
    }
  }

  const lock = LockService.getScriptLock()
  if (!lock.tryLock(30000)) throw new Error(`Could not obtain lock: ${lockName}`)
  try {
    return fn()
  } finally {
    lock.releaseLock()
  }
}

function writeSyncLogSafe_(step, status, rowsIn, rowsOut, seconds, error) {
  if (typeof writeSyncLog === 'function') {
    return writeSyncLog(step, status, rowsIn, rowsOut, seconds, error || '')
  }
  Logger.log(`[SYNCLOG missing] ${step} ${status} rows_in=${rowsIn} rows_out=${rowsOut} seconds=${seconds} error=${error || ''}`)
}

/* =========================
 * Tiny helpers
 * ========================= */

function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    sheet.setFrozenRows(1)
    return
  }

  const lastCol = Math.max(sheet.getLastColumn(), headers.length)
  const existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())

  const same =
    existing.length >= headers.length &&
    headers.every((h, i) => String(existing[i] || '').trim() === h)

  if (!same) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    sheet.setFrozenRows(1)
  }
}

function strOrBlank_(v) {
  if (v === null || v === undefined) return ''
  return String(v).trim()
}

function safeJson_(obj) {
  try { return JSON.stringify(obj || {}) } catch (e) { return '{}' }
}