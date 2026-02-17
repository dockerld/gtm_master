/**************************************************************
 * build_canon_users() — self-contained (FIXED last_login + days_since_last_login)
 *
 * Goal:
 *  - days_since_last_login MUST ALWAYS populate when raw_clerk_users has a last login field.
 *
 * Truth sources (priority order):
 *  1) raw_clerk_users row: email_key, email, name, created_at, clerk_user_id, last_login_date-ish
 *     - Supports Date objects like: Wed Nov 19 2025 00:00:00 GMT-0700 ...
 *     - Supports ISO strings like: 2025-12-23T20:13:50.179Z
 *  2) raw_clerk_memberships: org_id + org_role per user (prefer clerk_user_id join)
 *  3) raw_posthog_user_metrics: metrics (prefer email_key, fallback email)
 *  4) login_events: only used for logged_in_days_count + fallback last_login if clerk missing
 *  5) clerk_master fallback
 *
 * Output:
 *  - canon_users (overwrite), preserves manual override columns from prior canon_users
 *
 * Optional shared utils used if present:
 *  - lockWrap(step, fn)
 *  - writeSyncLog(step, status, rows_in, rows_out, seconds, error)
 *  - getOrCreateSheet(ss, name) OR getOrCreateSheet(name)
 *  - readHeaderMap(sheet, headerRow)
 *  - normalizeEmail(email)
 *  - batchSetValues(sheet, startRow, startCol, values, chunkSize)
 **************************************************************/

function build_canon_users() {
  const STEP = 'build_canon_users'

  const lock = (typeof lockWrap === 'function')
    ? lockWrap
    : (_name, fn) => fn()

  return lock(STEP, () => {
    const t0 = new Date()
    try {
      const ss = SpreadsheetApp.getActive()

      // ---------- Config ----------
      const CFG = {
        SHEETS: {
          CLERK_USERS: 'raw_clerk_users',
          CLERK_MEMBERSHIPS: 'raw_clerk_memberships',
          POSTHOG_METRICS: 'raw_posthog_user_metrics',
          LOGIN_EVENTS: 'login_events',
          CLERK_MASTER: 'clerk_master',
          CANON_USERS: 'canon_users'
        },

        CANON_HEADERS: [
          'email_key',
          'email',
          'name',

          'clerk_user_id',
          'created_at',

          'org_id',
          'org_role',

          'logged_in_days_count',
          'last_login_date',
          'days_with_ping',
          'days_since_last_login',

          'meetings_recorded',
          'hours_recorded',
          'ask_meeting',
          'ask_global',
          'client_page_views',
          'clients_count',
          'active_days',
          'action_items_synced',
          'meeting_notes_synced',

          'calendar_connected',
          'first_calendar_connected_date',
          'email_connected',
          'first_email_connected_date',

          'pm_karbon_connected',
          'pm_karbon_first_connected_date',
          'pm_keeper_connected',
          'pm_keeper_first_connected_date',
          'pm_financial_cents_connected',
          'pm_financial_cents_first_connected_date',

          // manual overrides
          'service_override',
          'white_glove_override',
          'in_onboarding_override',
          'tags_override',
          'note_override',

          // derived effective values
          'service_effective',
          'white_glove_effective',
          'in_onboarding_effective',
          'tags_effective',

          'updated_at'
        ],

        MANUAL_FIELDS: new Set([
          'service_override',
          'white_glove_override',
          'in_onboarding_override',
          'tags_override',
          'note_override'
        ])
      }

      // ---------- Sheet helpers ----------
      function getSheet_(name) {
        return ss.getSheetByName(name)
      }

      function getOrCreateSheetSafe_(name) {
        if (typeof getOrCreateSheet === 'function') {
          try { return getOrCreateSheet(ss, name) } catch (e) {}
          try { return getOrCreateSheet(name) } catch (e) {}
        }
        const sh = ss.getSheetByName(name)
        return sh || ss.insertSheet(name)
      }

      const shUsers = getSheet_(CFG.SHEETS.CLERK_USERS)
      const shMems = getSheet_(CFG.SHEETS.CLERK_MEMBERSHIPS)
      const shMetrics = getSheet_(CFG.SHEETS.POSTHOG_METRICS)

      if (!shUsers) throw new Error(`Missing input sheet: ${CFG.SHEETS.CLERK_USERS}`)
      if (!shMems) throw new Error(`Missing input sheet: ${CFG.SHEETS.CLERK_MEMBERSHIPS}`)
      if (!shMetrics) throw new Error(`Missing input sheet: ${CFG.SHEETS.POSTHOG_METRICS}`)

      // ---------- readHeaderMap fallback ----------
      function readHeaderMapSafe_(sheet, headerRow) {
        if (typeof readHeaderMap === 'function') return readHeaderMap(sheet, headerRow)

        const lastCol = sheet.getLastColumn()
        const header = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
        const map = {}
        header.forEach((h, i) => {
          const key = String(h || '').trim().toLowerCase()
          if (!key) return
          map[key] = i + 1 // 1-based
        })
        return { map }
      }

      // ---------- normalizeEmail fallback ----------
      function normalizeEmailSafe_(email) {
        if (typeof normalizeEmail === 'function') return normalizeEmail(email)
        return String(email || '').trim().toLowerCase()
      }

      // ---------- Raw reader ----------
      function readRaw_(sheet, headerRow) {
        const { map } = readHeaderMapSafe_(sheet, headerRow)
        const lastRow = sheet.getLastRow()
        const lastCol = sheet.getLastColumn()
        if (lastRow < headerRow + 1) {
          return { rows: [], has: () => false, col: () => { throw new Error('No rows') } }
        }
        const rows = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()
        return {
          rows,
          map, // keep for fuzzy header access
          has: (h) => map[String(h).toLowerCase()] != null,
          col: (h) => {
            const idx = map[String(h).toLowerCase()]
            if (!idx) throw new Error(`Missing header "${h}" on sheet "${sheet.getName()}"`)
            return idx - 1
          }
        }
      }

      const users = readRaw_(shUsers, 1)
      const mems = readRaw_(shMems, 1)
      const metrics = readRaw_(shMetrics, 1)

      // ---------- Date parsing (robust) ----------
      function parseUnknownDate_(v) {
        if (!v && v !== 0) return null
        if (v instanceof Date) return isNaN(v.getTime()) ? null : v

        const s = String(v || '').trim()
        if (!s) return null

        // YYYY-MM-DD
        if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
          const d = new Date(s + 'T00:00:00Z')
          return isNaN(d.getTime()) ? null : d
        }

        // numeric timestamp
        if (/^\d+$/.test(s)) {
          const n = Number(s)
          if (!isFinite(n)) return null
          const ms = n > 1e12 ? n : n * 1000
          const d = new Date(ms)
          return isNaN(d.getTime()) ? null : d
        }

        // ISO string like 2025-12-23T20:13:50.179Z or Date.toString()
        const d2 = new Date(s)
        return isNaN(d2.getTime()) ? null : d2
      }

      function asYMD_(value) {
        const d = parseUnknownDate_(value)
        if (!d) return ''
        // use UTC to avoid timezone “off by one day” surprises for ISO Z strings
        return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd')
      }

      function daysBetweenYMD_(ymdA, ymdB) {
        const aStr = String(ymdA || '').trim()
        const bStr = String(ymdB || '').trim()
        if (!/^\d{4}-\d{2}-\d{2}$/.test(aStr)) return ''
        if (!/^\d{4}-\d{2}-\d{2}$/.test(bStr)) return ''

        const a = new Date(aStr + 'T00:00:00Z')
        const b = new Date(bStr + 'T00:00:00Z')
        const diff = (b.getTime() - a.getTime()) / (1000 * 60 * 60 * 24)
        if (!isFinite(diff)) return ''
        return Math.floor(diff)
      }

      function computeDaysWithPing_(createdAtRaw, todayYMD) {
        if (!createdAtRaw) return ''
        const createdYMD = asYMD_(createdAtRaw)
        if (!createdYMD) return ''
        return daysBetweenYMD_(createdYMD, todayYMD)
      }

      // ---------- Pick last login from raw_clerk_users row (truth) ----------
      function pickLastLoginFromClerkUsersRow_(usersTbl, row) {
        // Strong preference list
        const candidates = [
          'last_login_date',
          'last_sign_in_at',
          'last_sign_in_date',
          'last_login_at',
          'last_active_at',
          'last_seen_at'
        ]

        for (const h of candidates) {
          if (usersTbl.has(h)) {
            const raw = row[usersTbl.col(h)]
            const ymd = asYMD_(raw)
            if (ymd) return ymd
          }
        }

        // Fuzzy fallback: any header containing last + (login|sign in|active|seen)
        const keys = Object.keys(usersTbl.map || {})
        const fuzzy = keys.find(k => {
          const kk = String(k || '').toLowerCase()
          const hasLast = kk.includes('last')
          const hasSignal =
            kk.includes('login') ||
            (kk.includes('sign') && kk.includes('in')) ||
            kk.includes('active') ||
            kk.includes('seen')
          return hasLast && hasSignal
        })

        if (fuzzy && usersTbl.map[fuzzy]) {
          const raw = row[usersTbl.map[fuzzy] - 1]
          const ymd = asYMD_(raw)
          if (ymd) return ymd
        }

        return ''
      }

      // ---------- Membership indexing ----------
      function buildMembershipIndex_(memsTbl) {
        const byUserIdAll = new Map()
        const byEmailAll = new Map()

        const userIdField =
          memsTbl.has('clerk_user_id') ? 'clerk_user_id' :
          memsTbl.has('user_id') ? 'user_id' :
          ''

        const emailKeyField = memsTbl.has('email_key') ? 'email_key' : ''
        const orgIdField = memsTbl.has('org_id') ? 'org_id' : ''
        const roleField =
          memsTbl.has('role') ? 'role' :
          memsTbl.has('org_role') ? 'org_role' :
          ''

        if (!userIdField) throw new Error('raw_clerk_memberships missing clerk_user_id (or user_id)')
        if (!orgIdField) throw new Error('raw_clerk_memberships missing org_id')

        function add_(mapAll, key, orgId, role) {
          if (!key) return
          if (!mapAll.has(key)) mapAll.set(key, [])
          mapAll.get(key).push({ org_id: orgId, role })
        }

        memsTbl.rows.forEach(r => {
          const userId = String(r[memsTbl.col(userIdField)] || '').trim()
          const orgId = String(r[memsTbl.col(orgIdField)] || '').trim()
          if (!userId || !orgId) return

          const role = roleField ? String(r[memsTbl.col(roleField)] || '').trim() : ''
          add_(byUserIdAll, userId, orgId, role)

          if (emailKeyField) {
            const emailKey = normalizeEmailSafe_(String(r[memsTbl.col(emailKeyField)] || ''))
            add_(byEmailAll, emailKey, orgId, role)
          }
        })

        const priority = (role) => {
          const rr = String(role || '').toLowerCase()
          if (rr.includes('owner')) return 1
          if (rr.includes('admin')) return 2
          if (rr.includes('manager')) return 3
          if (rr.includes('user')) return 4
          return 9
        }

        function chooseBest_(arr) {
          if (!arr || !arr.length) return { org_id: '', role: '' }
          const copy = arr.slice().sort((a, b) => {
            const pa = priority(a.role)
            const pb = priority(b.role)
            if (pa !== pb) return pa - pb
            return String(a.org_id).localeCompare(String(b.org_id))
          })
          return copy[0]
        }

        const byUserId = new Map()
        byUserIdAll.forEach((arr, userId) => byUserId.set(userId, chooseBest_(arr)))

        const byEmail = new Map()
        byEmailAll.forEach((arr, emailKey) => byEmail.set(emailKey, chooseBest_(arr)))

        return { byUserId, byEmail }
      }

      // ---------- Metrics indexing ----------
      function str_(tbl, row, field) {
        if (!tbl.has(field)) return ''
        return String(row[tbl.col(field)] || '').trim()
      }

      function num_(tbl, row, field) {
        if (!tbl.has(field)) return 0
        return Number(row[tbl.col(field)] ?? 0) || 0
      }

      function bool_(tbl, row, field) {
        if (!tbl.has(field)) return false
        const v = row[tbl.col(field)]
        if (v === true) return true
        if (v === false) return false
        const s = String(v || '').toLowerCase().trim()
        return s === 'yes' || s === 'true' || s === '1'
      }

      function emptyMetrics_() {
        return {
          meetings_recorded: 0,
          hours_recorded: 0,
          ask_meeting: 0,
          ask_global: 0,
          client_page_views: 0,
          clients_count: 0,
          active_days: 0,
          action_items_synced: 0,
          meeting_notes_synced: 0,

          calendar_connected: false,
          first_calendar_connected_date: '',

          email_connected: false,
          first_email_connected_date: '',

          pm_karbon_connected: false,
          pm_karbon_first_connected_date: '',
          pm_keeper_connected: false,
          pm_keeper_first_connected_date: '',
          pm_financial_cents_connected: false,
          pm_financial_cents_first_connected_date: ''
        }
      }

      function buildMetricsIndex_(metricsTbl) {
        const out = new Map()
        const emailKeyField = metricsTbl.has('email_key') ? 'email_key' : ''
        const emailField = metricsTbl.has('email') ? 'email' : ''

        metricsTbl.rows.forEach(r => {
          const key =
            emailKeyField ? normalizeEmailSafe_(String(r[metricsTbl.col(emailKeyField)] || '')) :
            emailField ? normalizeEmailSafe_(String(r[metricsTbl.col(emailField)] || '')) :
            ''
          if (!key) return

          out.set(key, {
            meetings_recorded: num_(metricsTbl, r, 'meetings_recorded'),
            hours_recorded: num_(metricsTbl, r, 'hours_recorded'),
            ask_meeting: num_(metricsTbl, r, 'ask_meeting'),
            ask_global: num_(metricsTbl, r, 'ask_global'),
            client_page_views: num_(metricsTbl, r, 'client_page_views'),
            clients_count: num_(metricsTbl, r, 'clients_count'),
            active_days: num_(metricsTbl, r, 'active_days'),
            action_items_synced: num_(metricsTbl, r, 'action_items_synced'),
            meeting_notes_synced: num_(metricsTbl, r, 'meeting_notes_synced'),

            calendar_connected: bool_(metricsTbl, r, 'calendar_connected'),
            first_calendar_connected_date: str_(metricsTbl, r, 'first_calendar_connected_date'),

            email_connected: bool_(metricsTbl, r, 'email_connected'),
            first_email_connected_date: str_(metricsTbl, r, 'first_email_connected_date'),

            pm_karbon_connected: bool_(metricsTbl, r, 'pm_karbon_connected'),
            pm_karbon_first_connected_date: str_(metricsTbl, r, 'pm_karbon_first_connected_date'),

            pm_keeper_connected: bool_(metricsTbl, r, 'pm_keeper_connected'),
            pm_keeper_first_connected_date: str_(metricsTbl, r, 'pm_keeper_first_connected_date'),

            pm_financial_cents_connected: bool_(metricsTbl, r, 'pm_financial_cents_connected'),
            pm_financial_cents_first_connected_date: str_(metricsTbl, r, 'pm_financial_cents_first_connected_date')
          })
        })

        return out
      }

      // ---------- Login rollups (keep for logged_in_days_count + fallback last_login) ----------
      function emptyLogin_() {
        return { logged_in_days_count: '', last_login_date: '' }
      }

      function buildLoginRollups_() {
        // Priority:
        // 1) login_events
        // 2) clerk_master (for rollups)
        // NOTE: last_login_date truth is raw_clerk_users row, but we keep this as fallback.
        // Returns Map(email_key -> { logged_in_days_count, last_login_date })
        // ---- 1) login_events ----
        const shEvents = ss.getSheetByName(CFG.SHEETS.LOGIN_EVENTS)
        if (shEvents && shEvents.getLastRow() >= 2) {
          const { map } = readHeaderMapSafe_(shEvents, 1)
          const cEmail = map['email']
          const cLoginDate = map['login_date']

          if (cEmail && cLoginDate) {
            const lastRow = shEvents.getLastRow()
            const lastCol = shEvents.getLastColumn()
            const data = shEvents.getRange(2, 1, lastRow - 1, lastCol).getValues()

            const datesByEmail = new Map() // email_key -> Set(yyyy-MM-dd)
            data.forEach(r => {
              const emailKey = normalizeEmailSafe_(String(r[cEmail - 1] || ''))
              const loginDate = asYMD_(r[cLoginDate - 1]) // normalize if it’s Date/ISO
              if (!emailKey || !loginDate) return
              if (!datesByEmail.has(emailKey)) datesByEmail.set(emailKey, new Set())
              datesByEmail.get(emailKey).add(loginDate)
            })

            const out = new Map()
            datesByEmail.forEach((set, emailKey) => {
              const dates = Array.from(set).sort()
              const last = dates.length ? dates[dates.length - 1] : ''
              out.set(emailKey, { logged_in_days_count: dates.length, last_login_date: last })
            })
            return out
          }
        }

        // ---- 2) clerk_master ----
        const shMaster = ss.getSheetByName(CFG.SHEETS.CLERK_MASTER)
        if (shMaster && shMaster.getLastRow() >= 2) {
          const { map } = readHeaderMapSafe_(shMaster, 1)
          const cEmail = map['email']
          const cLast = map['last_login_date']
          const cCount = map['logged_in_days_count']

          if (cEmail && (cLast || cCount)) {
            const lastRow = shMaster.getLastRow()
            const lastCol = shMaster.getLastColumn()
            const data = shMaster.getRange(2, 1, lastRow - 1, lastCol).getValues()

            const out = new Map()
            data.forEach(r => {
              const emailKey = normalizeEmailSafe_(String(r[cEmail - 1] || ''))
              if (!emailKey) return
              const lastLogin = cLast ? asYMD_(r[cLast - 1]) : ''
              const count = cCount ? (Number(r[cCount - 1] ?? '') || '') : ''
              out.set(emailKey, { logged_in_days_count: count, last_login_date: lastLogin })
            })
            return out
          }
        }

        return new Map()
      }

      // ---------- Manual override preservation ----------
      function readExistingManualOverrides_(canonSheet) {
        const lastRow = canonSheet.getLastRow()
        const lastCol = canonSheet.getLastColumn()
        if (lastRow < 2 || lastCol < 1) return {}

        const header = canonSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())
        const headerMap = {}
        header.forEach((h, i) => { if (h) headerMap[h.toLowerCase()] = i + 1 })

        if (!headerMap['email_key']) return {}

        const cEmailKey = headerMap['email_key'] - 1
        const cService = headerMap['service_override'] ? headerMap['service_override'] - 1 : null
        const cWG = headerMap['white_glove_override'] ? headerMap['white_glove_override'] - 1 : null
        const cOnb = headerMap['in_onboarding_override'] ? headerMap['in_onboarding_override'] - 1 : null
        const cTags = headerMap['tags_override'] ? headerMap['tags_override'] - 1 : null
        const cNote = headerMap['note_override'] ? headerMap['note_override'] - 1 : null

        const data = canonSheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
        const out = {}

        data.forEach(r => {
          const emailKey = normalizeEmailSafe_(String(r[cEmailKey] || ''))
          if (!emailKey) return
          out[emailKey] = {
            service_override: cService != null ? String(r[cService] || '').trim() : '',
            white_glove_override: cWG != null ? r[cWG] === true : false,
            in_onboarding_override: cOnb != null ? r[cOnb] === true : false,
            tags_override: cTags != null ? String(r[cTags] || '').trim() : '',
            note_override: cNote != null ? String(r[cNote] || '').trim() : ''
          }
        })

        return out
      }

      // ---------- Writer ----------
      function batchSetValuesSafe_(sheet, startRow, startCol, values, chunkSize) {
        if (!values || !values.length) return
        if (typeof batchSetValues === 'function') {
          return batchSetValues(sheet, startRow, startCol, values, chunkSize || 5000)
        }
        const size = chunkSize || 2000
        for (let i = 0; i < values.length; i += size) {
          const chunk = values.slice(i, i + size)
          sheet.getRange(startRow + i, startCol, chunk.length, chunk[0].length).setValues(chunk)
        }
      }

      function writeCanonOverwrite_(sheet, headers, rows) {
        sheet.clearContents()
        sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        sheet.setFrozenRows(1)
        if (rows && rows.length) batchSetValuesSafe_(sheet, 2, 1, rows, 5000)
        sheet.autoResizeColumns(1, headers.length)
      }

      // ---------- Build indices ----------
      const membershipIdx = buildMembershipIndex_(mems)
      const metricsByEmail = buildMetricsIndex_(metrics)
      const loginRollups = buildLoginRollups_()

      // ---------- Manual preservation ----------
      const canonSheet = getOrCreateSheetSafe_(CFG.SHEETS.CANON_USERS)
      const existingManual = readExistingManualOverrides_(canonSheet)

      // ---------- Build output ----------
      const today = new Date()
      const todayStr = Utilities.formatDate(today, 'UTC', 'yyyy-MM-dd') // use UTC to match our asYMD_

      const out = []
      const rowsIn = users.rows.length

      users.rows.forEach(r => {
        const clerkUserId =
          users.has('clerk_user_id') ? String(r[users.col('clerk_user_id')] || '').trim() :
          users.has('user_id') ? String(r[users.col('user_id')] || '').trim() :
          ''

        const email = users.has('email') ? String(r[users.col('email')] || '').trim() : ''
        const emailKey =
          users.has('email_key') ? normalizeEmailSafe_(String(r[users.col('email_key')] || '')) :
          normalizeEmailSafe_(email)

        if (!emailKey) return

        const name =
          users.has('name') ? String(r[users.col('name')] || '').trim() :
          users.has('full_name') ? String(r[users.col('full_name')] || '').trim() :
          (() => {
            const first = users.has('first_name') ? String(r[users.col('first_name')] || '').trim() : ''
            const last = users.has('last_name') ? String(r[users.col('last_name')] || '').trim() : ''
            return `${first} ${last}`.trim()
          })()

        const createdAtRaw = users.has('created_at') ? r[users.col('created_at')] : ''

        const mem =
          (clerkUserId && membershipIdx.byUserId.has(clerkUserId)) ? membershipIdx.byUserId.get(clerkUserId) :
          (membershipIdx.byEmail.has(emailKey) ? membershipIdx.byEmail.get(emailKey) : { org_id: '', role: '' })

        const m = metricsByEmail.get(emailKey) || emptyMetrics_()
        const login = loginRollups.get(emailKey) || emptyLogin_()

        // ✅ Truth: last login comes from raw_clerk_users row (handles Date objects + ISO strings)
        const lastLoginFromClerk = pickLastLoginFromClerkUsersRow_(users, r)
        const lastLoginYMD = lastLoginFromClerk || login.last_login_date || ''

        const daysWithPing = computeDaysWithPing_(createdAtRaw, todayStr)

        // ✅ Should always fill if lastLoginYMD exists
        const daysSinceLastLogin = lastLoginYMD ? daysBetweenYMD_(lastLoginYMD, todayStr) : ''

        const manual = existingManual[emailKey] || {
          service_override: '',
          white_glove_override: false,
          in_onboarding_override: false,
          tags_override: '',
          note_override: ''
        }

        const serviceEff = manual.service_override || ''
        const whiteGloveEff = manual.white_glove_override === true
        const inOnbEff = manual.in_onboarding_override === true
        const tagsEff = manual.tags_override || ''

        out.push([
          emailKey,
          email,
          name,

          clerkUserId,
          createdAtRaw instanceof Date ? createdAtRaw : String(createdAtRaw || '').trim(),

          mem.org_id,
          mem.role,

          login.logged_in_days_count,
          lastLoginYMD,              // ✅ write normalized last_login_date
          daysWithPing,
          daysSinceLastLogin,        // ✅ write computed difference

          m.meetings_recorded,
          m.hours_recorded,
          m.ask_meeting,
          m.ask_global,
          m.client_page_views,
          m.clients_count,
          m.active_days,
          m.action_items_synced,
          m.meeting_notes_synced,

          m.calendar_connected,
          m.first_calendar_connected_date,
          m.email_connected,
          m.first_email_connected_date,

          m.pm_karbon_connected,
          m.pm_karbon_first_connected_date,
          m.pm_keeper_connected,
          m.pm_keeper_first_connected_date,
          m.pm_financial_cents_connected,
          m.pm_financial_cents_first_connected_date,

          manual.service_override,
          manual.white_glove_override === true,
          manual.in_onboarding_override === true,
          manual.tags_override,
          manual.note_override,

          serviceEff,
          whiteGloveEff,
          inOnbEff,
          tagsEff,

          today
        ])
      })

      writeCanonOverwrite_(canonSheet, CFG.CANON_HEADERS, out)

      if (typeof writeSyncLog === 'function') {
        writeSyncLog(STEP, 'ok', rowsIn, out.length, (new Date() - t0) / 1000, '')
      }

      return { rows_in: rowsIn, rows_out: out.length }
    } catch (err) {
      const msg = String(err && err.message ? err.message : err)
      if (typeof writeSyncLog === 'function') {
        writeSyncLog(STEP, 'error', '', '', (new Date() - t0) / 1000, msg)
      }
      throw err
    }
  })
}
