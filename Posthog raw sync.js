/**************************************************************
 * PostHog Raw Sync (overwrite-only) — HARDENED AGAINST 504s
 *
 * Creates/overwrites:
 *  - raw_posthog_user_metrics
 *
 * Pulls (per email_key):
 *  - meetings_recorded
 *  - hours_recorded
 *  - ask_meeting
 *  - ask_global
 *  - client_page_views
 *  - active_days (distinct days with ANY events)
 *  - clients_count (ORG-level; fanned back to user)
 *  - calendar_connected + first_calendar_connected_date
 *  - email_connected + first_email_connected_date
 *  - PM providers + first connected dates (one column per provider)
 *
 * Key improvements:
 * 1) Retries + exponential backoff for PostHog 429/502/503/504 (common transient failures)
 * 2) Smaller batches + longer pauses (reduces load)
 * 3) Event queries bounded by a lookback window (reduces scan size → fewer timeouts)
 * 4) Writes partial progress only at end (same behavior), but logs where it failed
 *
 * Uses Script Properties:
 *  - POSTHOG_API_KEY
 *  - POSTHOG_PROJECT_ID (optional; falls back to config)
 *
 * Email source:
 *  - Reads emails from raw_clerk_users by default (recommended)
 *
 * Notes:
 * - No semicolons in HogQL
 * - Overwrite-only raw table. Canon tables handle editability rules.
 **************************************************************/

const POSTHOG_RAW_CFG = {
  PROJECT_ID_FALLBACK: '179975',
  API_BASE: 'https://app.posthog.com/api',

  // ↓ Lower batch size helps avoid large query payloads/timeouts
  BATCH_SIZE: 100,

  // ↓ Slightly more breathing room between batches
  PAUSE_MS: 300,

  // ↓ Write chunking to Sheets
  WRITE_CHUNK: 5000,

  // ↓ Event query bounds (reduce workload). Adjust if you truly need “all time”.
  EVENT_LOOKBACK_DAYS: 365,

  // ↓ Retry policy for PostHog API
  RETRY: {
    MAX_ATTEMPTS: 6,          // total attempts per query
    BASE_SLEEP_MS: 750,       // backoff base
    MAX_SLEEP_MS: 15000,      // cap
    JITTER_MS: 250            // jitter to avoid thundering herd
  },

  SHEETS: {
    SOURCE_USERS: 'raw_clerk_users',
    DEST: 'raw_posthog_user_metrics'
  },

  SOURCE_HEADERS: {
    EMAIL: 'email'
  }
}

function posthog_pull_user_metrics_to_raw() {
  const t0 = new Date()
  const props = PropertiesService.getScriptProperties()

  const apiKey = props.getProperty('POSTHOG_API_KEY')
  if (!apiKey) throw new Error('Missing POSTHOG_API_KEY in Script Properties')

  const projectId = props.getProperty('POSTHOG_PROJECT_ID') || POSTHOG_RAW_CFG.PROJECT_ID_FALLBACK
  if (!projectId) throw new Error('Missing POSTHOG_PROJECT_ID (or set PROJECT_ID_FALLBACK)')

  const ss = SpreadsheetApp.getActive()

  // 1) Get email list (prefer raw_clerk_users)
  const source = ss.getSheetByName(POSTHOG_RAW_CFG.SHEETS.SOURCE_USERS)
  if (!source) throw new Error(`Source sheet not found: ${POSTHOG_RAW_CFG.SHEETS.SOURCE_USERS}`)

  const emails = posthogReadEmails_(source, 1, POSTHOG_RAW_CFG.SOURCE_HEADERS.EMAIL)
  const uniqueEmailKeys = Array.from(new Set(emails.filter(Boolean).map(e => normalizeEmail(e))))

  Logger.log(`PostHog: unique emails to query = ${uniqueEmailKeys.length}`)
  if (uniqueEmailKeys.length === 0) {
    posthogWriteSyncLogSafe_(
      'posthog_pull_user_metrics_to_raw',
      'ok',
      0,
      0,
      (new Date() - t0) / 1000,
      'no emails to query'
    )
    return { rows_in: 0, rows_out: 0 }
  }

  // 2) Query PostHog in batches
  const metricsMap = new Map()     // email_key -> record (db metrics)
  const pageViewsMap = new Map()   // email_key -> client_page_views
  const activeDaysMap = new Map()  // email_key -> active_days (ANY events)

  for (let i = 0; i < uniqueEmailKeys.length; i += POSTHOG_RAW_CFG.BATCH_SIZE) {
    const batch = uniqueEmailKeys.slice(i, i + POSTHOG_RAW_CFG.BATCH_SIZE)
    const batchNum = Math.floor(i / POSTHOG_RAW_CFG.BATCH_SIZE) + 1
    Logger.log(`PostHog batch ${batchNum}: ${batch.length} emails`)

    // 2a) DB-backed metrics (postgres.* tables via HogQL)
    {
      const sql = posthogBuildHogQL_dbMetrics_(batch)
      const rows = posthogRunQuery_(apiKey, projectId, sql, `dbMetrics batch ${batchNum}`)

      rows.forEach(r => {
        const emailKey = normalizeEmail(r?.[0] || '')
        if (!emailKey) return

        metricsMap.set(emailKey, {
          email_key: emailKey,
          email: String(r?.[1] || ''),

          meetings_recorded: Number(r?.[2] ?? 0),
          hours_recorded: Number(r?.[3] ?? 0),
          ask_meeting: Number(r?.[4] ?? 0),
          ask_global: Number(r?.[5] ?? 0),
          clients_count: Number(r?.[6] ?? 0),

          calendar_connected: String(r?.[7] || '').toLowerCase() === 'yes',
          first_calendar_connected_date: String(r?.[8] || ''),

          email_connected: String(r?.[9] || '').toLowerCase() === 'yes',
          first_email_connected_date: String(r?.[10] || ''),

          pm_karbon_connected: String(r?.[11] || '').toLowerCase() === 'yes',
          pm_karbon_first_connected_date: String(r?.[12] || ''),

          pm_keeper_connected: String(r?.[13] || '').toLowerCase() === 'yes',
          pm_keeper_first_connected_date: String(r?.[14] || ''),

          pm_financial_cents_connected: String(r?.[15] || '').toLowerCase() === 'yes',
          pm_financial_cents_first_connected_date: String(r?.[16] || '')
        })
      })
    }

    // 2b) Event metrics (client page views from events/persons) — bounded by lookback
    {
      const sql = posthogBuildHogQL_clientPageViews_(batch, POSTHOG_RAW_CFG.EVENT_LOOKBACK_DAYS)
      const rows = posthogRunQuery_(apiKey, projectId, sql, `clientPageViews batch ${batchNum}`)

      rows.forEach(r => {
        const emailKey = normalizeEmail(r?.[0] || '')
        if (!emailKey) return
        pageViewsMap.set(emailKey, Number(r?.[1] ?? 0))
      })
    }

    // 2c) Active days (ANY events) — bounded by lookback
    {
      const sql = posthogBuildHogQL_activeDays_(batch, POSTHOG_RAW_CFG.EVENT_LOOKBACK_DAYS)
      const rows = posthogRunQuery_(apiKey, projectId, sql, `activeDays batch ${batchNum}`)

      rows.forEach(r => {
        const emailKey = normalizeEmail(r?.[0] || '')
        if (!emailKey) return
        activeDaysMap.set(emailKey, Number(r?.[1] ?? 0))
      })
    }

    Utilities.sleep(POSTHOG_RAW_CFG.PAUSE_MS)
  }

  // 3) Build output rows (one row per email_key we queried)
  const headers = [
    'email_key',
    'email',

    'meetings_recorded',
    'hours_recorded',
    'ask_meeting',
    'ask_global',
    'client_page_views',
    'active_days',
    'clients_count',

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

    'pulled_at'
  ]

  const pulledAt = new Date()
  const rowsOut = uniqueEmailKeys.map(emailKey => {
    const base = metricsMap.get(emailKey) || {
      email_key: emailKey,
      email: '',

      meetings_recorded: 0,
      hours_recorded: 0,
      ask_meeting: 0,
      ask_global: 0,
      clients_count: 0,

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

    const clientViews = pageViewsMap.has(emailKey) ? pageViewsMap.get(emailKey) : 0
    const activeDays  = activeDaysMap.has(emailKey) ? activeDaysMap.get(emailKey) : 0

    return [
      base.email_key,
      base.email,

      base.meetings_recorded,
      base.hours_recorded,
      base.ask_meeting,
      base.ask_global,
      clientViews,
      activeDays,
      base.clients_count,

      base.calendar_connected,
      base.first_calendar_connected_date || '',
      base.email_connected,
      base.first_email_connected_date || '',

      base.pm_karbon_connected,
      base.pm_karbon_first_connected_date || '',
      base.pm_keeper_connected,
      base.pm_keeper_first_connected_date || '',
      base.pm_financial_cents_connected,
      base.pm_financial_cents_first_connected_date || '',

      pulledAt
    ]
  })

  // 4) Overwrite destination
  const dest = getOrCreateSheetSafe_(ss, POSTHOG_RAW_CFG.SHEETS.DEST)
  posthogOverwriteSheet_(dest, headers, rowsOut)

  posthogWriteSyncLogSafe_(
    'posthog_pull_user_metrics_to_raw',
    'ok',
    uniqueEmailKeys.length,
    rowsOut.length,
    (new Date() - t0) / 1000,
    `lookback_days=${POSTHOG_RAW_CFG.EVENT_LOOKBACK_DAYS} batch_size=${POSTHOG_RAW_CFG.BATCH_SIZE}`
  )

  return { rows_in: uniqueEmailKeys.length, rows_out: rowsOut.length }
}

/* =========================
 * HogQL builders
 * ========================= */

function posthogBuildHogQL_dbMetrics_(emailKeys) {
  const quoted = emailKeys.map(e => `'${String(e).replace(/'/g, "''")}'`).join(', ')

  return `
WITH [${quoted}] AS input_emails

, base_users AS (
  SELECT
    u.id AS user_id,
    lower(u.email) AS email_key,
    u.email AS email
  FROM postgres.users AS u
  WHERE lower(u.email) IN (SELECT arrayJoin(input_emails))
)

 , meeting_bot_stats AS (
  SELECT
    mb.user_id,
    countIf(
      mb.recording_started_at IS NOT NULL
      AND mb.recording_ended_at IS NOT NULL
    ) AS meetings_recorded,
    round(
      sumIf(
        dateDiff('second', mb.recording_started_at, mb.recording_ended_at),
        mb.recording_started_at IS NOT NULL
        AND mb.recording_ended_at IS NOT NULL
      ) / 3600,
      2
    ) AS hours_recorded
  FROM postgres.meeting_bots AS mb
  GROUP BY mb.user_id
)

, ask_counts AS (
  SELECT
    t.resource_id AS user_id,
    countIf(t.id LIKE 'meeting%') AS ask_meeting,
    countIf(t.id LIKE 'global%')  AS ask_global
  FROM postgres.mastra.mastra_threads AS t
  GROUP BY t.resource_id
)

-- ORG-LEVEL client counts (no fan-out)
, user_orgs AS (
  SELECT DISTINCT
    bu.user_id,
    uo.org_id
  FROM base_users AS bu
  JOIN postgres.users_orgs AS uo
    ON uo.user_id = bu.user_id
)

, org_scope AS (
  SELECT DISTINCT org_id
  FROM user_orgs
)

, org_client_counts AS (
  SELECT
    os.org_id,
    countIf(
      c.id IS NOT NULL
      AND lower(trim(coalesce(c.name, ''))) != 'camden bean'
    ) AS clients_count
  FROM org_scope AS os
  LEFT JOIN postgres.clients AS c
    ON c.org_id = os.org_id
  GROUP BY os.org_id
)

, clients_count_per_user AS (
  SELECT
    uo.user_id,
    max(coalesce(occ.clients_count, 0)) AS clients_count
  FROM user_orgs AS uo
  LEFT JOIN org_client_counts AS occ
    ON occ.org_id = uo.org_id
  GROUP BY uo.user_id
)

, oauth_rollup AS (
  SELECT
    oc.user_id,

    if(countIf(oc.scope_type = 'CALENDAR') > 0, 'yes', 'no') AS calendar_connected,
    formatDateTime(minIf(oc.created_at, oc.scope_type = 'CALENDAR'), '%Y-%m-%d') AS first_calendar_connected_date,

    if(countIf(oc.scope_type = 'EMAIL') > 0, 'yes', 'no') AS email_connected,
    formatDateTime(minIf(oc.created_at, oc.scope_type = 'EMAIL'), '%Y-%m-%d') AS first_email_connected_date,

    if(countIf(oc.provider = 'KARBON') > 0, 'yes', 'no') AS pm_karbon_connected,
    formatDateTime(minIf(oc.created_at, oc.provider = 'KARBON'), '%Y-%m-%d') AS pm_karbon_first_connected_date,

    if(countIf(oc.provider = 'KEEPER') > 0, 'yes', 'no') AS pm_keeper_connected,
    formatDateTime(minIf(oc.created_at, oc.provider = 'KEEPER'), '%Y-%m-%d') AS pm_keeper_first_connected_date,

    if(countIf(oc.provider = 'FINANCIAL_CENTS') > 0, 'yes', 'no') AS pm_financial_cents_connected,
    formatDateTime(minIf(oc.created_at, oc.provider = 'FINANCIAL_CENTS'), '%Y-%m-%d') AS pm_financial_cents_first_connected_date

  FROM postgres.oauth_credentials AS oc
  GROUP BY oc.user_id
)

SELECT
  u.email_key,
  u.email,

  coalesce(mbs.meetings_recorded, 0) AS meetings_recorded,
  coalesce(mbs.hours_recorded, 0) AS hours_recorded,
  coalesce(a.ask_meeting, 0) AS ask_meeting,
  coalesce(a.ask_global, 0)  AS ask_global,

  coalesce(ccpu.clients_count, 0) AS clients_count,

  coalesce(o.calendar_connected, 'no') AS calendar_connected,
  coalesce(o.first_calendar_connected_date, '') AS first_calendar_connected_date,

  coalesce(o.email_connected, 'no') AS email_connected,
  coalesce(o.first_email_connected_date, '') AS first_email_connected_date,

  coalesce(o.pm_karbon_connected, 'no') AS pm_karbon_connected,
  coalesce(o.pm_karbon_first_connected_date, '') AS pm_karbon_first_connected_date,

  coalesce(o.pm_keeper_connected, 'no') AS pm_keeper_connected,
  coalesce(o.pm_keeper_first_connected_date, '') AS pm_keeper_first_connected_date,

  coalesce(o.pm_financial_cents_connected, 'no') AS pm_financial_cents_connected,
  coalesce(o.pm_financial_cents_first_connected_date, '') AS pm_financial_cents_first_connected_date

FROM base_users AS u
LEFT JOIN meeting_bot_stats       AS mbs  ON mbs.user_id = u.user_id
LEFT JOIN ask_counts              AS a    ON a.user_id = u.user_id
LEFT JOIN clients_count_per_user  AS ccpu ON ccpu.user_id = u.user_id
LEFT JOIN oauth_rollup            AS o    ON o.user_id = u.user_id

ORDER BY u.email_key
LIMIT 50000
  `.trim()
}

function posthogBuildHogQL_clientPageViews_(emailKeys, lookbackDays) {
  const quoted = emailKeys.map(e => `'${String(e).replace(/'/g, "''")}'`).join(', ')
  const days = Math.max(1, Number(lookbackDays || 365) || 365)

  return `
WITH [${quoted}] AS input_emails
SELECT
  lower(p.properties.email) AS email_key,
  count() AS client_page_views
FROM events AS e
JOIN persons AS p
  ON e.person_id = p.id
WHERE e.event = '$pageview'
  AND e.timestamp >= now() - INTERVAL ${days} DAY
  AND lower(p.properties.email) IN input_emails
  AND (
    lower(e.properties.$current_url) LIKE '%/clients%'
    OR lower(e.properties.$pathname) LIKE '%/clients%'
  )
GROUP BY email_key
ORDER BY client_page_views DESC
LIMIT 50000
  `.trim()
}

function posthogBuildHogQL_activeDays_(emailKeys, lookbackDays) {
  const quoted = emailKeys.map(e => `'${String(e).replace(/'/g, "''")}'`).join(', ')
  const days = Math.max(1, Number(lookbackDays || 365) || 365)

  return `
WITH [${quoted}] AS input_emails
SELECT
  lower(p.properties.email) AS email_key,
  count(DISTINCT toDate(e.timestamp)) AS active_days
FROM events AS e
JOIN persons AS p
  ON e.person_id = p.id
WHERE e.timestamp >= now() - INTERVAL ${days} DAY
  AND lower(p.properties.email) IN (SELECT arrayJoin(input_emails))
GROUP BY email_key
ORDER BY active_days DESC
LIMIT 50000
  `.trim()
}

/* =========================
 * PostHog API + helpers
 * ========================= */

function posthogRunQuery_(apiKey, projectId, hogql, label) {
  const payload = { query: { kind: 'HogQLQuery', query: hogql } }
  const url = `${POSTHOG_RAW_CFG.API_BASE}/projects/${projectId}/query`

  const maxAttempts = POSTHOG_RAW_CFG.RETRY.MAX_ATTEMPTS
  const baseSleep = POSTHOG_RAW_CFG.RETRY.BASE_SLEEP_MS
  const maxSleep = POSTHOG_RAW_CFG.RETRY.MAX_SLEEP_MS
  const jitter = POSTHOG_RAW_CFG.RETRY.JITTER_MS

  let lastErr = null

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: `Bearer ${apiKey}` },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    })

    const code = res.getResponseCode()
    const text = res.getContentText() || ''

    if (code >= 200 && code < 300) {
      const json = JSON.parse(text)
      return json.results || []
    }

    // Retry on transient statuses
    const shouldRetry = (code === 429 || code === 502 || code === 503 || code === 504)
    lastErr = new Error(`PostHog API error ${code}: ${text}`)

    if (!shouldRetry || attempt === maxAttempts) break

    const sleepMs = Math.min(
      maxSleep,
      Math.floor(baseSleep * Math.pow(2, attempt - 1) + Math.random() * jitter)
    )

    Logger.log(
      `[PostHog retry] ${label || 'query'} attempt ${attempt}/${maxAttempts} got ${code}. Sleeping ${sleepMs}ms`
    )
    Utilities.sleep(sleepMs)
  }

  throw lastErr
}

function posthogReadEmails_(sheet, headerRow, emailHeaderName) {
  const { map } = readHeaderMap(sheet, headerRow)
  const cEmail = map[String(emailHeaderName).toLowerCase()]
  if (!cEmail) throw new Error(`posthogReadEmails_: header not found: ${emailHeaderName}`)

  const lastRow = sheet.getLastRow()
  if (lastRow < headerRow + 1) return []

  const n = lastRow - headerRow
  const vals = sheet.getRange(headerRow + 1, cEmail, n, 1).getValues()
  return vals.map(r => String(r[0] || '').trim()).filter(Boolean)
}

function posthogOverwriteSheet_(sheet, headers, rows) {
  sheet.clearContents()
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
  sheet.setFrozenRows(1)

  if (rows && rows.length) {
    if (typeof batchSetValues === 'function') batchSetValues(sheet, 2, 1, rows, POSTHOG_RAW_CFG.WRITE_CHUNK)
    else sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows)
  }
  sheet.autoResizeColumns(1, headers.length)
}

/* =========================
 * Minimal shared utilities (fallbacks)
 * ========================= */

function posthogWriteSyncLogSafe_(step, status, rowsIn, rowsOut, seconds, error) {
  if (typeof writeSyncLog === 'function') return writeSyncLog(step, status, rowsIn, rowsOut, seconds, error || '')
  Logger.log(`[SYNCLOG missing] ${step} ${status} rows_in=${rowsIn} rows_out=${rowsOut} seconds=${seconds} error=${error || ''}`)
}
