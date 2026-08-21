/**************************************************************
 * Sauron (PRODUCTION)
 *
 * Builds the live "Sauron" sheet from a single PostHog HogQL query
 * (users + orgs + subscriptions + stripe + events). Raw dump — manual
 * columns are no longer preserved (that data lives in the CRM now).
 *
 * Also exports sauronQueryRun_, the shared columns-aware HogQL runner
 * used by the other query builders.
 *
 * render_sauron_view() runs in the daily pipeline.
 *
 * Script Properties used:
 *  - POSTHOG_API_KEY
 *  - POSTHOG_PROJECT_ID (optional; falls back to POSTHOG_RAW_CFG)
 **************************************************************/

const SAURON_QUERY_CFG = {
  OUT_SHEET: 'Sauron'
}

// The query is kept verbatim so it matches what was validated in PostHog.
const SAURON_QUERY_HOGQL = `
WITH
base AS (  -- one row per real user
  SELECT u.id AS user_id, u.email AS email, lower(u.email) AS email_l, u.name AS name, u.created_at AS user_created
  FROM postgres.users u WHERE u.email IS NOT NULL AND u.email != ''),
prim_org AS (  -- a user's primary org = earliest membership
  SELECT user_id, argMin(org_id, created_at) AS org_id FROM postgres.users_orgs GROUP BY user_id),
org_members AS (SELECT org_id, count(DISTINCT user_id) AS org_members FROM postgres.users_orgs GROUP BY org_id),
orgs AS (SELECT id AS org_id, name AS org_name, created_at AS org_created FROM postgres.orgs LIMIT 1 BY id),
subs AS (SELECT org_id, argMax(status, created_at) AS sub_status,
         argMax(coalesce(full_seat_count,0)+coalesce(lite_seat_count,0), created_at) AS seats
         FROM postgres.org_subscriptions GROUP BY org_id),
-- Paying = The Ring's exact gate (real paid invoice, live, not sheet-excluded)
paid_subs AS (SELECT subscription_id FROM stripe.invoice WHERE status='paid' AND toFloat(total)>0 AND subscription_id IS NOT NULL GROUP BY subscription_id),
-- Map an override-sheet org_id to the internal org id (internal DB id OR app/WorkOS id).
org_keys AS (SELECT id AS org_id, toString(id) AS k FROM postgres.orgs LIMIT 1 BY id UNION ALL SELECT id AS org_id, coalesce(nullIf(workos_id,''), external_id) AS k FROM postgres.orgs LIMIT 1 BY id),
-- MANAGED orgs: keyed on sheet org_id. Marks them paying; seats = full_seats+lite_seats.
managed AS (SELECT ok.org_id AS org_id, max(toFloat64OrNull(replaceRegexpAll(coalesce(toString(ovr.amount),''),'[^0-9.]',''))) AS amt, max(toIntOrZero(replaceRegexpAll(coalesce(toString(ovr.full_seats),''),'[^0-9]',''))) AS full_seats, max(toIntOrZero(replaceRegexpAll(coalesce(toString(ovr.lite_seats),''),'[^0-9]',''))) AS lite_seats FROM override_googlesheets_manual_stripe_changes ovr JOIN org_keys ok ON ok.k = toString(ovr.org_id) WHERE lower(ovr.exclude_reason)='managed' AND coalesce(toString(ovr.org_id),'') != '' GROUP BY ok.org_id),
-- Non-managed exclusions (not counted as paying), keyed on the sheet's customer_email.
sheet_excl AS (SELECT DISTINCT os.org_id FROM postgres.org_subscriptions os JOIN (SELECT id, lower(email) AS email FROM stripe.customer LIMIT 1 BY id) c ON c.id = os.stripe_customer_id WHERE c.email IN (SELECT lower(customer_email) FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) != 'managed' AND coalesce(customer_email,'') != '')),
ss AS (SELECT id, canceled_at, status FROM stripe.subscription LIMIT 1 BY id),
osub AS (SELECT org_id, stripe_subscription_id FROM postgres.org_subscriptions),
paying_orgs AS (SELECT DISTINCT osub.org_id AS org_id FROM osub JOIN ss ON ss.id=osub.stripe_subscription_id
  WHERE osub.stripe_subscription_id IN (SELECT subscription_id FROM paid_subs)
    AND osub.org_id NOT IN (SELECT org_id FROM sheet_excl)
    AND NOT (ss.status='canceled' OR (ss.canceled_at IS NOT NULL AND ss.canceled_at < now()))),
promo AS (SELECT pr.org_id AS org_id, argMax(pc.code, pr.redeemed_at) AS promo_code
  FROM postgres.promo_redemptions pr JOIN postgres.promo_codes pc ON pc.id=pr.promo_code_id GROUP BY pr.org_id),
clients AS (SELECT org_id, countIf(id IS NOT NULL AND lower(trim(coalesce(name,'')))!='camden bean') AS clients_count FROM postgres.clients GROUP BY org_id),
mb AS (SELECT user_id, countIf(recording_started_at IS NOT NULL AND recording_ended_at IS NOT NULL) AS meetings_recorded,
       round(sumIf(dateDiff('second', recording_started_at, recording_ended_at), recording_started_at IS NOT NULL AND recording_ended_at IS NOT NULL)/3600,2) AS hours_recorded
       FROM postgres.meeting_bots GROUP BY user_id),
ask AS (SELECT resource_id AS user_id, countIf(id LIKE 'meeting%') AS ask_meeting, countIf(id LIKE 'global%') AS ask_global FROM postgres.mastra.mastra_threads GROUP BY resource_id),
ai AS (SELECT user_id, countIf(synced_to_practice_management = true) AS action_items_synced FROM postgres.action_items GROUP BY user_id),
notes AS (SELECT user_id, count() AS meeting_notes_synced FROM postgres.meetings WHERE sync_status IS NOT NULL GROUP BY user_id),
oauth AS (SELECT user_id,
  if(countIf(scope_type='CALENDAR')>0,'yes','no') AS cal_connected,
  formatDateTime(minIf(created_at, scope_type='CALENDAR'),'%Y-%m-%d') AS cal_connected_date,
  if(countIf(scope_type='EMAIL')>0,'yes','no') AS email_connected,
  formatDateTime(minIf(created_at, scope_type='EMAIL'),'%Y-%m-%d') AS email_connected_date,
  arrayStringConcat(arraySort(groupUniqArrayIf(provider, provider IN ('KARBON','KEEPER','FINANCIAL_CENTS'))), ', ') AS pm,
  formatDateTime(minIf(created_at, provider IN ('KARBON','KEEPER','FINANCIAL_CENTS')),'%Y-%m-%d') AS pm_connected_date
  FROM postgres.oauth_credentials GROUP BY user_id),
ev AS (  -- event-based, keyed by email, 365-day lookback
  SELECT lower(p.properties.email) AS email_l,
    count(DISTINCT toDate(e.timestamp)) AS active_days,
    countIf((lower(e.properties['$current_url']) LIKE '%/clients%' OR lower(e.properties['$pathname']) LIKE '%/clients%') AND e.event='$pageview') AS client_page_views,
    countIf(e.event='auth_login_succeeded') AS logins,
    maxIf(e.timestamp, e.event='auth_login_succeeded') AS last_login
  FROM events e JOIN persons p ON p.id=e.person_id
  WHERE lower(p.properties.email) IN (SELECT email_l FROM base) AND e.timestamp > now() - INTERVAL 365 DAY
  GROUP BY email_l)
SELECT
  b.email AS email, b.name AS name, o.org_name AS org_name,
  dateDiff('day', b.user_created, now()) AS days_with_ping,
  if(ev.last_login IS NULL, NULL, dateDiff('day', ev.last_login, now())) AS days_since_last_login,
  coalesce(mb.meetings_recorded,0) AS meetings_recorded,
  coalesce(mb.hours_recorded,0) AS hours_recorded,
  coalesce(notes.meeting_notes_synced,0) AS meeting_notes_synced,
  coalesce(ai.action_items_synced,0) AS action_items_synced,
  coalesce(ev.logins,0) AS logged_in_count,
  coalesce(ev.active_days,0) AS active_days,
  if(mg.org_id IS NOT NULL AND (coalesce(mg.full_seats,0)+coalesce(mg.lite_seats,0))>0, coalesce(mg.full_seats,0)+coalesce(mg.lite_seats,0), coalesce(s.seats,0)) AS seats,
  coalesce(cl.clients_count,0) AS clients,
  if(po.org_id IS NOT NULL OR mg.org_id IS NOT NULL,'yes','no') AS paying,
  formatDateTime(b.user_created,'%Y-%m-%d') AS sign_up_date,
  coalesce(ask.ask_meeting,0) AS ask_meeting,
  coalesce(ask.ask_global,0) AS ask_global,
  coalesce(ev.client_page_views,0) AS client_page_views,
  coalesce(oauth.pm,'') AS pm,
  coalesce(oauth.pm_connected_date,'') AS pm_connected_date,
  coalesce(oauth.cal_connected,'no') AS cal_connected,
  coalesce(oauth.cal_connected_date,'') AS cal_connected_date,
  coalesce(oauth.email_connected,'no') AS email_connected,
  coalesce(oauth.email_connected_date,'') AS email_connected_date,
  coalesce(pr.promo_code,'') AS promo_code,
  coalesce(s.sub_status,'') AS status,
  formatDateTime(o.org_created,'%Y-%m-%d') AS org_sign_up_date,
  coalesce(om.org_members,0) AS org_members
FROM base b
LEFT JOIN prim_org pog ON pog.user_id=b.user_id
LEFT JOIN orgs o ON o.org_id=pog.org_id
LEFT JOIN org_members om ON om.org_id=pog.org_id
LEFT JOIN subs s ON s.org_id=pog.org_id
LEFT JOIN paying_orgs po ON po.org_id=pog.org_id
LEFT JOIN managed mg ON mg.org_id=pog.org_id
LEFT JOIN promo pr ON pr.org_id=pog.org_id
LEFT JOIN clients cl ON cl.org_id=pog.org_id
LEFT JOIN mb ON mb.user_id=b.user_id
LEFT JOIN ask ON ask.user_id=b.user_id
LEFT JOIN ai ON ai.user_id=b.user_id
LEFT JOIN notes ON notes.user_id=b.user_id
LEFT JOIN oauth ON oauth.user_id=b.user_id
LEFT JOIN ev ON ev.email_l=b.email_l
ORDER BY org_name, email
LIMIT 5000
`

/**
 * Build the LIVE "Sauron" sheet from the PostHog query (raw dump — manual
 * columns are no longer preserved; that lives in the CRM now).
 * Returns { rows_in, rows_out } for pipeline-style logging.
 */
function render_sauron_view() {
  const t0 = new Date()
  const ss = SpreadsheetApp.getActive()

  const props = PropertiesService.getScriptProperties()
  const apiKey = props.getProperty('POSTHOG_API_KEY')
  if (!apiKey) throw new Error('Missing POSTHOG_API_KEY in Script Properties')
  const projectId = props.getProperty('POSTHOG_PROJECT_ID') || POSTHOG_RAW_CFG.PROJECT_ID_FALLBACK

  const { columns, results } = sauronQueryRun_(apiKey, projectId, SAURON_QUERY_HOGQL, 'render_sauron_view')

  const headers = (columns && columns.length)
    ? columns
    : ['(no columns returned)']

  // Normalize each row to the header width; leave values as-is (numbers/strings/null).
  const rows = (results || []).map(r => {
    const row = new Array(headers.length)
    for (let i = 0; i < headers.length; i++) row[i] = (r && r[i] != null) ? r[i] : ''
    return row
  })

  const sh = ss.getSheetByName(SAURON_QUERY_CFG.OUT_SHEET) || ss.insertSheet(SAURON_QUERY_CFG.OUT_SHEET)
  sh.clearContents()
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#F3F4F6')
  sh.setFrozenRows(1)
  if (rows.length) {
    const chunk = 5000
    for (let i = 0; i < rows.length; i += chunk) {
      const part = rows.slice(i, i + chunk)
      sh.getRange(2 + i, 1, part.length, headers.length).setValues(part)
    }
  }
  try { sh.autoResizeColumns(1, headers.length) } catch (e) {}

  if (typeof writeSyncLog === 'function') {
    writeSyncLog('render_sauron_view', 'ok', rows.length, rows.length, (new Date() - t0) / 1000, '')
  }
  Logger.log(`render_sauron_view: ${rows.length} rows in ${((new Date() - t0) / 1000).toFixed(1)}s`)
  return { rows_in: rows.length, rows_out: rows.length }
}

/**
 * HogQL runner that returns BOTH columns and results (with retry/backoff).
 * Mirrors posthogRunQuery_ but keeps the column names so the test sheet
 * header always matches the query.
 */
function sauronQueryRun_(apiKey, projectId, hogql, label) {
  const payload = { query: { kind: 'HogQLQuery', query: hogql } }
  const url = `${POSTHOG_RAW_CFG.API_BASE}/projects/${projectId}/query`

  const R = POSTHOG_RAW_CFG.RETRY
  let lastErr = null

  for (let attempt = 1; attempt <= R.MAX_ATTEMPTS; attempt++) {
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
      return { columns: json.columns || [], results: json.results || [] }
    }

    const shouldRetry = (code === 429 || code === 502 || code === 503 || code === 504)
    lastErr = new Error(`PostHog API error ${code}: ${text}`)
    if (!shouldRetry || attempt === R.MAX_ATTEMPTS) break

    const sleepMs = Math.min(R.MAX_SLEEP_MS, Math.floor(R.BASE_SLEEP_MS * Math.pow(2, attempt - 1) + Math.random() * R.JITTER_MS))
    Logger.log(`[PostHog retry] ${label || 'query'} attempt ${attempt}/${R.MAX_ATTEMPTS} got ${code}. Sleeping ${sleepMs}ms`)
    Utilities.sleep(sleepMs)
  }

  throw lastErr
}
