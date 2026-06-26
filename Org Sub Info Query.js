/**************************************************************
 * Org Subscription Info Query (TEST / verification)
 *
 * Runs a PostHog HogQL query that selects the latest/active
 * subscription per org (server-side) and writes the result to a
 * NEW sheet for verification against the existing
 * render_org_subscription_info pipeline before any swap.
 *
 * - Does NOT touch the real "org_subscription_info" sheet.
 * - Headers come straight from the query's returned columns.
 * - Run via menu: Ping Ops → "TEST: Org Sub Info from PostHog query"
 *   or run render_org_sub_info_query_test() directly in the editor.
 *
 * Script Properties used (same as the rest of the PostHog code):
 *  - POSTHOG_API_KEY
 *  - POSTHOG_PROJECT_ID (optional; falls back to POSTHOG_RAW_CFG)
 **************************************************************/

const ORG_SUB_INFO_QUERY_CFG = {
  OUT_SHEET: 'org_subscription_info (Query Test)'
}

// CTE kept verbatim; wrapped with a final SELECT so it runs standalone.
const ORG_SUB_INFO_QUERY_HOGQL = `
WITH os AS (
  SELECT org_id, id AS latest_subscription_id, stripe_subscription_id, status AS app_status,
         created_at AS sub_created, trial_ends_at, current_period_end, cancel_at_period_end, stripe_customer_id
  FROM postgres.org_subscriptions
  ORDER BY (status = 'active') DESC,                              -- 1. active wins
           (status = 'trialing' AND trial_ends_at > now()) DESC,  -- 2. then a real (non-stale) trial
           created_at DESC                                        -- 3. else latest
  LIMIT 1 BY org_id
)
SELECT org_id, latest_subscription_id, stripe_subscription_id, app_status,
       sub_created, trial_ends_at, current_period_end, cancel_at_period_end, stripe_customer_id
FROM os
ORDER BY org_id
LIMIT 5000
`

/**
 * Run the org-subscription query and dump it to the test sheet.
 * Returns { rows_in, rows_out } for pipeline-style logging.
 */
function render_org_sub_info_query_test() {
  const t0 = new Date()
  const ss = SpreadsheetApp.getActive()

  const props = PropertiesService.getScriptProperties()
  const apiKey = props.getProperty('POSTHOG_API_KEY')
  if (!apiKey) throw new Error('Missing POSTHOG_API_KEY in Script Properties')
  const projectId = props.getProperty('POSTHOG_PROJECT_ID') || POSTHOG_RAW_CFG.PROJECT_ID_FALLBACK

  // Reuse the columns-aware HogQL runner (defined in Sauron Query.js).
  const { columns, results } = sauronQueryRun_(apiKey, projectId, ORG_SUB_INFO_QUERY_HOGQL, 'org_sub_info_query_test')

  const headers = (columns && columns.length) ? columns : ['(no columns returned)']

  const rows = (results || []).map(r => {
    const row = new Array(headers.length)
    for (let i = 0; i < headers.length; i++) row[i] = (r && r[i] != null) ? r[i] : ''
    return row
  })

  const sh = ss.getSheetByName(ORG_SUB_INFO_QUERY_CFG.OUT_SHEET) || ss.insertSheet(ORG_SUB_INFO_QUERY_CFG.OUT_SHEET)
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

  Logger.log(`render_org_sub_info_query_test: ${rows.length} rows in ${((new Date() - t0) / 1000).toFixed(1)}s`)
  return { rows_in: rows.length, rows_out: rows.length }
}
