/**************************************************************
 * ARR Raw Data Query (TEST / verification)
 *
 * Runs a PostHog HogQL query that builds arr_raw_data server-side
 * (ARR per org from Stripe revenue items, ring bucket, cohorts)
 * and writes the result to a NEW sheet for verification against
 * the existing render_arr_raw_data_view pipeline before any swap.
 *
 * - Does NOT touch the real "arr_raw_data" sheet.
 * - Headers come straight from the query's returned columns.
 * - Run via menu: Ping Ops → "TEST: ARR Raw Data from PostHog query"
 *   or run render_arr_raw_data_query_test() directly in the editor.
 *
 * Script Properties used (same as the rest of the PostHog code):
 *  - POSTHOG_API_KEY
 *  - POSTHOG_PROJECT_ID (optional; falls back to POSTHOG_RAW_CFG)
 **************************************************************/

const ARR_RAW_DATA_QUERY_CFG = {
  OUT_SHEET: 'arr_raw_data (Query Test)'
}

// Query kept verbatim.
const ARR_RAW_DATA_QUERY_HOGQL = `
WITH
inv_full AS (SELECT id, subscription_id, created_at FROM stripe.invoice WHERE coalesce(billing_reason,'') IN ('subscription_cycle','subscription_create') AND status='paid' LIMIT 1 BY id),
latest_inv AS (SELECT subscription_id, argMax(id, created_at) AS inv_id FROM inv_full GROUP BY subscription_id),
rev AS (SELECT id, invoice_id, toFloat(amount) AS amt, toStartOfMonth(timestamp) AS mth FROM stripe.revenue_item_revenue_view WHERE is_recurring=1 LIMIT 1 BY id),
sub_rev AS (SELECT li.subscription_id AS sid, round(SUM(rev.amt)/nullIf(uniq(rev.mth),0)*12,2) AS arr FROM latest_inv li JOIN rev ON rev.invoice_id=li.inv_id GROUP BY li.subscription_id),
paid AS (SELECT subscription_id FROM stripe.invoice WHERE status='paid' AND toFloat(total)>0 AND subscription_id IS NOT NULL GROUP BY subscription_id),
sheet_excl AS (SELECT subscription_id FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) != 'managed'),
ss AS (SELECT id, status, canceled_at, customer_id FROM stripe.subscription LIMIT 1 BY id),
fp_org AS (SELECT os2.org_id AS org_id, min(i.created_at) AS first_payment_at FROM postgres.org_subscriptions os2 JOIN stripe.invoice i ON i.subscription_id=os2.stripe_subscription_id WHERE i.status='paid' AND toFloat(i.total)>0 GROUP BY os2.org_id),
org_arr AS (SELECT osb.org_id AS org_id, round(sum(sr.arr),2) AS total_arr_paid FROM postgres.org_subscriptions osb JOIN ss ON ss.id=osb.stripe_subscription_id JOIN sub_rev sr ON sr.sid=osb.stripe_subscription_id WHERE osb.stripe_subscription_id IN (SELECT subscription_id FROM paid) AND osb.stripe_subscription_id NOT IN (SELECT subscription_id FROM sheet_excl) AND ss.status='active' AND sr.arr>0 GROUP BY osb.org_id),
picked AS (SELECT org_id, stripe_subscription_id, stripe_customer_id, status AS app_status, created_at AS sub_created, trial_ends_at FROM postgres.org_subscriptions ORDER BY (status='active') DESC, (status='trialing' AND trial_ends_at>now()) DESC, created_at DESC LIMIT 1 BY org_id),
sx AS (SELECT id AS sub_id, status AS stripe_status, canceled_at, coalesce(nullIf(plan.interval,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(items,''),'data'),1),'plan','interval')) AS intv, coalesce(nullIf(plan.product,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(items,''),'data'),1),'plan','product')) AS product_id FROM stripe.subscription LIMIT 1 BY id),
pm AS (SELECT DISTINCT customer_id FROM stripe.customerpaymentmethod),
prod AS (SELECT id, name FROM stripe.product LIMIT 1 BY id),
orgs AS (SELECT id, name, created_at FROM postgres.orgs LIMIT 1 BY id),
assembled AS (
  SELECT o.id AS org_id, o.name AS org_name, o.created_at AS org_creation_date,
    fp_org.first_payment_at AS first_payment_date, sx.canceled_at AS churn_date,
    sx.stripe_status AS current_status, prod.name AS plan_name,
    multiIf(sx.intv='year','yearly', sx.intv='month','monthly', sx.intv) AS billing_frequency,
    picked.sub_created AS subscription_start_date,
    (pm.customer_id IS NOT NULL) AS has_pm,
    coalesce(org_arr.total_arr_paid,0) AS arr_paid,
    multiIf(
      sx.stripe_status='active' AND fp_org.first_payment_at IS NOT NULL, 'paid',
      sx.stripe_status='active' AND pm.customer_id IS NOT NULL, 'intent_to_pay',
      sx.stripe_status='trialing' AND pm.customer_id IS NOT NULL, 'intent_to_pay',
      sx.stripe_status='trialing', 'free_trial', '') AS bucket0
  FROM orgs o
  LEFT JOIN picked  ON picked.org_id = o.id
  LEFT JOIN sx      ON sx.sub_id = picked.stripe_subscription_id
  LEFT JOIN fp_org  ON fp_org.org_id = o.id
  LEFT JOIN org_arr ON org_arr.org_id = o.id
  LEFT JOIN pm      ON pm.customer_id = picked.stripe_customer_id
  LEFT JOIN prod    ON prod.id = sx.product_id
)
SELECT
  org_id, org_name, org_creation_date, first_payment_date, churn_date,
  formatDateTime(org_creation_date, '%b %Y') AS sign_up_cohort_month,
  if(ring_bucket='paid' AND first_payment_date IS NOT NULL, formatDateTime(first_payment_date,'%b %Y'), '') AS paid_cohort_month,
  current_status, ring_bucket, plan_name, billing_frequency, total_arr, subscription_start_date
FROM (
  SELECT *,
    multiIf(bucket0='paid' AND arr_paid<0.01, if(has_pm,'intent_to_pay','free_trial'), bucket0) AS ring_bucket,
    multiIf(bucket0='paid' AND arr_paid>=0.01, arr_paid, 0) AS total_arr
  FROM assembled
) AS t
WHERE ring_bucket != ''
ORDER BY total_arr DESC
LIMIT 1000
`

/**
 * Run the arr_raw_data query and dump it to the test sheet.
 * Returns { rows_in, rows_out } for pipeline-style logging.
 */
function render_arr_raw_data_query_test() {
  const t0 = new Date()
  const ss = SpreadsheetApp.getActive()

  const props = PropertiesService.getScriptProperties()
  const apiKey = props.getProperty('POSTHOG_API_KEY')
  if (!apiKey) throw new Error('Missing POSTHOG_API_KEY in Script Properties')
  const projectId = props.getProperty('POSTHOG_PROJECT_ID') || POSTHOG_RAW_CFG.PROJECT_ID_FALLBACK

  // Reuse the columns-aware HogQL runner (defined in Sauron Query.js).
  const { columns, results } = sauronQueryRun_(apiKey, projectId, ARR_RAW_DATA_QUERY_HOGQL, 'arr_raw_data_query_test')

  const headers = (columns && columns.length) ? columns : ['(no columns returned)']

  const rows = (results || []).map(r => {
    const row = new Array(headers.length)
    for (let i = 0; i < headers.length; i++) row[i] = (r && r[i] != null) ? r[i] : ''
    return row
  })

  const sh = ss.getSheetByName(ARR_RAW_DATA_QUERY_CFG.OUT_SHEET) || ss.insertSheet(ARR_RAW_DATA_QUERY_CFG.OUT_SHEET)
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

  Logger.log(`render_arr_raw_data_query_test: ${rows.length} rows in ${((new Date() - t0) / 1000).toFixed(1)}s`)
  return { rows_in: rows.length, rows_out: rows.length }
}
