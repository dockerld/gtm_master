/**************************************************************
 * Ring Query (TEST / verification)
 *
 * Two PostHog HogQL queries that rebuild "The Ring":
 *   (A) RING SUMMARY  -> 3 rows: Paid / Intent to Pay / Trialing
 *                        (arr, subscriptions, total_seats)
 *   (B) RING DETAIL   -> one row per org (all three statuses)
 *
 * Writes both to a NEW "The Ring (Query Test)" sheet (summary block on
 * top, detail block below) for verification against the live render_ring_view
 * before any swap. Does NOT touch the real "The Ring" sheet.
 *
 * Conventions: status from app DB; Paid ARR = net run-rate from latest
 * full-cycle invoice; Intent/Trialing ARR = plan list-price annualized.
 *
 * Run via menu: Ping Ops -> "TEST: The Ring from PostHog query"
 *   or run render_ring_query_test() directly in the editor.
 * Uses sauronQueryRun_ (defined in "Sauron Query.js").
 **************************************************************/

const RING_QUERY_CFG = {
  OUT_SHEET: 'The Ring (Query Test)'
}

// (A) RING SUMMARY — 3 rows: Paid / Intent to Pay / Trialing
const RING_SUMMARY_HOGQL = `
WITH
inv_full AS (SELECT id, subscription_id, created_at FROM stripe.invoice WHERE coalesce(billing_reason,'') IN ('subscription_cycle','subscription_create') AND status='paid' LIMIT 1 BY id),
latest_inv AS (SELECT subscription_id, argMax(id, created_at) AS inv_id FROM inv_full GROUP BY subscription_id),
rev AS (SELECT id, invoice_id, toFloat(amount) AS amt, toStartOfMonth(timestamp) AS mth FROM stripe.revenue_item_revenue_view WHERE is_recurring=1 LIMIT 1 BY id),
sub_rev AS (SELECT li.subscription_id AS sid, round(SUM(rev.amt)/nullIf(uniq(rev.mth),0)*12,2) AS arr FROM latest_inv li JOIN rev ON rev.invoice_id=li.inv_id GROUP BY li.subscription_id),
paid AS (SELECT subscription_id FROM stripe.invoice WHERE status='paid' AND toFloat(total)>0 AND subscription_id IS NOT NULL GROUP BY subscription_id),
sheet_excl AS (SELECT subscription_id FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) != 'managed'),
fp_org AS (SELECT org_id FROM postgres.org_subscriptions os JOIN stripe.invoice i ON i.subscription_id=os.stripe_subscription_id WHERE i.status='paid' AND toFloat(i.total)>0 GROUP BY org_id),
org_arr AS (SELECT osb.org_id, round(sum(sr.arr),2) AS arr_actual FROM postgres.org_subscriptions osb JOIN sub_rev sr ON sr.sid=osb.stripe_subscription_id WHERE osb.stripe_subscription_id IN (SELECT subscription_id FROM paid) AND osb.stripe_subscription_id NOT IN (SELECT subscription_id FROM sheet_excl) AND osb.status!='canceled' AND sr.arr>0 GROUP BY osb.org_id),
picked AS (SELECT org_id, stripe_subscription_id, stripe_customer_id, status AS app_status, trial_ends_at, full_seat_count, lite_seat_count FROM postgres.org_subscriptions ORDER BY (stripe_subscription_id IN (SELECT subscription_id FROM paid) AND status!='canceled') DESC, (status='active') DESC, (status='trialing' AND trial_ends_at>now()) DESC, created_at DESC LIMIT 1 BY org_id),
sx AS (SELECT id AS sub_id, coalesce(nullIf(plan.interval,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(items,''),'data'),1),'plan','interval')) AS intv, arraySum(arrayMap(x -> JSONExtractInt(x,'plan','amount')*JSONExtractInt(x,'quantity'), JSONExtractArrayRaw(coalesce(items,''),'data')))/100.0 AS period_amt FROM stripe.subscription LIMIT 1 BY id),
pm AS (SELECT DISTINCT customer_id FROM stripe.customerpaymentmethod),
per_org AS (SELECT o.id AS org_id, coalesce(picked.full_seat_count,0)+coalesce(picked.lite_seat_count,0) AS seats, (pm.customer_id IS NOT NULL) AS has_pm, (fp_org.org_id IS NOT NULL) AS has_fp, (picked.app_status IN ('active','trialing')) AS is_live, coalesce(org_arr.arr_actual,0) AS arr_actual, round(coalesce(sx.period_amt,0)*if(sx.intv='month',12,1),2) AS arr_potential FROM (SELECT id FROM postgres.orgs LIMIT 1 BY id) o LEFT JOIN picked ON picked.org_id=o.id LEFT JOIN sx ON sx.sub_id=picked.stripe_subscription_id LEFT JOIN fp_org ON fp_org.org_id=o.id LEFT JOIN org_arr ON org_arr.org_id=o.id LEFT JOIN pm ON pm.customer_id=picked.stripe_customer_id),
staged AS (SELECT *, multiIf(has_fp AND arr_actual>=0.01,'Paid', is_live AND has_pm,'Intent to Pay', is_live,'Trialing','') AS status, if(has_fp AND arr_actual>=0.01, arr_actual, arr_potential) AS arr_row FROM per_org)
SELECT status, round(sum(arr_row),2) AS arr, count() AS subscriptions, sum(seats) AS total_seats
FROM staged WHERE status!='' GROUP BY status
ORDER BY multiIf(status='Paid',1, status='Intent to Pay',2, 3)
`

// (B) RING DETAIL — one row per org; all Paid + Intent to Pay + Trialing
const RING_DETAIL_HOGQL = `
WITH
inv_full AS (SELECT id, subscription_id, created_at FROM stripe.invoice WHERE coalesce(billing_reason,'') IN ('subscription_cycle','subscription_create') AND status='paid' LIMIT 1 BY id),
latest_inv AS (SELECT subscription_id, argMax(id, created_at) AS inv_id FROM inv_full GROUP BY subscription_id),
rev AS (SELECT id, invoice_id, toFloat(amount) AS amt, toStartOfMonth(timestamp) AS mth FROM stripe.revenue_item_revenue_view WHERE is_recurring=1 LIMIT 1 BY id),
sub_rev AS (SELECT li.subscription_id AS sid, round(SUM(rev.amt)/nullIf(uniq(rev.mth),0)*12,2) AS arr FROM latest_inv li JOIN rev ON rev.invoice_id=li.inv_id GROUP BY li.subscription_id),
paid AS (SELECT subscription_id FROM stripe.invoice WHERE status='paid' AND toFloat(total)>0 AND subscription_id IS NOT NULL GROUP BY subscription_id),
sheet_excl AS (SELECT subscription_id FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) != 'managed'),
fp_org AS (SELECT org_id, min(i.created_at) AS first_payment FROM postgres.org_subscriptions os JOIN stripe.invoice i ON i.subscription_id=os.stripe_subscription_id WHERE i.status='paid' AND toFloat(i.total)>0 GROUP BY org_id),
org_arr AS (SELECT osb.org_id, round(sum(sr.arr),2) AS arr_actual FROM postgres.org_subscriptions osb JOIN sub_rev sr ON sr.sid=osb.stripe_subscription_id WHERE osb.stripe_subscription_id IN (SELECT subscription_id FROM paid) AND osb.stripe_subscription_id NOT IN (SELECT subscription_id FROM sheet_excl) AND osb.status!='canceled' AND sr.arr>0 GROUP BY osb.org_id),
picked AS (SELECT org_id, stripe_subscription_id, stripe_customer_id, status AS app_status, trial_ends_at, full_seat_count, lite_seat_count FROM postgres.org_subscriptions ORDER BY (stripe_subscription_id IN (SELECT subscription_id FROM paid) AND status!='canceled') DESC, (status='active') DESC, (status='trialing' AND trial_ends_at>now()) DESC, created_at DESC LIMIT 1 BY org_id),
sx AS (SELECT id AS sub_id, coalesce(nullIf(plan.interval,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(items,''),'data'),1),'plan','interval')) AS intv, arraySum(arrayMap(x -> JSONExtractInt(x,'plan','amount')*JSONExtractInt(x,'quantity'), JSONExtractArrayRaw(coalesce(items,''),'data')))/100.0 AS period_amt FROM stripe.subscription LIMIT 1 BY id),
pm AS (SELECT DISTINCT customer_id FROM stripe.customerpaymentmethod),
cust AS (SELECT id, email, name FROM stripe.customer LIMIT 1 BY id),
orgs AS (SELECT id, name, created_at FROM postgres.orgs LIMIT 1 BY id),
per_org AS (SELECT o.id AS org_id, o.name AS org_name, o.created_at AS sign_up_date, cust.email AS customer_email, cust.name AS customer_name, fp_org.first_payment AS first_payment_at, picked.trial_ends_at AS trial_ends_at, sx.intv AS interval, coalesce(picked.full_seat_count,0)+coalesce(picked.lite_seat_count,0) AS seats, (pm.customer_id IS NOT NULL) AS has_pm, (fp_org.org_id IS NOT NULL) AS has_fp, (picked.app_status IN ('active','trialing')) AS is_live, coalesce(org_arr.arr_actual,0) AS arr_actual, round(coalesce(sx.period_amt,0)*if(sx.intv='month',12,1),2) AS arr_potential FROM orgs o LEFT JOIN picked ON picked.org_id=o.id LEFT JOIN sx ON sx.sub_id=picked.stripe_subscription_id LEFT JOIN fp_org ON fp_org.org_id=o.id LEFT JOIN org_arr ON org_arr.org_id=o.id LEFT JOIN pm ON pm.customer_id=picked.stripe_customer_id LEFT JOIN cust ON cust.id=picked.stripe_customer_id),
staged AS (SELECT *, multiIf(has_fp AND arr_actual>=0.01,'Paid', is_live AND has_pm,'Intent to Pay', is_live,'Trialing','') AS status FROM per_org)
SELECT customer_email, customer_name, org_name, status,
       formatDateTime(sign_up_date,'%Y-%m-%d') AS sign_up_date,
       if(first_payment_at IS NOT NULL, formatDateTime(first_payment_at,'%Y-%m-%d'), '') AS first_payment_at,
       if(status!='Paid' AND trial_ends_at>now(), dateDiff('day', now(), trial_ends_at), NULL) AS trial_days_remaining,
       interval,
       if(status='Paid', arr_actual, arr_potential) AS arr,
       seats
FROM staged WHERE status!=''
ORDER BY multiIf(status='Paid',1, status='Intent to Pay',2, 3), arr DESC
`

/**
 * Run both Ring queries and write them to the test sheet:
 * summary block on top, then a gap, then the detail block.
 */
function render_ring_query_test() {
  const t0 = new Date()
  const ss = SpreadsheetApp.getActive()

  const props = PropertiesService.getScriptProperties()
  const apiKey = props.getProperty('POSTHOG_API_KEY')
  if (!apiKey) throw new Error('Missing POSTHOG_API_KEY in Script Properties')
  const projectId = props.getProperty('POSTHOG_PROJECT_ID') || POSTHOG_RAW_CFG.PROJECT_ID_FALLBACK

  const summary = sauronQueryRun_(apiKey, projectId, RING_SUMMARY_HOGQL, 'ring_summary')
  const detail = sauronQueryRun_(apiKey, projectId, RING_DETAIL_HOGQL, 'ring_detail')

  const sh = ss.getSheetByName(RING_QUERY_CFG.OUT_SHEET) || ss.insertSheet(RING_QUERY_CFG.OUT_SHEET)
  sh.clear()

  const norm = (results, width) => (results || []).map(r => {
    const row = new Array(width)
    for (let i = 0; i < width; i++) row[i] = (r && r[i] != null) ? r[i] : ''
    return row
  })

  // --- Summary block ---
  const sumHeaders = (summary.columns && summary.columns.length) ? summary.columns : ['(no columns)']
  let row = 1
  sh.getRange(row, 1, 1, 1).setValues([['SUMMARY']]).setFontWeight('bold')
  row++
  sh.getRange(row, 1, 1, sumHeaders.length).setValues([sumHeaders]).setFontWeight('bold').setBackground('#F3F4F6')
  row++
  const sumRows = norm(summary.results, sumHeaders.length)
  if (sumRows.length) { sh.getRange(row, 1, sumRows.length, sumHeaders.length).setValues(sumRows); row += sumRows.length }

  // gap
  row += 2

  // --- Detail block ---
  const detHeaders = (detail.columns && detail.columns.length) ? detail.columns : ['(no columns)']
  sh.getRange(row, 1, 1, 1).setValues([['DETAIL']]).setFontWeight('bold')
  row++
  sh.getRange(row, 1, 1, detHeaders.length).setValues([detHeaders]).setFontWeight('bold').setBackground('#F3F4F6')
  row++
  const detRows = norm(detail.results, detHeaders.length)
  if (detRows.length) {
    const chunk = 5000
    for (let i = 0; i < detRows.length; i += chunk) {
      const part = detRows.slice(i, i + chunk)
      sh.getRange(row + i, 1, part.length, detHeaders.length).setValues(part)
    }
  }
  try { sh.autoResizeColumns(1, Math.max(detHeaders.length, sumHeaders.length)) } catch (e) {}

  Logger.log(`render_ring_query_test: summary=${sumRows.length} detail=${detRows.length} in ${((new Date() - t0) / 1000).toFixed(1)}s`)
  return { rows_in: detRows.length, rows_out: detRows.length }
}
