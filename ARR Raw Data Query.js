/**************************************************************
 * ARR Raw Data query (PRODUCTION)
 *
 * The HogQL query used by render_arr_raw_data_view() (in
 * "ARR Waterfall Data.js") to build the LIVE arr_raw_data sheet
 * server-side from PostHog (users/orgs/subscriptions/stripe).
 *
 * Output columns (match ARR_RAW_CFG.HEADERS exactly):
 *   org_id, org_name, org_creation_date, first_payment_date, churn_date,
 *   sign_up_cohort_month, paid_cohort_month, current_status, ring_bucket,
 *   plan_name, billing_frequency, total_arr, subscription_start_date
 *
 * Date columns are formatted to 'YYYY-MM-DD HH:mm:ss' (project tz =
 * US/Mountain); cohorts as 'MMM yyyy' (so the waterfall parses them).
 *
 * Run via the pipeline (part 2) or render_arr_raw_data_view() directly.
 * Uses sauronQueryRun_ (defined in "Sauron Query.js").
 **************************************************************/

const ARR_RAW_DATA_QUERY_HOGQL = `
WITH
inv_full AS (SELECT id, subscription_id, created_at FROM stripe.invoice WHERE coalesce(billing_reason,'') IN ('subscription_cycle','subscription_create') AND status='paid' LIMIT 1 BY id),
latest_inv AS (SELECT subscription_id, argMax(id, created_at) AS inv_id FROM inv_full GROUP BY subscription_id),
rev AS (SELECT id, invoice_id, toFloat(amount) AS amt, toStartOfMonth(timestamp) AS mth FROM stripe.revenue_item_revenue_view WHERE is_recurring=1 LIMIT 1 BY id),
sub_rev AS (SELECT li.subscription_id AS sid, round(SUM(rev.amt)/nullIf(uniq(rev.mth),0)*12,2) AS arr FROM latest_inv li JOIN rev ON rev.invoice_id=li.inv_id GROUP BY li.subscription_id),
paid AS (SELECT subscription_id FROM stripe.invoice WHERE status='paid' AND toFloat(total)>0 AND subscription_id IS NOT NULL GROUP BY subscription_id),
sheet_excl AS (SELECT subscription_id FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) != 'managed'),
fp_org AS (SELECT org_id, min(i.created_at) AS first_payment_at FROM postgres.org_subscriptions os JOIN stripe.invoice i ON i.subscription_id=os.stripe_subscription_id WHERE i.status='paid' AND toFloat(i.total)>0 GROUP BY org_id),
org_arr AS (SELECT osb.org_id, round(sum(sr.arr),2) AS total_arr_paid FROM postgres.org_subscriptions osb JOIN sub_rev sr ON sr.sid=osb.stripe_subscription_id WHERE osb.stripe_subscription_id IN (SELECT subscription_id FROM paid) AND osb.stripe_subscription_id NOT IN (SELECT subscription_id FROM sheet_excl) AND osb.status!='canceled' AND sr.arr>0 GROUP BY osb.org_id),
picked AS (SELECT org_id, stripe_subscription_id, stripe_customer_id, status AS app_status, created_at AS sub_created, trial_ends_at, cancel_at_period_end, current_period_end FROM postgres.org_subscriptions ORDER BY (stripe_subscription_id IN (SELECT subscription_id FROM paid) AND status!='canceled') DESC, (status='active') DESC, (status='trialing' AND trial_ends_at>now()) DESC, created_at DESC LIMIT 1 BY org_id),
sx AS (SELECT id AS sub_id, start_date AS sub_start, coalesce(nullIf(plan.interval,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(items,''),'data'),1),'plan','interval')) AS intv, coalesce(nullIf(plan.product,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(items,''),'data'),1),'plan','product')) AS product_id FROM stripe.subscription LIMIT 1 BY id),
pm AS (SELECT DISTINCT customer_id FROM stripe.customerpaymentmethod),
prod AS (SELECT id, name FROM stripe.product LIMIT 1 BY id),
orgs AS (SELECT id, name, created_at FROM postgres.orgs LIMIT 1 BY id),
managed_ovr AS (SELECT subscription_id, toFloat64OrNull(replaceRegexpAll(coalesce(amount,''),'[^0-9.]','')) AS ovr_arr FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason)='managed' AND subscription_id IS NOT NULL),
assembled AS (
  SELECT o.id AS org_id, o.name AS org_name, o.created_at AS org_creation_date,
    fp_org.first_payment_at AS first_payment_date,
    if(picked.cancel_at_period_end, picked.current_period_end, NULL) AS churn_date,
    picked.app_status AS current_status, prod.name AS plan_name,
    multiIf(sx.intv='year','yearly', sx.intv='month','monthly', sx.intv) AS billing_frequency,
    coalesce(sx.sub_start, picked.sub_created) AS subscription_start_date,
    (pm.customer_id IS NOT NULL) AS has_pm,
    (fp_org.org_id IS NOT NULL OR (mo.subscription_id IS NOT NULL AND mo.ovr_arr>0)) AS has_fp,
    (picked.app_status IN ('active','trialing')) AS is_live,
    if(mo.subscription_id IS NOT NULL AND mo.ovr_arr IS NOT NULL, mo.ovr_arr, coalesce(org_arr.total_arr_paid,0)) AS arr_paid
  FROM orgs o
  LEFT JOIN picked  ON picked.org_id = o.id
  LEFT JOIN sx      ON sx.sub_id = picked.stripe_subscription_id
  LEFT JOIN fp_org  ON fp_org.org_id = o.id
  LEFT JOIN org_arr ON org_arr.org_id = o.id
  LEFT JOIN pm      ON pm.customer_id = picked.stripe_customer_id
  LEFT JOIN prod    ON prod.id = sx.product_id
  LEFT JOIN managed_ovr mo ON mo.subscription_id = picked.stripe_subscription_id
),
final AS (
  SELECT *,
    multiIf(has_fp AND arr_paid>=0.01,'paid', is_live AND has_pm,'intent_to_pay', is_live,'free_trial','') AS ring_bucket
  FROM assembled
)
SELECT
  final.org_id AS org_id,
  final.org_name AS org_name,
  formatDateTime(final.org_creation_date, '%Y-%m-%d %H:%i:%S') AS org_creation_date,
  if(final.first_payment_date IS NOT NULL, formatDateTime(final.first_payment_date, '%Y-%m-%d %H:%i:%S'), '') AS first_payment_date,
  if(final.churn_date IS NOT NULL, formatDateTime(final.churn_date, '%Y-%m-%d %H:%i:%S'), '') AS churn_date,
  formatDateTime(final.org_creation_date, '%b %Y') AS sign_up_cohort_month,
  if(final.ring_bucket='paid' AND final.first_payment_date IS NOT NULL, formatDateTime(final.first_payment_date,'%b %Y'), '') AS paid_cohort_month,
  final.current_status AS current_status,
  final.ring_bucket AS ring_bucket,
  final.plan_name AS plan_name,
  final.billing_frequency AS billing_frequency,
  if(final.ring_bucket='paid', final.arr_paid, 0) AS total_arr,
  formatDateTime(final.subscription_start_date, '%Y-%m-%d %H:%i:%S') AS subscription_start_date
FROM final
WHERE final.ring_bucket != ''
ORDER BY total_arr DESC
LIMIT 1000
`
