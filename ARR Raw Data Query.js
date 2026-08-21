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
excl_emails AS (SELECT DISTINCT lower(customer_email) AS email FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) IN ('internal','duplicate','fake') AND coalesce(customer_email,'')!=''),
founder_x AS (SELECT uo.org_id AS org_id, argMin(lower(u.email), uo.created_at) AS email FROM postgres.users_orgs uo JOIN postgres.users u ON u.id=uo.user_id GROUP BY uo.org_id),
excluded_orgs AS (SELECT DISTINCT os.org_id AS org_id FROM postgres.org_subscriptions os JOIN (SELECT id, lower(email) AS email FROM stripe.customer LIMIT 1 BY id) c ON c.id=os.stripe_customer_id WHERE c.email IN (SELECT email FROM excl_emails) UNION DISTINCT SELECT org_id FROM founder_x WHERE email IN (SELECT email FROM excl_emails)),
-- ARR basis = CURRENT subscription items (contracted run-rate), annualized, net of active
-- discount. Replaces the old trailing "latest full-cycle invoice" method, which lagged
-- mid-cycle upgrades/downgrades (e.g. an annual seat add wouldn't show until next renewal).
-- Discount netting: percent_off/amount_off from live coupons; 'once' coupons ignored
-- (one-time, not recurring); expired coupons are already gone from subscription.discounts.
sxi AS (SELECT id AS sid, coalesce(nullIf(plan.interval,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(toString(items),''),'data'),1),'plan','interval')) AS intv, arraySum(arrayMap(x -> JSONExtractInt(x,'plan','amount')*JSONExtractInt(x,'quantity'), JSONExtractArrayRaw(coalesce(toString(items),''),'data')))/100.0 AS period_amt, arraySum(arrayMap(d -> if(JSONExtractString(d,'coupon','duration')!='once', toFloatOrZero(JSONExtractRaw(d,'coupon','percent_off')), 0.0), JSONExtractArrayRaw(coalesce(toString(discounts),'[]')))) AS pct_off, arraySum(arrayMap(d -> if(JSONExtractString(d,'coupon','duration')!='once', toFloatOrZero(JSONExtractRaw(d,'coupon','amount_off')), 0.0), JSONExtractArrayRaw(coalesce(toString(discounts),'[]'))))/100.0 AS amt_off_period FROM stripe.subscription LIMIT 1 BY id),
ci AS (SELECT sid, round(greatest(0, (period_amt*if(intv='month',12,1))*(1-least(pct_off,100.0)/100.0) - amt_off_period*if(intv='month',12,1)),2) AS arr FROM sxi),
paid AS (SELECT subscription_id FROM stripe.invoice WHERE status='paid' AND toFloat(total)>0 AND subscription_id IS NOT NULL GROUP BY subscription_id),
-- Map an override-sheet org_id to the internal org id (internal DB id OR app/WorkOS id).
org_keys AS (SELECT id AS org_id, toString(id) AS k FROM postgres.orgs LIMIT 1 BY id UNION ALL SELECT id AS org_id, coalesce(nullIf(workos_id,''), external_id) AS k FROM postgres.orgs LIMIT 1 BY id),
-- Non-managed exclusions (ARR not counted), keyed on the sheet's customer_email.
sheet_excl AS (SELECT DISTINCT os.org_id FROM postgres.org_subscriptions os JOIN (SELECT id, lower(email) AS email FROM stripe.customer LIMIT 1 BY id) c ON c.id = os.stripe_customer_id WHERE c.email IN (SELECT lower(customer_email) FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) != 'managed' AND coalesce(customer_email,'') != '')),
fp_org AS (SELECT org_id, min(i.created_at) AS first_payment_at FROM postgres.org_subscriptions os JOIN stripe.invoice i ON i.subscription_id=os.stripe_subscription_id WHERE i.status='paid' AND toFloat(i.total)>0 GROUP BY org_id),
org_arr AS (SELECT osb.org_id, round(sum(ci.arr),2) AS total_arr_paid FROM postgres.org_subscriptions osb JOIN ci ON ci.sid=osb.stripe_subscription_id WHERE osb.stripe_subscription_id IN (SELECT subscription_id FROM paid) AND osb.org_id NOT IN (SELECT org_id FROM sheet_excl) AND osb.status!='canceled' AND ci.arr>0 GROUP BY osb.org_id),
picked AS (SELECT org_id, stripe_subscription_id, stripe_customer_id, status AS app_status, created_at AS sub_created, trial_ends_at, cancel_at_period_end, current_period_end FROM postgres.org_subscriptions ORDER BY (stripe_subscription_id IN (SELECT subscription_id FROM paid) AND status!='canceled') DESC, (status='active') DESC, (status='trialing' AND trial_ends_at>now()) DESC, created_at DESC LIMIT 1 BY org_id),
sx AS (SELECT id AS sub_id, start_date AS sub_start, coalesce(nullIf(plan.interval,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(items,''),'data'),1),'plan','interval')) AS intv, coalesce(nullIf(plan.product,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(items,''),'data'),1),'plan','product')) AS product_id FROM stripe.subscription LIMIT 1 BY id),
pm AS (SELECT DISTINCT customer_id FROM stripe.customerpaymentmethod),
prod AS (SELECT id, name FROM stripe.product LIMIT 1 BY id),
orgs AS (SELECT id, name, created_at FROM postgres.orgs LIMIT 1 BY id),
-- MANAGED orgs: keyed on the sheet's org_id. ARR = 'amount' (injected directly, no Stripe sub/invoice needed).
managed AS (SELECT ok.org_id AS org_id, max(toFloat64OrNull(replaceRegexpAll(coalesce(toString(ovr.amount),''),'[^0-9.]',''))) AS amt, max(toFloat64OrNull(replaceRegexpAll(coalesce(toString(ovr.full_seats),''),'[^0-9.]',''))) AS full_seats, max(toFloat64OrNull(replaceRegexpAll(coalesce(toString(ovr.lite_seats),''),'[^0-9.]',''))) AS lite_seats FROM override_googlesheets_manual_stripe_changes ovr JOIN org_keys ok ON ok.k = toString(ovr.org_id) WHERE lower(ovr.exclude_reason)='managed' AND coalesce(toString(ovr.org_id),'') != '' GROUP BY ok.org_id),
assembled AS (
  SELECT o.id AS org_id, o.name AS org_name, o.created_at AS org_creation_date,
    fp_org.first_payment_at AS first_payment_date,
    if(picked.cancel_at_period_end, picked.current_period_end, NULL) AS churn_date,
    if(mg.org_id IS NOT NULL AND picked.app_status IS NULL, 'active', picked.app_status) AS current_status, prod.name AS plan_name,
    multiIf(sx.intv='year','yearly', sx.intv='month','monthly', sx.intv) AS billing_frequency,
    coalesce(sx.sub_start, picked.sub_created, o.created_at) AS subscription_start_date,
    (pm.customer_id IS NOT NULL) AS has_pm,
    (fp_org.org_id IS NOT NULL OR mg.org_id IS NOT NULL) AS has_fp,
    (picked.app_status IN ('active','trialing')) AS is_live,
    if(mg.org_id IS NOT NULL, coalesce(mg.amt,0), coalesce(org_arr.total_arr_paid,0)) AS arr_paid
  FROM orgs o
  LEFT JOIN picked  ON picked.org_id = o.id
  LEFT JOIN sx      ON sx.sub_id = picked.stripe_subscription_id
  LEFT JOIN fp_org  ON fp_org.org_id = o.id
  LEFT JOIN org_arr ON org_arr.org_id = o.id
  LEFT JOIN pm      ON pm.customer_id = picked.stripe_customer_id
  LEFT JOIN prod    ON prod.id = sx.product_id
  LEFT JOIN managed mg ON mg.org_id = o.id
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
  AND final.org_id NOT IN (SELECT org_id FROM excluded_orgs)
ORDER BY total_arr DESC
LIMIT 5000
`
