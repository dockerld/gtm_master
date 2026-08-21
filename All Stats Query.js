/**************************************************************
 * All the Stats (PRODUCTION) — PostHog HogQL, multi-section
 *
 * Builds the "All the Stats" sheet from a set of self-contained HogQL
 * queries (one per section), written as labeled blocks top-to-bottom.
 * No canon tables / raw pulls — straight from PostHog (postgres + stripe),
 * plus the synced snapshot/waterfall mirrors for the retention sections.
 *
 * Sections:
 *   1. ARR / MRR Summary by Stage
 *   3. Avg Revenue Metrics (Paid)
 *   4. Churned / Scheduled-Cancel Audit
 *   5. Trialing (No Payment Method)
 *   6. Conversion (signup->paid, promo->paid)
 *   7. NRR / GRR (from override_googlesheets_arr_snapshot)
 *   8. Net New ARR by Month / Waterfall (from override_googlesheets_arr_waterfall_facts)
 *
 * Run via pipeline or render_all_stats_view() in the editor.
 * Uses sauronQueryRun_ (defined in "Sauron Query.js").
 **************************************************************/

const ALL_STATS_CFG = {
  OUT_SHEET: 'All the Stats'
}

const ALLSTATS_SECTION_1 = `
WITH
sxi AS (SELECT id AS sid, coalesce(nullIf(plan.interval,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(toString(items),''),'data'),1),'plan','interval')) AS intv, arraySum(arrayMap(x -> JSONExtractInt(x,'plan','amount')*JSONExtractInt(x,'quantity'), JSONExtractArrayRaw(coalesce(toString(items),''),'data')))/100.0 AS period_amt, arraySum(arrayMap(d -> if(JSONExtractString(d,'coupon','duration')!='once', toFloatOrZero(JSONExtractRaw(d,'coupon','percent_off')), 0.0), JSONExtractArrayRaw(coalesce(toString(discounts),'[]')))) AS pct_off, arraySum(arrayMap(d -> if(JSONExtractString(d,'coupon','duration')!='once', toFloatOrZero(JSONExtractRaw(d,'coupon','amount_off')), 0.0), JSONExtractArrayRaw(coalesce(toString(discounts),'[]'))))/100.0 AS amt_off_period FROM stripe.subscription LIMIT 1 BY id),
sub_rev AS (SELECT sid, round(greatest(0, (period_amt*if(intv='month',12,1))*(1-least(pct_off,100.0)/100.0) - amt_off_period*if(intv='month',12,1)),2) AS arr FROM sxi),
paid AS (SELECT subscription_id FROM stripe.invoice WHERE status='paid' AND toFloat(total)>0 AND subscription_id IS NOT NULL GROUP BY subscription_id),
org_keys AS (SELECT id AS org_id, toString(id) AS k FROM postgres.orgs LIMIT 1 BY id UNION ALL SELECT id AS org_id, coalesce(nullIf(workos_id,''), external_id) AS k FROM postgres.orgs LIMIT 1 BY id),
sheet_excl AS (SELECT DISTINCT os.org_id FROM postgres.org_subscriptions os JOIN (SELECT id, lower(email) AS email FROM stripe.customer LIMIT 1 BY id) c ON c.id = os.stripe_customer_id WHERE c.email IN (SELECT lower(customer_email) FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) != 'managed' AND coalesce(customer_email,'') != '')),
fp_org AS (SELECT org_id FROM postgres.org_subscriptions os JOIN stripe.invoice i ON i.subscription_id=os.stripe_subscription_id WHERE i.status='paid' AND toFloat(i.total)>0 GROUP BY org_id),
managed AS (SELECT ok.org_id AS org_id, max(toFloat64OrNull(replaceRegexpAll(coalesce(toString(ovr.amount),''),'[^0-9.]',''))) AS amt, max(toFloat64OrNull(replaceRegexpAll(coalesce(toString(ovr.full_seats),''),'[^0-9.]',''))) AS full_seats, max(toFloat64OrNull(replaceRegexpAll(coalesce(toString(ovr.lite_seats),''),'[^0-9.]',''))) AS lite_seats FROM override_googlesheets_manual_stripe_changes ovr JOIN org_keys ok ON ok.k = toString(ovr.org_id) WHERE lower(ovr.exclude_reason)='managed' AND coalesce(toString(ovr.org_id),'') != '' GROUP BY ok.org_id),
org_arr AS (SELECT osb.org_id, round(sum(sr.arr),2) AS arr_actual FROM postgres.org_subscriptions osb JOIN sub_rev sr ON sr.sid=osb.stripe_subscription_id WHERE osb.stripe_subscription_id IN (SELECT subscription_id FROM paid) AND osb.org_id NOT IN (SELECT org_id FROM sheet_excl) AND osb.status!='canceled' AND sr.arr>0 GROUP BY osb.org_id),
picked AS (SELECT org_id, stripe_subscription_id, stripe_customer_id, status AS app_status, trial_ends_at, full_seat_count, lite_seat_count FROM postgres.org_subscriptions ORDER BY (stripe_subscription_id IN (SELECT subscription_id FROM paid) AND status!='canceled') DESC, (status='active') DESC, (status='trialing' AND trial_ends_at>now()) DESC, created_at DESC LIMIT 1 BY org_id),
sx AS (SELECT id AS sub_id, coalesce(nullIf(plan.interval,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(items,''),'data'),1),'plan','interval')) AS intv, arraySum(arrayMap(x -> JSONExtractInt(x,'plan','amount')*JSONExtractInt(x,'quantity'), JSONExtractArrayRaw(coalesce(items,''),'data')))/100.0 AS period_amt FROM stripe.subscription LIMIT 1 BY id),
pm AS (SELECT DISTINCT customer_id FROM stripe.customerpaymentmethod),
excl_emails AS (SELECT DISTINCT lower(customer_email) AS email FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) IN ('internal','duplicate','fake') AND coalesce(customer_email,'')!=''),
founder_x AS (SELECT uo.org_id AS org_id, argMin(lower(u.email), uo.created_at) AS email FROM postgres.users_orgs uo JOIN postgres.users u ON u.id=uo.user_id GROUP BY uo.org_id),
excluded_orgs AS (SELECT DISTINCT os.org_id AS org_id FROM postgres.org_subscriptions os JOIN (SELECT id, lower(email) AS email FROM stripe.customer LIMIT 1 BY id) c ON c.id=os.stripe_customer_id WHERE c.email IN (SELECT email FROM excl_emails) UNION DISTINCT SELECT org_id FROM founder_x WHERE email IN (SELECT email FROM excl_emails)),
per_org AS (SELECT o.id AS org_id, if(mg.org_id IS NOT NULL AND (coalesce(mg.full_seats,0)+coalesce(mg.lite_seats,0))>0, coalesce(mg.full_seats,0), coalesce(picked.full_seat_count,0)) AS full_seats, if(mg.org_id IS NOT NULL AND (coalesce(mg.full_seats,0)+coalesce(mg.lite_seats,0))>0, coalesce(mg.lite_seats,0), coalesce(picked.lite_seat_count,0)) AS lite_seats, (pm.customer_id IS NOT NULL) AS has_pm, (fp_org.org_id IS NOT NULL OR mg.org_id IS NOT NULL) AS has_fp, (picked.app_status IN ('active','trialing')) AS is_live, if(mg.org_id IS NOT NULL, coalesce(mg.amt,0), coalesce(org_arr.arr_actual,0)) AS arr_actual, round(coalesce(sx.period_amt,0) * if(sx.intv='month',12,1),2) AS arr_potential FROM (SELECT id FROM postgres.orgs LIMIT 1 BY id) o LEFT JOIN picked ON picked.org_id=o.id LEFT JOIN sx ON sx.sub_id=picked.stripe_subscription_id LEFT JOIN fp_org ON fp_org.org_id=o.id LEFT JOIN org_arr ON org_arr.org_id=o.id LEFT JOIN pm ON pm.customer_id=picked.stripe_customer_id LEFT JOIN managed mg ON mg.org_id=o.id WHERE o.id NOT IN (SELECT org_id FROM excluded_orgs)),
staged AS (SELECT *, multiIf(has_fp AND arr_actual>=0.01,'Paid', is_live AND has_pm,'Intent to Pay', is_live,'Trialing','') AS stage, if(has_fp AND arr_actual>=0.01, arr_actual, arr_potential) AS arr_row FROM per_org)
SELECT stage, count() AS firms, round(sum(arr_row),2) AS arr, round(sum(arr_row)/12,2) AS mrr,
       sum(full_seats+lite_seats) AS seats, sum(full_seats) AS full_seats_total, sum(lite_seats) AS lite_seats_total,
       round(sum(full_seats+lite_seats)/count(),2) AS avg_seats_per_firm
FROM staged WHERE stage!='' GROUP BY stage ORDER BY arr DESC
`

const ALLSTATS_SECTION_3 = `
WITH
sxi AS (SELECT id AS sid, coalesce(nullIf(plan.interval,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(toString(items),''),'data'),1),'plan','interval')) AS intv, arraySum(arrayMap(x -> JSONExtractInt(x,'plan','amount')*JSONExtractInt(x,'quantity'), JSONExtractArrayRaw(coalesce(toString(items),''),'data')))/100.0 AS period_amt, arraySum(arrayMap(d -> if(JSONExtractString(d,'coupon','duration')!='once', toFloatOrZero(JSONExtractRaw(d,'coupon','percent_off')), 0.0), JSONExtractArrayRaw(coalesce(toString(discounts),'[]')))) AS pct_off, arraySum(arrayMap(d -> if(JSONExtractString(d,'coupon','duration')!='once', toFloatOrZero(JSONExtractRaw(d,'coupon','amount_off')), 0.0), JSONExtractArrayRaw(coalesce(toString(discounts),'[]'))))/100.0 AS amt_off_period FROM stripe.subscription LIMIT 1 BY id),
sub_rev AS (SELECT sid, round(greatest(0, (period_amt*if(intv='month',12,1))*(1-least(pct_off,100.0)/100.0) - amt_off_period*if(intv='month',12,1)),2) AS arr FROM sxi),
paid AS (SELECT subscription_id FROM stripe.invoice WHERE status='paid' AND toFloat(total)>0 AND subscription_id IS NOT NULL GROUP BY subscription_id),
org_keys AS (SELECT id AS org_id, toString(id) AS k FROM postgres.orgs LIMIT 1 BY id UNION ALL SELECT id AS org_id, coalesce(nullIf(workos_id,''), external_id) AS k FROM postgres.orgs LIMIT 1 BY id),
sheet_excl AS (SELECT DISTINCT os.org_id FROM postgres.org_subscriptions os JOIN (SELECT id, lower(email) AS email FROM stripe.customer LIMIT 1 BY id) c ON c.id = os.stripe_customer_id WHERE c.email IN (SELECT lower(customer_email) FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) != 'managed' AND coalesce(customer_email,'') != '')),
managed AS (SELECT ok.org_id AS org_id, max(toFloat64OrNull(replaceRegexpAll(coalesce(toString(ovr.amount),''),'[^0-9.]',''))) AS amt, max(toFloat64OrNull(replaceRegexpAll(coalesce(toString(ovr.full_seats),''),'[^0-9.]',''))) AS full_seats, max(toFloat64OrNull(replaceRegexpAll(coalesce(toString(ovr.lite_seats),''),'[^0-9.]',''))) AS lite_seats FROM override_googlesheets_manual_stripe_changes ovr JOIN org_keys ok ON ok.k = toString(ovr.org_id) WHERE lower(ovr.exclude_reason)='managed' AND coalesce(toString(ovr.org_id),'') != '' GROUP BY ok.org_id),
sub_arr AS (SELECT osb.org_id, round(sum(sr.arr),2) AS arr_actual FROM postgres.org_subscriptions osb JOIN sub_rev sr ON sr.sid=osb.stripe_subscription_id WHERE osb.stripe_subscription_id IN (SELECT subscription_id FROM paid) AND osb.org_id NOT IN (SELECT org_id FROM sheet_excl) AND osb.status!='canceled' AND sr.arr>0 GROUP BY osb.org_id),
picked AS (SELECT org_id, stripe_subscription_id, status AS app_status, trial_ends_at, full_seat_count, lite_seat_count FROM postgres.org_subscriptions ORDER BY (stripe_subscription_id IN (SELECT subscription_id FROM paid) AND status!='canceled') DESC, (status='active') DESC, (status='trialing' AND trial_ends_at>now()) DESC, created_at DESC LIMIT 1 BY org_id),
excl_emails AS (SELECT DISTINCT lower(customer_email) AS email FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) IN ('internal','duplicate','fake') AND coalesce(customer_email,'')!=''),
founder_x AS (SELECT uo.org_id AS org_id, argMin(lower(u.email), uo.created_at) AS email FROM postgres.users_orgs uo JOIN postgres.users u ON u.id=uo.user_id GROUP BY uo.org_id),
excluded_orgs AS (SELECT DISTINCT os.org_id AS org_id FROM postgres.org_subscriptions os JOIN (SELECT id, lower(email) AS email FROM stripe.customer LIMIT 1 BY id) c ON c.id=os.stripe_customer_id WHERE c.email IN (SELECT email FROM excl_emails) UNION DISTINCT SELECT org_id FROM founder_x WHERE email IN (SELECT email FROM excl_emails)),
paid_orgs AS (SELECT o.id AS org_id, if(mg.org_id IS NOT NULL, coalesce(mg.amt,0), coalesce(sub_arr.arr_actual,0)) AS arr, (mg.org_id IS NOT NULL) AS is_managed, if(mg.org_id IS NOT NULL AND (coalesce(mg.full_seats,0)+coalesce(mg.lite_seats,0))>0, coalesce(mg.full_seats,0)+coalesce(mg.lite_seats,0), coalesce(picked.full_seat_count,0)+coalesce(picked.lite_seat_count,0)) AS seats FROM (SELECT id FROM postgres.orgs LIMIT 1 BY id) o LEFT JOIN sub_arr ON sub_arr.org_id=o.id LEFT JOIN managed mg ON mg.org_id=o.id LEFT JOIN picked ON picked.org_id=o.id WHERE if(mg.org_id IS NOT NULL, coalesce(mg.amt,0), coalesce(sub_arr.arr_actual,0))>=0.01 AND o.id NOT IN (SELECT org_id FROM excluded_orgs))
SELECT count() AS paid_firms, round(sum(arr),2) AS total_arr,
       round(sum(arr)/count(),2) AS avg_arr_per_firm,
       round(sumIf(arr, is_managed=0)/countIf(is_managed=0),2) AS avg_arr_per_firm_excl_managed,
       round(sum(arr)/sum(seats),2) AS avg_arr_per_seat,
       round(sum(seats)/count(),2) AS avg_seats_per_firm
FROM paid_orgs
`

const ALLSTATS_SECTION_4 = `
WITH
sxi AS (SELECT id AS sid, coalesce(nullIf(plan.interval,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(toString(items),''),'data'),1),'plan','interval')) AS intv, arraySum(arrayMap(x -> JSONExtractInt(x,'plan','amount')*JSONExtractInt(x,'quantity'), JSONExtractArrayRaw(coalesce(toString(items),''),'data')))/100.0 AS period_amt, arraySum(arrayMap(d -> if(JSONExtractString(d,'coupon','duration')!='once', toFloatOrZero(JSONExtractRaw(d,'coupon','percent_off')), 0.0), JSONExtractArrayRaw(coalesce(toString(discounts),'[]')))) AS pct_off, arraySum(arrayMap(d -> if(JSONExtractString(d,'coupon','duration')!='once', toFloatOrZero(JSONExtractRaw(d,'coupon','amount_off')), 0.0), JSONExtractArrayRaw(coalesce(toString(discounts),'[]'))))/100.0 AS amt_off_period FROM stripe.subscription LIMIT 1 BY id),
sub_rev AS (SELECT sid, round(greatest(0, (period_amt*if(intv='month',12,1))*(1-least(pct_off,100.0)/100.0) - amt_off_period*if(intv='month',12,1)),2) AS arr FROM sxi),
paid AS (SELECT subscription_id FROM stripe.invoice WHERE status='paid' AND toFloat(total)>0 AND subscription_id IS NOT NULL GROUP BY subscription_id),
sheet_excl AS (SELECT DISTINCT os.org_id FROM postgres.org_subscriptions os JOIN (SELECT id, lower(email) AS email FROM stripe.customer LIMIT 1 BY id) c ON c.id = os.stripe_customer_id WHERE c.email IN (SELECT lower(customer_email) FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) != 'managed' AND coalesce(customer_email,'') != '')),
firstpay AS (SELECT subscription_id AS sid, min(created_at) AS first_payment FROM stripe.invoice WHERE status='paid' AND toFloat(total)>0 GROUP BY subscription_id),
picked AS (SELECT org_id, stripe_subscription_id, stripe_customer_id, status AS app_status, cancel_at_period_end, current_period_end, full_seat_count, lite_seat_count, trial_ends_at FROM postgres.org_subscriptions ORDER BY (stripe_subscription_id IN (SELECT subscription_id FROM paid) AND status!='canceled') DESC, (status='active') DESC, (status='trialing' AND trial_ends_at>now()) DESC, created_at DESC LIMIT 1 BY org_id),
sx AS (SELECT id AS sub_id, coalesce(nullIf(plan.product,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(items,''),'data'),1),'plan','product')) AS product_id FROM stripe.subscription LIMIT 1 BY id),
prod AS (SELECT id, name FROM stripe.product LIMIT 1 BY id),
cust AS (SELECT id, email FROM stripe.customer LIMIT 1 BY id),
fp_org AS (SELECT org_id FROM postgres.org_subscriptions os JOIN stripe.invoice i ON i.subscription_id=os.stripe_subscription_id WHERE i.status='paid' AND toFloat(i.total)>0 GROUP BY org_id),
orgs AS (SELECT id, name FROM postgres.orgs LIMIT 1 BY id)
SELECT orgs.name AS org_name, cust.email AS customer_email, picked.stripe_subscription_id AS subscription_id,
       (coalesce(picked.full_seat_count,0)+coalesce(picked.lite_seat_count,0)) AS seats,
       round(coalesce(sub_rev.arr,0),2) AS arr, picked.app_status AS status,
       if(picked.app_status='canceled','canceled','scheduled_cancel') AS churn_type,
       formatDateTime(picked.current_period_end,'%Y-%m-%d') AS period_end_date,
       formatDateTime(firstpay.first_payment,'%Y-%m-%d') AS first_payment,
       prod.name AS plan
FROM orgs JOIN picked ON picked.org_id=orgs.id
LEFT JOIN sub_rev ON sub_rev.sid=picked.stripe_subscription_id
LEFT JOIN firstpay ON firstpay.sid=picked.stripe_subscription_id
LEFT JOIN sx ON sx.sub_id=picked.stripe_subscription_id
LEFT JOIN prod ON prod.id=sx.product_id
LEFT JOIN cust ON cust.id=picked.stripe_customer_id
WHERE orgs.id IN (SELECT org_id FROM fp_org)
  AND picked.org_id NOT IN (SELECT org_id FROM sheet_excl)
  AND (picked.app_status='canceled' OR (picked.app_status='active' AND picked.cancel_at_period_end=1))
ORDER BY arr DESC
LIMIT 5000
`

const ALLSTATS_SECTION_5 = `
WITH
paid AS (SELECT subscription_id FROM stripe.invoice WHERE status='paid' AND toFloat(total)>0 AND subscription_id IS NOT NULL GROUP BY subscription_id),
fp_org AS (SELECT org_id FROM postgres.org_subscriptions os JOIN stripe.invoice i ON i.subscription_id=os.stripe_subscription_id WHERE i.status='paid' AND toFloat(i.total)>0 GROUP BY org_id),
picked AS (SELECT org_id, stripe_subscription_id, stripe_customer_id, status AS app_status, created_at AS sub_created, trial_ends_at, full_seat_count, lite_seat_count FROM postgres.org_subscriptions ORDER BY (stripe_subscription_id IN (SELECT subscription_id FROM paid) AND status!='canceled') DESC, (status='active') DESC, (status='trialing' AND trial_ends_at>now()) DESC, created_at DESC LIMIT 1 BY org_id),
sx AS (SELECT id AS sub_id, arraySum(arrayMap(x -> JSONExtractInt(x,'plan','amount')*JSONExtractInt(x,'quantity'), JSONExtractArrayRaw(coalesce(items,''),'data')))/100.0 AS mrr, coalesce(nullIf(plan.interval,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(items,''),'data'),1),'plan','interval')) AS intv FROM stripe.subscription LIMIT 1 BY id),
pm AS (SELECT DISTINCT customer_id FROM stripe.customerpaymentmethod),
cust AS (SELECT id, email, name FROM stripe.customer LIMIT 1 BY id),
orgs AS (SELECT id, name FROM postgres.orgs LIMIT 1 BY id)
SELECT picked.org_id AS org_id, orgs.name AS org_name, cust.name AS customer_name, cust.email AS customer_email,
       dateDiff('day', picked.sub_created, now()) AS days_in_trial,
       (coalesce(picked.full_seat_count,0)+coalesce(picked.lite_seat_count,0)) AS seats,
       round(coalesce(sx.mrr,0),2) AS mrr,
       round(coalesce(sx.mrr,0)*if(sx.intv='month',12,1),2) AS arr_potential,
       picked.stripe_subscription_id AS subscription_id
FROM orgs JOIN picked ON picked.org_id=orgs.id
LEFT JOIN sx ON sx.sub_id=picked.stripe_subscription_id
LEFT JOIN pm ON pm.customer_id=picked.stripe_customer_id
LEFT JOIN cust ON cust.id=picked.stripe_customer_id
WHERE picked.app_status IN ('active','trialing') AND pm.customer_id IS NULL AND picked.org_id NOT IN (SELECT org_id FROM fp_org)
ORDER BY days_in_trial ASC
LIMIT 5000
`

const ALLSTATS_SECTION_6 = `
WITH
sxi AS (SELECT id AS sid, coalesce(nullIf(plan.interval,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(toString(items),''),'data'),1),'plan','interval')) AS intv, arraySum(arrayMap(x -> JSONExtractInt(x,'plan','amount')*JSONExtractInt(x,'quantity'), JSONExtractArrayRaw(coalesce(toString(items),''),'data')))/100.0 AS period_amt, arraySum(arrayMap(d -> if(JSONExtractString(d,'coupon','duration')!='once', toFloatOrZero(JSONExtractRaw(d,'coupon','percent_off')), 0.0), JSONExtractArrayRaw(coalesce(toString(discounts),'[]')))) AS pct_off, arraySum(arrayMap(d -> if(JSONExtractString(d,'coupon','duration')!='once', toFloatOrZero(JSONExtractRaw(d,'coupon','amount_off')), 0.0), JSONExtractArrayRaw(coalesce(toString(discounts),'[]'))))/100.0 AS amt_off_period FROM stripe.subscription LIMIT 1 BY id),
sub_rev AS (SELECT sid, round(greatest(0, (period_amt*if(intv='month',12,1))*(1-least(pct_off,100.0)/100.0) - amt_off_period*if(intv='month',12,1)),2) AS arr FROM sxi),
paidsub AS (SELECT subscription_id FROM stripe.invoice WHERE status='paid' AND toFloat(total)>0 AND subscription_id IS NOT NULL GROUP BY subscription_id),
sheet_excl AS (SELECT DISTINCT os.org_id FROM postgres.org_subscriptions os JOIN (SELECT id, lower(email) AS email FROM stripe.customer LIMIT 1 BY id) c ON c.id = os.stripe_customer_id WHERE c.email IN (SELECT lower(customer_email) FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) != 'managed' AND coalesce(customer_email,'') != '')),
org_arr AS (SELECT osb.org_id FROM postgres.org_subscriptions osb JOIN sub_rev sr ON sr.sid=osb.stripe_subscription_id WHERE osb.stripe_subscription_id IN (SELECT subscription_id FROM paidsub) AND osb.org_id NOT IN (SELECT org_id FROM sheet_excl) AND osb.status!='canceled' AND sr.arr>0 GROUP BY osb.org_id),
founder AS (SELECT uo.org_id AS org_id, argMin(u.email, uo.created_at) AS email FROM postgres.users_orgs uo JOIN postgres.users u ON u.id=uo.user_id GROUP BY uo.org_id),
noninternal AS (SELECT org_id FROM founder WHERE email IS NOT NULL AND email!='' AND email NOT ILIKE '%@pingassistant.com%' AND email NOT ILIKE '%@pingassistant.dev%' AND email NOT ILIKE '%@posthog.com%' AND email NOT IN ('gptfam69@gmail.com','chad@bookends.app','stormmgarnett@gmail.com')),
promo AS (SELECT DISTINCT org_id FROM postgres.promo_redemptions)
SELECT (SELECT count() FROM noninternal) AS total_signups,
       (SELECT count() FROM noninternal WHERE org_id IN (SELECT org_id FROM org_arr)) AS converted,
       round((SELECT count() FROM noninternal WHERE org_id IN (SELECT org_id FROM org_arr)) / (SELECT count() FROM noninternal) * 100, 1) AS signup_to_paid_pct,
       (SELECT count() FROM noninternal WHERE org_id IN (SELECT org_id FROM promo)) AS promo_signups,
       (SELECT count() FROM noninternal WHERE org_id IN (SELECT org_id FROM promo) AND org_id IN (SELECT org_id FROM org_arr)) AS promo_converted,
       round((SELECT count() FROM noninternal WHERE org_id IN (SELECT org_id FROM promo) AND org_id IN (SELECT org_id FROM org_arr)) / (SELECT count() FROM noninternal WHERE org_id IN (SELECT org_id FROM promo)) * 100, 1) AS promo_to_paid_pct
`

const ALLSTATS_SECTION_7 = `
WITH s AS (SELECT org_id, snapshot_date, max(toFloat(total_arr)) AS arr FROM override_googlesheets_arr_snapshot GROUP BY org_id, snapshot_date),
base AS (SELECT org_id, arr AS base_arr FROM s WHERE snapshot_date='2026-04-01' AND arr>0),
curr AS (SELECT org_id, arr AS curr_arr FROM s WHERE snapshot_date='2026-05-01')
SELECT round(sum(base_arr),2) AS base_total,
       round(sum(coalesce(curr.curr_arr,0)),2) AS retained_plus_expansion,
       round(sum(coalesce(curr.curr_arr,0))/sum(base_arr)*100,1) AS nrr_pct,
       round(sum(least(coalesce(curr.curr_arr,0), base_arr))/sum(base_arr)*100,1) AS grr_pct
FROM base LEFT JOIN curr ON curr.org_id=base.org_id
`

const ALLSTATS_SECTION_8 = `
WITH d AS (SELECT month, org_id, metric, max(toFloat(amount)) AS amount FROM override_googlesheets_arr_waterfall_facts GROUP BY month, org_id, metric)
SELECT month,
       round(sumIf(amount,metric='SOM'),2)       AS som,
       round(sumIf(amount,metric='New'),2)       AS new,
       round(sumIf(amount,metric='Upgrade'),2)   AS upgrade,
       round(sumIf(amount,metric='Downgrade'),2) AS downgrade,
       round(sumIf(amount,metric='Churn'),2)     AS churn,
       round(sumIf(amount,metric='EOM'),2)       AS eom
FROM d GROUP BY month ORDER BY month
LIMIT 5000
`

const ALLSTATS_SECTIONS = [
  { title: 'ARR / MRR Summary by Stage', hogql: ALLSTATS_SECTION_1 },
  { title: 'Avg Revenue Metrics (Paid)', hogql: ALLSTATS_SECTION_3 },
  { title: 'Churned / Scheduled-Cancel Audit', hogql: ALLSTATS_SECTION_4 },
  { title: 'Trialing (No Payment Method)', hogql: ALLSTATS_SECTION_5 },
  { title: 'Conversion (signup to paid, promo to paid)', hogql: ALLSTATS_SECTION_6 },
  { title: 'NRR / GRR', hogql: ALLSTATS_SECTION_7 },
  { title: 'Net New ARR by Month / Waterfall', hogql: ALLSTATS_SECTION_8 }
]

/**
 * Build the "All the Stats" sheet from the section queries (labeled blocks).
 */
function render_all_stats_view() {
  const t0 = new Date()
  const ss = SpreadsheetApp.getActive()

  const props = PropertiesService.getScriptProperties()
  const apiKey = props.getProperty('POSTHOG_API_KEY')
  if (!apiKey) throw new Error('Missing POSTHOG_API_KEY in Script Properties')
  const projectId = props.getProperty('POSTHOG_PROJECT_ID') || POSTHOG_RAW_CFG.PROJECT_ID_FALLBACK

  const sh = ss.getSheetByName(ALL_STATS_CFG.OUT_SHEET) || ss.insertSheet(ALL_STATS_CFG.OUT_SHEET)
  sh.clear()

  let row = 1
  let totalRows = 0
  let maxCols = 1

  for (const section of ALLSTATS_SECTIONS) {
    const { columns, results } = sauronQueryRun_(apiKey, projectId, section.hogql, 'all_stats:' + section.title)
    const headers = (columns && columns.length) ? columns : ['(no columns)']
    maxCols = Math.max(maxCols, headers.length)

    // Section title
    sh.getRange(row, 1, 1, 1).setValues([[section.title]]).setFontWeight('bold').setFontSize(12)
    row++
    // Column headers
    sh.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#F3F4F6')
    row++
    // Data rows
    const rows = (results || []).map(r => {
      const out = new Array(headers.length)
      for (let i = 0; i < headers.length; i++) out[i] = (r && r[i] != null) ? r[i] : ''
      return out
    })
    if (rows.length) {
      const chunk = 5000
      for (let i = 0; i < rows.length; i += chunk) {
        const part = rows.slice(i, i + chunk)
        sh.getRange(row + i, 1, part.length, headers.length).setValues(part)
      }
      row += rows.length
      totalRows += rows.length
    }
    // Gap between sections
    row += 2
  }

  try { sh.autoResizeColumns(1, maxCols) } catch (e) {}

  if (typeof writeSyncLog === 'function') {
    writeSyncLog('render_all_stats_view', 'ok', '', totalRows, (new Date() - t0) / 1000, '')
  }
  Logger.log(`render_all_stats_view: ${ALLSTATS_SECTIONS.length} sections, ${totalRows} data rows in ${((new Date() - t0) / 1000).toFixed(1)}s`)
  return { rows_in: ALLSTATS_SECTIONS.length, rows_out: totalRows }
}
