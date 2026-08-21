/**************************************************************
 * Ring Queries (PRODUCTION)
 *
 * The two HogQL queries that build the live "The Ring" sheet, used by
 * render_ring_view() in "Render The Ring.js". Kept EXACTLY in sync with
 * the HubSpot Ring logic (dashboard tile 0wTWWPar):
 *   (A) RING SUMMARY  -> 3 rows: Paid / Card Info Entered / Trialing
 *   (B) RING DETAIL   -> one row per org (all three statuses)
 *
 * Conventions:
 *   - Status from app DB (postgres.org_subscriptions.status).
 *   - Paid ARR = run-rate from latest full-cycle invoice, net of discounts,
 *     annualized (monthly x12, annual as-is).
 *   - MANAGED orgs (override sheet exclude_reason='managed') are valued at the
 *     override sheet's `amount` column (NOT run-rate). Applied inside org_arr.
 *   - Card Info Entered / Trialing ARR = POTENTIAL (plan list price annualized).
 **************************************************************/

// (A) RING SUMMARY — 3 rows: Paid / Card Info Entered / Trialing
const RING_SUMMARY_HOGQL = `
WITH
-- ARR basis = CURRENT subscription items (contracted run-rate), annualized, net of active
-- discount. Replaces the old trailing "latest full-cycle invoice" method that lagged
-- mid-cycle upgrades (annual seat adds wouldn't show until next renewal). 'once' coupons ignored.
sxi AS (SELECT id AS sid, coalesce(nullIf(plan.interval,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(toString(items),''),'data'),1),'plan','interval')) AS intv, arraySum(arrayMap(x -> JSONExtractInt(x,'plan','amount')*JSONExtractInt(x,'quantity'), JSONExtractArrayRaw(coalesce(toString(items),''),'data')))/100.0 AS period_amt, arraySum(arrayMap(d -> if(JSONExtractString(d,'coupon','duration')!='once', toFloatOrZero(JSONExtractRaw(d,'coupon','percent_off')), 0.0), JSONExtractArrayRaw(coalesce(toString(discounts),'[]')))) AS pct_off, arraySum(arrayMap(d -> if(JSONExtractString(d,'coupon','duration')!='once', toFloatOrZero(JSONExtractRaw(d,'coupon','amount_off')), 0.0), JSONExtractArrayRaw(coalesce(toString(discounts),'[]'))))/100.0 AS amt_off_period FROM stripe.subscription LIMIT 1 BY id),
ci AS (SELECT sid, round(greatest(0, (period_amt*if(intv='month',12,1))*(1-least(pct_off,100.0)/100.0) - amt_off_period*if(intv='month',12,1)),2) AS arr FROM sxi),
paid AS (SELECT subscription_id FROM stripe.invoice WHERE status='paid' AND toFloat(total)>0 AND subscription_id IS NOT NULL GROUP BY subscription_id),
sheet_excl AS (SELECT subscription_id FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) != 'managed'),
managed AS (SELECT subscription_id AS sid, toFloat(nullIf(replaceAll(replaceAll(toString(amount),'$',''),',',''),'')) AS amt FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason)='managed'),
fp_org AS (SELECT org_id FROM postgres.org_subscriptions os JOIN stripe.invoice i ON i.subscription_id=os.stripe_subscription_id WHERE i.status='paid' AND toFloat(i.total)>0 GROUP BY org_id),
org_arr AS (SELECT osb.org_id, round(sum(if(m.sid IS NOT NULL, m.amt, ci.arr)),2) AS arr_actual FROM postgres.org_subscriptions osb JOIN ci ON ci.sid=osb.stripe_subscription_id LEFT JOIN managed m ON m.sid=osb.stripe_subscription_id WHERE osb.stripe_subscription_id IN (SELECT subscription_id FROM paid) AND osb.stripe_subscription_id NOT IN (SELECT subscription_id FROM sheet_excl) AND osb.status!='canceled' AND if(m.sid IS NOT NULL, m.amt, ci.arr)>0 GROUP BY osb.org_id),
picked AS (SELECT org_id, stripe_subscription_id, stripe_customer_id, status AS app_status, trial_ends_at, full_seat_count, lite_seat_count FROM postgres.org_subscriptions ORDER BY (stripe_subscription_id IN (SELECT subscription_id FROM paid) AND status!='canceled') DESC, (status='active') DESC, (status='trialing' AND trial_ends_at>now()) DESC, created_at DESC LIMIT 1 BY org_id),
sx AS (SELECT id AS sub_id, coalesce(nullIf(plan.interval,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(items,''),'data'),1),'plan','interval')) AS intv, arraySum(arrayMap(x -> JSONExtractInt(x,'plan','amount')*JSONExtractInt(x,'quantity'), JSONExtractArrayRaw(coalesce(items,''),'data')))/100.0 AS period_amt FROM stripe.subscription LIMIT 1 BY id),
pm AS (SELECT DISTINCT customer_id FROM stripe.customerpaymentmethod),
-- Internal/test orgs hidden from The Ring: keyed on the override sheet's
-- customer_email column (flag a row exclude_reason='internal' with the org's
-- Stripe customer email). Auto-applies as you flag more in the sheet.
internal_orgs AS (SELECT DISTINCT os.org_id FROM postgres.org_subscriptions os JOIN (SELECT id, lower(email) AS email FROM stripe.customer LIMIT 1 BY id) c ON c.id = os.stripe_customer_id WHERE c.email IN (SELECT lower(customer_email) FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) IN ('internal','duplicate','fake') AND coalesce(customer_email,'') != '') UNION DISTINCT SELECT org_id FROM (SELECT uo.org_id AS org_id, argMin(lower(u.email), uo.created_at) AS femail FROM postgres.users_orgs uo JOIN postgres.users u ON u.id=uo.user_id GROUP BY uo.org_id) WHERE femail IN (SELECT lower(customer_email) FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) IN ('internal','duplicate','fake') AND coalesce(customer_email,'') != '')),
per_org AS (SELECT o.id AS org_id, coalesce(picked.full_seat_count,0)+coalesce(picked.lite_seat_count,0) AS seats, (pm.customer_id IS NOT NULL) AS has_pm, (fp_org.org_id IS NOT NULL) AS has_fp, (picked.app_status IN ('active','trialing')) AS is_live, coalesce(org_arr.arr_actual,0) AS arr_actual, round(coalesce(sx.period_amt,0)*if(sx.intv='month',12,1),2) AS arr_potential FROM (SELECT id FROM postgres.orgs LIMIT 1 BY id) o LEFT JOIN picked ON picked.org_id=o.id LEFT JOIN sx ON sx.sub_id=picked.stripe_subscription_id LEFT JOIN fp_org ON fp_org.org_id=o.id LEFT JOIN org_arr ON org_arr.org_id=o.id LEFT JOIN pm ON pm.customer_id=picked.stripe_customer_id WHERE o.id NOT IN (SELECT org_id FROM internal_orgs)),
staged AS (SELECT *, multiIf(has_fp AND arr_actual>=0.01,'Paid', is_live AND has_pm,'Card Info Entered', is_live,'Trialing','') AS status, if(has_fp AND arr_actual>=0.01, arr_actual, arr_potential) AS arr_row FROM per_org)
SELECT status, round(sum(arr_row),2) AS arr, count() AS subscriptions, sum(seats) AS total_seats
FROM staged WHERE status!='' GROUP BY status
ORDER BY multiIf(status='Paid',1, status='Card Info Entered',2, 3)
`

// (B) RING DETAIL — one row per org; all Paid + Card Info Entered + Trialing
const RING_DETAIL_HOGQL = `
WITH
-- ARR basis = CURRENT subscription items (contracted run-rate), annualized, net of active
-- discount. Replaces the old trailing "latest full-cycle invoice" method that lagged
-- mid-cycle upgrades (annual seat adds wouldn't show until next renewal). 'once' coupons ignored.
sxi AS (SELECT id AS sid, coalesce(nullIf(plan.interval,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(toString(items),''),'data'),1),'plan','interval')) AS intv, arraySum(arrayMap(x -> JSONExtractInt(x,'plan','amount')*JSONExtractInt(x,'quantity'), JSONExtractArrayRaw(coalesce(toString(items),''),'data')))/100.0 AS period_amt, arraySum(arrayMap(d -> if(JSONExtractString(d,'coupon','duration')!='once', toFloatOrZero(JSONExtractRaw(d,'coupon','percent_off')), 0.0), JSONExtractArrayRaw(coalesce(toString(discounts),'[]')))) AS pct_off, arraySum(arrayMap(d -> if(JSONExtractString(d,'coupon','duration')!='once', toFloatOrZero(JSONExtractRaw(d,'coupon','amount_off')), 0.0), JSONExtractArrayRaw(coalesce(toString(discounts),'[]'))))/100.0 AS amt_off_period FROM stripe.subscription LIMIT 1 BY id),
ci AS (SELECT sid, round(greatest(0, (period_amt*if(intv='month',12,1))*(1-least(pct_off,100.0)/100.0) - amt_off_period*if(intv='month',12,1)),2) AS arr FROM sxi),
paid AS (SELECT subscription_id FROM stripe.invoice WHERE status='paid' AND toFloat(total)>0 AND subscription_id IS NOT NULL GROUP BY subscription_id),
sheet_excl AS (SELECT subscription_id FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) != 'managed'),
managed AS (SELECT subscription_id AS sid, toFloat(nullIf(replaceAll(replaceAll(toString(amount),'$',''),',',''),'')) AS amt FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason)='managed'),
fp_org AS (SELECT org_id, min(i.created_at) AS first_payment FROM postgres.org_subscriptions os JOIN stripe.invoice i ON i.subscription_id=os.stripe_subscription_id WHERE i.status='paid' AND toFloat(i.total)>0 GROUP BY org_id),
org_arr AS (SELECT osb.org_id, round(sum(if(m.sid IS NOT NULL, m.amt, ci.arr)),2) AS arr_actual FROM postgres.org_subscriptions osb JOIN ci ON ci.sid=osb.stripe_subscription_id LEFT JOIN managed m ON m.sid=osb.stripe_subscription_id WHERE osb.stripe_subscription_id IN (SELECT subscription_id FROM paid) AND osb.stripe_subscription_id NOT IN (SELECT subscription_id FROM sheet_excl) AND osb.status!='canceled' AND if(m.sid IS NOT NULL, m.amt, ci.arr)>0 GROUP BY osb.org_id),
picked AS (SELECT org_id, stripe_subscription_id, stripe_customer_id, status AS app_status, trial_ends_at, full_seat_count, lite_seat_count FROM postgres.org_subscriptions ORDER BY (stripe_subscription_id IN (SELECT subscription_id FROM paid) AND status!='canceled') DESC, (status='active') DESC, (status='trialing' AND trial_ends_at>now()) DESC, created_at DESC LIMIT 1 BY org_id),
sx AS (SELECT id AS sub_id, coalesce(nullIf(plan.interval,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(items,''),'data'),1),'plan','interval')) AS intv, arraySum(arrayMap(x -> JSONExtractInt(x,'plan','amount')*JSONExtractInt(x,'quantity'), JSONExtractArrayRaw(coalesce(items,''),'data')))/100.0 AS period_amt FROM stripe.subscription LIMIT 1 BY id),
pm AS (SELECT DISTINCT customer_id FROM stripe.customerpaymentmethod),
cust AS (SELECT id, email, name FROM stripe.customer LIMIT 1 BY id),
orgs AS (SELECT id, name, created_at FROM postgres.orgs LIMIT 1 BY id),
-- Internal/test orgs hidden from The Ring: keyed on the override sheet's
-- customer_email column (flag a row exclude_reason='internal' with the org's
-- Stripe customer email). Auto-applies as you flag more in the sheet.
internal_orgs AS (SELECT DISTINCT os.org_id FROM postgres.org_subscriptions os JOIN (SELECT id, lower(email) AS email FROM stripe.customer LIMIT 1 BY id) c ON c.id = os.stripe_customer_id WHERE c.email IN (SELECT lower(customer_email) FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) IN ('internal','duplicate','fake') AND coalesce(customer_email,'') != '') UNION DISTINCT SELECT org_id FROM (SELECT uo.org_id AS org_id, argMin(lower(u.email), uo.created_at) AS femail FROM postgres.users_orgs uo JOIN postgres.users u ON u.id=uo.user_id GROUP BY uo.org_id) WHERE femail IN (SELECT lower(customer_email) FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) IN ('internal','duplicate','fake') AND coalesce(customer_email,'') != '')),
per_org AS (SELECT o.id AS org_id, o.name AS org_name, o.created_at AS sign_up_date, cust.email AS customer_email, cust.name AS customer_name, fp_org.first_payment AS first_payment_at, picked.trial_ends_at AS trial_ends_at, sx.intv AS interval, coalesce(picked.full_seat_count,0)+coalesce(picked.lite_seat_count,0) AS seats, (pm.customer_id IS NOT NULL) AS has_pm, (fp_org.org_id IS NOT NULL) AS has_fp, (picked.app_status IN ('active','trialing')) AS is_live, coalesce(org_arr.arr_actual,0) AS arr_actual, round(coalesce(sx.period_amt,0)*if(sx.intv='month',12,1),2) AS arr_potential FROM orgs o LEFT JOIN picked ON picked.org_id=o.id LEFT JOIN sx ON sx.sub_id=picked.stripe_subscription_id LEFT JOIN fp_org ON fp_org.org_id=o.id LEFT JOIN org_arr ON org_arr.org_id=o.id LEFT JOIN pm ON pm.customer_id=picked.stripe_customer_id LEFT JOIN cust ON cust.id=picked.stripe_customer_id WHERE o.id NOT IN (SELECT org_id FROM internal_orgs)),
staged AS (SELECT *, multiIf(has_fp AND arr_actual>=0.01,'Paid', is_live AND has_pm,'Card Info Entered', is_live,'Trialing','') AS status FROM per_org)
SELECT customer_email, customer_name, org_name, status,
       formatDateTime(sign_up_date,'%Y-%m-%d') AS sign_up_date,
       if(first_payment_at IS NOT NULL, formatDateTime(first_payment_at,'%Y-%m-%d'), '') AS first_payment_at,
       if(status!='Paid' AND trial_ends_at>now(), dateDiff('day', now(), trial_ends_at), NULL) AS trial_days_remaining,
       interval,
       if(status='Paid', arr_actual, arr_potential) AS arr,
       seats
FROM staged WHERE status!=''
ORDER BY multiIf(status='Paid',1, status='Card Info Entered',2, 3), arr DESC
LIMIT 5000
`
