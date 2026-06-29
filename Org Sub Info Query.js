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
  OUT_SHEET: 'org_subscription_info'
}

// Query kept verbatim (the `interval` identifier's backticks are escaped for
// the JS template literal). Full org_subscription_info equivalent.
const ORG_SUB_INFO_QUERY_HOGQL = `
WITH
os AS (   -- one sub per org: active → real trial → latest
  SELECT org_id, id AS latest_subscription_id, stripe_subscription_id, status AS app_status,
         created_at AS sub_created, trial_ends_at, current_period_end, cancel_at_period_end, stripe_customer_id
  FROM postgres.org_subscriptions
  ORDER BY (status='active') DESC, (status='trialing' AND trial_ends_at > now()) DESC, created_at DESC
  LIMIT 1 BY org_id
),
sx AS (   -- stripe sub detail (amounts handle multi-item via items[])
  SELECT id AS sub_id, customer_id, status AS stripe_status, canceled_at,
    JSONExtractString(coalesce(cancellation_details,''),'reason') AS churn_reason,
    trial_end, current_period_end AS cpe, cancel_at_period_end AS cape,
    coalesce(nullIf(plan.product,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(items,''),'data'),1),'plan','product')) AS product_id,
    coalesce(nullIf(plan.interval,''), JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(items,''),'data'),1),'plan','interval')) AS intv,
    coalesce(toInt(nullIf(toString(plan.interval_count),'')), JSONExtractInt(arrayElement(JSONExtractArrayRaw(coalesce(items,''),'data'),1),'plan','interval_count')) AS intv_count,
    arraySum(arrayMap(x -> JSONExtractInt(x,'quantity'), JSONExtractArrayRaw(coalesce(items,''),'data'))) AS quantity_total,
    arraySum(arrayMap(x -> JSONExtractInt(x,'plan','amount') * JSONExtractInt(x,'quantity'), JSONExtractArrayRaw(coalesce(items,''),'data'))) AS amount_cents,
    JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(discounts,'[]')),1),'coupon','name') AS coupon_name
  FROM stripe.subscription LIMIT 1 BY id
),
fp AS (SELECT subscription_id, min(created_at) AS first_payment_at FROM stripe.invoice WHERE status='paid' AND toFloat(total)>0 AND subscription_id IS NOT NULL GROUP BY subscription_id),
pm AS (SELECT customer_id, min(created_at) AS pm_created FROM stripe.customerpaymentmethod GROUP BY customer_id),
promo AS (SELECT pr.org_id AS org_id, count() AS n_promo, argMax(pc.code, pr.redeemed_at) AS promo_code, argMax(pc.name, pr.redeemed_at) AS promo_name, argMax(pc.trial_days, pr.redeemed_at) AS promo_trial_days, max(pr.redeemed_at) AS redeemed_at FROM postgres.promo_redemptions pr JOIN postgres.promo_codes pc ON pc.id=pr.promo_code_id GROUP BY pr.org_id),
prod AS (SELECT id, name FROM stripe.product LIMIT 1 BY id),
cust AS (SELECT id, email FROM stripe.customer LIMIT 1 BY id),
orgs AS (SELECT id, name, created_at, workos_id, external_id FROM postgres.orgs LIMIT 1 BY id)
SELECT
  os.org_id                                            AS org_id,        -- internal DB id
  coalesce(nullIf(orgs.workos_id,''), orgs.external_id) AS app_org_id,    -- WorkOS external id (fallback external_id)
  orgs.name                                            AS org_name,
  orgs.created_at                                      AS org_created_at,
  os.latest_subscription_id                            AS latest_subscription_id,
  os.stripe_subscription_id                            AS stripe_subscription_id,
  coalesce(nullIf(sx.stripe_status,''), os.app_status) AS status,
  toDate(os.sub_created)                               AS subscription_created_at_date,
  fp.first_payment_at                                  AS first_payment_at,
  sx.canceled_at                                       AS churn_date,
  prod.name                                            AS plan_name,
  os.stripe_customer_id                                AS stripe_customer_id,
  cust.email                                           AS customer_email,
  if(pm.customer_id IS NOT NULL,'yes','no')            AS has_payment_method,
  pm.pm_created                                        AS payment_method_created_at,
  sx.intv                                              AS \`interval\`,
  sx.intv_count                                        AS interval_count,
  sx.quantity_total                                    AS quantity_total,
  round(sx.amount_cents/nullIf(sx.quantity_total,0)/100.0,2)                                                              AS unit_price,
  round(sx.amount_cents/100.0,2)                                                                                          AS amount,
  round(multiIf(sx.intv='month', sx.amount_cents/nullIf(sx.quantity_total,0), sx.intv='year', sx.amount_cents/nullIf(sx.quantity_total,0)/12, sx.amount_cents/nullIf(sx.quantity_total,0))/100.0,2) AS unit_price_monthly,
  round(multiIf(sx.intv='month', sx.amount_cents/nullIf(sx.quantity_total,0)*12, sx.intv='year', sx.amount_cents/nullIf(sx.quantity_total,0), sx.amount_cents/nullIf(sx.quantity_total,0))/100.0,2) AS unit_price_yearly,
  round(multiIf(sx.intv='month', sx.amount_cents, sx.intv='year', sx.amount_cents/12, sx.amount_cents)/100.0,2)          AS amount_monthly,
  round(multiIf(sx.intv='month', sx.amount_cents*12, sx.intv='year', sx.amount_cents, sx.amount_cents)/100.0,2)          AS amount_yearly,
  coalesce(sx.trial_end, os.trial_ends_at)             AS trial_ends_at,
  coalesce(sx.cpe, os.current_period_end)              AS current_period_end,
  greatest(0, dateDiff('day', now(), coalesce(sx.trial_end, os.trial_ends_at))) AS trial_days_remaining,
  if(coalesce(promo.n_promo,0)>0,'yes','no')           AS promo_used,
  coalesce(promo.promo_code,'')                        AS last_promo_used,
  ''                                                   AS redemption_location,
  promo.redeemed_at                                    AS redeemed_at,
  coalesce(nullIf(sx.coupon_name,''), promo.promo_code, '') AS promo_code,
  coalesce(promo.promo_name,'')                        AS promo_name,
  promo.promo_trial_days                               AS trial_days,
  coalesce(sx.cape, os.cancel_at_period_end)           AS cancel_at_period_end,
  sx.churn_reason                                      AS churn_reason
FROM os
LEFT JOIN sx    ON sx.sub_id = os.stripe_subscription_id
LEFT JOIN orgs  ON orgs.id = os.org_id
LEFT JOIN fp    ON fp.subscription_id = os.stripe_subscription_id
LEFT JOIN cust  ON cust.id = os.stripe_customer_id
LEFT JOIN pm    ON pm.customer_id = os.stripe_customer_id
LEFT JOIN promo ON promo.org_id = os.org_id
LEFT JOIN prod  ON prod.id = sx.product_id
ORDER BY amount DESC
LIMIT 5000
`

/**
 * Build the LIVE org_subscription_info sheet from the PostHog query.
 * (Pipeline part1 calls this; the old combine builder is preserved as
 * render_org_subscription_info_legacy_ in "Render Org Subscription Info.js".)
 * Returns { rows_in, rows_out } for pipeline-style logging.
 */
function render_org_subscription_info() {
  const t0 = new Date()
  const ss = SpreadsheetApp.getActive()

  const props = PropertiesService.getScriptProperties()
  const apiKey = props.getProperty('POSTHOG_API_KEY')
  if (!apiKey) throw new Error('Missing POSTHOG_API_KEY in Script Properties')
  const projectId = props.getProperty('POSTHOG_PROJECT_ID') || POSTHOG_RAW_CFG.PROJECT_ID_FALLBACK

  // Reuse the columns-aware HogQL runner (defined in Sauron Query.js).
  const { columns, results } = sauronQueryRun_(apiKey, projectId, ORG_SUB_INFO_QUERY_HOGQL, 'render_org_subscription_info')

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

  if (typeof writeSyncLog === 'function') {
    writeSyncLog('render_org_subscription_info', 'ok', rows.length, rows.length, (new Date() - t0) / 1000, '')
  }
  Logger.log(`render_org_subscription_info: ${rows.length} rows in ${((new Date() - t0) / 1000).toFixed(1)}s`)
  return { rows_in: rows.length, rows_out: rows.length }
}
