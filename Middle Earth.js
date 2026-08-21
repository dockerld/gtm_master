/**************************************************************
 * Middle Earth
 *
 * Builds the "Middle Earth" sheet from a single PostHog HogQL query
 * (org identity, owner, paying/seats, promo, billing, health score).
 * This replaces the old combine-based render_middle_earth_view.
 *
 * Run via menu: Ping Ops → "Render Middle Earth"
 *   or run render_middle_earth() directly in the editor.
 *
 * org_id  = internal DB id (orgs.id)
 * app_org_id = WorkOS external id (coalesce(workos_id, external_id))
 *
 * Script Properties used:
 *  - POSTHOG_API_KEY
 *  - POSTHOG_PROJECT_ID (optional; falls back to POSTHOG_RAW_CFG)
 * Uses sauronQueryRun_ (defined in "Sauron Query.js").
 **************************************************************/

const MIDDLE_EARTH_CFG = {
  OUT_SHEET: 'Middle Earth'
}

const MIDDLE_EARTH_HOGQL = `
WITH
picked AS (  -- one sub per org: active → real trial → latest
  SELECT org_id, stripe_subscription_id, stripe_customer_id, owner_user_id,
         coalesce(full_seat_count,0)+coalesce(lite_seat_count,0) AS seats, trial_ends_at
  FROM postgres.org_subscriptions
  ORDER BY (status='active') DESC, (status='trialing' AND trial_ends_at > now()) DESC, created_at DESC
  LIMIT 1 BY org_id
),
sub_agg AS (SELECT org_id, arrayStringConcat(groupUniqArray(stripe_subscription_id), ', ') AS stripe_subscription_ids FROM postgres.org_subscriptions GROUP BY org_id),
paid_subs AS (SELECT subscription_id FROM stripe.invoice WHERE status='paid' AND toFloat(total)>0 AND subscription_id IS NOT NULL GROUP BY subscription_id),
-- Map an override-sheet org_id to the internal org id (internal DB id OR app/WorkOS id).
org_keys AS (SELECT id AS org_id, toString(id) AS k FROM postgres.orgs LIMIT 1 BY id UNION ALL SELECT id AS org_id, coalesce(nullIf(workos_id,''), external_id) AS k FROM postgres.orgs LIMIT 1 BY id),
-- MANAGED orgs: keyed on sheet org_id. Marks them paying; seats = full_seats+lite_seats.
managed AS (SELECT ok.org_id AS org_id, max(toFloat64OrNull(replaceRegexpAll(coalesce(toString(ovr.amount),''),'[^0-9.]',''))) AS amt, max(toIntOrZero(replaceRegexpAll(coalesce(toString(ovr.full_seats),''),'[^0-9]',''))) AS full_seats, max(toIntOrZero(replaceRegexpAll(coalesce(toString(ovr.lite_seats),''),'[^0-9]',''))) AS lite_seats FROM override_googlesheets_manual_stripe_changes ovr JOIN org_keys ok ON ok.k = toString(ovr.org_id) WHERE lower(ovr.exclude_reason)='managed' AND coalesce(toString(ovr.org_id),'') != '' GROUP BY ok.org_id),
-- Non-managed exclusions (not counted as paying), keyed on the sheet's customer_email.
sheet_excl AS (SELECT DISTINCT os.org_id FROM postgres.org_subscriptions os JOIN (SELECT id, lower(email) AS email FROM stripe.customer LIMIT 1 BY id) c ON c.id = os.stripe_customer_id WHERE c.email IN (SELECT lower(customer_email) FROM override_googlesheets_manual_stripe_changes WHERE lower(exclude_reason) != 'managed' AND coalesce(customer_email,'') != '')),
ssub AS (SELECT id, canceled_at, status FROM stripe.subscription LIMIT 1 BY id),
paying_sub AS (SELECT osb.org_id AS org_id, coalesce(osb.full_seat_count,0)+coalesce(osb.lite_seat_count,0) AS seats
  FROM postgres.org_subscriptions osb JOIN ssub ON ssub.id=osb.stripe_subscription_id
  WHERE osb.stripe_subscription_id IN (SELECT subscription_id FROM paid_subs)
    AND osb.org_id NOT IN (SELECT org_id FROM sheet_excl)
    AND NOT (ssub.status='canceled' OR (ssub.canceled_at IS NOT NULL AND ssub.canceled_at < now()))),
seats_paying AS (SELECT org_id, sum(seats) AS stripe_seats_paying_sum, count()>0 AS is_paying FROM paying_sub GROUP BY org_id),
promo AS (SELECT pr.org_id AS org_id, argMax(pc.code, pr.redeemed_at) AS promo_code,
          arrayStringConcat(groupUniqArray(pc.code), ', ') AS app_promo_codes
          FROM postgres.promo_redemptions pr JOIN postgres.promo_codes pc ON pc.id=pr.promo_code_id GROUP BY pr.org_id),
u AS (SELECT id, email, name FROM postgres.users LIMIT 1 BY id),
cust AS (SELECT id, email FROM stripe.customer LIMIT 1 BY id),
hc AS (SELECT workos_org_id, max(health_score) AS health_score FROM hubspot.companies GROUP BY workos_org_id),
stripe_promo AS (SELECT id AS sid,
  JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(discounts,'[]')),1),'coupon','name')       AS coupon_name,
  JSONExtractFloat (arrayElement(JSONExtractArrayRaw(coalesce(discounts,'[]')),1),'coupon','percent_off') AS percent_off,
  JSONExtractFloat (arrayElement(JSONExtractArrayRaw(coalesce(discounts,'[]')),1),'coupon','amount_off')  AS amount_off,
  JSONExtractString(arrayElement(JSONExtractArrayRaw(coalesce(discounts,'[]')),1),'promotion_code')       AS promotion_code_id
  FROM stripe.subscription LIMIT 1 BY id),
promo_names AS (
  SELECT 'promo_1S6YXl8UpQbjZAlZ5St46VX1' AS promotion_code_id, 'TIMALYN-BOWENS' AS promo_code_label
  -- UNION ALL SELECT 'promo_xxxxx', 'OTHER-CODE'
)
SELECT
  o.id                                            AS org_id,        -- internal DB id
  coalesce(nullIf(o.workos_id,''), o.external_id) AS app_org_id,    -- WorkOS external id (fallback external_id)
  o.name                                          AS org_name,
  ''                                              AS org_slug,           -- not in DB (see below)
  o.created_at                                    AS org_created_at,
  picked.owner_user_id                            AS owner_user_id,
  u.email                                         AS owner_email,
  u.name                                          AS owner_name,
  if(coalesce(sp.is_paying,0)=1 OR mg.org_id IS NOT NULL,'yes','no')       AS is_paying,
  if(mg.org_id IS NOT NULL AND (coalesce(mg.full_seats,0)+coalesce(mg.lite_seats,0))>0, coalesce(mg.full_seats,0)+coalesce(mg.lite_seats,0), coalesce(picked.seats,0))                        AS seats,
  coalesce(
    nullIf(pn.promo_code_label, ''),
    if(spr.coupon_name != '', concat(spr.coupon_name,
      multiIf(spr.percent_off>0, concat(' (', toString(toInt(round(spr.percent_off))), '% off)'),
              spr.amount_off>0,  concat(' ($', toString(toInt(round(spr.amount_off/100))), ' off)'), '')), NULL),
    nullIf(promo.promo_code, '')
  )                                               AS promo_code,
  cust.email                                      AS billing_email,
  picked.stripe_customer_id                       AS billing_customer_id,
  sub_agg.stripe_subscription_ids                 AS stripe_subscription_ids,
  picked.stripe_customer_id                       AS stripe_customer_id,
  coalesce(sp.stripe_seats_paying_sum,0)          AS stripe_seats_paying_sum,
  picked.trial_ends_at                            AS trial_ends_at,
  hc.health_score                                 AS health_score,
  coalesce(promo.app_promo_codes,'')              AS app_promo_codes,
  o.updated_at                                    AS updated_at
FROM (SELECT id, name, created_at, updated_at, workos_id, external_id FROM postgres.orgs LIMIT 1 BY id) o
LEFT JOIN picked       ON picked.org_id = o.id
LEFT JOIN sub_agg      ON sub_agg.org_id = o.id
LEFT JOIN seats_paying sp ON sp.org_id = o.id
LEFT JOIN managed mg   ON mg.org_id = o.id
LEFT JOIN promo        ON promo.org_id = o.id
LEFT JOIN stripe_promo spr ON spr.sid = picked.stripe_subscription_id
LEFT JOIN promo_names pn ON pn.promotion_code_id = spr.promotion_code_id
LEFT JOIN u            ON u.id = picked.owner_user_id
LEFT JOIN cust         ON cust.id = picked.stripe_customer_id
LEFT JOIN hc           ON hc.workos_org_id = o.workos_id
ORDER BY is_paying DESC, org_name
LIMIT 5000
`

/**
 * Run the Middle Earth query and write the "Middle Earth" sheet.
 * Returns { rows_in, rows_out } for pipeline-style logging.
 */
function render_middle_earth() {
  const t0 = new Date()
  const ss = SpreadsheetApp.getActive()

  const props = PropertiesService.getScriptProperties()
  const apiKey = props.getProperty('POSTHOG_API_KEY')
  if (!apiKey) throw new Error('Missing POSTHOG_API_KEY in Script Properties')
  const projectId = props.getProperty('POSTHOG_PROJECT_ID') || POSTHOG_RAW_CFG.PROJECT_ID_FALLBACK

  // Reuse the columns-aware HogQL runner (defined in Sauron Query.js).
  const { columns, results } = sauronQueryRun_(apiKey, projectId, MIDDLE_EARTH_HOGQL, 'render_middle_earth')

  const headers = (columns && columns.length) ? columns : ['(no columns returned)']

  const rows = (results || []).map(r => {
    const row = new Array(headers.length)
    for (let i = 0; i < headers.length; i++) row[i] = (r && r[i] != null) ? r[i] : ''
    return row
  })

  const sh = ss.getSheetByName(MIDDLE_EARTH_CFG.OUT_SHEET) || ss.insertSheet(MIDDLE_EARTH_CFG.OUT_SHEET)
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

  Logger.log(`render_middle_earth: ${rows.length} rows in ${((new Date() - t0) / 1000).toFixed(1)}s`)
  return { rows_in: rows.length, rows_out: rows.length }
}
