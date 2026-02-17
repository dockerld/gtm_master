/**************************************************************
 * One-time backfill helper for promo redemptions.
 *
 * Output sheet: promo_redemptions_backfill
 * Columns:
 * - redemption_id, org_id, org_name, redeemed_at, promo_code_id, promo_code,
 *   subscription_id, customer_id, customer_email
 **************************************************************/

const PROMO_BACKFILL_CFG = {
  INPUTS: {
    PROMO_REDEMPTIONS: 'promo_redemptions',
    PROMO_CODES: 'promo_codes',
    CLERK_USERS: 'raw_clerk_users',
    CLERK_MEMBERSHIPS: 'raw_clerk_memberships',
    CLERK_ORGS: 'raw_clerk_orgs',
    STRIPE_SUBS: 'raw_stripe_subscriptions'
  },
  OUTPUT_SHEET: 'promo_redemptions_backfill'
}

function render_promo_redemptions_backfill() {
  return ALLSTATS_lockWrap_('render_promo_redemptions_backfill', () => {
    const t0 = new Date()
    const ss = SpreadsheetApp.getActive()

    const shRedemptions = ss.getSheetByName(PROMO_BACKFILL_CFG.INPUTS.PROMO_REDEMPTIONS)

    const shPromoCodes = ss.getSheetByName(PROMO_BACKFILL_CFG.INPUTS.PROMO_CODES)
    const shUsers = ss.getSheetByName(PROMO_BACKFILL_CFG.INPUTS.CLERK_USERS)
    const shMems = ss.getSheetByName(PROMO_BACKFILL_CFG.INPUTS.CLERK_MEMBERSHIPS)
    const shOrgs = ss.getSheetByName(PROMO_BACKFILL_CFG.INPUTS.CLERK_ORGS)
    const shStripe = ss.getSheetByName(PROMO_BACKFILL_CFG.INPUTS.STRIPE_SUBS)
    if (!shStripe) throw new Error('Missing input sheet: raw_stripe_subscriptions')

    const redemptions = shRedemptions
      ? ALLSTATS_readSheetObjects_(shRedemptions, 1)
      : PRBACK_fetchPromoRedemptionsFromPosthog_()
    const promoCodes = shPromoCodes ? ALLSTATS_readSheetObjects_(shPromoCodes, 1) : []
    const users = shUsers ? ALLSTATS_readSheetObjects_(shUsers, 1) : []
    const mems = shMems ? ALLSTATS_readSheetObjects_(shMems, 1) : []
    const orgs = shOrgs ? ALLSTATS_readSheetObjects_(shOrgs, 1) : []
    const stripe = ALLSTATS_readSheetObjects_(shStripe, 1)

    const promoCodeById = ALLSTATS_buildPromoCodeLookupById_(promoCodes)
    const orgNameByOrgId = PRBACK_buildOrgNameByOrgId_(orgs)
    const subIdsByOrgId = PRBACK_buildSubIdsByOrgId_(users, mems)
    const stripeBySubId = PRBACK_buildStripeBySubId_(stripe)

    const outRows = []

    for (const r of redemptions) {
      const redemptionId = ALLSTATS_str_(r.id || r.redemption_id)
      const orgId = ALLSTATS_str_(r.org_id)
      const orgName = orgNameByOrgId.get(orgId) || ''
      const redeemedAt = ALLSTATS_toDateOrNull_(r.redeemed_at) || ALLSTATS_str_(r.redeemed_at)
      const promoCodeId = ALLSTATS_str_(r.promo_code_id || r.promocode_id || r.code_id)
      const promoCode =
        ALLSTATS_str_(r.promo_code) ||
        ALLSTATS_str_(r.code) ||
        ALLSTATS_str_(r.promo_name) ||
        (promoCodeId ? (promoCodeById.get(promoCodeId) || promoCodeId) : '')

      const subIds = Array.from(subIdsByOrgId.get(orgId) || [])
      if (!subIds.length) {
        outRows.push([
          redemptionId,
          orgId,
          orgName,
          redeemedAt,
          promoCodeId,
          promoCode,
          '',
          '',
          ''
        ])
        continue
      }

      for (const subId of subIds) {
        const s = stripeBySubId.get(subId) || {}
        const customerId = ALLSTATS_str_(s.stripe_customer_id || s.customer_id)
        const customerEmail = ALLSTATS_str_(s.customer_email || s.email || s.billing_email)

        outRows.push([
          redemptionId,
          orgId,
          orgName,
          redeemedAt,
          promoCodeId,
          promoCode,
          subId,
          customerId,
          customerEmail
        ])
      }
    }

    const out = ALLSTATS_getOrCreateSheet_(ss, PROMO_BACKFILL_CFG.OUTPUT_SHEET)
    out.clear()

    const headers = [
      'redemption_id',
      'org_id',
      'org_name',
      'redeemed_at',
      'promo_code_id',
      'promo_code',
      'subscription_id',
      'customer_id',
      'customer_email'
    ]

    out.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold')
      .setBackground('#F3F4F6')

    if (outRows.length) {
      out.getRange(2, 1, outRows.length, headers.length).setValues(outRows)
      out.getRange(2, 4, outRows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss')
    }

    out.setFrozenRows(1)
    out.autoResizeColumns(1, headers.length)

    const seconds = (new Date() - t0) / 1000
    ALLSTATS_writeSyncLog_('render_promo_redemptions_backfill', 'ok', redemptions.length, outRows.length, seconds, '')
    return { rows_in: redemptions.length, rows_out: outRows.length }
  })
}

function PRBACK_fetchPromoRedemptionsFromPosthog_() {
  const cfg = PRBACK_getPosthogConfig_()

  const redemptionTables = [
    'postgres.promo_redemptions',
    'postgres.promo_redemption'
  ]

  const promoTables = [
    'postgres.promo_codes',
    'postgres.promo_code',
    ''
  ]

  let lastErr = null
  for (const rTable of redemptionTables) {
    for (const pTable of promoTables) {
      const sql = PRBACK_buildPromoRedemptionsHogQL_(rTable, pTable)
      const label = `promo_redemptions [${rTable}${pTable ? ' + ' + pTable : ''}]`
      try {
        const rows = PRBACK_runPosthogQuery_(cfg.apiKey, cfg.projectId, sql, label)
        return rows.map(r => ({
          id: String(r && r[0] != null ? r[0] : ''),
          org_id: String(r && r[1] != null ? r[1] : ''),
          redeemed_at: String(r && r[2] != null ? r[2] : ''),
          promo_code_id: String(r && r[3] != null ? r[3] : ''),
          promo_code: String(r && r[4] != null ? r[4] : '')
        }))
      } catch (err) {
        lastErr = err
      }
    }
  }

  const msg = lastErr
    ? String(lastErr && lastErr.message ? lastErr.message : lastErr)
    : 'Unknown PostHog query failure.'
  throw new Error('Could not load promo_redemptions from PostHog tables or sheet. ' + msg)
}

function PRBACK_getPosthogConfig_() {
  const props = PropertiesService.getScriptProperties()
  const apiKey = props.getProperty('POSTHOG_API_KEY')
  if (!apiKey) throw new Error('Missing POSTHOG_API_KEY in Script Properties')

  let projectId = props.getProperty('POSTHOG_PROJECT_ID')
  if (!projectId && typeof POSTHOG_RAW_CFG !== 'undefined' && POSTHOG_RAW_CFG && POSTHOG_RAW_CFG.PROJECT_ID_FALLBACK) {
    projectId = POSTHOG_RAW_CFG.PROJECT_ID_FALLBACK
  }
  if (!projectId) throw new Error('Missing POSTHOG_PROJECT_ID (and no fallback configured)')

  return { apiKey, projectId }
}

function PRBACK_buildPromoRedemptionsHogQL_(redemptionTable, promoTable) {
  const rTable = String(redemptionTable || '').trim()
  const pTable = String(promoTable || '').trim()
  if (!rTable) throw new Error('Missing redemption table')

  const promoJoin = pTable
    ? `\nLEFT JOIN ${pTable} AS pc ON toString(pc.id) = toString(pr.promo_code_id)`
    : ''
  const promoExpr = pTable
    ? `coalesce(toString(pc.code), toString(pc.name), toString(pr.promo_code_id), '')`
    : `toString(pr.promo_code_id)`

  return `
SELECT
  toString(pr.id) AS redemption_id,
  toString(pr.org_id) AS org_id,
  toString(pr.redeemed_at) AS redeemed_at,
  toString(pr.promo_code_id) AS promo_code_id,
  ${promoExpr} AS promo_code
FROM ${rTable} AS pr${promoJoin}
ORDER BY pr.redeemed_at DESC
LIMIT 200000
  `.trim()
}

function PRBACK_runPosthogQuery_(apiKey, projectId, hogql, label) {
  if (typeof posthogRunQuery_ === 'function') {
    return posthogRunQuery_(apiKey, projectId, hogql, label)
  }

  const payload = { query: { kind: 'HogQLQuery', query: hogql } }
  const url = `https://app.posthog.com/api/projects/${projectId}/query`
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${apiKey}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  })
  const code = res.getResponseCode()
  const text = res.getContentText() || ''
  if (code < 200 || code >= 300) throw new Error(`PostHog API error ${code}: ${text}`)
  const json = JSON.parse(text)
  return json.results || []
}

function PRBACK_buildOrgNameByOrgId_(orgRows) {
  const out = new Map()
  for (const r of (orgRows || [])) {
    const orgId = ALLSTATS_str_(r.org_id)
    if (!orgId) continue
    const orgName = ALLSTATS_str_(r.org_name || r.org_slug)
    if (orgName) out.set(orgId, orgName)
  }
  return out
}

function PRBACK_buildSubIdsByOrgId_(userRows, membershipRows) {
  const subIdsByOrgId = new Map()

  const usersByEmailKey = new Map()
  for (const u of (userRows || [])) {
    const emailKey = ALLSTATS_str_(u.email_key) || ALLSTATS_normalizeEmail_(ALLSTATS_str_(u.email))
    if (!emailKey) continue

    const subId = ALLSTATS_str_(u.stripe_subscription_id || u.stripeSubscriptionId)
    if (subId) {
      if (!usersByEmailKey.has(emailKey)) usersByEmailKey.set(emailKey, new Set())
      usersByEmailKey.get(emailKey).add(subId)
    }

    const orgIdDirect = ALLSTATS_str_(u.org_id)
    if (orgIdDirect && subId) {
      if (!subIdsByOrgId.has(orgIdDirect)) subIdsByOrgId.set(orgIdDirect, new Set())
      subIdsByOrgId.get(orgIdDirect).add(subId)
    }
  }

  for (const m of (membershipRows || [])) {
    const orgId = ALLSTATS_str_(m.org_id)
    if (!orgId) continue

    const emailKey = ALLSTATS_str_(m.email_key) || ALLSTATS_normalizeEmail_(ALLSTATS_str_(m.email))
    if (!emailKey) continue

    const subIds = usersByEmailKey.get(emailKey)
    if (!subIds || !subIds.size) continue

    if (!subIdsByOrgId.has(orgId)) subIdsByOrgId.set(orgId, new Set())
    const bucket = subIdsByOrgId.get(orgId)
    subIds.forEach(id => bucket.add(id))
  }

  return subIdsByOrgId
}

function PRBACK_buildStripeBySubId_(stripeRows) {
  const out = new Map()
  for (const s of (stripeRows || [])) {
    const subId =
      ALLSTATS_str_(s.stripe_subscription_id) ||
      ALLSTATS_str_(s.subscription_id) ||
      ALLSTATS_str_(s.id)
    if (!subId) continue
    if (!out.has(subId)) out.set(subId, s)
  }
  return out
}
