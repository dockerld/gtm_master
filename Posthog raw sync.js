/**************************************************************
 * PostHog Raw Sync (overwrite-only) — HARDENED AGAINST 504s
 *
 * Creates/overwrites:
 *  - raw_posthog_user_metrics
 *
 * Pulls (per email_key):
 *  - meetings_recorded
 *  - hours_recorded
 *  - ask_meeting
 *  - ask_global
 *  - stripe_subscription_id (from Postgres Stripe subscription metadata.orgId -> users_orgs)
 *  - client_page_views
 *  - active_days (distinct days with ANY events)
 *  - clients_count (ORG-level; fanned back to user)
 *  - calendar_connected + first_calendar_connected_date
 *  - email_connected + first_email_connected_date
 *  - PM providers + first connected dates (one column per provider)
 *  - other_integrations (comma-separated list of providers not in the core set)
 *  - action_items_synced
 *  - meeting_notes_synced
 *
 * Key improvements:
 * 1) Retries + exponential backoff for PostHog 429/502/503/504 (common transient failures)
 * 2) Smaller batches + longer pauses (reduces load)
 * 3) Event queries bounded by a lookback window (reduces scan size → fewer timeouts)
 * 4) Writes partial progress only at end (same behavior), but logs where it failed
 *
 * Uses Script Properties:
 *  - POSTHOG_API_KEY
 *  - POSTHOG_PROJECT_ID (optional; falls back to config)
 *
 * Email source:
 *  - Reads emails from raw_clerk_users by default (recommended)
 *
 * Notes:
 * - No semicolons in HogQL
 * - Overwrite-only raw table. Canon tables handle editability rules.
 **************************************************************/

const POSTHOG_RAW_CFG = {
  PROJECT_ID_FALLBACK: '179975',
  API_BASE: 'https://app.posthog.com/api',

  // ↓ Lower batch size helps avoid large query payloads/timeouts
  BATCH_SIZE: 100,

  // ↓ Slightly more breathing room between batches
  PAUSE_MS: 300,

  // ↓ Write chunking to Sheets
  WRITE_CHUNK: 5000,

  // ↓ Event query bounds (reduce workload). Adjust if you truly need “all time”.
  EVENT_LOOKBACK_DAYS: 365,

  // ↓ Retry policy for PostHog API
  RETRY: {
    MAX_ATTEMPTS: 6,          // total attempts per query
    BASE_SLEEP_MS: 750,       // backoff base
    MAX_SLEEP_MS: 15000,      // cap
    JITTER_MS: 250            // jitter to avoid thundering herd
  },

  // Different environments may expose this table as singular or plural.
  // We try these candidates in order and cache the first one that works.
  STRIPE_SUB_TABLE_CANDIDATES: [
    'postgres.stripe_subscription',
    'postgres.stripe_subscriptions'
  ],

  SHEETS: {
    SOURCE_USERS: 'raw_clerk_users',
    DEST: 'raw_posthog_user_metrics',
    DEST_ORGS: 'raw_posthog_orgs',
    DEST_ORG_SUBSCRIPTIONS: 'raw_posthog_org_subscriptions',
    DEST_PROMO_REDEMPTIONS: 'promo_redemptions'
  },

  SOURCE_HEADERS: {
    EMAIL: 'email'
  }
}

let POSTHOG_STRIPE_TABLE_RESOLVED = null

function posthog_pull_user_metrics_to_raw() {
  const t0 = new Date()
  const props = PropertiesService.getScriptProperties()

  const apiKey = props.getProperty('POSTHOG_API_KEY')
  if (!apiKey) throw new Error('Missing POSTHOG_API_KEY in Script Properties')

  const projectId = props.getProperty('POSTHOG_PROJECT_ID') || POSTHOG_RAW_CFG.PROJECT_ID_FALLBACK
  if (!projectId) throw new Error('Missing POSTHOG_PROJECT_ID (or set PROJECT_ID_FALLBACK)')

  const ss = SpreadsheetApp.getActive()

  // 1) Get email list (prefer raw_clerk_users)
  const source = ss.getSheetByName(POSTHOG_RAW_CFG.SHEETS.SOURCE_USERS)
  if (!source) throw new Error(`Source sheet not found: ${POSTHOG_RAW_CFG.SHEETS.SOURCE_USERS}`)

  const emails = posthogReadEmails_(source, 1, POSTHOG_RAW_CFG.SOURCE_HEADERS.EMAIL)
  const uniqueEmailKeys = Array.from(new Set(emails.filter(Boolean).map(e => normalizeEmail(e))))

  Logger.log(`PostHog: unique emails to query = ${uniqueEmailKeys.length}`)
  if (uniqueEmailKeys.length === 0) {
    posthogWriteSyncLogSafe_(
      'posthog_pull_user_metrics_to_raw',
      'ok',
      0,
      0,
      (new Date() - t0) / 1000,
      'no emails to query'
    )
    return { rows_in: 0, rows_out: 0 }
  }

  // 2) Query PostHog in batches
  const metricsMap = new Map()     // email_key -> record (db metrics)
  const pageViewsMap = new Map()   // email_key -> client_page_views
  const activeDaysMap = new Map()  // email_key -> active_days (ANY events)

  for (let i = 0; i < uniqueEmailKeys.length; i += POSTHOG_RAW_CFG.BATCH_SIZE) {
    const batch = uniqueEmailKeys.slice(i, i + POSTHOG_RAW_CFG.BATCH_SIZE)
    const batchNum = Math.floor(i / POSTHOG_RAW_CFG.BATCH_SIZE) + 1
    Logger.log(`PostHog batch ${batchNum}: ${batch.length} emails`)

    // 2a) DB-backed metrics (postgres.* tables via HogQL)
    {
      const rows = posthogRunDbMetricsQueryWithStripe_(apiKey, projectId, batch, batchNum)

      rows.forEach(r => {
        const emailKey = normalizeEmail(r?.[0] || '')
        if (!emailKey) return

        metricsMap.set(emailKey, {
          email_key: emailKey,
          email: String(r?.[1] || ''),

          meetings_recorded: Number(r?.[2] ?? 0),
          hours_recorded: Number(r?.[3] ?? 0),
          ask_meeting: Number(r?.[4] ?? 0),
          ask_global: Number(r?.[5] ?? 0),
          clients_count: Number(r?.[6] ?? 0),

          calendar_connected: String(r?.[7] || '').toLowerCase() === 'yes',
          first_calendar_connected_date: String(r?.[8] || ''),

          email_connected: String(r?.[9] || '').toLowerCase() === 'yes',
          first_email_connected_date: String(r?.[10] || ''),

          pm_karbon_connected: String(r?.[11] || '').toLowerCase() === 'yes',
          pm_karbon_first_connected_date: String(r?.[12] || ''),

          pm_keeper_connected: String(r?.[13] || '').toLowerCase() === 'yes',
          pm_keeper_first_connected_date: String(r?.[14] || ''),

          pm_financial_cents_connected: String(r?.[15] || '').toLowerCase() === 'yes',
          pm_financial_cents_first_connected_date: String(r?.[16] || ''),

          other_integrations: String(r?.[17] || ''),

          action_items_synced: Number(r?.[18] ?? 0),
          meeting_notes_synced: Number(r?.[19] ?? 0),
          stripe_subscription_id: String(r?.[20] || '')
        })
      })
    }

    // 2b) Event metrics (client page views from events/persons) — bounded by lookback
    {
      const sql = posthogBuildHogQL_clientPageViews_(batch, POSTHOG_RAW_CFG.EVENT_LOOKBACK_DAYS)
      const rows = posthogRunQuery_(apiKey, projectId, sql, `clientPageViews batch ${batchNum}`)

      rows.forEach(r => {
        const emailKey = normalizeEmail(r?.[0] || '')
        if (!emailKey) return
        pageViewsMap.set(emailKey, Number(r?.[1] ?? 0))
      })
    }

    // 2c) Active days (ANY events) — bounded by lookback
    {
      const sql = posthogBuildHogQL_activeDays_(batch, POSTHOG_RAW_CFG.EVENT_LOOKBACK_DAYS)
      const rows = posthogRunQuery_(apiKey, projectId, sql, `activeDays batch ${batchNum}`)

      rows.forEach(r => {
        const emailKey = normalizeEmail(r?.[0] || '')
        if (!emailKey) return
        activeDaysMap.set(emailKey, Number(r?.[1] ?? 0))
      })
    }

    Utilities.sleep(POSTHOG_RAW_CFG.PAUSE_MS)
  }

  // 3) Build output rows (one row per email_key we queried)
  const headers = [
    'email_key',
    'email',

    'meetings_recorded',
    'hours_recorded',
    'ask_meeting',
    'ask_global',
    'client_page_views',
    'active_days',
    'clients_count',

    'calendar_connected',
    'first_calendar_connected_date',
    'email_connected',
    'first_email_connected_date',

    'pm_karbon_connected',
    'pm_karbon_first_connected_date',
    'pm_keeper_connected',
    'pm_keeper_first_connected_date',
    'pm_financial_cents_connected',
    'pm_financial_cents_first_connected_date',
    'other_integrations',
    'action_items_synced',
    'meeting_notes_synced',
    'stripe_subscription_id',

    'pulled_at'
  ]

  const pulledAt = new Date()
  const rowsOut = uniqueEmailKeys.map(emailKey => {
    const base = metricsMap.get(emailKey) || {
      email_key: emailKey,
      email: '',

      meetings_recorded: 0,
      hours_recorded: 0,
      ask_meeting: 0,
      ask_global: 0,
      clients_count: 0,

      calendar_connected: false,
      first_calendar_connected_date: '',

      email_connected: false,
      first_email_connected_date: '',

      pm_karbon_connected: false,
      pm_karbon_first_connected_date: '',

      pm_keeper_connected: false,
      pm_keeper_first_connected_date: '',

      pm_financial_cents_connected: false,
      pm_financial_cents_first_connected_date: '',

      other_integrations: '',

      action_items_synced: 0,
      meeting_notes_synced: 0,
      stripe_subscription_id: ''
    }

    const clientViews = pageViewsMap.has(emailKey) ? pageViewsMap.get(emailKey) : 0
    const activeDays  = activeDaysMap.has(emailKey) ? activeDaysMap.get(emailKey) : 0

    return [
      base.email_key,
      base.email,

      base.meetings_recorded,
      base.hours_recorded,
      base.ask_meeting,
      base.ask_global,
      clientViews,
      activeDays,
      base.clients_count,

      base.calendar_connected,
      base.first_calendar_connected_date || '',
      base.email_connected,
      base.first_email_connected_date || '',

      base.pm_karbon_connected,
      base.pm_karbon_first_connected_date || '',
      base.pm_keeper_connected,
      base.pm_keeper_first_connected_date || '',
      base.pm_financial_cents_connected,
      base.pm_financial_cents_first_connected_date || '',
      base.other_integrations,
      base.action_items_synced,
      base.meeting_notes_synced,
      base.stripe_subscription_id || '',

      pulledAt
    ]
  })

  // 4) Overwrite destination
  const dest = getOrCreateSheetSafe_(ss, POSTHOG_RAW_CFG.SHEETS.DEST)
  posthogOverwriteSheet_(dest, headers, rowsOut)

  posthogWriteSyncLogSafe_(
    'posthog_pull_user_metrics_to_raw',
    'ok',
    uniqueEmailKeys.length,
    rowsOut.length,
    (new Date() - t0) / 1000,
    `lookback_days=${POSTHOG_RAW_CFG.EVENT_LOOKBACK_DAYS} batch_size=${POSTHOG_RAW_CFG.BATCH_SIZE}`
  )

  return { rows_in: uniqueEmailKeys.length, rows_out: rowsOut.length }
}

/**
 * Pull org-level subscription data from PostHog Postgres tables.
 * Writes overwrite-only sheet: raw_posthog_org_subscriptions
 */
function posthog_pull_org_subscriptions_to_raw() {
  const t0 = new Date()
  const props = PropertiesService.getScriptProperties()

  const apiKey = props.getProperty('POSTHOG_API_KEY')
  if (!apiKey) throw new Error('Missing POSTHOG_API_KEY in Script Properties')

  const projectId = props.getProperty('POSTHOG_PROJECT_ID') || POSTHOG_RAW_CFG.PROJECT_ID_FALLBACK
  if (!projectId) throw new Error('Missing POSTHOG_PROJECT_ID (or set PROJECT_ID_FALLBACK)')

  const ss = SpreadsheetApp.getActive()

  const rows = posthogQueryOrgSubscriptions_(apiKey, projectId)
  const sourceTag = 'posthog'
  const pulledAt = new Date()
  const stripeSeatsBySubId = posthogBuildStripeSeatsBySubscriptionId_(ss)
  const stripeDetailsBySubId = posthogBuildStripeSubscriptionDetailsById_(ss)

  const headers = [
    'id',
    'app_org_id',
    'org_id',
    'stripe_customer_id',
    'stripe_subscription_id',
    'status',
    'owner_user_id',
    'full_seat_count',
    'lite_seat_count',
    'billing_interval',
    'current_period_start',
    'current_period_end',
    'cancel_at_period_end',
    'trial_started_at',
    'trial_ends_at',
    'has_used_trial',
    'created_at',
    'updated_at',
    'is_active',
    'is_canceled',
    'stripe_interval',
    'stripe_interval_count',
    'stripe_amount',
    'stripe_amount_yearly',
    'stripe_discount_percent_all',
    'stripe_discount_amount_off_all',
    'stripe_discount_duration_all',
    'stripe_promo_code_all',
    'stripe_promo_code',
    'pulled_at'
  ]

  const rowsOut = (rows || []).map(r => {
    const stripeSubscriptionId = String(r && r[3] != null ? r[3] : '')
    const status = String(r && r[4] != null ? r[4] : '').toLowerCase()
    let fullSeatCount = posthogToNumOrZero_(r && r[6])
    const liteSeatCount = posthogToNumOrZero_(r && r[7])
    const stripeMeta = stripeSubscriptionId ? (stripeDetailsBySubId.get(stripeSubscriptionId) || {}) : {}

    // Some PostHog org_subscriptions schemas omit seat fields;
    // in that case, use Stripe quantity_total as best-effort full seats.
    if (fullSeatCount <= 0 && liteSeatCount <= 0 && stripeSubscriptionId && stripeSeatsBySubId.has(stripeSubscriptionId)) {
      fullSeatCount = stripeSeatsBySubId.get(stripeSubscriptionId) || 0
    }

    return [
      String(r && r[0] != null ? r[0] : ''),
      String(r && r[1] != null ? r[1] : ''),
      String(r && r[1] != null ? r[1] : ''),
      String(r && r[2] != null ? r[2] : ''),
      stripeSubscriptionId,
      String(r && r[4] != null ? r[4] : ''),
      String(r && r[5] != null ? r[5] : ''),
      fullSeatCount,
      liteSeatCount,
      String(r && r[8] != null ? r[8] : ''),
      String(r && r[9] != null ? r[9] : ''),
      String(r && r[10] != null ? r[10] : ''),
      String(r && r[11] != null ? r[11] : ''),
      String(r && r[12] != null ? r[12] : ''),
      String(r && r[13] != null ? r[13] : ''),
      String(r && r[14] != null ? r[14] : ''),
      String(r && r[15] != null ? r[15] : ''),
      String(r && r[16] != null ? r[16] : ''),
      status === 'active',
      status === 'canceled',
      String(stripeMeta.interval || ''),
      posthogToNumOrZero_(stripeMeta.interval_count),
      posthogToNumOrZero_(stripeMeta.amount),
      posthogToNumOrZero_(stripeMeta.amount_yearly),
      String(stripeMeta.discount_percent_all || ''),
      String(stripeMeta.discount_amount_off_all || ''),
      String(stripeMeta.discount_duration_all || ''),
      String(stripeMeta.promo_code_all || ''),
      String(stripeMeta.promo_code || ''),
      pulledAt
    ]
  })

  const dest = getOrCreateSheetSafe_(ss, POSTHOG_RAW_CFG.SHEETS.DEST_ORG_SUBSCRIPTIONS)
  posthogOverwriteSheet_(dest, headers, rowsOut)

  posthogWriteSyncLogSafe_(
    'posthog_pull_org_subscriptions_to_raw',
    'ok',
    rowsOut.length,
    rowsOut.length,
    (new Date() - t0) / 1000,
    `source=${sourceTag}`
  )

  return { rows_in: rowsOut.length, rows_out: rowsOut.length }
}

/**
 * Pull org-level metadata (all orgs) and append subscription rollups.
 * Writes overwrite-only sheet: raw_posthog_orgs
 */
function posthog_pull_orgs_to_raw() {
  const t0 = new Date()
  const props = PropertiesService.getScriptProperties()

  const apiKey = props.getProperty('POSTHOG_API_KEY')
  if (!apiKey) throw new Error('Missing POSTHOG_API_KEY in Script Properties')

  const projectId = props.getProperty('POSTHOG_PROJECT_ID') || POSTHOG_RAW_CFG.PROJECT_ID_FALLBACK
  if (!projectId) throw new Error('Missing POSTHOG_PROJECT_ID (or set PROJECT_ID_FALLBACK)')

  const ss = SpreadsheetApp.getActive()
  const rows = posthogQueryOrgs_(apiKey, projectId)
  const pulledAt = new Date()
  const subSheet = ss.getSheetByName(POSTHOG_RAW_CFG.SHEETS.DEST_ORG_SUBSCRIPTIONS)
  const subRows = subSheet ? posthogReadSheetObjectsSafe_(subSheet, 1) : []
  const subAggByOrgId = posthogBuildOrgSubAggByOrgId_(subRows)

  const headers = [
    'app_org_id',
    'org_id',
    'org_name',
    'org_status',
    'billing_email',
    'owner_user_id',
    'created_at',
    'updated_at',
    'subscription_count',
    'latest_subscription_id',
    'latest_subscription_status',
    'latest_subscription_updated_at',
    'active_subscription_count',
    'latest_active_subscription_id',
    'latest_active_subscription_updated_at',
    'pulled_at'
  ]

  const seen = new Set()
  const orgMetaById = new Map()
  ;(rows || []).forEach(r => {
    const orgId = String(r && r[0] != null ? r[0] : '').trim()
    if (!orgId) return
    seen.add(orgId)

    const orgName = String(r && r[1] != null ? r[1] : '')
    const orgStatus = String(r && r[2] != null ? r[2] : '').toLowerCase().trim()
    const ownerUserId = String(r && r[3] != null ? r[3] : '')
    const billingEmail = String(r && r[4] != null ? r[4] : '')
    const createdAt = String(r && r[5] != null ? r[5] : '')
    const updatedAt = String(r && r[6] != null ? r[6] : '')
    const subSortAt = String(r && r[7] != null ? r[7] : '')

    if (!orgMetaById.has(orgId)) {
      orgMetaById.set(orgId, {
        org_name: orgName,
        org_status: orgStatus,
        owner_user_id: ownerUserId,
        billing_email: billingEmail,
        created_at: createdAt,
        updated_at: updatedAt,
        _sort_at: subSortAt
      })
      return
    }

    const cur = orgMetaById.get(orgId)
    if (!cur.org_name && orgName) cur.org_name = orgName
    if (!cur.created_at && createdAt) cur.created_at = createdAt
    if (!cur.updated_at && updatedAt) cur.updated_at = updatedAt

    const curScore = posthogOrgStatusRank_(cur.org_status)
    const newScore = posthogOrgStatusRank_(orgStatus)
    const curTs = posthogIsoToMs_(cur._sort_at)
    const newTs = posthogIsoToMs_(subSortAt)
    if (newScore > curScore || (newScore === curScore && newTs >= curTs)) {
      cur.org_status = orgStatus || cur.org_status
      cur.owner_user_id = ownerUserId || cur.owner_user_id
      cur.billing_email = billingEmail || cur.billing_email
      cur._sort_at = subSortAt || cur._sort_at
    }
  })

  const rowsOut = []
  orgMetaById.forEach((meta, orgId) => {
    const agg = subAggByOrgId.get(orgId) || posthogEmptyOrgSubAgg_()
    rowsOut.push([
      orgId,
      orgId,
      String(meta.org_name || ''),
      String(meta.org_status || ''),
      String(meta.billing_email || ''),
      String(meta.owner_user_id || ''),
      String(meta.created_at || ''),
      String(meta.updated_at || ''),
      agg.subscription_count,
      agg.latest_subscription_id,
      agg.latest_subscription_status,
      agg.latest_subscription_updated_at,
      agg.active_subscription_count,
      agg.latest_active_subscription_id,
      agg.latest_active_subscription_updated_at,
      pulledAt
    ])
  })

  subAggByOrgId.forEach((agg, orgId) => {
    if (!orgId || seen.has(orgId)) return
    rowsOut.push([
      orgId,
      orgId,
      '',
      '',
      '',
      '',
      '',
      '',
      agg.subscription_count,
      agg.latest_subscription_id,
      agg.latest_subscription_status,
      agg.latest_subscription_updated_at,
      agg.active_subscription_count,
      agg.latest_active_subscription_id,
      agg.latest_active_subscription_updated_at,
      pulledAt
    ])
  })

  const dest = getOrCreateSheetSafe_(ss, POSTHOG_RAW_CFG.SHEETS.DEST_ORGS)
  posthogOverwriteSheet_(dest, headers, rowsOut)

  posthogWriteSyncLogSafe_(
    'posthog_pull_orgs_to_raw',
    'ok',
    rowsOut.length,
    rowsOut.length,
    (new Date() - t0) / 1000,
    ''
  )

  return { rows_in: rowsOut.length, rows_out: rowsOut.length }
}

/**
 * Pull org-level promo redemption data from PostHog Postgres tables.
 * Writes append-only sheet: promo_redemptions
 */
function posthog_pull_promo_redemptions_to_raw() {
  const t0 = new Date()
  const props = PropertiesService.getScriptProperties()

  const apiKey = props.getProperty('POSTHOG_API_KEY')
  if (!apiKey) throw new Error('Missing POSTHOG_API_KEY in Script Properties')

  const projectId = props.getProperty('POSTHOG_PROJECT_ID') || POSTHOG_RAW_CFG.PROJECT_ID_FALLBACK
  if (!projectId) throw new Error('Missing POSTHOG_PROJECT_ID (or set PROJECT_ID_FALLBACK)')

  const ss = SpreadsheetApp.getActive()
  const pulledAt = new Date()
  const inAppRows = posthogBuildInAppPromoRows_(posthogQueryPromoRedemptions_(apiKey, projectId), pulledAt)
  const stripeRows = posthogBuildStripePromoRows_(ss, pulledAt)
  const allRows = inAppRows.concat(stripeRows)

  const headers = posthogPromoRedemptionHeaders_()
  const dest = getOrCreateSheetSafe_(ss, POSTHOG_RAW_CFG.SHEETS.DEST_PROMO_REDEMPTIONS)
  const upsert = posthogUpsertPromoRedemptions_(dest, headers, allRows)

  posthogWriteSyncLogSafe_(
    'posthog_pull_promo_redemptions_to_raw',
    'ok',
    allRows.length,
    upsert.inserted + upsert.updated,
    (new Date() - t0) / 1000,
    `in_app=${inAppRows.length} stripe=${stripeRows.length} inserted=${upsert.inserted} updated=${upsert.updated}`
  )

  return { rows_in: allRows.length, rows_out: upsert.inserted + upsert.updated }
}

function posthogPromoRedemptionHeaders_() {
  return [
    'source_key',
    'redemption_location',
    'id',
    'app_org_id',
    'org_id',
    'stripe_subscription_id',
    'stripe_customer_id',
    'redeemed_at',
    'promo_code',
    'promo_name',
    'trial_days',
    'discount_percent',
    'discount_duration',
    'discount_duration_months',
    'discount_start_at',
    'discount_end_at',
    'promo_type',
    'pulled_at'
  ]
}

function posthogBuildInAppPromoRows_(rows, pulledAt) {
  return (rows || []).map(r => {
    const id = String(r && r[0] != null ? r[0] : '')
    const orgId = String(r && r[1] != null ? r[1] : '')
    const redeemedAt = String(r && r[2] != null ? r[2] : '')
    const promoCode = String(r && r[3] != null ? r[3] : '')
    const promoName = String(r && r[4] != null ? r[4] : '')
    const trialDays = posthogToNumOrZero_(r && r[5])
    const promoType = String(r && r[6] != null ? r[6] : '')
    const sourceKey = posthogBuildPromoSourceKey_({
      location: 'in_app',
      id,
      orgId,
      redeemedAt,
      promoCode
    })
    return [
      sourceKey,
      'in_app',
      id,
      orgId,
      orgId,
      '',
      '',
      redeemedAt,
      promoCode,
      promoName,
      trialDays,
      '',
      '',
      '',
      '',
      '',
      promoType,
      pulledAt
    ]
  })
}

function posthogBuildStripePromoRows_(ss, pulledAt) {
  const shStripe = ss.getSheetByName('raw_stripe_subscriptions')
  if (!shStripe) return []

  const stripeRows = posthogReadSheetObjectsSafe_(shStripe, 1)
  const orgBySubId = posthogBuildOrgIdBySubscriptionFromCanon_(ss)
  const orgByCustomerId = posthogBuildOrgIdByCustomerFromCanon_(ss)
  const orgBySubIdFromPosthog = posthogBuildOrgIdBySubscriptionFromPosthogOrgSubs_(ss)
  const out = []

  ;(stripeRows || []).forEach(r => {
    const subId = String(r.stripe_subscription_id || r.subscription_id || r.id || '').trim()
    if (!subId) return
    const customerId = String(r.stripe_customer_id || r.customer_id || '').trim()
    const metadataOrgId = posthogExtractOrgIdFromStripeMetadata_(r.metadata_json)
    const mappedOrgId =
      String(metadataOrgId || '').trim() ||
      String(orgBySubId.get(subId) || '').trim() ||
      String(orgBySubIdFromPosthog.get(subId) || '').trim() ||
      String(orgByCustomerId.get(customerId) || '').trim()

    const codeAll = posthogCsvList_(r.promo_code_all)
    const startAll = posthogCsvList_(r.discount_start_at_all)
    const endAll = posthogCsvList_(r.discount_end_at_all)
    const pctAll = posthogCsvList_(r.discount_percent_all)
    const amtAll = posthogCsvList_(r.discount_amount_off_all)
    const durationAll = posthogCsvList_(r.discount_duration_all)
    const durationMonthsAll = posthogCsvList_(r.discount_duration_months_all)
    const promoSingle = String(r.promo_code || '').trim()
    const discountCount = posthogToNumOrZero_(r.discount_count)

    const maxN = Math.max(
      codeAll.length,
      startAll.length,
      endAll.length,
      pctAll.length,
      amtAll.length,
      durationAll.length,
      durationMonthsAll.length,
      discountCount
    )

    if (maxN <= 0 && !promoSingle) return

    const rowCount = Math.max(1, maxN)
    for (let i = 0; i < rowCount; i++) {
      const promoCode = String(codeAll[i] || promoSingle || '').trim()
      const redeemedAt =
        String(startAll[i] || '').trim() ||
        String(r.discount_start_at || '').trim() ||
        String(r.first_payment_at || '').trim() ||
        String(r.created_at || '').trim()
      const endAt = String(endAll[i] || '').trim() || String(r.discount_end_at || '').trim()
      const pct = posthogToNumOrZero_(pctAll[i] !== undefined ? pctAll[i] : r.discount_percent)
      const amt = posthogToNumOrZero_(amtAll[i] !== undefined ? amtAll[i] : r.discount_amount_off)
      const dur = String(durationAll[i] || r.discount_duration || '').trim()
      const durMonths = posthogToNumOrZero_(durationMonthsAll[i] !== undefined ? durationMonthsAll[i] : r.discount_duration_months)
      const promoName = promoCode
      const promoType = 'stripe_discount'

      // If no promo code and no discount signal on this slot, skip.
      if (!promoCode && pct <= 0 && amt <= 0 && !dur && !durMonths) continue

      const sourceKey = posthogBuildPromoSourceKey_({
        location: 'stripe',
        subId,
        promoCode: promoCode || '__discount__',
        redeemedAt,
        slot: i + 1
      })

      out.push([
        sourceKey,
        'stripe',
        '',
        mappedOrgId,
        mappedOrgId,
        subId,
        customerId,
        redeemedAt,
        promoCode,
        promoName,
        '',
        pct > 0 ? pct : '',
        dur,
        durMonths > 0 ? durMonths : '',
        redeemedAt,
        endAt,
        promoType,
        pulledAt
      ])
    }
  })

  return out
}

function posthogBuildPromoSourceKey_(parts) {
  const p = parts || {}
  const location = String(p.location || '').trim()
  if (location === 'in_app') {
    const id = String(p.id || '').trim()
    if (id) return `in_app:${id}`
    return `in_app:${String(p.orgId || '').trim()}:${String(p.promoCode || '').trim()}:${String(p.redeemedAt || '').trim()}`
  }
  if (location === 'stripe') {
    return `stripe:${String(p.subId || '').trim()}:${String(p.promoCode || '').trim()}:${String(p.redeemedAt || '').trim()}:slot${String(p.slot || 1)}`
  }
  return `${location}:${String(p.id || '').trim()}`
}

function posthogAppendUniquePromoRedemptions_(sheet, headers, rows) {
  const existingHeaders = posthogEnsureSheetHeaders_(sheet, headers)
  const keyIdx = existingHeaders.findIndex(h => String(h || '').trim().toLowerCase() === 'source_key')
  if (keyIdx < 0) throw new Error('promo_redemptions missing source_key header')

  const existing = new Set()
  const lastRow = sheet.getLastRow()
  if (lastRow >= 2) {
    const vals = sheet.getRange(2, keyIdx + 1, lastRow - 1, 1).getValues()
    vals.forEach(v => {
      const key = String(v && v[0] != null ? v[0] : '').trim()
      if (key) existing.add(key)
    })
  }

  const toAppend = []
  ;(rows || []).forEach(r => {
    const key = String(r && r[0] != null ? r[0] : '').trim()
    if (!key || existing.has(key)) return
    existing.add(key)
    toAppend.push(r)
  })

  if (!toAppend.length) return 0

  const startRow = Math.max(2, sheet.getLastRow() + 1)
  posthogBatchSetValuesSafe_(sheet, startRow, 1, toAppend, POSTHOG_RAW_CFG.WRITE_CHUNK || 5000)
  return toAppend.length
}

function posthogUpsertPromoRedemptions_(sheet, headers, rows) {
  const existingHeaders = posthogEnsureSheetHeaders_(sheet, headers)
  const keyIdx = existingHeaders.findIndex(h => String(h || '').trim().toLowerCase() === 'source_key')
  if (keyIdx < 0) throw new Error('promo_redemptions missing source_key header')

  const keyToRow = new Map()
  const lastRow = sheet.getLastRow()
  if (lastRow >= 2) {
    const vals = sheet.getRange(2, keyIdx + 1, lastRow - 1, 1).getValues()
    vals.forEach((v, i) => {
      const key = String(v && v[0] != null ? v[0] : '').trim()
      if (!key) return
      keyToRow.set(key, i + 2)
    })
  }

  const toAppend = []
  const toUpdate = []
  ;(rows || []).forEach(r => {
    const key = String(r && r[0] != null ? r[0] : '').trim()
    if (!key) return
    const rowNum = keyToRow.get(key)
    if (rowNum) {
      toUpdate.push({ rowNum, row: r })
    } else {
      keyToRow.set(key, -1) // prevent duplicate appends in same batch
      toAppend.push(r)
    }
  })

  // Update existing rows (full-row overwrite to keep schema aligned).
  toUpdate.forEach(u => {
    sheet.getRange(u.rowNum, 1, 1, u.row.length).setValues([u.row])
  })

  if (toAppend.length) {
    const startRow = Math.max(2, sheet.getLastRow() + 1)
    posthogBatchSetValuesSafe_(sheet, startRow, 1, toAppend, POSTHOG_RAW_CFG.WRITE_CHUNK || 5000)
  }

  return { inserted: toAppend.length, updated: toUpdate.length }
}

function posthogEnsureSheetHeaders_(sheet, headers) {
  const needed = headers || []
  if (!needed.length) return []

  const lastRow = sheet.getLastRow()
  const lastCol = Math.max(sheet.getLastColumn(), needed.length, 1)
  const existing = lastRow >= 1
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())
    : []

  let changed = false
  const out = existing.slice()
  needed.forEach(h => {
    if (out.indexOf(h) >= 0) return
    out.push(h)
    changed = true
  })

  if (out.length < needed.length) {
    while (out.length < needed.length) out.push('')
    changed = true
  }

  if (!lastRow) changed = true
  if (changed) {
    sheet.getRange(1, 1, 1, out.length).setValues([out])
  }
  return out.length ? out : needed.slice()
}

function posthogBatchSetValuesSafe_(sheet, startRow, startCol, values, chunkSize) {
  const size = Math.max(1, Number(chunkSize || 5000) || 5000)
  for (let i = 0; i < values.length; i += size) {
    const chunk = values.slice(i, i + size)
    sheet.getRange(startRow + i, startCol, chunk.length, chunk[0].length).setValues(chunk)
  }
}

function posthogBuildOrgIdBySubscriptionFromCanon_(ss) {
  const out = new Map()
  const sh = ss.getSheetByName('canon_orgs')
  if (!sh) return out
  const rows = posthogReadSheetObjectsSafe_(sh, 1)
  ;(rows || []).forEach(r => {
    const orgId = String(r.app_org_id || r.org_id || '').trim()
    if (!orgId) return
    posthogCsvList_(r.stripe_subscription_ids).forEach(subId => {
      const sid = String(subId || '').trim()
      if (!sid || out.has(sid)) return
      out.set(sid, orgId)
    })
  })
  return out
}

function posthogBuildOrgIdByCustomerFromCanon_(ss) {
  const out = new Map()
  const sh = ss.getSheetByName('canon_orgs')
  if (!sh) return out
  const rows = posthogReadSheetObjectsSafe_(sh, 1)
  ;(rows || []).forEach(r => {
    const orgId = String(r.app_org_id || r.org_id || '').trim()
    if (!orgId) return
    posthogCsvList_(r.stripe_customer_ids).forEach(customerId => {
      const cid = String(customerId || '').trim()
      if (!cid || out.has(cid)) return
      out.set(cid, orgId)
    })
    const billingCustomer = String(r.billing_customer_id || '').trim()
    if (billingCustomer && !out.has(billingCustomer)) out.set(billingCustomer, orgId)
  })
  return out
}

function posthogBuildOrgIdBySubscriptionFromPosthogOrgSubs_(ss) {
  const out = new Map()
  const sh = ss.getSheetByName(POSTHOG_RAW_CFG.SHEETS.DEST_ORG_SUBSCRIPTIONS)
  if (!sh) return out
  const rows = posthogReadSheetObjectsSafe_(sh, 1)
  ;(rows || []).forEach(r => {
    const subId = String(r.stripe_subscription_id || r.subscription_id || r.id || '').trim()
    const orgId = String(r.app_org_id || r.org_id || '').trim()
    if (!subId || !orgId || out.has(subId)) return
    out.set(subId, orgId)
  })
  return out
}

function posthogBuildStripeSubscriptionDetailsById_(ss) {
  const out = new Map()
  const sh = ss.getSheetByName('raw_stripe_subscriptions')
  if (!sh) return out
  const rows = posthogReadSheetObjectsSafe_(sh, 1)
  ;(rows || []).forEach(r => {
    const subId = String(r.stripe_subscription_id || r.subscription_id || r.id || '').trim()
    if (!subId) return
    out.set(subId, {
      interval: r.interval,
      interval_count: r.interval_count,
      amount: r.amount,
      amount_yearly: r.amount_yearly,
      discount_percent_all: r.discount_percent_all,
      discount_amount_off_all: r.discount_amount_off_all,
      discount_duration_all: r.discount_duration_all,
      promo_code_all: r.promo_code_all,
      promo_code: r.promo_code
    })
  })
  return out
}

function posthogExtractOrgIdFromStripeMetadata_(metadataRaw) {
  const raw = String(metadataRaw || '').trim()
  if (!raw) return ''
  try {
    const obj = JSON.parse(raw)
    if (!obj || typeof obj !== 'object') return ''
    return String(obj.orgId || obj.org_id || '').trim()
  } catch (err) {
    return ''
  }
}

function posthogCsvList_(v) {
  const s = String(v == null ? '' : v).trim()
  if (!s) return []
  return s.split(',').map(x => String(x || '').trim()).filter(Boolean)
}

/* =========================
 * HogQL builders
 * ========================= */

function posthogBuildHogQL_dbMetrics_(emailKeys, stripeSubTableExpr) {
  const quoted = emailKeys.map(e => `'${String(e).replace(/'/g, "''")}'`).join(', ')
  const stripeTable = String(stripeSubTableExpr || '').trim()

  const stripeCtes = stripeTable
    ? `
, stripe_sub_per_org AS (
  SELECT
    JSONExtractString(toString(ss.metadata), 'orgId') AS org_id,
    any(ss.id) AS stripe_subscription_id
  FROM ${stripeTable} AS ss
  WHERE length(coalesce(JSONExtractString(toString(ss.metadata), 'orgId'), '')) > 0
  GROUP BY org_id
)

, stripe_sub_per_user AS (
  SELECT
    uo.user_id,
    any(sso.stripe_subscription_id) AS stripe_subscription_id
  FROM user_orgs AS uo
  LEFT JOIN stripe_sub_per_org AS sso
    ON toString(sso.org_id) = toString(uo.org_id)
  GROUP BY uo.user_id
)
`
    : `
, stripe_sub_per_user AS (
  SELECT
    bu.user_id,
    '' AS stripe_subscription_id
  FROM base_users AS bu
)
`

  return `
WITH [${quoted}] AS input_emails

, base_users AS (
  SELECT
    u.id AS user_id,
    lower(u.email) AS email_key,
    u.email AS email
  FROM postgres.users AS u
  WHERE lower(u.email) IN (SELECT arrayJoin(input_emails))
)

 , meeting_bot_stats AS (
  SELECT
    mb.user_id,
    countIf(
      mb.recording_started_at IS NOT NULL
      AND mb.recording_ended_at IS NOT NULL
    ) AS meetings_recorded,
    round(
      sumIf(
        dateDiff('second', mb.recording_started_at, mb.recording_ended_at),
        mb.recording_started_at IS NOT NULL
        AND mb.recording_ended_at IS NOT NULL
      ) / 3600,
      2
    ) AS hours_recorded
  FROM postgres.meeting_bots AS mb
  GROUP BY mb.user_id
)

, ask_counts AS (
  SELECT
    t.resource_id AS user_id,
    countIf(t.id LIKE 'meeting%') AS ask_meeting,
    countIf(t.id LIKE 'global%')  AS ask_global
  FROM postgres.mastra.mastra_threads AS t
  GROUP BY t.resource_id
)

-- ORG-LEVEL client counts (no fan-out)
, user_orgs AS (
  SELECT DISTINCT
    bu.user_id,
    uo.org_id
  FROM base_users AS bu
  JOIN postgres.users_orgs AS uo
    ON uo.user_id = bu.user_id
)

, org_scope AS (
  SELECT DISTINCT org_id
  FROM user_orgs
)

, org_client_counts AS (
  SELECT
    os.org_id,
    countIf(
      c.id IS NOT NULL
      AND lower(trim(coalesce(c.name, ''))) != 'camden bean'
    ) AS clients_count
  FROM org_scope AS os
  LEFT JOIN postgres.clients AS c
    ON c.org_id = os.org_id
  GROUP BY os.org_id
)

, clients_count_per_user AS (
  SELECT
    uo.user_id,
    max(coalesce(occ.clients_count, 0)) AS clients_count
  FROM user_orgs AS uo
  LEFT JOIN org_client_counts AS occ
    ON occ.org_id = uo.org_id
  GROUP BY uo.user_id
)
${stripeCtes}

, oauth_rollup AS (
  SELECT
    oc.user_id,

    if(countIf(oc.scope_type = 'CALENDAR') > 0, 'yes', 'no') AS calendar_connected,
    formatDateTime(minIf(oc.created_at, oc.scope_type = 'CALENDAR'), '%Y-%m-%d') AS first_calendar_connected_date,

    if(countIf(oc.scope_type = 'EMAIL') > 0, 'yes', 'no') AS email_connected,
    formatDateTime(minIf(oc.created_at, oc.scope_type = 'EMAIL'), '%Y-%m-%d') AS first_email_connected_date,

    if(countIf(oc.provider = 'KARBON') > 0, 'yes', 'no') AS pm_karbon_connected,
    formatDateTime(minIf(oc.created_at, oc.provider = 'KARBON'), '%Y-%m-%d') AS pm_karbon_first_connected_date,

    if(countIf(oc.provider = 'KEEPER') > 0, 'yes', 'no') AS pm_keeper_connected,
    formatDateTime(minIf(oc.created_at, oc.provider = 'KEEPER'), '%Y-%m-%d') AS pm_keeper_first_connected_date,

    if(countIf(oc.provider = 'FINANCIAL_CENTS') > 0, 'yes', 'no') AS pm_financial_cents_connected,
    formatDateTime(minIf(oc.created_at, oc.provider = 'FINANCIAL_CENTS'), '%Y-%m-%d') AS pm_financial_cents_first_connected_date,

    arrayStringConcat(
      arraySort(
        groupUniqArrayIf(
          oc.provider,
          oc.provider NOT IN ('FINANCIAL_CENTS', 'KEEPER', 'KARBON', 'GOOGLE', 'COMPOSIO', 'MICROSOFT')
          AND length(coalesce(oc.provider, '')) > 0
        )
      ),
      ', '
    ) AS other_integrations

  FROM postgres.oauth_credentials AS oc
  GROUP BY oc.user_id
)

, action_items_counts AS (
  SELECT
    ai.user_id,
    countIf(ai.synced_to_practice_management = true) AS action_items_synced
  FROM postgres.action_items AS ai
  GROUP BY ai.user_id
)

, meeting_notes_counts AS (
  SELECT
    m.user_id,
    count() AS meeting_notes_synced
  FROM postgres.meetings AS m
  WHERE m.sync_status IS NOT NULL
  GROUP BY m.user_id
)

SELECT
  u.email_key,
  u.email,

  coalesce(mbs.meetings_recorded, 0) AS meetings_recorded,
  coalesce(mbs.hours_recorded, 0) AS hours_recorded,
  coalesce(a.ask_meeting, 0) AS ask_meeting,
  coalesce(a.ask_global, 0)  AS ask_global,

  coalesce(ccpu.clients_count, 0) AS clients_count,

  coalesce(o.calendar_connected, 'no') AS calendar_connected,
  coalesce(o.first_calendar_connected_date, '') AS first_calendar_connected_date,

  coalesce(o.email_connected, 'no') AS email_connected,
  coalesce(o.first_email_connected_date, '') AS first_email_connected_date,

  coalesce(o.pm_karbon_connected, 'no') AS pm_karbon_connected,
  coalesce(o.pm_karbon_first_connected_date, '') AS pm_karbon_first_connected_date,

  coalesce(o.pm_keeper_connected, 'no') AS pm_keeper_connected,
  coalesce(o.pm_keeper_first_connected_date, '') AS pm_keeper_first_connected_date,

  coalesce(o.pm_financial_cents_connected, 'no') AS pm_financial_cents_connected,
  coalesce(o.pm_financial_cents_first_connected_date, '') AS pm_financial_cents_first_connected_date,

  coalesce(o.other_integrations, '') AS other_integrations,

  coalesce(aic.action_items_synced, 0) AS action_items_synced,
  coalesce(mnc.meeting_notes_synced, 0) AS meeting_notes_synced,
  coalesce(sspu.stripe_subscription_id, '') AS stripe_subscription_id

FROM base_users AS u
LEFT JOIN meeting_bot_stats       AS mbs  ON mbs.user_id = u.user_id
LEFT JOIN ask_counts              AS a    ON a.user_id = u.user_id
LEFT JOIN clients_count_per_user  AS ccpu ON ccpu.user_id = u.user_id
LEFT JOIN oauth_rollup            AS o    ON o.user_id = u.user_id
LEFT JOIN action_items_counts     AS aic  ON aic.user_id = u.user_id
LEFT JOIN meeting_notes_counts    AS mnc  ON mnc.user_id = u.user_id
LEFT JOIN stripe_sub_per_user     AS sspu ON sspu.user_id = u.user_id

ORDER BY u.email_key
LIMIT 50000
  `.trim()
}

function posthogRunDbMetricsQueryWithStripe_(apiKey, projectId, emailKeys, batchNum) {
  const labelBase = `dbMetrics batch ${batchNum}`

  const candidates = POSTHOG_STRIPE_TABLE_RESOLVED
    ? [POSTHOG_STRIPE_TABLE_RESOLVED]
    : (POSTHOG_RAW_CFG.STRIPE_SUB_TABLE_CANDIDATES || []).slice()

  // Final fallback keeps the job running even if stripe table naming differs.
  candidates.push('')

  let lastErr = null
  for (const tableExpr of candidates) {
    const sql = posthogBuildHogQL_dbMetrics_(emailKeys, tableExpr)
    const tag = tableExpr ? tableExpr : 'no_stripe_join_fallback'

    try {
      const rows = posthogRunQuery_(apiKey, projectId, sql, `${labelBase} [${tag}]`)

      if (tableExpr && !POSTHOG_STRIPE_TABLE_RESOLVED) {
        POSTHOG_STRIPE_TABLE_RESOLVED = tableExpr
        Logger.log(`PostHog: using Stripe table "${tableExpr}" for stripe_subscription_id`)
      }
      return rows
    } catch (err) {
      lastErr = err

      if (POSTHOG_STRIPE_TABLE_RESOLVED) throw err
      if (!tableExpr) break

      Logger.log(`PostHog: dbMetrics failed with Stripe table "${tableExpr}" (${String(err && err.message ? err.message : err)}). Trying next fallback.`)
    }
  }

  throw lastErr || new Error('PostHog dbMetrics failed with all Stripe table fallbacks.')
}

function posthogBuildHogQL_clientPageViews_(emailKeys, lookbackDays) {
  const quoted = emailKeys.map(e => `'${String(e).replace(/'/g, "''")}'`).join(', ')
  const days = Math.max(1, Number(lookbackDays || 365) || 365)

  return `
WITH [${quoted}] AS input_emails
SELECT
  lower(p.properties.email) AS email_key,
  count() AS client_page_views
FROM events AS e
JOIN persons AS p
  ON e.person_id = p.id
WHERE e.event = '$pageview'
  AND e.timestamp >= now() - INTERVAL ${days} DAY
  AND lower(p.properties.email) IN input_emails
  AND (
    lower(e.properties.$current_url) LIKE '%/clients%'
    OR lower(e.properties.$pathname) LIKE '%/clients%'
  )
GROUP BY email_key
ORDER BY client_page_views DESC
LIMIT 50000
  `.trim()
}

function posthogBuildHogQL_activeDays_(emailKeys, lookbackDays) {
  const quoted = emailKeys.map(e => `'${String(e).replace(/'/g, "''")}'`).join(', ')
  const days = Math.max(1, Number(lookbackDays || 365) || 365)

  return `
WITH [${quoted}] AS input_emails
SELECT
  lower(p.properties.email) AS email_key,
  count(DISTINCT toDate(e.timestamp)) AS active_days
FROM events AS e
JOIN persons AS p
  ON e.person_id = p.id
WHERE e.timestamp >= now() - INTERVAL ${days} DAY
  AND lower(p.properties.email) IN (SELECT arrayJoin(input_emails))
GROUP BY email_key
ORDER BY active_days DESC
LIMIT 50000
  `.trim()
}

function posthogQueryOrgSubscriptions_(apiKey, projectId) {
  const tables = [
    'postgres.org_subscriptions',
    'postgres.org_subscription'
  ]

  let lastErr = null
  for (const tableExpr of tables) {
    // First try the richer schema.
    let sql = posthogBuildHogQL_orgSubscriptions_(tableExpr)
    let label = `orgSubscriptions rich [${tableExpr}]`
    try {
      return posthogRunQuery_(apiKey, projectId, sql, label)
    } catch (err) {
      lastErr = err
    }

    // Then try minimal shape for workspaces with fewer columns.
    sql = posthogBuildHogQL_orgSubscriptionsMinimal_(tableExpr)
    label = `orgSubscriptions minimal [${tableExpr}]`
    try {
      return posthogRunQuery_(apiKey, projectId, sql, label)
    } catch (err) {
      lastErr = err
    }
  }

  throw (lastErr || new Error('Could not query org_subscriptions tables'))
}

function posthogQueryPromoRedemptions_(apiKey, projectId) {
  let lastErr = null

  let sql = posthogBuildHogQL_promoRedemptions_()
  try {
    return posthogRunQuery_(apiKey, projectId, sql, 'promoRedemptions rich')
  } catch (err) {
    lastErr = err
  }

  sql = posthogBuildHogQL_promoRedemptionsMinimal_()
  try {
    return posthogRunQuery_(apiKey, projectId, sql, 'promoRedemptions minimal')
  } catch (err) {
    lastErr = err
  }

  throw (lastErr || new Error('Could not query promo_redemptions table'))
}

function posthogQueryOrgs_(apiKey, projectId) {
  const tableCandidates = ['postgres.orgs', 'postgres.organizations']
  let lastErr = null
  for (const tableExpr of tableCandidates) {
    let sql = posthogBuildHogQL_orgsJoined_(tableExpr)
    try {
      return posthogRunQuery_(apiKey, projectId, sql, `orgs joined [${tableExpr}]`)
    } catch (err) {
      lastErr = err
    }

    sql = posthogBuildHogQL_orgsCore_(tableExpr)
    try {
      return posthogRunQuery_(apiKey, projectId, sql, `orgs core [${tableExpr}]`)
    } catch (err) {
      lastErr = err
    }

    sql = posthogBuildHogQL_orgsMinimal_(tableExpr)
    try {
      return posthogRunQuery_(apiKey, projectId, sql, `orgs minimal [${tableExpr}]`)
    } catch (err) {
      lastErr = err
    }

    sql = posthogBuildHogQL_orgsIdOnly_(tableExpr)
    try {
      return posthogRunQuery_(apiKey, projectId, sql, `orgs id_only [${tableExpr}]`)
    } catch (err) {
      lastErr = err
    }
  }

  throw (lastErr || new Error('Could not query orgs table'))
}

function posthogBuildHogQL_orgSubscriptions_(tableExpr) {
  const t = String(tableExpr || '').trim()
  if (!t) throw new Error('Missing org_subscriptions table')

  return `
SELECT
  toString(os.id) AS id,
  toString(os.org_id) AS org_id,
  toString(os.stripe_customer_id) AS stripe_customer_id,
  toString(os.stripe_subscription_id) AS stripe_subscription_id,
  toString(os.status) AS status,
  toString(os.owner_user_id) AS owner_user_id,
  toFloat(os.full_seat_count) AS full_seat_count,
  toFloat(os.lite_seat_count) AS lite_seat_count,
  toString(os.billing_interval) AS billing_interval,
  toString(os.current_period_start) AS current_period_start,
  toString(os.current_period_end) AS current_period_end,
  toString(os.cancel_at_period_end) AS cancel_at_period_end,
  toString(os.trial_started_at) AS trial_started_at,
  toString(os.trial_ends_at) AS trial_ends_at,
  toString(os.has_used_trial) AS has_used_trial,
  toString(os.created_at) AS created_at,
  toString(os.updated_at) AS updated_at
FROM ${t} AS os
WHERE length(coalesce(toString(os.stripe_subscription_id), '')) > 0
ORDER BY os.updated_at DESC
LIMIT 300000
  `.trim()
}

function posthogBuildHogQL_orgSubscriptionsMinimal_(tableExpr) {
  const t = String(tableExpr || '').trim()
  if (!t) throw new Error('Missing org_subscriptions table')

  return `
SELECT
  toString(os.id) AS id,
  toString(os.org_id) AS org_id,
  '' AS stripe_customer_id,
  toString(os.stripe_subscription_id) AS stripe_subscription_id,
  '' AS status,
  '' AS owner_user_id,
  0 AS full_seat_count,
  0 AS lite_seat_count,
  '' AS billing_interval,
  '' AS current_period_start,
  '' AS current_period_end,
  '' AS cancel_at_period_end,
  '' AS trial_started_at,
  '' AS trial_ends_at,
  '' AS has_used_trial,
  '' AS created_at,
  '' AS updated_at
FROM ${t} AS os
WHERE length(coalesce(toString(os.stripe_subscription_id), '')) > 0
ORDER BY os.org_id, os.stripe_subscription_id
LIMIT 300000
  `.trim()
}

function posthogBuildHogQL_orgsMinimal_(tableExpr) {
  const t = String(tableExpr || '').trim()
  if (!t) throw new Error('Missing orgs table')
  return `
SELECT
  toString(o.id) AS org_id,
  toString(o.name) AS org_name,
  '' AS org_status,
  '' AS billing_email,
  '' AS owner_user_id,
  toString(o.created_at) AS created_at,
  toString(o.updated_at) AS updated_at
FROM ${t} AS o
ORDER BY o.id
LIMIT 300000
  `.trim()
}

function posthogBuildHogQL_orgsCore_(tableExpr) {
  const t = String(tableExpr || '').trim()
  if (!t) throw new Error('Missing orgs table')
  return `
SELECT
  toString(o.id) AS org_id,
  toString(o.name) AS org_name,
  '' AS org_status,
  '' AS billing_email,
  '' AS owner_user_id,
  toString(o.created_at) AS created_at,
  toString(o.updated_at) AS updated_at
FROM ${t} AS o
ORDER BY o.updated_at DESC
LIMIT 300000
  `.trim()
}

function posthogBuildHogQL_orgsIdOnly_(tableExpr) {
  const t = String(tableExpr || '').trim()
  if (!t) throw new Error('Missing orgs table')
  return `
SELECT
  toString(o.id) AS org_id,
  '' AS org_name,
  '' AS org_status,
  '' AS billing_email,
  '' AS owner_user_id,
  '' AS created_at,
  '' AS updated_at
FROM ${t} AS o
ORDER BY o.id
LIMIT 300000
  `.trim()
}

function posthogBuildHogQL_orgsJoined_(tableExpr) {
  const t = String(tableExpr || '').trim()
  if (!t) throw new Error('Missing orgs table')
  return `
SELECT
  toString(o.id) AS org_id,
  toString(o.name) AS org_name,
  toString(os.status) AS org_status,
  toString(os.owner_user_id) AS owner_user_id,
  toString(u.email) AS billing_email,
  toString(o.created_at) AS created_at,
  toString(o.updated_at) AS updated_at,
  toString(coalesce(os.current_period_end, os.updated_at, os.created_at)) AS sub_sort_at
FROM ${t} AS o
LEFT JOIN postgres.org_subscriptions AS os
  ON toString(os.org_id) = toString(o.id)
LEFT JOIN postgres.users AS u
  ON toString(u.id) = toString(os.owner_user_id)
ORDER BY o.created_at DESC
LIMIT 300000
  `.trim()
}

function posthogBuildHogQL_promoRedemptions_() {
  return `
SELECT
  toString(r.id) AS id,
  toString(r.org_id) AS org_id,
  toString(r.redeemed_at) AS redeemed_at,
  toString(p.code) AS promo_code,
  toString(p.name) AS promo_name,
  toString(p.trial_days) AS trial_days,
  toString(p.type) AS promo_type
FROM postgres.promo_redemptions AS r
LEFT JOIN postgres.promo_codes AS p
  ON toString(r.promo_code_id) = toString(p.id)
ORDER BY r.redeemed_at DESC
LIMIT 300000
  `.trim()
}

function posthogBuildHogQL_promoRedemptionsMinimal_() {
  return `
SELECT
  toString(r.id) AS id,
  toString(r.org_id) AS org_id,
  toString(r.redeemed_at) AS redeemed_at,
  toString(r.promo_code_id) AS promo_code,
  '' AS promo_name,
  0 AS trial_days,
  '' AS promo_type
FROM postgres.promo_redemptions AS r
ORDER BY r.redeemed_at DESC
LIMIT 300000
  `.trim()
}

function posthogBuildHogQL_orgSubscriptionsFromStripe_(tableExpr) {
  const t = String(tableExpr || '').trim()
  if (!t) throw new Error('Missing Stripe subscription table')

  return `
SELECT DISTINCT
  toString(ss.id) AS id,
  JSONExtractString(toString(ss.metadata), 'orgId') AS org_id,
  '' AS stripe_customer_id,
  toString(ss.id) AS stripe_subscription_id,
  '' AS status,
  '' AS owner_user_id,
  0 AS full_seat_count,
  0 AS lite_seat_count,
  '' AS billing_interval,
  '' AS current_period_start,
  '' AS current_period_end,
  '' AS cancel_at_period_end,
  '' AS trial_started_at,
  '' AS trial_ends_at,
  '' AS has_used_trial,
  '' AS created_at,
  '' AS updated_at
FROM ${t} AS ss
WHERE length(coalesce(JSONExtractString(toString(ss.metadata), 'orgId'), '')) > 0
ORDER BY org_id, stripe_subscription_id
LIMIT 300000
  `.trim()
}

function posthogReadSheetObjectsSafe_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1 || lastCol < 1) return []

  const header = sheet
    .getRange(headerRow, 1, 1, lastCol)
    .getValues()[0]
    .map(h => String(h || '').trim().toLowerCase().replace(/\s+/g, '_'))

  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()
  return data.map(r => {
    const obj = {}
    header.forEach((h, i) => {
      if (!h) return
      obj[h] = r[i]
    })
    return obj
  })
}

function posthogBuildStripeSeatsBySubscriptionId_(ss) {
  const out = new Map()
  const sh = ss.getSheetByName('raw_stripe_subscriptions')
  if (!sh) return out
  const rows = posthogReadSheetObjectsSafe_(sh, 1)
  ;(rows || []).forEach(r => {
    const subId = String(r.stripe_subscription_id || r.subscription_id || r.id || '').trim()
    if (!subId) return
    const qty = posthogToNumOrZero_(r.quantity_total)
    if (!out.has(subId) || qty > (out.get(subId) || 0)) out.set(subId, qty)
  })
  return out
}

function posthogToNumOrZero_(v) {
  const n = Number(v)
  return isFinite(n) ? n : 0
}

function posthogEmptyOrgSubAgg_() {
  return {
    subscription_count: 0,
    latest_subscription_id: '',
    latest_subscription_status: '',
    latest_subscription_updated_at: '',
    active_subscription_count: 0,
    latest_active_subscription_id: '',
    latest_active_subscription_updated_at: ''
  }
}

function posthogBuildOrgSubAggByOrgId_(subRows) {
  const out = new Map()
  ;(subRows || []).forEach(r => {
    const orgId = String(r.app_org_id || r.org_id || '').trim()
    const subId = String(r.stripe_subscription_id || r.subscription_id || r.id || '').trim()
    if (!orgId || !subId) return
    const status = String(r.status || '').toLowerCase().trim()
    const updatedAt = String(r.updated_at || r.created_at || '').trim()
    const ts = posthogIsoToMs_(updatedAt)

    if (!out.has(orgId)) out.set(orgId, posthogEmptyOrgSubAgg_())
    const agg = out.get(orgId)
    agg.subscription_count += 1

    const latestTs = posthogIsoToMs_(agg.latest_subscription_updated_at)
    if (!agg.latest_subscription_id || ts >= latestTs) {
      agg.latest_subscription_id = subId
      agg.latest_subscription_status = status
      agg.latest_subscription_updated_at = updatedAt
    }

    if (status === 'active') {
      agg.active_subscription_count += 1
      const latestActiveTs = posthogIsoToMs_(agg.latest_active_subscription_updated_at)
      if (!agg.latest_active_subscription_id || ts >= latestActiveTs) {
        agg.latest_active_subscription_id = subId
        agg.latest_active_subscription_updated_at = updatedAt
      }
    }
  })
  return out
}

function posthogIsoToMs_(isoLike) {
  const s = String(isoLike || '').trim()
  if (!s) return 0
  const d = new Date(s)
  return isNaN(d.getTime()) ? 0 : d.getTime()
}

function posthogOrgStatusRank_(status) {
  const s = String(status || '').toLowerCase().trim()
  if (s === 'active') return 3
  if (s === 'trialing') return 2
  if (s === 'canceled') return 1
  return 0
}

/* =========================
 * PostHog API + helpers
 * ========================= */

function posthogRunQuery_(apiKey, projectId, hogql, label) {
  const payload = { query: { kind: 'HogQLQuery', query: hogql } }
  const url = `${POSTHOG_RAW_CFG.API_BASE}/projects/${projectId}/query`

  const maxAttempts = POSTHOG_RAW_CFG.RETRY.MAX_ATTEMPTS
  const baseSleep = POSTHOG_RAW_CFG.RETRY.BASE_SLEEP_MS
  const maxSleep = POSTHOG_RAW_CFG.RETRY.MAX_SLEEP_MS
  const jitter = POSTHOG_RAW_CFG.RETRY.JITTER_MS

  let lastErr = null

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: `Bearer ${apiKey}` },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    })

    const code = res.getResponseCode()
    const text = res.getContentText() || ''

    if (code >= 200 && code < 300) {
      const json = JSON.parse(text)
      return json.results || []
    }

    // Retry on transient statuses
    const shouldRetry = (code === 429 || code === 502 || code === 503 || code === 504)
    lastErr = new Error(`PostHog API error ${code}: ${text}`)

    if (!shouldRetry || attempt === maxAttempts) break

    const sleepMs = Math.min(
      maxSleep,
      Math.floor(baseSleep * Math.pow(2, attempt - 1) + Math.random() * jitter)
    )

    Logger.log(
      `[PostHog retry] ${label || 'query'} attempt ${attempt}/${maxAttempts} got ${code}. Sleeping ${sleepMs}ms`
    )
    Utilities.sleep(sleepMs)
  }

  throw lastErr
}

function posthogReadEmails_(sheet, headerRow, emailHeaderName) {
  const { map } = readHeaderMap(sheet, headerRow)
  const cEmail = map[String(emailHeaderName).toLowerCase()]
  if (!cEmail) throw new Error(`posthogReadEmails_: header not found: ${emailHeaderName}`)

  const lastRow = sheet.getLastRow()
  if (lastRow < headerRow + 1) return []

  const n = lastRow - headerRow
  const vals = sheet.getRange(headerRow + 1, cEmail, n, 1).getValues()
  return vals.map(r => String(r[0] || '').trim()).filter(Boolean)
}

function posthogOverwriteSheet_(sheet, headers, rows) {
  sheet.clearContents()
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
  sheet.setFrozenRows(1)

  if (rows && rows.length) {
    if (typeof batchSetValues === 'function') batchSetValues(sheet, 2, 1, rows, POSTHOG_RAW_CFG.WRITE_CHUNK)
    else sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows)
  }
  sheet.autoResizeColumns(1, headers.length)
}

/* =========================
 * Minimal shared utilities (fallbacks)
 * ========================= */

function posthogWriteSyncLogSafe_(step, status, rowsIn, rowsOut, seconds, error) {
  if (typeof writeSyncLog === 'function') return writeSyncLog(step, status, rowsIn, rowsOut, seconds, error || '')
  Logger.log(`[SYNCLOG missing] ${step} ${status} rows_in=${rowsIn} rows_out=${rowsOut} seconds=${seconds} error=${error || ''}`)
}
