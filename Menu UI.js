/**************************************************************
 * Custom Menu (Triggers / UI)
 *
 * Adds a menu to your spreadsheet:
 *  - Run daily pipeline
 *  - Run only PostHog
 *  - Run only Stripe
 *  - Run only Clerk
 *  - Rebuild canon tables
 *  - Push UpSale targets to Notion  ✅ NEW
 *
 * Notes:
 * - Each action uses LockService via lockWrap()
 * - Each action logs each step via writeSyncLog()
 * - Assumes you already have:
 *   run_daily_pipeline()
 *   stripe_pull_subscriptions_to_raw()
 *   posthog_pull_user_metrics_to_raw()
 *   clerk_pull_users_to_raw()
 *   clerk_pull_orgs_to_raw()
 *   clerk_pull_memberships_to_raw()
 *   syncClerkUsers()
 *   build_canon_orgs()
 *   build_canon_users()
 *   render_org_info_view()
 *   render_sauron_view()
 *   render_ring_view()
 *   render_arr_raw_data_view()
 *   write_arr_snapshot()
 *   render_arr_waterfall_facts()
 *   render_onboarding_stats()
 *   render_org_conversion_stats()
 *   notion_push_upsale_targets_from_org_info()   ✅ NEW
 *   write_daily_snapshot()
 *   writeSyncLog(step, status, rows_in, rows_out, seconds, error)
 *   lockWrap(fn)  (your shared utility)
 **************************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ping Ops')
    .addItem('Run daily pipeline', 'ui_run_daily_pipeline')
    .addSeparator()
    .addItem('Run only PostHog', 'ui_run_only_posthog')
    .addItem('Run only Stripe', 'ui_run_only_stripe')
    .addItem('Run only Clerk', 'ui_run_only_clerk')
    .addSeparator()
    .addItem('Rebuild canon tables', 'ui_rebuild_canon_tables')
    .addSeparator()
    .addItem('Run ARR refresh', 'ui_run_arr_refresh')
    .addItem('Run Conversion & Onboarding stats', 'ui_run_conversion_onboarding_stats')
    .addItem('Run Conversion audit', 'ui_run_conversion_audit')
    .addSeparator()
    .addItem('Push UpSale targets to Notion', 'ui_push_upsale_targets_to_notion') // ✅ NEW
    .addToUi()
}

/* =========================
 * UI handlers
 * ========================= */

function ui_run_daily_pipeline() {
  return uiRunWrapped_('ui_run_daily_pipeline', () => {
    run_daily_pipeline()
  })
}

function ui_run_only_posthog() {
  return uiRunWrapped_('ui_run_only_posthog', () => {
    runSteps_([
      { name: 'posthog_pull_user_metrics_to_raw', fn: posthog_pull_user_metrics_to_raw },
      { name: 'build_canon_users', fn: build_canon_users },
      { name: 'render_sauron_view', fn: render_sauron_view },
      { name: 'render_ring_view', fn: render_ring_view }
      // { name: 'write_daily_snapshot', fn: write_daily_snapshot }
    ])
  })
}

function ui_run_only_stripe() {
  return uiRunWrapped_('ui_run_only_stripe', () => {
    runSteps_([
      { name: 'stripe_pull_subscriptions_to_raw', fn: stripe_pull_subscriptions_to_raw },
      { name: 'build_canon_orgs', fn: build_canon_orgs },
      { name: 'render_sauron_view', fn: render_sauron_view },
      { name: 'render_ring_view', fn: render_ring_view }
    ])
  })
}

function ui_run_only_clerk() {
  return uiRunWrapped_('ui_run_only_clerk', () => {
    runSteps_([
      { name: 'clerk_pull_users_to_raw', fn: clerk_pull_users_to_raw },
      { name: 'clerk_pull_orgs_to_raw', fn: clerk_pull_orgs_to_raw },
      { name: 'clerk_pull_memberships_to_raw', fn: clerk_pull_memberships_to_raw },
      { name: 'syncClerkUsers (login events)', fn: syncClerkUsers },
      { name: 'build_canon_orgs', fn: build_canon_orgs },
      { name: 'build_canon_users', fn: build_canon_users },
      { name: 'render_sauron_view', fn: render_sauron_view },
      { name: 'render_ring_view', fn: render_ring_view }
    ])
  })
}

function ui_rebuild_canon_tables() {
  return uiRunWrapped_('ui_rebuild_canon_tables', () => {
    runSteps_([
      { name: 'build_canon_orgs', fn: build_canon_orgs },
      { name: 'build_canon_users', fn: build_canon_users },
      { name: 'render_org_info_view', fn: render_org_info_view },
      { name: 'render_sauron_view', fn: render_sauron_view },
      { name: 'render_ring_view', fn: render_ring_view }
    ])
  })
}

function ui_run_arr_refresh() {
  return uiRunWrapped_('ui_run_arr_refresh', () => {
    runSteps_([
      { name: 'render_arr_raw_data_view', fn: render_arr_raw_data_view },
      { name: 'write_arr_snapshot', fn: write_arr_snapshot },
      { name: 'render_arr_waterfall_facts', fn: render_arr_waterfall_facts }
    ])
  })
}

function ui_run_conversion_onboarding_stats() {
  return uiRunWrapped_('ui_run_conversion_onboarding_stats', () => {
    runSteps_([
      { name: 'render_conversion_onboarding_stats', fn: render_conversion_onboarding_stats }
    ])
  })
}

function ui_run_onboarding_stats() {
  return ui_run_conversion_onboarding_stats()
}

function ui_run_conversion_stats() {
  return ui_run_conversion_onboarding_stats()
}

function ui_run_conversion_audit() {
  return uiRunWrapped_('ui_run_conversion_audit', () => {
    runSteps_([
      { name: 'render_org_conversion_audit', fn: render_org_conversion_audit }
    ])
  })
}

/**
 * ✅ NEW: Manual button to push UpSale targets from org_info -> Notion
 */
function ui_push_upsale_targets_to_notion() {
  return uiRunWrapped_('ui_push_upsale_targets_to_notion', () => {
    runSteps_([
      { name: 'notion_push_upsale_targets_from_org_info', fn: notion_push_upsale_targets_from_org_info }
    ])
  })
}

/* =========================
 * Helpers
 * ========================= */

function uiRunWrapped_(name, fn) {
  return lockWrap(name, () => {
    const ss = SpreadsheetApp.getActive()
    ss.toast('Running…', 'Ping Ops', 5)

    const t0 = new Date()
    try {
      fn()
      const seconds = ((new Date()) - t0) / 1000
      writeSyncLog(name, 'ok', '', '', seconds, '')
      ss.toast('Done ✅', 'Ping Ops', 5)
    } catch (err) {
      const seconds = ((new Date()) - t0) / 1000
      const msg = String(err && err.message ? err.message : err)
      writeSyncLog(name, 'error', '', '', seconds, msg)
      ss.toast('Failed ❌ (check Sync Log)', 'Ping Ops', 8)
      throw err
    }
  })
}

function runSteps_(steps) {
  for (const step of steps) {
    const t0 = new Date()
    try {
      const out = step.fn() // may return { rows_in, rows_out }
      const seconds = ((new Date()) - t0) / 1000
      const rowsIn = out && out.rows_in != null ? out.rows_in : ''
      const rowsOut = out && out.rows_out != null ? out.rows_out : ''
      writeSyncLog(step.name, 'ok', rowsIn, rowsOut, seconds, '')
    } catch (err) {
      const seconds = ((new Date()) - t0) / 1000
      const msg = String(err && err.message ? err.message : err)
      writeSyncLog(step.name, 'error', '', '', seconds, msg)
      throw new Error(`${step.name} failed: ${msg}`)
    }
  }
}
