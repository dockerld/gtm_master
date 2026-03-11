/**************************************************************
 * Custom Menu (Triggers / UI)
 *
 * Adds a menu to your spreadsheet:
 *  - Run daily pipeline
 *  - Run only PostHog
 *  - Run only Stripe
 *  - Run only Clerk
 *  - Run The Ring only
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
 *   posthog_pull_orgs_to_raw()
 *   posthog_pull_org_subscriptions_to_raw()
 *   posthog_pull_promo_redemptions_to_raw()
 *   clerk_pull_users_to_raw()
 *   clerk_pull_orgs_to_raw()
 *   clerk_pull_memberships_to_raw()
 *   syncClerkUsers()
 *   build_canon_orgs()
 *   build_canon_users()
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
    .addItem('Resync Promo Redemptions', 'ui_resync_promo_redemptions')
    .addItem('Run only Stripe', 'ui_run_only_stripe')
    .addItem('Run only Clerk', 'ui_run_only_clerk')
    .addItem('Run The Ring only', 'ui_run_only_ring')
    .addItem('Render Promo Trial Page', 'ui_render_promo_trial_page')
    .addItem('Render Org Subscription Info', 'ui_render_org_subscription_info')
    .addItem('Render All the Stats', 'ui_render_all_stats')
    .addItem('Publish The Good Stuff', 'ui_publish_the_good_stuff')
    .addItem('Send Ring Weekly Test (Docker)', 'ui_send_ring_weekly_test_docker')
    .addSeparator()
    .addItem('Rebuild canon tables', 'ui_rebuild_canon_tables')
    .addSeparator()
    .addItem('Run ARR refresh', 'ui_run_arr_refresh')
    .addItem('Run Conversion & Onboarding stats', 'ui_run_conversion_onboarding_stats')
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

function ui_confirmed_run_daily_pipeline() {
  const ui = SpreadsheetApp.getUi()
  const result = ui.alert(
    'Confirm Resync',
    'Are you sure you want to resync? This will take 5-10 min to resync everything.',
    ui.ButtonSet.YES_NO
  )
  if (result !== ui.Button.YES) return

  return uiRunWrapped_('ui_confirmed_run_daily_pipeline', () => {
    run_daily_pipeline()
  })
}

function ui_run_only_posthog() {
  return uiRunWrapped_('ui_run_only_posthog', () => {
    runSteps_([
      { name: 'posthog_pull_user_metrics_to_raw', fn: posthog_pull_user_metrics_to_raw },
      { name: 'posthog_pull_org_subscriptions_to_raw', fn: posthog_pull_org_subscriptions_to_raw },
      { name: 'posthog_pull_orgs_to_raw', fn: posthog_pull_orgs_to_raw },
      { name: 'posthog_pull_promo_redemptions_to_raw', fn: posthog_pull_promo_redemptions_to_raw },
      { name: 'build_canon_orgs', fn: build_canon_orgs },
      { name: 'build_canon_users', fn: build_canon_users },
      { name: 'render_sauron_view', fn: render_sauron_view },
      { name: 'render_ring_view', fn: render_ring_view }
      // { name: 'write_daily_snapshot', fn: write_daily_snapshot }
    ])
  })
}

function ui_resync_promo_redemptions() {
  return uiRunWrapped_('ui_resync_promo_redemptions', () => {
    runSteps_([
      { name: 'posthog_pull_promo_redemptions_to_raw', fn: posthog_pull_promo_redemptions_to_raw }
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

function ui_run_only_ring() {
  return uiRunWrapped_('ui_run_only_ring', () => {
    runSteps_([
      { name: 'render_ring_view', fn: render_ring_view }
    ])
  })
}

function ui_render_all_stats() {
  return uiRunWrapped_('ui_render_all_stats', () => {
    runSteps_([
      { name: 'render_all_stats_view', fn: render_all_stats_view }
    ])
  })
}

function ui_render_promo_trial_page() {
  return uiRunWrapped_('ui_render_promo_trial_page', () => {
    runSteps_([
      { name: 'render_promo_trial_page', fn: render_promo_trial_page }
    ])
  })
}

function ui_render_org_subscription_info() {
  return uiRunWrapped_('ui_render_org_subscription_info', () => {
    runSteps_([
      { name: 'render_org_subscription_info', fn: render_org_subscription_info }
    ])
  })
}

function ui_publish_the_good_stuff() {
  return uiRunWrapped_('ui_publish_the_good_stuff', () => {
    runSteps_([
      { name: 'publish_the_good_stuff', fn: publish_the_good_stuff }
    ])
  })
}

function ui_send_ring_weekly_test_docker() {
  return uiRunWrapped_('ui_send_ring_weekly_test_docker', () => {
    runSteps_([
      { name: 'send_ring_weekly_email_test_docker', fn: send_ring_weekly_email_test_docker }
    ])
  })
}

function ui_rebuild_canon_tables() {
  return uiRunWrapped_('ui_rebuild_canon_tables', () => {
    runSteps_([
      { name: 'build_canon_orgs', fn: build_canon_orgs },
      { name: 'build_canon_users', fn: build_canon_users },
      { name: 'render_sauron_view', fn: render_sauron_view },
      { name: 'render_ring_view', fn: render_ring_view }
    ])
  })
}

function ui_run_arr_refresh() {
  return uiRunWrapped_('ui_run_arr_refresh', () => {
    runSteps_([
      { name: 'render_arr_raw_data_view', fn: render_arr_raw_data_view },
      { name: 'one_time_migrate_arr_snapshot_to_new_schema', fn: one_time_migrate_arr_snapshot_to_new_schema },
      { name: 'one_time_add_first_payment_cohort_to_arr_snapshot', fn: one_time_add_first_payment_cohort_to_arr_snapshot },
      { name: 'write_arr_snapshot', fn: write_arr_snapshot },
      { name: 'render_arr_waterfall_facts', fn: render_arr_waterfall_facts }
    ])
  })
}

function ui_run_arr_mapping_audit() {
  return uiRunWrapped_('ui_run_arr_mapping_audit', () => {
    runSteps_([
      { name: 'render_arr_subscription_mapping_audit', fn: render_arr_subscription_mapping_audit }
    ])
  })
}

function ui_audit_all_stats_vs_ring() {
  return uiRunWrapped_('ui_audit_all_stats_vs_ring', () => {
    runSteps_([
      { name: 'render_all_stats_vs_ring_audit', fn: render_all_stats_vs_ring_audit }
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
