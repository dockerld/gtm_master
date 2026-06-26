/**************************************************************
 * Custom Menu (Triggers / UI)
 *
 * Adds a menu to your spreadsheet:
 *  - Run daily pipeline
 *  - Run only PostHog
 *  - Run only Stripe
 *  - Run The Ring only
 *  - Rebuild canon tables
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
 *   writeSyncLog(step, status, rows_in, rows_out, seconds, error)
 *   lockWrap(fn)  (your shared utility)
 **************************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ping Ops')
    .addItem('Run daily pipeline (full)', 'ui_run_daily_pipeline')
    .addItem('Run pipeline: Part 1 (pulls + canon)', 'ui_run_daily_pipeline_part1')
    .addItem('Run pipeline: Part 2 (ARR + analytics)', 'ui_run_daily_pipeline_part2')
    .addSeparator()
    .addItem('Run only PostHog', 'ui_run_only_posthog')
    .addItem('Run only Stripe', 'ui_run_only_stripe')
    .addItem('Run The Ring only', 'ui_run_only_ring')
    .addItem('Render All the Stats', 'ui_render_all_stats')
    .addItem('Publish The Good Stuff', 'ui_publish_the_good_stuff')
    .addSeparator()
    .addItem('Rebuild canon tables', 'ui_rebuild_canon_tables')
    .addSeparator()
    .addItem('Run ARR refresh', 'ui_run_arr_refresh')
    .addItem('Generate CSM Commission Report', 'ui_generate_csm_commission_report')
    .addItem('Render Weekly SS Report', 'ui_render_weekly_ss_report')
    .addSeparator()
    .addItem('TEST: Sauron from PostHog query', 'ui_render_sauron_query_test')
    .addItem('TEST: Org Sub Info from PostHog query', 'ui_render_org_sub_info_query_test')
    .addItem('TEST: Canon Orgs from PostHog query', 'ui_render_canon_orgs_query_test')
    .addItem('TEST: ARR Raw Data from PostHog query', 'ui_render_arr_raw_data_query_test')
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

function ui_run_daily_pipeline_part1() {
  return uiRunWrapped_('ui_run_daily_pipeline_part1', () => {
    run_daily_pipeline_part1()
  })
}

function ui_run_daily_pipeline_part2() {
  return uiRunWrapped_('ui_run_daily_pipeline_part2', () => {
    run_daily_pipeline_part2()
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

function ui_run_only_ring() {
  return uiRunWrapped_('ui_run_only_ring', () => {
    runSteps_([
      { name: 'stripe_pull_subscriptions_to_raw', fn: stripe_pull_subscriptions_to_raw },
      { name: 'render_org_subscription_info',     fn: render_org_subscription_info },
      { name: 'render_ring_view',                 fn: render_ring_view }
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

function ui_publish_the_good_stuff() {
  return uiRunWrapped_('ui_publish_the_good_stuff', () => {
    runSteps_([
      { name: 'publish_the_good_stuff', fn: publish_the_good_stuff }
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
      { name: 'write_arr_snapshot', fn: write_arr_snapshot },
      { name: 'render_arr_waterfall_facts', fn: render_arr_waterfall_facts }
    ])
  })
}

function ui_generate_csm_commission_report() {
  return uiRunWrapped_('ui_generate_csm_commission_report', () => {
    runSteps_([
      { name: 'render_csm_commission_report', fn: render_csm_commission_report }
    ])
  })
}

function ui_render_weekly_ss_report() {
  return uiRunWrapped_('ui_render_weekly_ss_report', () => {
    runSteps_([
      { name: 'render_weekly_ss_report', fn: render_weekly_ss_report }
    ])
  })
}

function ui_render_sauron_query_test() {
  return uiRunWrapped_('ui_render_sauron_query_test', () => {
    runSteps_([
      { name: 'render_sauron_query_test', fn: render_sauron_query_test }
    ])
  })
}

function ui_render_org_sub_info_query_test() {
  return uiRunWrapped_('ui_render_org_sub_info_query_test', () => {
    runSteps_([
      { name: 'render_org_sub_info_query_test', fn: render_org_sub_info_query_test }
    ])
  })
}

function ui_render_canon_orgs_query_test() {
  return uiRunWrapped_('ui_render_canon_orgs_query_test', () => {
    runSteps_([
      { name: 'render_canon_orgs_query_test', fn: render_canon_orgs_query_test }
    ])
  })
}

function ui_render_arr_raw_data_query_test() {
  return uiRunWrapped_('ui_render_arr_raw_data_query_test', () => {
    runSteps_([
      { name: 'render_arr_raw_data_query_test', fn: render_arr_raw_data_query_test }
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
      ss.toast('Failed ❌', 'Ping Ops', 8)
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
