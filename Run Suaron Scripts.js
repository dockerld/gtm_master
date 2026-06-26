/**************************************************************
 * run_daily_pipeline_part1() — Data pulls + canon builds + Sauron
 * run_daily_pipeline_part2() — ARR + analytics + CSM Commission
 *
 * Split into two to stay under the Apps Script 30-min limit.
 * Trigger schedule: Part 1 at 2am, Part 2 at 3am.
 *
 * run_daily_pipeline() is kept as a backstop for manual full runs.
 *
 * Notes:
 * - Each step is wrapped so one failure does not stop later steps
 * - Whole pipeline is protected by LockService to avoid overlaps
 * - Sends email alert if any step fails
 * - Per-step timing is written to the "pipeline_log" sheet
 **************************************************************/

const PIPELINE_PART1_STEPS_ = [
  { name: 'clerk_pull_users_to_raw',               fn: () => clerk_pull_users_to_raw() },
  { name: 'clerk_pull_orgs_to_raw',                fn: () => clerk_pull_orgs_to_raw() },
  { name: 'clerk_pull_memberships_to_raw',         fn: () => clerk_pull_memberships_to_raw() },
  { name: 'syncClerkUsers',                        fn: () => syncClerkUsers() },

  { name: 'stripe_pull_subscriptions_to_raw',      fn: () => stripe_pull_subscriptions_to_raw() },

  { name: 'posthog_pull_user_metrics_to_raw',      fn: () => posthog_pull_user_metrics_to_raw() },
  { name: 'posthog_pull_org_subscriptions_to_raw', fn: () => posthog_pull_org_subscriptions_to_raw() },
  { name: 'posthog_pull_orgs_to_raw',              fn: () => posthog_pull_orgs_to_raw() },
  { name: 'posthog_pull_promo_redemptions_to_raw', fn: () => posthog_pull_promo_redemptions_to_raw() },

  { name: 'render_org_subscription_info',          fn: () => render_org_subscription_info() },
  { name: 'build_canon_orgs',                      fn: () => build_canon_orgs() },
  { name: 'build_canon_users',                     fn: () => build_canon_users() },
  { name: 'render_sauron_view',                    fn: () => render_sauron_view() }
]

const PIPELINE_PART2_STEPS_ = [
  { name: 'render_arr_raw_data_view',              fn: () => render_arr_raw_data_view() },
  { name: 'write_arr_snapshot',                    fn: () => write_arr_snapshot() },
  { name: 'render_arr_waterfall_facts',            fn: () => render_arr_waterfall_facts() },
  { name: 'render_paying_users_snapshot',          fn: () => render_paying_users_snapshot() },

  { name: 'render_ring_view',                      fn: () => render_ring_view() },
  { name: 'render_all_stats_view',                 fn: () => render_all_stats_view() },

  { name: 'render_conversion_onboarding_stats',    fn: () => render_conversion_onboarding_stats() },
  { name: 'render_csm_commission_report',          fn: () => render_csm_commission_report() }
]

function run_daily_pipeline_part1() {
  return runPipeline_('run_daily_pipeline_part1', PIPELINE_PART1_STEPS_)
}

function run_daily_pipeline_part2() {
  return runPipeline_('run_daily_pipeline_part2', PIPELINE_PART2_STEPS_)
}

// Backstop: run everything in one shot (may exceed 30-min limit).
// Kept for manual testing only — the daily triggers should use part1/part2.
function run_daily_pipeline() {
  return runPipeline_('run_daily_pipeline', PIPELINE_PART1_STEPS_.concat(PIPELINE_PART2_STEPS_))
}

function runPipeline_(pipelineName, steps) {
  return lockWrapCompat_(pipelineName, () => {
    const pipelineStart = new Date()
    const results = []

    for (const step of steps) {
      const t0 = new Date()
      const res = runStepSafe_(step.name, step.fn, t0)
      results.push(res)
      writePipelineLog_(pipelineName, res)
    }

    const totalSeconds = ((new Date()) - pipelineStart) / 1000
    const errors = results.filter(r => r.status === 'error')

    writePipelineLog_(pipelineName, {
      step: '__TOTAL__',
      status: errors.length ? 'error' : 'ok',
      seconds: totalSeconds,
      error: errors.length ? `${errors.length} step(s) failed` : ''
    })

    if (errors.length) {
      sendPipelineErrorEmail_(errors, totalSeconds, pipelineName)
    }

    return { pipeline: pipelineName, total_seconds: totalSeconds, steps: results }
  })
}

function writePipelineLog_(pipelineName, res) {
  try {
    const ss = SpreadsheetApp.getActive()
    let sh = ss.getSheetByName('pipeline_log')
    if (!sh) {
      sh = ss.insertSheet('pipeline_log')
      sh.getRange(1, 1, 1, 6).setValues([['run_at', 'pipeline', 'step', 'status', 'seconds', 'error']])
        .setFontWeight('bold').setBackground('#F3F4F6')
    }
    sh.appendRow([new Date(), pipelineName, res.step, res.status, res.seconds || 0, res.error || ''])
  } catch (e) {
    Logger.log('writePipelineLog_ failed: ' + e)
  }
}

function runStepSafe_(name, fn, t0) {
  try {
    var out = fn()
    var seconds = ((new Date()) - t0) / 1000
    var rowsIn = out && out.rows_in != null ? out.rows_in : ''
    var rowsOut = out && out.rows_out != null ? out.rows_out : ''
    return { step: name, status: 'ok', seconds: seconds, rows_in: rowsIn, rows_out: rowsOut }
  } catch (err) {
    var seconds2 = ((new Date()) - t0) / 1000
    var msg = String(err && err.message ? err.message : err)
    return { step: name, status: 'error', seconds: seconds2, error: msg }
  }
}

/* =========================
 * Compatibility wrapper (LockService)
 * ========================= */

function lockWrap(lockNameOrFn, maybeFn) {
  let lockName = 'lockWrap'
  let fn = lockNameOrFn

  // Support: lockWrap('name', fn)
  if (typeof lockNameOrFn === 'string') {
    lockName = lockNameOrFn
    fn = maybeFn
  }

  // Support: lockWrap(fn)
  if (typeof fn !== 'function') {
    throw new Error('lockWrap: fn must be a function')
  }

  const lock = LockService.getScriptLock()
  const ok = lock.tryLock(300000) // 5 minutes
  if (!ok) throw new Error(`Could not acquire lock: ${lockName}`)

  try {
    return fn()
  } finally {
    lock.releaseLock()
  }
}

/* =========================
 * Small helper
 * ========================= */

function safeJson_(obj) {
  try {
    return JSON.stringify(obj)
  } catch (e) {
    return String(obj)
  }
}

/* =========================
 * Pipeline error email
 * ========================= */

const PIPELINE_ALERT_RECIPIENTS = [
  'docker@pingassistant.com'
]

function sendPipelineErrorEmail_(errors, totalSeconds, pipelineName) {
  try {
    const tz = Session.getScriptTimeZone()
    const now = Utilities.formatDate(new Date(), tz, 'MMM dd, yyyy h:mm a')
    const label = pipelineName || 'pipeline'

    const rows = errors.map(e =>
      '<tr>' +
        '<td style="padding:8px 12px;border-bottom:1px solid #E2E8F0;font-weight:bold;color:#DC2626">' + (e.step || '?') + '</td>' +
        '<td style="padding:8px 12px;border-bottom:1px solid #E2E8F0">' + String(e.seconds || 0).substring(0, 6) + 's</td>' +
        '<td style="padding:8px 12px;border-bottom:1px solid #E2E8F0;color:#64748B;font-size:13px">' + (e.error || 'unknown') + '</td>' +
      '</tr>'
    ).join('')

    const html =
      '<div style="font-family:sans-serif;max-width:600px">' +
        '<h2 style="color:#DC2626;margin-bottom:4px">Pipeline Alert: ' + label + '</h2>' +
        '<p style="color:#64748B;margin-top:0">' + now + ' &mdash; ' + errors.length + ' step(s) failed in ' + Math.round(totalSeconds) + 's total</p>' +
        '<table style="border-collapse:collapse;width:100%">' +
          '<tr style="background:#F8FAFC">' +
            '<th style="padding:8px 12px;text-align:left;border-bottom:2px solid #CBD5E1">Step</th>' +
            '<th style="padding:8px 12px;text-align:left;border-bottom:2px solid #CBD5E1">Time</th>' +
            '<th style="padding:8px 12px;text-align:left;border-bottom:2px solid #CBD5E1">Error</th>' +
          '</tr>' +
          rows +
        '</table>' +
      '</div>'

    GmailApp.sendEmail(
      PIPELINE_ALERT_RECIPIENTS.join(','),
      'Pipeline Alert: ' + label + ' — ' + errors.length + ' step(s) failed',
      errors.map(e => e.step + ': ' + (e.error || 'unknown')).join('\n'),
      { htmlBody: html }
    )
  } catch (e) {
    Logger.log('Failed to send pipeline error email: ' + e)
  }
}
