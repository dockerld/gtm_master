/**************************************************************
 * run_daily_pipeline()
 *
 * Orchestrates the daily job in the recommended order.
 *
 * Order:
 * 1) Clerk raw pulls + login event sync
 * 2) Stripe raw pull
 * 3) PostHog raw pull
 * 4) build_canon_orgs
 * 5) build_canon_users
 * 6) render_sauron_view
 * 7) writeSyncLog summary
 *
 * Notes:
 * - Each step is wrapped so one failure does not stop later steps
 * - Whole pipeline is protected by LockService to avoid overlaps
 * - Uses writeSyncLog(step, status, rows_in, rows_out, seconds, error)
 * - Compatible with either lockWrap(lockName, fn) OR lockWrap(fn)
 **************************************************************/

function run_daily_pipeline() {
  lockWrapCompat_('run_daily_pipeline', () => {
    const pipelineStart = new Date()
    const results = []

    const steps = [
      { name: 'clerk_pull_users_to_raw',         fn: clerk_pull_users_to_raw },
      { name: 'clerk_pull_orgs_to_raw',          fn: clerk_pull_orgs_to_raw },
      { name: 'clerk_pull_memberships_to_raw',   fn: clerk_pull_memberships_to_raw },

      // Your existing login history job (append-only login_events + clerk_master rollups)
      { name: 'syncClerkUsers',                  fn: syncClerkUsers },

      { name: 'stripe_pull_subscriptions_to_raw', fn: stripe_pull_subscriptions_to_raw },
      { name: 'posthog_pull_user_metrics_to_raw', fn: posthog_pull_user_metrics_to_raw },

      { name: 'build_canon_orgs',                fn: build_canon_orgs },
      { name: 'build_canon_users',               fn: build_canon_users },
      { name: 'render_org_info_view',            fn: render_org_info_view },

      { name: 'render_sauron_view',              fn: render_sauron_view },
      { name: 'render_ring_view', fn: render_ring_view }
    ]

    for (const step of steps) {
      const t0 = new Date()
      const res = runStepSafe_(step.name, step.fn, t0)
      results.push(res)
    }

    // Final pipeline summary row
    const totalSeconds = ((new Date()) - pipelineStart) / 1000
    const errors = results.filter(r => r.status === 'error')

    writeSyncLog(
      'run_daily_pipeline',
      errors.length ? 'error' : 'ok',
      '',
      '',
      totalSeconds,
      errors.length ? safeJson_(errors) : ''
    )

    return { total_seconds: totalSeconds, steps: results }
  })
}

/**
 * Runs one step and writes a sync log row no matter what.
 * If the step function returns an object like { rows_in, rows_out },
 * we’ll capture it. Otherwise we just log blanks.
 *
 * IMPORTANT:
 * - We do NOT double-log if the step itself already calls writeSyncLog internally.
 *   To support both styles, set STEP_SELF_LOGS = true per step below if needed.
 *   (Default here: false, so pipeline logs every step.)
 */
function runStepSafe_(name, fn, t0) {
  // If you have steps that already call writeSyncLog and you don't want duplicates,
  // add them here:
  const STEP_SELF_LOGS = new Set([
    // 'clerk_pull_users_to_raw',
    // 'clerk_pull_orgs_to_raw',
    // 'clerk_pull_memberships_to_raw',
    // 'stripe_pull_subscriptions_to_raw',
    // 'posthog_pull_user_metrics_to_raw',
    // 'build_canon_orgs',
    // 'build_canon_users',
    // 'render_sauron_view',
    // 'syncClerkUsers',
    // 'render_ring_view'
    // 'render_org_info_view
  ])

  try {
    const out = fn() // may return { rows_in, rows_out }
    const seconds = ((new Date()) - t0) / 1000

    const rowsIn = out && out.rows_in != null ? out.rows_in : ''
    const rowsOut = out && out.rows_out != null ? out.rows_out : ''

    if (!STEP_SELF_LOGS.has(name)) {
      writeSyncLog(name, 'ok', rowsIn, rowsOut, seconds, '')
    }

    return { step: name, status: 'ok', seconds, rows_in: rowsIn, rows_out: rowsOut }
  } catch (err) {
    const seconds = ((new Date()) - t0) / 1000
    const msg = String(err && err.message ? err.message : err)

    if (!STEP_SELF_LOGS.has(name)) {
      writeSyncLog(name, 'error', '', '', seconds, msg)
    }

    return { step: name, status: 'error', seconds, error: msg }
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