/**************************************************************
 * Shared Utilities (one file)
 * Project: Ping "Sauron" Sheet Pipeline
 *
 * Drop this in a single Apps Script file, e.g. "00_utils.gs"
 *
 * Notes:
 * - Header-based: readHeaderMap() builds {headerLower: colIndex1Based}
 * - Email key: normalizeEmail() -> lower(trim(email))
 * - Batch writes: batchSetValues() writes in chunks to avoid limits
 * - Upserts: buildIndexByKey() builds {key: rowIndex0BasedInArray}
 * - Locking: lockWrap() prevents overlapping runs
 **************************************************************/

const UTIL_CFG = {
  DEFAULT_BATCH_ROWS: 5000,
  LOCK_TIMEOUT_MS: 5 * 60 * 1000, // 5 minutes
};

/**
 * Get or create a sheet by name.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet(ss, name) {
  if (!ss) throw new Error('getOrCreateSheet: ss is required');
  if (!name) throw new Error('getOrCreateSheet: name is required');
  const existing = ss.getSheetByName(name);
  return existing || ss.insertSheet(name);
}

/**
 * Read header row and return a map of header -> 1-based col index.
 * Keys are normalized to lowercase trimmed strings.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} headerRow
 * @returns {{ map: Object<string, number>, headers: string[] }}
 */
function readHeaderMap(sheet, headerRow) {
  if (!sheet) throw new Error('readHeaderMap: sheet is required');
  if (!headerRow || headerRow < 1) throw new Error('readHeaderMap: headerRow must be >= 1');

  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return { map: {}, headers: [] };

  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || '').trim());

  const map = {};
  headers.forEach((h, i) => {
    const key = String(h || '').trim().toLowerCase();
    if (!key) return;
    // If duplicate headers exist, keep the first one (predictable)
    if (map[key] == null) map[key] = i + 1;
  });

  return { map, headers };
}

/**
 * Normalize an email address into a stable key.
 * @param {string} email
 * @returns {string} email_key
 */
function normalizeEmail(email) {
  if (email == null) return '';
  return String(email).trim().toLowerCase();
}

/**
 * Batch write a 2D array to a sheet in chunks.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} startRow 1-based
 * @param {number} startCol 1-based
 * @param {any[][]} values 2D array
 * @param {number=} batchRows optional override
 */
function batchSetValues(sheet, startRow, startCol, values, batchRows) {
  if (!sheet) throw new Error('batchSetValues: sheet is required');
  if (!startRow || startRow < 1) throw new Error('batchSetValues: startRow must be >= 1');
  if (!startCol || startCol < 1) throw new Error('batchSetValues: startCol must be >= 1');
  if (!Array.isArray(values) || values.length === 0) return;

  const rowsPerBatch = batchRows || UTIL_CFG.DEFAULT_BATCH_ROWS;
  const numCols = values[0].length;

  for (let i = 0; i < values.length; i += rowsPerBatch) {
    const chunk = values.slice(i, i + rowsPerBatch);
    sheet.getRange(startRow + i, startCol, chunk.length, numCols).setValues(chunk);
  }
}

/**
 * Build an index map from a 2D array of rows keyed by a specific column index (0-based).
 *
 * @param {any[][]} rows
 * @param {number} keyColIndex0 0-based index within each row
 * @param {Object=} opts
 * @param {boolean=} opts.lowercase
 * @param {boolean=} opts.trim
 * @returns {Object<string, number>} key -> rowIndex0Based
 */
function buildIndexByKey(rows, keyColIndex0, opts) {
  if (!Array.isArray(rows)) throw new Error('buildIndexByKey: rows must be an array');
  if (keyColIndex0 == null || keyColIndex0 < 0) throw new Error('buildIndexByKey: keyColIndex0 must be >= 0');

  const o = Object.assign({ lowercase: true, trim: true }, opts || {});
  const index = {};

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || [];
    let key = r[keyColIndex0];

    if (key == null) continue;
    key = String(key);

    if (o.trim) key = key.trim();
    if (o.lowercase) key = key.toLowerCase();
    if (!key) continue;

    // Keep first occurrence (stable). Later duplicates are ignored.
    if (index[key] == null) index[key] = i;
  }

  return index;
}

/**
 * Append a row to sync_log (creates the tab + header if needed).
 *
 * Columns:
 * timestamp | step | status | rows_in | rows_out | seconds | error
 *
 * @param {string} step
 * @param {string} status e.g. "ok" | "error"
 * @param {number=} rowsIn
 * @param {number=} rowsOut
 * @param {number=} seconds
 * @param {string=} errorMsg
 */
// No-op: sync_log sheet removed. All callers still reference this function
// so we keep the signature to avoid runtime errors.
function writeSyncLog() {}

/**
 * Wrap a function call in a document lock to prevent overlapping runs.
 *
 * Usage:
 *   lockWrap('daily_pipeline', () => run_daily_pipeline_impl_());
 *
 * @param {string} lockName used only for log labeling
 * @param {Function} fn
 * @param {Object=} opts
 * @param {number=} opts.timeoutMs
 * @returns {any} return value of fn
 */
function lockWrap(lockName, fn, opts) {
  if (typeof fn !== 'function') throw new Error('lockWrap: fn must be a function');

  const timeoutMs = (opts && opts.timeoutMs) || UTIL_CFG.LOCK_TIMEOUT_MS;
  const lock = LockService.getDocumentLock();

  const t0 = new Date();
  const got = lock.tryLock(timeoutMs);

  if (!got) {
    const msg = `lockWrap: could not acquire lock (${lockName || 'job'}) within ${timeoutMs}ms`;
    writeSyncLog(lockName || 'job', 'error', '', '', (new Date() - t0) / 1000, msg);
    throw new Error(msg);
  }

  try {
    return fn();
  } catch (err) {
    const msg = String(err && err.message ? err.message : err);
    writeSyncLog(lockName || 'job', 'error', '', '', (new Date() - t0) / 1000, msg);
    throw err;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**************************************************************
 * Shared Conversion Org List
 *
 * Builds a filtered, enriched org list from canon_orgs + Stripe
 * for use in conversion metrics across all sheets.
 *
 * Returns array of objects:
 *   { org_id, org_name, owner_email, org_created_at, trial_ends_at,
 *     first_payment_at, is_paying, stripe_subscription_ids }
 *
 * Excludes:
 *   - Org name contains "ping" or "test" (case-insensitive)
 *   - Owner email ends with @pingassistant.com
 *   - All subs in Manual Stripe Changes with exclude reason
 **************************************************************/
function buildConversionOrgList_(ss) {
  // 1) Read canon_orgs
  const shCanon = ss.getSheetByName('canon_orgs')
  if (!shCanon) return []
  const canonOrgs = CONVUTIL_readSheetObjects_(shCanon, 1)

  // 2) Read raw_stripe_subscriptions for first_payment_at
  const shStripe = ss.getSheetByName('raw_stripe_subscriptions')
  const stripeRows = shStripe ? CONVUTIL_readSheetObjects_(shStripe, 1) : []
  const stripeBySubId = new Map()
  for (const r of stripeRows) {
    const subId = CONVUTIL_str_(r.stripe_subscription_id || r.subscription_id || r.id)
    if (subId) stripeBySubId.set(subId, r)
  }

  // 3) Build excluded sub IDs from Manual Stripe Changes
  const excludedSubIds = CONVUTIL_buildExcludedSubIds_(ss)

  // 4) Enrich and filter each canon org
  const result = []
  for (const org of canonOrgs) {
    // Exclude internal/test orgs by name
    const orgName = CONVUTIL_str_(org.org_name || org.org_slug).toLowerCase()
    if (orgName.includes('ping') || orgName.includes('test')) continue

    // Exclude by owner email
    const ownerEmail = CONVUTIL_str_(org.owner_email).toLowerCase()
    if (ownerEmail.endsWith('@pingassistant.com')) continue

    // Parse subscription IDs
    const subIds = CONVUTIL_csvList_(org.stripe_subscription_ids)
    const nonExcludedSubs = subIds.filter(id => !excludedSubIds.has(id))

    // If org has subs but ALL are excluded, skip
    if (subIds.length > 0 && nonExcludedSubs.length === 0) continue

    // Find earliest first_payment_at across non-excluded subs
    let firstPaymentAt = null
    for (const subId of nonExcludedSubs) {
      const stripe = stripeBySubId.get(subId)
      if (!stripe) continue
      const fp = CONVUTIL_parseDate_(stripe.first_payment_at)
      if (fp && (!firstPaymentAt || fp < firstPaymentAt)) {
        firstPaymentAt = fp
      }
    }

    // trial_ends_at from canon_orgs
    const trialEndsAt = CONVUTIL_parseDate_(org.trial_ends_at)

    // Determine if trial is resolved (ended) or still active
    const now = new Date()
    const status = CONVUTIL_str_(org.active_subscription_status || org.org_status).toLowerCase()
    const isCurrentlyTrialing = (status === 'trialing' || status === 'active') && !firstPaymentAt
    const trialResolved = !isCurrentlyTrialing

    // Promo / trial extension: covers in-app promo, Stripe promo, AND manual Stripe extension
    const hasPromoCode = !!(CONVUTIL_str_(org.combined_promo_codes) ||
      CONVUTIL_str_(org.promo_code) ||
      CONVUTIL_str_(org.stripe_promo_codes) ||
      CONVUTIL_str_(org.app_promo_codes))
    const trialExtended = CONVUTIL_str_(org.trial_extended).toLowerCase() === 'true'
    const trialExtensionSource = CONVUTIL_str_(org.trial_extension_source)
    const hasPromo = hasPromoCode || trialExtended

    result.push({
      org_id: CONVUTIL_str_(org.app_org_id || org.org_id),
      org_name: CONVUTIL_str_(org.org_name),
      owner_email: CONVUTIL_str_(org.owner_email),
      org_created_at: CONVUTIL_parseDate_(org.org_created_at),
      trial_ends_at: trialEndsAt,
      first_payment_at: firstPaymentAt,
      is_paying: CONVUTIL_str_(org.is_paying),
      is_currently_trialing: isCurrentlyTrialing,
      trial_resolved: trialResolved,
      ever_had_sub: nonExcludedSubs.length > 0,
      has_promo: hasPromo,
      trial_extended: trialExtended,
      trial_extension_source: trialExtensionSource,
      promo_codes: CONVUTIL_str_(org.combined_promo_codes) || CONVUTIL_str_(org.promo_code) || '',
      org_status: CONVUTIL_str_(org.org_status)
    })
  }
  return result
}

// ── Conversion util helpers (prefixed to avoid collisions) ──

function CONVUTIL_readSheetObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1 || lastCol < 1) return []

  const header = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || '').trim().toLowerCase().replace(/\s+/g, '_'))
  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()

  return data.map(r => {
    const obj = {}
    header.forEach((h, i) => { if (h) obj[h] = r[i] })
    return obj
  })
}

function CONVUTIL_str_(v) {
  if (v === null || v === undefined) return ''
  return String(v).trim()
}

function CONVUTIL_csvList_(v) {
  const s = CONVUTIL_str_(v)
  if (!s) return []
  return s.split(',').map(x => x.trim()).filter(Boolean)
}

function CONVUTIL_parseDate_(v) {
  if (!v) return null
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v
  const s = CONVUTIL_str_(v)
  if (!s) return null
  const d = new Date(s)
  return isNaN(d.getTime()) ? null : d
}

function CONVUTIL_toMonthKey_(d) {
  if (!d) return ''
  const date = (d instanceof Date) ? d : CONVUTIL_parseDate_(d)
  if (!date) return ''
  const y = date.getFullYear()
  const m = String(date.getMonth() + 1).padStart(2, '0')
  return y + '-' + m
}

function CONVUTIL_buildExcludedSubIds_(ss) {
  const out = new Set()
  const sh = ss.getSheetByName('Manual Stripe Changes')
  if (!sh) return out
  const rows = CONVUTIL_readSheetObjects_(sh, 1)
  const EXCLUDE_REASONS = new Set(['internal', 'partner', 'free subscription'])
  for (const r of rows) {
    const reason = CONVUTIL_str_(r.exclude_reason).toLowerCase()
    if (!EXCLUDE_REASONS.has(reason)) continue
    const subId = CONVUTIL_str_(r.subscription_id || r.stripe_subscription_id || r.subscription)
    if (subId) out.add(subId)
  }
  return out
}