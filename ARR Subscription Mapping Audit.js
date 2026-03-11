/**************************************************************
 * render_arr_subscription_mapping_audit()
 *
 * Audits active Stripe subscriptions against canonical org linkage.
 *
 * Output sheet: "arr_subscription_mapping_audit"
 * - Summary metrics
 * - Active subscriptions not linked to any org
 * - Orgs with multiple active subscriptions
 **************************************************************/

const ARRMAP_CFG = {
  OUT_SHEET: 'arr_subscription_mapping_audit',
  STRIPE_SHEET: 'raw_stripe_subscriptions',
  CANON_ORGS_SHEET: 'canon_orgs',
  ORG_SUBSCRIPTIONS_SHEET: 'raw_posthog_org_subscriptions',
  MANUAL_CHANGES_SHEET: 'Manual Stripe Changes'
}

function render_arr_subscription_mapping_audit() {
  return ARRMAP_lockWrap_('render_arr_subscription_mapping_audit', () => {
    const t0 = new Date()
    const ss = SpreadsheetApp.getActive()

    const shStripe = ss.getSheetByName(ARRMAP_CFG.STRIPE_SHEET)
    const shCanonOrgs = ss.getSheetByName(ARRMAP_CFG.CANON_ORGS_SHEET)
    const shOrgSubs = ss.getSheetByName(ARRMAP_CFG.ORG_SUBSCRIPTIONS_SHEET)
    const shManual = ss.getSheetByName(ARRMAP_CFG.MANUAL_CHANGES_SHEET)

    if (!shStripe) throw new Error(`Missing sheet: ${ARRMAP_CFG.STRIPE_SHEET}`)
    if (!shCanonOrgs) throw new Error(`Missing sheet: ${ARRMAP_CFG.CANON_ORGS_SHEET}`)

    const stripeRows = ARRMAP_readSheetObjects_(shStripe, 1)
    const canonOrgs = ARRMAP_readSheetObjects_(shCanonOrgs, 1)
    const orgSubs = shOrgSubs ? ARRMAP_readSheetObjects_(shOrgSubs, 1) : []
    const excludedSubIds = ARRMAP_internalExcludedSubIds_(shManual)

    const orgNameById = new Map()
    ;(canonOrgs || []).forEach(o => {
      const appId = ARRMAP_str_(o.app_org_id)
      const clerkId = ARRMAP_str_(o.clerk_org_id || o.org_id)
      const id = appId || clerkId
      if (!id && !clerkId) return
      const name = ARRMAP_str_(o.org_name || o.org_slug || o.org)
      if (id) orgNameById.set(id, name)
      if (clerkId) orgNameById.set(clerkId, name)
    })

    const orgIdsBySubId = ARRMAP_buildOrgIdsBySubIdFromCanon_(canonOrgs)
    ARRMAP_mergeOrgIdsBySubIdFromOrgSubscriptions_(orgIdsBySubId, orgSubs)

    const activeSubs = []
    const unmapped = []
    const activeSubIdsByOrgId = new Map()

    let excludedInternal = 0

    ;(stripeRows || []).forEach(r => {
      const subId = ARRMAP_str_(
        r.stripe_subscription_id || r.subscription_id || r.subscription || r.id
      )
      if (!subId) return

      if (excludedSubIds.has(subId)) {
        excludedInternal += 1
        return
      }

      const status = ARRMAP_str_(r.status).toLowerCase()
      if (status !== 'active') return

      const orgIdSet = new Set(orgIdsBySubId.get(subId) || [])
      const orgIds = Array.from(orgIdSet)
      activeSubs.push({ subId, orgIds, row: r })

      if (!orgIds.length) {
        unmapped.push({
          sub_id: subId,
          customer_email: ARRMAP_str_(r.customer_email || r.email || r.billing_email),
          interval: ARRMAP_str_(r.interval),
          amount: ARRMAP_num_(r.amount),
          amount_yearly: ARRMAP_num_(r.amount_yearly),
          promo_codes: ARRMAP_str_(r.promo_code_all || r.promo_code),
          discount_percent: ARRMAP_str_(r.discount_percent_all || r.discount_percent),
          discount_amount_off: ARRMAP_str_(r.discount_amount_off_all || r.discount_amount_off)
        })
      }

      orgIds.forEach(orgId => {
        if (!activeSubIdsByOrgId.has(orgId)) activeSubIdsByOrgId.set(orgId, new Set())
        activeSubIdsByOrgId.get(orgId).add(subId)
      })
    })

    const multiOrgRows = []
    activeSubIdsByOrgId.forEach((subSet, orgId) => {
      if ((subSet || new Set()).size <= 1) return
      const ids = Array.from(subSet)
      let arrSum = 0
      ids.forEach(id => {
        const row = stripeRows.find(r => {
          const sid = ARRMAP_str_(r.stripe_subscription_id || r.subscription_id || r.subscription || r.id)
          return sid === id
        })
        if (row) arrSum += ARRMAP_listArr_(row)
      })

      multiOrgRows.push({
        org_id: orgId,
        org_name: orgNameById.get(orgId) || '',
        active_subscription_count: ids.length,
        subscription_ids: ids.join(', '),
        list_arr_sum: arrSum
      })
    })
    multiOrgRows.sort((a, b) => b.active_subscription_count - a.active_subscription_count)

    const out = ARRMAP_getOrCreateSheet_(ss, ARRMAP_CFG.OUT_SHEET)
    out.clear()

    const summary = [
      ['metric', 'value'],
      ['active_subscriptions_total', activeSubs.length],
      ['active_subscriptions_internal_excluded', excludedInternal],
      ['active_subscriptions_unmapped_to_org', unmapped.length],
      ['orgs_with_multiple_active_subscriptions', multiOrgRows.length]
    ]
    out.getRange(1, 1, summary.length, 2).setValues(summary)

    const unmappedHeader = [
      'sub_id', 'customer_email', 'interval', 'amount', 'amount_yearly',
      'discount_percent', 'discount_amount_off', 'promo_codes'
    ]
    const unmappedStart = summary.length + 2
    out.getRange(unmappedStart, 1, 1, unmappedHeader.length).setValues([unmappedHeader])
    if (unmapped.length) {
      const rows = unmapped.map(r => [
        r.sub_id,
        r.customer_email,
        r.interval,
        r.amount,
        r.amount_yearly,
        r.discount_percent,
        r.discount_amount_off,
        r.promo_codes
      ])
      out.getRange(unmappedStart + 1, 1, rows.length, unmappedHeader.length).setValues(rows)
    }

    const multiHeader = [
      'org_id', 'org_name', 'active_subscription_count', 'subscription_ids', 'list_arr_sum'
    ]
    const multiStart = unmappedStart + 2 + Math.max(1, unmapped.length)
    out.getRange(multiStart, 1, 1, multiHeader.length).setValues([multiHeader])
    if (multiOrgRows.length) {
      const rows = multiOrgRows.map(r => [
        r.org_id,
        r.org_name,
        r.active_subscription_count,
        r.subscription_ids,
        r.list_arr_sum
      ])
      out.getRange(multiStart + 1, 1, rows.length, multiHeader.length).setValues(rows)
    }

    out.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#F3F3F3')
    out.getRange(unmappedStart, 1, 1, unmappedHeader.length).setFontWeight('bold').setBackground('#F3F3F3')
    out.getRange(multiStart, 1, 1, multiHeader.length).setFontWeight('bold').setBackground('#F3F3F3')
    out.autoResizeColumns(1, Math.max(unmappedHeader.length, multiHeader.length))

    const seconds = (new Date() - t0) / 1000
    if (typeof writeSyncLog === 'function') {
      writeSyncLog('render_arr_subscription_mapping_audit', 'ok', stripeRows.length, unmapped.length + multiOrgRows.length, seconds, '')
    }

    return { rows_in: stripeRows.length, rows_out: unmapped.length + multiOrgRows.length }
  })
}

function ARRMAP_buildOrgIdsBySubIdFromCanon_(canonOrgs) {
  const out = new Map()
  ;(canonOrgs || []).forEach(r => {
    const appOrgId = ARRMAP_str_(r.app_org_id)
    const clerkOrgId = ARRMAP_str_(r.clerk_org_id || r.org_id)
    const orgId = appOrgId || clerkOrgId
    if (!orgId) return
    const subIds = ARRMAP_csvList_(r.stripe_subscription_ids)
    subIds.forEach(subIdRaw => {
      const subId = ARRMAP_str_(subIdRaw)
      if (!subId) return
      if (!out.has(subId)) out.set(subId, new Set())
      out.get(subId).add(orgId)
    })
  })
  return out
}

function ARRMAP_mergeOrgIdsBySubIdFromOrgSubscriptions_(orgIdsBySubId, orgSubsRows) {
  if (!(orgIdsBySubId instanceof Map)) return
  ;(orgSubsRows || []).forEach(r => {
    const subId = ARRMAP_str_(r.stripe_subscription_id || r.subscription_id || r.subscription || r.id)
    const orgId = ARRMAP_str_(r.app_org_id || r.org_id)
    if (!subId || !orgId) return
    const set = orgIdsBySubId.get(subId) || new Set()
    set.add(orgId)
    orgIdsBySubId.set(subId, set)
  })
}

function ARRMAP_csvList_(v) {
  const s = ARRMAP_str_(v)
  if (!s) return []
  return s.split(',').map(x => ARRMAP_str_(x)).filter(Boolean)
}

function ARRMAP_internalExcludedSubIds_(sheet) {
  const out = new Set()
  if (!sheet) return out
  const rows = ARRMAP_readSheetObjects_(sheet, 1)
  ;(rows || []).forEach(r => {
    const reason = ARRMAP_str_(r.exclude_reason).toLowerCase()
    const EXCLUDE_REASONS = new Set(['internal', 'partner', 'free subscription'])
    if (!EXCLUDE_REASONS.has(reason)) return
    const subId =
      ARRMAP_str_(r.subscription_id) ||
      ARRMAP_str_(r.stripe_subscription_id) ||
      ARRMAP_str_(r.subscription)
    if (subId) out.add(subId)
  })
  return out
}

function ARRMAP_listArr_(row) {
  const amountYearly = ARRMAP_num_(row.amount_yearly)
  if (amountYearly > 0) return amountYearly

  const amount = ARRMAP_num_(row.amount)
  const interval = ARRMAP_str_(row.interval).toLowerCase()
  const intervalCount = Math.max(1, ARRMAP_num_(row.interval_count) || 1)
  if (amount <= 0) return 0
  if (interval === 'year' || interval === 'annual' || interval === 'yr') return amount / intervalCount
  if (interval === 'month' || interval === 'mo') return amount * (12 / intervalCount)
  return amount * 12
}

function ARRMAP_readSheetObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1) return []

  const header = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => ARRMAP_str_(h))
  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()

  return data.map(r => {
    const obj = {}
    header.forEach((h, i) => {
      if (!h) return
      obj[ARRMAP_key_(h)] = r[i]
    })
    return obj
  })
}

function ARRMAP_getOrCreateSheet_(ss, name) {
  if (typeof getOrCreateSheetCompat_ === 'function') return getOrCreateSheetCompat_(ss, name)
  if (typeof getOrCreateSheet === 'function') {
    try { return getOrCreateSheet(ss, name) } catch (e) {}
    try { return getOrCreateSheet(name) } catch (e) {}
  }
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function ARRMAP_lockWrap_(name, fn) {
  if (typeof lockWrapCompat_ === 'function') return lockWrapCompat_(name, fn)
  if (typeof lockWrap === 'function') {
    try { return lockWrap(name, fn) } catch (e) { return lockWrap(fn) }
  }
  const lock = LockService.getScriptLock()
  if (!lock.tryLock(300000)) throw new Error(`Could not acquire lock: ${name}`)
  try { return fn() } finally { lock.releaseLock() }
}

function ARRMAP_key_(h) {
  return ARRMAP_str_(h)
    .toLowerCase()
    .replace(/\s+/g, '_')
}

function ARRMAP_str_(v) {
  if (v === null || v === undefined) return ''
  return String(v).trim()
}

function ARRMAP_num_(v) {
  const n = Number(v)
  return isFinite(n) ? n : 0
}

function ARRMAP_normEmail_(v) {
  const s = ARRMAP_str_(v).toLowerCase()
  if (!s) return ''
  return s.replace(/\+[^@]+(?=@)/, '')
}
