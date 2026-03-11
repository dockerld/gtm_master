/**************************************************************
 * render_all_stats_vs_ring_audit()
 *
 * Compares stage counts from "All the Stats" against "The Ring"
 * and writes a detailed ring subscription list for investigation.
 **************************************************************/

const ALLSTATS_RING_AUDIT_CFG = {
  SHEET_NAME: 'AllStats vs Ring Audit',
  ALL_STATS_SHEET: 'All the Stats',
  RING_SHEET: 'The Ring',
  RING_HEADER_ROW: 3,
  RING_DATA_START_ROW: 4,
  RING_START_COL: 2
}

function render_all_stats_vs_ring_audit() {
  return ALLSTATS_lockWrap_('render_all_stats_vs_ring_audit', () => {
    const t0 = new Date()
    const ss = SpreadsheetApp.getActive()

    const allStatsSheet = ss.getSheetByName(ALLSTATS_RING_AUDIT_CFG.ALL_STATS_SHEET)
    if (!allStatsSheet) throw new Error('Missing sheet: All the Stats')

    const ringSheet = ss.getSheetByName(ALLSTATS_RING_AUDIT_CFG.RING_SHEET)
    if (!ringSheet) throw new Error('Missing sheet: The Ring')

    const stageCounts = ASRA_readStageCountsFromAllStats_(allStatsSheet)
    const ringRows = ASRA_readRingRows_(ringSheet)

    const ringPaidSubs = ringRows.filter(r => r.status === 'Paid')
    const ringPromoSubs = ringRows.filter(r => r.status === 'Promo Trial')

    const ringPaidUnique = ASRA_uniqueOrgCount_(ringPaidSubs)
    const ringPromoUnique = ASRA_uniqueOrgCount_(ringPromoSubs)

    const allStatsPaid = Number(stageCounts.paid || 0)
    const allStatsPromo = Number(stageCounts.promoTrial || 0)
    const allStatsComparable = ASRA_buildAllStatsComparableRows_(ss)

    const paidDiffRows = ASRA_buildSideDiffRows_(
      'Paid',
      ringPaidSubs,
      allStatsComparable.paidRows
    )
    const promoDiffRows = ASRA_buildSideDiffRows_(
      'Promo Trial',
      ringPromoSubs,
      allStatsComparable.promoRows
    )
    const sideDiffRows = [].concat(paidDiffRows, promoDiffRows)

    const out = ASRA_getOrCreateSheet_(ss, ALLSTATS_RING_AUDIT_CFG.SHEET_NAME)
    out.clear()

    let row = 1

    out.getRange(row, 1, 1, 8).merge()
    out.getRange(row, 1)
      .setValue('All Stats vs Ring Audit')
      .setFontWeight('bold')
      .setFontSize(16)
      .setBackground('#F3F4F6')
    row += 1

    const tz = Session.getScriptTimeZone()
    out.getRange(row, 1, 1, 8).merge()
    out.getRange(row, 1)
      .setValue('Generated at: ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss'))
      .setFontSize(10)
      .setFontColor('#6B7280')
    row += 2

    const allStatsHeaders = ['Metric', 'Value']
    const allStatsRows = [
      ['All Stats - Paid (Firms)', allStatsPaid],
      ['All Stats - Promo Trial (Firms)', allStatsPromo]
    ]

    out.getRange(row, 1, 1, allStatsHeaders.length).setValues([allStatsHeaders])
      .setFontWeight('bold')
      .setBackground('#DBEAFE')

    out.getRange(row + 1, 1, allStatsRows.length, allStatsHeaders.length).setValues(allStatsRows)
    out.getRange(row + 1, 2, allStatsRows.length, 1).setNumberFormat('0')
    row += allStatsRows.length + 2

    const summaryHeaders = ['Metric', 'Value']
    const summaryRows = [
      ['Ring - Paid (Subscriptions)', ringPaidSubs.length],
      ['Ring - Promo Trial (Subscriptions)', ringPromoSubs.length],
      ['Ring - Paid (Unique Orgs Derived)', ringPaidUnique],
      ['Ring - Promo Trial (Unique Orgs Derived)', ringPromoUnique],
      ['Diff Paid (All Stats Firms - Ring Unique Orgs)', allStatsPaid - ringPaidUnique],
      ['Diff Promo Trial (All Stats Firms - Ring Unique Orgs)', allStatsPromo - ringPromoUnique]
    ]

    out.getRange(row, 1, 1, summaryHeaders.length).setValues([summaryHeaders])
      .setFontWeight('bold')
      .setBackground('#E5E7EB')

    out.getRange(row + 1, 1, summaryRows.length, summaryHeaders.length).setValues(summaryRows)
    out.getRange(row + 1, 2, summaryRows.length, 1).setNumberFormat('0')
    row += summaryRows.length + 3

    out.getRange(row, 1, 1, 8).merge()
    out.getRange(row, 1)
      .setValue('Which Rows Make Up The Difference')
      .setFontWeight('bold')
      .setBackground('#FEF3C7')
    row += 1

    const sideDiffHeaders = [
      'side',
      'status',
      'org_key',
      'org_name',
      'customer_name',
      'customer_email',
      'ring_rows_for_key',
      'all_stats_rows_for_key'
    ]

    out.getRange(row, 1, 1, sideDiffHeaders.length).setValues([sideDiffHeaders])
      .setFontWeight('bold')
      .setBackground('#F3F4F6')

    const safeSideDiffRows = sideDiffRows.length
      ? sideDiffRows
      : [['(none)', '', '', '', '', '', '', '']]

    out.getRange(row + 1, 1, safeSideDiffRows.length, sideDiffHeaders.length).setValues(safeSideDiffRows)
    out.getRange(row + 1, 7, safeSideDiffRows.length, 2).setNumberFormat('0')
    row += safeSideDiffRows.length + 3

    out.getRange(row, 1, 1, 8).merge()
    out.getRange(row, 1)
      .setValue('Ring Subscription List (Paid + Promo Trial)')
      .setFontWeight('bold')
      .setBackground('#EEF2FF')
    row += 1

    const detailHeaders = [
      'ring_row',
      'status',
      'org_name',
      'customer_name',
      'customer_email',
      'first_payment_at',
      'org_key',
      'duplicate_org_in_status'
    ]

    out.getRange(row, 1, 1, detailHeaders.length).setValues([detailHeaders])
      .setFontWeight('bold')
      .setBackground('#F3F4F6')

    const filtered = ringRows.filter(r => r.status === 'Paid' || r.status === 'Promo Trial')

    const dupCountByStatusOrg = new Map()
    for (const r of filtered) {
      const key = r.status + '|' + r.org_key
      dupCountByStatusOrg.set(key, (dupCountByStatusOrg.get(key) || 0) + 1)
    }

    const detailRows = filtered.map(r => {
      const key = r.status + '|' + r.org_key
      const isDup = (dupCountByStatusOrg.get(key) || 0) > 1 ? 'yes' : ''
      return [
        r.ring_row,
        r.status,
        r.org_name,
        r.customer_name,
        r.customer_email,
        r.first_payment_at || '',
        r.org_key,
        isDup
      ]
    })

    if (detailRows.length) {
      out.getRange(row + 1, 1, detailRows.length, detailHeaders.length).setValues(detailRows)
      out.getRange(row + 1, 1, detailRows.length, 1).setNumberFormat('0')
      out.getRange(row + 1, 6, detailRows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss')
    }

    out.setFrozenRows(1)
    out.autoResizeColumns(1, detailHeaders.length)

    const seconds = (new Date() - t0) / 1000
    ALLSTATS_writeSyncLog_('render_all_stats_vs_ring_audit', 'ok', ringRows.length, detailRows.length, seconds, '')
    return { rows_in: ringRows.length, rows_out: detailRows.length }
  })
}

function ASRA_readStageCountsFromAllStats_(sheet) {
  const vals = sheet.getDataRange().getValues()
  if (!vals.length) return { paid: 0, promoTrial: 0 }

  let sectionRow = -1
  for (let r = 0; r < vals.length; r++) {
    const a = String(vals[r][0] || '').trim()
    if (a === '# of Paying Firms / Seats by Stage') {
      sectionRow = r
      break
    }
  }

  if (sectionRow < 0) return { paid: 0, promoTrial: 0 }

  const out = { paid: 0, promoTrial: 0 }

  // Header expected one row below section label, then stage rows.
  for (let r = sectionRow + 2; r < vals.length; r++) {
    const stage = String(vals[r][0] || '').trim()
    if (!stage) break
    const firms = Number(vals[r][1] || 0)

    if (stage === 'Paid') out.paid = isFinite(firms) ? firms : 0
    if (stage === 'Promo Trial') out.promoTrial = isFinite(firms) ? firms : 0
  }

  return out
}

function ASRA_readRingRows_(sheet) {
  const maxColsFromStart = sheet.getLastColumn() - ALLSTATS_RING_AUDIT_CFG.RING_START_COL + 1
  if (maxColsFromStart <= 0) return []

  const headerRaw = sheet
    .getRange(ALLSTATS_RING_AUDIT_CFG.RING_HEADER_ROW, ALLSTATS_RING_AUDIT_CFG.RING_START_COL, 1, maxColsFromStart)
    .getValues()[0]

  const headerWidth = ASRA_contiguousWidth_(headerRaw)
  if (headerWidth <= 0) return []

  const headers = headerRaw.slice(0, headerWidth).map(h => String(h || '').trim())
  const idx = {}
  headers.forEach((h, i) => { idx[h] = i })

  const lastRow = sheet.getLastRow()
  const numRows = Math.max(0, lastRow - ALLSTATS_RING_AUDIT_CFG.RING_DATA_START_ROW + 1)
  if (!numRows) return []

  const data = sheet
    .getRange(ALLSTATS_RING_AUDIT_CFG.RING_DATA_START_ROW, ALLSTATS_RING_AUDIT_CFG.RING_START_COL, numRows, headerWidth)
    .getValues()

  const out = []
  for (let i = 0; i < data.length; i++) {
    const r = data[i]
    const status = ASRA_str_(r[idx['Status']])
    if (!status) continue

    const orgName = ASRA_str_(r[idx['Org Name']])
    const customerName = ASRA_str_(r[idx['Customer Name']])
    const customerEmail = ASRA_str_(r[idx['Customer Email']])
    const firstPaymentAt = (idx['First Payment At'] != null) ? r[idx['First Payment At']] : ''

    out.push({
      ring_row: ALLSTATS_RING_AUDIT_CFG.RING_DATA_START_ROW + i,
      status,
      org_name: orgName,
      customer_name: customerName,
      customer_email: customerEmail,
      first_payment_at: firstPaymentAt,
      org_key: ASRA_orgKey_(orgName, customerEmail, customerName)
    })
  }

  return out
}

function ASRA_uniqueOrgCount_(rows) {
  const s = new Set()
  ;(rows || []).forEach(r => s.add(r.org_key))
  return s.size
}

function ASRA_buildSideDiffRows_(statusLabel, ringRows, allStatsRows) {
  const ringByKey = ASRA_groupByOrgKey_(ringRows || [])
  const allByKey = ASRA_groupByOrgKey_(allStatsRows || [])

  const ringKeys = new Set(Array.from(ringByKey.keys()))
  const allKeys = new Set(Array.from(allByKey.keys()))

  const out = []

  for (const key of ringKeys) {
    if (allKeys.has(key)) continue
    const sampleRing = (ringByKey.get(key) || [])[0] || {}
    out.push([
      'Ring only',
      statusLabel,
      key,
      sampleRing.org_name || '',
      sampleRing.customer_name || '',
      sampleRing.customer_email || '',
      (ringByKey.get(key) || []).length,
      0
    ])
  }

  for (const key of allKeys) {
    if (ringKeys.has(key)) continue
    const sampleAll = (allByKey.get(key) || [])[0] || {}
    out.push([
      'All Stats only',
      statusLabel,
      key,
      sampleAll.org_name || '',
      sampleAll.customer_name || '',
      sampleAll.customer_email || '',
      0,
      (allByKey.get(key) || []).length
    ])
  }

  out.sort((a, b) => {
    if (a[1] !== b[1]) return String(a[1]).localeCompare(String(b[1]))
    if (a[0] !== b[0]) return String(a[0]).localeCompare(String(b[0]))
    return String(a[2]).localeCompare(String(b[2]))
  })

  return out
}

function ASRA_groupByOrgKey_(rows) {
  const map = new Map()
  for (const r of (rows || [])) {
    const key = ASRA_str_(r && r.org_key)
    if (!key) continue
    if (!map.has(key)) map.set(key, [])
    map.get(key).push(r)
  }
  return map
}

function ASRA_buildAllStatsComparableRows_(ss) {
  const inputs = (typeof ALL_STATS_CFG !== 'undefined' && ALL_STATS_CFG && ALL_STATS_CFG.INPUTS)
    ? ALL_STATS_CFG.INPUTS
    : {
        STRIPE_SUBS: 'raw_stripe_subscriptions',
        ORG_SUBSCRIPTIONS: 'org_subscription_info',
        POSTHOG_PROMO_REDEMPTIONS: 'promo_redemptions',
        MANUAL_CHANGES: 'Manual Stripe Changes',
        CANON_ORGS: 'canon_orgs',
        CLERK_USERS: 'raw_clerk_users',
        CLERK_MEMBERSHIPS: 'raw_clerk_memberships',
        CLERK_ORGS: 'raw_clerk_orgs',
        POSTHOG_USERS: 'raw_posthog_user_metrics'
      }

  const shStripe = ss.getSheetByName(inputs.STRIPE_SUBS)
  if (!shStripe) throw new Error('Missing input sheet: ' + inputs.STRIPE_SUBS)

  const shOrgSubscriptions = ss.getSheetByName(inputs.ORG_SUBSCRIPTIONS)
  if (!shOrgSubscriptions) throw new Error('Missing input sheet: ' + inputs.ORG_SUBSCRIPTIONS)
  const shCanonOrgs = ss.getSheetByName(inputs.CANON_ORGS)
  if (!shCanonOrgs) throw new Error('Missing input sheet: ' + inputs.CANON_ORGS)
  const shManual = ss.getSheetByName(inputs.MANUAL_CHANGES)
  const shPromoRedemptions = ss.getSheetByName(inputs.POSTHOG_PROMO_REDEMPTIONS)
  const shUsers = ss.getSheetByName(inputs.CLERK_USERS)
  const shMems = ss.getSheetByName(inputs.CLERK_MEMBERSHIPS)
  const shOrgs = ss.getSheetByName(inputs.CLERK_ORGS)
  const shPosthog = ss.getSheetByName(inputs.POSTHOG_USERS)

  const stripeRows = ALLSTATS_readSheetObjects_(shStripe, 1)
  const orgSubscriptions = ALLSTATS_loadOrgSubscriptions_(shOrgSubscriptions)
  const orgSubsIndex = ALLSTATS_buildOrgSubscriptionsIndex_(orgSubscriptions)
  const manualBySubId = ALLSTATS_buildManualStripeChangesBySubId_(shManual)
  const promoRedemptions = ALLSTATS_loadPromoRedemptions_(shPromoRedemptions)
  const canonOrgs = ALLSTATS_readSheetObjects_(shCanonOrgs, 1)

  const clerkUsers = shUsers ? ALLSTATS_readSheetObjects_(shUsers, 1) : []
  const clerkMems = shMems ? ALLSTATS_readSheetObjects_(shMems, 1) : []
  const clerkOrgs = shOrgs ? ALLSTATS_readSheetObjects_(shOrgs, 1) : []
  const posthogUsers = shPosthog ? ALLSTATS_readSheetObjects_(shPosthog, 1) : []
  const indexes = ALLSTATS_buildIndexes_(clerkUsers, clerkMems, clerkOrgs, posthogUsers, canonOrgs)

  const orgAggByKey = new Map()
  for (const row of stripeRows) {
    const sub = ALLSTATS_normalizeSubscription_(row, manualBySubId, indexes, orgSubsIndex)
    if (!sub.include) continue
    const key = ALLSTATS_orgKey_(sub)
    if (!orgAggByKey.has(key)) orgAggByKey.set(key, ALLSTATS_newOrgAgg_(sub, key))
    ALLSTATS_addSubToOrgAgg_(orgAggByKey.get(key), sub)
  }

  const orgs = Array.from(orgAggByKey.values()).map(ALLSTATS_finalizeOrgAgg_)
  const promoByOrgId = ALLSTATS_buildPromoRedemptionByOrgId_(promoRedemptions)
  ALLSTATS_applyPromoEligibilityFromRedemptions_(orgs, promoByOrgId)

  const toComparable = o => ({
    org_key: ASRA_orgKey_(o.org_name, o.customer_email, o.customer_name),
    org_name: o.org_name || '',
    customer_name: o.customer_name || '',
    customer_email: o.customer_email || ''
  })

  return {
    paidRows: orgs.filter(o => o.has_paid).map(toComparable),
    promoRows: orgs.filter(o => !o.has_paid && o.has_promo).map(toComparable)
  }
}

function ASRA_orgKey_(orgName, customerEmail, customerName) {
  const org = ASRA_str_(orgName).toLowerCase()
  if (org) return 'org:' + org
  const email = ASRA_str_(customerEmail).toLowerCase()
  if (email) return 'email:' + email
  const name = ASRA_str_(customerName).toLowerCase()
  if (name) return 'name:' + name
  return 'unknown'
}

function ASRA_contiguousWidth_(headerRowArray) {
  let w = 0
  for (let i = 0; i < headerRowArray.length; i++) {
    const v = String(headerRowArray[i] || '').trim()
    if (!v) break
    w++
  }
  return w
}

function ASRA_str_(v) {
  if (v === null || v === undefined) return ''
  return String(v).trim()
}

function ASRA_getOrCreateSheet_(ss, name) {
  if (typeof getOrCreateSheetCompat_ === 'function') {
    try { return getOrCreateSheetCompat_(ss, name) } catch (e) {}
  }
  if (typeof getOrCreateSheet === 'function') {
    try { return getOrCreateSheet(ss, name) } catch (e) {}
    try { return getOrCreateSheet(name) } catch (e) {}
  }
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}
