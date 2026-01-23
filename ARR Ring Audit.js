/**************************************************************
 * render_arr_ring_audit()
 *
 * Compares ARR totals from:
 * - The Ring (active Stripe subs, excluding 100% forever)
 * - arr_raw_data (org rollup)
 * - arr_snapshot (latest snapshot date, if present)
 *
 * Outputs:
 * - "ARR vs Ring audit" (summary + per-org delta)
 * - "ARR vs Ring audit subs" (per-sub detail)
 **************************************************************/

const ARRRING_CFG = {
  OUT_SHEET: 'ARR vs Ring audit',
  OUT_SUBS_SHEET: 'ARR vs Ring audit subs',

  STRIPE_SHEET: 'raw_stripe_subscriptions',
  USERS_SHEET: 'raw_clerk_users',
  MEMS_SHEET: 'raw_clerk_memberships',
  ORGS_SHEET: 'raw_clerk_orgs',
  ARR_RAW_SHEET: 'arr_raw_data',
  ARR_SNAPSHOT_SHEET: 'arr_snapshot'
}

function render_arr_ring_audit() {
  return ARRRING_lockWrapCompat_('render_arr_ring_audit', () => {
    const t0 = new Date()
    const ss = SpreadsheetApp.getActive()

    const shStripe = ss.getSheetByName(ARRRING_CFG.STRIPE_SHEET)
    const shUsers = ss.getSheetByName(ARRRING_CFG.USERS_SHEET)
    const shMems = ss.getSheetByName(ARRRING_CFG.MEMS_SHEET)
    const shOrgs = ss.getSheetByName(ARRRING_CFG.ORGS_SHEET)
    const shArrRaw = ss.getSheetByName(ARRRING_CFG.ARR_RAW_SHEET)
    const shArrSnap = ss.getSheetByName(ARRRING_CFG.ARR_SNAPSHOT_SHEET)

    if (!shStripe) throw new Error(`Missing sheet: ${ARRRING_CFG.STRIPE_SHEET}`)
    if (!shUsers) throw new Error(`Missing sheet: ${ARRRING_CFG.USERS_SHEET}`)
    if (!shMems) throw new Error(`Missing sheet: ${ARRRING_CFG.MEMS_SHEET}`)
    if (!shOrgs) throw new Error(`Missing sheet: ${ARRRING_CFG.ORGS_SHEET}`)
    if (!shArrRaw) throw new Error(`Missing sheet: ${ARRRING_CFG.ARR_RAW_SHEET}`)

    const subs = ARRRING_readSheetObjects_(shStripe, 1)
    const users = ARRRING_readSheetObjects_(shUsers, 1)
    const mems = ARRRING_readSheetObjects_(shMems, 1)
    const orgs = ARRRING_readSheetObjects_(shOrgs, 1)

    const orgNameById = new Map()
    orgs.forEach(o => {
      const id = ARRRING_str_(o.org_id)
      if (!id) return
      const name = ARRRING_str_(o.org_name || o.org_slug)
      orgNameById.set(id, name)
    })

    const membershipsByOrgId = ARRRING_buildMembershipsByOrgId_(mems)
    const userByEmailKey = ARRRING_buildUsersByEmailKey_(users)
    const subIdsByOrgId = ARRRING_buildSubIdsByOrgId_(membershipsByOrgId, userByEmailKey, users)
    const orgIdsBySubId = ARRRING_buildOrgIdsBySubId_(subIdsByOrgId)

    const arrRawInfo = ARRRING_readArrRaw_(shArrRaw)
    const arrRawByOrgId = arrRawInfo.byOrgId
    const arrRawTotals = arrRawInfo.totals

    const snapInfo = shArrSnap ? ARRRING_readArrSnapshot_(shArrSnap) : null

    const ringSubRows = []
    const ringArrByOrgId = new Map()
    let ringTotalArr = 0
    let ringTotalArrMapped = 0
    let ringUnmappedArr = 0
    let ringActiveSubs = 0
    let ringUnmappedSubs = 0

    subs.forEach(r => {
      const status = ARRRING_str_(r.status).toLowerCase()
      const discountPercent = ARRRING_num_(r.discount_percent)
      const discountDuration = ARRRING_str_(r.discount_duration).toLowerCase()
      const discountDurationMonths = ARRRING_num_(r.discount_duration_months)

      const interval = ARRRING_str_(r.interval).toLowerCase()
      const amount = ARRRING_moneyAmount_(r.amount)
      const amountYearly = ARRRING_num_(r.amount_yearly)

      const subId = ARRRING_str_(
        r.stripe_subscription_id || r.subscription_id || r.subscription || r.id
      )

      const includedInRing = (status === 'active') &&
        !(discountPercent === 100 && discountDuration === 'forever')

      let excludedReason = ''
      if (!includedInRing) {
        if (status !== 'active') excludedReason = 'inactive'
        if (discountPercent === 100 && discountDuration === 'forever') {
          excludedReason = excludedReason ? `${excludedReason},100%_forever` : '100%_forever'
        }
      }

      const ringArr = includedInRing ? ARRRING_computeArr_(amount, interval) : 0
      const arrRawLike = ARRRING_computeArrRawLike_(amount, amountYearly, interval)

      if (includedInRing) {
        ringActiveSubs += 1
        ringTotalArr += ringArr
      }

      const orgIds = orgIdsBySubId.get(subId) || []
      if (includedInRing) {
        if (!orgIds.length) {
          ringUnmappedSubs += 1
          ringUnmappedArr += ringArr
        } else {
          orgIds.forEach(orgId => {
            const prev = ringArrByOrgId.get(orgId) || { arr: 0, count: 0 }
            prev.arr += ringArr
            prev.count += 1
            ringArrByOrgId.set(orgId, prev)
            ringTotalArrMapped += ringArr
          })
        }
      }

      ringSubRows.push([
        subId,
        status,
        interval,
        amount,
        amountYearly || '',
        ringArr || '',
        arrRawLike || '',
        discountPercent || 0,
        discountDuration || '',
        discountDurationMonths || '',
        includedInRing,
        excludedReason,
        orgIds.join(', ')
      ])
    })

    const subHeader = [
      'subscription_id',
      'status',
      'interval',
      'amount',
      'amount_yearly',
      'arr_ring',
      'arr_raw_like',
      'discount_percent',
      'discount_duration',
      'discount_duration_months',
      'included_in_ring',
      'excluded_reason',
      'org_ids'
    ]

    const unmappedRows = ringSubRows.filter(r => r[10] === true && !r[12])

    // Summary sheet
    const shOut = ARRRING_getOrCreateSheetCompat_(ss, ARRRING_CFG.OUT_SHEET)
    shOut.clearContents()

    const summary = [
      ['metric', 'value'],
      ['ring_total_arr', ringTotalArr],
      ['ring_total_arr_mapped', ringTotalArrMapped],
      ['ring_total_arr_unmapped', ringUnmappedArr],
      ['ring_active_subs', ringActiveSubs],
      ['ring_active_subs_unmapped', ringUnmappedSubs],
      ['arr_raw_total_arr_all', arrRawTotals.all],
      ['arr_raw_total_arr_active', arrRawTotals.active],
      ['arr_snapshot_latest_date', snapInfo ? snapInfo.latestDateLabel : ''],
      ['arr_snapshot_latest_total_arr', snapInfo ? snapInfo.latestTotal : ''],
      ['delta_ring_minus_arr_raw_active', ringTotalArr - arrRawTotals.active],
      ['delta_ring_minus_snapshot_latest', snapInfo ? (ringTotalArr - snapInfo.latestTotal) : '']
    ]

    shOut.getRange(1, 1, summary.length, 2).setValues(summary)

    const orgHeader = [
      'org_id',
      'org_name',
      'ring_arr_sum',
      'ring_active_subs',
      'arr_raw_total_arr',
      'arr_raw_status',
      'delta_ring_minus_arr_raw',
      'notes'
    ]

    const orgRows = []
    const orgIdSet = new Set()
    ringArrByOrgId.forEach((_, orgId) => orgIdSet.add(orgId))
    arrRawByOrgId.forEach((_, orgId) => orgIdSet.add(orgId))

    orgIdSet.forEach(orgId => {
      const ringInfo = ringArrByOrgId.get(orgId) || { arr: 0, count: 0 }
      const arrRaw = arrRawByOrgId.get(orgId) || {}
      const arrRawTotal = ARRRING_num_(arrRaw.total_arr)
      const status = ARRRING_str_(arrRaw.current_status)

      const notes = []
      if (!arrRawByOrgId.has(orgId)) notes.push('missing_in_arr_raw')
      if (ringInfo.count > 1) notes.push('multiple_active_subs')
      if (ringInfo.arr > 0 && arrRawTotal === 0) notes.push('arr_raw_zero')
      if (ringInfo.arr === 0 && arrRawTotal > 0) notes.push('no_active_subs_in_ring')

      orgRows.push([
        orgId,
        orgNameById.get(orgId) || '',
        ringInfo.arr,
        ringInfo.count,
        arrRawTotal,
        status,
        ringInfo.arr - arrRawTotal,
        notes.join(';')
      ])
    })

    orgRows.sort((a, b) => Math.abs(b[6]) - Math.abs(a[6]))

    const orgStartRow = summary.length + 2
    shOut.getRange(orgStartRow, 1, 1, orgHeader.length).setValues([orgHeader])
    if (orgRows.length) {
      shOut.getRange(orgStartRow + 1, 1, orgRows.length, orgHeader.length).setValues(orgRows)
    }

    const unmappedTitleRow = orgStartRow + 1 + orgRows.length + 2
    shOut.getRange(unmappedTitleRow, 1).setValue('Unmapped subscriptions (included in Ring)').setFontWeight('bold')
    shOut.getRange(unmappedTitleRow + 1, 1, 1, subHeader.length).setValues([subHeader])
    if (unmappedRows.length) {
      shOut.getRange(unmappedTitleRow + 2, 1, unmappedRows.length, subHeader.length).setValues(unmappedRows)
    }

    // Subs sheet
    const shSubs = ARRRING_getOrCreateSheetCompat_(ss, ARRRING_CFG.OUT_SUBS_SHEET)
    shSubs.clearContents()

    shSubs.getRange(1, 1, 1, subHeader.length).setValues([subHeader])
    if (ringSubRows.length) {
      shSubs.getRange(2, 1, ringSubRows.length, subHeader.length).setValues(ringSubRows)
    }

    // Formatting
    shOut.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#F3F3F3')
    shOut.getRange(orgStartRow, 1, 1, orgHeader.length).setFontWeight('bold').setBackground('#F3F3F3')
    shOut.getRange(unmappedTitleRow + 1, 1, 1, subHeader.length).setFontWeight('bold').setBackground('#F3F3F3')
    shOut.getRange(2, 2, summary.length - 1, 1).setNumberFormat('0.00')

    shSubs.getRange(1, 1, 1, subHeader.length).setFontWeight('bold').setBackground('#F3F3F3')
    const moneyCols = [4, 5, 6, 7]
    moneyCols.forEach(col => {
      shSubs.getRange(2, col, Math.max(0, ringSubRows.length), 1).setNumberFormat('0.00')
    })

    shOut.autoResizeColumns(1, Math.max(orgHeader.length, subHeader.length))
    shSubs.autoResizeColumns(1, subHeader.length)

    const seconds = (new Date() - t0) / 1000
    if (typeof writeSyncLog === 'function') {
      writeSyncLog('render_arr_ring_audit', 'ok', subs.length, ringSubRows.length, seconds, '')
    }

    return { rows_out: ringSubRows.length }
  })
}

/* =========================
 * Helpers
 * ========================= */

function ARRRING_readArrRaw_(sheet) {
  const headerRow = ARRRING_findHeaderRow_(sheet, ['org_id'])
  const header = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0]
  const headerMap = ARRRING_headerMap_(header)

  const orgIdIdx = headerMap['org_id']
  const totalArrIdx = headerMap['total_arr']
  const statusIdx = headerMap['current_status']

  const byOrgId = new Map()
  let totalAll = 0
  let totalActive = 0

  const lastRow = sheet.getLastRow()
  if (lastRow <= headerRow) {
    return { byOrgId, totals: { all: 0, active: 0 } }
  }

  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, header.length).getValues()
  data.forEach(r => {
    const orgId = ARRRING_str_(r[orgIdIdx])
    if (!orgId) return
    const totalArr = ARRRING_num_(r[totalArrIdx])
    const status = ARRRING_str_(statusIdx != null ? r[statusIdx] : '')

    byOrgId.set(orgId, { total_arr: totalArr, current_status: status })

    totalAll += totalArr
    if (status.toLowerCase() === 'active') totalActive += totalArr
  })

  return { byOrgId, totals: { all: totalAll, active: totalActive } }
}

function ARRRING_readArrSnapshot_(sheet) {
  const headerRow = ARRRING_findHeaderRow_(sheet, ['snapshot_date'])
  const header = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0]
  const headerMap = ARRRING_headerMap_(header)

  const snapIdx = headerMap['snapshot_date']
  const totalArrIdx = headerMap['total_arr']
  if (snapIdx == null || totalArrIdx == null) return null

  const lastRow = sheet.getLastRow()
  if (lastRow <= headerRow) return null

  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, header.length).getValues()
  let latestMs = 0
  data.forEach(r => {
    const d = ARRRING_parseDate_(r[snapIdx])
    if (d && d.getTime() > latestMs) latestMs = d.getTime()
  })

  if (!latestMs) return null

  let total = 0
  data.forEach(r => {
    const d = ARRRING_parseDate_(r[snapIdx])
    if (d && d.getTime() === latestMs) {
      total += ARRRING_num_(r[totalArrIdx])
    }
  })

  const tz = Session.getScriptTimeZone()
  const latestDateLabel = Utilities.formatDate(new Date(latestMs), tz, 'yyyy-MM-dd')

  return { latestTotal: total, latestDateLabel }
}

function ARRRING_buildMembershipsByOrgId_(mems) {
  const out = new Map()
  ;(mems || []).forEach(m => {
    const orgId = ARRRING_str_(m.org_id)
    if (!orgId) return
    const email = ARRRING_str_(m.email)
    const emailKey = ARRRING_normEmail_(m.email_key || email)
    if (!emailKey) return
    if (!out.has(orgId)) out.set(orgId, [])
    out.get(orgId).push({ email, email_key: emailKey })
  })
  return out
}

function ARRRING_buildUsersByEmailKey_(users) {
  const out = new Map()
  ;(users || []).forEach(u => {
    const email = ARRRING_str_(u.email)
    const emailKey = ARRRING_normEmail_(u.email_key || email)
    if (!emailKey) return
    out.set(emailKey, u)
  })
  return out
}

function ARRRING_buildSubIdsByOrgId_(membershipsByOrgId, userByEmailKey, users) {
  const out = new Map()
  membershipsByOrgId.forEach((members, orgId) => {
    members.forEach(m => {
      const key = ARRRING_normEmail_(m.email_key || m.email)
      const u = userByEmailKey.get(key)
      const subId = u ? ARRRING_str_(u.stripe_subscription_id || u.stripeSubscriptionId) : ''
      if (!subId) return
      if (!out.has(orgId)) out.set(orgId, new Set())
      out.get(orgId).add(subId)
    })
  })
  ;(users || []).forEach(u => {
    const orgId = ARRRING_str_(u.org_id)
    const subId = ARRRING_str_(u.stripe_subscription_id || u.stripeSubscriptionId)
    if (!orgId || !subId) return
    if (!out.has(orgId)) out.set(orgId, new Set())
    out.get(orgId).add(subId)
  })
  return out
}

function ARRRING_buildOrgIdsBySubId_(subIdsByOrgId) {
  const out = new Map()
  subIdsByOrgId.forEach((subSet, orgId) => {
    Array.from(subSet || []).forEach(subId => {
      if (!subId) return
      if (!out.has(subId)) out.set(subId, [])
      out.get(subId).push(orgId)
    })
  })
  return out
}

function ARRRING_findHeaderRow_(sheet, headerNames) {
  const lastRow = sheet.getLastRow()
  const scanRows = Math.min(5, Math.max(1, lastRow))
  const headerList = (headerNames || []).map(h => String(h || '').toLowerCase())

  for (let r = 1; r <= scanRows; r++) {
    const row = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getValues()[0]
    const rowKeys = row.map(h => String(h || '').toLowerCase().trim())
    const hasAll = headerList.every(h => rowKeys.indexOf(h) >= 0)
    if (hasAll) return r
  }
  return 1
}

function ARRRING_headerMap_(headerRow) {
  const out = {}
  ;(headerRow || []).forEach((h, i) => {
    const key = ARRRING_key_(h)
    if (!key) return
    if (!(key in out)) out[key] = i
  })
  return out
}

function ARRRING_computeArr_(amount, interval) {
  const amt = Number(amount || 0) || 0
  const intv = String(interval || '').toLowerCase().trim()
  if (intv === 'year' || intv === 'annual' || intv === 'yr') return amt
  return amt * 12
}

function ARRRING_computeArrRawLike_(amount, amountYearly, interval) {
  if (amountYearly != null && amountYearly !== '' && !isNaN(Number(amountYearly))) {
    return Number(amountYearly) || 0
  }
  return ARRRING_computeArr_(amount, interval)
}

function ARRRING_moneyAmount_(raw) {
  if (raw === null || raw === undefined || raw === '') return 0
  const n = ARRRING_num_(raw)
  if (!isFinite(n)) return 0
  return Math.round(n * 100) / 100
}

function ARRRING_parseDate_(v) {
  if (!v) return null
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v
  const d = new Date(String(v || ''))
  return isNaN(d.getTime()) ? null : d
}

function ARRRING_str_(v) {
  if (v === null || v === undefined) return ''
  return String(v).trim()
}

function ARRRING_num_(v) {
  if (v === null || v === undefined || v === '') return 0
  if (typeof v === 'number') return v
  const s = String(v).replace(/[^0-9.\-]/g, '').trim()
  const n = Number(s)
  return isNaN(n) ? 0 : n
}

function ARRRING_normEmail_(v) {
  const s = String(v || '').trim().toLowerCase()
  if (!s) return ''
  return s.replace(/\+[^@]+(?=@)/, '')
}

function ARRRING_key_(h) {
  return String(h || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_')
}

function ARRRING_readSheetObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1) return []

  const header = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || '').trim())

  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()

  return data.map(r => {
    const obj = {}
    header.forEach((h, i) => {
      if (!h) return
      obj[ARRRING_key_(h)] = r[i]
    })
    return obj
  })
}

function ARRRING_getOrCreateSheetCompat_(ss, name) {
  if (typeof getOrCreateSheet === 'function') {
    try { return getOrCreateSheet(ss, name) } catch (e) {}
    try { return getOrCreateSheet(name) } catch (e) {}
  }
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function ARRRING_lockWrapCompat_(lockName, fn) {
  if (typeof lockWrap === 'function') {
    try { return lockWrap(lockName, fn) } catch (e) { return lockWrap(fn) }
  }
  const lock = LockService.getScriptLock()
  const ok = lock.tryLock(300000)
  if (!ok) throw new Error(`Could not acquire lock: ${lockName}`)
  try { return fn() } finally { lock.releaseLock() }
}
