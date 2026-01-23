/**************************************************************
 * render_org_info_view()
 *
 * Updates "org_info" with:
 * - Org ID, Org Name, Org Owner, Service (manual)
 * - Trial Start Date + Trial End Date (same org-based logic as arr_raw_data)
 * - Subscription Start Date + Purchase Date (from Stripe)
 * - In clerk (member count from raw_clerk_memberships)
 * - Seats (paid seats from Stripe quantity_total via stripe_subscription_id)
 * - Diff (In clerk - Seats)
 * - UpSale (manual checkbox) ✅ NEW: appended at end and preserved by Org ID
 *
 * Conditional formatting on Diff:
 * - Diff > 0  (more members than seats)  -> RED
 * - Diff < 0  (fewer members than seats) -> YELLOW
 * - Diff = 0  (same)                     -> GREEN
 *
 * Manual values that persist:
 * - Service (by Org ID)
 * - UpSale (checkbox) (by Org ID) ✅ NEW
 **************************************************************/

const ORG_INFO_CFG = {
  SHEET_NAME: 'org_info',

  INPUTS: {
    CANON_ORGS: 'canon_orgs',
    CLERK_MEMBERSHIPS: 'raw_clerk_memberships',
    CLERK_USERS_RAW: 'raw_clerk_users',
    STRIPE_SUBSCRIPTIONS_RAW: 'raw_stripe_subscriptions'
  },

  HEADER_ROW: 1,
  DATA_START_ROW: 2,

  // Trial standard length (same as arr_raw_data)
  TRIAL_DAYS: 14,

  // ✅ UpSale moved to the END and will be preserved
  HEADERS: [
    'Org ID',
    'Org Name',
    'Org Owner',
    'Service',
    'In clerk',
    'Seats',
    'Diff',
    'subscription_start_date',
    'purchase_date',
    'trial_start_date',
    'trial_end_date',
    'UpSale'
  ],

  SERVICE_OPTIONS: ['White Glove', 'Hands On']
}

function render_org_info_view() {
  lockWrapCompat_('render_org_info_view', () => {
    const t0 = new Date()
    try {
      const ss = SpreadsheetApp.getActive()
      const sh = ORGINFO_getOrCreateSheet_(ss, ORG_INFO_CFG.SHEET_NAME)

      const shOrgs = ss.getSheetByName(ORG_INFO_CFG.INPUTS.CANON_ORGS)
      const shMems = ss.getSheetByName(ORG_INFO_CFG.INPUTS.CLERK_MEMBERSHIPS)
      const shClerkUsers = ss.getSheetByName(ORG_INFO_CFG.INPUTS.CLERK_USERS_RAW)
      const shStripeSubs = ss.getSheetByName(ORG_INFO_CFG.INPUTS.STRIPE_SUBSCRIPTIONS_RAW)

      if (!shOrgs) throw new Error(`Missing input sheet: ${ORG_INFO_CFG.INPUTS.CANON_ORGS}`)
      if (!shMems) throw new Error(`Missing input sheet: ${ORG_INFO_CFG.INPUTS.CLERK_MEMBERSHIPS}`)
      if (!shClerkUsers) throw new Error(`Missing input sheet: ${ORG_INFO_CFG.INPUTS.CLERK_USERS_RAW}`)
      if (!shStripeSubs) throw new Error(`Missing input sheet: ${ORG_INFO_CFG.INPUTS.STRIPE_SUBSCRIPTIONS_RAW}`)

      const orgs = ORGINFO_readSheetObjects_(shOrgs, 1)
      const mems = ORGINFO_readSheetObjects_(shMems, 1)
      const users = ORGINFO_readSheetObjects_(shClerkUsers, 1)
      const subs = ORGINFO_readSheetObjects_(shStripeSubs, 1)

      // Preserve manual "Service" + "UpSale" from existing org_info by Org ID
      const existingManualByOrgId = ORGINFO_readExistingManualByOrgId_(sh)

      // Build org_id -> {members:Set(emailKey), owners:[], admins:[], any:[]}
      const memAgg = ORGINFO_buildMembershipAgg_(mems)

      // Trial logic indexes (org-based, same as arr_raw_data)
      const membershipsByOrgId = ORGINFO_buildMembershipsByOrgId_(mems) // orgId -> [{email,email_key,role,created_at}]
      const userByEmailKey = ORGINFO_buildUsersByEmailKey_(users)       // email_key -> user obj
      const subIdsByOrgId = ORGINFO_buildSubIdsByOrgId_(membershipsByOrgId, userByEmailKey, users)
      const stripeBySubId = ORGINFO_buildStripeBySubscriptionId_(subs)

      // Seats from Stripe: build indexes
      const subIdByEmailKey = ORGINFO_buildSubIdByEmailKey_(users)                 // email_key -> stripe_subscription_id
      const qtyBySubId = ORGINFO_buildQuantityBySubId_(subs)                      // stripe_subscription_id -> quantity_total
      const seatsByOrgId = ORGINFO_buildSeatsByOrgId_(memAgg, subIdByEmailKey, qtyBySubId) // org_id -> seats

      const out = []

      const sortedOrgs = orgs
        .map(o => ({
          orgId: String(o.org_id || '').trim(),
          orgName: String(o.org_name || '').trim(),
          orgCreatedAt: o.org_created_at || o.created_at || ''
        }))
        .filter(o => o.orgId || o.orgName)
        .sort((a, b) => (a.orgName || '').localeCompare(b.orgName || '') || (a.orgId || '').localeCompare(b.orgId || ''))

      for (const o of sortedOrgs) {
        const orgId = o.orgId
        const orgName = o.orgName
        const orgCreatedAt = o.orgCreatedAt

        const agg = memAgg.get(orgId) || { members: new Set(), owners: [], admins: [], any: [] }
        const inClerk = agg.members.size

        const ownerEmail =
          (agg.owners[0] || '') ||
          (agg.admins[0] || '') ||
          (agg.any[0] || '') ||
          ''

        const manual = existingManualByOrgId.get(orgId) || {}
        const service = String(manual.service || '').trim()

        // ✅ preserve checkbox state (true/false) by orgId
        const upsale = manual.upsale === true

        const seats = ORGINFO_safeInt_(seatsByOrgId.get(orgId) || 0)
        const diff = (Number(inClerk) || 0) - (Number(seats) || 0)

        const trialOwnerEmail = ORGINFO_pickOrgOwnerEmail_(membershipsByOrgId.get(orgId) || []) || ownerEmail
        const orgCreatedIso = ORGINFO_toIsoOrBlank_(orgCreatedAt)
        const trialStart = orgCreatedIso
        const subRows = ORGINFO_rowsFromSubIds_(subIdsByOrgId.get(orgId), stripeBySubId)
        const trialEnd = ORGINFO_computeTrialEndIso_({
          trialStartIso: trialStart,
          subscriptions: subRows,
          trialDays: ORG_INFO_CFG.TRIAL_DAYS
        })

        const subscriptionStart = ORGINFO_minIso_(
          subRows.map(r => ORGINFO_toIsoOrBlank_(r.created_at)).filter(Boolean)
        )
        const purchaseDate = ORGINFO_minIso_(
          subRows.map(r => ORGINFO_toIsoOrBlank_(r.first_payment_at)).filter(Boolean)
        )

        out.push([
          orgId,
          orgName,
          ownerEmail,
          service,
          inClerk,
          seats,
          diff,
          ORGINFO_isoToDateOrBlank_(subscriptionStart),
          ORGINFO_isoToDateOrBlank_(purchaseDate),
          ORGINFO_isoToDateOrBlank_(trialStart),
          ORGINFO_isoToDateOrBlank_(trialEnd),
          upsale
        ])
      }

      // Rebuild sheet
      sh.clear()

      // Headers
      sh.getRange(ORG_INFO_CFG.HEADER_ROW, 1, 1, ORG_INFO_CFG.HEADERS.length).setValues([ORG_INFO_CFG.HEADERS])
      sh.setFrozenRows(1)

      // Data
      if (out.length) {
        ORGINFO_batchSetValues_(sh, ORG_INFO_CFG.DATA_START_ROW, 1, out, 2000)
      }

      // Basic formatting
      ORGINFO_applyFormats_(sh, out.length)

      // Service dropdown
      ORGINFO_applyServiceDropdown_(sh)

      // ✅ UpSale checkbox column formatting + validation
      ORGINFO_applyUpsaleCheckboxes_(sh)

      // Diff conditional formatting
      ORGINFO_applyDiffConditionalFormatting_(sh)

      sh.autoResizeColumns(1, ORG_INFO_CFG.HEADERS.length)

      if (typeof writeSyncLog === 'function') {
        writeSyncLog(
          'render_org_info_view',
          'ok',
          orgs.length,
          out.length,
          (new Date() - t0) / 1000,
          ''
        )
      } else {
        Logger.log(`[render_org_info_view] ok rows_in=${orgs.length} rows_out=${out.length}`)
      }

      return { rows_in: orgs.length, rows_out: out.length }
    } catch (err) {
      if (typeof writeSyncLog === 'function') {
        writeSyncLog('render_org_info_view', 'error', '', '', '', String(err && err.message ? err.message : err))
      }
      throw err
    }
  })
}

/* =========================
 * Preserve manual Service + UpSale
 * ========================= */

function ORGINFO_readExistingManualByOrgId_(sheet) {
  const out = new Map()

  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < 2 || lastCol < 1) return out

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())
  const orgIdIdx = header.findIndex(h => h.toLowerCase() === 'org id')
  const serviceIdx = header.findIndex(h => h.toLowerCase() === 'service')
  const upsaleIdx = header.findIndex(h => h.toLowerCase() === 'upsale')

  if (orgIdIdx < 0) return out

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
  data.forEach(r => {
    const orgId = String(r[orgIdIdx] || '').trim()
    if (!orgId) return

    const service = (serviceIdx >= 0) ? String(r[serviceIdx] || '').trim() : ''
    const upsaleRaw = (upsaleIdx >= 0) ? r[upsaleIdx] : false
    const upsale = (upsaleRaw === true) || String(upsaleRaw || '').toLowerCase().trim() === 'true'

    out.set(orgId, { service, upsale })
  })

  return out
}

/* =========================
 * Seats from Stripe (indexes)
 * ========================= */

function ORGINFO_buildSubIdByEmailKey_(source) {
  const out = new Map()

  if (Array.isArray(source)) {
    ;(source || []).forEach(u => {
      const emailKeyRaw = String(u.email_key || u.email || '').trim()
      const emailKey = normalizeEmailCompat_(emailKeyRaw)
      const subId = String(u.stripe_subscription_id || u.stripeSubscriptionId || '').trim()
      if (!emailKey || !subId) return
      out.set(emailKey, subId)
    })
    return out
  }

  const sheet = source
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < 2) return out

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim().toLowerCase())
  const emailKeyIdx = header.findIndex(h => h === 'email_key')
  const subIdIdx = header.findIndex(h => h === 'stripe_subscription_id')

  if (emailKeyIdx < 0 || subIdIdx < 0) {
    throw new Error('raw_clerk_users must have headers: email_key, stripe_subscription_id')
  }

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
  data.forEach(r => {
    const emailKey = String(r[emailKeyIdx] || '').trim()
    const subId = String(r[subIdIdx] || '').trim()
    if (!emailKey || !subId) return
    out.set(emailKey, subId)
  })

  return out
}

function ORGINFO_buildQuantityBySubId_(source) {
  const out = new Map()

  if (Array.isArray(source)) {
    ;(source || []).forEach(r => {
      const subId = String(r.stripe_subscription_id || r.subscription_id || r.id || '').trim()
      if (!subId) return
      const qty = ORGINFO_safeInt_(r.quantity_total)
      out.set(subId, qty)
    })
    return out
  }

  const sheet = source
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < 2) return out

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim().toLowerCase())
  const subIdIdx = header.findIndex(h => h === 'stripe_subscription_id')
  const qtyIdx = header.findIndex(h => h === 'quantity_total')

  if (subIdIdx < 0 || qtyIdx < 0) {
    throw new Error('raw_stripe_subscriptions must have headers: stripe_subscription_id, quantity_total')
  }

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
  data.forEach(r => {
    const subId = String(r[subIdIdx] || '').trim()
    if (!subId) return
    const qty = ORGINFO_safeInt_(r[qtyIdx])
    out.set(subId, qty)
  })

  return out
}

function ORGINFO_buildSeatsByOrgId_(memAgg, subIdByEmailKey, qtyBySubId) {
  const out = new Map()

  ;(memAgg || new Map()).forEach((agg, orgId) => {
    const memberKeys = Array.from(agg.members || [])
    let maxSeats = 0

    memberKeys.forEach(emailKey => {
      const subId = subIdByEmailKey.get(emailKey)
      if (!subId) return
      const qty = ORGINFO_safeInt_(qtyBySubId.get(subId))
      if (qty > maxSeats) maxSeats = qty
    })

    out.set(orgId, maxSeats)
  })

  return out
}

/* =========================
 * Membership aggregation
 * ========================= */

function ORGINFO_buildMembershipAgg_(mems) {
  const map = new Map()

  ;(mems || []).forEach(m => {
    const orgId = String(m.org_id || '').trim()
    if (!orgId) return

    const email = String(m.email || '').trim()
    const emailKey = normalizeEmailCompat_(email) || email
    if (!emailKey) return

    const role = String(m.role || '').toLowerCase().trim()

    if (!map.has(orgId)) {
      map.set(orgId, { members: new Set(), owners: [], admins: [], any: [] })
    }
    const agg = map.get(orgId)

    agg.members.add(emailKey)
    agg.any.push(email)

    const isOwner = role.includes('owner') || role === 'owner'
    const isAdmin = role.includes('admin') || role === 'admin' || role === 'org:admin'

    if (isOwner) agg.owners.push(email)
    else if (isAdmin) agg.admins.push(email)
  })

  map.forEach(agg => {
    agg.any = ORGINFO_dedupeKeepOrder_(agg.any)
    agg.owners = ORGINFO_dedupeKeepOrder_(agg.owners)
    agg.admins = ORGINFO_dedupeKeepOrder_(agg.admins)
  })

  return map
}

/* =========================
 * Trial dates (org-based)
 * ========================= */

function ORGINFO_buildMembershipsByOrgId_(mems) {
  const out = new Map()
  ;(mems || []).forEach(m => {
    const orgId = ORGINFO_str_(m.org_id)
    if (!orgId) return
    const email = ORGINFO_str_(m.email) || ORGINFO_str_(m.email_key)
    const emailKey = ORGINFO_normEmail_(m.email_key || email)
    const role = ORGINFO_str_(m.role).toLowerCase()
    const createdAt = ORGINFO_toIsoOrBlank_(m.created_at)

    if (!out.has(orgId)) out.set(orgId, [])
    out.get(orgId).push({ email, email_key: emailKey, role, created_at: createdAt })
  })
  return out
}

function ORGINFO_pickOrgOwnerEmail_(members) {
  const arr = (members || []).slice()

  // Sort by created_at ascending (earliest)
  arr.sort((a, b) => {
    const ams = ORGINFO_toMs_(a.created_at) || 0
    const bms = ORGINFO_toMs_(b.created_at) || 0
    return ams - bms
  })

  const owners = arr.filter(m => (m.role || '').includes('owner'))
  if (owners.length) return owners[0].email || owners[0].email_key || ''

  const admins = arr.filter(m => (m.role || '').includes('admin'))
  if (admins.length) return admins[0].email || admins[0].email_key || ''

  if (arr.length) return arr[0].email || arr[0].email_key || ''
  return ''
}

function ORGINFO_buildUsersByEmailKey_(users) {
  const out = new Map()
  ;(users || []).forEach(u => {
    const email = ORGINFO_str_(u.email)
    const emailKeyRaw = ORGINFO_str_(u.email_key || email)
    const emailKey = ORGINFO_normEmail_(emailKeyRaw)
    if (!emailKey) return
    out.set(emailKey, u)
  })
  return out
}

function ORGINFO_buildSubIdsByOrgId_(membershipsByOrgId, userByEmailKey, users) {
  const out = new Map()

  ;(membershipsByOrgId || new Map()).forEach((members, orgId) => {
    const set = out.get(orgId) || new Set()
    ;(members || []).forEach(m => {
      const key = ORGINFO_normEmail_(m.email_key || m.email)
      if (!key) return
      const u = userByEmailKey.get(key)
      const subId = ORGINFO_str_(u && (u.stripe_subscription_id || u.stripeSubscriptionId))
      if (subId) set.add(subId)
    })
    out.set(orgId, set)
  })

  ;(users || []).forEach(u => {
    const orgId = ORGINFO_str_(u.org_id)
    const subId = ORGINFO_str_(u.stripe_subscription_id || u.stripeSubscriptionId)
    if (!orgId || !subId) return
    const set = out.get(orgId) || new Set()
    set.add(subId)
    out.set(orgId, set)
  })

  return out
}

function ORGINFO_buildStripeBySubscriptionId_(subs) {
  const out = new Map()
  ;(subs || []).forEach(s => {
    const id = ORGINFO_str_(s.stripe_subscription_id || s.subscription_id || s.id)
    if (!id) return
    out.set(id, s)
  })
  return out
}

function ORGINFO_rowsFromSubIds_(subIdSet, stripeBySubId) {
  const rows = []
  ;(subIdSet || new Set()).forEach(id => {
    const row = stripeBySubId.get(id)
    if (row) rows.push(row)
  })
  return rows
}

function ORGINFO_computeTrialEndIso_({ trialStartIso, subscriptions, trialDays }) {
  const ts = ORGINFO_parseIsoDate_(trialStartIso)
  if (!ts) return ''

  const standardEnd = new Date(ts.getTime() + Number(trialDays || 14) * 24 * 60 * 60 * 1000)
  const candidate = ORGINFO_pickTrialExtensionSub_(subscriptions, ts, standardEnd)

  if (candidate) {
    const end = ORGINFO_addMonths_(candidate.start, candidate.months)
    return end.toISOString()
  }

  return standardEnd.toISOString()
}

function ORGINFO_pickTrialExtensionSub_(subscriptions, trialStart, standardEnd) {
  const startMs = trialStart.getTime()
  const endMs = standardEnd.getTime()
  let best = null

  ;(subscriptions || []).forEach(sub => {
    const pct = Number(sub.discount_percent)
    const months = Number(sub.discount_duration_months)
    if (pct !== 100 || !isFinite(months) || months <= 0) return

    const startIso = ORGINFO_getSubscriptionStartIso_(sub)
    const start = ORGINFO_parseIsoDate_(startIso)
    if (!start) return

    const t = start.getTime()
    if (t < startMs || t > endMs) return

    if (!best || t < best.start.getTime()) best = { start, months }
  })

  return best
}

function ORGINFO_getSubscriptionStartIso_(sub) {
  if (!sub) return ''
  return ORGINFO_toIsoOrBlank_(sub.created_at)
}

function ORGINFO_addMonths_(dateObj, months) {
  const d = new Date(dateObj.getTime())
  const m = Number(months) || 0
  const day = d.getUTCDate()
  d.setUTCMonth(d.getUTCMonth() + m)

  // Best-effort clamp for month length differences:
  if (d.getUTCDate() !== day) d.setUTCDate(0)
  return d
}

/* =========================
 * Sheet IO
 * ========================= */

function ORGINFO_readSheetObjects_(sheet, headerRow) {
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
      obj[ORGINFO_key_(h)] = r[i]
    })
    return obj
  })
}

function ORGINFO_key_(h) {
  return String(h || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_')
}

/* =========================
 * Formatting + validation
 * ========================= */

function ORGINFO_applyFormats_(sheet, numDataRows) {
  const headerRange = sheet.getRange(1, 1, 1, ORG_INFO_CFG.HEADERS.length)
  headerRange
    .setFontWeight('bold')
    .setBackground('#F6F4F0')
    .setFontColor('#2F2B27')

  if (!numDataRows) return

  const range = sheet.getRange(2, 1, numDataRows, ORG_INFO_CFG.HEADERS.length)
  range.setVerticalAlignment('middle').setWrap(false)

  const colInClerk = ORG_INFO_CFG.HEADERS.indexOf('In clerk') + 1
  const colSeats = ORG_INFO_CFG.HEADERS.indexOf('Seats') + 1
  const colDiff = ORG_INFO_CFG.HEADERS.indexOf('Diff') + 1
  const colSubStart = ORG_INFO_CFG.HEADERS.indexOf('subscription_start_date') + 1
  const colPurchase = ORG_INFO_CFG.HEADERS.indexOf('purchase_date') + 1
  const colTrialStart = ORG_INFO_CFG.HEADERS.indexOf('trial_start_date') + 1
  const colTrialEnd = ORG_INFO_CFG.HEADERS.indexOf('trial_end_date') + 1

  if (colInClerk > 0) sheet.getRange(2, colInClerk, numDataRows, 1).setNumberFormat('0').setHorizontalAlignment('center')
  if (colSeats > 0) sheet.getRange(2, colSeats, numDataRows, 1).setNumberFormat('0').setHorizontalAlignment('center')
  if (colDiff > 0) sheet.getRange(2, colDiff, numDataRows, 1).setNumberFormat('0').setHorizontalAlignment('center')
  if (colSubStart > 0) sheet.getRange(2, colSubStart, numDataRows, 1).setNumberFormat('MM-dd-yy')
  if (colPurchase > 0) sheet.getRange(2, colPurchase, numDataRows, 1).setNumberFormat('MM-dd-yy')
  if (colTrialStart > 0) sheet.getRange(2, colTrialStart, numDataRows, 1).setNumberFormat('MM-dd-yy')
  if (colTrialEnd > 0) sheet.getRange(2, colTrialEnd, numDataRows, 1).setNumberFormat('MM-dd-yy')
}

function ORGINFO_applyServiceDropdown_(sheet) {
  const serviceCol = ORG_INFO_CFG.HEADERS.indexOf('Service') + 1
  if (serviceCol <= 0) return

  const maxRows = sheet.getMaxRows()
  const startRow = ORG_INFO_CFG.DATA_START_ROW
  const numRows = Math.max(1, maxRows - startRow + 1)

  const list = [''].concat(ORG_INFO_CFG.SERVICE_OPTIONS.map(String))

  const rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(true)
    .build()

  sheet.getRange(startRow, serviceCol, numRows, 1).setDataValidation(rule)
}

// ✅ NEW
function ORGINFO_applyUpsaleCheckboxes_(sheet) {
  const upsaleCol = ORG_INFO_CFG.HEADERS.indexOf('UpSale') + 1
  if (upsaleCol <= 0) return

  const startRow = ORG_INFO_CFG.DATA_START_ROW
  const numRows = Math.max(1, sheet.getMaxRows() - startRow + 1)

  // Ensure checkbox validation (Google Sheets checkboxes are boolean TRUE/FALSE)
  const rule = SpreadsheetApp
    .newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(true)
    .build()

  sheet.getRange(startRow, upsaleCol, numRows, 1).setDataValidation(rule)
}

function ORGINFO_applyDiffConditionalFormatting_(sheet) {
  const colDiff = ORG_INFO_CFG.HEADERS.indexOf('Diff') + 1
  const colSeats = ORG_INFO_CFG.HEADERS.indexOf('Seats') + 1
  if (colDiff <= 0 || colSeats <= 0) return

  const startRow = ORG_INFO_CFG.DATA_START_ROW
  const numRows = Math.max(1, sheet.getMaxRows() - startRow + 1)

  const range = sheet.getRange(startRow, colDiff, numRows, 1)

  // Remove prior rules that target this column (best-effort)
  const existing = sheet.getConditionalFormatRules() || []
  const kept = existing.filter(rule => {
    try {
      const rs = rule.getRanges() || []
      return !rs.some(r => r.getColumn() === colDiff)
    } catch (e) {
      return true
    }
  })

  const diffLetter = ORGINFO_colLetter_(colDiff)
  const seatsLetter = ORGINFO_colLetter_(colSeats)

  // Only color when Seats > 0
  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([range])
    .whenFormulaSatisfied(`=AND($${seatsLetter}${startRow}>0,$${diffLetter}${startRow}>0)`)
    .setBackground('#F8D7DA')
    .setFontColor('#2F2B27')
    .build()

  const yellowRule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([range])
    .whenFormulaSatisfied(`=AND($${seatsLetter}${startRow}>0,$${diffLetter}${startRow}<0)`)
    .setBackground('#FFF3CD')
    .setFontColor('#2F2B27')
    .build()

  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([range])
    .whenFormulaSatisfied(`=AND($${seatsLetter}${startRow}>0,$${diffLetter}${startRow}=0)`)
    .setBackground('#D4EDDA')
    .setFontColor('#2F2B27')
    .build()

  sheet.setConditionalFormatRules(kept.concat([redRule, yellowRule, greenRule]))
}

/* =========================
 * Tiny utils
 * ========================= */

function ORGINFO_str_(v) {
  if (v === null || v === undefined) return ''
  return String(v).trim()
}

function ORGINFO_normEmail_(v) {
  const s = String(v || '').trim().toLowerCase()
  if (!s) return ''
  return s.replace(/\+[^@]+(?=@)/, '')
}

function ORGINFO_toIsoOrBlank_(v) {
  if (!v) return ''
  if (v instanceof Date) return v.toISOString()

  const s = String(v || '').trim()
  if (!s) return ''

  // ISO string already
  if (s.includes('T') && s.endsWith('Z')) return s

  // unix seconds/millis
  if (/^\d+$/.test(s)) {
    const n = Number(s)
    const ms = n > 1e12 ? n : n * 1000
    const d = new Date(ms)
    return isNaN(d.getTime()) ? '' : d.toISOString()
  }

  const d = new Date(s)
  return isNaN(d.getTime()) ? '' : d.toISOString()
}

function ORGINFO_parseIsoDate_(iso) {
  const s = String(iso || '').trim()
  if (!s) return null
  const d = new Date(s)
  return isNaN(d.getTime()) ? null : d
}

function ORGINFO_isoToDateOrBlank_(iso) {
  const d = ORGINFO_parseIsoDate_(iso)
  return d ? d : ''
}

function ORGINFO_toMs_(iso) {
  const d = ORGINFO_parseIsoDate_(iso)
  return d ? d.getTime() : 0
}

function ORGINFO_firstNonEmpty_() {
  for (let i = 0; i < arguments.length; i++) {
    const v = ORGINFO_str_(arguments[i])
    if (v) return v
  }
  return ''
}

function ORGINFO_minIso_(isos) {
  let best = null
  ;(isos || []).forEach(s => {
    const d = ORGINFO_parseIsoDate_(s)
    if (!d) return
    if (!best || d.getTime() < best.getTime()) best = d
  })
  return best ? best.toISOString() : ''
}

function ORGINFO_safeInt_(v) {
  const n = Number(v)
  if (!isFinite(n) || isNaN(n)) return 0
  return Math.max(0, Math.floor(n))
}

function ORGINFO_dedupeKeepOrder_(arr) {
  const seen = new Set()
  const out = []
  ;(arr || []).forEach(v => {
    const s = String(v || '').trim()
    if (!s) return
    if (seen.has(s)) return
    seen.add(s)
    out.push(s)
  })
  return out
}

function ORGINFO_colLetter_(colNum1Based) {
  let n = Number(colNum1Based)
  let s = ''
  while (n > 0) {
    const r = (n - 1) % 26
    s = String.fromCharCode(65 + r) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

function normalizeEmailCompat_(email) {
  if (typeof normalizeEmail === 'function') return normalizeEmail(email)
  return String(email || '').trim().toLowerCase()
}

/* =========================
 * Compatibility wrappers
 * ========================= */

function ORGINFO_getOrCreateSheet_(ss, name) {
  if (typeof getOrCreateSheet === 'function') {
    try { return getOrCreateSheet(ss, name) } catch (e) {}
    try { return getOrCreateSheet(name) } catch (e) {}
  }
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function ORGINFO_batchSetValues_(sheet, startRow, startCol, values, chunkSize) {
  if (!values || !values.length) return
  const size = chunkSize || 2000
  for (let i = 0; i < values.length; i += size) {
    const chunk = values.slice(i, i + size)
    sheet.getRange(startRow + i, startCol, chunk.length, chunk[0].length).setValues(chunk)
  }
}

function lockWrapCompat_(lockName, fn) {
  if (typeof lockWrap === 'function') {
    try { return lockWrap(lockName, fn) } catch (e) { return lockWrap(fn) }
  }
  const lock = LockService.getScriptLock()
  const ok = lock.tryLock(300000)
  if (!ok) throw new Error(`Could not acquire lock: ${lockName}`)
  try { return fn() } finally { lock.releaseLock() }
}
