/**************************************************************
 * render_stripe_multi_sub_audit()
 *
 * Audits raw_stripe_subscriptions to find orgs with >1 subscription.
 * Maps Stripe subs -> orgs via raw_clerk_users + raw_clerk_memberships.
 **************************************************************/

const STRA_CFG = {
  SHEET_NAME: 'Stripe multi-sub audit',

  SOURCE_STRIPE: 'raw_stripe_subscriptions',
  SOURCE_USERS: 'raw_clerk_users',
  SOURCE_MEMBERSHIPS: 'raw_clerk_memberships',
  SOURCE_ORGS: 'raw_clerk_orgs',

  HEADER_ROW: 1,
  DATA_START_ROW: 2
}

function render_stripe_multi_sub_audit() {
  return lockWrapCompat_('render_stripe_multi_sub_audit', () => {
    const t0 = new Date()
    const ss = SpreadsheetApp.getActive()

    const shOut = getOrCreateSheetCompat_(ss, STRA_CFG.SHEET_NAME)
    const shStripe = ss.getSheetByName(STRA_CFG.SOURCE_STRIPE)
    const shUsers = ss.getSheetByName(STRA_CFG.SOURCE_USERS)
    const shMems = ss.getSheetByName(STRA_CFG.SOURCE_MEMBERSHIPS)
    const shOrgs = ss.getSheetByName(STRA_CFG.SOURCE_ORGS)

    if (!shStripe) throw new Error(`Missing sheet: ${STRA_CFG.SOURCE_STRIPE}`)
    if (!shUsers) throw new Error(`Missing sheet: ${STRA_CFG.SOURCE_USERS}`)
    if (!shMems) throw new Error(`Missing sheet: ${STRA_CFG.SOURCE_MEMBERSHIPS}`)
    if (!shOrgs) throw new Error(`Missing sheet: ${STRA_CFG.SOURCE_ORGS}`)

    const subs = STRA_readSheetObjects_(shStripe, 1)
    const users = STRA_readSheetObjects_(shUsers, 1)
    const mems = STRA_readSheetObjects_(shMems, 1)
    const orgs = STRA_readSheetObjects_(shOrgs, 1)

    const orgNameById = new Map()
    orgs.forEach(o => {
      const id = STRA_str_(o.org_id)
      if (!id) return
      const name = STRA_str_(o.org_name || o.org_slug)
      orgNameById.set(id, name)
    })

    const membershipsByOrgId = STRA_buildMembershipsByOrgId_(mems)
    const userByEmailKey = STRA_buildUsersByEmailKey_(users)
    const stripeBySubId = STRA_buildStripeBySubscriptionId_(subs)
    const subIdsByOrgId = STRA_buildSubIdsByOrgId_(membershipsByOrgId, userByEmailKey, users)

    const rows = []
    subIdsByOrgId.forEach((subIdSet, orgId) => {
      const ids = Array.from(subIdSet || [])
      if (ids.length <= 1) return

      const statuses = new Set()
      ids.forEach(id => {
        const row = stripeBySubId.get(id)
        if (!row) return
        const status = STRA_str_(row.status).toLowerCase()
        if (status) statuses.add(status)
      })

      const members = membershipsByOrgId.get(orgId) || []
      const ownerEmail = STRA_pickOrgOwnerEmail_(members)
      const orgName = orgNameById.get(orgId) || ''

      rows.push([
        orgId,
        orgName,
        ownerEmail,
        ids.length,
        ids.join(', '),
        Array.from(statuses).join(', ')
      ])
    })

    rows.sort((a, b) => {
      const aCount = Number(a[3]) || 0
      const bCount = Number(b[3]) || 0
      if (aCount !== bCount) return bCount - aCount
      return String(a[1] || '').localeCompare(String(b[1] || ''))
    })

    const headers = [
      'org_id',
      'org_name',
      'org_email',
      'subscription_count',
      'subscription_ids',
      'subscription_statuses'
    ]

    shOut.clearContents()
    shOut.getRange(STRA_CFG.HEADER_ROW, 1, 1, headers.length).setValues([headers])
    if (rows.length) {
      shOut.getRange(STRA_CFG.DATA_START_ROW, 1, rows.length, headers.length).setValues(rows)
    }
    shOut.setFrozenRows(STRA_CFG.HEADER_ROW)
    shOut.autoResizeColumns(1, headers.length)

    const seconds = (new Date() - t0) / 1000
    if (typeof writeSyncLog === 'function') {
      writeSyncLog('render_stripe_multi_sub_audit', 'ok', rows.length, rows.length, seconds, '')
    }

    return { rows_out: rows.length }
  })
}

/* =========================
 * Helpers
 * ========================= */

function STRA_readSheetObjects_(sheet, headerRow) {
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
      obj[STRA_key_(h)] = r[i]
    })
    return obj
  })
}

function STRA_key_(h) {
  return String(h || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_')
}

function STRA_buildMembershipsByOrgId_(mems) {
  const out = new Map()
  ;(mems || []).forEach(m => {
    const orgId = STRA_str_(m.org_id)
    if (!orgId) return
    const email = STRA_str_(m.email)
    const emailKey = STRA_normEmail_(m.email_key || email)
    const role = STRA_str_(m.role).toLowerCase()
    const createdAt = STRA_str_(m.created_at)

    if (!out.has(orgId)) out.set(orgId, [])
    out.get(orgId).push({ email, email_key: emailKey, role, created_at: createdAt })
  })
  return out
}

function STRA_buildUsersByEmailKey_(users) {
  const out = new Map()
  ;(users || []).forEach(u => {
    const email = STRA_str_(u.email)
    const emailKey = STRA_normEmail_(u.email_key || email)
    if (!emailKey) return
    out.set(emailKey, u)
  })
  return out
}

function STRA_buildSubIdsByOrgId_(membershipsByOrgId, userByEmailKey, users) {
  const out = new Map()
  membershipsByOrgId.forEach((members, orgId) => {
    const set = new Set()
    ;(members || []).forEach(m => {
      const key = STRA_normEmail_(m.email_key || m.email)
      if (!key) return
      const u = userByEmailKey.get(key)
      const subId = STRA_str_(u && (u.stripe_subscription_id || u.stripeSubscriptionId))
      if (subId) set.add(subId)
    })
    out.set(orgId, set)
  })
  ;(users || []).forEach(u => {
    const orgId = STRA_str_(u.org_id)
    const subId = STRA_str_(u && (u.stripe_subscription_id || u.stripeSubscriptionId))
    if (!orgId || !subId) return
    const set = out.get(orgId) || new Set()
    set.add(subId)
    out.set(orgId, set)
  })
  return out
}

function STRA_buildStripeBySubscriptionId_(subs) {
  const out = new Map()
  ;(subs || []).forEach(s => {
    const id = STRA_str_(s.stripe_subscription_id || s.subscription_id || s.id)
    if (!id) return
    out.set(id, s)
  })
  return out
}

function STRA_pickOrgOwnerEmail_(members) {
  const arr = (members || []).slice()
  arr.sort((a, b) => {
    const ams = STRA_toMs_(a.created_at) || 0
    const bms = STRA_toMs_(b.created_at) || 0
    return ams - bms
  })

  const owners = arr.filter(m => (m.role || '').includes('owner'))
  if (owners.length && owners[0].email) return owners[0].email

  const admins = arr.filter(m => (m.role || '').includes('admin'))
  if (admins.length && admins[0].email) return admins[0].email

  if (arr.length && arr[0].email) return arr[0].email
  return ''
}

function STRA_str_(v) {
  if (v === null || v === undefined) return ''
  return String(v).trim()
}

function STRA_normEmail_(v) {
  const s = String(v || '').trim().toLowerCase()
  if (!s) return ''
  return s.replace(/\+[^@]+(?=@)/, '')
}

function STRA_toMs_(iso) {
  const d = STRA_parseDate_(iso)
  return d ? d.getTime() : 0
}

function STRA_parseDate_(v) {
  if (!v) return null
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v
  const s = String(v || '').trim()
  if (!s) return null
  if (/^\d+$/.test(s)) {
    const n = Number(s)
    const ms = n > 1e12 ? n : n * 1000
    const d = new Date(ms)
    return isNaN(d.getTime()) ? null : d
  }
  const d = new Date(s)
  return isNaN(d.getTime()) ? null : d
}
