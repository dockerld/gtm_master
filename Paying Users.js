/**************************************************************
 * render_paying_users_snapshot()
 *
 * Builds/updates "Paying Users" as an append-only monthly snapshot
 * of MRR by org (active subscriptions only).
 *
 * Behavior:
 * - Uses raw_stripe_subscriptions as source of truth
 * - Maps subscriptions to orgs via raw_clerk_users + raw_clerk_memberships
 * - Appends a new month column on every run (even same month)
 * - Writes MRR per org for the snapshot month only
 * - Includes 100% discount months as 0 MRR (keeps org listed)
 * - Yearly subs are normalized to monthly
 **************************************************************/

const PAYING_CFG = {
  SHEET_NAME: 'Paying Users',

  SOURCE_STRIPE: 'raw_stripe_subscriptions',
  SOURCE_USERS: 'raw_clerk_users',
  SOURCE_MEMBERSHIPS: 'raw_clerk_memberships',
  SOURCE_ORGS: 'raw_clerk_orgs',

  SUMMARY_ROW_START: 2,
  SUMMARY_ROWS: 4,
  HEADER_ROW: 7,
  DATA_START_ROW: 8,

  ID_HEADER: 'org_id',
  NAME_HEADER: 'org_name',
  EMAIL_HEADER: 'org_email',

  MONTH_FMT: 'MMM-yyyy',
  FORECAST_MONTHS: 24,
  BACKFILL_START_MONTH_KEY: '2025-09'
}

function render_paying_users_snapshot() {
  return lockWrapCompat_('render_paying_users_snapshot', () => {
    return PAYING_renderSnapshot_({ reset: false })
  })
}

function reset_paying_users_sheet() {
  return lockWrapCompat_('reset_paying_users_sheet', () => {
    return PAYING_renderSnapshot_({ reset: true })
  })
}

function PAYING_renderSnapshot_(opts) {
  const t0 = new Date()
  const ss = SpreadsheetApp.getActive()

  const sh = getOrCreateSheetCompat_(ss, PAYING_CFG.SHEET_NAME)
  const shStripe = ss.getSheetByName(PAYING_CFG.SOURCE_STRIPE)
  const shUsers = ss.getSheetByName(PAYING_CFG.SOURCE_USERS)
  const shMems = ss.getSheetByName(PAYING_CFG.SOURCE_MEMBERSHIPS)
  const shOrgs = ss.getSheetByName(PAYING_CFG.SOURCE_ORGS)

  if (!shStripe) throw new Error(`Missing sheet: ${PAYING_CFG.SOURCE_STRIPE}`)
  if (!shUsers) throw new Error(`Missing sheet: ${PAYING_CFG.SOURCE_USERS}`)
  if (!shMems) throw new Error(`Missing sheet: ${PAYING_CFG.SOURCE_MEMBERSHIPS}`)
  if (!shOrgs) throw new Error(`Missing sheet: ${PAYING_CFG.SOURCE_ORGS}`)

  const allowPastUpdate = !!(opts && opts.reset)

  if (opts && opts.reset) {
    sh.clear()
  }

  PAYING_ensureSummaryLayout_(sh)

  const subs = PAYING_readSheetObjects_(shStripe, 1)
  const users = PAYING_readSheetObjects_(shUsers, 1)
  const mems = PAYING_readSheetObjects_(shMems, 1)
  const orgs = PAYING_readSheetObjects_(shOrgs, 1)

  const orgNameById = new Map()
  orgs.forEach(o => {
    const id = PAYING_str_(o.org_id)
    if (!id) return
    const name = PAYING_str_(o.org_name || o.org_slug)
    orgNameById.set(id, name)
  })

  const membershipsByOrgId = PAYING_buildMembershipsByOrgId_(mems)
  const userByEmailKey = PAYING_buildUsersByEmailKey_(users)
  const stripeBySubId = PAYING_buildStripeBySubscriptionId_(subs)
  const subIdsByOrgId = PAYING_buildSubIdsByOrgId_(membershipsByOrgId, userByEmailKey)

  const orgSubs = new Map()
  const orgNoActive = new Map()
  let earliestStartKey = ''

  subIdsByOrgId.forEach((subIdSet, orgId) => {
    const rows = PAYING_rowsFromSubIds_(subIdSet, stripeBySubId)
    if (!rows.length) return

    rows.forEach(r => {
      const startKey = PAYING_subStartKey_(r)
      if (!startKey) return
      if (!earliestStartKey || startKey < earliestStartKey) earliestStartKey = startKey
    })

    const activeRows = rows.filter(r => PAYING_str_(r.status).toLowerCase() === 'active')
    if (!activeRows.length) {
      const members = membershipsByOrgId.get(orgId) || []
      const ownerEmail = PAYING_pickOrgOwnerEmail_(members)
      const orgName = orgNameById.get(orgId) || ''
      orgNoActive.set(orgId, { org_name: orgName, org_email: ownerEmail })
      return
    }

    orgSubs.set(orgId, activeRows)

    activeRows.forEach(r => {
      const startKey = PAYING_subStartKey_(r)
      if (!startKey) return
      if (!earliestStartKey || startKey < earliestStartKey) earliestStartKey = startKey
    })
  })

  const overrideStart = PAYING_monthKeyFromHeader_(PAYING_CFG.BACKFILL_START_MONTH_KEY)
  if (overrideStart) earliestStartKey = overrideStart

  if (!earliestStartKey) {
    Logger.log('No active subscriptions with a valid start date. Paying Users not updated.')
    return
  }

  const snapshotDate = new Date()
  const currentMonthKey = PAYING_monthKeyFromDate_(snapshotDate)
  const endMonthKey = PAYING_addMonths_(currentMonthKey, PAYING_CFG.FORECAST_MONTHS)
  const monthKeys = PAYING_monthRange_(earliestStartKey, endMonthKey)
  const monthIndexByKey = new Map(monthKeys.map((k, i) => [k, i]))

  const orgMonthData = new Map()
  orgSubs.forEach((activeRows, orgId) => {
    const values = new Array(monthKeys.length).fill(0)
    let orgStartIdx = null

    activeRows.forEach(r => {
      const startKey = PAYING_subStartKey_(r)
      if (!startKey) return

      let startIdx = PAYING_monthDiff_(earliestStartKey, startKey)
      if (startIdx < 0) startIdx = 0
      if (startIdx >= monthKeys.length) return
      if (orgStartIdx == null || startIdx < orgStartIdx) orgStartIdx = startIdx

      const baseMrr = PAYING_calcSubBaseMrr_(r)
      const freeInfo = PAYING_discountFreeMonths_(r)

      for (let i = startIdx; i < monthKeys.length; i++) {
        let val = baseMrr
        if (freeInfo.forever) {
          val = 0
        } else if (freeInfo.months > 0 && (i - startIdx) < freeInfo.months) {
          val = 0
        }
        values[i] += val
      }
    })

    if (orgStartIdx == null) return

    const members = membershipsByOrgId.get(orgId) || []
    const ownerEmail = PAYING_pickOrgOwnerEmail_(members)
    const orgName = orgNameById.get(orgId) || ''

    orgMonthData.set(orgId, {
      org_name: orgName,
      org_email: ownerEmail,
      start_idx: orgStartIdx,
      values
    })
  })

  let lastRow = sh.getLastRow()
  let lastCol = sh.getLastColumn()

  if (lastRow < PAYING_CFG.HEADER_ROW) {
    const headers = [PAYING_CFG.ID_HEADER, PAYING_CFG.NAME_HEADER, PAYING_CFG.EMAIL_HEADER]
      .concat(monthKeys.map(k => PAYING_monthLabelFromKey_(k)))
    sh.getRange(PAYING_CFG.HEADER_ROW, 1, 1, headers.length).setValues([headers])
    sh.setFrozenRows(PAYING_CFG.HEADER_ROW)
    lastRow = PAYING_CFG.HEADER_ROW
    lastCol = headers.length
  } else {
    const headerRow = sh
      .getRange(PAYING_CFG.HEADER_ROW, 1, 1, Math.max(lastCol, 3))
      .getValues()[0]
      .map(h => String(h || '').trim())

    const headerOk =
      headerRow[0] === PAYING_CFG.ID_HEADER &&
      headerRow[1] === PAYING_CFG.NAME_HEADER &&
      headerRow[2] === PAYING_CFG.EMAIL_HEADER

    if (!headerOk) {
      sh.getRange(PAYING_CFG.HEADER_ROW, 1, 1, 3).setValues([[
        PAYING_CFG.ID_HEADER,
        PAYING_CFG.NAME_HEADER,
        PAYING_CFG.EMAIL_HEADER
      ]])
      lastCol = Math.max(lastCol, 3)
    }

    const existingMonthKeys = new Set()
    headerRow.slice(3).forEach(h => {
      const key = PAYING_monthKeyFromHeader_(h)
      if (key) existingMonthKeys.add(key)
    })

    const missing = monthKeys.filter(k => !existingMonthKeys.has(k))
    if (missing.length) {
      const labels = missing.map(k => PAYING_monthLabelFromKey_(k))
      sh.getRange(PAYING_CFG.HEADER_ROW, lastCol + 1, 1, labels.length).setValues([labels])
      lastCol += labels.length
    }
  }

  const idCol = 1
  const nameCol = 2
  const emailCol = 3
  const monthCol = lastCol

  const rowMap = new Map()
  if (lastRow >= PAYING_CFG.DATA_START_ROW) {
    const numRows = lastRow - PAYING_CFG.HEADER_ROW
    const idValues = sh
      .getRange(PAYING_CFG.DATA_START_ROW, idCol, numRows, 1)
      .getValues()
    idValues.forEach((r, idx) => {
      const id = String(r[0] || '').trim()
      if (id) rowMap.set(id, PAYING_CFG.DATA_START_ROW + idx)
    })
  }

  const newRows = []
  orgMonthData.forEach((v, orgId) => {
    if (rowMap.has(orgId)) return
    const row = new Array(lastCol).fill('')
    row[idCol - 1] = orgId
    row[nameCol - 1] = v.org_name
    row[emailCol - 1] = v.org_email
    newRows.push(row)
    rowMap.set(orgId, lastRow + newRows.length)
  })

  if (newRows.length) {
    sh.getRange(lastRow + 1, 1, newRows.length, lastCol).setValues(newRows)
    lastRow += newRows.length
  }

  if (lastRow >= PAYING_CFG.DATA_START_ROW) {
    const numRows = lastRow - PAYING_CFG.HEADER_ROW
    const nameValues = sh.getRange(PAYING_CFG.DATA_START_ROW, nameCol, numRows, 1).getValues()
    const emailValues = sh.getRange(PAYING_CFG.DATA_START_ROW, emailCol, numRows, 1).getValues()

    const headerRow = sh
      .getRange(PAYING_CFG.HEADER_ROW, 1, 1, lastCol)
      .getValues()[0]

    const headerMonthKeys = headerRow.slice(3).map(h => PAYING_monthKeyFromHeader_(h))
    const monthValues = sh.getRange(PAYING_CFG.DATA_START_ROW, 4, numRows, headerMonthKeys.length).getValues()

    orgMonthData.forEach((v, orgId) => {
      const rowIdx = rowMap.get(orgId)
      if (!rowIdx) return
      const offset = rowIdx - PAYING_CFG.DATA_START_ROW
      if (offset < 0 || offset >= numRows) return
      if (v.org_name) nameValues[offset][0] = v.org_name
      if (v.org_email) emailValues[offset][0] = v.org_email

      for (let c = 0; c < headerMonthKeys.length; c++) {
        const key = headerMonthKeys[c]
        if (!key) continue
        const idx = monthIndexByKey.get(key)
        if (idx == null || idx < v.start_idx) continue
        if (!allowPastUpdate && key < currentMonthKey) continue
        monthValues[offset][c] = v.values[idx]
      }
    })

    orgNoActive.forEach((v, orgId) => {
      const rowIdx = rowMap.get(orgId)
      if (!rowIdx) return
      const offset = rowIdx - PAYING_CFG.DATA_START_ROW
      if (offset < 0 || offset >= numRows) return
      if (v.org_name) nameValues[offset][0] = v.org_name
      if (v.org_email) emailValues[offset][0] = v.org_email

      for (let c = 0; c < headerMonthKeys.length; c++) {
        const key = headerMonthKeys[c]
        if (!key) continue
        if (!allowPastUpdate && key < currentMonthKey) continue
        monthValues[offset][c] = 0
      }
    })

    sh.getRange(PAYING_CFG.DATA_START_ROW, nameCol, numRows, 1).setValues(nameValues)
    sh.getRange(PAYING_CFG.DATA_START_ROW, emailCol, numRows, 1).setValues(emailValues)
    sh.getRange(PAYING_CFG.DATA_START_ROW, 4, numRows, headerMonthKeys.length).setValues(monthValues)
    sh.getRange(PAYING_CFG.DATA_START_ROW, 4, numRows, headerMonthKeys.length).setNumberFormat('0')
  }

  sh.setFrozenColumns(3)
  sh.autoResizeColumns(1, Math.min(lastCol, 6))
  PAYING_applySummaryFormulas_(sh, lastCol)

  const seconds = (new Date() - t0) / 1000
  if (typeof writeSyncLog === 'function') {
    writeSyncLog('render_paying_users_snapshot', 'ok', orgMonthData.size, lastRow - 1, seconds, '')
  }
}

/* =========================
 * MRR logic
 * ========================= */

function PAYING_calcSubBaseMrr_(sub) {
  const amountMonthlyRaw = PAYING_num_(sub.amount_monthly)
  const amount = PAYING_num_(sub.amount)
  const interval = PAYING_str_(sub.interval).toLowerCase()
  const intervalCount = Math.max(1, PAYING_num_(sub.interval_count) || 1)

  if (amountMonthlyRaw) return amountMonthlyRaw

  let months = 1
  if (interval === 'month') months = intervalCount
  if (interval === 'year') months = intervalCount * 12
  return months ? (amount / months) : 0
}

function PAYING_discountFreeMonths_(sub) {
  const discountPercent = PAYING_num_(sub.discount_percent)
  const discountDuration = PAYING_str_(sub.discount_duration).toLowerCase()
  const discountMonthsRaw = PAYING_num_(sub.discount_duration_months)

  if (discountPercent !== 100) return { months: 0, forever: false }
  if (discountDuration === 'forever') return { months: 0, forever: true }

  let months = discountMonthsRaw
  if (!months && discountDuration === 'once') months = 1
  return { months: months || 0, forever: false }
}

function PAYING_monthDiff_(startKey, endKey) {
  const a = PAYING_parseMonthKey_(startKey)
  const b = PAYING_parseMonthKey_(endKey)
  if (!a || !b) return 0
  return (b.y - a.y) * 12 + (b.m - a.m)
}

/* =========================
 * Sheet helpers
 * ========================= */

function PAYING_readSheetObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1) return []

  const header = sheet
    .getRange(headerRow, 1, 1, lastCol)
    .getValues()[0]
    .map(h => String(h || '').trim())

  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()

  return data.map(r => {
    const obj = {}
    header.forEach((h, i) => {
      if (!h) return
      obj[PAYING_key_(h)] = r[i]
    })
    return obj
  })
}

function PAYING_ensureSummaryLayout_(sheet) {
  const headerRow = PAYING_CFG.HEADER_ROW
  const summaryRows = PAYING_CFG.SUMMARY_ROWS

  const row1 = sheet.getRange(1, 1, 1, 3).getValues()[0].map(v => String(v || '').trim())
  if (row1[0] === PAYING_CFG.ID_HEADER) {
    sheet.insertRowsBefore(1, headerRow - 1)
  }

  const labels = [['Target'], ['ARR'], ['% to target'], ['Month over month']]
  sheet.getRange(PAYING_CFG.SUMMARY_ROW_START, 1, summaryRows, 1).setValues(labels)

  const blankRow = PAYING_CFG.SUMMARY_ROW_START + summaryRows
  sheet.getRange(blankRow, 1, 1, 1).setValue('')
}

function PAYING_applySummaryFormulas_(sheet, lastCol) {
  const startCol = 4
  const numCols = lastCol - startCol + 1
  if (numCols <= 0) return

  const arrRow = PAYING_CFG.SUMMARY_ROW_START + 1
  for (let i = 0; i < numCols; i++) {
    const col = startCol + i
    const cell = sheet.getRange(arrRow, col)
    const existingFormula = cell.getFormula()
    const existingValue = cell.getValue()
    if (!existingFormula && (existingValue === '' || existingValue === null)) {
      const colLetter = PAYING_colLetter_(col)
      cell.setFormula(`=SUM(${colLetter}${PAYING_CFG.DATA_START_ROW}:${colLetter})`)
    }
  }

  const pctRow = PAYING_CFG.SUMMARY_ROW_START + 2
  const momRow = PAYING_CFG.SUMMARY_ROW_START + 3

  const pctFormulas = [[]]
  const momFormulas = [[]]

  for (let i = 0; i < numCols; i++) {
    const col = startCol + i
    const colLetter = PAYING_colLetter_(col)
    const targetRef = `${colLetter}${PAYING_CFG.SUMMARY_ROW_START}`
    const arrRef = `${colLetter}${arrRow}`

    pctFormulas[0].push(`=IF(${targetRef}="","",${arrRef}/${targetRef})`)

    if (i === 0) {
      momFormulas[0].push('')
    } else {
      const prevLetter = PAYING_colLetter_(col - 1)
      const prevArr = `${prevLetter}${arrRow}`
      momFormulas[0].push(`=IF(OR(${arrRef}="",${prevArr}=""),"",${arrRef}/${prevArr}-1)`)
    }
  }

  sheet.getRange(pctRow, startCol, 1, numCols).setFormulas([pctFormulas[0]])
  sheet.getRange(momRow, startCol, 1, numCols).setFormulas([momFormulas[0]])
  sheet.getRange(pctRow, startCol, 1, numCols).setNumberFormat('0.00%')
  sheet.getRange(momRow, startCol, 1, numCols).setNumberFormat('0.00%')
}

function PAYING_colLetter_(col) {
  let n = Number(col)
  let s = ''
  while (n > 0) {
    const mod = (n - 1) % 26
    s = String.fromCharCode(65 + mod) + s
    n = Math.floor((n - mod) / 26)
  }
  return s
}

function PAYING_key_(h) {
  return String(h || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_')
}

function PAYING_buildMembershipsByOrgId_(mems) {
  const out = new Map()
  ;(mems || []).forEach(m => {
    const orgId = PAYING_str_(m.org_id)
    if (!orgId) return

    const email = PAYING_str_(m.email)
    const emailKey = PAYING_normEmail_(m.email_key || email)
    const role = PAYING_str_(m.role).toLowerCase()
    const createdAt = PAYING_str_(m.created_at)

    if (!out.has(orgId)) out.set(orgId, [])
    out.get(orgId).push({ email, email_key: emailKey, role, created_at: createdAt })
  })
  return out
}

function PAYING_buildUsersByEmailKey_(users) {
  const out = new Map()
  ;(users || []).forEach(u => {
    const email = PAYING_str_(u.email)
    const emailKey = PAYING_normEmail_(u.email_key || email)
    if (!emailKey) return
    out.set(emailKey, u)
  })
  return out
}

function PAYING_buildSubIdsByOrgId_(membershipsByOrgId, userByEmailKey) {
  const out = new Map()
  membershipsByOrgId.forEach((members, orgId) => {
    const set = new Set()
    ;(members || []).forEach(m => {
      const key = PAYING_normEmail_(m.email_key || m.email)
      if (!key) return
      const u = userByEmailKey.get(key)
      const subId = PAYING_str_(u && (u.stripe_subscription_id || u.stripeSubscriptionId))
      if (subId) set.add(subId)
    })
    out.set(orgId, set)
  })
  return out
}

function PAYING_buildStripeBySubscriptionId_(subs) {
  const out = new Map()
  ;(subs || []).forEach(s => {
    const id = PAYING_str_(s.stripe_subscription_id || s.subscription_id || s.id)
    if (!id) return
    out.set(id, s)
  })
  return out
}

function PAYING_rowsFromSubIds_(subIdSet, stripeBySubId) {
  const rows = []
  ;(subIdSet || new Set()).forEach(id => {
    const row = stripeBySubId.get(id)
    if (row) rows.push(row)
  })
  return rows
}

function PAYING_pickOrgOwnerEmail_(members) {
  const arr = (members || []).slice()
  arr.sort((a, b) => {
    const ams = PAYING_toMs_(a.created_at) || 0
    const bms = PAYING_toMs_(b.created_at) || 0
    return ams - bms
  })

  const owners = arr.filter(m => (m.role || '').includes('owner'))
  if (owners.length && owners[0].email) return owners[0].email

  const admins = arr.filter(m => (m.role || '').includes('admin'))
  if (admins.length && admins[0].email) return admins[0].email

  if (arr.length && arr[0].email) return arr[0].email
  return ''
}

/* =========================
 * Date helpers
 * ========================= */

function PAYING_monthLabel_(d, fmt) {
  const s = Utilities.formatDate(d, Session.getScriptTimeZone(), fmt)
  return String(s || '').replace(' ', '-')
}

function PAYING_monthKeyFromDate_(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return ''
  const y = d.getUTCFullYear()
  const m = String(d.getUTCMonth() + 1).padStart(2, '0')
  return `${y}-${m}`
}

function PAYING_monthKeyFromIso_(iso) {
  const d = PAYING_parseDate_(iso)
  return d ? PAYING_monthKeyFromDate_(d) : ''
}

function PAYING_monthLabelFromKey_(key) {
  const info = PAYING_parseMonthKey_(key)
  if (!info) return ''
  const d = new Date(Date.UTC(info.y, info.m - 1, 1))
  return PAYING_monthLabel_(d, PAYING_CFG.MONTH_FMT)
}

function PAYING_monthKeyFromHeader_(v) {
  if (!v) return ''
  if (v instanceof Date && !isNaN(v.getTime())) return PAYING_monthKeyFromDate_(v)

  const s = String(v || '').trim()
  if (!s) return ''
  if (/^\d{4}-\d{2}$/.test(s)) return s

  const m = /^([A-Za-z]{3})[-\s](\d{4})$/.exec(s)
  if (!m) return ''

  const monthNum = PAYING_monthNumFromAbbrev_(m[1])
  if (!monthNum) return ''
  return `${m[2]}-${String(monthNum).padStart(2, '0')}`
}

function PAYING_parseMonthKey_(key) {
  const s = String(key || '').trim()
  if (!/^\d{4}-\d{2}$/.test(s)) return null
  const y = Number(s.slice(0, 4))
  const m = Number(s.slice(5, 7))
  if (!isFinite(y) || !isFinite(m)) return null
  return { y, m }
}

function PAYING_monthRange_(startKey, endKey) {
  const start = PAYING_parseMonthKey_(startKey)
  const end = PAYING_parseMonthKey_(endKey)
  if (!start || !end) return []

  const out = []
  let y = start.y
  let m = start.m
  while (y < end.y || (y === end.y && m <= end.m)) {
    out.push(`${y}-${String(m).padStart(2, '0')}`)
    m += 1
    if (m > 12) {
      m = 1
      y += 1
    }
  }
  return out
}

function PAYING_addMonths_(key, n) {
  const info = PAYING_parseMonthKey_(key)
  if (!info) return ''
  const d = new Date(Date.UTC(info.y, info.m - 1 + Number(n || 0), 1))
  return PAYING_monthKeyFromDate_(d)
}

function PAYING_subStartKey_(sub) {
  const startIso = PAYING_str_(sub.created_at || sub.current_period_start)
  return PAYING_monthKeyFromIso_(startIso)
}

function PAYING_parseDate_(v) {
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

/* =========================
 * Generic utils
 * ========================= */

function PAYING_str_(v) {
  if (v === null || v === undefined) return ''
  return String(v).trim()
}

function PAYING_num_(v) {
  const n = Number(v)
  return isFinite(n) ? n : 0
}

function PAYING_normEmail_(v) {
  const s = String(v || '').trim().toLowerCase()
  if (!s) return ''
  return s.replace(/\+[^@]+(?=@)/, '')
}

function PAYING_toMs_(iso) {
  const d = PAYING_parseDate_(iso)
  return d ? d.getTime() : 0
}

function PAYING_monthNumFromAbbrev_(abbr) {
  const s = String(abbr || '').trim().toLowerCase()
  const map = {
    jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
    jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12
  }
  return map[s] || 0
}
