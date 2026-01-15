/**************************************************************
 * Debug org member counts
 *
 * Compares:
 * - raw_clerk_orgs.members_count  (sheet value)
 * - computed count from raw_clerk_memberships (unique clerk_user_id per org_id)
 *
 * Outputs to:
 * - debug_org_member_counts
 *
 * Notes:
 * - Uses raw sheets only (no Clerk API calls) so it's fast + deterministic.
 **************************************************************/

const ORG_COUNT_DEBUG_CFG = {
  ORGS_SHEET: 'raw_clerk_orgs',
  MEMS_SHEET: 'raw_clerk_memberships',
  OUT_SHEET: 'debug_org_member_counts',
  HEADER_ROW: 1
}

function debug_org_member_counts() {
  const ss = SpreadsheetApp.getActive()

  const shOrgs = ss.getSheetByName(ORG_COUNT_DEBUG_CFG.ORGS_SHEET)
  const shMems = ss.getSheetByName(ORG_COUNT_DEBUG_CFG.MEMS_SHEET)
  if (!shOrgs) throw new Error(`Missing sheet: ${ORG_COUNT_DEBUG_CFG.ORGS_SHEET}`)
  if (!shMems) throw new Error(`Missing sheet: ${ORG_COUNT_DEBUG_CFG.MEMS_SHEET}`)

  // ---- Read orgs ----
  const orgs = dbg_readSheetObjects_(shOrgs, ORG_COUNT_DEBUG_CFG.HEADER_ROW)

  // ---- Read memberships & compute unique member counts per org ----
  const mems = dbg_readSheetObjects_(shMems, ORG_COUNT_DEBUG_CFG.HEADER_ROW)

  const membersByOrg = new Map() // org_id -> Set(clerk_user_id)
  mems.forEach(m => {
    const orgId = String(m.org_id || '').trim()
    if (!orgId) return

    const userId = String(m.clerk_user_id || '').trim()
    if (!userId) return

    if (!membersByOrg.has(orgId)) membersByOrg.set(orgId, new Set())
    membersByOrg.get(orgId).add(userId)
  })

  // ---- Build output ----
  const outHeaders = [
    'org_id',
    'org_name',
    'members_count_in_raw_orgs',
    'members_count_from_memberships',
    'delta (raw - computed)',
    'raw_missing?',
    'computed_missing?',
    'sample_member_emails (first 5)',
    'notes'
  ]

  // Build sample emails index for quick sanity
  const emailsByOrg = new Map() // org_id -> Set(email)
  mems.forEach(m => {
    const orgId = String(m.org_id || '').trim()
    if (!orgId) return
    const email = String(m.email || '').trim()
    if (!email) return
    if (!emailsByOrg.has(orgId)) emailsByOrg.set(orgId, new Set())
    emailsByOrg.get(orgId).add(email)
  })

  const rows = orgs.map(o => {
    const orgId = String(o.org_id || '').trim()
    const orgName = String(o.org_name || o.org_slug || '').trim()

    const rawCount = dbg_safeInt_(o.members_count)
    const computedCount = membersByOrg.has(orgId) ? membersByOrg.get(orgId).size : 0

    const delta = rawCount - computedCount

    const sampleEmails = emailsByOrg.has(orgId)
      ? Array.from(emailsByOrg.get(orgId)).slice(0, 5).join(', ')
      : ''

    let notes = ''
    if (!orgId) notes = 'missing org_id in raw_clerk_orgs row'
    else if (rawCount === 0 && computedCount > 0) notes = 'raw shows 0 but memberships has members (orgs pull likely missing members_count)'
    else if (rawCount > 0 && computedCount === 0) notes = 'raw shows >0 but memberships has 0 (memberships pull may be incomplete)'
    else if (delta !== 0) notes = 'counts disagree'

    return [
      orgId,
      orgName,
      rawCount,
      computedCount,
      delta,
      rawCount === 0 ? 'maybe' : '',
      computedCount === 0 ? 'maybe' : '',
      sampleEmails,
      notes
    ]
  })

  // Sort: biggest discrepancies first
  rows.sort((a, b) => Math.abs(b[4]) - Math.abs(a[4]))

  // ---- Write results ----
  const outSheet = dbg_getOrCreateSheet_(ss, ORG_COUNT_DEBUG_CFG.OUT_SHEET)
  outSheet.clearContents()
  outSheet.getRange(1, 1, 1, outHeaders.length).setValues([outHeaders])
  outSheet.setFrozenRows(1)

  if (rows.length) {
    outSheet.getRange(2, 1, rows.length, outHeaders.length).setValues(rows)
  }

  // Light formatting
  outSheet.autoResizeColumns(1, outHeaders.length)

  // Conditional formatting on delta
  const deltaCol = outHeaders.indexOf('delta (raw - computed)') + 1
  const lastRow = outSheet.getLastRow()
  if (lastRow >= 2) {
    const range = outSheet.getRange(2, deltaCol, lastRow - 1, 1)
    const rules = outSheet.getConditionalFormatRules() || []

    const keep = rules.filter(r => {
      try {
        return !(r.getRanges() || []).some(rr => rr.getColumn() === deltaCol)
      } catch (e) {
        return true
      }
    })

    const red = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([range])
      .whenNumberGreaterThan(0)
      .setBackground('#F8D7DA') // red-ish
      .build()

    const yellow = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([range])
      .whenNumberLessThan(0)
      .setBackground('#FFF3CD') // yellow-ish
      .build()

    const green = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([range])
      .whenNumberEqualTo(0)
      .setBackground('#D4EDDA') // green-ish
      .build()

    outSheet.setConditionalFormatRules(keep.concat([red, yellow, green]))
  }

  Logger.log(`Wrote ${rows.length} rows to ${ORG_COUNT_DEBUG_CFG.OUT_SHEET}`)
}

/* =========================
 * Helpers
 * ========================= */

function dbg_getOrCreateSheet_(ss, name) {
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function dbg_readSheetObjects_(sheet, headerRow) {
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
      obj[String(h).trim().toLowerCase().replace(/\s+/g, '_')] = r[i]
    })
    return obj
  })
}

function dbg_safeInt_(v) {
  const n = Number(v)
  if (!isFinite(n) || isNaN(n)) return 0
  return Math.floor(n)
}





function debug_clerk_org_payload_keys() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLERK_SECRET_KEY')
  if (!apiKey) throw new Error('Missing CLERK_SECRET_KEY')

  const url = 'https://api.clerk.com/v1/organizations?limit=3&offset=0'
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: `Bearer ${apiKey}` },
    muteHttpExceptions: true
  })

  const code = res.getResponseCode()
  const bodyText = res.getContentText()
  if (code >= 300) throw new Error(`Clerk API error ${code}: ${bodyText}`)

  const json = JSON.parse(bodyText)
  const orgs = Array.isArray(json) ? json : (Array.isArray(json.data) ? json.data : [])

  Logger.log(`Got ${orgs.length} orgs`)
  orgs.forEach((o, i) => {
    Logger.log(`--- ORG ${i + 1} ---`)
    Logger.log(`id=${o.id} name=${o.name}`)
    Logger.log(`keys=${Object.keys(o).join(', ')}`)
    Logger.log(`members_count=${o.members_count}`)
    Logger.log(`membersCount=${o.membersCount}`)
  })
}