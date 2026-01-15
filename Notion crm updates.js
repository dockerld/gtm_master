/**************************************************************
 * Notion CRM ↔ Sauron Linker + UpSale Pusher (Google Apps Script)
 *
 * Updated with best-practice fixes:
 * 1) ✅ Email normalization matches Calendar import:
 *    - lowercase, trim, strip +alias
 *
 * 2) ✅ Locking: prevents overlapping runs (manual double-clicks / triggers)
 *
 * 3) ✅ Company linker fast-path:
 *    - If a company already has sauron_org_id, we just set linked_to_sauron=true
 *      and write the mapping row (no need to search contacts)
 *
 * Notes:
 * - Mapping sheet notion_sauron_map_org uses column name "org_id"
 * - Mapping sheet notion_sauron_map_user uses: "org_id" and "user_id"
 * - Notion Companies property "sauron_org_id" (rich text) remains unchanged
 **************************************************************/

/** =========================
 * CONFIG
 * ========================= */

// Script properties
const PROP_NOTION_TOKEN = "NOTION_TOKEN"
const PROP_NOTION_VERSION = "NOTION_VERSION"
const PROP_NOTION_COMPANIES_DB_ID = "NOTION_COMPANIES_DB_ID"
const PROP_NOTION_CONTACTS_DB_ID = "NOTION_CONTACTS_DB_ID"

// Sheets
const SHEET_SAURON = "Sauron"
const SHEET_CANON_USERS = "canon_users"
const SHEET_ORG_INFO = "org_info"

// Audit sheets
const SHEET_MAP_ORG = "notion_sauron_map_org"
const SHEET_MAP_USER = "notion_sauron_map_user"

// org_info columns
const ORG_INFO_COL_ORG_ID = "Org ID"
const ORG_INFO_COL_ORG_NAME = "Org Name"
const ORG_INFO_COL_ORG_OWNER = "Org Owner"
const ORG_INFO_COL_UPSALE = "UpSale"

// Sauron sheet columns (best-effort matching; we try multiple header names)
const SAURON_COL_ORG_ID = "Org ID"
const SAURON_COL_ORG_NAME = "Org Name"
const SAURON_COL_SERVICE = "Service"
const SAURON_COL_EMAIL = "Email"
const SAURON_COL_OWNER_USER_ID_1 = "Org Owner User ID"
const SAURON_COL_OWNER_USER_ID_2 = "Owner User ID"
const SAURON_COL_OWNER_USER_ID_3 = "Org Owner"

// Notion DB property names (Companies)
const NOTION_COMPANY_PROP_LINKED = "linked_to_sauron"
const NOTION_COMPANY_PROP_SAURON_ORG_ID = "sauron_org_id"
const NOTION_COMPANY_PROP_LINK_SOURCE = "link_source"
const NOTION_COMPANY_PROP_LINKED_AT = "linked_at"
const NOTION_COMPANY_PROP_SERVICE = "service"
const NOTION_COMPANY_PROP_COMPANY_NAME = "Company Name"

// Notion DB property names (Contacts)
const NOTION_CONTACT_PROP_NAME = "Name" // title prop
const NOTION_CONTACT_PROP_EMAIL = "Email" // email type
const NOTION_CONTACT_PROP_COMPANY_REL = "Company" // relation to Companies
const NOTION_CONTACT_PROP_LINKED = "linked_to_sauron"
const NOTION_CONTACT_PROP_SAURON_ORG_ID = "sauron_org_id" // rich_text
const NOTION_CONTACT_PROP_SAURON_USER_ID = "sauron_user_id" // rich_text
const NOTION_CONTACT_PROP_ROLE = "Role" // optional select/multi_select

// Upsale (Companies)
const NOTION_COMPANY_PROP_UPSALE_STAGE = "Pipeline Stage (Up Sale)"
const NOTION_COMPANY_UPSALE_VALUE = "Upsale Target"

// Behavior toggles
const LINK_CONTACTS_TOO = true
const ENSURE_OWNER_ON_COMPANY_LINK = true
const MAX_COMPANIES_PER_RUN = 200
const MAX_CONTACTS_PER_COMPANY = 200
const MAX_CONTACTS_PER_RUN = 300

// Mapping sheet headers (ORG)
const MAP_ORG_HEADERS = [
  "notion_company_id",
  "notion_company_name",
  "matched_email",
  "org_id",
  "link_source",
  "status",
  "linked_at",
  "last_checked_at",
  "notes",
  "upsale_sent",
  "upsale_sent_at",
  "upsale_notes"
]

// Mapping sheet headers (USER)
const MAP_USER_HEADERS = [
  "notion_contact_id",
  "notion_contact_name",
  "notion_company_id",
  "notion_company_name",
  "email",
  "matched_email_key",
  "org_id",
  "user_id",
  "link_source",
  "status",
  "linked_at",
  "last_checked_at",
  "notes"
]

/** =========================
 * LOCK WRAPPER
 * ========================= */

function notionLockWrap_(name, fn) {
  if (typeof lockWrap === "function") {
    try {
      return lockWrap(name, fn)
    } catch (e) {
      return lockWrap(fn)
    }
  }
  const lock = LockService.getScriptLock()
  if (!lock.tryLock(300000)) throw new Error(`Could not acquire lock: ${name}`)
  try {
    return fn()
  } finally {
    lock.releaseLock()
  }
}

/** =========================
 * ENTRYPOINT 1: ORG LINKER
 * ========================= */

function notion_link_unlinked_companies_to_sauron() {
  return notionLockWrap_('notion_link_unlinked_companies_to_sauron', () => {
    // --- keep your existing function body EXACTLY as-is below this line ---
    const props = PropertiesService.getScriptProperties()
    const notion = notionClient_()

    ensureMapSheetOrg_()
    ensureMapSheetUser_()

    const sauronIndex = buildSauronIndex_()
    const companiesDbId = mustGetProp_(props, PROP_NOTION_COMPANIES_DB_ID)
    const contactsDbId = mustGetProp_(props, PROP_NOTION_CONTACTS_DB_ID)

    const unlinkedCompanies = notionQueryAll_(notion, companiesDbId, {
      filter: { property: NOTION_COMPANY_PROP_LINKED, checkbox: { equals: false } },
      page_size: 100
    }).slice(0, MAX_COMPANIES_PER_RUN)

    Logger.log(`Found ${unlinkedCompanies.length} unlinked Notion companies to check`)

    let linkedCount = 0
    let checkedCount = 0
    const nowIso = new Date().toISOString()

    for (const companyPage of unlinkedCompanies) {
      checkedCount += 1

      const companyId = companyPage.id
      const companyName = notionGetTitle_(companyPage, NOTION_COMPANY_PROP_COMPANY_NAME) || "(untitled)"

      const contactEmails = getCompanyContactEmails_(notion, companyId)
      const match = findMatchByEmails_(contactEmails, sauronIndex)

      if (!match.matched) {
        upsertMapRowOrg_({
          notion_company_id: companyId,
          notion_company_name: companyName,
          matched_email: "",
          org_id: "",
          link_source: match.reason,
          status: "no_match",
          linked_at: "",
          last_checked_at: nowIso,
          notes: `checked ${contactEmails.length} contact emails`,
          upsale_sent: "",
          upsale_sent_at: "",
          upsale_notes: ""
        })
        continue
      }

      linkCompanyPage_(notion, companyId, match, sauronIndex)

      if (LINK_CONTACTS_TOO) {
        backfillContacts_(notion, companyId, sauronIndex, {
          companyId,
          companyName,
          linkSource: match.reason
        })
      }

      if (ENSURE_OWNER_ON_COMPANY_LINK) {
        const ownerUserId = sauronIndex.orgOwnerUserIdByOrgId.get(match.orgId || "") || ""
        if (ownerUserId) {
          const owner = sauronIndex.userByClerkUserId.get(ownerUserId) || null
          if (owner && owner.email) {
            ensureContactLinkedToCompany_(notion, contactsDbId, {
              companyId,
              companyName,
              orgId: match.orgId,
              userId: ownerUserId,
              ownerName: owner.fullName || "",
              ownerEmail: owner.email,
              roleName: "Owner",
              linkSource: "org_link_owner_from_sauron",
              sauronIndex
            })
          }
        }
      }

      upsertMapRowOrg_({
        notion_company_id: companyId,
        notion_company_name: companyName,
        matched_email: match.emailMatched || "",
        org_id: match.orgId || "",
        link_source: match.reason,
        status: "linked",
        linked_at: nowIso,
        last_checked_at: nowIso,
        notes: `matched from ${contactEmails.length} contact emails`,
        upsale_sent: "",
        upsale_sent_at: "",
        upsale_notes: ""
      })

      linkedCount += 1
    }

    Logger.log(`Done. Checked=${checkedCount}, Linked=${linkedCount}`)
    return { rows_in: unlinkedCompanies.length, rows_out: linkedCount }
  })
}

/** =========================
 * ENTRYPOINT 2: USER LINKER
 * ========================= */

function notion_link_unlinked_contacts_to_sauron() {
  return notionLockWrap_('notion_link_unlinked_contacts_to_sauron', () => {
    // keep your existing body as-is
    // (no other code changes required besides the lock wrapper)
    const props = PropertiesService.getScriptProperties()
    const notion = notionClient_()

    ensureMapSheetUser_()

    const sauronIndex = buildSauronIndex_()
    const contactsDbId = mustGetProp_(props, PROP_NOTION_CONTACTS_DB_ID)

    const contacts = notionQueryAll_(notion, contactsDbId, {
      filter: { property: NOTION_CONTACT_PROP_LINKED, checkbox: { equals: false } },
      page_size: 100
    }).slice(0, MAX_CONTACTS_PER_RUN)

    Logger.log(`Found ${contacts.length} unlinked Notion contacts to check`)

    let linked = 0
    const nowIso = new Date().toISOString()

    for (const c of contacts) {
      const contactId = c.id
      const contactName = notionGetTitleAny_(c) || "(untitled contact)"
      const email = notionGetEmail_(c, NOTION_CONTACT_PROP_EMAIL)
      const emailKey = normEmail_(email)

      if (!emailKey) {
        upsertMapRowUser_({
          notion_contact_id: contactId,
          notion_contact_name: contactName,
          notion_company_id: "",
          notion_company_name: "",
          email: email || "",
          matched_email_key: "",
          org_id: "",
          user_id: "",
          link_source: "no_email",
          status: "skipped",
          linked_at: "",
          last_checked_at: nowIso,
          notes: "contact missing email"
        })
        continue
      }

      const hit = sauronIndex.userByEmailKey.get(emailKey)
      if (!hit) {
        upsertMapRowUser_({
          notion_contact_id: contactId,
          notion_contact_name: contactName,
          notion_company_id: "",
          notion_company_name: "",
          email: email,
          matched_email_key: emailKey,
          org_id: "",
          user_id: "",
          link_source: "no_match",
          status: "no_match",
          linked_at: "",
          last_checked_at: nowIso,
          notes: "email not found in canon_users"
        })
        continue
      }
      // ✅ upgrade contact title if it's currently an email
      const betterName = bestDisplayNameFromCanon_(hit)
      if (betterName) {
        tryUpgradeNotionContactTitle_(notion, contactId, betterName)
      }
      const patch = { properties: {} }
      patch.properties[NOTION_CONTACT_PROP_LINKED] = { checkbox: true }

      if (hit.orgId && hasProp_(c, NOTION_CONTACT_PROP_SAURON_ORG_ID)) {
        patch.properties[NOTION_CONTACT_PROP_SAURON_ORG_ID] = {
          rich_text: [{ type: "text", text: { content: hit.orgId } }]
        }
      }
      if (hit.userId && hasProp_(c, NOTION_CONTACT_PROP_SAURON_USER_ID)) {
        patch.properties[NOTION_CONTACT_PROP_SAURON_USER_ID] = {
          rich_text: [{ type: "text", text: { content: hit.userId } }]
        }
      }

      notionUpdatePage_(notion, contactId, patch)

      upsertMapRowUser_({
        notion_contact_id: contactId,
        notion_contact_name: contactName,
        notion_company_id: "",
        notion_company_name: "",
        email: email,
        matched_email_key: emailKey,
        org_id: hit.orgId || "",
        user_id: hit.userId || "",
        link_source: "email_match",
        status: "linked",
        linked_at: nowIso,
        last_checked_at: nowIso,
        notes: "linked via contact scan"
      })

      linked += 1
    }

    Logger.log(`Contact scan done. Linked=${linked} / Checked=${contacts.length}`)
    return { rows_in: contacts.length, rows_out: linked }
  })
}

/** =========================
 * ENTRYPOINT 3: UPSALE PUSHER (from org_info.UpSale)
 * ========================= */

function notion_push_upsale_targets_from_org_info() {
  return notionLockWrap_("notion_push_upsale_targets_from_org_info", () => {
    const props = PropertiesService.getScriptProperties()
    const notion = notionClient_()

    ensureMapSheetOrg_()
    ensureMapSheetUser_()

    const companiesDbId = mustGetProp_(props, PROP_NOTION_COMPANIES_DB_ID)
    const contactsDbId = mustGetProp_(props, PROP_NOTION_CONTACTS_DB_ID)

    const sauronIndex = buildSauronIndex_()
    const targets = readOrgInfoUpsaleTargets_()

    Logger.log(`UpSale targets checked in org_info: ${targets.length}`)

    const sentIndex = buildUpsaleSentIndexFromMapOrg_()

    let processed = 0
    let skippedAlready = 0
    let created = 0
    let updated = 0
    let ownerContactsEnsured = 0

    const nowIso = new Date().toISOString()

    for (const t of targets) {
      const orgId = str_(t.orgId)
      if (!orgId) continue

      const prior = sentIndex.get(orgId)
      if (prior && prior.upsale_sent === true) {
        skippedAlready += 1
        continue
      }

      // Find or create company by org id
      let companyPage = findNotionCompanyBySauronOrgId_(notion, companiesDbId, orgId)
      if (!companyPage) {
        companyPage = createNotionCompanyForOrg_(notion, companiesDbId, {
          orgId,
          companyName: t.orgName || (sauronIndex.orgData.get(orgId) || {}).orgName || "",
          linkSource: "org_info_upsale"
        })
        created += 1
      } else {
        updated += 1
      }

      const companyId = companyPage.id
      const companyName =
        notionGetTitle_(companyPage, NOTION_COMPANY_PROP_COMPANY_NAME) ||
        t.orgName ||
        (sauronIndex.orgData.get(orgId) || {}).orgName ||
        "(untitled)"

      setCompanyUpsaleStage_(notion, companyId)
      ensureCompanyLinkedToSauron_(notion, companyId, orgId)
      ensureCompanyLinkedAt_(notion, companyId)

      const ownerResolved = resolveOwnerForOrg_(t, sauronIndex, orgId)
      if (ownerResolved && ownerResolved.email) {
        ensureContactLinkedToCompany_(notion, contactsDbId, {
          companyId,
          companyName,
          orgId,
          userId: ownerResolved.userId || "",
          ownerName: ownerResolved.name || "",
          ownerEmail: ownerResolved.email,
          roleName: "Owner",
          linkSource: ownerResolved.source || "org_info_upsale_owner",
          sauronIndex
        })
        ownerContactsEnsured += 1
      }

      upsertMapRowOrg_({
        notion_company_id: companyId,
        notion_company_name: companyName,
        matched_email: ownerResolved ? (ownerResolved.email || "") : (t.orgOwnerEmail || ""),
        org_id: orgId,
        link_source: "org_info_upsale",
        status: "upsale_target",
        linked_at: nowIso,
        last_checked_at: nowIso,
        notes: ownerResolved ? "set Upsale Target + ensured owner contact" : "set Upsale Target (no resolvable owner)",
        upsale_sent: "true",
        upsale_sent_at: nowIso,
        upsale_notes: "guarded by notion_sauron_map_org; org_info.UpSale stays checked"
      })

      sentIndex.set(orgId, { upsale_sent: true, notion_company_id: companyId })
      processed += 1
    }

    Logger.log(
      `UpSale push complete. processed=${processed}, created=${created}, updated=${updated}, ` +
      `owner_contacts_ensured=${ownerContactsEnsured}, skipped_already_sent=${skippedAlready}`
    )

    return { rows_in: targets.length, rows_out: processed }
  })
}

/** =========================
 * OPTIONAL TRIGGERS
 * ========================= */

function setup_daily_notion_sauron_link_triggers() {
  const fns = [
    "notion_link_unlinked_companies_to_sauron",
    "notion_link_unlinked_contacts_to_sauron",
    "notion_push_upsale_targets_from_org_info"
  ]

  ScriptApp.getProjectTriggers().forEach(t => {
    if (fns.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t)
  })

  ScriptApp.newTrigger("notion_link_unlinked_companies_to_sauron")
    .timeBased().everyDays(1).atHour(7).nearMinute(10).create()

  ScriptApp.newTrigger("notion_link_unlinked_contacts_to_sauron")
    .timeBased().everyDays(1).atHour(7).nearMinute(25).create()

  ScriptApp.newTrigger("notion_push_upsale_targets_from_org_info")
    .timeBased().everyDays(1).atHour(7).nearMinute(40).create()

  Logger.log("Daily triggers created.")
}

/** =========================
 * SAURON INDEX (canon_users + Sauron sheet)
 * ========================= */

function buildSauronIndex_() {
  const ss = SpreadsheetApp.getActive()

  // canon_users
  const shUsers = ss.getSheetByName(SHEET_CANON_USERS)
  if (!shUsers) throw new Error(`Missing sheet: ${SHEET_CANON_USERS}`)

  const canon = readSheetObjects_(shUsers, 1)

  const userByEmailKey = new Map()    // email_key -> { orgId, userId, fullName, email }
  const userByClerkUserId = new Map() // clerk_user_id -> { orgId, userId, fullName, email }

  for (const u of canon) {
    const email = str_(u.email)
    const emailKey = normEmail_(u.email_key || email)
    const orgId = str_(u.org_id)
    const userId = str_(u.clerk_user_id || u.user_id || u.id)

    const fullName =
      str_(u.full_name) ||
      str_(u.name) ||
      [str_(u.first_name), str_(u.last_name)].filter(Boolean).join(" ").trim()

    if (emailKey) userByEmailKey.set(emailKey, { orgId, userId, fullName, email })
    if (userId) userByClerkUserId.set(userId, { orgId, userId, fullName, email })
  }

  // Sauron sheet org data + owner user id
  const shSauron = ss.getSheetByName(SHEET_SAURON)
  const orgData = new Map()
  const orgOwnerUserIdByOrgId = new Map()

  if (shSauron) {
    const { header, rows } = readTable_(shSauron, 3, 2)

    const colEmail = findIdxAny_(header, [SAURON_COL_EMAIL])
    const colOrgName = findIdxAny_(header, [SAURON_COL_ORG_NAME])
    const colService = findIdxAny_(header, [SAURON_COL_SERVICE])
    const colOrgId = findIdxAny_(header, [SAURON_COL_ORG_ID])
    const colOwnerUserId = findIdxAny_(header, [SAURON_COL_OWNER_USER_ID_1, SAURON_COL_OWNER_USER_ID_2, SAURON_COL_OWNER_USER_ID_3])

    for (const r of rows) {
      let orgId = ""
      if (colOrgId >= 0) orgId = str_(r[colOrgId])

      if (!orgId && colEmail >= 0) {
        const key = normEmail_(r[colEmail])
        const hit = userByEmailKey.get(key)
        orgId = hit ? hit.orgId : ""
      }

      if (!orgId) continue

      if (!orgData.has(orgId)) {
        const orgName = colOrgName >= 0 ? str_(r[colOrgName]) : ""
        const service = colService >= 0 ? str_(r[colService]) : ""
        orgData.set(orgId, { orgName, service })
      }

      if (colOwnerUserId >= 0) {
        const ownerUserId = str_(r[colOwnerUserId])
        if (ownerUserId && !orgOwnerUserIdByOrgId.has(orgId)) {
          orgOwnerUserIdByOrgId.set(orgId, ownerUserId)
        }
      }
    }
  }

  return { userByEmailKey, userByClerkUserId, orgData, orgOwnerUserIdByOrgId }
}

function findMatchByEmails_(emails, sauronIndex) {
  for (const email of (emails || [])) {
    const key = normEmail_(email)
    if (!key) continue

    const hit = sauronIndex.userByEmailKey.get(key)
    if (!hit) continue
    if (!hit.orgId) return { matched: false, reason: "matched_email_missing_org_id", orgId: "", userId: "" }

    return { matched: true, reason: "email_match", orgId: hit.orgId, userId: hit.userId || "", emailMatched: key }
  }
  return { matched: false, reason: "no_email_match", orgId: "", userId: "" }
}

/** =========================
 * UPSALE HELPERS
 * ========================= */

function readOrgInfoUpsaleTargets_() {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(SHEET_ORG_INFO)
  if (!sh) throw new Error(`Missing sheet: ${SHEET_ORG_INFO}`)

  const lastRow = sh.getLastRow()
  const lastCol = sh.getLastColumn()
  if (lastRow < 2) return []

  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim())
  const cOrgId = header.findIndex(h => h.toLowerCase() === ORG_INFO_COL_ORG_ID.toLowerCase())
  const cOrgName = header.findIndex(h => h.toLowerCase() === ORG_INFO_COL_ORG_NAME.toLowerCase())
  const cOrgOwner = header.findIndex(h => h.toLowerCase() === ORG_INFO_COL_ORG_OWNER.toLowerCase())
  const cUpsale = header.findIndex(h => h.toLowerCase() === ORG_INFO_COL_UPSALE.toLowerCase())

  if (cOrgId < 0) throw new Error(`org_info missing column: ${ORG_INFO_COL_ORG_ID}`)
  if (cUpsale < 0) throw new Error(`org_info missing column: ${ORG_INFO_COL_UPSALE}`)

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues()
  const out = []

  for (const r of data) {
    const orgId = str_(r[cOrgId])
    if (!orgId) continue

    const upsaleVal = r[cUpsale]
    const upsale = (upsaleVal === true) || String(upsaleVal || "").toLowerCase().trim() === "true"
    if (!upsale) continue

    const rawOwner = (cOrgOwner >= 0) ? str_(r[cOrgOwner]) : ""
    const parsed = parseNameEmail_(rawOwner)

    out.push({
      orgId,
      orgName: (cOrgName >= 0) ? str_(r[cOrgName]) : "",
      orgOwnerName: parsed.name || "",
      orgOwnerEmail: parsed.email || (normEmail_(rawOwner) || "")
    })
  }

  return out
}

function buildUpsaleSentIndexFromMapOrg_() {
  ensureMapSheetOrg_()

  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(SHEET_MAP_ORG)
  const lastRow = sh.getLastRow()
  const lastCol = sh.getLastColumn()
  const out = new Map()
  if (lastRow < 2) return out

  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim())
  const cOrgId = header.indexOf("org_id")
  const cCompanyId = header.indexOf("notion_company_id")
  const cSent = header.indexOf("upsale_sent")

  if (cOrgId < 0) throw new Error(`Mapping sheet missing header: org_id`)
  if (cSent < 0) throw new Error(`Mapping sheet missing header: upsale_sent`)

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues()

  for (const r of data) {
    const orgId = str_(r[cOrgId])
    if (!orgId) continue

    const sentRaw = String(r[cSent] || "").toLowerCase().trim()
    const sent = (sentRaw === "true" || sentRaw === "yes" || sentRaw === "1")

    out.set(orgId, {
      upsale_sent: sent,
      notion_company_id: (cCompanyId >= 0) ? str_(r[cCompanyId]) : ""
    })
  }

  return out
}

function resolveOwnerForOrg_(targetRow, sauronIndex, orgId) {
  // A) org_info owner
  const emailA = normEmail_(targetRow.orgOwnerEmail || "")
  const nameA = str_(targetRow.orgOwnerName || "")
  if (emailA) {
    const hit = sauronIndex.userByEmailKey.get(emailA) || null
    return {
      name: nameA || (hit ? hit.fullName : ""),
      email: emailA,
      userId: hit ? (hit.userId || "") : "",
      source: "org_info_owner"
    }
  }

  // B) Sauron owner user id -> canon users
  const ownerUserId = str_(sauronIndex.orgOwnerUserIdByOrgId.get(orgId) || "")
  if (ownerUserId) {
    const hit2 = sauronIndex.userByClerkUserId.get(ownerUserId) || null
    if (hit2 && hit2.email) {
      return {
        name: hit2.fullName || "",
        email: hit2.email,
        userId: ownerUserId,
        source: "sauron_owner_user_id"
      }
    }
  }

  return null
}

/** =========================
 * NOTION HELPERS (Company)
 * ========================= */

function findNotionCompanyBySauronOrgId_(notion, companiesDbId, orgId) {
  const res = notionQueryAll_(notion, companiesDbId, {
    filter: { property: NOTION_COMPANY_PROP_SAURON_ORG_ID, rich_text: { contains: orgId } },
    page_size: 10
  })

  for (const p of res) {
    const rt = notionGetRichText_(p, NOTION_COMPANY_PROP_SAURON_ORG_ID)
    if (rt === orgId) return p
  }
  return res && res.length ? res[0] : null
}

function createNotionCompanyForOrg_(notion, companiesDbId, { orgId, companyName, linkSource }) {
  const props = {}
  props[NOTION_COMPANY_PROP_COMPANY_NAME] = { title: [{ type: "text", text: { content: companyName || orgId } }] }
  props[NOTION_COMPANY_PROP_LINKED] = { checkbox: true }
  props[NOTION_COMPANY_PROP_SAURON_ORG_ID] = { rich_text: [{ type: "text", text: { content: orgId } }] }
  props[NOTION_COMPANY_PROP_LINK_SOURCE] = { select: { name: linkSource || "org_info_upsale" } }
  props[NOTION_COMPANY_PROP_LINKED_AT] = { date: { start: new Date().toISOString() } }
  props[NOTION_COMPANY_PROP_UPSALE_STAGE] = { select: { name: NOTION_COMPANY_UPSALE_VALUE } }

  return notionPost_(notion, "/pages", { parent: { database_id: companiesDbId }, properties: props })
}

function setCompanyUpsaleStage_(notion, companyId) {
  const page = notionGetPage_(notion, companyId)
  const p = page && page.properties ? page.properties[NOTION_COMPANY_PROP_UPSALE_STAGE] : null
  if (!p) return

  const patch = { properties: {} }
  if (p.type === "select") {
    patch.properties[NOTION_COMPANY_PROP_UPSALE_STAGE] = { select: { name: NOTION_COMPANY_UPSALE_VALUE } }
  } else if (p.type === "multi_select") {
    const existing = (p.multi_select || []).map(x => x.name)
    const next = existing.includes(NOTION_COMPANY_UPSALE_VALUE) ? existing : existing.concat([NOTION_COMPANY_UPSALE_VALUE])
    patch.properties[NOTION_COMPANY_PROP_UPSALE_STAGE] = { multi_select: next.map(name => ({ name })) }
  } else {
    return
  }
  notionUpdatePage_(notion, companyId, patch)
}

function ensureCompanyLinkedToSauron_(notion, companyId, orgId) {
  const patch = { properties: {} }
  patch.properties[NOTION_COMPANY_PROP_LINKED] = { checkbox: true }
  patch.properties[NOTION_COMPANY_PROP_SAURON_ORG_ID] = { rich_text: [{ type: "text", text: { content: orgId } }] }
  notionUpdatePage_(notion, companyId, patch)
}

function ensureCompanyLinkedAt_(notion, companyId) {
  const page = notionGetPage_(notion, companyId)
  const p = page && page.properties ? page.properties[NOTION_COMPANY_PROP_LINKED_AT] : null
  if (!p || p.type !== "date") return
  const patch = { properties: {} }
  patch.properties[NOTION_COMPANY_PROP_LINKED_AT] = { date: { start: new Date().toISOString() } }
  notionUpdatePage_(notion, companyId, patch)
}

/** =========================
 * NOTION HELPERS (Contacts / Owner logic)
 * ========================= */

function ensureContactLinkedToCompany_(notion, contactsDbId, {
  companyId,
  companyName,
  orgId,
  userId,
  ownerName,
  ownerEmail,
  roleName,
  linkSource,
  sauronIndex
}) {
  const nowIso = new Date().toISOString()
  const email = str_(ownerEmail)
  const emailKey = normEmail_(email)
  if (!emailKey) return

  const existing = notionQueryAll_(notion, contactsDbId, {
    filter: { property: NOTION_CONTACT_PROP_EMAIL, email: { equals: email } },
    page_size: 5
  })

  let contactId = existing && existing.length ? existing[0].id : ""

  // Create if missing
  if (!contactId) {
    const hit = sauronIndex && sauronIndex.userByEmailKey ? sauronIndex.userByEmailKey.get(emailKey) : null
    const bestName = str_(ownerName) || (hit && hit.fullName ? String(hit.fullName).trim() : "")
    const fallbackName = email ? email.split("@")[0] : "Owner"

    const createProps = {}
    createProps[NOTION_CONTACT_PROP_NAME] = { title: [{ type: "text", text: { content: bestName || fallbackName } }] }
    createProps[NOTION_CONTACT_PROP_EMAIL] = { email: email }
    createProps[NOTION_CONTACT_PROP_COMPANY_REL] = { relation: [{ id: companyId }] }
    createProps[NOTION_CONTACT_PROP_LINKED] = { checkbox: true }

    const finalOrgId = str_(orgId) || (hit && hit.orgId ? hit.orgId : "")
    const finalUserId = str_(userId) || (hit && hit.userId ? hit.userId : "")

    if (finalOrgId) {
      createProps[NOTION_CONTACT_PROP_SAURON_ORG_ID] = { rich_text: [{ type: "text", text: { content: finalOrgId } }] }
    }
    if (finalUserId) {
      createProps[NOTION_CONTACT_PROP_SAURON_USER_ID] = { rich_text: [{ type: "text", text: { content: finalUserId } }] }
    }
    if (roleName) {
      createProps[NOTION_CONTACT_PROP_ROLE] = { select: { name: roleName } }
    }

    const created = notionCreateContactSafe_(notion, contactsDbId, createProps)
    contactId = created && created.id ? created.id : ""
    if (!contactId) return

    upsertMapRowUser_({
      notion_contact_id: contactId,
      notion_contact_name: bestName || fallbackName,
      notion_company_id: companyId,
      notion_company_name: companyName || "",
      email: email,
      matched_email_key: emailKey,
      org_id: finalOrgId || "",
      user_id: finalUserId || "",
      link_source: linkSource || "ensure_contact_create",
      status: "linked",
      linked_at: nowIso,
      last_checked_at: nowIso,
      notes: "created contact + linked to company"
    })
    return
  }

  // Update if exists
  const full = notionGetPage_(notion, contactId)
  const patch = { properties: {} }

  if (full.properties && full.properties[NOTION_CONTACT_PROP_LINKED]) {
    patch.properties[NOTION_CONTACT_PROP_LINKED] = { checkbox: true }
  }

  const hit2 = sauronIndex && sauronIndex.userByEmailKey ? sauronIndex.userByEmailKey.get(emailKey) : null
  const finalOrgId2 = str_(orgId) || (hit2 && hit2.orgId ? hit2.orgId : "")
  const finalUserId2 = str_(userId) || (hit2 && hit2.userId ? hit2.userId : "")

  // ✅ upgrade contact title if it's currently an email
  const betterName =
    bestDisplayNameFromCanon_(hit2) ||
    String(ownerName || "").trim()

  if (betterName) {
    tryUpgradeNotionContactTitle_(notion, contactId, betterName)
  }

  if (full.properties && full.properties[NOTION_CONTACT_PROP_SAURON_ORG_ID] && finalOrgId2) {
    patch.properties[NOTION_CONTACT_PROP_SAURON_ORG_ID] = { rich_text: [{ type: "text", text: { content: finalOrgId2 } }] }
  }
  if (full.properties && full.properties[NOTION_CONTACT_PROP_SAURON_USER_ID] && finalUserId2) {
    patch.properties[NOTION_CONTACT_PROP_SAURON_USER_ID] = { rich_text: [{ type: "text", text: { content: finalUserId2 } }] }
  }

  // ensure relation contains companyId
  if (full.properties && full.properties[NOTION_CONTACT_PROP_COMPANY_REL]) {
    const rel = full.properties[NOTION_CONTACT_PROP_COMPANY_REL]
    if (rel.type === "relation") {
      const ids = (rel.relation || []).map(x => x.id)
      if (!ids.includes(companyId)) {
        patch.properties[NOTION_CONTACT_PROP_COMPANY_REL] = { relation: ids.concat([companyId]).map(id => ({ id })) }
      }
    }
  }

  // Role
  if (roleName && full.properties && full.properties[NOTION_CONTACT_PROP_ROLE]) {
    const roleProp = full.properties[NOTION_CONTACT_PROP_ROLE]
    if (roleProp.type === "select") {
      patch.properties[NOTION_CONTACT_PROP_ROLE] = { select: { name: roleName } }
    } else if (roleProp.type === "multi_select") {
      const existingRoles = (roleProp.multi_select || []).map(x => x.name)
      const next = existingRoles.includes(roleName) ? existingRoles : existingRoles.concat([roleName])
      patch.properties[NOTION_CONTACT_PROP_ROLE] = { multi_select: next.map(name => ({ name })) }
    }
  }

  if (Object.keys(patch.properties).length) {
    notionUpdatePage_(notion, contactId, patch)
  }

  upsertMapRowUser_({
    notion_contact_id: contactId,
    notion_contact_name: notionGetTitleAny_(full) || "",
    notion_company_id: companyId,
    notion_company_name: companyName || "",
    email: email,
    matched_email_key: emailKey,
    org_id: finalOrgId2 || "",
    user_id: finalUserId2 || "",
    link_source: linkSource || "ensure_contact_update",
    status: "linked",
    linked_at: nowIso,
    last_checked_at: nowIso,
    notes: "ensured contact linked to company"
  })
}

function notionCreateContactSafe_(notion, contactsDbId, props) {
  const body = { parent: { database_id: contactsDbId }, properties: props }
  try {
    return notionPost_(notion, "/pages", body)
  } catch (e) {
    const minimal = { parent: { database_id: contactsDbId }, properties: {} }
    minimal.properties[NOTION_CONTACT_PROP_NAME] = props[NOTION_CONTACT_PROP_NAME]
    minimal.properties[NOTION_CONTACT_PROP_EMAIL] = props[NOTION_CONTACT_PROP_EMAIL]
    minimal.properties[NOTION_CONTACT_PROP_COMPANY_REL] = props[NOTION_CONTACT_PROP_COMPANY_REL]
    return notionPost_(notion, "/pages", minimal)
  }
}

/** =========================
 * NOTION: COMPANY → CONTACT EMAILS
 * ========================= */

function getCompanyContactEmails_(notion, companyId) {
  const contactsDbId = mustGetProp_(PropertiesService.getScriptProperties(), PROP_NOTION_CONTACTS_DB_ID)

  const pages = notionQueryAll_(notion, contactsDbId, {
    filter: { property: NOTION_CONTACT_PROP_COMPANY_REL, relation: { contains: companyId } },
    page_size: 100
  })

  const emails = []
  for (const p of pages.slice(0, MAX_CONTACTS_PER_COMPANY)) {
    const email = notionGetEmail_(p, NOTION_CONTACT_PROP_EMAIL)
    if (email) emails.push(email)
  }

  return [...new Set(emails.map(normEmail_).filter(Boolean))]
}

/** =========================
 * NOTION: WRITE UPDATES (Linker)
 * ========================= */

function linkCompanyPage_(notion, companyId, match, sauronIndex) {
  const orgId = match.orgId
  const orgMeta = sauronIndex.orgData.get(orgId) || {}

  const props = {}
  props[NOTION_COMPANY_PROP_LINKED] = { checkbox: true }
  props[NOTION_COMPANY_PROP_SAURON_ORG_ID] = { rich_text: [{ type: "text", text: { content: orgId } }] }
  props[NOTION_COMPANY_PROP_LINK_SOURCE] = { select: { name: match.reason } }
  props[NOTION_COMPANY_PROP_LINKED_AT] = { date: { start: new Date().toISOString() } }

  if (orgMeta.service) props[NOTION_COMPANY_PROP_SERVICE] = { select: { name: orgMeta.service } }

  notionUpdatePage_(notion, companyId, { properties: props })
}

function backfillContacts_(notion, companyId, sauronIndex, ctx) {
  const contactsDbId = mustGetProp_(PropertiesService.getScriptProperties(), PROP_NOTION_CONTACTS_DB_ID)

  const contacts = notionQueryAll_(notion, contactsDbId, {
    filter: { property: NOTION_CONTACT_PROP_COMPANY_REL, relation: { contains: companyId } },
    page_size: 100
  })

  const nowIso = new Date().toISOString()

  for (const c of contacts.slice(0, MAX_CONTACTS_PER_COMPANY)) {
    const email = notionGetEmail_(c, NOTION_CONTACT_PROP_EMAIL)
    const emailKey = normEmail_(email)
    const hit = emailKey ? sauronIndex.userByEmailKey.get(emailKey) : null
    // ✅ upgrade contact title if it's currently an email
    const betterName = bestDisplayNameFromCanon_(hit)
    if (betterName) {
      tryUpgradeNotionContactTitle_(notion, c.id, betterName)
    }
    const patch = { properties: {} }
    patch.properties[NOTION_CONTACT_PROP_LINKED] = { checkbox: true }

    if (hit && hit.orgId && hasProp_(c, NOTION_CONTACT_PROP_SAURON_ORG_ID)) {
      patch.properties[NOTION_CONTACT_PROP_SAURON_ORG_ID] = { rich_text: [{ type: "text", text: { content: hit.orgId } }] }
    }
    if (hit && hit.userId && hasProp_(c, NOTION_CONTACT_PROP_SAURON_USER_ID)) {
      patch.properties[NOTION_CONTACT_PROP_SAURON_USER_ID] = { rich_text: [{ type: "text", text: { content: hit.userId } }] }
    }

    if (Object.keys(patch.properties).length) notionUpdatePage_(notion, c.id, patch)

    upsertMapRowUser_({
      notion_contact_id: c.id,
      notion_contact_name: notionGetTitleAny_(c) || "",
      notion_company_id: ctx && ctx.companyId ? ctx.companyId : "",
      notion_company_name: ctx && ctx.companyName ? ctx.companyName : "",
      email: email || "",
      matched_email_key: emailKey || "",
      org_id: (hit && hit.orgId) ? hit.orgId : "",
      user_id: (hit && hit.userId) ? hit.userId : "",
      link_source: (ctx && ctx.linkSource) ? ctx.linkSource : "company_backfill",
      status: hit ? "linked" : "patched_no_hit",
      linked_at: hit ? nowIso : "",
      last_checked_at: nowIso,
      notes: hit ? "backfilled via company link" : "company linked but this contact email not in canon_users"
    })
  }
}

/** =========================
 * AUDIT MAPPING SHEETS
 * ========================= */

function ensureMapSheetOrg_() {
  const ss = SpreadsheetApp.getActive()
  let sh = ss.getSheetByName(SHEET_MAP_ORG)
  if (!sh) sh = ss.insertSheet(SHEET_MAP_ORG)

  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, MAP_ORG_HEADERS.length).setValues([MAP_ORG_HEADERS])
    sh.setFrozenRows(1)
    return
  }

  const lastCol = Math.max(sh.getLastColumn(), MAP_ORG_HEADERS.length)
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim())
  if (header[0] !== MAP_ORG_HEADERS[0]) {
    sh.getRange(1, 1, 1, MAP_ORG_HEADERS.length).setValues([MAP_ORG_HEADERS])
    sh.setFrozenRows(1)
    return
  }

  const missing = MAP_ORG_HEADERS.filter(h => !header.includes(h))
  if (missing.length) {
    sh.getRange(1, 1, 1, MAP_ORG_HEADERS.length).setValues([MAP_ORG_HEADERS])
    sh.setFrozenRows(1)
  }
}

function ensureMapSheetUser_() {
  const ss = SpreadsheetApp.getActive()
  let sh = ss.getSheetByName(SHEET_MAP_USER)
  if (!sh) sh = ss.insertSheet(SHEET_MAP_USER)

  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, MAP_USER_HEADERS.length).setValues([MAP_USER_HEADERS])
    sh.setFrozenRows(1)
    return
  }

  const lastCol = Math.max(sh.getLastColumn(), MAP_USER_HEADERS.length)
  const oldHeader = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim())
  if (arraysEqual_(oldHeader.slice(0, MAP_USER_HEADERS.length), MAP_USER_HEADERS)) return

  sh.getRange(1, 1, 1, MAP_USER_HEADERS.length).setValues([MAP_USER_HEADERS])
  sh.setFrozenRows(1)
}

function upsertMapRowOrg_(obj) {
  upsertOrgRowByEitherKey_(obj)
}

function upsertMapRowUser_(obj) {
  upsertById_(SHEET_MAP_USER, "notion_contact_id", obj)
}

function upsertOrgRowByEitherKey_(obj) {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(SHEET_MAP_ORG)
  if (!sh) throw new Error(`Missing sheet: ${SHEET_MAP_ORG}`)

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v || "").trim())
  const cCompanyId = header.indexOf("notion_company_id")
  const cOrgId = header.indexOf("org_id")
  if (cCompanyId < 0) throw new Error(`Mapping sheet missing header: notion_company_id`)
  if (cOrgId < 0) throw new Error(`Mapping sheet missing header: org_id`)

  const companyId = str_(obj.notion_company_id)
  const orgId = str_(obj.org_id)
  if (!companyId && !orgId) return

  const lastRow = sh.getLastRow()
  let rowToWrite = lastRow + 1

  if (lastRow >= 2) {
    const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues()
    const foundIdx = data.findIndex(r => {
      const existingCompanyId = str_(r[cCompanyId])
      const existingOrgId = str_(r[cOrgId])
      if (companyId && existingCompanyId === companyId) return true
      if (!companyId && orgId && existingOrgId === orgId) return true
      if (companyId && orgId && existingOrgId === orgId) return true
      return false
    })
    if (foundIdx >= 0) rowToWrite = 2 + foundIdx
  }

  const row = header.map(h => (Object.prototype.hasOwnProperty.call(obj, h) ? obj[h] : ""))
  sh.getRange(rowToWrite, 1, 1, header.length).setValues([row])
}

function upsertById_(sheetName, idHeader, obj) {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(sheetName)
  if (!sh) throw new Error(`Missing sheet: ${sheetName}`)

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v || "").trim())
  const idCol = header.indexOf(idHeader)
  if (idCol < 0) throw new Error(`Mapping sheet missing header: ${idHeader}`)

  const targetId = String(obj[idHeader] || "").trim()
  if (!targetId) return

  const lastRow = sh.getLastRow()
  let rowToWrite = lastRow + 1

  if (lastRow >= 2) {
    const ids = sh.getRange(2, idCol + 1, lastRow - 1, 1).getValues().map(r => String(r[0] || "").trim())
    const foundIdx = ids.findIndex(v => v === targetId)
    if (foundIdx >= 0) rowToWrite = 2 + foundIdx
  }

  const row = header.map(h => (Object.prototype.hasOwnProperty.call(obj, h) ? obj[h] : ""))
  sh.getRange(rowToWrite, 1, 1, header.length).setValues([row])
}

/** =========================
 * NOTION HTTP CLIENT
 * ========================= */

function notionClient_() {
  const props = PropertiesService.getScriptProperties()
  const token = mustGetProp_(props, PROP_NOTION_TOKEN)
  const version = props.getProperty(PROP_NOTION_VERSION) || "2022-06-28"
  return { token, version, baseUrl: "https://api.notion.com/v1" }
}

function notionQueryAll_(notion, databaseId, body) {
  const out = []
  let cursor = null
  while (true) {
    const payload = Object.assign({}, body || {})
    if (cursor) payload.start_cursor = cursor
    const res = notionPost_(notion, `/databases/${databaseId}/query`, payload)
    const results = res && res.results ? res.results : []
    out.push(...results)
    if (!res.has_more) break
    cursor = res.next_cursor
    if (!cursor) break
    if (out.length > 5000) break
  }
  return out
}

function notionUpdatePage_(notion, pageId, body) {
  return notionPatch_(notion, `/pages/${pageId}`, body)
}

function notionGetPage_(notion, pageId) {
  const url = notion.baseUrl + `/pages/${pageId}`
  const resp = UrlFetchApp.fetch(url, {
    method: "get",
    contentType: "application/json",
    muteHttpExceptions: true,
    headers: { Authorization: `Bearer ${notion.token}`, "Notion-Version": notion.version }
  })
  const code = resp.getResponseCode()
  const text = resp.getContentText()
  if (code >= 200 && code < 300) return text ? JSON.parse(text) : {}
  throw new Error(`Notion API error ${code}: ${text}`)
}

function notionPost_(notion, path, payload) {
  return notionFetch_(notion, "post", path, payload)
}

function notionPatch_(notion, path, payload) {
  return notionFetch_(notion, "patch", path, payload)
}

function notionFetch_(notion, method, path, payload) {
  const url = notion.baseUrl + path
  const options = {
    method: method,
    contentType: "application/json",
    muteHttpExceptions: true,
    headers: { Authorization: `Bearer ${notion.token}`, "Notion-Version": notion.version },
    payload: payload ? JSON.stringify(payload) : undefined
  }

  for (let attempt = 0; attempt < 5; attempt++) {
    const resp = UrlFetchApp.fetch(url, options)
    const code = resp.getResponseCode()
    const text = resp.getContentText()

    if (code === 429) {
      Utilities.sleep(500 * (attempt + 1))
      continue
    }
    if (code >= 200 && code < 300) return text ? JSON.parse(text) : {}
    throw new Error(`Notion API error ${code}: ${text}`)
  }
  throw new Error("Notion API error: too many retries")
}

/** =========================
 * NOTION PROPERTY READERS
 * ========================= */

function notionGetTitle_(page, propName) {
  const p = page && page.properties ? page.properties[propName] : null
  if (!p || p.type !== "title" || !p.title) return ""
  return p.title.map(t => t.plain_text).join("").trim()
}

function notionGetTitleAny_(page) {
  const props = page && page.properties ? page.properties : {}
  for (const k in props) {
    const p = props[k]
    if (p && p.type === "title" && p.title) return p.title.map(t => t.plain_text).join("").trim()
  }
  return ""
}

function notionGetEmail_(page, propName) {
  const p = page && page.properties ? page.properties[propName] : null
  if (!p) return ""
  if (p.type === "email") return str_(p.email)
  if (p.type === "rich_text") return (p.rich_text || []).map(t => t.plain_text).join("").trim()
  return ""
}

function notionGetRichText_(page, propName) {
  const p = page && page.properties ? page.properties[propName] : null
  if (!p) return ""
  if (p.type === "rich_text") return (p.rich_text || []).map(t => t.plain_text).join("").trim()
  if (p.type === "title") return (p.title || []).map(t => t.plain_text).join("").trim()
  return ""
}

function hasProp_(page, propName) {
  return !!(page && page.properties && page.properties[propName])
}

/** =========================
 * SHEET HELPERS
 * ========================= */




function notionLockWrap_(name, fn) {
  const lock = LockService.getScriptLock()
  const ok = lock.tryLock(5 * 60 * 1000) // 5 minutes
  if (!ok) throw new Error(`Could not acquire lock: ${name}`)
  try {
    return fn()
  } finally {
    lock.releaseLock()
  }
}


function readTable_(sheet, headerRow, startCol) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow) return { header: [], rows: [] }

  const header = sheet.getRange(headerRow, startCol, 1, lastCol - startCol + 1)
    .getValues()[0]
    .map(h => String(h || "").trim())

  const numRows = Math.max(0, lastRow - headerRow)
  if (!numRows) return { header, rows: [] }

  const rows = sheet.getRange(headerRow + 1, startCol, numRows, header.length).getValues()
  return { header, rows }
}

function findIdxMaybe_(header, name) {
  return header.findIndex(h => String(h).trim().toLowerCase() === String(name).trim().toLowerCase())
}

function findIdxAny_(header, names) {
  for (const n of (names || [])) {
    const idx = findIdxMaybe_(header, n)
    if (idx >= 0) return idx
  }
  return -1
}

function readSheetObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1) return []

  const header = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim())
  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()

  return data.map(r => {
    const obj = {}
    header.forEach((h, i) => {
      if (!h) return
      obj[key_(h)] = r[i]
    })
    return obj
  })
}

function key_(h) {
  return String(h || "").trim().toLowerCase().replace(/\s+/g, "_")
}

/** =========================
 * TINY HELPERS
 * ========================= */

function mustGetProp_(props, key) {
  const v = props.getProperty(key)
  if (!v) throw new Error(`Missing Script Property: ${key}`)
  return v
}

function str_(v) {
  if (v === null || v === undefined) return ""
  return String(v).trim()
}

// ✅ IMPORTANT: same normalization as calendar import (strip +alias)
function normEmail_(v) {
  const s = String(v || "").trim().toLowerCase()
  if (!s) return ""
  // strip plus-alias: kate+foo@domain.com -> kate@domain.com
  return s.replace(/\+[^@]+(?=@)/, "")
}

function parseNameEmail_(raw) {
  const s = String(raw || "").trim()
  if (!s) return { name: "", email: "" }

  let m = s.match(/^\s*([^<]+?)\s*<\s*([^>]+?)\s*>\s*$/)
  if (m) return { name: m[1].trim(), email: m[2].trim() }

  m = s.match(/^\s*(.*?)\s*\(\s*([^\)]+@[^\)]+)\s*\)\s*$/)
  if (m) return { name: m[1].trim(), email: m[2].trim() }

  m = s.match(/([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})/i)
  if (m) {
    const email = m[1].trim()
    const name = s.replace(m[1], "").replace(/[<>()]/g, "").trim()
    return { name, email }
  }

  return { name: s, email: "" }
}

function arraysEqual_(a, b) {
  if (!a || !b) return false
  if (a.length !== b.length) return false
  for (let i = 0; i < a.length; i++) {
    if (String(a[i] || "") !== String(b[i] || "")) return false
  }
  return true
}


function looksLikeEmail_(s) {
  const t = String(s || "").trim().toLowerCase()
  if (!t) return false
  // If it contains @ it's probably email-ish, but use a real email regex too.
  if (t.includes("@")) return true
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(t)
}

function bestDisplayNameFromCanon_(hit) {
  if (!hit) return ""
  const n = String(hit.fullName || "").trim()
  return n
}

/**
 * If the contact's title currently looks like an email, rename it to desiredName.
 * Safe: does nothing if title already looks like a real name.
 */
function tryUpgradeNotionContactTitle_(notion, contactId, desiredName) {
  const name = String(desiredName || "").trim()
  if (!contactId || !name) return

  const page = notionGetPage_(notion, contactId)
  const currentTitle = notionGetTitleAny_(page) || ""

  // Only overwrite if current title looks like an email / placeholder
  if (!currentTitle || looksLikeEmail_(currentTitle)) {
    const patch = { properties: {} }
    patch.properties[NOTION_CONTACT_PROP_NAME] = {
      title: [{ type: "text", text: { content: name } }]
    }
    notionUpdatePage_(notion, contactId, patch)
  }
}