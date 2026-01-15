/**************************************************************
 * Calendar → Notion CRM Import (Google Apps Script) — FULL FILE (IMPROVED)
 *
 * What it does:
 * - Looks BACK only (past N days) on Camden’s calendar
 * - Collects external attendee emails (normalized)
 * - Upserts Notion Contacts by Email (Email is PRIMARY KEY)
 * - Resolves org_id from canon_users (truth)
 * - Resolves org_name from org_info (manual truth)
 * - Finds/creates Notion Company:
 *    - If org_id exists: find by Companies.sauron_org_id (rich_text) else create with real org name
 *    - Else: find/create Unknown Company (domain.com) (business domains only)
 * - Ensures Contact relates to Company (Company relation)
 * - Writes audit rows to:
 *    - notion_sauron_map_org (upsert by org_id or notion_company_id)
 *    - notion_sauron_map_user (upsert by notion_contact_id)
 * - Optional: runs your linker jobs afterwards (if functions exist)
 *
 * Safe to run multiple times per day:
 * - Contacts upsert by Email
 * - Companies upsert by sauron_org_id (if known) else by domain bucket
 * - Map sheets upsert by stable keys
 *
 * Script Properties required:
 * - NOTION_TOKEN
 * - NOTION_COMPANIES_DB_ID
 * - NOTION_CONTACTS_DB_ID
 *
 * Optional Script Properties:
 * - NOTION_VERSION (defaults to "2022-06-28")
 * - CALENDAR_ID (defaults to "camden@pingassistant.com")
 **************************************************************/

/** =========================
 * CONFIG
 * ========================= */

const CALCRM_PROP_CALENDAR_ID = "CALENDAR_ID"

// Past-only window
const CALCRM_LOOKBACK_DAYS = 2

// Limits
const CALCRM_MAX_EVENTS_PER_RUN = 50
const CALCRM_MAX_UNIQUE_EMAILS_PER_RUN = 150

// Sheets (truth sources)
const CALCRM_SHEET_CANON_USERS = "canon_users" // email_key/email -> org_id
const CALCRM_SHEET_ORG_INFO = "org_info"       // Org ID -> Org Name (manual)

// Mapping sheets (audit)
const CALCRM_SHEET_MAP_ORG = "notion_sauron_map_org"
const CALCRM_SHEET_MAP_USER = "notion_sauron_map_user"

// org_info column names (do NOT change org_info sheet)
const CALCRM_ORG_INFO_COL_ORG_ID = "Org ID"
const CALCRM_ORG_INFO_COL_ORG_NAME = "Org Name"

// Notion Companies DB props
const CALCRM_NOTION_COMPANY_TITLE = "Company Name"           // title
const CALCRM_NOTION_COMPANY_SAURON_ORG_ID = "sauron_org_id"  // rich_text
const CALCRM_NOTION_COMPANY_DOMAIN = "domain"                // rich_text OPTIONAL (auto-fallback if missing)

// Notion Contacts DB props
const CALCRM_NOTION_CONTACT_TITLE = "Name"                  // title
const CALCRM_NOTION_CONTACT_EMAIL = "Email"                 // email type
const CALCRM_NOTION_CONTACT_COMPANY_REL = "Company"         // relation to Companies

// Internal email filtering
const CALCRM_INTERNAL_EMAIL_DOMAINS = [
  "pingassistant.com",
  "ping-assistant.com"
]

// Personal domains: if not in canon_users, do NOT create a company bucket
const CALCRM_PERSONAL_DOMAINS = new Set([
  "gmail.com", "googlemail.com",
  "yahoo.com", "hotmail.com", "outlook.com", "live.com",
  "icloud.com", "me.com", "mac.com",
  "aol.com", "proton.me", "protonmail.com"
])

// Mapping sheet headers (matches your newer schema expectations)
const CALCRM_MAP_ORG_HEADERS = [
  "notion_company_id",
  "notion_company_name",
  "org_id",
  "matched_email",
  "link_source",
  "status",
  "linked_at",
  "last_checked_at",
  "notes",
  "upsale_sent",
  "upsale_sent_at",
  "upsale_notes"
]

const CALCRM_MAP_USER_HEADERS = [
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

// Optional: auto-run linker jobs after calendar import
const CALCRM_RUN_LINKERS_AFTER_IMPORT = true

/** =========================
 * ENTRYPOINT
 * ========================= */

function calcrm_notion_calendar_import_from_camden() {
  const props = PropertiesService.getScriptProperties()
  const notion = calcrm_notionClient_()

  const companiesDbId = calcrm_mustGetProp_(props, "NOTION_COMPANIES_DB_ID")
  const contactsDbId  = calcrm_mustGetProp_(props, "NOTION_CONTACTS_DB_ID")

  // Ensure audit sheets exist (and have correct headers)
  calcrm_ensureMapSheet_(CALCRM_SHEET_MAP_ORG, CALCRM_MAP_ORG_HEADERS)
  calcrm_ensureMapSheet_(CALCRM_SHEET_MAP_USER, CALCRM_MAP_USER_HEADERS)

  const calendarId = props.getProperty(CALCRM_PROP_CALENDAR_ID) || "camden@pingassistant.com"
  const cal = CalendarApp.getCalendarById(calendarId)
  if (!cal) throw new Error(`Calendar not found: ${calendarId}`)

  // Build indices
  const canon = calcrm_buildCanonEmailToOrgIdIndex_() // email_key -> org_id
  const orgInfo = calcrm_buildOrgIdToOrgNameIndex_()  // org_id -> org_name

  // Past-only window
  const now = new Date()
  const start = new Date(now.getTime() - CALCRM_LOOKBACK_DAYS * 24 * 60 * 60 * 1000)
  const end = now

  const events = cal.getEvents(start, end) || []
  const slice = events.slice(0, CALCRM_MAX_EVENTS_PER_RUN)

  Logger.log(`Calendar import: events in window=${events.length}, processing=${slice.length}`)
  Logger.log(`Calendar import window: start=${start.toISOString()} end=${end.toISOString()}`)

  // Collect unique external attendee emails (normalized)
  const emailSet = new Set()

  for (const ev of slice) {
    const guests = ev.getGuestList(true) || []
    for (const g of guests) {
      const emailKey = calcrm_normEmail_(g.getEmail())
      if (!emailKey) continue
      if (calcrm_isInternalEmail_(emailKey)) continue
      emailSet.add(emailKey)
      if (emailSet.size >= CALCRM_MAX_UNIQUE_EMAILS_PER_RUN) break
    }
    if (emailSet.size >= CALCRM_MAX_UNIQUE_EMAILS_PER_RUN) break
  }

  const emails = Array.from(emailSet)
  Logger.log(`Calendar import: unique external attendee emails=${emails.length}`)

  // Stats
  let createdContacts = 0
  let updatedContacts = 0
  let createdCompanies = 0
  let foundCompanies = 0

  // Caches
  const companyByOrgId = new Map()   // org_id -> company page
  const companyByDomain = new Map()  // domain -> company page

  // If domain property is missing on Companies DB, we will stop trying to set it
  let domainPropertyExists = true

  const nowIso = new Date().toISOString()

  for (const emailKey of emails) {
    // Resolve org_id (canon_users)
    const orgId = canon.emailToOrgId.get(emailKey) || ""
    const domain = calcrm_getDomain_(emailKey)

    const shouldUseDomainCompany = (!orgId && domain && !CALCRM_PERSONAL_DOMAINS.has(domain))

    // Ensure company
    let companyPage = null
    let desiredCompanyName = ""

    if (orgId) {
      // ✅ Use org_info for the real name
      desiredCompanyName = orgInfo.orgIdToName.get(orgId) || ""
      if (!desiredCompanyName) desiredCompanyName = `Unknown Company (${domain || "linked"})`

      companyPage = companyByOrgId.get(orgId) || null
      if (!companyPage) {
        const found = calcrm_findCompanyBySauronOrgId_(notion, companiesDbId, orgId)
        if (found) {
          companyPage = found
          foundCompanies += 1
        } else {
          companyPage = calcrm_createCompanySafe_(notion, companiesDbId, {
            title: desiredCompanyName,
            sauronOrgId: orgId,
            domain: (domain && !CALCRM_PERSONAL_DOMAINS.has(domain)) ? domain : "",
            domainPropertyExists
          })
          if (companyPage && companyPage._domain_prop_missing === true) domainPropertyExists = false
          createdCompanies += 1
        }

        // ✅ If found/created but still has placeholder name, rename to the real org_info name
        calcrm_tryRenameCompanyIfUnknown_(notion, companyPage, desiredCompanyName)
        companyByOrgId.set(orgId, companyPage)
      } else {
        calcrm_tryRenameCompanyIfUnknown_(notion, companyPage, desiredCompanyName)
      }
    } else if (shouldUseDomainCompany) {
      desiredCompanyName = `Unknown Company (${domain})`

      companyPage = companyByDomain.get(domain) || null
      if (!companyPage) {
        const found = domainPropertyExists
          ? calcrm_findCompanyByDomainProp_(notion, companiesDbId, domain)
          : calcrm_findCompanyByDomainTitle_(notion, companiesDbId, domain)

        if (found) {
          companyPage = found
          foundCompanies += 1
        } else {
          companyPage = calcrm_createCompanySafe_(notion, companiesDbId, {
            title: desiredCompanyName,
            sauronOrgId: "",
            domain,
            domainPropertyExists
          })
          if (companyPage && companyPage._domain_prop_missing === true) domainPropertyExists = false
          createdCompanies += 1
        }

        companyByDomain.set(domain, companyPage)
      }
    }

    // Upsert contact by email
    const existingContact = calcrm_findContactByEmail_(notion, contactsDbId, emailKey)

    let contactPage = existingContact
    let contactName = ""

    if (existingContact) {
      contactName = calcrm_notionGetTitleAny_(existingContact) || emailKey
      if (companyPage && companyPage.id) {
        calcrm_ensureContactRelatedToCompany_(notion, existingContact, companyPage.id)
      }
      updatedContacts += 1
    } else {
      const payload = calcrm_buildContactCreatePayload_({
        contactsDbId,
        name: emailKey,
        email: emailKey,
        companyId: companyPage ? companyPage.id : ""
      })

      contactPage = calcrm_notionPost_(notion, "/pages", payload)
      contactName = emailKey
      if (contactPage && contactPage.id) createdContacts += 1
    }

    // Audit mapping sheets
    if (companyPage && companyPage.id) {
      const title = calcrm_notionGetTitle_(companyPage, CALCRM_NOTION_COMPANY_TITLE) || desiredCompanyName || "(untitled)"
      calcrm_upsertMapOrgRow_({
        notion_company_id: companyPage.id,
        notion_company_name: title,
        org_id: orgId || "",
        matched_email: emailKey,
        link_source: orgId ? "calendar_canon_match" : "calendar_domain_match",
        status: orgId ? "linked_to_org" : "domain_bucket",
        linked_at: nowIso,
        last_checked_at: nowIso,
        notes: orgId ? "org_id from canon_users; org name from org_info" : "no org_id; domain bucket",
        upsale_sent: "",
        upsale_sent_at: "",
        upsale_notes: ""
      })
    }

    if (contactPage && contactPage.id) {
      const companyIdForUser = (companyPage && companyPage.id) ? companyPage.id : ""
      const companyTitleForUser = (companyPage && companyPage.id)
        ? (calcrm_notionGetTitle_(companyPage, CALCRM_NOTION_COMPANY_TITLE) || desiredCompanyName || "")
        : ""

      calcrm_upsertMapUserRow_({
        notion_contact_id: contactPage.id,
        notion_contact_name: contactName || emailKey,
        notion_company_id: companyIdForUser,
        notion_company_name: companyTitleForUser,
        email: emailKey,
        matched_email_key: emailKey,
        org_id: orgId || "",
        user_id: "", // calendar job doesn't set user_id; linker can fill later
        link_source: orgId ? "calendar_canon_match" : (shouldUseDomainCompany ? "calendar_domain_match" : "calendar_no_company"),
        status: orgId ? "linked_to_org" : (shouldUseDomainCompany ? "domain_bucket" : "no_company"),
        linked_at: nowIso,
        last_checked_at: nowIso,
        notes: orgId ? "matched to canon_users.org_id" : "no canon match"
      })
    }
  }

  Logger.log(
    `Calendar import done. ` +
    `companies(created=${createdCompanies}, found=${foundCompanies}), ` +
    `contacts(created=${createdContacts}, updated_or_found=${updatedContacts})`
  )

  // ✅ Optional: run linker jobs after import (safe, and will “finish” connections)
  if (CALCRM_RUN_LINKERS_AFTER_IMPORT) {
    calcrm_runLinkersBestEffort_()
  }

  return {
    emails: emails.length,
    created_contacts: createdContacts,
    updated_contacts: updatedContacts,
    created_companies: createdCompanies,
    found_companies: foundCompanies
  }
}

/** =========================
 * Build indices
 * ========================= */

function calcrm_buildCanonEmailToOrgIdIndex_() {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(CALCRM_SHEET_CANON_USERS)
  if (!sh) throw new Error(`Missing sheet: ${CALCRM_SHEET_CANON_USERS}`)

  const lastRow = sh.getLastRow()
  const lastCol = sh.getLastColumn()
  const map = new Map()
  if (lastRow < 2) return { emailToOrgId: map }

  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim().toLowerCase())
  const cEmailKey = header.indexOf("email_key")
  const cEmail = header.indexOf("email")
  const cOrgId = header.indexOf("org_id")

  if (cOrgId < 0) throw new Error(`canon_users missing header: org_id`)

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues()
  for (const r of data) {
    const emailKey = (cEmailKey >= 0) ? calcrm_normEmail_(r[cEmailKey]) : ""
    const email = (cEmail >= 0) ? calcrm_normEmail_(r[cEmail]) : ""
    const key = emailKey || email
    if (!key) continue

    const orgId = String(r[cOrgId] || "").trim()
    if (!orgId) continue

    map.set(key, orgId)
  }

  return { emailToOrgId: map }
}

function calcrm_buildOrgIdToOrgNameIndex_() {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(CALCRM_SHEET_ORG_INFO)
  const out = new Map()
  if (!sh || sh.getLastRow() < 2) return { orgIdToName: out }

  const lastRow = sh.getLastRow()
  const lastCol = sh.getLastColumn()
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim())

  const cOrgId = header.findIndex(h => String(h).trim().toLowerCase() === CALCRM_ORG_INFO_COL_ORG_ID.toLowerCase())
  const cOrgName = header.findIndex(h => String(h).trim().toLowerCase() === CALCRM_ORG_INFO_COL_ORG_NAME.toLowerCase())

  if (cOrgId < 0) throw new Error(`org_info missing column: ${CALCRM_ORG_INFO_COL_ORG_ID}`)
  if (cOrgName < 0) throw new Error(`org_info missing column: ${CALCRM_ORG_INFO_COL_ORG_NAME}`)

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues()
  for (const r of data) {
    const orgId = String(r[cOrgId] || "").trim()
    if (!orgId) continue
    const orgName = String(r[cOrgName] || "").trim()
    if (!orgName) continue
    out.set(orgId, orgName)
  }

  return { orgIdToName: out }
}

/** =========================
 * Notion finders
 * ========================= */

function calcrm_findContactByEmail_(notion, contactsDbId, email) {
  const res = calcrm_notionQueryAll_(notion, contactsDbId, {
    filter: { property: CALCRM_NOTION_CONTACT_EMAIL, email: { equals: email } },
    page_size: 5
  })
  return (res && res.length) ? res[0] : null
}

function calcrm_findCompanyBySauronOrgId_(notion, companiesDbId, orgId) {
  const res = calcrm_notionQueryAll_(notion, companiesDbId, {
    filter: { property: CALCRM_NOTION_COMPANY_SAURON_ORG_ID, rich_text: { contains: orgId } },
    page_size: 10
  })

  for (const p of res) {
    const rt = calcrm_notionGetRichText_(p, CALCRM_NOTION_COMPANY_SAURON_ORG_ID)
    if (String(rt || "").trim() === String(orgId || "").trim()) return p
  }

  return (res && res.length) ? res[0] : null
}

function calcrm_findCompanyByDomainProp_(notion, companiesDbId, domain) {
  const res = calcrm_notionQueryAll_(notion, companiesDbId, {
    filter: { property: CALCRM_NOTION_COMPANY_DOMAIN, rich_text: { contains: domain } },
    page_size: 10
  })

  for (const p of res) {
    const rt = calcrm_notionGetRichText_(p, CALCRM_NOTION_COMPANY_DOMAIN)
    if (calcrm_safeLower_(rt) === calcrm_safeLower_(domain)) return p
  }

  return (res && res.length) ? res[0] : null
}

function calcrm_findCompanyByDomainTitle_(notion, companiesDbId, domain) {
  const token = `(${domain})`
  const res = calcrm_notionQueryAll_(notion, companiesDbId, {
    filter: { property: CALCRM_NOTION_COMPANY_TITLE, title: { contains: token } },
    page_size: 10
  })
  return (res && res.length) ? res[0] : null
}

/** =========================
 * Notion writers
 * ========================= */

function calcrm_createCompanySafe_(notion, companiesDbId, { title, sauronOrgId, domain, domainPropertyExists }) {
  const includeDomain = domainPropertyExists === true

  try {
    return calcrm_createCompany_(notion, companiesDbId, { title, sauronOrgId, domain: includeDomain ? domain : "" })
  } catch (e) {
    const msg = String(e && e.message ? e.message : e)
    if (msg.includes("domain is not a property that exists")) {
      const page = calcrm_createCompany_(notion, companiesDbId, { title, sauronOrgId, domain: "" })
      page._domain_prop_missing = true
      return page
    }
    throw e
  }
}

function calcrm_createCompany_(notion, companiesDbId, { title, sauronOrgId, domain }) {
  const props = {}

  props[CALCRM_NOTION_COMPANY_TITLE] = {
    title: [{ type: "text", text: { content: title || "Unknown Company" } }]
  }

  if (sauronOrgId) {
    props[CALCRM_NOTION_COMPANY_SAURON_ORG_ID] = {
      rich_text: [{ type: "text", text: { content: sauronOrgId } }]
    }
  }

  if (domain) {
    props[CALCRM_NOTION_COMPANY_DOMAIN] = {
      rich_text: [{ type: "text", text: { content: domain } }]
    }
  }

  return calcrm_notionPost_(notion, "/pages", {
    parent: { database_id: companiesDbId },
    properties: props
  })
}

function calcrm_tryRenameCompanyIfUnknown_(notion, companyPage, desiredName) {
  try {
    if (!companyPage || !companyPage.id) return
    const wants = String(desiredName || "").trim()
    if (!wants) return

    const currentTitle = calcrm_notionGetTitle_(companyPage, CALCRM_NOTION_COMPANY_TITLE) || ""
    if (!currentTitle.toLowerCase().startsWith("unknown company")) return

    const patch = { properties: {} }
    patch.properties[CALCRM_NOTION_COMPANY_TITLE] = {
      title: [{ type: "text", text: { content: wants } }]
    }
    calcrm_notionPatch_(notion, `/pages/${companyPage.id}`, patch)
  } catch (e) {
    // non-fatal
  }
}

function calcrm_buildContactCreatePayload_({ contactsDbId, name, email, companyId }) {
  const props = {}

  props[CALCRM_NOTION_CONTACT_TITLE] = {
    title: [{ type: "text", text: { content: name || email } }]
  }

  props[CALCRM_NOTION_CONTACT_EMAIL] = { email: email }

  if (companyId) {
    props[CALCRM_NOTION_CONTACT_COMPANY_REL] = { relation: [{ id: companyId }] }
  }

  return { parent: { database_id: contactsDbId }, properties: props }
}

function calcrm_ensureContactRelatedToCompany_(notion, contactPage, companyId) {
  if (!contactPage || !contactPage.id || !companyId) return

  const rel = contactPage.properties && contactPage.properties[CALCRM_NOTION_CONTACT_COMPANY_REL]
  const current = (rel && rel.type === "relation" && Array.isArray(rel.relation)) ? rel.relation : []
  const has = current.some(x => x && x.id === companyId)
  if (has) return

  const patch = { properties: {} }
  patch.properties[CALCRM_NOTION_CONTACT_COMPANY_REL] = { relation: current.concat([{ id: companyId }]) }
  calcrm_notionPatch_(notion, `/pages/${contactPage.id}`, patch)
}

/** =========================
 * Audit map sheets
 * ========================= */

function calcrm_ensureMapSheet_(sheetName, headers) {
  const ss = SpreadsheetApp.getActive()
  let sh = ss.getSheetByName(sheetName)
  if (!sh) sh = ss.insertSheet(sheetName)

  const lastRow = sh.getLastRow()
  const lastCol = sh.getLastColumn()
  const firstCell = (lastRow >= 1 && lastCol >= 1) ? String(sh.getRange(1, 1).getValue() || "").trim() : ""

  if (firstCell !== headers[0]) {
    sh.clear()
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
    sh.setFrozenRows(1)
    return
  }

  // keep headers consistent
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
  sh.setFrozenRows(1)
}

function calcrm_upsertMapOrgRow_(obj) {
  // Upsert by org_id if present, else by notion_company_id
  calcrm_upsertRowByEitherKey_(
    CALCRM_SHEET_MAP_ORG,
    ["org_id", "notion_company_id"],
    CALCRM_MAP_ORG_HEADERS,
    obj
  )
}

function calcrm_upsertMapUserRow_(obj) {
  calcrm_upsertRowByKey_(
    CALCRM_SHEET_MAP_USER,
    "notion_contact_id",
    CALCRM_MAP_USER_HEADERS,
    obj
  )
}

function calcrm_upsertRowByKey_(sheetName, keyHeader, headers, obj) {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(sheetName)
  if (!sh) throw new Error(`Missing sheet: ${sheetName}`)

  const key = String(obj[keyHeader] || "").trim()
  if (!key) return

  const lastRow = sh.getLastRow()
  let rowToWrite = lastRow + 1

  if (lastRow >= 2) {
    const keyCol = headers.indexOf(keyHeader) + 1
    const keys = sh.getRange(2, keyCol, lastRow - 1, 1).getValues().map(r => String(r[0] || "").trim())
    const found = keys.findIndex(v => v === key)
    if (found >= 0) rowToWrite = 2 + found
  }

  const row = headers.map(h => Object.prototype.hasOwnProperty.call(obj, h) ? obj[h] : "")
  sh.getRange(rowToWrite, 1, 1, headers.length).setValues([row])
}

function calcrm_upsertRowByEitherKey_(sheetName, keyHeaders, headers, obj) {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(sheetName)
  if (!sh) throw new Error(`Missing sheet: ${sheetName}`)

  const lastRow = sh.getLastRow()
  let rowToWrite = lastRow + 1

  if (lastRow >= 2) {
    const data = sh.getRange(2, 1, lastRow - 1, headers.length).getValues()
    const found = data.findIndex(r => {
      return keyHeaders.some(kh => {
        const idx = headers.indexOf(kh)
        if (idx < 0) return false
        const incoming = String(obj[kh] || "").trim()
        const existing = String(r[idx] || "").trim()
        return incoming && existing && incoming === existing
      })
    })
    if (found >= 0) rowToWrite = 2 + found
  }

  const row = headers.map(h => Object.prototype.hasOwnProperty.call(obj, h) ? obj[h] : "")
  sh.getRange(rowToWrite, 1, 1, headers.length).setValues([row])
}

/** =========================
 * Email normalization + rules
 * ========================= */

function calcrm_normEmail_(v) {
  const s = String(v || "").trim().toLowerCase()
  if (!s) return ""
  return s.replace(/\+[^@]+(?=@)/, "")
}

function calcrm_isInternalEmail_(emailKey) {
  const domain = calcrm_getDomain_(emailKey)
  if (!domain) return false
  return CALCRM_INTERNAL_EMAIL_DOMAINS.includes(domain)
}

function calcrm_getDomain_(email) {
  const e = String(email || "").trim().toLowerCase()
  const i = e.indexOf("@")
  if (i < 0) return ""
  return e.slice(i + 1).trim()
}

/** =========================
 * Run linkers after import (best effort)
 * ========================= */

function calcrm_runLinkersBestEffort_() {
  // Run contacts linker first, then companies linker
  try {
    if (typeof notion_link_unlinked_contacts_to_sauron === "function") {
      Logger.log("Calendar import: running notion_link_unlinked_contacts_to_sauron()")
      notion_link_unlinked_contacts_to_sauron()
    } else {
      Logger.log("Calendar import: notion_link_unlinked_contacts_to_sauron not found (skipping)")
    }
  } catch (e) {
    Logger.log("Calendar import: contacts linker failed (non-fatal): " + String(e && e.message ? e.message : e))
  }

  try {
    if (typeof notion_link_unlinked_companies_to_sauron === "function") {
      Logger.log("Calendar import: running notion_link_unlinked_companies_to_sauron()")
      notion_link_unlinked_companies_to_sauron()
    } else {
      Logger.log("Calendar import: notion_link_unlinked_companies_to_sauron not found (skipping)")
    }
  } catch (e) {
    Logger.log("Calendar import: companies linker failed (non-fatal): " + String(e && e.message ? e.message : e))
  }
}

/** =========================
 * Notion client (namespaced)
 * ========================= */

function calcrm_notionClient_() {
  const props = PropertiesService.getScriptProperties()
  const token = calcrm_mustGetProp_(props, "NOTION_TOKEN")
  const version = props.getProperty("NOTION_VERSION") || "2022-06-28"
  return { token, version, baseUrl: "https://api.notion.com/v1" }
}

function calcrm_notionQueryAll_(notion, databaseId, body) {
  const out = []
  let cursor = null

  while (true) {
    const payload = Object.assign({}, body || {})
    if (cursor) payload.start_cursor = cursor

    const res = calcrm_notionPost_(notion, `/databases/${databaseId}/query`, payload)
    const results = res && res.results ? res.results : []
    out.push(...results)

    if (!res.has_more) break
    cursor = res.next_cursor
    if (!cursor) break
    if (out.length > 5000) break
  }
  return out
}

function calcrm_notionPost_(notion, path, payload) {
  return calcrm_notionFetch_(notion, "post", path, payload)
}

function calcrm_notionPatch_(notion, path, payload) {
  return calcrm_notionFetch_(notion, "patch", path, payload)
}

function calcrm_notionFetch_(notion, method, path, payload) {
  const url = notion.baseUrl + path
  const options = {
    method,
    contentType: "application/json",
    muteHttpExceptions: true,
    headers: {
      Authorization: `Bearer ${notion.token}`,
      "Notion-Version": notion.version
    },
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
 * Notion readers (namespaced)
 * ========================= */

function calcrm_notionGetRichText_(page, propName) {
  const p = page && page.properties ? page.properties[propName] : null
  if (!p) return ""
  if (p.type === "rich_text") return (p.rich_text || []).map(t => t.plain_text).join("").trim()
  if (p.type === "title") return (p.title || []).map(t => t.plain_text).join("").trim()
  return ""
}

function calcrm_notionGetTitle_(page, propName) {
  const p = page && page.properties ? page.properties[propName] : null
  if (!p || p.type !== "title") return ""
  return (p.title || []).map(t => t.plain_text).join("").trim()
}

function calcrm_notionGetTitleAny_(page) {
  const props = page && page.properties ? page.properties : {}
  for (const k in props) {
    const p = props[k]
    if (p && p.type === "title") return (p.title || []).map(t => t.plain_text).join("").trim()
  }
  return ""
}

/** =========================
 * Tiny utils
 * ========================= */

function calcrm_mustGetProp_(props, key) {
  const v = props.getProperty(key)
  if (!v) throw new Error(`Missing Script Property: ${key}`)
  return v
}

function calcrm_safeLower_(v) {
  return String(v || "").trim().toLowerCase()
}