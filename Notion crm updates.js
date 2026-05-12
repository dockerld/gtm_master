/**************************************************************
 * Notion CRM Sync (Google Apps Script)
 *
 * One-way sync: Google Sheets → Notion CRM
 * - Sheets = source of truth for app data (canon_orgs, canon_users)
 * - Notion = source of truth for sales pipeline
 *
 * Link keys:
 * - Companies: sauron_org_id (permanent, once stamped)
 * - Contacts:  sauron_user_id (permanent, once stamped)
 *
 * Core functions:
 * 1. link_unlinked_orgs()      — match unlinked companies via contact emails
 * 2. link_unlinked_contacts()  — match unlinked contacts via email
 * 3. sync_orgs_to_notion()     — push latest org data + auto-create missing
 * 4. sync_users_to_notion()    — push latest user data
 **************************************************************/

/** =========================
 * CONFIG
 * ========================= */

const PROP_NOTION_TOKEN = "NOTION_TOKEN"
const PROP_NOTION_VERSION = "NOTION_VERSION"
const PROP_NOTION_COMPANIES_DB_ID = "NOTION_COMPANIES_DB_ID"
const PROP_NOTION_CONTACTS_DB_ID = "NOTION_CONTACTS_DB_ID"

// Sheets
const SHEET_CANON_ORGS = "canon_orgs"
const SHEET_CANON_USERS = "canon_users"

// Notion Company properties
const NOTION_COMPANY_PROP_NAME = "Company Name"
const NOTION_COMPANY_PROP_LINKED = "linked_to_sauron"
const NOTION_COMPANY_PROP_SAURON_ORG_ID = "sauron_org_id"
const NOTION_COMPANY_PROP_LINK_SOURCE = "link_source"
const NOTION_COMPANY_PROP_LINKED_AT = "linked_at"
const NOTION_COMPANY_PROP_PAID_SEATS = "Current Seats"
const NOTION_COMPANY_PROP_IS_PAYING = "Is Paying"
const NOTION_COMPANY_PROP_SUBSCRIPTION_STATUS = "Subscription Status"
const NOTION_COMPANY_PROP_ORG_CREATED_AT = "Created on Ping"
const NOTION_COMPANY_PROP_FIRM_SIZE = "Firm Size"
const NOTION_COMPANY_PROP_LAST_SYNCED = "Last Synced At"
const NOTION_COMPANY_PROP_PIPELINE_STAGE = "Pipeline Stage"
const NOTION_COMPANY_PROP_TRIAL_END_DATE = "Trial End Date"
const NOTION_COMPANY_PROP_HEALTH_SCORE = "Health Score"

// Notion Contact properties
const NOTION_CONTACT_PROP_NAME = "Name"
const NOTION_CONTACT_PROP_EMAIL = "Email"
const NOTION_CONTACT_PROP_COMPANY_REL = "Company"
const NOTION_CONTACT_PROP_LINKED = "linked_to_sauron"
const NOTION_CONTACT_PROP_SAURON_ORG_ID = "sauron_org_id"
const NOTION_CONTACT_PROP_SAURON_USER_ID = "sauron_user_id"
const NOTION_CONTACT_PROP_PING_CREATED = "Ping account created"
const NOTION_CONTACT_PROP_LAST_SYNCED = "Last Synced At"

// Firm Size
const FIRM_SIZE_RANK = { "Solo": 1, "Small (<10)": 2, "Medium (10+)": 3, "Large (50+)": 4 }

// Domain
const CRM_PERSONAL_DOMAINS = new Set([
  "gmail.com", "googlemail.com", "yahoo.com", "hotmail.com",
  "outlook.com", "live.com", "icloud.com", "me.com", "mac.com",
  "aol.com", "proton.me", "protonmail.com"
])
const NOTION_COMPANY_PROP_DOMAIN = "domain"

// Limits
const MAX_PER_RUN = 300

/** =========================
 * FIRM SIZE HELPERS
 * ========================= */

function seatsToFirmSize_(seats) {
  const n = Number(seats) || 0
  if (n >= 50) return "Large (50+)"
  if (n >= 10) return "Medium (10+)"
  if (n >= 2) return "Small (<10)"
  return "Solo"
}

function shouldUpdateFirmSize_(currentNotion, fromSheet) {
  return (FIRM_SIZE_RANK[fromSheet] || 0) > (FIRM_SIZE_RANK[currentNotion] || 0)
}

function notionGetSelect_(page, propName) {
  const p = page && page.properties ? page.properties[propName] : null
  if (!p || p.type !== "select" || !p.select) return ""
  return p.select.name || ""
}

/** =========================
 * SHEET INDEX BUILDERS
 * ========================= */

function buildOrgIndex_() {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(SHEET_CANON_ORGS)
  if (!sh) throw new Error(`Missing sheet: ${SHEET_CANON_ORGS}`)
  const rows = readSheetObjects_(sh, 1)
  const map = new Map()
  for (const r of rows) {
    const orgId = str_(r.org_id)
    if (orgId) map.set(orgId, r)
  }
  return map
}

function buildUserIndex_() {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(SHEET_CANON_USERS)
  if (!sh) throw new Error(`Missing sheet: ${SHEET_CANON_USERS}`)
  const rows = readSheetObjects_(sh, 1)
  const byEmail = new Map()
  const byUserId = new Map()
  const byDomain = new Map()
  for (const r of rows) {
    r._org_id = str_(r.org_id)
    const emailKey = normEmail_(str_(r.email_key || r.email))
    const userId = str_(r.clerk_user_id)
    if (emailKey) {
      byEmail.set(emailKey, r)
      const domain = getDomain_(emailKey)
      if (domain && !CRM_PERSONAL_DOMAINS.has(domain) && !byDomain.has(domain)) {
        byDomain.set(domain, r)
      }
    }
    if (userId) byUserId.set(userId, r)
  }
  return { byEmail, byUserId, byDomain }
}

/** =========================
 * CORE 1: LINK UNLINKED ORGS
 * ========================= */

function link_unlinked_orgs() {
  const notion = notionClient_()
  const props = PropertiesService.getScriptProperties()
  const companiesDbId = mustGetProp_(props, PROP_NOTION_COMPANIES_DB_ID)
  const contactsDbId = mustGetProp_(props, PROP_NOTION_CONTACTS_DB_ID)
  const userIndex = buildUserIndex_()
  const nowIso = new Date().toISOString()

  const unlinked = notionQueryAll_(notion, companiesDbId, {
    filter: { property: NOTION_COMPANY_PROP_LINKED, checkbox: { equals: false } },
    page_size: 100
  }).slice(0, MAX_PER_RUN)

  Logger.log(`link_unlinked_orgs: found ${unlinked.length} unlinked companies`)

  let linked = 0

  for (const company of unlinked) {
    const companyId = company.id
    const existingOrgId = notionGetRichText_(company, NOTION_COMPANY_PROP_SAURON_ORG_ID)

    // Fast path: already has org_id (e.g. from Calendar-crm), just stamp linked
    if (existingOrgId) {
      notionUpdatePage_(notion, companyId, { properties: {
        [NOTION_COMPANY_PROP_LINKED]: { checkbox: true },
        [NOTION_COMPANY_PROP_LINK_SOURCE]: { select: { name: "pre_stamped" } },
        [NOTION_COMPANY_PROP_LINKED_AT]: { date: { start: nowIso } }
      }})
      linked++
      continue
    }

    // Email match: get company contacts, match emails to canon_users
    const contactEmails = getCompanyContactEmails_(notion, contactsDbId, companyId)
    let matchedOrgId = ""
    let matchedContacts = []
    let linkSource = "email_match"

    for (const { emailKey, contactId } of contactEmails) {
      const hit = userIndex.byEmail.get(emailKey)
      if (hit && str_(hit._org_id)) {
        matchedOrgId = str_(hit._org_id)
        matchedContacts.push({ contactId, hit })
        break
      }
    }

    // Domain fallback: if no email match, try matching contact domains
    if (!matchedOrgId) {
      for (const { emailKey } of contactEmails) {
        const domain = getDomain_(emailKey)
        if (!domain || CRM_PERSONAL_DOMAINS.has(domain)) continue
        const hit = userIndex.byDomain.get(domain)
        if (hit && str_(hit._org_id)) {
          matchedOrgId = str_(hit._org_id)
          linkSource = "domain_match"
          break
        }
      }
    }

    if (!matchedOrgId) continue

    // Stamp company
    const companyName = notionGetTitleAny_(company) || companyId
    Logger.log(`  LINKED ORG: "${companyName}" → ${matchedOrgId} (${linkSource})`)
    notionUpdatePage_(notion, companyId, { properties: {
      [NOTION_COMPANY_PROP_SAURON_ORG_ID]: { rich_text: [{ type: "text", text: { content: matchedOrgId } }] },
      [NOTION_COMPANY_PROP_LINKED]: { checkbox: true },
      [NOTION_COMPANY_PROP_LINK_SOURCE]: { select: { name: linkSource } },
      [NOTION_COMPANY_PROP_LINKED_AT]: { date: { start: nowIso } }
    }})

    // Side effect: stamp matched contacts
    for (const { contactId, hit } of matchedContacts) {
      const contactPatch = {
        [NOTION_CONTACT_PROP_LINKED]: { checkbox: true },
        [NOTION_CONTACT_PROP_SAURON_ORG_ID]: { rich_text: [{ type: "text", text: { content: matchedOrgId } }] }
      }
      const userId = str_(hit.clerk_user_id)
      if (userId) {
        contactPatch[NOTION_CONTACT_PROP_SAURON_USER_ID] = { rich_text: [{ type: "text", text: { content: userId } }] }
      }
      notionUpdatePage_(notion, contactId, { properties: contactPatch })
    }

    linked++
    if (linked % 50 === 0) Utilities.sleep(100)
  }

  Logger.log(`link_unlinked_orgs done: checked=${unlinked.length}, linked=${linked}`)
  return { rows_in: unlinked.length, rows_out: linked }
}

function getCompanyContactEmails_(notion, contactsDbId, companyId) {
  const contacts = notionQueryAll_(notion, contactsDbId, {
    filter: { property: NOTION_CONTACT_PROP_COMPANY_REL, relation: { contains: companyId } },
    page_size: 100
  }).slice(0, 200)

  const results = []
  for (const c of contacts) {
    const email = notionGetEmail_(c, NOTION_CONTACT_PROP_EMAIL)
    const emailKey = normEmail_(email)
    if (emailKey) results.push({ emailKey, contactId: c.id })
  }
  return results
}

function findContactByEmail_(notion, contactsDbId, email) {
  const results = notionQueryAll_(notion, contactsDbId, {
    filter: { property: NOTION_CONTACT_PROP_EMAIL, email: { equals: email } },
    page_size: 1
  })
  return results.length > 0 ? results[0] : null
}

function createOwnerContactIfNeeded_(notion, contactsDbId, companyPageId, org, orgId, nowIso) {
  const ownerEmail = str_(org.owner_email)
  if (!ownerEmail) return

  const existing = findContactByEmail_(notion, contactsDbId, ownerEmail)
  if (existing) {
    // Ensure existing contact is related to this company
    const rel = existing.properties && existing.properties[NOTION_CONTACT_PROP_COMPANY_REL]
    const relIds = (rel && rel.type === "relation" && Array.isArray(rel.relation))
      ? rel.relation.map(r => r.id) : []
    if (!relIds.includes(companyPageId)) {
      notionUpdatePage_(notion, existing.id, { properties: {
        [NOTION_CONTACT_PROP_COMPANY_REL]: { relation: relIds.concat([companyPageId]).map(id => ({ id })) },
        [NOTION_CONTACT_PROP_LINKED]: { checkbox: true },
        [NOTION_CONTACT_PROP_SAURON_ORG_ID]: { rich_text: [{ type: "text", text: { content: orgId } }] }
      }})
    }
    // Stamp user_id if not already set
    const existingUserId = notionGetRichText_(existing, NOTION_CONTACT_PROP_SAURON_USER_ID)
    if (!existingUserId && str_(org.owner_user_id)) {
      notionUpdatePage_(notion, existing.id, { properties: {
        [NOTION_CONTACT_PROP_SAURON_USER_ID]: { rich_text: [{ type: "text", text: { content: str_(org.owner_user_id) } }] }
      }})
    }
    return
  }

  // Create new contact
  const ownerName = str_(org.owner_name) || ownerEmail
  const ownerUserId = str_(org.owner_user_id)
  const contactProps = {
    [NOTION_CONTACT_PROP_NAME]: { title: [{ type: "text", text: { content: ownerName } }] },
    [NOTION_CONTACT_PROP_EMAIL]: { email: ownerEmail },
    [NOTION_CONTACT_PROP_COMPANY_REL]: { relation: [{ id: companyPageId }] },
    [NOTION_CONTACT_PROP_LINKED]: { checkbox: true },
    [NOTION_CONTACT_PROP_SAURON_ORG_ID]: { rich_text: [{ type: "text", text: { content: orgId } }] },
    [NOTION_CONTACT_PROP_LAST_SYNCED]: { date: { start: nowIso } }
  }
  if (ownerUserId) {
    contactProps[NOTION_CONTACT_PROP_SAURON_USER_ID] = { rich_text: [{ type: "text", text: { content: ownerUserId } }] }
  }
  const contactsDsId = notionResolveDataSourceId_(notion, contactsDbId)
  notionPost_(notion, "/pages", {
    parent: { data_source_id: contactsDsId },
    properties: contactProps
  })
}

/** =========================
 * CORE 2: LINK UNLINKED CONTACTS
 * ========================= */

function link_unlinked_contacts() {
  const notion = notionClient_()
  const props = PropertiesService.getScriptProperties()
  const companiesDbId = mustGetProp_(props, PROP_NOTION_COMPANIES_DB_ID)
  const contactsDbId = mustGetProp_(props, PROP_NOTION_CONTACTS_DB_ID)
  const userIndex = buildUserIndex_()
  const nowIso = new Date().toISOString()

  const unlinked = notionQueryAll_(notion, contactsDbId, {
    filter: { property: NOTION_CONTACT_PROP_LINKED, checkbox: { equals: false } },
    page_size: 100
  }).slice(0, MAX_PER_RUN)

  Logger.log(`link_unlinked_contacts: found ${unlinked.length} unlinked contacts`)

  let linked = 0

  for (const contact of unlinked) {
    const contactId = contact.id
    const email = notionGetEmail_(contact, NOTION_CONTACT_PROP_EMAIL)
    const emailKey = normEmail_(email)
    if (!emailKey) continue

    const hit = userIndex.byEmail.get(emailKey)
    if (!hit || !str_(hit._org_id)) continue

    const orgId = str_(hit._org_id)
    const userId = str_(hit.clerk_user_id)

    // Stamp contact
    const contactPatch = {
      [NOTION_CONTACT_PROP_LINKED]: { checkbox: true },
      [NOTION_CONTACT_PROP_SAURON_ORG_ID]: { rich_text: [{ type: "text", text: { content: orgId } }] }
    }
    if (userId) {
      contactPatch[NOTION_CONTACT_PROP_SAURON_USER_ID] = { rich_text: [{ type: "text", text: { content: userId } }] }
    }
    // Set Ping account created if available
    const createdAt = parseDate_(hit.created_at)
    if (createdAt) {
      contactPatch[NOTION_CONTACT_PROP_PING_CREATED] = { date: { start: createdAt } }
    }
    notionUpdatePage_(notion, contactId, { properties: contactPatch })
    Logger.log(`  LINKED CONTACT: "${email}" → org=${orgId}${userId ? ", user=" + userId : ""}`)

    // Fix email-as-title
    const name = str_(hit.name)
    if (name && !looksLikeEmail_(name)) {
      tryUpgradeNotionContactTitle_(notion, contactId, name)
    }

    // If parent company is unlinked, link it too
    const companyRel = contact.properties && contact.properties[NOTION_CONTACT_PROP_COMPANY_REL]
    const companyIds = (companyRel && companyRel.type === "relation" && Array.isArray(companyRel.relation))
      ? companyRel.relation.map(r => r.id)
      : []

    for (const cid of companyIds) {
      try {
        const companyPage = notionGetPage_(notion, cid)
        const alreadyLinked = companyPage.properties &&
          companyPage.properties[NOTION_COMPANY_PROP_LINKED] &&
          companyPage.properties[NOTION_COMPANY_PROP_LINKED].checkbox === true
        if (alreadyLinked) continue

        notionUpdatePage_(notion, cid, { properties: {
          [NOTION_COMPANY_PROP_SAURON_ORG_ID]: { rich_text: [{ type: "text", text: { content: orgId } }] },
          [NOTION_COMPANY_PROP_LINKED]: { checkbox: true },
          [NOTION_COMPANY_PROP_LINK_SOURCE]: { select: { name: "contact_email_match" } },
          [NOTION_COMPANY_PROP_LINKED_AT]: { date: { start: nowIso } }
        }})
      } catch (e) {
        Logger.log(`link_unlinked_contacts: failed to link company ${cid}: ${e.message}`)
      }
    }

    linked++
    if (linked % 50 === 0) Utilities.sleep(100)
  }

  Logger.log(`link_unlinked_contacts done: checked=${unlinked.length}, linked=${linked}`)
  return { rows_in: unlinked.length, rows_out: linked }
}

/** =========================
 * CORE 3: SYNC ORGS TO NOTION
 * ========================= */

function sync_orgs_to_notion() {
  const notion = notionClient_()
  const props = PropertiesService.getScriptProperties()
  const companiesDbId = mustGetProp_(props, PROP_NOTION_COMPANIES_DB_ID)
  const contactsDbId = mustGetProp_(props, PROP_NOTION_CONTACTS_DB_ID)
  const orgIndex = buildOrgIndex_()
  const nowIso = new Date().toISOString()

  // Phase A: Auto-create — find orgs in sheet not yet in Notion
  const existingCompanies = notionQueryAll_(notion, companiesDbId, {
    filter: { property: NOTION_COMPANY_PROP_SAURON_ORG_ID, rich_text: { is_not_empty: true } },
    page_size: 100
  })

  const existingOrgIds = new Set()
  for (const c of existingCompanies) {
    const oid = notionGetRichText_(c, NOTION_COMPANY_PROP_SAURON_ORG_ID)
    if (oid) existingOrgIds.add(oid)
  }

  // Build domain index of existing companies for dedup
  const allCompaniesForDomain = notionQueryAll_(notion, companiesDbId, { page_size: 100 })
  const existingByDomain = new Map()
  for (const c of allCompaniesForDomain) {
    const d = notionGetRichText_(c, NOTION_COMPANY_PROP_DOMAIN)
    if (d && !CRM_PERSONAL_DOMAINS.has(d)) existingByDomain.set(d, c)
  }

  let created = 0
  let domainLinked = 0
  for (const [orgId, org] of orgIndex) {
    if (existingOrgIds.has(orgId)) continue

    const orgName = str_(org.org_name) || orgId
    const seats = Number(org.seats) || 0
    const firmSize = seatsToFirmSize_(seats)
    const isPaying = String(org.is_paying).toLowerCase() === "true" || org.is_paying === true
    const subStatus = str_(org.org_status) || "Expired"
    const createdAt = parseDate_(org.org_created_at)
    const ownerEmail = str_(org.owner_email)
    const domain = getDomain_(ownerEmail)
    const trialEndDate = parseDateOnly_(org.trial_ends_at)
    const healthScore = Number(org.health_score) || 0

    // Dedup: if a company exists with this domain but no org_id, link it instead
    if (domain && !CRM_PERSONAL_DOMAINS.has(domain) && existingByDomain.has(domain)) {
      const existing = existingByDomain.get(domain)
      const existingOid = notionGetRichText_(existing, NOTION_COMPANY_PROP_SAURON_ORG_ID)
      if (!existingOid) {
        notionUpdatePage_(notion, existing.id, { properties: {
          [NOTION_COMPANY_PROP_SAURON_ORG_ID]: { rich_text: [{ type: "text", text: { content: orgId } }] },
          [NOTION_COMPANY_PROP_LINKED]: { checkbox: true },
          [NOTION_COMPANY_PROP_LINK_SOURCE]: { select: { name: "domain_dedup" } },
          [NOTION_COMPANY_PROP_LINKED_AT]: { date: { start: nowIso } },
          [NOTION_COMPANY_PROP_PAID_SEATS]: { number: seats },
          [NOTION_COMPANY_PROP_FIRM_SIZE]: { select: { name: firmSize } },
          [NOTION_COMPANY_PROP_IS_PAYING]: { checkbox: isPaying },
          [NOTION_COMPANY_PROP_LAST_SYNCED]: { date: { start: nowIso } },
          [NOTION_COMPANY_PROP_HEALTH_SCORE]: { number: healthScore },
          ...(trialEndDate ? { [NOTION_COMPANY_PROP_TRIAL_END_DATE]: { date: { start: trialEndDate } } } : {})
        }})
        Logger.log(`  DOMAIN-LINKED: "${orgName}" (${orgId}) → existing company by domain ${domain}`)
        existingOrgIds.add(orgId)
        domainLinked++
        continue
      }
    }

    const createProps = {
      [NOTION_COMPANY_PROP_NAME]: { title: [{ type: "text", text: { content: orgName } }] },
      [NOTION_COMPANY_PROP_SAURON_ORG_ID]: { rich_text: [{ type: "text", text: { content: orgId } }] },
      [NOTION_COMPANY_PROP_LINKED]: { checkbox: true },
      [NOTION_COMPANY_PROP_LINK_SOURCE]: { select: { name: "auto_create" } },
      [NOTION_COMPANY_PROP_LINKED_AT]: { date: { start: nowIso } },
      [NOTION_COMPANY_PROP_PAID_SEATS]: { number: seats },
      [NOTION_COMPANY_PROP_FIRM_SIZE]: { select: { name: firmSize } },
      [NOTION_COMPANY_PROP_IS_PAYING]: { checkbox: isPaying },
      [NOTION_COMPANY_PROP_HEALTH_SCORE]: { number: healthScore },
      [NOTION_COMPANY_PROP_LAST_SYNCED]: { date: { start: nowIso } }
    }
    if (subStatus) {
      createProps[NOTION_COMPANY_PROP_SUBSCRIPTION_STATUS] = { select: { name: subStatus } }
    }
    if (createdAt) {
      createProps[NOTION_COMPANY_PROP_ORG_CREATED_AT] = { date: { start: createdAt } }
    }
    if (domain && !CRM_PERSONAL_DOMAINS.has(domain)) {
      createProps[NOTION_COMPANY_PROP_DOMAIN] = { rich_text: [{ type: "text", text: { content: domain } }] }
    }
    if (isPaying) {
      createProps[NOTION_COMPANY_PROP_PIPELINE_STAGE] = { select: { name: "Closed Won" } }
    }
    if (trialEndDate) {
      createProps[NOTION_COMPANY_PROP_TRIAL_END_DATE] = { date: { start: trialEndDate } }
    }

    const companiesDsId = notionResolveDataSourceId_(notion, companiesDbId)
    const companyPage = notionPost_(notion, "/pages", {
      parent: { data_source_id: companiesDsId },
      properties: createProps
    })

    // Auto-create owner contact
    if (ownerEmail && companyPage && companyPage.id) {
      createOwnerContactIfNeeded_(notion, contactsDbId, companyPage.id, org, orgId, nowIso)
    }

    // PDL enrichment (only when domain is set)
    if (companyPage && companyPage.id && domain && !CRM_PERSONAL_DOMAINS.has(domain)) {
      enrichWithPdl_(companyPage.id)
    }

    Logger.log(`  CREATED: "${orgName}" (${orgId}) — seats=${seats}, paying=${isPaying}, status=${subStatus || "none"}`)
    created++
    if (created % 50 === 0) Utilities.sleep(100)
  }

  Logger.log(`sync_orgs_to_notion: auto-created ${created}, domain-linked ${domainLinked}`)

  // Phase B: Sync existing — re-query to include newly created
  const allCompanies = notionQueryAll_(notion, companiesDbId, {
    filter: { property: NOTION_COMPANY_PROP_SAURON_ORG_ID, rich_text: { is_not_empty: true } },
    page_size: 100
  })

  let synced = 0
  let skipped = 0
  for (const company of allCompanies) {
    const orgId = notionGetRichText_(company, NOTION_COMPANY_PROP_SAURON_ORG_ID)
    const org = orgIndex.get(orgId)
    if (!org) continue

    const seats = Number(org.seats) || 0
    const isPaying = String(org.is_paying).toLowerCase() === "true" || org.is_paying === true
    const subStatus = str_(org.org_status) || "Expired"
    const createdAt = parseDate_(org.org_created_at)
    const ownerEmail = str_(org.owner_email)
    const domain = getDomain_(ownerEmail)
    const trialEndDate = parseDateOnly_(org.trial_ends_at)
    const healthScore = Number(org.health_score) || 0
    const sheetFirmSize = seatsToFirmSize_(seats)
    const companyName = notionGetTitleAny_(company) || orgId

    // Read current Notion values to diff
    const curSeats = (company.properties && company.properties[NOTION_COMPANY_PROP_PAID_SEATS] && company.properties[NOTION_COMPANY_PROP_PAID_SEATS].number) || 0
    const curPaying = (company.properties && company.properties[NOTION_COMPANY_PROP_IS_PAYING] && company.properties[NOTION_COMPANY_PROP_IS_PAYING].checkbox) || false
    const curStatus = notionGetSelect_(company, NOTION_COMPANY_PROP_SUBSCRIPTION_STATUS)
    const curFirmSize = notionGetSelect_(company, NOTION_COMPANY_PROP_FIRM_SIZE)
    const curDomain = notionGetRichText_(company, NOTION_COMPANY_PROP_DOMAIN)
    const curPipeline = notionGetSelect_(company, NOTION_COMPANY_PROP_PIPELINE_STAGE)
    const curTrialEnd = (company.properties && company.properties[NOTION_COMPANY_PROP_TRIAL_END_DATE] && company.properties[NOTION_COMPANY_PROP_TRIAL_END_DATE].date && company.properties[NOTION_COMPANY_PROP_TRIAL_END_DATE].date.start) || ""
    const curHealthScore = (company.properties && company.properties[NOTION_COMPANY_PROP_HEALTH_SCORE] && company.properties[NOTION_COMPANY_PROP_HEALTH_SCORE].number) || 0

    // Check if anything actually changed
    const seatsChanged = curSeats !== seats
    const payingChanged = curPaying !== isPaying
    const statusChanged = subStatus && curStatus.toLowerCase() !== subStatus.toLowerCase()
    const firmSizeChanged = shouldUpdateFirmSize_(curFirmSize, sheetFirmSize)
    const domainChanged = domain && !CRM_PERSONAL_DOMAINS.has(domain) && curDomain.toLowerCase() !== domain.toLowerCase()
    const pipelineChanged = isPaying && curPipeline.toLowerCase() !== "closed won"
    const trialEndChanged = trialEndDate && curTrialEnd !== trialEndDate
    const healthScoreChanged = curHealthScore !== healthScore

    if (!seatsChanged && !payingChanged && !statusChanged && !firmSizeChanged && !domainChanged && !pipelineChanged && !trialEndChanged && !healthScoreChanged) {
      skipped++
      continue
    }

    const patch = {
      [NOTION_COMPANY_PROP_PAID_SEATS]: { number: seats },
      [NOTION_COMPANY_PROP_IS_PAYING]: { checkbox: isPaying },
      [NOTION_COMPANY_PROP_LAST_SYNCED]: { date: { start: nowIso } }
    }

    if (subStatus) {
      patch[NOTION_COMPANY_PROP_SUBSCRIPTION_STATUS] = { select: { name: subStatus } }
    }
    if (createdAt) {
      patch[NOTION_COMPANY_PROP_ORG_CREATED_AT] = { date: { start: createdAt } }
    }
    if (domainChanged) {
      patch[NOTION_COMPANY_PROP_DOMAIN] = { rich_text: [{ type: "text", text: { content: domain } }] }
    }
    if (firmSizeChanged) {
      patch[NOTION_COMPANY_PROP_FIRM_SIZE] = { select: { name: sheetFirmSize } }
    }
    if (isPaying) {
      patch[NOTION_COMPANY_PROP_PIPELINE_STAGE] = { select: { name: "Closed Won" } }
    }
    if (trialEndChanged) {
      patch[NOTION_COMPANY_PROP_TRIAL_END_DATE] = { date: { start: trialEndDate } }
    }
    if (healthScoreChanged) {
      patch[NOTION_COMPANY_PROP_HEALTH_SCORE] = { number: healthScore }
    }

    const changes = []
    if (seatsChanged) changes.push(`seats: ${curSeats}→${seats}`)
    if (payingChanged) changes.push(`paying: ${curPaying}→${isPaying}`)
    if (statusChanged) changes.push(`status: ${curStatus}→${subStatus}`)
    if (firmSizeChanged) changes.push(`firmSize: ${curFirmSize}→${sheetFirmSize}`)
    if (domainChanged) changes.push(`domain: ${curDomain}→${domain}`)
    if (pipelineChanged) changes.push(`pipeline: ${curPipeline}→Closed Won`)
    if (trialEndChanged) changes.push(`trialEnd: ${curTrialEnd}→${trialEndDate}`)
    if (healthScoreChanged) changes.push(`healthScore: ${curHealthScore}→${healthScore}`)

    notionUpdatePage_(notion, company.id, { properties: patch })
    Logger.log(`  SYNCED: "${companyName}" — ${changes.join(", ")}`)
    synced++
    if (synced % 50 === 0) Utilities.sleep(100)
  }

  // Phase C: Ensure every linked company has at least one contact (owner)
  let ensured = 0
  for (const company of allCompanies) {
    const orgId = notionGetRichText_(company, NOTION_COMPANY_PROP_SAURON_ORG_ID)
    const org = orgIndex.get(orgId)
    if (!org || !str_(org.owner_email)) continue

    const contacts = getCompanyContactEmails_(notion, contactsDbId, company.id)
    if (contacts.length > 0) continue

    createOwnerContactIfNeeded_(notion, contactsDbId, company.id, org, orgId, nowIso)
    ensured++
    if (ensured % 50 === 0) Utilities.sleep(100)
  }

  Logger.log(`sync_orgs_to_notion done: created=${created}, domain_linked=${domainLinked}, synced=${synced}, unchanged=${skipped}, ensured_owners=${ensured}`)
  return { rows_in: allCompanies.length, rows_out: synced }
}

/** =========================
 * CORE 4: SYNC USERS TO NOTION
 * ========================= */

function sync_users_to_notion() {
  const notion = notionClient_()
  const props = PropertiesService.getScriptProperties()
  const contactsDbId = mustGetProp_(props, PROP_NOTION_CONTACTS_DB_ID)
  const userIndex = buildUserIndex_()
  const nowIso = new Date().toISOString()

  const linkedContacts = notionQueryAll_(notion, contactsDbId, {
    filter: { property: NOTION_CONTACT_PROP_SAURON_USER_ID, rich_text: { is_not_empty: true } },
    page_size: 100
  })

  Logger.log(`sync_users_to_notion: found ${linkedContacts.length} linked contacts`)

  let synced = 0
  let skipped = 0
  for (const contact of linkedContacts) {
    const userId = notionGetRichText_(contact, NOTION_CONTACT_PROP_SAURON_USER_ID)
    const user = userIndex.byUserId.get(userId)
    if (!user) continue

    const createdAt = parseDate_(user.created_at)
    const name = str_(user.name)
    const contactEmail = notionGetEmail_(contact, NOTION_CONTACT_PROP_EMAIL) || ""
    const contactTitle = notionGetTitleAny_(contact) || contactEmail || userId

    // Check if anything actually changed
    const curCreatedAt = (contact.properties && contact.properties[NOTION_CONTACT_PROP_PING_CREATED] && contact.properties[NOTION_CONTACT_PROP_PING_CREATED].date && contact.properties[NOTION_CONTACT_PROP_PING_CREATED].date.start) || ""
    const createdChanged = createdAt && curCreatedAt !== createdAt
    const curTitle = notionGetTitleAny_(contact) || ""
    const nameNeedsUpgrade = name && !looksLikeEmail_(name) && (looksLikeEmail_(curTitle) || !curTitle)

    if (!createdChanged && !nameNeedsUpgrade) {
      skipped++
      continue
    }

    const patch = {
      [NOTION_CONTACT_PROP_LAST_SYNCED]: { date: { start: nowIso } }
    }
    if (createdAt) {
      patch[NOTION_CONTACT_PROP_PING_CREATED] = { date: { start: createdAt } }
    }

    const changes = []
    if (createdChanged) changes.push(`created: ${curCreatedAt || "none"}→${createdAt}`)
    if (nameNeedsUpgrade) changes.push(`name: "${curTitle}"→"${name}"`)

    notionUpdatePage_(notion, contact.id, { properties: patch })

    // Fix email-as-title
    if (nameNeedsUpgrade) {
      tryUpgradeNotionContactTitle_(notion, contact.id, name)
    }

    Logger.log(`  SYNCED CONTACT: "${contactTitle}" — ${changes.join(", ")}`)
    synced++
    if (synced % 50 === 0) Utilities.sleep(100)
  }

  Logger.log(`sync_users_to_notion done: checked=${linkedContacts.length}, synced=${synced}, unchanged=${skipped}`)
  return { rows_in: linkedContacts.length, rows_out: synced }
}

/** =========================
 * DEDUP CLEANUP
 * ========================= */

function crm_dedup_cleanup() {
  const notion = notionClient_()
  const props = PropertiesService.getScriptProperties()
  const companiesDbId = mustGetProp_(props, PROP_NOTION_COMPANIES_DB_ID)
  const contactsDbId = mustGetProp_(props, PROP_NOTION_CONTACTS_DB_ID)

  let contactsArchived = 0
  let companiesArchived = 0
  let domainDupes = []

  // 1. Contact dedup by email
  const allContacts = notionQueryAll_(notion, contactsDbId, { page_size: 100 })
  const contactsByEmail = new Map()
  for (const c of allContacts) {
    const email = normEmail_(notionGetEmail_(c, NOTION_CONTACT_PROP_EMAIL))
    if (!email) continue
    if (!contactsByEmail.has(email)) contactsByEmail.set(email, [])
    contactsByEmail.get(email).push(c)
  }

  for (const [email, pages] of contactsByEmail) {
    if (pages.length <= 1) continue

    // Score each: linked > has company > has real name
    const scored = pages.map(p => {
      let score = 0
      const linked = p.properties && p.properties[NOTION_CONTACT_PROP_LINKED]
      if (linked && linked.checkbox === true) score += 10
      const userId = notionGetRichText_(p, NOTION_CONTACT_PROP_SAURON_USER_ID)
      if (userId) score += 5
      const rel = p.properties && p.properties[NOTION_CONTACT_PROP_COMPANY_REL]
      if (rel && rel.relation && rel.relation.length > 0) score += 3
      const title = notionGetTitleAny_(p)
      if (title && !looksLikeEmail_(title)) score += 1
      return { page: p, score }
    }).sort((a, b) => b.score - a.score)

    const keeper = scored[0].page

    // Merge company relations from losers to keeper
    const keeperRel = keeper.properties && keeper.properties[NOTION_CONTACT_PROP_COMPANY_REL]
    const keeperRelIds = (keeperRel && keeperRel.relation) ? keeperRel.relation.map(r => r.id) : []
    const allRelIds = new Set(keeperRelIds)

    for (let i = 1; i < scored.length; i++) {
      const loser = scored[i].page
      const loserRel = loser.properties && loser.properties[NOTION_CONTACT_PROP_COMPANY_REL]
      if (loserRel && loserRel.relation) {
        for (const r of loserRel.relation) allRelIds.add(r.id)
      }
      notionPatch_(notion, `/pages/${loser.id}`, { archived: true })
      contactsArchived++
    }

    // Update keeper with merged relations if changed
    if (allRelIds.size > keeperRelIds.length) {
      notionUpdatePage_(notion, keeper.id, { properties: {
        [NOTION_CONTACT_PROP_COMPANY_REL]: { relation: Array.from(allRelIds).map(id => ({ id })) }
      }})
    }
  }

  // 2. Company dedup by sauron_org_id
  const allCompanies = notionQueryAll_(notion, companiesDbId, { page_size: 100 })
  const companiesByOrgId = new Map()
  for (const c of allCompanies) {
    const oid = notionGetRichText_(c, NOTION_COMPANY_PROP_SAURON_ORG_ID)
    if (!oid) continue
    if (!companiesByOrgId.has(oid)) companiesByOrgId.set(oid, [])
    companiesByOrgId.get(oid).push(c)
  }

  for (const [oid, pages] of companiesByOrgId) {
    if (pages.length <= 1) continue

    // Keep first (oldest created), archive rest
    const sorted = pages.sort((a, b) => {
      const aDate = a.created_time || ""
      const bDate = b.created_time || ""
      return aDate.localeCompare(bDate)
    })

    const keeper = sorted[0]
    for (let i = 1; i < sorted.length; i++) {
      // Re-parent contacts from loser to keeper
      try {
        const loserContacts = notionQueryAll_(notion, contactsDbId, {
          filter: { property: NOTION_CONTACT_PROP_COMPANY_REL, relation: { contains: sorted[i].id } },
          page_size: 100
        })
        for (const contact of loserContacts) {
          const rel = contact.properties && contact.properties[NOTION_CONTACT_PROP_COMPANY_REL]
          const ids = (rel && rel.relation) ? rel.relation.map(r => r.id).filter(id => id !== sorted[i].id) : []
          ids.push(keeper.id)
          notionUpdatePage_(notion, contact.id, { properties: {
            [NOTION_CONTACT_PROP_COMPANY_REL]: { relation: [...new Set(ids)].map(id => ({ id })) }
          }})
        }
      } catch (e) {
        Logger.log(`crm_dedup: failed to re-parent contacts for company ${sorted[i].id}: ${e.message}`)
      }
      notionPatch_(notion, `/pages/${sorted[i].id}`, { archived: true })
      companiesArchived++
    }
  }

  // 3. Company domain dupes (log only — weaker signal)
  const companiesByDomain = new Map()
  for (const c of allCompanies) {
    const d = notionGetRichText_(c, NOTION_COMPANY_PROP_DOMAIN)
    if (!d || CRM_PERSONAL_DOMAINS.has(d)) continue
    if (!companiesByDomain.has(d)) companiesByDomain.set(d, [])
    companiesByDomain.get(d).push(c)
  }
  for (const [domain, pages] of companiesByDomain) {
    if (pages.length <= 1) continue
    const names = pages.map(p => notionGetTitle_(p, NOTION_COMPANY_PROP_NAME) || "(untitled)")
    domainDupes.push({ domain, count: pages.length, names })
  }

  if (domainDupes.length > 0) {
    Logger.log(`crm_dedup: ${domainDupes.length} potential domain duplicates (manual review):`)
    for (const d of domainDupes) Logger.log(`  ${d.domain}: ${d.count} companies — ${d.names.join(", ")}`)
  }

  Logger.log(`crm_dedup done: contacts_archived=${contactsArchived}, companies_archived=${companiesArchived}, domain_dupes_flagged=${domainDupes.length}`)
  return { contactsArchived, companiesArchived, domainDupesFlagged: domainDupes.length }
}


/** =========================
 * ONE-TIME MIGRATION: re-stamp sauron_org_id from app_org_id to Clerk org_id
 * Run once manually, then delete this function.
 * ========================= */

function migrate_sauron_org_id_to_clerk_org_id() {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(SHEET_CANON_ORGS)
  if (!sh) throw new Error(`Missing sheet: ${SHEET_CANON_ORGS}`)
  const rows = readSheetObjects_(sh, 1)

  // Build reverse map: app_org_id → org_id (Clerk)
  const appToClerk = new Map()
  for (const r of rows) {
    const appOrgId = str_(r.app_org_id)
    const clerkOrgId = str_(r.org_id)
    if (appOrgId && clerkOrgId) appToClerk.set(appOrgId, clerkOrgId)
  }

  Logger.log(`migrate: built reverse map with ${appToClerk.size} entries`)

  const notion = notionClient_()
  const props = PropertiesService.getScriptProperties()
  const companiesDbId = mustGetProp_(props, PROP_NOTION_COMPANIES_DB_ID)

  const allCompanies = notionQueryAll_(notion, companiesDbId, {
    filter: { property: NOTION_COMPANY_PROP_SAURON_ORG_ID, rich_text: { is_not_empty: true } },
    page_size: 100
  })

  let migrated = 0
  let skipped = 0
  let alreadyClerk = 0

  for (const company of allCompanies) {
    const currentId = notionGetRichText_(company, NOTION_COMPANY_PROP_SAURON_ORG_ID)
    if (!currentId) { skipped++; continue }

    // Already a Clerk org_id — no migration needed
    if (/^org_[A-Za-z0-9]/.test(currentId)) { alreadyClerk++; continue }

    const clerkId = appToClerk.get(currentId)
    if (!clerkId) {
      Logger.log(`migrate: no Clerk org_id found for app_org_id=${currentId}, skipping`)
      skipped++
      continue
    }

    notionUpdatePage_(notion, company.id, { properties: {
      [NOTION_COMPANY_PROP_SAURON_ORG_ID]: { rich_text: [{ type: "text", text: { content: clerkId } }] }
    }})
    migrated++
  }

  Logger.log(`migrate done: migrated=${migrated}, already_clerk=${alreadyClerk}, skipped=${skipped}, total=${allCompanies.length}`)
  return { migrated, alreadyClerk, skipped, total: allCompanies.length }
}

/** =========================
 * TRIGGER SETUP
 * ========================= */

function setup_daily_notion_crm_triggers() {
  const newFns = [
    "link_unlinked_orgs",
    "link_unlinked_contacts",
    "build_crm_lists",
    "sync_segments_to_notion",
    "crm_dedup_cleanup"
  ]
  const oldFns = [
    "notion_link_unlinked_companies_to_sauron",
    "notion_link_unlinked_contacts_to_sauron",
    "notion_push_upsale_targets_from_org_info",
    "sync_orgs_to_notion",
    "sync_users_to_notion",
    ...newFns
  ]

  // Clean up old + current triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (oldFns.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t)
  })

  // Link first (discovers new connections), then sync (pushes data)
  ScriptApp.newTrigger("link_unlinked_orgs").timeBased().everyDays(1).atHour(7).nearMinute(10).create()
  ScriptApp.newTrigger("link_unlinked_contacts").timeBased().everyDays(1).atHour(7).nearMinute(20).create()
  ScriptApp.newTrigger("build_crm_lists").timeBased().everyDays(1).atHour(7).nearMinute(45).create()
  ScriptApp.newTrigger("sync_segments_to_notion").timeBased().everyDays(1).atHour(7).nearMinute(50).create()
  ScriptApp.newTrigger("crm_dedup_cleanup").timeBased().everyDays(1).atHour(8).nearMinute(0).create()

  Logger.log("Daily CRM sync triggers created.")
}

/** =========================
 * NOTION HTTP CLIENT
 * ========================= */

function notionClient_() {
  const props = PropertiesService.getScriptProperties()
  const token = mustGetProp_(props, PROP_NOTION_TOKEN)
  const version = props.getProperty(PROP_NOTION_VERSION) || "2025-09-03"
  return { token, version, baseUrl: "https://api.notion.com/v1", _dsCache: {} }
}

/**
 * Resolve a database_id to its first data_source_id (cached per client).
 * Required for API version 2025-09-03+ with multi-source databases.
 */
function notionResolveDataSourceId_(notion, databaseId) {
  if (notion._dsCache[databaseId]) return notion._dsCache[databaseId]
  const db = notionFetch_(notion, "get", `/databases/${databaseId}`, null)
  const ds = db && db.data_sources && db.data_sources.length ? db.data_sources[0].id : null
  if (!ds) throw new Error(`Could not resolve data_source_id for database ${databaseId}`)
  notion._dsCache[databaseId] = ds
  return ds
}

function notionQueryAll_(notion, databaseId, body) {
  const dsId = notionResolveDataSourceId_(notion, databaseId)
  const out = []
  let cursor = null
  while (true) {
    const payload = Object.assign({}, body || {})
    if (cursor) payload.start_cursor = cursor
    const res = notionPost_(notion, `/data_sources/${dsId}/query`, payload)
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
    method,
    contentType: "application/json",
    muteHttpExceptions: true,
    headers: { Authorization: `Bearer ${notion.token}`, "Notion-Version": notion.version },
    payload: payload ? JSON.stringify(payload) : undefined
  }

  for (let attempt = 0; attempt < 7; attempt++) {
    if (attempt > 0) Utilities.sleep(1000 * attempt)

    const resp = UrlFetchApp.fetch(url, options)
    const code = resp.getResponseCode()
    const text = resp.getContentText()

    // Rate limit or bandwidth quota — back off and retry
    if (code === 429 || (code === 502 && text.includes("Bandwidth")) || code === 503) {
      Utilities.sleep(2000 * (attempt + 1))
      continue
    }
    if (code >= 200 && code < 300) {
      // Baseline throttle: ~3 req/sec max
      Utilities.sleep(350)
      return text ? JSON.parse(text) : {}
    }
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

function normEmail_(v) {
  const s = String(v || "").trim().toLowerCase()
  if (!s) return ""
  return s.replace(/\+[^@]+(?=@)/, "")
}

function looksLikeEmail_(s) {
  const t = String(s || "").trim().toLowerCase()
  if (!t) return false
  return t.includes("@")
}

function getDomain_(email) {
  const e = String(email || "").trim().toLowerCase()
  const i = e.indexOf("@")
  return i < 0 ? "" : e.slice(i + 1).trim()
}

function parseDate_(v) {
  if (!v) return ""
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return ""
    return v.toISOString()
  }
  const s = String(v).trim()
  if (!s) return ""
  const d = new Date(s)
  if (isNaN(d.getTime())) return ""
  return d.toISOString()
}

/** Returns just YYYY-MM-DD for Notion date-only columns. */
function parseDateOnly_(v) {
  const iso = parseDate_(v)
  return iso ? iso.slice(0, 10) : ""
}

function tryUpgradeNotionContactTitle_(notion, contactId, desiredName) {
  const name = String(desiredName || "").trim()
  if (!contactId || !name) return

  const page = notionGetPage_(notion, contactId)
  const currentTitle = notionGetTitleAny_(page) || ""

  if (!currentTitle || looksLikeEmail_(currentTitle)) {
    notionUpdatePage_(notion, contactId, { properties: {
      [NOTION_CONTACT_PROP_NAME]: {
        title: [{ type: "text", text: { content: name } }]
      }
    }})
  }
}

/**
 * Call PDL enrichment webhook for a newly created Notion company page.
 * Requires WEBHOOK_SECRET in Script Properties.
 */
function enrichWithPdl_(notionPageId) {
  const secret = PropertiesService.getScriptProperties().getProperty('WEBHOOK_SECRET')
  if (!secret) {
    Logger.log('enrichWithPdl_: WEBHOOK_SECRET not set, skipping')
    return
  }
  try {
    const res = UrlFetchApp.fetch('https://pingdomsteward-production.up.railway.app/admin/enrich/pdl', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-webhook-secret': secret },
      payload: JSON.stringify({ pageId: notionPageId }),
      muteHttpExceptions: true
    })
    const code = res.getResponseCode()
    if (code === 200) {
      Logger.log(`  PDL enriched: ${notionPageId}`)
    } else if (code === 404) {
      Logger.log(`  PDL no match: ${notionPageId}`)
    } else {
      Logger.log(`  PDL error ${code}: ${notionPageId}`)
    }
  } catch (e) {
    Logger.log(`  PDL fetch failed: ${e.message}`)
  }
}
