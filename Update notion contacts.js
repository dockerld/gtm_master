/**
 * Repair job:
 * Fix Notion Contacts whose Name/title looks like an email, by using canon_users name.
 *
 * Safe:
 * - Only updates if current title looks like an email (contains "@")
 * - Only updates if we find a non-empty canonical name
 *
 * Run manually whenever you want:
 *   notion_fix_contact_names_from_canon()
 */
function notion_fix_contact_names_from_canon() {
  return notionLockWrap_("notion_fix_contact_names_from_canon", () => {
    const props = PropertiesService.getScriptProperties()
    const notion = notionClient_()
    const contactsDbId = mustGetProp_(props, PROP_NOTION_CONTACTS_DB_ID)

    const sauronIndex = buildSauronIndex_()

    // Query contacts where Name contains "@"
    // (This is the reliable "email-ish title" flag)
    const candidates = notionQueryAll_(notion, contactsDbId, {
      filter: {
        property: NOTION_CONTACT_PROP_NAME,
        title: { contains: "@" }
      },
      page_size: 100
    }).slice(0, 500) // safety cap

    Logger.log(`Fix names: found ${candidates.length} contacts with @ in title`)

    let fixed = 0
    let skipped = 0
    const nowIso = new Date().toISOString()

    for (const c of candidates) {
      const contactId = c.id
      if (!contactId) continue

      const currentTitle = notionGetTitleAny_(c) || ""
      if (!looksLikeEmail_(currentTitle)) {
        skipped += 1
        continue
      }

      const email = notionGetEmail_(c, NOTION_CONTACT_PROP_EMAIL)
      const emailKey = normEmail_(email)
      if (!emailKey) {
        skipped += 1
        continue
      }

      const hit = sauronIndex.userByEmailKey.get(emailKey)
      const betterName = bestDisplayNameFromCanon_(hit)

      // must have a real name and it must not look like an email
      if (!betterName || looksLikeEmail_(betterName)) {
        skipped += 1
        continue
      }

      // ✅ Update title
      const patch = { properties: {} }
      patch.properties[NOTION_CONTACT_PROP_NAME] = {
        title: [{ type: "text", text: { content: betterName } }]
      }
      notionUpdatePage_(notion, contactId, patch)

      // ✅ Optional: update mapping sheet too (so you can audit repairs)
      // Only if map sheet exists and headers match your schema.
      try {
        ensureMapSheetUser_()
        upsertMapRowUser_({
          notion_contact_id: contactId,
          notion_contact_name: betterName,
          notion_company_id: "",
          notion_company_name: "",
          email: email,
          matched_email_key: emailKey,
          org_id: hit && hit.orgId ? hit.orgId : "",
          user_id: hit && hit.userId ? hit.userId : "",
          link_source: "repair_title_from_canon",
          status: "fixed_name",
          linked_at: "",
          last_checked_at: nowIso,
          notes: `updated title from "${currentTitle}" -> "${betterName}"`
        })
      } catch (e) {
        // ignore map failures, repair should still work
      }

      fixed += 1
    }

    Logger.log(`Fix names done. fixed=${fixed}, skipped=${skipped}`)
    return { rows_in: candidates.length, rows_out: fixed }
  })
}