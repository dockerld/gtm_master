/**************************************************************
 * Webhook Receiver (Google Apps Script Web App) — RUNS IMMEDIATELY
 *
 * ✅ What this version does:
 * - Validates shared secret (query param ?secret=... OR JSON body.secret)
 * - Logs EVERYTHING into webhook_inbox (including full payload, clipped)
 * - Runs the pipeline immediately (no queue):
 *    1) calcrm_notion_calendar_import_from_camden()
 *    2) notion_link_unlinked_contacts_to_sauron()
 *    3) notion_link_unlinked_companies_to_sauron()
 *
 * ✅ Notes:
 * - This WILL take longer than "queue" mode. If Notion times out, switch back to queue.
 * - To reduce timeout risk, make sure your calendar import has small limits (you already did).
 *
 * Sheets:
 * - webhook_inbox (logs)
 *
 * Script Properties required:
 * - WEBHOOK_SHARED_SECRET
 **************************************************************/

const WEBHOOK_RUN = {
  SHEET: "webhook_inbox",
  HEADER: [
    "received_at",
    "request_id",
    "status",              // running | done | error | auth_error
    "ok",
    "error",
    "secret_source",
    "query_params_json",
    "raw_body",
    "parsed_json",
    "steps_json",
    "processed_at",
    "duration_ms"
  ],
  MAX_RAW_BODY_CHARS: 50000
}

/**
 * Web App entry
 * - Runs pipeline immediately (no queue)
 * - Logs everything in webhook_inbox
 */
function doPost(e) {
  const requestId = Utilities.getUuid()
  const receivedAt = new Date().toISOString()

  let status = "running"
  let ok = false
  let error = ""
  let secretSource = "none"
  let stepsJson = "[]"
  let processedAt = ""
  let durationMs = ""

  const t0 = Date.now()

  const queryParams = (e && e.parameter) ? e.parameter : {}
  const queryParamsJson = safeJson_(queryParams)

  const rawBody = (e && e.postData && typeof e.postData.contents === "string")
    ? e.postData.contents
    : ""

  const rawBodyClipped = clip_(rawBody, WEBHOOK_RUN.MAX_RAW_BODY_CHARS)
  const parsed = safeParseJson_(rawBody)
  const parsedJson = safeJson_(parsed.value)

  try {
    ensureWebhookSheet_()

    // --- AUTH (query param or body) ---
    const expected = mustGetScriptProp_("WEBHOOK_SHARED_SECRET")

    const fromQuery = String(queryParams.secret || "").trim()
    const fromBody =
      parsed.ok
        ? String(
            (parsed.value && parsed.value.secret) ||
            (parsed.value && parsed.value.data && parsed.value.data.secret) ||
            ""
          ).trim()
        : ""

    const provided = fromQuery || fromBody
    if (fromQuery) secretSource = "query"
    else if (fromBody) secretSource = "json_body"

    if (!provided) {
      status = "auth_error"
      error = "Unauthorized: missing secret (?secret=... or JSON body.secret)"
      durationMs = String(Date.now() - t0)
      processedAt = new Date().toISOString()
      writeWebhookRow_(receivedAt, requestId, status, ok, error, secretSource, queryParamsJson, rawBodyClipped, parsedJson, "[]", processedAt, durationMs)
      return jsonResponse_({ ok: false, request_id: requestId, error })
    }

    if (provided !== expected) {
      status = "auth_error"
      error = "Unauthorized: invalid secret"
      durationMs = String(Date.now() - t0)
      processedAt = new Date().toISOString()
      writeWebhookRow_(receivedAt, requestId, status, ok, error, secretSource, queryParamsJson, rawBodyClipped, parsedJson, "[]", processedAt, durationMs)
      return jsonResponse_({ ok: false, request_id: requestId, error })
    }

    // --- RUN PIPELINE ---
    const steps = []

    steps.push(runStep_("calcrm_notion_calendar_import_from_camden", () => {
      if (typeof calcrm_notion_calendar_import_from_camden !== "function") {
        throw new Error("Missing function: calcrm_notion_calendar_import_from_camden")
      }
      return calcrm_notion_calendar_import_from_camden()
    }))

    steps.push(runStep_("notion_link_unlinked_contacts_to_sauron", () => {
      if (typeof notion_link_unlinked_contacts_to_sauron !== "function") {
        throw new Error("Missing function: notion_link_unlinked_contacts_to_sauron")
      }
      return notion_link_unlinked_contacts_to_sauron()
    }))

    steps.push(runStep_("notion_link_unlinked_companies_to_sauron", () => {
      if (typeof notion_link_unlinked_companies_to_sauron !== "function") {
        throw new Error("Missing function: notion_link_unlinked_companies_to_sauron")
      }
      return notion_link_unlinked_companies_to_sauron()
    }))

    // Determine overall result
    const allOk = steps.every(s => s && s.ok === true)
    ok = allOk
    status = allOk ? "done" : "error"
    if (!allOk) {
      const firstFail = steps.find(s => s && s.ok === false)
      error = firstFail ? String(firstFail.error || "unknown error") : "unknown error"
    }

    stepsJson = JSON.stringify(steps)
    processedAt = new Date().toISOString()
    durationMs = String(Date.now() - t0)

    // Log final row
    writeWebhookRow_(
      receivedAt,
      requestId,
      status,
      ok,
      error,
      secretSource,
      queryParamsJson,
      rawBodyClipped,
      parsedJson,
      stepsJson,
      processedAt,
      durationMs
    )

    return jsonResponse_({
      ok,
      request_id: requestId,
      status,
      duration_ms: Number(durationMs),
      steps
    })
  } catch (err) {
    status = "error"
    ok = false
    error = String(err && err.message ? err.message : err)
    stepsJson = stepsJson || "[]"
    processedAt = new Date().toISOString()
    durationMs = String(Date.now() - t0)

    try {
      ensureWebhookSheet_()
      writeWebhookRow_(
        receivedAt,
        requestId,
        status,
        ok,
        error,
        secretSource,
        queryParamsJson,
        rawBodyClipped,
        parsedJson,
        stepsJson,
        processedAt,
        durationMs
      )
    } catch (_) {}

    return jsonResponse_({ ok: false, request_id: requestId, status, error, duration_ms: Number(durationMs) })
  }
}

/* =========================
 * Helpers
 * ========================= */

function runStep_(name, fn) {
  const t0 = Date.now()
  try {
    const result = fn()
    return { name, ok: true, ms: Date.now() - t0, result: result == null ? null : result }
  } catch (err) {
    return { name, ok: false, ms: Date.now() - t0, error: String(err && err.message ? err.message : err) }
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON)
}

function ensureWebhookSheet_() {
  const ss = SpreadsheetApp.getActive()
  let sh = ss.getSheetByName(WEBHOOK_RUN.SHEET)
  if (!sh) sh = ss.insertSheet(WEBHOOK_RUN.SHEET)

  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, WEBHOOK_RUN.HEADER.length).setValues([WEBHOOK_RUN.HEADER])
    sh.setFrozenRows(1)
    return
  }

  // enforce header
  sh.getRange(1, 1, 1, WEBHOOK_RUN.HEADER.length).setValues([WEBHOOK_RUN.HEADER])
  sh.setFrozenRows(1)
}

function writeWebhookRow_(receivedAt, requestId, status, ok, error, secretSource, queryParamsJson, rawBody, parsedJson, stepsJson, processedAt, durationMs) {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(WEBHOOK_RUN.SHEET)
  if (!sh) throw new Error("Missing webhook_inbox sheet")

  sh.appendRow([
    receivedAt,
    requestId,
    status,
    ok === true,
    error || "",
    secretSource || "",
    queryParamsJson || "{}",
    rawBody || "",
    parsedJson || "{}",
    stepsJson || "[]",
    processedAt || "",
    durationMs || ""
  ])
}

function mustGetScriptProp_(key) {
  const v = PropertiesService.getScriptProperties().getProperty(key)
  if (!v) throw new Error(`Missing Script Property: ${key}`)
  return v
}

function safeParseJson_(s) {
  try {
    if (!s) return { ok: true, value: {} }
    return { ok: true, value: JSON.parse(s) }
  } catch (e) {
    return { ok: false, value: { _parse_error: String(e && e.message ? e.message : e) } }
  }
}

function safeJson_(obj) {
  try { return JSON.stringify(obj == null ? {} : obj) } catch (e) { return "{}" }
}

function clip_(s, max) {
  const str = String(s || "")
  if (str.length <= max) return str
  return str.slice(0, max) + `\n...[clipped ${str.length - max} chars]`
}