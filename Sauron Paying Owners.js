/**************************************************************
 * render_sauron_paying_owners()
 *
 * Creates/overwrites a sheet with users from "Sauron" who are:
 * - Paying = TRUE
 * - Hierarchy = Owner (normalized; "org:owner" -> "owner")
 *
 * Output columns (in order):
 * Email, Name, Org Name, Days with Ping, Days since last Login,
 * Meetings Recorded, Hours Recorded, Meeting Notes Synced,
 * Action Items Synced, Logged in #, # of Clients, PM,
 * Cal Connected, Email Connected, Org Sign Up Date, Org Members
 *
 * Notes:
 * - Reads Sauron headers from Row 3 starting at Col B
 * - Reads Sauron data from Row 4 starting at Col B
 * - Overwrites output sheet each run
 **************************************************************/

const SAURON_OWNER_CFG = {
  SOURCE_SHEET: 'Sauron',
  OUT_SHEET: 'Sauron Paying Owners',

  // Sauron layout
  HEADER_ROW: 3,
  START_COL: 2,      // Col B
  DATA_START_ROW: 4,

  FILTER_HEADERS: {
    PAYING: 'Paying',
    HIERARCHY: 'Hierarchy'
  },

  OUT_HEADERS: [
    'Email',
    'Name',
    'Org Name',
    'Days with Ping',
    'Days since last Login',
    'Meetings Recorded',
    'Hours Recorded',
    'Meeting Notes Synced',
    'Action Items Synced',
    'Logged in #',
    '# of Clients',
    'PM',
    'Cal Connected',
    'Email Connected',
    'Org Sign Up Date',
    'Org Members'
  ],

  WRITE_CHUNK: 3000
}

function render_sauron_paying_owners() {
  return lockWrapCompat_('render_sauron_paying_owners', () => {
    const t0 = new Date()
    const ss = SpreadsheetApp.getActive()

    const src = ss.getSheetByName(SAURON_OWNER_CFG.SOURCE_SHEET)
    if (!src) throw new Error(`Source sheet not found: ${SAURON_OWNER_CFG.SOURCE_SHEET}`)

    const outSheet = getOrCreateSheetCompat_(ss, SAURON_OWNER_CFG.OUT_SHEET)

    // ---- Read Sauron headers (contiguous from B3) ----
    const maxColsFromStart = src.getLastColumn() - SAURON_OWNER_CFG.START_COL + 1
    if (maxColsFromStart <= 0) throw new Error('Sauron has no columns in the expected region')

    const rawHeaderRow = src
      .getRange(SAURON_OWNER_CFG.HEADER_ROW, SAURON_OWNER_CFG.START_COL, 1, maxColsFromStart)
      .getValues()[0]
      .map(h => String(h || '').trim())

    const headerWidth = SAURON_ownerContigHeaderWidth_(rawHeaderRow)
    if (headerWidth <= 0) throw new Error('Sauron header row appears empty starting at B3')

    const srcHeaders = rawHeaderRow.slice(0, headerWidth)
    const headerMap = SAURON_ownerHeaderMap_(srcHeaders)

    // ---- Required headers ----
    const cPaying = SAURON_ownerCol_(headerMap, SAURON_OWNER_CFG.FILTER_HEADERS.PAYING)
    const cHierarchy = SAURON_ownerCol_(headerMap, SAURON_OWNER_CFG.FILTER_HEADERS.HIERARCHY)

    SAURON_OWNER_CFG.OUT_HEADERS.forEach(h => SAURON_ownerCol_(headerMap, h))

    // ---- Read data ----
    const lastRow = src.getLastRow()
    if (lastRow < SAURON_OWNER_CFG.DATA_START_ROW) {
      SAURON_ownerWriteOut_(outSheet, SAURON_OWNER_CFG.OUT_HEADERS, [])
      return
    }

    const numRows = lastRow - SAURON_OWNER_CFG.DATA_START_ROW + 1
    const srcData = src.getRange(
      SAURON_OWNER_CFG.DATA_START_ROW,
      SAURON_OWNER_CFG.START_COL,
      numRows,
      headerWidth
    ).getValues()

    const outRows = []

    for (const row of srcData) {
      const paying = SAURON_ownerToBool_(row[cPaying])
      if (!paying) continue

      const roleRaw = String(row[cHierarchy] || '').trim()
      const role = SAURON_ownerNormalizeRole_(roleRaw)
      if (role.toLowerCase() !== 'owner') continue

      outRows.push(SAURON_OWNER_CFG.OUT_HEADERS.map(h => row[headerMap[h.toLowerCase()]]))
    }

    SAURON_ownerWriteOut_(outSheet, SAURON_OWNER_CFG.OUT_HEADERS, outRows)

    if (typeof writeSyncLog === 'function') {
      writeSyncLog(
        'render_sauron_paying_owners',
        'ok',
        numRows,
        outRows.length,
        (new Date() - t0) / 1000,
        ''
      )
    }
  })
}

/* =========================
 * Local helpers
 * ========================= */

function SAURON_ownerContigHeaderWidth_(headerRowArray) {
  let w = 0
  for (let i = 0; i < headerRowArray.length; i++) {
    if (!headerRowArray[i]) break
    w++
  }
  return w
}

function SAURON_ownerHeaderMap_(headers) {
  const map = {}
  headers.forEach((h, i) => {
    const key = String(h || '').trim().toLowerCase()
    if (!key) return
    if (map[key] == null) map[key] = i
  })
  return map
}

function SAURON_ownerCol_(map, headerName) {
  const key = String(headerName || '').trim().toLowerCase()
  if (!key || map[key] == null) {
    throw new Error(`Sauron missing required header: ${headerName}`)
  }
  return map[key]
}

function SAURON_ownerToBool_(v) {
  if (v === true) return true
  if (v === false) return false
  const s = String(v || '').toLowerCase().trim()
  return s === 'true' || s === 'yes' || s === '1'
}

function SAURON_ownerNormalizeRole_(role) {
  const s = String(role || '').trim()
  if (!s) return ''
  const idx = s.lastIndexOf(':')
  return idx >= 0 ? s.slice(idx + 1).trim() : s
}

function SAURON_ownerWriteOut_(sheet, headers, rows) {
  sheet.clearContents()
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
  sheet.setFrozenRows(1)
  if (rows && rows.length) {
    batchSetValuesCompat_(sheet, 2, 1, rows, SAURON_OWNER_CFG.WRITE_CHUNK)
  }
  sheet.autoResizeColumns(1, headers.length)
}
