/**
 * Toggle "Paying is checked" in all filter views on "Sauron".
 *
 * - Header row is row 3, header text is "Paying".
 * - Works on FILTER VIEWS using Advanced Sheets API.
 * - Uses DocumentProperties to remember ON/OFF state.
 * - Visual indicator: cell C1 on Sauron sheet
 */

const SAURON_PAYING_FILTER_HEADER_ROW_INDEX = 3
const SAURON_PAYING_FILTER_HEADER_TEXT = 'Paying'
const SAURON_PAYING_FILTER_SHEET_NAME = 'Sauron'
const SAURON_PAYING_FILTER_TOGGLE_PROP_KEY = 'SAURON_PAYING_FILTER_TOGGLE_STATE' // "ON" or "OFF"

function SAURON_toggleCheckedFilter() {
  const ss = SpreadsheetApp.getActive()
  const sheet = ss.getSheetByName(SAURON_PAYING_FILTER_SHEET_NAME)

  if (!sheet) throw new Error(`Sheet "${SAURON_PAYING_FILTER_SHEET_NAME}" not found.`)

  // Find the "Paying" column in header row 3
  const lastColumn = sheet.getLastColumn()
  const headerValues = sheet
    .getRange(SAURON_PAYING_FILTER_HEADER_ROW_INDEX, 1, 1, lastColumn)
    .getValues()[0]

  const payingColIndex1Based = headerValues.indexOf(SAURON_PAYING_FILTER_HEADER_TEXT) + 1
  if (payingColIndex1Based <= 0) {
    throw new Error(
      `Could not find a column with header "${SAURON_PAYING_FILTER_HEADER_TEXT}" in row ${SAURON_PAYING_FILTER_HEADER_ROW_INDEX}`
    )
  }

  // Sheets API uses 0-based column index as key
  const payingColIndex0Based = payingColIndex1Based - 1
  const colKey = String(payingColIndex0Based)

  const spreadsheetId = ss.getId()
  const sheetId = sheet.getSheetId()

  // Determine current toggle state from properties
  const props = PropertiesService.getDocumentProperties()
  const currentState = props.getProperty(SAURON_PAYING_FILTER_TOGGLE_PROP_KEY) || 'OFF'
  const turningOn = currentState === 'OFF' // OFF -> ON, ON -> OFF

  // Visual indicator: change only cell C1
  const indicatorCell = sheet.getRange('C1')
  indicatorCell.setBackground(turningOn ? '#C6EFCE' : '#FFEB9C') // green on, yellow off

  // Get filterViews for this sheet via Advanced Sheets API
  const spreadsheet = Sheets.Spreadsheets.get(spreadsheetId, {
    fields: 'sheets(properties(sheetId),filterViews)'
  })

  const targetSheet = (spreadsheet.sheets || []).find(
    s => s.properties && s.properties.sheetId === sheetId
  )

  if (!targetSheet || !targetSheet.filterViews || targetSheet.filterViews.length === 0) {
    SpreadsheetApp.getUi().alert(`No filter views found on sheet "${SAURON_PAYING_FILTER_SHEET_NAME}".`)
    return
  }

  const requests = []

  targetSheet.filterViews.forEach(fv => {
    const criteria = fv.criteria || {}

    if (turningOn) {
      // Turn ON: show only checked -> hide FALSE and blanks
      criteria[colKey] = { hiddenValues: ['FALSE', ''] }
    } else {
      // Turn OFF: allow checked + unchecked + blanks
      criteria[colKey] = { hiddenValues: [] }
    }

    fv.criteria = criteria

    requests.push({
      updateFilterView: {
        filter: fv,
        fields: 'criteria'
      }
    })
  })

  if (requests.length) {
    Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId)
  }

  // Flip and store the new state
  props.setProperty(SAURON_PAYING_FILTER_TOGGLE_PROP_KEY, turningOn ? 'ON' : 'OFF')
}