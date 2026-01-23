/**************************************************************
 * render_conversion_onboarding_stats()
 *
 * Builds "Conversion & Onboarding stats" with two tables stacked:
 * - Conversion stats (top)
 * - Onboarding stats (bottom)
 **************************************************************/

const COMBINED_STATS_CFG = {
  SHEET_NAME: 'Conversion & Onboarding stats',
  CONV_TITLE: 'Conversion stats',
  ONB_TITLE: 'Onboarding stats',
  GAP_ROWS: 2
}

function render_conversion_onboarding_stats() {
  return CONV_lockWrapCompat_('render_conversion_onboarding_stats', () => {
    return COMBINED_renderConversionOnboarding_({ logStepName: 'render_conversion_onboarding_stats' })
  })
}

function COMBINED_renderConversionOnboarding_(opts) {
  const t0 = new Date()
  const ss = SpreadsheetApp.getActive()

  const shOut = getOrCreateSheetCompat_(ss, COMBINED_STATS_CFG.SHEET_NAME)
  const shOrgs = ss.getSheetByName(CONV_CFG.INPUTS.CLERK_ORGS)
  const shOrgInfo = ss.getSheetByName(CONV_CFG.INPUTS.ORG_INFO)
  const shPosthog = ss.getSheetByName(ONB_CFG.POSTHOG_SHEET)
  const shClerk = ss.getSheetByName(ONB_CFG.CLERK_SHEET)

  if (!shOrgs) throw new Error(`Missing input sheet: ${CONV_CFG.INPUTS.CLERK_ORGS}`)
  if (!shOrgInfo) throw new Error(`Missing input sheet: ${CONV_CFG.INPUTS.ORG_INFO}`)
  if (!shPosthog) throw new Error(`Missing input sheet: ${ONB_CFG.POSTHOG_SHEET}`)
  if (!shClerk) throw new Error(`Missing input sheet: ${ONB_CFG.CLERK_SHEET}`)

  const tz = Session.getScriptTimeZone()

  const statsByMonth = CONV_collectStatsByMonth_(shOrgs, shOrgInfo, tz)
  const convRows = CONV_buildRows_(statsByMonth)

  const createdByEmailKey = ONB_buildCreatedAtIndex_(shClerk)
  const onbStats = ONB_buildStats_(shPosthog, createdByEmailKey, tz)
  const onbRows = ONB_buildRows_(onbStats.byMonth, onbStats.summary)

  shOut.clearContents()

  let row = 1
  shOut.getRange(row, 1).setValue(COMBINED_STATS_CFG.CONV_TITLE).setFontWeight('bold')
  row += 1

  const convHeaderRow = row
  shOut.getRange(convHeaderRow, 1, 1, CONV_CFG.HEADERS.length).setValues([CONV_CFG.HEADERS])
  row += 1

  const convDataStart = row
  if (convRows.length) {
    shOut.getRange(convDataStart, 1, convRows.length, CONV_CFG.HEADERS.length).setValues(convRows)
  }
  row += convRows.length

  row += COMBINED_STATS_CFG.GAP_ROWS

  shOut.getRange(row, 1).setValue(COMBINED_STATS_CFG.ONB_TITLE).setFontWeight('bold')
  row += 1

  const onbHeaderRow = row
  shOut.getRange(onbHeaderRow, 1, 1, ONB_CFG.HEADERS.length).setValues([ONB_CFG.HEADERS])
  row += 1

  const onbDataStart = row
  if (onbRows.length) {
    shOut.getRange(onbDataStart, 1, onbRows.length, ONB_CFG.HEADERS.length).setValues(onbRows)
  }

  if (convRows.length) {
    CONV_applyFormatsAt_(shOut, convHeaderRow, convDataStart, convRows.length)
  } else {
    CONV_applyFormatsAt_(shOut, convHeaderRow, convDataStart, 0)
  }

  if (onbRows.length) {
    ONB_applyFormatsAt_(shOut, onbHeaderRow, onbDataStart, onbRows.length, onbStats.monthCount, onbStats.summaryCount)
  } else {
    ONB_applyFormatsAt_(shOut, onbHeaderRow, onbDataStart, 0, onbStats.monthCount, onbStats.summaryCount)
  }

  shOut.autoResizeColumns(1, Math.max(CONV_CFG.HEADERS.length, ONB_CFG.HEADERS.length))

  const seconds = (new Date() - t0) / 1000
  if (typeof writeSyncLog === 'function' && opts && opts.logStepName) {
    writeSyncLog(opts.logStepName, 'ok', convRows.length + onbRows.length, '', seconds, '')
  }

  return { rows_out: convRows.length + onbRows.length }
}

function COMBINED_removeOldSheets_(ss, names) {
  ;(names || []).forEach(name => {
    const sh = ss.getSheetByName(name)
    if (sh) ss.deleteSheet(sh)
  })
}
