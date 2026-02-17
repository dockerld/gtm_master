/**************************************************************
 * Publish "The Good Stuff"
 *
 * Copies a polished dashboard + table view from "The Ring"
 * into an external spreadsheet tab named "The Good Stuff".
 **************************************************************/

const GOOD_STUFF_CFG = {
  SOURCE: {
    RING_SHEET: 'The Ring',
    HEADER_ROW: 3,
    DATA_START_ROW: 4,
    START_COL: 2 // Ring table starts at column B
  },
  TARGET: {
    SPREADSHEET_ID: '147yUcx8Eb7LE-jhAALwfmddIIOoBYcvEJpXhJR0c8qc',
    SHEET_NAME: 'The Good Stuff'
  },
  KPI: {
    COMBINED_COL: 2,        // B2:D2 in The Ring
    PAID_ONLY_COL: 6,       // F2:H2 in The Ring
    PAID_WITH_FIRST_COL: 10 // J2:L2 in The Ring
  },
  ANNUAL_GOAL_ARR: 1000000,
  TABLE_START_ROW: 13
}

function publish_the_good_stuff() {
  return GOOD_lockWrapCompat_('publish_the_good_stuff', () => {
    const sourceSs = SpreadsheetApp.getActive()
    const ring = sourceSs.getSheetByName(GOOD_STUFF_CFG.SOURCE.RING_SHEET)
    if (!ring) throw new Error(`Missing source sheet: ${GOOD_STUFF_CFG.SOURCE.RING_SHEET}`)

    const targetSs = SpreadsheetApp.openById(GOOD_STUFF_CFG.TARGET.SPREADSHEET_ID)
    const out = GOOD_getOrCreateSheet_(targetSs, GOOD_STUFF_CFG.TARGET.SHEET_NAME)

    const metrics = GOOD_readRingMetrics_(ring)
    const goals = GOOD_readGoals_(sourceSs, metrics.combined.arr)
    const table = GOOD_readRingTable_(ring)

    GOOD_resetAndStyleCanvas_(out)
    GOOD_writeHeader_(out)
    GOOD_writeGoalStrip_(out, goals)
    GOOD_writeKpiCards_(out, metrics)
    GOOD_writeTable_(out, GOOD_STUFF_CFG.TABLE_START_ROW, table.headers, table.rows)

    return {
      rows_in: table.rows.length,
      rows_out: table.rows.length
    }
  })
}

function GOOD_readRingMetrics_(ringSheet) {
  function readTriplet(colStart) {
    const vals = ringSheet.getRange(2, colStart, 1, 3).getValues()[0]
    return {
      arr: GOOD_num_(vals[0]),
      subscriptions: GOOD_num_(vals[1]),
      totalSeats: GOOD_num_(vals[2])
    }
  }

  return {
    combined: readTriplet(GOOD_STUFF_CFG.KPI.COMBINED_COL),
    paidOnly: readTriplet(GOOD_STUFF_CFG.KPI.PAID_ONLY_COL),
    paidWithFirstPayment: readTriplet(GOOD_STUFF_CFG.KPI.PAID_WITH_FIRST_COL)
  }
}

function GOOD_readGoals_(sourceSs, currentArr) {
  const monthlyGoal = GOOD_getMonthlyArrGoal_(sourceSs)
  const annualGoal = GOOD_STUFF_CFG.ANNUAL_GOAL_ARR

  const monthlyPct = monthlyGoal > 0 ? (currentArr / monthlyGoal) : 0
  const annualPct = annualGoal > 0 ? (currentArr / annualGoal) : 0

  return {
    currentArr: GOOD_num_(currentArr),
    monthlyGoal: GOOD_num_(monthlyGoal),
    annualGoal: GOOD_num_(annualGoal),
    monthlyPct: GOOD_clamp01_(monthlyPct),
    annualPct: GOOD_clamp01_(annualPct)
  }
}

function GOOD_getMonthlyArrGoal_(sourceSs) {
  if (typeof getMonthlyArrGoalFromGoalsSheet_ === 'function') {
    try {
      const n = Number(getMonthlyArrGoalFromGoalsSheet_())
      if (isFinite(n) && n > 0) return n
    } catch (e) {}
  }

  const sh = sourceSs.getSheetByName('Goals')
  if (!sh) return 0

  const lastCol = sh.getLastColumn()
  if (lastCol < 1) return 0

  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
  const arrRow = sh.getRange(2, 1, 1, lastCol).getValues()[0]

  const tz = Session.getScriptTimeZone()
  const thisMonthKey = Utilities.formatDate(new Date(), tz, 'MMM-yyyy')

  let idx = -1
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim() === thisMonthKey) {
      idx = i
      break
    }
  }

  if (idx === -1) {
    for (let c = arrRow.length - 1; c >= 0; c--) {
      const v = Number(arrRow[c])
      if (isFinite(v) && v > 0) {
        idx = c
        break
      }
    }
  }

  if (idx < 0) return 0
  const n = Number(arrRow[idx])
  return isFinite(n) ? n : 0
}

function GOOD_readRingTable_(ringSheet) {
  const maxColsFromStart = ringSheet.getLastColumn() - GOOD_STUFF_CFG.SOURCE.START_COL + 1
  if (maxColsFromStart <= 0) return { headers: [], rows: [] }

  const headerRaw = ringSheet
    .getRange(GOOD_STUFF_CFG.SOURCE.HEADER_ROW, GOOD_STUFF_CFG.SOURCE.START_COL, 1, maxColsFromStart)
    .getValues()[0]

  const headerWidth = GOOD_contiguousWidth_(headerRaw)
  if (headerWidth <= 0) return { headers: [], rows: [] }

  const headers = headerRaw.slice(0, headerWidth).map(h => String(h || '').trim())

  const lastRow = ringSheet.getLastRow()
  const numRows = Math.max(0, lastRow - GOOD_STUFF_CFG.SOURCE.DATA_START_ROW + 1)
  const rows = numRows > 0
    ? ringSheet
      .getRange(GOOD_STUFF_CFG.SOURCE.DATA_START_ROW, GOOD_STUFF_CFG.SOURCE.START_COL, numRows, headerWidth)
      .getValues()
      .filter(r => r.some(v => String(v || '').trim() !== ''))
    : []

  return { headers, rows }
}

function GOOD_resetAndStyleCanvas_(sheet) {
  sheet.clear()
  try { sheet.setHiddenGridlines(true) } catch (e) {}

  for (let i = 1; i <= 12; i++) {
    sheet.setColumnWidth(i, 180)
  }
}

function GOOD_writeHeader_(sheet) {
  const titleRange = sheet.getRange(1, 1, 1, 12)
  titleRange.merge()
  titleRange
    .setValue('THE GOOD STUFF')
    .setFontSize(24)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#FFFFFF')
    .setFontColor('#111111')

  const tz = Session.getScriptTimeZone()
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss')
  const subtitle = sheet.getRange(2, 1, 1, 12)
  subtitle.merge()
  subtitle
    .setValue(`Live snapshot from The Ring • Updated ${stamp}`)
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#FFFFFF')
    .setFontColor('#374151')

  sheet.setRowHeight(1, 42)
  sheet.setRowHeight(2, 24)
}

function GOOD_writeGoalStrip_(sheet, goals) {
  const labels = ['Current ARR', 'Monthly Goal', 'Annual Goal']
  const values = [goals.currentArr, goals.monthlyGoal, goals.annualGoal]
  const starts = [1, 5, 9]

  for (let i = 0; i < starts.length; i++) {
    const c = starts[i]
    const labelR = sheet.getRange(9, c, 1, 4)
    labelR.merge()
    labelR
      .setValue(labels[i])
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#3B82F6')
      .setFontColor('#F8FAFC')

    const valueR = sheet.getRange(10, c, 1, 4)
    valueR.merge()
    valueR
      .setValue(values[i])
      .setNumberFormat('$#,##0.00')
      .setFontSize(16)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#F8FAFC')
      .setFontColor('#0F172A')
  }

  const progress = sheet.getRange(11, 1, 1, 12)
  progress.merge()
  progress
    .setValue(
      `Progress • Monthly: ${GOOD_pctText_(goals.monthlyPct)}    |    Annual: ${GOOD_pctText_(goals.annualPct)}`
    )
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setBackground('#E0F2FE')
    .setFontColor('#0C4A6E')
}

function GOOD_writeKpiCards_(sheet, metrics) {
  GOOD_writeKpiCard_(sheet, 4, 1, 'Paid + Promo Trial', metrics.combined, '#2563EB')
  GOOD_writeKpiCard_(sheet, 4, 5, 'Paid Only', metrics.paidOnly, '#7C3AED')
  GOOD_writeKpiCard_(sheet, 4, 9, 'Paid + First Payment', metrics.paidWithFirstPayment, '#EA580C')
}

function GOOD_writeKpiCard_(sheet, topRow, startCol, title, data, accent) {
  const card = sheet.getRange(topRow, startCol, 4, 4)
  card
    .setBorder(true, true, true, true, true, true, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment('middle')

  const titleR = sheet.getRange(topRow, startCol, 1, 4)
  titleR.merge()
  titleR
    .setValue(title)
    .setBackground(accent)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')

  const arrR = sheet.getRange(topRow + 1, startCol, 1, 4)
  arrR.merge()
  arrR
    .setValue(data.arr || 0)
    .setNumberFormat('$#,##0.00')
    .setFontSize(20)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#F8FAFC')
    .setFontColor('#0F172A')

  const labelSubsR = sheet.getRange(topRow + 2, startCol, 1, 2)
  labelSubsR.merge()
  labelSubsR
    .setValue('Subscriptions')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#EEF2FF')
    .setFontColor('#334155')

  const labelSeatsR = sheet.getRange(topRow + 2, startCol + 2, 1, 2)
  labelSeatsR.merge()
  labelSeatsR
    .setValue('Seats')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#EEF2FF')
    .setFontColor('#334155')

  const valueSubsR = sheet.getRange(topRow + 3, startCol, 1, 2)
  valueSubsR.merge()
  valueSubsR
    .setValue(data.subscriptions || 0)
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setBackground('#FFFFFF')
    .setFontColor('#0F172A')
    .setNumberFormat('0')

  const valueSeatsR = sheet.getRange(topRow + 3, startCol + 2, 1, 2)
  valueSeatsR.merge()
  valueSeatsR
    .setValue(data.totalSeats || 0)
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setBackground('#FFFFFF')
    .setFontColor('#0F172A')
    .setNumberFormat('0')
}

function GOOD_writeTable_(sheet, startRow, headers, rows) {
  if (!headers || !headers.length) return

  const headR = sheet.getRange(startRow, 1, 1, headers.length)
  headR.setValues([headers])
  headR
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#2563EB')
    .setFontColor('#F8FAFC')

  if (rows && rows.length) {
    const dataR = sheet.getRange(startRow + 1, 1, rows.length, headers.length)
    dataR.setValues(rows)
    dataR.setVerticalAlignment('middle')

    const h = headers.map(x => String(x || '').trim().toLowerCase())
    const col = (name) => h.indexOf(name) + 1

    const cAmount = col('amount')
    const cMrr = col('mrr')
    const cArr = col('arr')
    const cDisc = col('discount %')
    const cSeats = col('seats')
    const cFirstPay = col('first payment at')

    if (cAmount > 0) sheet.getRange(startRow + 1, cAmount, rows.length, 1).setNumberFormat('$#,##0.00')
    if (cMrr > 0) sheet.getRange(startRow + 1, cMrr, rows.length, 1).setNumberFormat('$#,##0.00')
    if (cArr > 0) sheet.getRange(startRow + 1, cArr, rows.length, 1).setNumberFormat('$#,##0.00')
    if (cDisc > 0) sheet.getRange(startRow + 1, cDisc, rows.length, 1).setNumberFormat('0.##%')
    if (cSeats > 0) sheet.getRange(startRow + 1, cSeats, rows.length, 1).setNumberFormat('0')
    if (cFirstPay > 0) sheet.getRange(startRow + 1, cFirstPay, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss')

    try { sheet.getRange(startRow, 1, rows.length + 1, headers.length).applyRowBanding() } catch (e) {}
  }

  sheet.setFrozenRows(startRow)
  const usedCols = Math.max(headers.length, 12)
  for (let c = 1; c <= usedCols; c++) {
    sheet.setColumnWidth(c, 180)
  }
}

function GOOD_contiguousWidth_(headerRowArray) {
  let width = 0
  for (let i = 0; i < headerRowArray.length; i++) {
    const v = String(headerRowArray[i] || '').trim()
    if (!v) break
    width += 1
  }
  return width
}

function GOOD_pctText_(pct01) {
  const n = GOOD_clamp01_(pct01) * 100
  return `${n.toFixed(1)}%`
}

function GOOD_num_(v) {
  const n = Number(v)
  return isFinite(n) ? n : 0
}

function GOOD_clamp01_(n) {
  const x = Number(n)
  if (!isFinite(x)) return 0
  return Math.max(0, Math.min(1, x))
}

function GOOD_getOrCreateSheet_(ss, name) {
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function GOOD_lockWrapCompat_(name, fn) {
  if (typeof lockWrap === 'function') {
    try { return lockWrap(name, fn) } catch (e) { return lockWrap(fn) }
  }
  const lock = LockService.getScriptLock()
  const ok = lock.tryLock(300000)
  if (!ok) throw new Error(`Could not acquire lock: ${name}`)
  try { return fn() } finally { lock.releaseLock() }
}
