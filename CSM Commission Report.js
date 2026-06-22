/**************************************************************
 * render_csm_commission_report()
 *
 * Builds "CSM Commission Report" from arr_waterfall_facts.
 *
 * Two sections:
 * 1. Org-level detail: each org's New/Upgrade/Downgrade per month
 * 2. Monthly summary with commission calc:
 *    - Commission structure:
 *      60% weight: 105% NRR target
 *      30% weight: 25% conversion rate target
 *      10% weight: Team's monthly ARR quota
 *
 * Output sheet: "CSM Commission Report"
 **************************************************************/

const CSM_CFG = {
  OUT_SHEET: 'CSM Commission Report',
  SOURCE_SHEET: 'arr_waterfall_facts',
  GOALS_SHEET: 'Goals',
  CONVERSION_STATS_SHEET: 'Conversion stats',

  EXTERNAL_SPREADSHEET_ID: '1jhxP7GdSfO78zwABHL175pX6-1M2rLTp9hbPcIuRHhg',
  EXTERNAL_SHEET_NAME: 'Commission',

  DETAIL_HEADERS: [
    'month',
    'org_id',
    'org_name',
    'plan_name',
    'billing_frequency',
    'paid_cohort_month',
    'metric',
    'amount'
  ],

  SUMMARY_HEADERS: [
    'month',
    'som_total',
    'new_arr',
    'upgrade_arr',
    'downgrade_arr',
    'churn_arr',
    'eom_total',
    'net_change',
    'nrr_percent',
    'nrr_target_met',
    'new_customers_count',
    'paid_orgs_count'
  ],

  // Commission targets
  NRR_TARGET: 1.05,
  CONVERSION_TARGET: 0.25,
  NRR_WEIGHT: 0.60,
  CONVERSION_WEIGHT: 0.30,
  QUOTA_WEIGHT: 0.10,

  // Input cell labels
  INPUT_LABEL_BG: '#FFF9E6',
  INPUT_CELL_BG: '#FFFDE7',

  CURRENCY_FMT: '$#,##0.00',
  PCT_FMT: '0.0%',
  COUNT_FMT: '0'
}

function render_csm_commission_report() {
  return CSM_lockWrap_('render_csm_commission_report', () => {
    const t0 = new Date()
    const ss = SpreadsheetApp.getActive()

    const src = ss.getSheetByName(CSM_CFG.SOURCE_SHEET)
    if (!src) throw new Error('Missing sheet: ' + CSM_CFG.SOURCE_SHEET)

    const srcRows = CSM_readSheetObjects_(src, 1)
    if (!srcRows.length) throw new Error('No data in ' + CSM_CFG.SOURCE_SHEET)

    // ── Build org-level detail rows (New, Upgrade, Downgrade only) ──
    const detailRows = []
    const monthAgg = new Map()

    for (const r of srcRows) {
      const month = CSM_str_(r.month)
      const metric = CSM_str_(r.metric)
      const amount = CSM_num_(r.amount)
      const orgId = CSM_str_(r.org_id)
      const orgName = CSM_str_(r.org_name)

      if (!month) continue

      // Aggregate by month
      if (!monthAgg.has(month)) {
        monthAgg.set(month, { som: 0, newArr: 0, upgrade: 0, downgrade: 0, churn: 0, eom: 0, newCount: 0, paidOrgs: new Set() })
      }
      const agg = monthAgg.get(month)

      if (metric === 'SOM') {
        agg.som += amount
        if (amount > 0 && orgId) agg.paidOrgs.add(orgId)
      } else if (metric === 'EOM') {
        agg.eom += amount
        if (amount > 0 && orgId) agg.paidOrgs.add(orgId)
      } else if (metric === 'New' && amount > 0) {
        agg.newArr += amount
        agg.newCount += 1
        if (orgId) agg.paidOrgs.add(orgId)
        detailRows.push([month, orgId, orgName, CSM_str_(r.plan_name), CSM_str_(r.billing_frequency), CSM_str_(r.paid_cohort_month), 'New', amount])
      } else if (metric === 'Upgrade' && amount > 0) {
        agg.upgrade += amount
        detailRows.push([month, orgId, orgName, CSM_str_(r.plan_name), CSM_str_(r.billing_frequency), CSM_str_(r.paid_cohort_month), 'Upgrade', amount])
      } else if (metric === 'Downgrade' && amount > 0) {
        agg.downgrade += amount
        detailRows.push([month, orgId, orgName, CSM_str_(r.plan_name), CSM_str_(r.billing_frequency), CSM_str_(r.paid_cohort_month), 'Downgrade', amount])
      } else if (metric === 'Churn' && amount > 0) {
        agg.churn += amount
        detailRows.push([month, orgId, orgName, CSM_str_(r.plan_name), CSM_str_(r.billing_frequency), CSM_str_(r.paid_cohort_month), 'Churn', amount])
      }
    }

    // Sort detail rows by month desc, then metric, then amount desc
    const metricOrder = { 'New': 0, 'Upgrade': 1, 'Downgrade': 2, 'Churn': 3 }
    detailRows.sort((a, b) => {
      if (a[0] !== b[0]) return b[0].localeCompare(a[0])
      const ma = metricOrder[a[6]] || 99
      const mb = metricOrder[b[6]] || 99
      if (ma !== mb) return ma - mb
      return b[7] - a[7]
    })

    // ── Build monthly summary ──
    const months = Array.from(monthAgg.keys()).sort()
    const summaryRows = months.map(month => {
      const a = monthAgg.get(month)
      const netChange = a.newArr + a.upgrade - a.downgrade - a.churn
      // NRR excludes new customers: (SOM + Expansion - Contraction - Churn) / SOM
      const nrr = a.som > 0 ? ((a.som + a.upgrade - a.downgrade - a.churn) / a.som) : 0
      const nrrTargetMet = nrr >= CSM_CFG.NRR_TARGET ? 'Yes' : 'No'
      return [
        month,
        a.som,
        a.newArr,
        a.upgrade,
        a.downgrade,
        a.churn,
        a.eom,
        netChange,
        nrr,
        nrrTargetMet,
        a.newCount,
        a.paidOrgs.size
      ]
    })
    // Most recent month first
    summaryRows.reverse()

    // ── Write output directly to external sheet ──
    const targetSs = SpreadsheetApp.openById(CSM_CFG.EXTERNAL_SPREADSHEET_ID)
    const out = CSM_getOrCreateSheet_(targetSs, CSM_CFG.EXTERNAL_SHEET_NAME)

    // Preserve user-entered comp inputs and commission config before clearing
    const savedInputs = CSM_readSavedInputs_(out)
    out.clear()
    try { out.clearNotes() } catch (e) {}

    // Section 1: Monthly Summary
    let row = 1
    out.getRange(row, 1).setValue('MONTHLY SUMMARY').setFontWeight('bold').setFontSize(12)
    row += 1
    out.getRange(row, 1, 1, CSM_CFG.SUMMARY_HEADERS.length).setValues([CSM_CFG.SUMMARY_HEADERS])
      .setFontWeight('bold').setBackground('#F3F4F6')
    row += 1

    if (summaryRows.length) {
      out.getRange(row, 1, summaryRows.length, 1).setNumberFormat('@')
      out.getRange(row, 1, summaryRows.length, CSM_CFG.SUMMARY_HEADERS.length).setValues(summaryRows)

      // Format currency columns (som, new, upgrade, downgrade, churn, eom, net_change)
      for (const col of [2, 3, 4, 5, 6, 7, 8]) {
        out.getRange(row, col, summaryRows.length, 1).setNumberFormat(CSM_CFG.CURRENCY_FMT)
      }
      // NRR percent
      out.getRange(row, 9, summaryRows.length, 1).setNumberFormat(CSM_CFG.PCT_FMT)
      // Counts
      for (const col of [11, 12]) {
        out.getRange(row, col, summaryRows.length, 1).setNumberFormat(CSM_CFG.COUNT_FMT)
      }
      row += summaryRows.length
    }

    // ── Monthly conversion rates from org_subscription_info ──
    // conversion = orgs that converted in month / orgs whose trial ended in month
    const convByMonth = CSM_buildMonthlyConversionRates_(ss)

    // ARR quota from Goals sheet (same source as Ring email)
    const goalsData = CSM_getGoalsData_(ss)

    // Section 2: Comp inputs & commission config (editable yellow cells, preserved across runs)
    row += 2
    out.getRange(row, 1).setValue('COMP & COMMISSION CONFIG').setFontWeight('bold').setFontSize(12)
    row += 1
    out.getRange(row, 1, 1, 3).setValues([['Setting', 'Value', '']])
      .setFontWeight('bold').setBackground('#F3F4F6')
    row += 1

    const inputStartRow = row
    const inputDefs = [
      { label: 'Monthly Base Salary',       key: 'base_salary',       fmt: CSM_CFG.CURRENCY_FMT, dflt: 0 },
      { label: 'Monthly Commission Potential', key: 'comm_potential', fmt: CSM_CFG.CURRENCY_FMT, dflt: 0 },
      { label: 'NRR Target',                key: 'nrr_target',        fmt: CSM_CFG.PCT_FMT,      dflt: CSM_CFG.NRR_TARGET },
      { label: 'NRR Weight (% of commission)', key: 'nrr_weight',    fmt: CSM_CFG.PCT_FMT,      dflt: CSM_CFG.NRR_WEIGHT },
      { label: 'Conversion Target',         key: 'conv_target',       fmt: CSM_CFG.PCT_FMT,      dflt: CSM_CFG.CONVERSION_TARGET },
      { label: 'Conversion Weight (% of commission)', key: 'conv_weight', fmt: CSM_CFG.PCT_FMT, dflt: CSM_CFG.CONVERSION_WEIGHT },
      { label: 'Quota Weight (% of commission)', key: 'quota_weight', fmt: CSM_CFG.PCT_FMT,      dflt: CSM_CFG.QUOTA_WEIGHT }
    ]
    for (let i = 0; i < inputDefs.length; i++) {
      const def = inputDefs[i]
      const saved = savedInputs[def.key]
      const val = (saved !== undefined && saved !== '' && saved !== null) ? saved : def.dflt
      out.getRange(row + i, 1).setValue(def.label).setFontWeight('bold')
      out.getRange(row + i, 2).setValue(val)
        .setBackground(CSM_CFG.INPUT_CELL_BG)
        .setNumberFormat(def.fmt)
    }
    row += inputDefs.length

    // Cell references for formulas
    const baseSalaryCell = 'B' + inputStartRow
    const commPotentialCell = 'B' + (inputStartRow + 1)
    const nrrTargetCell = 'B' + (inputStartRow + 2)
    const nrrWeightCell = 'B' + (inputStartRow + 3)
    const convTargetCell = 'B' + (inputStartRow + 4)
    const convWeightCell = 'B' + (inputStartRow + 5)
    const quotaWeightCell = 'B' + (inputStartRow + 6)

    // Section 3: Commission Calculators (current month + last month)
    const latestMonth = months.length ? months[months.length - 1] : null
    const prevMonth = months.length >= 2 ? months[months.length - 2] : null

    const monthsToRender = []
    if (latestMonth) monthsToRender.push({ key: latestMonth, label: 'CURRENT MONTH' })
    if (prevMonth) monthsToRender.push({ key: prevMonth, label: 'LAST MONTH' })

    for (const m of monthsToRender) {
      const agg = monthAgg.get(m.key)
      if (!agg) continue

      const nrr = agg.som > 0
        ? ((agg.som + agg.upgrade - agg.downgrade - agg.churn) / agg.som)
        : 0
      const convData = convByMonth.get(m.key) || { converted: 0, trialsEnded: 0, rate: 0 }
      const eomArr = agg.eom
      // Quota: how much growth needed from last month's EOM to hit this month's cumulative quota
      const thisMonthQuota = (goalsData.quotaByMonth.get(m.key) || 0)
      const prevMonthEom = agg.som  // SOM = last month's EOM
      const quotaGrowthTarget = thisMonthQuota - prevMonthEom
      const actualArrGrowth = agg.eom - prevMonthEom

      row += 2
      out.getRange(row, 1).setValue(m.label + ' — ' + m.key).setFontWeight('bold').setFontSize(12)
      row += 1

      const calcHeaders = ['Component', 'Weight', 'Target', 'Actual', 'Attainment', 'Commission Earned']
      out.getRange(row, 1, 1, calcHeaders.length).setValues([calcHeaders])
        .setFontWeight('bold').setBackground('#F3F4F6')
      row += 1

      const calcStartRow = row

      // NRR — weight & target reference config cells; attainment = (actual - 1) / (target - 1)
      out.getRange(row, 1).setValue('NRR')
      out.getRange(row, 2).setFormula('=' + nrrWeightCell).setNumberFormat(CSM_CFG.PCT_FMT)
      out.getRange(row, 3).setFormula('=' + nrrTargetCell).setNumberFormat(CSM_CFG.PCT_FMT)
      out.getRange(row, 4).setValue(nrr).setNumberFormat(CSM_CFG.PCT_FMT)
      // Attainment: MIN(1, MAX(0, (actual - 1) / (target - 1)))
      out.getRange(row, 5).setFormula('=MIN(1,MAX(0,(D' + row + '-1)/(C' + row + '-1)))')
        .setNumberFormat(CSM_CFG.PCT_FMT)
      out.getRange(row, 6).setFormula('=E' + row + '*B' + row + '*' + commPotentialCell)
        .setNumberFormat(CSM_CFG.CURRENCY_FMT)
      row += 1

      // Conversion Rate — weight & target reference config cells
      const convLabel = 'Conversion Rate (' + convData.converted + '/' + convData.trialsEnded + ' trials ended)'
      out.getRange(row, 1).setValue(convLabel)
      out.getRange(row, 2).setFormula('=' + convWeightCell).setNumberFormat(CSM_CFG.PCT_FMT)
      out.getRange(row, 3).setFormula('=' + convTargetCell).setNumberFormat(CSM_CFG.PCT_FMT)
      out.getRange(row, 4).setValue(convData.rate).setNumberFormat(CSM_CFG.PCT_FMT)
      // Attainment: MIN(1, MAX(0, actual / target))
      out.getRange(row, 5).setFormula('=MIN(1,MAX(0,D' + row + '/C' + row + '))')
        .setNumberFormat(CSM_CFG.PCT_FMT)
      out.getRange(row, 6).setFormula('=E' + row + '*B' + row + '*' + commPotentialCell)
        .setNumberFormat(CSM_CFG.CURRENCY_FMT)
      row += 1

      // ARR Quota — growth needed from last month's EOM to hit cumulative quota
      const quotaLabel = 'Monthly ARR Quota (last EOM: ' +
        CSM_fmtMoney_(prevMonthEom) + ' / quota: ' + CSM_fmtMoney_(thisMonthQuota) + ')'
      out.getRange(row, 1).setValue(quotaLabel)
      out.getRange(row, 2).setFormula('=' + quotaWeightCell).setNumberFormat(CSM_CFG.PCT_FMT)
      out.getRange(row, 3).setValue(quotaGrowthTarget).setNumberFormat(CSM_CFG.CURRENCY_FMT)
      out.getRange(row, 4).setValue(actualArrGrowth).setNumberFormat(CSM_CFG.CURRENCY_FMT)
      // Attainment: MIN(1, actual growth / target growth)
      out.getRange(row, 5).setFormula('=IF(C' + row + '>0,MIN(1,D' + row + '/C' + row + '),0)')
        .setNumberFormat(CSM_CFG.PCT_FMT)
      out.getRange(row, 6).setFormula('=E' + row + '*B' + row + '*' + commPotentialCell)
        .setNumberFormat(CSM_CFG.CURRENCY_FMT)
      row += 1

      // Totals
      row += 1
      out.getRange(row, 1, 1, 6).setBackground('#E8F5E9')
      out.getRange(row, 1).setValue('TOTAL COMMISSION EARNED').setFontWeight('bold')
      out.getRange(row, 6).setFormula('=SUM(F' + calcStartRow + ':F' + (calcStartRow + 2) + ')')
        .setNumberFormat(CSM_CFG.CURRENCY_FMT).setFontWeight('bold')
      row += 1

      out.getRange(row, 1, 1, 6).setBackground('#E8F5E9')
      out.getRange(row, 1).setValue('TOTAL MONTHLY COMP (Base + Commission)').setFontWeight('bold')
      out.getRange(row, 6).setFormula('=' + baseSalaryCell + '+F' + (row - 1))
        .setNumberFormat(CSM_CFG.CURRENCY_FMT).setFontWeight('bold')
      row += 1
    }

    // Section 4: Next month targets (based on latest month)
    const latestAgg = latestMonth ? monthAgg.get(latestMonth) : null
    const latestConv = latestMonth ? (convByMonth.get(latestMonth) || { rate: 0 }) : { rate: 0 }
    const latestNrr = latestAgg && latestAgg.som > 0
      ? ((latestAgg.som + latestAgg.upgrade - latestAgg.downgrade - latestAgg.churn) / latestAgg.som)
      : 0
    const nextSom = latestAgg ? latestAgg.eom : 0
    const nrrExpansionNeeded = nextSom * (CSM_CFG.NRR_TARGET - 1)
    const currentNetExpansion = latestAgg
      ? (latestAgg.upgrade - latestAgg.downgrade - latestAgg.churn)
      : 0
    const nrrGap = nrrExpansionNeeded - currentNetExpansion

    row += 2
    out.getRange(row, 1).setValue('NEXT MONTH TARGETS').setFontWeight('bold').setFontSize(12)
    row += 1
    const targetHeaders = ['Metric', 'Current', 'Target', 'Gap']
    out.getRange(row, 1, 1, targetHeaders.length).setValues([targetHeaders])
      .setFontWeight('bold').setBackground('#F3F4F6')
    row += 1

    out.getRange(row, 1).setValue('NRR')
    out.getRange(row, 2).setValue(latestNrr).setNumberFormat(CSM_CFG.PCT_FMT)
    out.getRange(row, 3).setValue(CSM_CFG.NRR_TARGET).setNumberFormat(CSM_CFG.PCT_FMT)
    out.getRange(row, 4).setValue(Math.max(0, CSM_CFG.NRR_TARGET - latestNrr)).setNumberFormat(CSM_CFG.PCT_FMT)
    row += 1

    out.getRange(row, 1).setValue('Net expansion needed (upgrade - downgrade - churn)')
    out.getRange(row, 2).setValue(currentNetExpansion).setNumberFormat(CSM_CFG.CURRENCY_FMT)
    out.getRange(row, 3).setValue(nrrExpansionNeeded).setNumberFormat(CSM_CFG.CURRENCY_FMT)
    out.getRange(row, 4).setValue(Math.max(0, nrrGap)).setNumberFormat(CSM_CFG.CURRENCY_FMT)
    row += 1

    out.getRange(row, 1).setValue('Projected next month SOM')
    out.getRange(row, 2).setValue(nextSom).setNumberFormat(CSM_CFG.CURRENCY_FMT)
    row += 1

    out.getRange(row, 1).setValue('Conversion Rate')
    out.getRange(row, 2).setValue(latestConv.rate).setNumberFormat(CSM_CFG.PCT_FMT)
    out.getRange(row, 3).setValue(CSM_CFG.CONVERSION_TARGET).setNumberFormat(CSM_CFG.PCT_FMT)
    out.getRange(row, 4).setValue(Math.max(0, CSM_CFG.CONVERSION_TARGET - latestConv.rate)).setNumberFormat(CSM_CFG.PCT_FMT)
    row += 1

    // Next month quota: growth needed from latest EOM to next month's cumulative quota
    const now = new Date()
    const nextMonthKey = String(now.getFullYear()) + '-' + String(now.getMonth() + 2).padStart(2, '0')
    const nextMonthQuota = goalsData.quotaByMonth.get(nextMonthKey) || 0
    const latestEom = latestAgg ? latestAgg.eom : 0
    const nextGrowthTarget = nextMonthQuota - latestEom
    out.getRange(row, 1).setValue('Monthly ARR Quota (EOM→quota)')
    out.getRange(row, 2).setValue(latestEom).setNumberFormat(CSM_CFG.CURRENCY_FMT)
    out.getRange(row, 3).setValue(nextMonthQuota).setNumberFormat(CSM_CFG.CURRENCY_FMT)
    out.getRange(row, 4).setValue(Math.max(0, nextGrowthTarget)).setNumberFormat(CSM_CFG.CURRENCY_FMT)
    row += 1

    // Section 5: Org-level detail
    row += 2
    out.getRange(row, 1).setValue('ORG-LEVEL DETAIL (New / Upgrade / Downgrade / Churn)').setFontWeight('bold').setFontSize(12)
    row += 1
    out.getRange(row, 1, 1, CSM_CFG.DETAIL_HEADERS.length).setValues([CSM_CFG.DETAIL_HEADERS])
      .setFontWeight('bold').setBackground('#F3F4F6')
    row += 1

    if (detailRows.length) {
      out.getRange(row, 1, detailRows.length, 1).setNumberFormat('@')
      out.getRange(row, 1, detailRows.length, CSM_CFG.DETAIL_HEADERS.length).setValues(detailRows)
      // Format amount column
      out.getRange(row, 8, detailRows.length, 1).setNumberFormat(CSM_CFG.CURRENCY_FMT)
    }

    out.setFrozenRows(0)
    out.autoResizeColumns(1, Math.max(CSM_CFG.SUMMARY_HEADERS.length, CSM_CFG.DETAIL_HEADERS.length))

    const seconds = (new Date() - t0) / 1000
    if (typeof writeSyncLog === 'function') {
      writeSyncLog('render_csm_commission_report', 'ok', srcRows.length, detailRows.length + summaryRows.length, seconds, '')
    }

    return { rows_in: srcRows.length, detail_rows: detailRows.length, summary_rows: summaryRows.length }
  })
}


// ── Helpers ──

function CSM_readSheetObjects_(sheet, headerRow) {
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < headerRow + 1 || lastCol < 1) return []

  const header = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => CSM_str_(h).toLowerCase().replace(/\s+/g, '_'))
  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues()

  return data.map(r => {
    const obj = {}
    header.forEach((h, i) => { if (h) obj[h] = r[i] })
    return obj
  })
}

function CSM_readSavedInputs_(sheet) {
  // Read existing comp/commission config values before the sheet gets cleared
  const out = {}
  if (!sheet) return out
  const lastRow = sheet.getLastRow()
  const lastCol = sheet.getLastColumn()
  if (lastRow < 2 || lastCol < 2) return out

  const data = sheet.getRange(1, 1, lastRow, Math.min(lastCol, 2)).getValues()
  const keyMap = {
    'monthly base salary': 'base_salary',
    'monthly commission potential': 'comm_potential',
    'nrr target': 'nrr_target',
    'nrr weight (% of commission)': 'nrr_weight',
    'conversion target': 'conv_target',
    'conversion weight (% of commission)': 'conv_weight',
    'quota weight (% of commission)': 'quota_weight'
  }

  for (const r of data) {
    const label = String(r[0] || '').trim().toLowerCase()
    const mapped = keyMap[label]
    if (mapped && r[1] !== '' && r[1] !== null && r[1] !== undefined) {
      out[mapped] = r[1]
    }
  }
  return out
}

function CSM_getOrCreateSheet_(ss, name) {
  if (typeof getOrCreateSheetCompat_ === 'function') return getOrCreateSheetCompat_(ss, name)
  const sh = ss.getSheetByName(name)
  return sh || ss.insertSheet(name)
}

function CSM_lockWrap_(name, fn) {
  if (typeof lockWrapCompat_ === 'function') return lockWrapCompat_(name, fn)
  const lock = LockService.getScriptLock()
  if (!lock.tryLock(300000)) throw new Error('Could not acquire lock: ' + name)
  try { return fn() } finally { lock.releaseLock() }
}

function CSM_toDateValue_(v) {
  if (!v) return ''
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : v
  const s = CSM_str_(v)
  if (!s) return ''
  const d = new Date(s)
  return isNaN(d.getTime()) ? s : d
}

function CSM_isInternalOrg_(r) {
  const name = CSM_str_(r.org_name || r.org_slug || r.org).toLowerCase()
  if (name.includes('ping') || name.includes('test')) return true
  const email = CSM_str_(r.customer_email || r.billing_email || r.email).toLowerCase()
  if (email.endsWith('@pingassistant.com')) return true
  return false
}

function CSM_buildExcludedSubIds_(ss) {
  const out = new Set()
  const sh = ss.getSheetByName('Manual Stripe Changes')
  if (!sh) return out
  const rows = CSM_readSheetObjects_(sh, 1)
  const EXCLUDE_REASONS = new Set(['internal', 'partner', 'free subscription'])
  for (const r of rows) {
    const reason = CSM_str_(r.exclude_reason).toLowerCase()
    if (!EXCLUDE_REASONS.has(reason)) continue
    const subId = CSM_str_(r.subscription_id || r.stripe_subscription_id || r.subscription)
    if (subId) out.add(subId)
  }
  return out
}

function CSM_str_(v) {
  if (v === null || v === undefined) return ''
  return String(v).trim()
}

function CSM_num_(v) {
  const n = Number(v)
  return isFinite(n) ? n : 0
}

function CSM_fmtMoney_(v) {
  const n = CSM_num_(v)
  return '$' + n.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 })
}

function CSM_buildMonthlyConversionRates_(ss) {
  // Uses shared canon_orgs-based org list (filtered, enriched with first_payment_at)
  const orgs = buildConversionOrgList_(ss)
  const out = new Map()

  // Separate buckets: trials ended by trial_ends_at month, conversions by first_payment_at month
  const trialBuckets = new Map()
  const convBuckets = new Map()

  for (const org of orgs) {
    const trialMonth = CONVUTIL_toMonthKey_(org.trial_ends_at)
    if (trialMonth) {
      trialBuckets.set(trialMonth, (trialBuckets.get(trialMonth) || 0) + 1)
    }

    const payMonth = CONVUTIL_toMonthKey_(org.first_payment_at)
    if (payMonth) {
      convBuckets.set(payMonth, (convBuckets.get(payMonth) || 0) + 1)
    }
  }

  const allMonths = new Set([...trialBuckets.keys(), ...convBuckets.keys()])
  for (const month of allMonths) {
    const trialsEnded = trialBuckets.get(month) || 0
    const converted = convBuckets.get(month) || 0
    out.set(month, {
      trialsEnded,
      converted,
      rate: trialsEnded > 0 ? (converted / trialsEnded) : 0
    })
  }
  return out
}

function CSM_parseToMonthKey_(v) {
  if (!v) return ''
  let d = null
  if (v instanceof Date) {
    d = isNaN(v.getTime()) ? null : v
  } else {
    const s = CSM_str_(v)
    if (!s) return ''
    d = new Date(s)
    if (isNaN(d.getTime())) return ''
  }
  if (!d) return ''
  const y = d.getFullYear()
  const m = String(d.getMonth() + 1).padStart(2, '0')
  return y + '-' + m
}

function CSM_getGoalsData_(ss) {
  // Returns quota ARR per month key so we can compute deltas
  const sh = ss.getSheetByName(CSM_CFG.GOALS_SHEET)
  const empty = { quotaByMonth: new Map(), goalByMonth: new Map() }
  if (!sh) return empty

  const lastCol = sh.getLastColumn()
  if (lastCol < 2) return empty

  const quotaByMonth = CSM_readGoalsRow_(sh, 12, 13, lastCol)
  const goalByMonth = CSM_readGoalsRow_(sh, 6, 7, lastCol)

  return { quotaByMonth, goalByMonth }
}

function CSM_readGoalsRow_(sh, headerRowNum, valueRowNum, lastCol) {
  const out = new Map()
  const maxRow = sh.getLastRow()
  if (maxRow < valueRowNum || lastCol < 1) return out

  const headers = sh.getRange(headerRowNum, 1, 1, lastCol).getValues()[0]
  const values = sh.getRange(valueRowNum, 1, 1, lastCol).getValues()[0]

  for (let i = 0; i < headers.length; i++) {
    let hKey = ''
    if (headers[i] instanceof Date && !isNaN(headers[i].getTime())) {
      const d = headers[i]
      hKey = String(d.getFullYear()) + '-' + String(d.getMonth() + 1).padStart(2, '0')
    } else {
      const h = CSM_str_(headers[i])
      if (/^\d{4}-\d{2}/.test(h)) hKey = h.slice(0, 7)
    }
    if (hKey) out.set(hKey, CSM_num_(values[i]))
  }
  return out
}

function CSM_getQuotaDelta_(goalsData, monthKey) {
  // Quota delta = this month's cumulative quota - last month's cumulative quota
  const thisQuota = goalsData.quotaByMonth.get(monthKey) || 0
  const prevKey = CSM_prevMonthKey_(monthKey)
  const prevQuota = prevKey ? (goalsData.quotaByMonth.get(prevKey) || 0) : 0
  return { target: thisQuota - prevQuota, thisQuota, prevQuota }
}

function CSM_prevMonthKey_(monthKey) {
  const parts = monthKey.split('-')
  if (parts.length < 2) return ''
  let y = Number(parts[0])
  let m = Number(parts[1])
  if (m === 1) { y -= 1; m = 12 } else { m -= 1 }
  return String(y) + '-' + String(m).padStart(2, '0')
}

