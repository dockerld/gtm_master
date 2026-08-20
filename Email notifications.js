/**************************************************************
 * send_ring_weekly_email
 *
 * Sends "The Ring Weekly" every Monday.
 * Safe ASCII-only Apps Script (no emoji parse errors).
 **************************************************************/

const RING_WEEKLY_CFG = {
  SHEET_NAME: 'The Ring',
  ARR_GOAL: 1000000,

  RECIPIENTS: [
    'docker@pingassistant.com',
    'camden@pingassistant.com',
    'chad@pingassistant.com',
    'ben@pingassistant.com',
    'carson@pingassistant.com'
  ],

  SUBJECT: 'The Ring Weekly'
}

function send_ring_weekly_email() {
  return send_ring_weekly_email_to_(RING_WEEKLY_CFG.RECIPIENTS)
}

// Test send to just yourself. This no-arg wrapper shows up in the Apps Script
// editor "Run" dropdown (send_ring_weekly_email_to_ can't, since it takes an arg).
function test_ring_weekly_email_to_me() {
  return send_ring_weekly_email_to_(['docker@pingassistant.com'])
}

function send_ring_weekly_email_to_(recipients) {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName(RING_WEEKLY_CFG.SHEET_NAME)
  if (!sh) throw new Error('Missing "The Ring" sheet')

  const arr = Number(sh.getRange('B2').getValue()) || 0
  const subs = Number(sh.getRange('C2').getValue()) || 0
  const seats = Number(sh.getRange('D2').getValue()) || 0

  const html = buildRingWeeklyHtml_(arr, subs, seats)

  GmailApp.sendEmail(
    recipients.join(','),
    RING_WEEKLY_CFG.SUBJECT,
    'Your email client does not support HTML.',
    { htmlBody: html }
  )

  return { recipients: recipients.length }
}

/* ============================================================
 * HTML BUILDER (Gmail-safe)
 * ============================================================ */

function buildRingWeeklyHtml_(arr, subs, seats) {
  const COLORS = {
    orange: '#FB923C',
    purple: '#8F88F9',
    cream:  '#F6F4F0',
    ink:    '#2F2B27'
  }

  const RING_URL = 'https://docs.google.com/spreadsheets/d/147yUcx8Eb7LE-jhAALwfmddIIOoBYcvEJpXhJR0c8qc/edit?gid=1300412141#gid=1300412141'

  const today = new Date()
  const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEE, MMM d, yyyy')

  // ---- Monthly goal/quota (month-delta model) ----
  // The bar fills from this month's STARTING ARR (baseline) toward the goal,
  // resetting to zero each month. We show "$X to quota" / "$Y to goal" as the
  // remaining gap, and the quota sits as a vertical marker inside the bar.
  const goalsMonth = getMonthlyGoalAndQuotaFromGoalsSheet_()
  const goalTarget = goalsMonth.goalArr     // e.g. 400000 (Goals row 4)
  const quotaTarget = goalsMonth.quotaArr   // e.g. 360000 (Goals row 5)

  const baseline = getMonthStartBaseline_()          // July 2026 -> 300000
  const gained = arr - baseline                      // filled amount this month
  const goalDelta = Math.max(0, goalTarget - baseline)   // full bar length
  const quotaDelta = Math.max(0, quotaTarget - baseline) // quota marker position
  const toQuota = quotaTarget - arr                  // >0 remaining, <=0 hit
  const toGoal = goalTarget - arr

  const fillPct = goalDelta > 0 ? clamp01_(gained / goalDelta) : (gained > 0 ? 1 : 0)
  const quotaMarkerPct = goalDelta > 0 ? clamp01_(quotaDelta / goalDelta) : 0

  const quotaHit = quotaTarget > 0 && arr >= quotaTarget
  const goalHit = goalTarget > 0 && arr >= goalTarget

  const monthLabel = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMMM')

  // Headline chips (\u2705 check mark, \uD83C\uDF89 party popper \u2014 ASCII-safe escapes)
  const quotaBig = quotaHit ? '+' + fmtMoney_(arr - quotaTarget) : fmtMoney_(toQuota)
  const quotaSub = quotaHit ? 'over quota \u2705' : 'to quota'
  const goalBig  = goalHit ? '+' + fmtMoney_(arr - goalTarget) : fmtMoney_(toGoal)
  const goalSub  = goalHit ? 'over goal \uD83C\uDF89' : 'to goal'

  // ---- Annual goal (absolute cumulative) ----
  const annualGoalArr = 1000000
  const annualPct = clamp01_(annualGoalArr > 0 ? (arr / annualGoalArr) : 0)
  const annualGoalPctText = fmtPct_(annualPct)
  const annualGoalPctWidth = fmtPctWidth_(annualPct)

  // ---- Format values ----
  const arrValue = fmtMoney_(arr)
  const gainedValue = (gained >= 0 ? '+' : '-') + fmtMoney_(Math.abs(gained))
  const baselineValue = fmtMoney_(baseline)
  const startLabel = fmtK_(baseline)          // "$300K"
  const goalLabelK = fmtK_(goalTarget)        // "$400K"
  const quotaLabelK = fmtK_(quotaTarget)      // "$360K"
  const annualGoalValueK = fmtK_(annualGoalArr)
  const fillPctText = fmtPct_(fillPct)

  const subscriptions = String(subs ?? '')
  const totalSeats = String(seats ?? '')
  const monthlyBarHtml = buildMonthlyBarWithQuotaMarkerHtml_(fillPct, quotaMarkerPct, COLORS)

  return `<!DOCTYPE html>
<html>
  <body style="margin:0;padding:0;background:${COLORS.cream};font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;color:${COLORS.ink};-webkit-font-smoothing:antialiased;">
    <div style="max-width:680px;margin:0 auto;padding:32px 18px;">
      <div style="background:linear-gradient(135deg, rgba(143,136,249,0.22), rgba(251,146,60,0.18));border-radius:30px;padding:14px;">
        <div style="background:#ffffff;border-radius:24px;padding:32px;box-shadow:0 14px 40px rgba(47,43,39,0.10);">

          <!-- Header -->
          <table width="100%" cellpadding="0" cellspacing="0" role="presentation">
            <tr>
              <td align="left" valign="middle">
                <div style="margin-bottom:10px;line-height:1;">
                  <span style="width:9px;height:9px;border-radius:50%;background:${COLORS.orange};display:inline-block;vertical-align:middle;"></span>
                  <span style="width:9px;height:9px;border-radius:50%;background:${COLORS.purple};display:inline-block;vertical-align:middle;margin-left:4px;"></span>
                  <span style="width:9px;height:9px;border-radius:50%;background:${COLORS.ink};display:inline-block;vertical-align:middle;margin-left:4px;"></span>
                  <span style="font-size:11px;letter-spacing:0.26em;font-weight:800;color:#9B968F;margin-left:10px;vertical-align:middle;">THE RING</span>
                </div>
                <div style="font-size:27px;font-weight:900;letter-spacing:-0.01em;">Weekly Update</div>
                <div style="margin-top:5px;font-size:13px;color:#9B968F;">${escapeHtml_(dateStr)}</div>
              </td>
              <td align="right" valign="middle">
                <a href="${escapeHtml_(RING_URL)}" target="_blank" style="text-decoration:none;">
                  <span style="display:inline-block;background:linear-gradient(90deg,${COLORS.purple},${COLORS.orange});color:#ffffff;padding:11px 18px;border-radius:12px;font-size:13px;font-weight:800;box-shadow:0 8px 16px rgba(143,136,249,0.32);white-space:nowrap;">Open The Ring &rarr;</span>
                </a>
              </td>
            </tr>
          </table>

          <div style="height:1px;background:rgba(47,43,39,0.08);margin:24px 0;"></div>

          <!-- ARR hero + monthly progress -->
          <div style="font-size:11px;letter-spacing:0.2em;font-weight:800;color:#9B968F;">ARR &middot; ${escapeHtml_(monthLabel.toUpperCase())} PROGRESS</div>
          <table width="100%" cellpadding="0" cellspacing="0" role="presentation" style="margin-top:10px;">
            <tr>
              <td align="left" valign="bottom">
                <div style="font-size:44px;font-weight:900;line-height:1;letter-spacing:-0.02em;">${escapeHtml_(arrValue)}</div>
                <div style="margin-top:7px;font-size:12px;color:#9B968F;">Started ${escapeHtml_(monthLabel)} at ${escapeHtml_(baselineValue)}</div>
              </td>
              <td align="right" valign="bottom">
                <span style="display:inline-block;background:rgba(143,136,249,0.14);color:#5A51D6;font-size:13px;font-weight:800;padding:8px 13px;border-radius:999px;white-space:nowrap;">&#9650; ${escapeHtml_(gainedValue)} in ${escapeHtml_(monthLabel)}</span>
              </td>
            </tr>
          </table>

          <!-- Meter -->
          <div style="margin-top:20px;">${monthlyBarHtml}</div>
          <table width="100%" cellpadding="0" cellspacing="0" role="presentation" style="margin-top:8px;">
            <tr>
              <td align="left" style="font-size:11px;color:#9B968F;font-weight:700;">${escapeHtml_(startLabel)} start</td>
              <td align="center" style="font-size:11px;color:${COLORS.ink};font-weight:800;">${escapeHtml_(fillPctText)} of goal</td>
              <td align="right" style="font-size:11px;color:#9B968F;font-weight:700;">Goal ${escapeHtml_(goalLabelK)}</td>
            </tr>
          </table>
          <div style="margin-top:6px;font-size:10.5px;color:#B4AFA8;text-align:center;">Filled bar = ${escapeHtml_(monthLabel)} gain &middot; vertical line = quota (${escapeHtml_(quotaLabelK)})</div>

          <!-- Headline chips: to quota / to goal -->
          <table width="100%" cellpadding="0" cellspacing="0" role="presentation" style="margin-top:18px;">
            <tr>
              <td width="50%" valign="top" style="padding-right:8px;">
                <div style="background:rgba(143,136,249,0.10);border:1px solid rgba(143,136,249,0.28);border-radius:16px;padding:16px 18px;">
                  <div style="font-size:23px;font-weight:900;color:${COLORS.ink};line-height:1;">${escapeHtml_(quotaBig)}</div>
                  <div style="margin-top:7px;font-size:11px;letter-spacing:0.1em;font-weight:800;color:#5A51D6;text-transform:uppercase;">${escapeHtml_(quotaSub)}</div>
                </div>
              </td>
              <td width="50%" valign="top" style="padding-left:8px;">
                <div style="background:rgba(251,146,60,0.10);border:1px solid rgba(251,146,60,0.30);border-radius:16px;padding:16px 18px;">
                  <div style="font-size:23px;font-weight:900;color:${COLORS.ink};line-height:1;">${escapeHtml_(goalBig)}</div>
                  <div style="margin-top:7px;font-size:11px;letter-spacing:0.1em;font-weight:800;color:#C2670F;text-transform:uppercase;">${escapeHtml_(goalSub)}</div>
                </div>
              </td>
            </tr>
          </table>

          <!-- Subs + Seats -->
          <table width="100%" cellspacing="0" cellpadding="0" role="presentation" style="margin-top:14px;">
            <tr>
              <td width="50%" valign="top" style="padding-right:8px;">
                <div style="background:#FBFAF8;border:1px solid rgba(47,43,39,0.08);border-radius:16px;padding:22px;text-align:center;">
                  <div style="font-size:11px;letter-spacing:0.16em;font-weight:800;color:#9B968F;">SUBSCRIPTIONS</div>
                  <div style="font-size:38px;font-weight:900;color:${COLORS.purple};line-height:1;margin-top:9px;">${escapeHtml_(subscriptions)}</div>
                  <div style="margin-top:7px;font-size:12px;color:#9B968F;">Active subs</div>
                </div>
              </td>
              <td width="50%" valign="top" style="padding-left:8px;">
                <div style="background:#FBFAF8;border:1px solid rgba(47,43,39,0.08);border-radius:16px;padding:22px;text-align:center;">
                  <div style="font-size:11px;letter-spacing:0.16em;font-weight:800;color:#9B968F;">TOTAL SEATS</div>
                  <div style="font-size:38px;font-weight:900;color:${COLORS.orange};line-height:1;margin-top:9px;">${escapeHtml_(totalSeats)}</div>
                  <div style="margin-top:7px;font-size:12px;color:#9B968F;">Seats in Stripe</div>
                </div>
              </td>
            </tr>
          </table>

          <!-- Annual goal -->
          <div style="margin-top:14px;background:#FBFAF8;border:1px solid rgba(47,43,39,0.08);border-radius:16px;padding:18px 20px;">
            <table width="100%" cellpadding="0" cellspacing="0" role="presentation">
              <tr>
                <td align="left" style="font-size:11px;letter-spacing:0.16em;font-weight:800;color:#9B968F;">ANNUAL GOAL</td>
                <td align="right" style="font-size:12px;font-weight:800;color:${COLORS.ink};">${escapeHtml_(annualGoalPctText)} &middot; ${escapeHtml_(annualGoalValueK)}</td>
              </tr>
            </table>
            <div style="margin-top:12px;height:12px;background:rgba(47,43,39,0.08);border-radius:999px;overflow:hidden;">
              <div style="width:${escapeHtml_(annualGoalPctWidth)};height:100%;background:linear-gradient(90deg,${COLORS.purple},${COLORS.orange});border-radius:999px;"></div>
            </div>
          </div>

          <!-- Notes -->
          <div style="margin-top:14px;background:${COLORS.cream};border-radius:16px;padding:18px 20px;font-size:13px;color:#5C5851;">
            <div style="font-size:11px;letter-spacing:0.16em;font-weight:800;color:#9B968F;margin-bottom:8px;">QUICK NOTES</div>
            <div style="line-height:1.7;">
              KPIs pulled live from <strong style="color:${COLORS.ink};">The Ring</strong>.<br>
              See something off? Ping Docker.<br>
              Have a magical week, Stinky boys.
            </div>
          </div>

          <div style="margin-top:20px;font-size:11px;color:#B4AFA8;text-align:center;">Sent by Ping Ops &middot; The Ring</div>

        </div>
      </div>
    </div>
  </body>
</html>`
}

/* =========================
 * Goals reader:
 * - Goal section: month headers in row 6, ARR values in row 7
 * - Quota section: month headers in row 12, ARR values in row 13
 * - Legacy fallback: row 1/2
 * ========================= */

function getMonthlyArrGoalFromGoalsSheet_() {
  return getMonthlyGoalAndQuotaFromGoalsSheet_().goalArr
}

function getMonthlyGoalAndQuotaFromGoalsSheet_() {
  const ss = SpreadsheetApp.getActive()
  const sh = ss.getSheetByName('Goals')
  if (!sh) return { goalArr: 0, quotaArr: 0, monthKey: '' }

  const lastCol = sh.getLastColumn()
  if (lastCol < 2) return { goalArr: 0, quotaArr: 0, monthKey: '' }

  const tz = Session.getScriptTimeZone()
  const thisMonthKey = Utilities.formatDate(new Date(), tz, 'MMM-yyyy') // "Dec-2025"

  // Current Goals layout: month headers in row 3, Goal ARR in row 4, Quota in row 5.
  let goalArr = findMonthValueInRowPair_(sh, 3, 4, thisMonthKey)
  let quotaArr = findMonthValueInRowPair_(sh, 3, 5, thisMonthKey)

  // Legacy fallbacks for older sheet layouts (rows 6/7 goal, 12/13 quota, 1/2 goal)
  if (!(goalArr > 0)) goalArr = findMonthValueInRowPair_(sh, 6, 7, thisMonthKey)
  if (!(quotaArr > 0)) quotaArr = findMonthValueInRowPair_(sh, 12, 13, thisMonthKey)
  if (!(goalArr > 0)) goalArr = findMonthValueInRowPair_(sh, 1, 2, thisMonthKey)

  if (!(goalArr > 0)) {
    // Last-resort fallback: latest positive value from the Goal ARR row (row 4)
    const goalRow = sh.getRange(4, 1, 1, lastCol).getValues()[0]
    goalArr = findLatestPositive_(goalRow)
  }

  return {
    goalArr: isFinite(goalArr) ? Number(goalArr) : 0,
    quotaArr: isFinite(quotaArr) ? Number(quotaArr) : 0,
    monthKey: thisMonthKey
  }
}

function findMonthValueInRowPair_(sheet, headerRowNumber, valueRowNumber, monthKey) {
  const lastCol = sheet.getLastColumn()
  if (lastCol < 1) return 0

  const headers = sheet.getRange(headerRowNumber, 1, 1, lastCol).getDisplayValues()[0]
  const values = sheet.getRange(valueRowNumber, 1, 1, lastCol).getValues()[0]

  let idx = -1
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim() === String(monthKey || '').trim()) {
      idx = i
      break
    }
  }

  if (idx < 0) return 0
  const n = Number(values[idx])
  return isFinite(n) ? n : 0
}

function findLatestPositive_(rowValues) {
  for (let i = rowValues.length - 1; i >= 0; i--) {
    const n = Number(rowValues[i])
    if (isFinite(n) && n > 0) return n
  }
  return 0
}

/* =========================
 * Month-start ARR baseline
 *
 * The monthly bar fills from the month's STARTING ARR (baseline) toward the
 * goal, and resets to zero every month. Baseline = ARR at the 1st of the
 * current month = the YYYY-MM-01 row set in arr_snapshot (prior month's close).
 *
 * July 2026 is overridden to 300000: pre-July snapshots weren't trustworthy
 * (the dedup bug), so we seed a flat start. Future months read arr_snapshot
 * automatically. Add an override here ONLY if a month's snapshot is ever bad.
 * ========================= */

const RING_BASELINE_OVERRIDES = {
  '2026-07': 300000
}

function getMonthStartBaseline_() {
  const tz = Session.getScriptTimeZone()
  const now = new Date()
  const ymKey = Utilities.formatDate(now, tz, 'yyyy-MM')   // "2026-07"
  const monthStartKey = ymKey + '-01'                      // "2026-07-01"

  if (Object.prototype.hasOwnProperty.call(RING_BASELINE_OVERRIDES, ymKey)) {
    return Number(RING_BASELINE_OVERRIDES[ymKey]) || 0
  }

  const fromSnap = sumSnapshotArrForDate_(monthStartKey)
  if (fromSnap > 0) return fromSnap

  // Fallback: the most recent snapshot month on/before this month.
  return latestSnapshotArrOnOrBefore_(monthStartKey)
}

function ringSnapshotArrIndexes_(sh) {
  const lastCol = sh.getLastColumn()
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())
  return {
    di: header.findIndex(h => h.toLowerCase() === 'snapshot_date'),
    ai: header.findIndex(h => h.toLowerCase() === 'total_arr'),
    lastCol: lastCol
  }
}

function ringSnapshotDateKey_(v) {
  return (typeof ARR_snap_normSnapshotKey_ === 'function')
    ? ARR_snap_normSnapshotKey_(v)
    : String(v || '').trim()
}

function sumSnapshotArrForDate_(monthStartKey) {
  const sh = SpreadsheetApp.getActive().getSheetByName('arr_snapshot')
  if (!sh) return 0
  const lastRow = sh.getLastRow()
  if (lastRow < 2) return 0
  const { di, ai, lastCol } = ringSnapshotArrIndexes_(sh)
  if (di < 0 || ai < 0) return 0

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues()
  let sum = 0
  for (const r of data) {
    if (ringSnapshotDateKey_(r[di]) !== monthStartKey) continue
    const n = Number(r[ai])
    if (isFinite(n)) sum += n
  }
  return sum
}

function latestSnapshotArrOnOrBefore_(monthStartKey) {
  const sh = SpreadsheetApp.getActive().getSheetByName('arr_snapshot')
  if (!sh) return 0
  const lastRow = sh.getLastRow()
  if (lastRow < 2) return 0
  const { di, ai, lastCol } = ringSnapshotArrIndexes_(sh)
  if (di < 0 || ai < 0) return 0

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues()
  const byDate = {}
  for (const r of data) {
    const key = ringSnapshotDateKey_(r[di])
    if (!key || key > monthStartKey) continue
    const n = Number(r[ai])
    byDate[key] = (byDate[key] || 0) + (isFinite(n) ? n : 0)
  }
  let best = '', sum = 0
  for (const k in byDate) {
    if (!best || k > best) { best = k; sum = byDate[k] }
  }
  return sum
}

/* =========================
 * Formatting helpers
 * ========================= */

function fmtMoney_(n) {
  const x = Number(n)
  if (!isFinite(x)) return '$0.00'
  return '$' + x.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
}

function fmtPct_(pct01) {
  const p = clamp01_(pct01) * 100
  return p.toFixed(1) + '%'
}

// Compact money for scale labels: 300000 -> "$300K", 1000000 -> "$1M".
function fmtK_(n) {
  const x = Number(n) || 0
  if (Math.abs(x) >= 1000000) return '$' + (x / 1000000).toFixed(x % 1000000 === 0 ? 0 : 1) + 'M'
  if (Math.abs(x) >= 1000) return '$' + Math.round(x / 1000) + 'K'
  return '$' + Math.round(x)
}

function fmtPctWidth_(pct01) {
  // cap the bar at 100% width visually
  const p = Math.max(0, Math.min(1, Number(pct01) || 0)) * 100
  return p.toFixed(2) + '%'
}

function clamp01_(n) {
  const x = Number(n)
  if (!isFinite(x)) return 0
  return Math.max(0, Math.min(1, x))
}

function buildMonthlyBarWithQuotaMarkerHtml_(fillPct01, markerPct01, colors) {
  const fillPct = GOOD_safePct_(fillPct01)
  const markerPct = GOOD_safePct_(markerPct01)
  const trackColor = '#ECEAE6'
  const fillColor = colors.purple
  const markerColor = colors.ink

  // Keep this visibly thick in strict email clients.
  const markerWidthPct = 1.6

  const leftOfMarker = Math.max(0, markerPct - markerWidthPct / 2)
  const rightOfMarker = Math.min(100, markerPct + markerWidthPct / 2)

  let cells = []
  if (fillPct <= leftOfMarker) {
    // Fill ends before marker
    cells = [
      { w: fillPct, bg: fillColor },
      { w: leftOfMarker - fillPct, bg: trackColor },
      { w: rightOfMarker - leftOfMarker, bg: markerColor },
      { w: 100 - rightOfMarker, bg: trackColor }
    ]
  } else if (fillPct <= rightOfMarker) {
    // Fill overlaps marker
    cells = [
      { w: leftOfMarker, bg: fillColor },
      { w: rightOfMarker - leftOfMarker, bg: markerColor },
      { w: 100 - rightOfMarker, bg: trackColor }
    ]
  } else {
    // Fill extends past marker
    cells = [
      { w: leftOfMarker, bg: fillColor },
      { w: rightOfMarker - leftOfMarker, bg: markerColor },
      { w: fillPct - rightOfMarker, bg: fillColor },
      { w: 100 - fillPct, bg: trackColor }
    ]
  }

  const cellsHtml = cells
    .filter(c => c.w > 0.01)
    .map(c => {
      const w = c.w.toFixed(3)
      return `<td width="${escapeHtml_(w)}%" style="padding:0;margin:0;height:14px;width:${escapeHtml_(w)}%;font-size:0;line-height:0;background:${escapeHtml_(c.bg)};">&nbsp;</td>`
    })
    .join('')

  return `<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="table-layout:fixed;border-collapse:collapse;background:${trackColor};border-radius:999px;overflow:hidden;"><tr>${cellsHtml}</tr></table>`
}

function GOOD_safePct_(v) {
  const n = Number(v)
  if (!isFinite(n)) return 0
  return Math.max(0, Math.min(100, n * 100))
}

function escapeHtml_(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;')
}
