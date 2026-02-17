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
    'david@pingassistant.com'
  ],

  SUBJECT: 'The Ring Weekly'
}

function send_ring_weekly_email() {
  return send_ring_weekly_email_to_(RING_WEEKLY_CFG.RECIPIENTS)
}

function send_ring_weekly_email_test_docker() {
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

  // Goals
  const goalsMonth = getMonthlyGoalAndQuotaFromGoalsSheet_()
  const monthlyGoalArr = goalsMonth.goalArr
  const monthlyQuotaArr = goalsMonth.quotaArr
  const annualGoalArr = 1000000

  // Percent helpers
  const monthlyPct = clamp01_(monthlyGoalArr > 0 ? (arr / monthlyGoalArr) : 0)
  const annualPct = clamp01_(annualGoalArr > 0 ? (arr / annualGoalArr) : 0)
  const quotaPctOfGoal = clamp01_(monthlyGoalArr > 0 ? (monthlyQuotaArr / monthlyGoalArr) : 0)

  const monthlyGoalPctText = fmtPct_(monthlyPct)
  const annualGoalPctText = fmtPct_(annualPct)

  const annualGoalPctWidth = fmtPctWidth_(annualPct)

  // Format values
  const arrValue = fmtMoney_(arr)
  const monthlyGoalValue = fmtMoney_(monthlyGoalArr)
  const monthlyQuotaValue = fmtMoney_(monthlyQuotaArr)
  const annualGoalValue = fmtMoney_(annualGoalArr)

  const subscriptions = String(subs ?? '')
  const totalSeats = String(seats ?? '')
  const monthlyBarHtml = buildMonthlyBarWithQuotaMarkerHtml_(monthlyPct, quotaPctOfGoal, COLORS)

  return `<!DOCTYPE html>
<html>
  <body style="margin:0;padding:0;background:${COLORS.cream};font-family:Arial,sans-serif;color:${COLORS.ink};">
    <div style="max-width:760px;margin:40px auto;padding:24px;">
      <div style="background:linear-gradient(135deg, rgba(143,136,249,0.18), rgba(251,146,60,0.16));border-radius:28px;padding:18px;">
        <div style="background:#ffffff;border-radius:22px;padding:26px;box-shadow:0 10px 30px rgba(47,43,39,0.08);">

          <!-- Header -->
            <table width="100%" cellpadding="0" cellspacing="0" role="presentation">
              <tr>
                <!-- Left: title -->
                <td align="left" valign="top">
                  <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
                    <span style="width:10px;height:10px;border-radius:50%;background:#FB923C;display:inline-block;"></span>
                    <span style="width:10px;height:10px;border-radius:50%;background:#8F88F9;display:inline-block;"></span>
                    <span style="width:10px;height:10px;border-radius:50%;background:#2F2B27;display:inline-block;"></span>
                    <div style="font-size:12px;letter-spacing:0.22em;opacity:0.65;font-weight:800;">
                      THE RING WEEKLY
                    </div>
                  </div>

                  <div style="font-size:28px;font-weight:900;">Your Biweekly Update</div>
                  <div style="margin-top:6px;font-size:13px;opacity:0.65;">${escapeHtml_(dateStr)}</div>
                </td>

                <!-- Right: button -->
                <td align="right" valign="top">
                  <a href="${escapeHtml_(RING_URL)}" target="_blank" style="text-decoration:none;">
                    <span style="
                      display:inline-block;
                      background:linear-gradient(90deg,#8F88F9,#FB923C);
                      color:#fff;
                      padding:12px 18px;
                      border-radius:14px;
                      font-size:13px;
                      font-weight:900;
                      box-shadow:0 10px 18px rgba(47,43,39,0.14);
                      white-space:nowrap;
                    ">
                      Open The Ring →
                    </span>
                  </a>
                </td>
              </tr>
            </table>

          <div style="height:1px;background:rgba(47,43,39,0.08);margin:18px 0;"></div>

          <!-- ARR + Monthly goal -->
          <div style="background:linear-gradient(135deg, rgba(251,146,60,0.10), rgba(143,136,249,0.10));border-radius:20px;padding:22px;border:1px solid rgba(47,43,39,0.06);margin-bottom:18px;">
            <div style="display:flex;justify-content:space-between;gap:12px;">
              <div>
                <div style="font-size:12px;letter-spacing:0.18em;opacity:0.7;font-weight:900;">ARR</div>
                <div style="font-size:38px;font-weight:950;line-height:1;margin-top:6px;">
                  ${escapeHtml_(arrValue)}
                </div>
              </div>
              <div style="text-align:right;font-size:12px;opacity:0.7;white-space:nowrap;">
                <div style="font-weight:800;">Monthly goal: ${escapeHtml_(monthlyGoalValue)}</div>
                <div style="font-weight:800;">Quota: ${escapeHtml_(monthlyQuotaValue)}</div>
                <div>${escapeHtml_(monthlyGoalPctText)} to goal</div>
              </div>
            </div>

            <div style="margin-top:16px;">
              ${monthlyBarHtml}
            </div>
            <div style="margin-top:6px;font-size:11px;opacity:0.75;">Quota marker shown as vertical line.</div>
          </div>

          <!-- Subs + Seats (email-safe table layout) -->
          <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse:separate;margin-bottom:18px;">
            <tr>
              <td width="50%" valign="top" style="padding-right:10px;">
                <div style="background:#fff;border-radius:20px;padding:26px;border:1px solid rgba(47,43,39,0.08);text-align:center;box-shadow:0 8px 18px rgba(47,43,39,0.06);">
                  <div style="font-size:12px;letter-spacing:0.18em;opacity:0.65;font-weight:950;">SUBSCRIPTIONS</div>
                  <div style="font-size:40px;font-weight:950;color:${COLORS.purple};line-height:1;margin-top:10px;">
                    ${escapeHtml_(subscriptions)}
                  </div>
                  <div style="margin-top:8px;font-size:12px;opacity:0.65;">Active subs</div>
                </div>
              </td>

              <td width="50%" valign="top" style="padding-left:10px;">
                <div style="background:#fff;border-radius:20px;padding:26px;border:1px solid rgba(47,43,39,0.08);text-align:center;box-shadow:0 8px 18px rgba(47,43,39,0.06);">
                  <div style="font-size:12px;letter-spacing:0.18em;opacity:0.65;font-weight:950;">TOTAL SEATS</div>
                  <div style="font-size:38px;font-weight:950;color:${COLORS.orange};line-height:1;margin-top:10px;">
                    ${escapeHtml_(totalSeats)}
                  </div>
                  <div style="margin-top:8px;font-size:12px;opacity:0.65;">Seats in Stripe</div>
                </div>
              </td>
            </tr>
          </table>

          <!-- Annual goal -->
          <div style="background:#fff;border-radius:20px;padding:18px;border:1px solid rgba(47,43,39,0.08);margin-bottom:16px;">
            <div style="display:flex;justify-content:space-between;gap:10px;font-size:12px;opacity:0.7;font-weight:800;flex-wrap:wrap;">
              <div>ANNUAL GOAL BAR</div>
              <div>Goal: ${escapeHtml_(annualGoalValue)} · ${escapeHtml_(annualGoalPctText)} to goal</div>
            </div>
            <div style="margin-top:12px;height:10px;background:rgba(47,43,39,0.10);border-radius:999px;overflow:hidden;">
              <div style="width:${escapeHtml_(annualGoalPctWidth)};height:100%;background:linear-gradient(90deg,${COLORS.purple},${COLORS.orange});"></div>
            </div>
          </div>

          <!-- Notes -->
          <div style="background:${COLORS.cream};border-radius:18px;padding:18px;border:1px solid rgba(47,43,39,0.06);font-size:13px;">
            <strong>Quick notes</strong>
            <ul style="padding-left:18px;margin:8px 0 0;">
              <li>KPIs pulled straight from <strong>The Ring</strong> sheet</li>
              <li>If something looks off, tell Docker PLZ</li>
              <li>Have a magical week, Stinky boys</li>
            </ul>
          </div>

          <div style="margin-top:14px;font-size:11px;opacity:0.55;text-align:center;">
            Sent by Ping Ops · colors #FB923C #8F88F9 #F6F4F0 #2F2B27
          </div>

        </div>
      </div>
    </div>
  </body>
</html>`
}

/* =========================
 * Goals reader:
 * - Goal section: month headers in row 12, ARR values in row 13
 * - Quota section: month headers in row 6, ARR values in row 7
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

  // New layout: Goal section (rows 12/13), Quota section (rows 6/7)
  let goalArr = findMonthValueInRowPair_(sh, 12, 13, thisMonthKey)
  const quotaArr = findMonthValueInRowPair_(sh, 6, 7, thisMonthKey)

  // Legacy fallback (rows 1/2) for older sheets
  if (!(goalArr > 0)) {
    goalArr = findMonthValueInRowPair_(sh, 1, 2, thisMonthKey)
  }

  if (!(goalArr > 0)) {
    // Last-resort fallback: latest positive value from Goal ARR row (row 13)
    const goalRow = sh.getRange(13, 1, 1, lastCol).getValues()[0]
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
  const trackColor = '#D9D7D3'
  const fillColor = colors.purple
  const markerColor = colors.ink

  // Keep this visibly thick in strict email clients.
  const markerWidthPct = 1.8

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
      return `<td width="${escapeHtml_(w)}%" style="padding:0;margin:0;height:10px;width:${escapeHtml_(w)}%;font-size:0;line-height:0;background:${escapeHtml_(c.bg)};">&nbsp;</td>`
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
