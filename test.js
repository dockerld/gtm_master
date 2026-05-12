/**
 * Count email assistant usage from Stripe subscription items.
 * Writes results to an "Email Assistant Usage" sheet.
 */
function count_email_assistant_subscriptions() {
  const EMAIL_ASSISTANT_PRODUCT_IDS = new Set([
    'prod_U7KjtgNHqnC1rj',
    'prod_U1vVdVWVVBj2rE',
    'prod_U1vVGsH6hsmdIV',
    'prod_U1vV3wJbRmELcD',
    'prod_U1vVOf9XxMDvB8',
    'prod_TzyxUIZ1alizJQ',
    'prod_TzyxkKaI5itFPe'
  ])

  const ss = SpreadsheetApp.getActive()

  // Read subscription items
  const shItems = ss.getSheetByName('raw_stripe_subscription_items')
  if (!shItems) throw new Error('Missing sheet: raw_stripe_subscription_items')
  const itemData = shItems.getDataRange().getValues()
  const itemHeaders = itemData[0].map(h => String(h).trim().toLowerCase())
  const iSubId = itemHeaders.indexOf('stripe_subscription_id')
  const iProductId = itemHeaders.indexOf('product_id')
  const iProductName = itemHeaders.indexOf('product_name')
  const iQty = itemHeaders.indexOf('quantity')

  // Read subscriptions for status + customer info
  const shSubs = ss.getSheetByName('raw_stripe_subscriptions')
  if (!shSubs) throw new Error('Missing sheet: raw_stripe_subscriptions')
  const subData = shSubs.getDataRange().getValues()
  const subHeaders = subData[0].map(h => String(h).trim().toLowerCase())
  const sSubId = subHeaders.indexOf('stripe_subscription_id')
  const sStatus = subHeaders.indexOf('status')
  const sEmail = subHeaders.indexOf('customer_email')
  const sCreated = subHeaders.indexOf('created_at')

  // Build sub info map
  const subInfoBySubId = new Map()
  for (let i = 1; i < subData.length; i++) {
    const sid = String(subData[i][sSubId] || '').trim()
    if (!sid) continue
    subInfoBySubId.set(sid, {
      status: String(subData[i][sStatus] || '').trim().toLowerCase(),
      email: sEmail >= 0 ? String(subData[i][sEmail] || '').trim() : '',
      created: sCreated >= 0 ? subData[i][sCreated] : ''
    })
  }

  // Read canon_orgs for org name lookup by email
  const shCanon = ss.getSheetByName('canon_orgs')
  const orgNameByEmail = new Map()
  const orgNameBySubId = new Map()
  if (shCanon) {
    const canonData = shCanon.getDataRange().getValues()
    const canonHeaders = canonData[0].map(h => String(h).trim().toLowerCase())
    const cOrgName = canonHeaders.indexOf('org_name')
    const cOwnerEmail = canonHeaders.indexOf('owner_email')
    const cBillingEmail = canonHeaders.indexOf('billing_email')
    const cStripeSubs = canonHeaders.indexOf('stripe_subscription_ids')
    for (let i = 1; i < canonData.length; i++) {
      const name = cOrgName >= 0 ? String(canonData[i][cOrgName] || '').trim() : ''
      if (!name) continue
      if (cOwnerEmail >= 0) {
        const email = String(canonData[i][cOwnerEmail] || '').trim().toLowerCase()
        if (email) orgNameByEmail.set(email, name)
      }
      if (cBillingEmail >= 0) {
        const email = String(canonData[i][cBillingEmail] || '').trim().toLowerCase()
        if (email) orgNameByEmail.set(email, name)
      }
      if (cStripeSubs >= 0) {
        const subs = String(canonData[i][cStripeSubs] || '').split(',').map(s => s.trim()).filter(Boolean)
        for (const s of subs) orgNameBySubId.set(s, name)
      }
    }
  }

  // Collect detail rows and summary counts
  const detailRows = []
  const summaryByProduct = {}
  let totalActive = 0, totalTrialing = 0, totalOther = 0

  for (let i = 1; i < itemData.length; i++) {
    const productId = String(itemData[i][iProductId] || '').trim()
    if (!EMAIL_ASSISTANT_PRODUCT_IDS.has(productId)) continue

    const subId = String(itemData[i][iSubId] || '').trim()
    const qty = Math.max(0, Number(itemData[i][iQty]) || 0)
    const productName = String(itemData[i][iProductName] || '').trim()
    const subInfo = subInfoBySubId.get(subId) || {}
    const status = subInfo.status || ''
    const email = subInfo.email || ''
    const created = subInfo.created || ''
    const orgName = orgNameBySubId.get(subId) || orgNameByEmail.get(email.toLowerCase()) || ''

    if (!summaryByProduct[productName]) summaryByProduct[productName] = { active: 0, trialing: 0, other: 0 }

    if (status === 'active') {
      totalActive += qty
      summaryByProduct[productName].active += qty
    } else if (status === 'trialing') {
      totalTrialing += qty
      summaryByProduct[productName].trialing += qty
    } else {
      totalOther += qty
      summaryByProduct[productName].other += qty
    }

    detailRows.push([orgName, email, productName, productId, subId, status, qty, created])
  }

  // Sort: active first, then trialing, then other; within each group by org name
  const statusOrder = { active: 0, trialing: 1 }
  detailRows.sort((a, b) => {
    const sa = statusOrder[a[5]] !== undefined ? statusOrder[a[5]] : 2
    const sb = statusOrder[b[5]] !== undefined ? statusOrder[b[5]] : 2
    if (sa !== sb) return sa - sb
    return String(a[0]).localeCompare(String(b[0]))
  })

  // Write to sheet
  const sheetName = 'Email Assistant Usage'
  let sh = ss.getSheetByName(sheetName)
  if (sh) sh.clear()
  else sh = ss.insertSheet(sheetName)

  let row = 1

  // Summary section
  sh.getRange(row, 1).setValue('Email Assistant Usage').setFontWeight('bold').setFontSize(14)
  row += 1
  sh.getRange(row, 1).setValue('Generated: ' + new Date().toLocaleString())
  row += 2

  // Summary table
  const summaryHeaders = ['Metric', 'Active', 'Trialing', 'Other', 'Total']
  sh.getRange(row, 1, 1, summaryHeaders.length).setValues([summaryHeaders]).setFontWeight('bold').setBackground('#F3F4F6')
  row += 1
  sh.getRange(row, 1, 1, 5).setValues([['Total Seats', totalActive, totalTrialing, totalOther, totalActive + totalTrialing + totalOther]])
  row += 1

  for (const [name, counts] of Object.entries(summaryByProduct).sort((a, b) => b[1].active - a[1].active)) {
    sh.getRange(row, 1, 1, 5).setValues([[name, counts.active, counts.trialing, counts.other, counts.active + counts.trialing + counts.other]])
    row += 1
  }

  row += 1

  // Detail table
  const detailHeaders = ['Org Name', 'Email', 'Product', 'Product ID', 'Subscription ID', 'Status', 'Seats', 'Created']
  sh.getRange(row, 1, 1, detailHeaders.length).setValues([detailHeaders]).setFontWeight('bold').setBackground('#F3F4F6')
  row += 1

  if (detailRows.length) {
    sh.getRange(row, 1, detailRows.length, detailHeaders.length).setValues(detailRows)
    row += detailRows.length
  }

  // Auto-resize columns
  for (let c = 1; c <= detailHeaders.length; c++) sh.autoResizeColumn(c)

  Logger.log(`Email Assistant Usage: ${detailRows.length} rows written to "${sheetName}" sheet`)
}
