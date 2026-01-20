/**************************************************************
 * render_arr_snapshot_audit()
 *
 * Audits arr_snapshot data types and aggregation consistency.
 * Outputs summary + detail into "arr_snapshot_audit".
 **************************************************************/

const ARR_SNAP_AUDIT_CFG = {
  SNAP_SHEET: 'arr_snapshot',
  OUT_SHEET: 'arr_snapshot_audit',

  HEADER_ROW: 1,
  DATA_START_ROW: 2,

  SNAPSHOT_DATE_HEADER: 'snapshot_date',
  COHORT_HEADER: 'trial_cohort_month',
  ARR_HEADER: 'total_arr'
}

function render_arr_snapshot_audit() {
  return lockWrapCompat_('render_arr_snapshot_audit', () => {
    const ss = SpreadsheetApp.getActive()
    const snap = ss.getSheetByName(ARR_SNAP_AUDIT_CFG.SNAP_SHEET)
    if (!snap) throw new Error(`Missing sheet: ${ARR_SNAP_AUDIT_CFG.SNAP_SHEET}`)

    const lastRow = snap.getLastRow()
    const lastCol = snap.getLastColumn()
    if (lastRow < ARR_SNAP_AUDIT_CFG.DATA_START_ROW) {
      throw new Error('arr_snapshot has no data rows to audit.')
    }

    const header = snap
      .getRange(ARR_SNAP_AUDIT_CFG.HEADER_ROW, 1, 1, lastCol)
      .getValues()[0]
      .map(h => String(h || '').trim())

    const snapIdx = header.findIndex(h => h.toLowerCase() === ARR_SNAP_AUDIT_CFG.SNAPSHOT_DATE_HEADER)
    const cohortIdx = header.findIndex(h => h.toLowerCase() === ARR_SNAP_AUDIT_CFG.COHORT_HEADER)
    const arrIdx = header.findIndex(h => h.toLowerCase() === ARR_SNAP_AUDIT_CFG.ARR_HEADER)

    if (snapIdx < 0) throw new Error(`arr_snapshot missing header: ${ARR_SNAP_AUDIT_CFG.SNAPSHOT_DATE_HEADER}`)
    if (cohortIdx < 0) throw new Error(`arr_snapshot missing header: ${ARR_SNAP_AUDIT_CFG.COHORT_HEADER}`)
    if (arrIdx < 0) throw new Error(`arr_snapshot missing header: ${ARR_SNAP_AUDIT_CFG.ARR_HEADER}`)

    const data = snap
      .getRange(ARR_SNAP_AUDIT_CFG.DATA_START_ROW, 1, lastRow - 1, lastCol)
      .getValues()

    const summaryMap = new Map()
    const detailMap = new Map()

    data.forEach(r => {
      const snapVal = r[snapIdx]
      const snapKey = ARR_audit_snapshotKey_(snapVal) || '(blank)'
      const snapSummary = ARR_audit_getSummary_(summaryMap, snapKey)

      snapSummary.rows += 1
      ARR_audit_incSnapType_(snapSummary, snapVal)

      const trialInfo = ARR_audit_trialInfo_(r[cohortIdx])
      ARR_audit_incTrialType_(snapSummary, trialInfo.type)

      const arrInfo = ARR_audit_arrInfo_(r[arrIdx])
      ARR_audit_incArrType_(snapSummary, arrInfo)
      snapSummary.total_sum += arrInfo.value

      if (!detailMap.has(snapKey)) detailMap.set(snapKey, new Map())
      const snapDetail = detailMap.get(snapKey)
      if (!snapDetail.has(trialInfo.key)) {
        snapDetail.set(trialInfo.key, {
          rows: 0,
          total_sum: 0,
          arr_number: 0,
          arr_text_number: 0,
          arr_blank: 0,
          arr_text_other: 0,
          trial_date: 0,
          trial_text_date: 0,
          trial_text: 0,
          trial_blank: 0
        })
      }
      const d = snapDetail.get(trialInfo.key)
      d.rows += 1
      d.total_sum += arrInfo.value
      ARR_audit_incArrType_(d, arrInfo)
      ARR_audit_incTrialType_(d, trialInfo.type)
    })

    const summaryRows = []
    Array.from(summaryMap.keys()).sort().forEach(key => {
      const s = summaryMap.get(key)
      summaryRows.push([
        key,
        s.rows,
        s.total_sum,
        s.arr_number,
        s.arr_text_number,
        s.arr_blank,
        s.arr_text_other,
        s.trial_date,
        s.trial_text_date,
        s.trial_text,
        s.trial_blank,
        s.snap_date,
        s.snap_text,
        s.snap_blank
      ])
    })

    const detailRows = []
    Array.from(detailMap.keys()).sort().forEach(snapKey => {
      const trialMap = detailMap.get(snapKey)
      Array.from(trialMap.keys()).sort().forEach(trialKey => {
        const d = trialMap.get(trialKey)
        detailRows.push([
          snapKey,
          trialKey,
          d.rows,
          d.total_sum,
          d.arr_number,
          d.arr_text_number,
          d.arr_blank,
          d.arr_text_other,
          d.trial_date,
          d.trial_text_date,
          d.trial_text,
          d.trial_blank
        ])
      })
    })

    const out = getOrCreateSheetCompat_(ss, ARR_SNAP_AUDIT_CFG.OUT_SHEET)
    out.clear()

    const summaryHeader = [
      'snapshot_date',
      'rows',
      'total_arr_sum',
      'arr_number_count',
      'arr_text_number_count',
      'arr_blank_count',
      'arr_text_other_count',
      'trial_date_count',
      'trial_text_date_count',
      'trial_text_count',
      'trial_blank_count',
      'snapshot_date_count',
      'snapshot_text_count',
      'snapshot_blank_count'
    ]
    out.getRange(1, 1, 1, summaryHeader.length).setValues([summaryHeader])
    if (summaryRows.length) {
      out.getRange(2, 1, summaryRows.length, summaryHeader.length).setValues(summaryRows)
    }

    const detailStartRow = summaryRows.length + 3
    const detailHeader = [
      'snapshot_date',
      'trial_cohort_key',
      'rows',
      'total_arr_sum',
      'arr_number_count',
      'arr_text_number_count',
      'arr_blank_count',
      'arr_text_other_count',
      'trial_date_count',
      'trial_text_date_count',
      'trial_text_count',
      'trial_blank_count'
    ]
    out.getRange(detailStartRow, 1, 1, detailHeader.length).setValues([detailHeader])
    if (detailRows.length) {
      out.getRange(detailStartRow + 1, 1, detailRows.length, detailHeader.length).setValues(detailRows)
    }

    out.setFrozenRows(1)
    out.autoResizeColumns(1, Math.max(summaryHeader.length, detailHeader.length))
  })
}

function ARR_audit_getSummary_(map, key) {
  if (!map.has(key)) {
    map.set(key, {
      rows: 0,
      total_sum: 0,
      arr_number: 0,
      arr_text_number: 0,
      arr_blank: 0,
      arr_text_other: 0,
      trial_date: 0,
      trial_text_date: 0,
      trial_text: 0,
      trial_blank: 0,
      snap_date: 0,
      snap_text: 0,
      snap_blank: 0
    })
  }
  return map.get(key)
}

function ARR_audit_snapshotKey_(v) {
  if (!v) return ''
  if (v instanceof Date && !isNaN(v.getTime())) {
    return ARR_audit_dateToKey_(v)
  }
  return String(v || '').trim()
}

function ARR_audit_dateToKey_(d) {
  const y = d.getUTCFullYear()
  const m = String(d.getUTCMonth() + 1).padStart(2, '0')
  const day = String(d.getUTCDate()).padStart(2, '0')
  return `${y}-${m}-${day}`
}

function ARR_audit_incSnapType_(summary, v) {
  if (!v) {
    summary.snap_blank += 1
    return
  }
  if (v instanceof Date && !isNaN(v.getTime())) {
    summary.snap_date += 1
    return
  }
  summary.snap_text += 1
}

function ARR_audit_trialInfo_(v) {
  if (!v) return { key: '(blank)', type: 'blank' }

  if (v instanceof Date && !isNaN(v.getTime())) {
    const y = v.getUTCFullYear()
    const m = String(v.getUTCMonth() + 1).padStart(2, '0')
    return { key: `${y}-${m}`, type: 'date' }
  }

  if (typeof v === 'number' && isFinite(v)) {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000))
    if (!isNaN(d.getTime())) {
      const y = d.getUTCFullYear()
      const m = String(d.getUTCMonth() + 1).padStart(2, '0')
      return { key: `${y}-${m}`, type: 'text_date' }
    }
  }

  const s = String(v || '').trim()
  if (!s) return { key: '(blank)', type: 'blank' }

  const d = new Date(s)
  if (!isNaN(d.getTime())) {
    const y = d.getUTCFullYear()
    const m = String(d.getUTCMonth() + 1).padStart(2, '0')
    return { key: `${y}-${m}`, type: 'text_date' }
  }

  return { key: s, type: 'text' }
}

function ARR_audit_incTrialType_(summary, type) {
  if (type === 'date') summary.trial_date += 1
  else if (type === 'text_date') summary.trial_text_date += 1
  else if (type === 'text') summary.trial_text += 1
  else summary.trial_blank += 1
}

function ARR_audit_arrInfo_(v) {
  if (v === '' || v === null || v === undefined) {
    return { type: 'blank', value: 0 }
  }
  if (typeof v === 'number' && isFinite(v)) {
    return { type: 'number', value: v }
  }
  const s = String(v || '').trim()
  if (!s) return { type: 'blank', value: 0 }
  if (/^-?\d+(\.\d+)?$/.test(s)) {
    return { type: 'text_number', value: Number(s) }
  }
  return { type: 'text_other', value: 0 }
}

function ARR_audit_incArrType_(summary, info) {
  if (info.type === 'number') summary.arr_number += 1
  else if (info.type === 'text_number') summary.arr_text_number += 1
  else if (info.type === 'text_other') summary.arr_text_other += 1
  else summary.arr_blank += 1
}
