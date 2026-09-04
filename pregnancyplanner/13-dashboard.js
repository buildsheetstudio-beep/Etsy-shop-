'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DASH = sheetMap['🌸 Dashboard'];

// 14 cols, 55 rows — built last so all cross-tab formulas resolve

(async () => {
  const reqs = [];

  // Row 0: Main title
  reqs.push({ mergeCells: { range: gridRange(DASH, 0, 1, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 0, 1, 0, 14),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DASH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });

  // Row 1: Disclaimer
  reqs.push({ mergeCells: { range: gridRange(DASH, 1, 2, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 1, 2, 0, 14),
    cell: { userEnteredFormat: { backgroundColor: hex(C.warmCream), textFormat: { foregroundColor: hex(C.warmGrey), italic: true, fontSize: 8 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment,wrapStrategy)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DASH, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Row 2: Due date input label area (A-E) + spacer
  reqs.push({ mergeCells: { range: gridRange(DASH, 2, 3, 0, 2), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(DASH, 2, 3, 2, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(DASH, 2, 3, 5, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 2, 3, 0, 2),
    cell: { userEnteredFormat: { backgroundColor: hex(C.taupeGold), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'RIGHT' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 2, 3, 2, 5),
    cell: { userEnteredFormat: { backgroundColor: hex(C.ivory), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 13 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER', numberFormat: { type: 'DATE', pattern: 'DD MMMM YYYY' } } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment,numberFormat)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DASH, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 38 }, fields: 'pixelSize' } });

  // Row 3: Spacer (cream)
  reqs.push({ mergeCells: { range: gridRange(DASH, 3, 4, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DASH, 3, 4, 0, 14), cell: { userEnteredFormat: { backgroundColor: hex(C.warmCream) } }, fields: 'userEnteredFormat.backgroundColor' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DASH, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 12 }, fields: 'pixelSize' } });

  // ── KPI Section rows 4-6 ──────────────────────────────────────────────────
  // 7 KPI tiles across 14 cols (2 cols each)
  const kpiLabels = [
    'Current Week','Days to Due Date','Trimester',
    'Next Appt','Nursery Ready','Budget Spent','Meals Frozen',
  ];
  kpiLabels.forEach((lbl, i) => {
    const c0 = i * 2, c1 = c0 + 2;
    reqs.push({ mergeCells: { range: gridRange(DASH, 4, 5, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(DASH, 4, 5, c0, c1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.softLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    reqs.push({ mergeCells: { range: gridRange(DASH, 5, 7, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(DASH, 5, 7, c0, c1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.whisperLav), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 18 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DASH, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DASH, dimension: 'ROWS', startIndex: 5, endIndex: 7 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // Row 7: spacer
  reqs.push({ mergeCells: { range: gridRange(DASH, 7, 8, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DASH, 7, 8, 0, 14), cell: { userEnteredFormat: { backgroundColor: hex(C.warmCream) } }, fields: 'userEnteredFormat.backgroundColor' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DASH, dimension: 'ROWS', startIndex: 7, endIndex: 8 }, properties: { pixelSize: 12 }, fields: 'pixelSize' } });

  // ── Progress section row 8-12 ─────────────────────────────────────────────
  // Left: Pregnancy progress (cols 0-6), Right: Budget & Bag progress (cols 7-13)
  reqs.push({ mergeCells: { range: gridRange(DASH, 8, 9, 0, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(DASH, 8, 9, 7, 14), mergeType: 'MERGE_ALL' } });
  [0,7].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(DASH, 8, 9, c, c+7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.taupeGold), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'LEFT' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }}));
  reqs.push({ updateDimensionProperties: { range: { sheetId: DASH, dimension: 'ROWS', startIndex: 8, endIndex: 9 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Progress labels + bar rows 9-12
  const progLabels = ['Pregnancy Progress','Hospital Bag Packed','Nursery Complete','Budget Used'];
  progLabels.forEach((lbl, i) => {
    const r = 9 + i;
    const c0 = i < 2 ? 0 : 7;
    const c1 = i < 2 ? 7 : 14;
    const rr = i < 2 ? i : i;
    reqs.push({ mergeCells: { range: gridRange(DASH, r, r+1, c0, c0+3), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DASH, r, r+1, c0+3, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(DASH, r, r+1, c0, c0+3),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderMist), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'LEFT' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(DASH, r, r+1, c0+3, c1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.ivory), textFormat: { foregroundColor: hex(C.deepPlum), fontSize: 9 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
    }});
    reqs.push({ updateDimensionProperties: { range: { sheetId: DASH, dimension: 'ROWS', startIndex: r, endIndex: r+1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  });

  // Row 13: spacer
  reqs.push({ mergeCells: { range: gridRange(DASH, 13, 14, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DASH, 13, 14, 0, 14), cell: { userEnteredFormat: { backgroundColor: hex(C.warmCream) } }, fields: 'userEnteredFormat.backgroundColor' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DASH, dimension: 'ROWS', startIndex: 13, endIndex: 14 }, properties: { pixelSize: 12 }, fields: 'pixelSize' } });

  // ── Quick stats section row 14-25 ────────────────────────────────────────
  // Left col (0-6): Pregnancy stats, Right col (7-13): Baby stats
  reqs.push({ mergeCells: { range: gridRange(DASH, 14, 15, 0, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(DASH, 14, 15, 7, 14), mergeType: 'MERGE_ALL' } });
  [0,7].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(DASH, 14, 15, c, c+7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.taupeGold), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'LEFT' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }}));
  reqs.push({ updateDimensionProperties: { range: { sheetId: DASH, dimension: 'ROWS', startIndex: 14, endIndex: 15 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Stats rows 15-25 (label + value pairs)
  for (let r = 15; r < 26; r++) {
    reqs.push({ mergeCells: { range: gridRange(DASH, r, r+1, 0, 3), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DASH, r, r+1, 3, 7), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DASH, r, r+1, 7, 10), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DASH, r, r+1, 10, 14), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(DASH, r, r+1, 0, 3),
      cell: { userEnteredFormat: { backgroundColor: hex(r%2===1 ? C.lavenderMist : C.paleLavender), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'RIGHT' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(DASH, r, r+1, 3, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(r%2===1 ? C.ivory : C.whisperLav), textFormat: { foregroundColor: hex(C.deepLavender), bold: true, fontSize: 9 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'LEFT' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(DASH, r, r+1, 7, 10),
      cell: { userEnteredFormat: { backgroundColor: hex(r%2===1 ? C.lavenderMist : C.paleLavender), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'RIGHT' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(DASH, r, r+1, 10, 14),
      cell: { userEnteredFormat: { backgroundColor: hex(r%2===1 ? C.ivory : C.whisperLav), textFormat: { foregroundColor: hex(C.deepLavender), bold: true, fontSize: 9 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'LEFT' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
    }});
    reqs.push({ updateDimensionProperties: { range: { sheetId: DASH, dimension: 'ROWS', startIndex: r, endIndex: r+1 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  }

  // ── Chart placeholder row 26+ ─────────────────────────────────────────────
  // Chart data helper cols (hidden area) — rows 27-54 for chart source
  reqs.push({ mergeCells: { range: gridRange(DASH, 26, 27, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 26, 27, 0, 14),
    cell: { userEnteredFormat: { backgroundColor: hex(C.taupeGold), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'LEFT' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DASH, dimension: 'ROWS', startIndex: 26, endIndex: 27 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Chart data rows 27-54: symptom frequency + mood by week
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 27, 54, 0, 6),
    cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { foregroundColor: hex(C.deepPlum), fontSize: 8 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Column widths
  const colW = Array(14).fill(100);
  colW.forEach((w, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: w }, fields: 'pixelSize',
  }}));

  const border = { style: 'SOLID', color: hex(C.lavenderBorder) };
  reqs.push({ updateBorders: { range: gridRange(DASH, 4, 27, 0, 14), innerHorizontal: border, innerVertical: border, top: border, bottom: border, left: border, right: border }});

  await batchUpdate(id, reqs, 'dash-format');

  // ── Values ────────────────────────────────────────────────────────────────
  const vdata = [];
  vdata.push({ range: `'🌸 Dashboard'!A1`, values: [['🌸 All-in-One Pregnancy & Postpartum Planner']] });
  vdata.push({ range: `'🌸 Dashboard'!A2`, values: [['⚠️ For personal tracking only. Always consult your healthcare provider. This planner does not constitute medical advice.']] });

  // Due date input row 2
  vdata.push({ range: `'🌸 Dashboard'!A3`, values: [['Due Date:']] });
  vdata.push({ range: `'🌸 Dashboard'!C3`, values: [['2025-12-15']] });

  // KPI labels row 4 (every 2 cols)
  vdata.push({ range: `'🌸 Dashboard'!A5:N5`, values: [[
    'Current Week','','Days to Due Date','','Trimester','',
    'Next Appointment','','Nursery % Ready','','Budget % Spent','','Meals Frozen','',
  ]]});

  // KPI values row 5-6 (merged cells, just write to first col)
  vdata.push({ range: `'🌸 Dashboard'!A6`, values: [[ '=IFERROR(IF(TODAY()<=$B$3,40-INT(($B$3-TODAY())/7),40+INT((TODAY()-$B$3)/7)),"Enter due date")' ]] });
  vdata.push({ range: `'🌸 Dashboard'!C6`, values: [[ '=IFERROR(IF($B$3>=TODAY(),$B$3-TODAY(),"Baby is here! 🎉"),"—")' ]] });
  vdata.push({ range: `'🌸 Dashboard'!E6`, values: [[ '=IFERROR(IF(A6<=13,"1st ✨",IF(A6<=27,"2nd 🌺","3rd 🌸")),"—")' ]] });
  vdata.push({ range: `'🌸 Dashboard'!G6`, values: [[ '=IFERROR(TEXT(MINIFS(\'🏥 Appointments Log\'!$A$6:$A$210,\'🏥 Appointments Log\'!$H$6:$H$210,FALSE,\'🏥 Appointments Log\'!$A$6:$A$210,">"&TODAY()),"DD MMM YYYY"),"None scheduled")' ]] });
  vdata.push({ range: `'🌸 Dashboard'!I6`, values: [[ '=IFERROR(TEXT(COUNTIF(\'🛏️ Nursery Planner\'!$G$6:$G$35,TRUE)/COUNTA(\'🛏️ Nursery Planner\'!$B$6:$B$35),"0%"),"—")' ]] });
  vdata.push({ range: `'🌸 Dashboard'!K6`, values: [[ '=IFERROR(TEXT(\'💰 Baby Budget\'!$C$3/\'💰 Baby Budget\'!$A$3,"0%"),"—")' ]] });
  vdata.push({ range: `'🌸 Dashboard'!M6`, values: [[ '=IFERROR(COUNTIF(\'🍱 Freezer Meal Planner\'!$G$5:$G$34,"Frozen"),"—")' ]] });

  // Progress section
  vdata.push({ range: `'🌸 Dashboard'!A9`, values: [['  🤰 Pregnancy Progress']] });
  vdata.push({ range: `'🌸 Dashboard'!H9`, values: [['  📦 Bag & Nursery Progress']] });

  vdata.push({ range: `'🌸 Dashboard'!A10`, values: [['Weeks Complete:']] });
  vdata.push({ range: `'🌸 Dashboard'!D10`, values: [[ `=IFERROR(SPARKLINE({A6,40-A6},{"charttype","bar";"color1","#6B5B8B";"color2","#DDD5E8"}),"")` ]] });
  vdata.push({ range: `'🌸 Dashboard'!A11`, values: [['Baby Size:']] });
  vdata.push({ range: `'🌸 Dashboard'!D11`, values: [[ '=IFERROR(VLOOKUP(A6,\'📋 Reference Data\'!$A$2:$B$40,2,FALSE),"—")' ]] });
  vdata.push({ range: `'🌸 Dashboard'!H10`, values: [['Bag Packed:']] });
  vdata.push({ range: `'🌸 Dashboard'!K10`, values: [[ `=IFERROR(SPARKLINE({\'🎒 Hospital Bag Checklist\'!B3,\'🎒 Hospital Bag Checklist\'!A3-\'🎒 Hospital Bag Checklist\'!B3},{"charttype","bar";"color1","#6B5B8B";"color2","#DDD5E8"}),"")` ]] });
  vdata.push({ range: `'🌸 Dashboard'!H11`, values: [['Nursery Done:']] });
  vdata.push({ range: `'🌸 Dashboard'!K11`, values: [[ `=IFERROR(SPARKLINE({COUNTIF(\'🛏️ Nursery Planner\'!G6:G35,TRUE),COUNTA(\'🛏️ Nursery Planner\'!B6:B35)-COUNTIF(\'🛏️ Nursery Planner\'!G6:G35,TRUE)},{"charttype","bar";"color1","#B8A06A";"color2","#DDD5E8"}),"")` ]] });
  vdata.push({ range: `'🌸 Dashboard'!H12`, values: [['Meals Frozen:']] });
  vdata.push({ range: `'🌸 Dashboard'!K12`, values: [[ `=IFERROR(SPARKLINE({COUNTIF(\'🍱 Freezer Meal Planner\'!G5:G34,"Frozen"),\'🍱 Freezer Meal Planner\'!E3-COUNTIF(\'🍱 Freezer Meal Planner\'!G5:G34,"Frozen")},{"charttype","bar";"color1","#5A7A6A";"color2","#DDD5E8"}),"")` ]] });

  // Stats section headers
  vdata.push({ range: `'🌸 Dashboard'!A15`, values: [['  📊 Pregnancy Stats']] });
  vdata.push({ range: `'🌸 Dashboard'!H15`, values: [['  👶 Baby & Planning Stats']] });

  // Left stats (rows 15-25)
  const leftStats = [
    ['Total Symptoms Logged:', '=IFERROR(\'🤒 Symptom & Wellbeing Tracker\'!A4,"—")'],
    ['Most Common Symptom:', '=IFERROR(\'🤒 Symptom & Wellbeing Tracker\'!C4,"—")'],
    ['Appointments Completed:', '=IFERROR(\'🏥 Appointments Log\'!B4,"—")'],
    ['Appointments Upcoming:', '=IFERROR(\'🏥 Appointments Log\'!C4,"—")'],
    ['Avg Mood Score:', '=IFERROR(\'🤒 Symptom & Wellbeing Tracker\'!D4,"—")'],
    ['Current Milestone:', '=IFERROR(VLOOKUP(A6,\'📋 Reference Data\'!$A$2:$C$40,3,FALSE),"—")'],
    ['Weekly Journal Entries:', '=IFERROR(COUNTA(\'📅 Weekly Pregnancy Journal\'!D4:D42),"—")'],
    ['Symptoms This Week:', '=IFERROR(COUNTIF(\'🤒 Symptom & Wellbeing Tracker\'!B6:B210,A6),"—")'],
    ['Severe Symptoms:', '=IFERROR(\'🤒 Symptom & Wellbeing Tracker\'!B4,"—")'],
    ['Next Appt Provider:', '=IFERROR(INDEX(\'🏥 Appointments Log\'!$B$6:$B$210,MATCH(MINIFS(\'🏥 Appointments Log\'!$A$6:$A$210,\'🏥 Appointments Log\'!$H$6:$H$210,FALSE,\'🏥 Appointments Log\'!$A$6:$A$210,">"&TODAY()),\'🏥 Appointments Log\'!$A$6:$A$210,0)),"—")'],
    ['Birth Plan Status:', '=IFERROR(INDEX(\'📝 Notes & Contacts\'!$D$22:$D$33,MATCH("Birth Plan",\'📝 Notes & Contacts\'!$C$22:$C$33,0)),"Draft")'],
  ];
  leftStats.forEach(([lbl, val], i) => {
    const r = 16 + i;
    vdata.push({ range: `'🌸 Dashboard'!A${r+1}`, values: [[lbl]] });
    vdata.push({ range: `'🌸 Dashboard'!D${r+1}`, values: [[val]] });
  });

  // Right stats
  const rightStats = [
    ['Total Baby Names:', '=IFERROR(\'👶 Baby Names\'!A4,"—")'],
    ['Top Name Score:', '=IFERROR(\'👶 Baby Names\'!D4,"—")'],
    ['Favourite Name:', '=IFERROR(\'👶 Baby Names\'!E4,"—")'],
    ['Nursery Items Purchased:', '=IFERROR(\'🛏️ Nursery Planner\'!B3,"—")'],
    ['Nursery Budget Remaining:', '=IFERROR("$"&TEXT(\'🛏️ Nursery Planner\'!F3,"#,##0.00"),"—")'],
    ['Hospital Bag % Packed:', '=IFERROR(\'🎒 Hospital Bag Checklist\'!D3,"—")'],
    ['Total Budget Allocated:', '=IFERROR("$"&TEXT(\'💰 Baby Budget\'!B3,"#,##0.00"),"—")'],
    ['Budget Remaining:', '=IFERROR("$"&TEXT(\'💰 Baby Budget\'!D3,"#,##0.00"),"—")'],
    ['Meals Target:', '=IFERROR(\'🍱 Freezer Meal Planner\'!E3,"—")'],
    ['Freezer Meals Ready:', '=IFERROR(\'🍱 Freezer Meal Planner\'!B3,"—")'],
    ['Key Contacts Saved:', '=IFERROR(COUNTA(\'📝 Notes & Contacts\'!B4:B13),"—")'],
  ];
  rightStats.forEach(([lbl, val], i) => {
    const r = 16 + i;
    vdata.push({ range: `'🌸 Dashboard'!H${r+1}`, values: [[lbl]] });
    vdata.push({ range: `'🌸 Dashboard'!K${r+1}`, values: [[val]] });
  });

  // Chart data area header
  vdata.push({ range: `'🌸 Dashboard'!A27`, values: [['  📈 Chart Data (Auto-Calculated)']] });

  // Chart data: mood by week (for mood line chart)
  vdata.push({ range: `'🌸 Dashboard'!A28:F28`, values: [['Week','Mood Score','Symptom Count','','','']] });
  const chartRows = [];
  for (let w = 4; w <= 42; w++) {
    chartRows.push([
      w,
      `=IFERROR(AVERAGEIF('🤒 Symptom & Wellbeing Tracker'!$B$6:$B$210,${w},'🤒 Symptom & Wellbeing Tracker'!$I$6:$I$210),"")`,
      `=IFERROR(COUNTIF('🤒 Symptom & Wellbeing Tracker'!$B$6:$B$210,${w}),"")`,
      '','','',
    ]);
  }
  vdata.push({ range: `'🌸 Dashboard'!A29:F${28+chartRows.length}`, values: chartRows });

  await valuesBatchUpdate(id, vdata, 'dash-values');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
