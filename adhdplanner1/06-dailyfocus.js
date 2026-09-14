'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Daily Focus'];
const S = "'Daily Focus'";

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,500,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Title row
  vals.push({ range: `${S}!A1`, values: [['DAILY FOCUS — ADHD FOCUS FLOW PLANNER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Date + Day Type row 2
  vals.push({ range: `${S}!A2`, values: [['Date:','','Day Type:','','Today\'s Intention:','','','','','','','']] });
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coralTint), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  // Format col B row 2 as date
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,1,2), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'ddd mm/dd/yyyy' },
    backgroundColor: hex(C.white), textFormat: { bold: true, fontSize: 11 },
  }}, fields: 'userEnteredFormat' }});
  vals.push({ range: `${S}!B2`, values: [['=TODAY()']] });
  vals.push({ range: `${S}!D2`, values: [['Normal Day']] });

  // Today\'s Intention span C2:L2 handled by wide col

  // Spacer row 3
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,12), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // ===== Section 1: Morning Check-in (rows 4-9) =====
  vals.push({ range: `${S}!A4`, values: [['MORNING CHECK-IN']] });
  fmt.push({ mergeCells: { range: gridRange(SID,3,4,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,3,4,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coral), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  const MORNING_LABELS = [
    ['Energy Level (1-5):','','Mood:','','Sleep Quality (1-5):','','Hours Slept:','','','','',''],
    ['Today\'s Capacity:','','Primary Focus:','','','','','','','','',''],
    ['One thing I\'m grateful for:','','','','','','','','','','',''],
    ['One thing that\'s worrying me:','','','','','','','','','','',''],
    ['My intention for today:','','','','','','','','','','',''],
  ];
  vals.push({ range: `${S}!A5`, values: MORNING_LABELS });
  fmt.push({ repeatCell: { range: gridRange(SID,4,9,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coralTint), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 9 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Spacer row 10
  fmt.push({ repeatCell: { range: gridRange(SID,9,10,0,12), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 9, endIndex: 10 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // ===== Section 2: TOP 3 PRIORITIES (rows 11-14) =====
  vals.push({ range: `${S}!A11`, values: [['TOP 3 PRIORITIES FOR TODAY']] });
  fmt.push({ mergeCells: { range: gridRange(SID,10,11,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,10,11,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.blue), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  const P3_HDRS = [['#','Task / Priority','Effort','Time Block','Done?','Notes','','','','','','']];
  vals.push({ range: `${S}!A12`, values: P3_HDRS });
  fmt.push({ repeatCell: { range: gridRange(SID,11,12,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.headerMid), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 11, endIndex: 12 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  const P3_DATA = [
    ['1','Send weekly status update email','Quick','9:00-9:15 AM',false,'','','','','','',''],
    ['2','Finish Q3 report executive summary','Deep Focus','10:00-11:30 AM',false,'No interruptions','','','','','',''],
    ['3','Pay credit card bill','Quick','12:30-12:35 PM',false,'','','','','','',''],
  ];
  vals.push({ range: `${S}!A13`, values: P3_DATA });
  fmt.push({ repeatCell: { range: gridRange(SID,12,15,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.blueTint), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 12, endIndex: 15 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Spacer row 16
  fmt.push({ repeatCell: { range: gridRange(SID,15,16,0,12), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 15, endIndex: 16 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // ===== Section 3: TASK QUEUE (rows 17-36) =====
  vals.push({ range: `${S}!A17`, values: [['TASK QUEUE — TODAY']] });
  fmt.push({ mergeCells: { range: gridRange(SID,16,17,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,16,17,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.teal), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 16, endIndex: 17 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  const TQ_HDRS = [['Task Name','Focus Stage','Priority','Effort','Time Est.','Start Time','Status','Rescue','Done?','Notes','','']];
  vals.push({ range: `${S}!A18`, values: TQ_HDRS });
  fmt.push({ repeatCell: { range: gridRange(SID,17,18,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 17, endIndex: 18 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  const TQ_DATA = [
    ['Respond to urgent Slack messages','Now','Must Do','Quick','15 min','8:30 AM','In Progress','Keep',false,'','',''],
    ['Review calendar for the week','Now','Must Do','Quick','10 min','8:45 AM','In Progress','Keep',false,'','',''],
    ['Send weekly status update email','Now','Must Do','Quick','15 min','9:00 AM','Ready','Keep',false,'','',''],
    ['Fix broken link on website','Now','Important','Quick','15 min','9:15 AM','Ready','Keep',false,'Waiting on IT','',''],
    ['Write Q3 report executive summary','Now','Must Do','Deep Focus','90 min','10:00 AM','Ready','Keep',false,'No interruptions','',''],
    ['Return Amazon package','Next','Important','Quick','20 min','12:00 PM','Ready','Keep',false,'Box in garage','',''],
    ['Pay credit card bill','Now','Must Do','Quick','5 min','12:30 PM','Ready','Keep',false,'','',''],
    ['Brainstorm new project ideas','Now','Important','Standard','30 min','2:00 PM','Ready','Keep',false,'','',''],
    ['Read industry newsletter','Now','Useful','Quick','15 min','3:00 PM','Ready','Keep',false,'','',''],
    ['Review and triage full task list','Now','Must Do','Quick','15 min','4:00 PM','Ready','Keep',false,'Monday habit','',''],
    ['Set weekly priorities','Now','Must Do','Quick','15 min','4:15 PM','Ready','Keep',false,'','',''],
    ['Water the plants','Now','Must Do','Quick','5 min','5:00 PM','Ready','Keep',false,'','',''],
    ['Find recipe for dinner','Now','Important','Quick','10 min','5:30 PM','Ready','Keep',false,'','',''],
    ['Stretch for 10 minutes','Now','Must Do','Quick','10 min','6:30 PM','Ready','Keep',false,'','',''],
    ['Write down top 3 wins','Now','Must Do','Quick','10 min','9:00 PM','Ready','Keep',false,'','',''],
    ['Meditate 5 minutes','Now','Must Do','Quick','5 min','9:30 PM','Ready','Keep',false,'','',''],
    ['Text Jake about moving help','Now','Important','Quick','5 min','7:00 PM','Ready','Keep',false,'He offered to help','',''],
    ['Defrost something for dinner','Now','Important','Quick','2 min','4:45 PM','Ready','Keep',false,'','',''],
    ['Grocery list review','Next','Important','Quick','10 min','6:00 PM','Ready','Keep',false,'','',''],
    ['Review tomorrow\'s calendar','Now','Must Do','Quick','10 min','9:45 PM','Ready','Keep',false,'Night habit','',''],
  ];
  vals.push({ range: `${S}!A19`, values: TQ_DATA });
  for (let r = 18; r < 18+TQ_DATA.length; r++) {
    const bg = r % 2 === 0 ? C.tealTint : C.white;
    fmt.push({ repeatCell: { range: gridRange(SID,r,r+1,0,12), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 18, endIndex: 18+TQ_DATA.length }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // 10 blank task queue rows
  const blankStart = 18 + TQ_DATA.length;
  for (let r = blankStart; r < blankStart+10; r++) {
    const bg = r % 2 === 0 ? C.tealTint : C.white;
    fmt.push({ repeatCell: { range: gridRange(SID,r,r+1,0,12), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: blankStart, endIndex: blankStart+10 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  const spacer1 = blankStart + 10;

  // Spacer
  fmt.push({ repeatCell: { range: gridRange(SID,spacer1,spacer1+1,0,12), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: spacer1, endIndex: spacer1+1 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  const eveningRow = spacer1 + 1;

  // ===== Section 4: EVENING REVIEW =====
  vals.push({ range: `${S}!A${eveningRow+1}`, values: [['EVENING REVIEW']] });
  fmt.push({ mergeCells: { range: gridRange(SID,eveningRow,eveningRow+1,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,eveningRow,eveningRow+1,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.lavender), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: eveningRow, endIndex: eveningRow+1 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  const EVE_LBL = [
    [`Tasks Completed Today:`,`=COUNTIF(I19:I${18+TQ_DATA.length+10},TRUE)`,'','','','','','','','','',''],
    ['Energy Level End of Day (1-5):','','Mood End of Day:','','','','','','','','',''],
    ['3 Wins Today:','','','','','','','','','','',''],
    ['Biggest Challenge:','','','','','','','','','','',''],
    ['What to carry over tomorrow:','','','','','','','','','','',''],
    ['Notes & Brain Dump:','','','','','','','','','','',''],
  ];
  vals.push({ range: `${S}!A${eveningRow+2}`, values: EVE_LBL });
  fmt.push({ repeatCell: { range: gridRange(SID,eveningRow+1,eveningRow+7,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.lavTint), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: eveningRow+1, endIndex: eveningRow+7 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // ===== Summary stats row =====
  const statsRow = eveningRow + 8;
  vals.push({ range: `${S}!A${statsRow+1}`, values: [['TODAY\'S STATS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,statsRow,statsRow+1,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,statsRow,statsRow+1,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.sand), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: statsRow, endIndex: statsRow+1 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  vals.push({ range: `${S}!A${statsRow+2}`, values: [
    ['Tasks Queued','Tasks Done','Completion Rate','Top Priority Done?','Day Rating (1-5)','','','','','','',''],
  ]});
  fmt.push({ repeatCell: { range: gridRange(SID,statsRow+1,statsRow+2,0,5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.sandTint), textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});

  vals.push({ range: `${S}!A${statsRow+3}`, values: [
    [
      `=COUNTA(A19:A${18+TQ_DATA.length+10})`,
      `=COUNTIF(I19:I${18+TQ_DATA.length+10},TRUE)`,
      `=IFERROR(TEXT(COUNTIF(I19:I${18+TQ_DATA.length+10},TRUE)/COUNTA(A19:A${18+TQ_DATA.length+10}),"0%"),"0%")`,
      `=IF(COUNTIF(I13:I15,TRUE)=3,"Yes ✓","Pending")`,
      '',
      '','','','','','','',
    ],
  ]});
  fmt.push({ repeatCell: { range: gridRange(SID,statsRow+2,statsRow+3,0,5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.sandTint), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.sand), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: statsRow+1, endIndex: statsRow+3 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Column widths
  const WIDTHS = [200,100,90,90,90,90,90,100,60,200,80,80];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1-2
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '06-dailyfocus values');
  await batchUpdate(id, fmt, '06-dailyfocus format');
  console.log('✅  Daily Focus done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
