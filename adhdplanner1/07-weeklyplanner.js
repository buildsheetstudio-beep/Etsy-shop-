'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Weekly Planner'];
const S = "'Weekly Planner'";

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,200,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Title row 1
  vals.push({ range: `${S}!A1`, values: [['WEEKLY PLANNER — ADHD FOCUS FLOW PLANNER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Week navigation row 2: Week of + date input + stats
  vals.push({ range: `${S}!A2`, values: [['Week of (Monday):','','Week Theme / Goal:','','','','','Tasks Planned','Tasks Done','Completion Rate','','','','']] });
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coralTint), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,1,2), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'mm/dd/yyyy' }, backgroundColor: hex(C.white),
    textFormat: { bold: true, fontSize: 11 },
  }}, fields: 'userEnteredFormat' }});
  // Default week start = Monday of current week
  vals.push({ range: `${S}!B2`, values: [['=TODAY()-WEEKDAY(TODAY(),3)']] });

  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // Stats formulas row 2, cols H-J (index 7-9)
  vals.push({ range: `${S}!I2`, values: [['=COUNTA(C8:C200)']] });
  vals.push({ range: `${S}!J2`, values: [['=COUNTIF(H8:H200,TRUE)']] });
  vals.push({ range: `${S}!K2`, values: [['=IFERROR(TEXT(J2/I2,"0%"),"0%")']] });
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,8,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coralTint), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.coral) },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,7,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coralTint), textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.text) },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});

  // Spacer row 3
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,14), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // Weekly priorities section rows 4-6
  vals.push({ range: `${S}!A4`, values: [['WEEKLY PRIORITIES (Top 5)']] });
  fmt.push({ mergeCells: { range: gridRange(SID,3,4,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,3,4,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coral), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  vals.push({ range: `${S}!A5`, values: [['#','Priority / Goal','Category','Due','Done?','','','','','','','','','']] });
  fmt.push({ repeatCell: { range: gridRange(SID,4,5,0,5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.headerMid), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  const WP_DATA = [
    ['1','Complete Q3 project report draft','Work','09/15/2026',false,'','','','','','','','',''],
    ['2','Prepare Friday presentation','Work','09/17/2026',false,'','','','','','','','',''],
    ['3','Renew car registration','Errand','09/21/2026',false,'','','','','','','','',''],
    ['4','Buy birthday gift for Mom','Personal','09/24/2026',false,'','','','','','','','',''],
    ['5','Return Amazon package','Errand','09/19/2026',false,'','','','','','','','',''],
  ];
  vals.push({ range: `${S}!A6`, values: WP_DATA });
  fmt.push({ repeatCell: { range: gridRange(SID,5,10,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coralTint), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,5,10,3,4), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'mm/dd/yyyy' },
  }}, fields: 'userEnteredFormat(numberFormat)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 10 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Spacer row 11
  fmt.push({ repeatCell: { range: gridRange(SID,10,11,0,14), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // ===== Main task grid (rows 12 header, rows 13+ data) =====
  vals.push({ range: `${S}!A12`, values: [['ALL TASKS THIS WEEK']] });
  fmt.push({ mergeCells: { range: gridRange(SID,11,12,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,11,12,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.blue), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 11, endIndex: 12 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Task grid header row 13 (index 12)
  // Note: frozenRowCount can't span merged title that covers all 14 cols — safe here as we're freezing only rows
  const TG_HDRS = ['Task Name','Focus Stage','Priority','Effort','Time','Day','Start Time','Done?','Linked Task ID','Notes','','','',''];
  vals.push({ range: `${S}!A13`, values: [TG_HDRS] });
  fmt.push({ repeatCell: { range: gridRange(SID,12,13,0,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 12, endIndex: 13 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Sample week of tasks (row 14+, index 13+)
  const TODAY = new Date('2026-09-14');
  function dayName(offset) {
    const days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    const dt = new Date(TODAY);
    dt.setDate(dt.getDate() + offset);
    return days[dt.getDay()];
  }
  const TASKS = [
    ['Review calendar & set weekly priorities','Now','Must Do','Quick','15 min','Monday','8:45 AM',false,'','Sunday night or Monday morning'],
    ['Respond to Slack messages','Now','Must Do','Quick','15 min','Monday','8:30 AM',false,'',''],
    ['Send weekly status update email','Now','Must Do','Quick','10 min','Monday','9:00 AM',false,'',''],
    ['Fix broken link on website','Now','Important','Quick','15 min','Monday','9:15 AM',false,'TASK-00102',''],
    ['Write Q3 report executive summary','Now','Must Do','Deep Focus','90 min','Monday','10:00 AM',false,'TASK-00001','No interruptions'],
    ['Return Amazon package','Next','Important','Quick','20 min','Monday','12:00 PM',false,'TASK-00069','Box in garage'],
    ['Pay credit card bill','Now','Must Do','Quick','5 min','Monday','12:30 PM',false,'TASK-00002',''],
    ['Text Jake about moving help','Now','Important','Quick','5 min','Monday','7:00 PM',false,'',''],
    ['Meditate 5 minutes','Now','Must Do','Quick','5 min','Monday','9:00 PM',false,'','Daily habit'],
    ['Build presentation slides structure','Now','Must Do','Standard','45 min','Tuesday','9:00 AM',false,'TASK-00005',''],
    ['Write Q3 report body — sections 1-3','Now','Must Do','Deep Focus','90 min','Tuesday','10:00 AM',false,'TASK-00001','Deep focus block'],
    ['Renew car registration','Next','Must Do','Quick','20 min','Tuesday','1:00 PM',false,'TASK-00006',''],
    ['Brainstorm new project ideas','Now','Important','Standard','30 min','Tuesday','2:00 PM',false,'',''],
    ['Order prescription refills','Next','Important','Quick','5 min','Tuesday','3:00 PM',false,'TASK-00041',''],
    ['Grocery shopping','Now','Must Do','Quick','45 min','Tuesday','5:00 PM',false,'TASK-00019',''],
    ['Finish presentation slides','Now','Must Do','Deep Focus','90 min','Wednesday','9:00 AM',false,'TASK-00005',''],
    ['Practice presentation run-through','Now','Must Do','Standard','30 min','Wednesday','11:00 AM',false,'TASK-00005',''],
    ['Write Q3 report body — sections 4-5','Now','Must Do','Deep Focus','60 min','Wednesday','2:00 PM',false,'TASK-00001',''],
    ['Schedule annual physical','Next','Important','Quick','10 min','Wednesday','3:00 PM',false,'TASK-00013',''],
    ['Book flights for holiday trip','Next','Important','Standard','45 min','Wednesday','4:00 PM',false,'TASK-00032',''],
    ['Q3 report final proofreading','Now','Must Do','Standard','30 min','Thursday','9:00 AM',false,'TASK-00001',''],
    ['Send Q3 report to manager','Now','Must Do','Quick','5 min','Thursday','10:00 AM',false,'TASK-00001',''],
    ['Buy birthday gift for Mom','Next','Important','Quick','30 min','Thursday','1:00 PM',false,'TASK-00007',''],
    ['Book restaurant for Mom\'s birthday','Next','Important','Quick','15 min','Thursday','2:00 PM',false,'','Sep 25'],
    ['Team performance review prep','Now','Must Do','Standard','45 min','Thursday','3:00 PM',false,'TASK-00010',''],
    ['Friday presentation delivery','Now','Must Do','Standard','60 min','Friday','10:00 AM',false,'TASK-00005','Team meeting'],
    ['Weekly review and reflection','Now','Must Do','Standard','30 min','Friday','3:30 PM',false,'','Wins/misses'],
    ['Set priorities for next week','Now','Must Do','Quick','15 min','Friday','4:00 PM',false,'',''],
    ['Plan weekend priorities','Now','Important','Quick','10 min','Friday','5:00 PM',false,'',''],
    ['Batch cook meals for next week','Later','Useful','Deep Focus','90 min','Saturday','10:00 AM',false,'','Optional'],
    ['Call Mom','Now','Important','Quick','30 min','Saturday','2:00 PM',false,'','She called twice'],
    ['Explore new hobby options','Later','Optional','Standard','','Saturday','','false','',''],
    ['Review next week calendar','Now','Must Do','Quick','10 min','Sunday','7:00 PM',false,'',''],
    ['Set next week\'s Top 3 Priorities','Now','Must Do','Quick','15 min','Sunday','7:30 PM',false,'',''],
    ['Brain dump anything new on mind','Now','Must Do','Quick','10 min','Sunday','8:00 PM',false,'',''],
  ];

  vals.push({ range: `${S}!A14`, values: TASKS });
  for (let r = 13; r < 13+TASKS.length; r++) {
    const bg = r % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,r,r+1,0,14), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 13, endIndex: 13+TASKS.length+10 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // Blank rows after sample data
  const blankStart = 13 + TASKS.length;
  for (let r = blankStart; r < blankStart+10; r++) {
    const bg = r % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,r,r+1,0,14), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
    }}, fields: 'userEnteredFormat' }});
  }

  // Spacer
  const spacerRow = blankStart + 10;
  fmt.push({ repeatCell: { range: gridRange(SID,spacerRow,spacerRow+1,0,14), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: spacerRow, endIndex: spacerRow+1 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // Weekly reflection section
  const refRow = spacerRow + 1;
  vals.push({ range: `${S}!A${refRow+1}`, values: [['WEEKLY REFLECTION']] });
  fmt.push({ mergeCells: { range: gridRange(SID,refRow,refRow+1,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,refRow,refRow+1,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.lavender), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: refRow, endIndex: refRow+1 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  const REF_LBL = [
    ['This week\'s biggest win:','','','','','','','','','','','','',''],
    ['What didn\'t get done (and why):','','','','','','','','','','','','',''],
    ['One thing to do differently next week:','','','','','','','','','','','','',''],
    ['Carry-overs to next week:','','','','','','','','','','','','',''],
    ['Overall week rating (1-5):','','Energy trend this week:','','','','','','','','','','',''],
  ];
  vals.push({ range: `${S}!A${refRow+2}`, values: REF_LBL });
  fmt.push({ repeatCell: { range: gridRange(SID,refRow+1,refRow+6,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.lavTint), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: refRow+1, endIndex: refRow+6 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Column widths
  const WIDTHS = [220,100,90,90,80,90,90,60,110,200,80,80,80,80];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1-3 only (title row merges across all cols)
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '07-weeklyplanner values');
  await batchUpdate(id, fmt, '07-weeklyplanner format');
  console.log('✅  Weekly Planner done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
