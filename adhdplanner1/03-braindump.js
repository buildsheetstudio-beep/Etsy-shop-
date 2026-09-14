'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Brain Dump'];
const S = "'Brain Dump'";

(async () => {
  const fmt = [];
  const vals = [];

  // Expand grid to 2007 rows first
  const { getSheets } = require('./lib');
  const sheets = await getSheets();
  await sheets.spreadsheets.batchUpdate({ spreadsheetId: id, requestBody: { requests: [
    { appendDimension: { sheetId: SID, dimension: 'ROWS', length: 1100 } },
  ]}});

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,2007,0,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Title row
  vals.push({ range: `${S}!A1`, values: [['BRAIN DUMP — ADHD FOCUS FLOW PLANNER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,10), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Motivational message row 2
  vals.push({ range: `${S}!A2`, values: [['CAPTURE FIRST. ORGANIZE LATER.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,10), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coralTint), textFormat: { bold: true, italic: true, fontSize: 10, foregroundColor: hex(C.coral), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // Summary cards row 3 (labels) and row 4 (values)
  const CARD_LABELS = ['Captured Items','Ready to Organize','Moved to Master','Unprocessed','Captured This Week','','','','',''];
  vals.push({ range: `${S}!A3`, values: [CARD_LABELS] });
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.headerMid), textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});

  // Summary formulas row 4
  const CARD_FMLAS = [
    `=COUNTA(C8:C2007)`,
    `=COUNTIF(G8:G2007,TRUE)`,
    `=COUNTIF(H8:H2007,TRUE)`,
    `=COUNTA(C8:C2007)-COUNTIF(G8:G2007,TRUE)`,
    `=COUNTIFS(B8:B2007,">="&TODAY()-7,B8:B2007,"<="&TODAY())`,
    '','','','','',
  ];
  vals.push({ range: `${S}!A4`, values: [CARD_FMLAS] });
  fmt.push({ repeatCell: { range: gridRange(SID,3,4,0,5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coralTint), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.coral), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 4 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Spacer row 5
  fmt.push({ repeatCell: { range: gridRange(SID,4,5,0,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.softGray),
  }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // Instructions row 6
  vals.push({ range: `${S}!A6`, values: [['Use this tab to capture EVERYTHING on your mind — tasks, ideas, worries, reminders. Don\'t judge or sort. Just dump it here and organize later.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,5,6,0,10), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,5,6,0,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.blueTint), textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Header row 7
  const HDRS = ['ID','Captured Date','Captured Item','Task Type','Due Date','Effort','Ready?','Moved?','Linked Task ID','Notes'];
  vals.push({ range: `${S}!A7`, values: [HDRS] });
  fmt.push({ repeatCell: { range: gridRange(SID,6,7,0,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // ID formula column A (rows 8:2007)
  fmt.push({ repeatCell: { range: gridRange(SID,7,2007,0,1), cell: {
    userEnteredValue: { formulaValue: `=IF(C8="","","CAP-"&TEXT(ROW()-7,"00000"))` },
    userEnteredFormat: {
      backgroundColor: hex(C.formulaCell), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.medGray) },
      horizontalAlignment: 'CENTER',
    },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Data rows background (alternating)
  for (let r = 7; r < 2007; r++) {
    const bg = r % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,r,r+1,1,10), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }

  // Date format col B (Captured Date) and col E (Due Date)
  fmt.push({ repeatCell: { range: gridRange(SID,7,2007,1,2), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'mm/dd/yyyy' },
  }}, fields: 'userEnteredFormat(numberFormat)' }});
  fmt.push({ repeatCell: { range: gridRange(SID,7,2007,4,5), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'mm/dd/yyyy' },
  }}, fields: 'userEnteredFormat(numberFormat)' }});

  // Column widths
  const WIDTHS = [90,100,280,110,100,100,70,70,110,200];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Row heights for data rows (compact)
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 2007 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // Freeze rows 1-7 only (no column freeze — title row merges across all cols)
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' }});

  // Sample data (80+ entries)
  const TODAY = new Date('2026-09-14');
  function d(offsetDays) {
    const dt = new Date(TODAY);
    dt.setDate(dt.getDate() + offsetDays);
    return `${dt.getMonth()+1}/${dt.getDate()}/${dt.getFullYear()}`;
  }
  const SAMPLE = [
    [d(-30),'Call dentist to schedule cleaning','Appointment',d(14),'Quick','TRUE','TRUE','TASK-00003','Overdue for 6 months'],
    [d(-28),'Finish Q3 project report','Work',d(-7),'Deep Focus','TRUE','TRUE','TASK-00012','Ask Sarah for final numbers'],
    [d(-26),'Renew car registration','Errand',d(7),'Quick','TRUE','FALSE','','Expires end of month'],
    [d(-25),'Research meal prep ideas for the week','Personal',d(-20),'Standard','TRUE','TRUE','TASK-00007','Pinterest board saved'],
    [d(-24),'Buy birthday gift for Mom','Errand',d(10),'Quick','TRUE','FALSE','','Her birthday Sept 25'],
    [d(-23),'Set up automatic bill payments','Personal','','Standard','TRUE','TRUE','TASK-00019',''],
    [d(-22),'Clean out email inbox','Work','','Quick','TRUE','FALSE','','1,200 unread emails'],
    [d(-21),'Fix squeaky door hinge','Household','','Quick','FALSE','FALSE','','Need WD-40'],
    [d(-20),'Book flights for holiday trip','Personal',d(45),'Standard','TRUE','FALSE','','Compare prices on Tuesday'],
    [d(-19),'Write performance self-review','Work',d(30),'Deep Focus','TRUE','TRUE','TASK-00024',''],
    [d(-18),'Return Amazon package','Errand',d(5),'Quick','TRUE','FALSE','','Box in garage'],
    [d(-17),'Learn new project management tool','Work','','Deep Focus','FALSE','FALSE','','Asana vs Monday'],
    [d(-16),'Schedule annual physical','Appointment',d(60),'Quick','TRUE','FALSE','',''],
    [d(-15),'Organize digital photos','Personal','','Deep Focus','FALSE','FALSE','','5000+ unsorted'],
    [d(-14),'Reply to client proposal email','Work',d(-12),'Quick','TRUE','TRUE','TASK-00031',''],
    [d(-13),'Fix leaky faucet in bathroom','Household','','Standard','FALSE','FALSE','','Need plumber maybe'],
    [d(-12),'Read that book on ADHD strategies','Personal','','Deep Focus','FALSE','FALSE','','On nightstand'],
    [d(-11),'Update resume with new projects','Work','','Standard','TRUE','FALSE','',''],
    [d(-10),'Grocery shopping for the week','Errand',d(2),'Quick','TRUE','FALSE','',''],
    [d(-9),'Meditate daily for 30 days','Personal','','Quick','TRUE','FALSE','','Start with 5 min'],
    [d(-8),'Cancel unused gym subscription','Personal','','Quick','TRUE','TRUE','TASK-00038','$45/month wasted'],
    [d(-7),'Prepare presentation slides','Work',d(3),'Deep Focus','TRUE','TRUE','TASK-00041',''],
    [d(-6),'Research standing desk options','Work','','Standard','FALSE','FALSE','','Back pain getting worse'],
    [d(-5),'Call Mom back','Personal',d(-4),'Quick','TRUE','FALSE','','She called twice'],
    [d(-4),'Pay credit card bill','Personal',d(1),'Quick','TRUE','FALSE','','$847 balance'],
    [d(-3),'Reorganize home office','Household','','Deep Focus','FALSE','FALSE','','Can\'t find anything'],
    [d(-2),'Order new headphones for focus','Work','','Quick','TRUE','FALSE','','Noise-cancelling'],
    [d(-1),'Review and triage task list','Personal',d(0),'Quick','TRUE','FALSE','','Do this every Monday'],
    [d(0),'Brainstorm new project ideas','Work','','Standard','FALSE','FALSE','',''],
    [d(0),'Set weekly priorities','Personal',d(0),'Quick','TRUE','FALSE','',''],
    [d(0),'Text Jake about moving help','Personal',d(3),'Quick','FALSE','FALSE','','He offered last week'],
    [d(0),'Fix broken link on website','Work',d(2),'Quick','TRUE','FALSE','',''],
    [d(0),'Find recipe for dinner tonight','Personal',d(0),'Quick','FALSE','FALSE','',''],
    [d(-1),'Set up emergency fund auto-transfer','Personal','','Standard','FALSE','FALSE','','$200/month goal'],
    [d(-2),'Research better sleep habits','Personal','','Standard','FALSE','FALSE','',''],
    [d(-3),'File expense reports from last trip','Work',d(-1),'Standard','TRUE','TRUE','TASK-00052',''],
    [d(-4),'Learn keyboard shortcuts for VS Code','Work','','Standard','FALSE','FALSE','',''],
    [d(-5),'Schedule hair appointment','Personal',d(21),'Quick','FALSE','FALSE','',''],
    [d(-6),'Back up computer files','Personal','','Standard','TRUE','FALSE','','3 months since last backup'],
    [d(-7),'Write thank you notes','Personal',d(-5),'Quick','TRUE','FALSE','','For birthday gifts'],
    [d(-8),'Order prescription refills','Personal',d(7),'Quick','TRUE','TRUE','TASK-00059',''],
    [d(-9),'Plan team bonding activity','Work',d(14),'Standard','FALSE','FALSE','','Budget $500'],
    [d(-10),'Update LinkedIn profile','Work','','Standard','FALSE','FALSE','',''],
    [d(-11),'Research home insurance options','Personal','','Deep Focus','FALSE','FALSE','','Renewal in 2 months'],
    [d(-12),'Buy new running shoes','Errand','','Quick','FALSE','FALSE','','Old ones are worn out'],
    [d(-13),'Clean car interior','Household','','Standard','FALSE','FALSE','',''],
    [d(-14),'Set up parental controls on tablet','Household','','Standard','FALSE','FALSE','','For kids'],
    [d(-15),'Finish reading industry newsletter','Work','','Quick','FALSE','FALSE','',''],
    [d(-16),'Call to dispute bank charge','Personal',d(-10),'Standard','TRUE','FALSE','','$23 mystery charge'],
    [d(-17),'Plan workout schedule for next month','Personal','','Standard','FALSE','FALSE','',''],
    [d(-18),'Research tax deduction options','Personal','','Deep Focus','FALSE','FALSE','',''],
    [d(-19),'Send invoice to freelance client','Work',d(-14),'Quick','TRUE','TRUE','TASK-00068','$1200 owed'],
    [d(-20),'Get quotes for fence repair','Household','','Standard','FALSE','FALSE','',''],
    [d(-21),'Unsubscribe from marketing emails','Personal','','Quick','FALSE','FALSE','',''],
    [d(-22),'Register for conference next month','Work',d(30),'Standard','TRUE','FALSE','','Early bird ends soon'],
    [d(-23),'Create budget spreadsheet','Personal','','Deep Focus','FALSE','FALSE','',''],
    [d(-24),'Set up reminders for medication','Personal','','Quick','TRUE','TRUE','TASK-00075',''],
    [d(-25),'Research best productivity apps','Personal','','Standard','FALSE','FALSE','',''],
    [d(-26),'Plan family vacation for winter','Personal',d(90),'Deep Focus','FALSE','FALSE','',''],
    [d(-27),'Write blog post draft','Work','','Deep Focus','FALSE','FALSE','','Topic: ADHD at work'],
    [d(-28),'Fix posture — get ergonomic chair','Work','','Standard','FALSE','FALSE','',''],
    [d(-29),'Audit monthly subscriptions','Personal','','Standard','FALSE','FALSE','','Probably $200+ wasted'],
    [d(-30),'Create emergency contact list','Personal','','Quick','FALSE','FALSE','',''],
    [d(-10),'Learn to say no more often','Personal','','Deep Focus','FALSE','FALSE','',''],
    [d(-8),'Call building manager about AC','Household',d(-5),'Quick','TRUE','FALSE','','Still too hot upstairs'],
    [d(-6),'Find a therapist for ADHD coaching','Personal','','Standard','FALSE','FALSE','',''],
    [d(-4),'Read "Getting Things Done"','Personal','','Deep Focus','FALSE','FALSE','','Already have the book'],
    [d(-3),'Set up body doubling session','Personal',d(7),'Quick','FALSE','FALSE','','Tuesday afternoon'],
    [d(-2),'Review credit report','Personal','','Standard','FALSE','FALSE','','Free annual check'],
    [d(-1),'Defrost something for dinner','Personal',d(0),'Quick','FALSE','FALSE','',''],
    [d(0),'Stretch for 10 minutes','Personal',d(0),'Quick','FALSE','FALSE','','Back is stiff'],
    [d(0),'Send weekly status update email','Work',d(0),'Quick','FALSE','FALSE','',''],
    [d(0),'Water the plants','Household',d(0),'Quick','FALSE','FALSE','','They look sad'],
    [d(-1),'Look up hyperfocus techniques','Personal','','Standard','FALSE','FALSE','',''],
    [d(-2),'Prepare for performance review','Work',d(14),'Deep Focus','FALSE','FALSE','',''],
    [d(-3),'Organize receipts for taxes','Personal','','Standard','FALSE','FALSE','',''],
    [d(-4),'Research ADHD-friendly routines','Personal','','Standard','FALSE','FALSE','',''],
    [d(-5),'Set up weekly review system','Personal','','Standard','FALSE','FALSE','',''],
    [d(-6),'Declutter junk drawer','Household','','Quick','FALSE','FALSE','',''],
    [d(-7),'Write down top 3 wins this week','Personal',d(0),'Quick','FALSE','FALSE','',''],
    [d(-8),'Find accountability partner','Personal','','Standard','FALSE','FALSE','',''],
    [d(-9),'Practice one deep work session','Work',d(2),'Deep Focus','FALSE','FALSE','','2 hours, no phone'],
  ];

  const rows = SAMPLE.map(r => {
    // [CapturedDate, CapturedItem, TaskType, DueDate, Effort, Ready, Moved, LinkedID, Notes]
    const ready = r[5] === 'TRUE';
    const moved = r[6] === 'TRUE';
    return ['', r[0], r[1], r[2], r[3], r[4], ready, moved, r[7], r[8]];
  });
  vals.push({ range: `${S}!A8`, values: rows });

  await valuesBatchUpdate(id, vals, '03-braindump values');
  await batchUpdate(id, fmt, '03-braindump format');
  console.log('✅  Brain Dump done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
