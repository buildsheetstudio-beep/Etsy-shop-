'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Master Tasks'];
const S = "'Master Tasks'";

(async () => {
  const fmt = [];
  const vals = [];

  // Expand grid to 5007 rows first
  const { getSheets } = require('./lib');
  const sheets = await getSheets();
  await sheets.spreadsheets.batchUpdate({ spreadsheetId: id, requestBody: { requests: [
    { appendDimension: { sheetId: SID, dimension: 'ROWS', length: 4100 } },
  ]}});

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,5007,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Title row
  vals.push({ range: `${S}!A1`, values: [['MASTER TASKS — ADHD FOCUS FLOW PLANNER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Summary cards row 2 (labels) row 3 (values)
  const LBL = ['Total Tasks','In Progress','Due This Week','Overdue','Completed','Ready/Next','High Priority','','','','','','',''];
  vals.push({ range: `${S}!A2`, values: [LBL] });
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.headerMid), textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});

  const FMLA = [
    `=COUNTA(C8:C5007)`,
    `=COUNTIF(E8:E5007,"In Progress")`,
    `=COUNTIFS(F8:F5007,">="&TODAY(),F8:F5007,"<="&TODAY()+7,E8:E5007,"<>Completed",E8:E5007,"<>Archived")`,
    `=COUNTIFS(F8:F5007,"<"&TODAY(),E8:E5007,"<>Completed",E8:E5007,"<>Archived",C8:C5007,"<>")`,
    `=COUNTIF(E8:E5007,"Completed")`,
    `=COUNTIF(B8:B5007,"Next")+COUNTIF(B8:B5007,"Now")`,
    `=COUNTIF(D8:D5007,"Must Do")`,
    '','','','','','','',
  ];
  vals.push({ range: `${S}!A3`, values: [FMLA] });
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coralTint), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.coral), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Spacer row 4
  fmt.push({ repeatCell: { range: gridRange(SID,3,4,0,14), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // Instructions row 5
  vals.push({ range: `${S}!A5`, values: [['Your centralized task hub. Organize tasks by Focus Stage, Priority, and Effort. Link to Task Breakdown for complex tasks. Mark Done when complete.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,4,5,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,4,5,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.blueTint), textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Spacer row 6
  fmt.push({ repeatCell: { range: gridRange(SID,5,6,0,14), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // Header row 7
  const HDRS = ['Task ID','Focus Stage','Task Name','Priority','Status','Due Date','Effort Level','Time Est.','Task Type','Energy Match','Breakdown?','Rescue Action','Source (Brain Dump)','Notes'];
  vals.push({ range: `${S}!A7`, values: [HDRS] });
  fmt.push({ repeatCell: { range: gridRange(SID,6,7,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // ID formula col A
  fmt.push({ repeatCell: { range: gridRange(SID,7,5007,0,1), cell: {
    userEnteredValue: { formulaValue: `=IF(C8="","","TASK-"&TEXT(ROW()-7,"00000"))` },
    userEnteredFormat: {
      backgroundColor: hex(C.formulaCell), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.medGray) },
      horizontalAlignment: 'CENTER',
    },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Alternating row backgrounds
  for (let r = 7; r < 5007; r++) {
    const bg = r % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,r,r+1,1,14), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }

  // Date format col F
  fmt.push({ repeatCell: { range: gridRange(SID,7,5007,5,6), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'mm/dd/yyyy' },
  }}, fields: 'userEnteredFormat(numberFormat)' }});

  // Column widths
  const WIDTHS = [90,100,240,100,100,100,110,90,100,100,90,110,130,180];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Row heights
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 5007 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // Freeze rows 1-7 only (title row merges across all cols)
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' }});

  // Sample data
  const TODAY = new Date('2026-09-14');
  function d(offset) {
    const dt = new Date(TODAY);
    dt.setDate(dt.getDate() + offset);
    return `${dt.getMonth()+1}/${dt.getDate()}/${dt.getFullYear()}`;
  }
  // [FocusStage, TaskName, Priority, Status, DueDate, EffortLevel, TimeEst, TaskType, EnergyMatch, Breakdown, RescueAction, Source, Notes]
  const DATA = [
    ['Now','Finish Q3 project report','Must Do','In Progress',d(1),'Deep Focus','90+ min','Work','High Effort',true,'Keep','CAP-00002','Ask Sarah for final numbers'],
    ['Now','Pay credit card bill','Must Do','Ready',d(0),'Quick','5 min','Personal','Low Effort',false,'Keep','CAP-00025',''],
    ['Now','Send weekly status update','Must Do','In Progress',d(0),'Quick','10 min','Work','Low Effort',false,'Keep','CAP-00071',''],
    ['Now','File expense reports','Must Do','In Progress',d(-1),'Standard','30 min','Work','Medium Effort',false,'Reschedule','CAP-00036','Overdue'],
    ['Now','Prepare presentation slides','Must Do','In Progress',d(3),'Deep Focus','90+ min','Work','High Effort',true,'Keep','CAP-00022','Team meeting Friday'],
    ['Next','Call dentist','Important','Ready',d(14),'Quick','10 min','Appointment','Low Effort',false,'Keep','CAP-00001','Schedule cleaning'],
    ['Next','Renew car registration','Important','Ready',d(7),'Quick','15 min','Errand','Low Effort',false,'Keep','CAP-00003',''],
    ['Next','Buy birthday gift for Mom','Important','Ready',d(10),'Quick','30 min','Errand','Low Effort',false,'Keep','CAP-00005','Her birthday Sept 25'],
    ['Next','Reply to client proposal email','Important','Completed',d(-12),'Quick','10 min','Work','Low Effort',false,'Keep','CAP-00015','Done'],
    ['Next','Order prescription refills','Important','Completed',d(7),'Quick','5 min','Appointment','Low Effort',false,'Keep','CAP-00041',''],
    ['Next','Update resume','Important','Ready','','Standard','60 min','Work','Medium Effort',false,'Keep','CAP-00018',''],
    ['Next','Schedule annual physical','Important','Ready',d(60),'Quick','10 min','Appointment','Low Effort',false,'Keep','CAP-00013',''],
    ['Next','Cancel unused gym subscription','Important','Completed','','Quick','5 min','Personal','Low Effort',false,'Keep','CAP-00021','Saved $45/month'],
    ['Later','Organize digital photos','Useful','Inbox','','Deep Focus','90+ min','Personal','High Effort',false,'Break Down','CAP-00014',''],
    ['Later','Learn new PM tool','Useful','Inbox','','Deep Focus','90+ min','Work','High Effort',false,'Keep','CAP-00012',''],
    ['Later','Fix leaky faucet','Useful','Inbox','','Standard','60 min','Household','Medium Effort',false,'Keep','CAP-00016',''],
    ['Later','Research home insurance','Useful','Inbox','','Deep Focus','90+ min','Personal','High Effort',false,'Keep','CAP-00044',''],
    ['Later','Reorganize home office','Useful','Inbox','','Deep Focus','90+ min','Household','High Effort',false,'Break Down','CAP-00027',''],
    ['Later','Set up emergency fund transfer','Useful','Inbox','','Standard','30 min','Personal','Medium Effort',false,'Keep','CAP-00034',''],
    ['Later','Plan family vacation for winter','Useful','Inbox',d(90),'Deep Focus','90+ min','Personal','High Effort',true,'Break Down','CAP-00059',''],
    ['Later','Create budget spreadsheet','Optional','Inbox','','Deep Focus','90+ min','Personal','High Effort',false,'Keep','CAP-00057',''],
    ['Later','Learn keyboard shortcuts','Optional','Inbox','','Standard','45 min','Work','Medium Effort',false,'Keep','CAP-00037',''],
    ['Later','Audit monthly subscriptions','Optional','Inbox','','Standard','45 min','Personal','Medium Effort',false,'Keep','CAP-00062',''],
    ['Later','Update LinkedIn profile','Optional','Inbox','','Standard','30 min','Work','Medium Effort',false,'Keep','CAP-00043',''],
    ['Later','Write blog post draft','Optional','Inbox','','Deep Focus','90+ min','Work','High Effort',false,'Keep','CAP-00058','Topic: ADHD at work'],
    ['Event','Team performance review meeting','Must Do','Scheduled',d(14),'Standard','60 min','Event','Medium Effort',false,'Keep','','Prepare talking points'],
    ['Event','Industry conference registration','Important','Scheduled',d(30),'Standard','30 min','Event','Medium Effort',false,'Keep','CAP-00055',''],
    ['Event','Mom birthday dinner','Important','Scheduled',d(11),'Quick','','Event','Low Effort',false,'Keep','CAP-00005','Book restaurant'],
    ['Revisit','Fix broken link on website','Important','Waiting',d(2),'Quick','15 min','Work','Low Effort',false,'Keep','CAP-00033','Waiting on IT'],
    ['Revisit','Get quotes for fence repair','Useful','Waiting','','Standard','30 min','Household','Medium Effort',false,'Reschedule','CAP-00053','Asked 3 contractors'],
    ['Now','Grocery shopping','Must Do','Ready',d(2),'Quick','45 min','Errand','Low Effort',false,'Keep','CAP-00019',''],
    ['Now','Water the plants','Must Do','Ready',d(0),'Quick','5 min','Household','Low Effort',false,'Keep','CAP-00072',''],
    ['Next','Book flights for holiday trip','Important','Ready',d(45),'Standard','45 min','Personal','Medium Effort',false,'Keep','CAP-00009',''],
    ['Next','Register for conference','Important','Ready',d(30),'Standard','30 min','Work','Medium Effort',false,'Keep','CAP-00055','Early bird ends soon'],
    ['Next','Set up medication reminders','Important','Completed','','Quick','10 min','Personal','Low Effort',false,'Keep','CAP-00057','Done'],
    ['Later','Back up computer files','Useful','Inbox','','Standard','30 min','Personal','Medium Effort',false,'Keep','CAP-00039',''],
    ['Later','Research ADHD-friendly routines','Useful','Inbox','','Standard','45 min','Personal','Medium Effort',false,'Keep','CAP-00076',''],
    ['Later','Set up weekly review system','Useful','Inbox','','Standard','45 min','Personal','Medium Effort',false,'Keep','CAP-00077',''],
    ['Later','Find a therapist for ADHD coaching','Useful','Inbox','','Standard','30 min','Personal','Medium Effort',false,'Keep','CAP-00066',''],
    ['Later','Find accountability partner','Optional','Inbox','','Quick','15 min','Personal','Low Effort',false,'Keep','CAP-00081',''],
    ['Now','Practice one deep work session','Must Do','Ready',d(2),'Deep Focus','90+ min','Work','High Effort',false,'Keep','CAP-00082','2 hours, no phone'],
    ['Next','Prepare for performance review','Important','Ready',d(14),'Deep Focus','90+ min','Work','High Effort',false,'Keep','CAP-00073',''],
    ['Next','Plan team bonding activity','Important','Inbox',d(14),'Standard','60 min','Work','Medium Effort',true,'Keep','CAP-00042','Budget $500'],
    ['Later','Write thank you notes','Useful','Inbox',d(-5),'Quick','30 min','Personal','Low Effort',false,'Reschedule','CAP-00040','Overdue'],
    ['Later','Declutter junk drawer','Optional','Inbox','','Quick','30 min','Household','Low Effort',false,'Keep','CAP-00079',''],
    ['Revisit','Call building manager about AC','Important','Waiting',d(-5),'Quick','10 min','Household','Low Effort',false,'Reschedule','CAP-00065','Still not fixed'],
    ['Revisit','Dispute bank charge','Important','Waiting',d(-10),'Standard','30 min','Personal','Low Effort',false,'Keep','CAP-00049','$23 mystery charge'],
    ['Now','Write down top 3 wins this week','Must Do','Ready',d(0),'Quick','10 min','Personal','Low Effort',false,'Keep','CAP-00080',''],
    ['Next','Review credit report','Important','Inbox','','Standard','30 min','Personal','Medium Effort',false,'Keep','CAP-00070',''],
    ['Later','Research tax deduction options','Useful','Inbox','','Deep Focus','90+ min','Personal','High Effort',false,'Keep','CAP-00051',''],
    ['Later','Organize receipts for taxes','Useful','Inbox','','Standard','45 min','Personal','Medium Effort',false,'Keep','CAP-00074',''],
    ['Later','Plan workout schedule','Useful','Inbox','','Standard','30 min','Personal','Medium Effort',false,'Keep','CAP-00050',''],
    ['Later','Research better sleep habits','Optional','Inbox','','Standard','30 min','Personal','Low Effort',false,'Keep','CAP-00035',''],
    ['Later','Create emergency contact list','Optional','Inbox','','Quick','15 min','Personal','Low Effort',false,'Keep','CAP-00063',''],
    ['Later','Unsubscribe from marketing emails','Optional','Inbox','','Quick','20 min','Personal','Low Effort',false,'Keep','CAP-00054',''],
    ['Later','Clean car interior','Optional','Inbox','','Standard','60 min','Household','Medium Effort',false,'Keep','CAP-00046',''],
    ['Later','Set up parental controls','Useful','Inbox','','Standard','30 min','Household','Medium Effort',false,'Keep','CAP-00047',''],
    ['Later','Read "Getting Things Done"','Optional','Inbox','','Deep Focus','90+ min','Personal','High Effort',false,'Keep','CAP-00067',''],
    ['Later','Research standing desk options','Optional','Inbox','','Standard','30 min','Work','Low Effort',false,'Keep','CAP-00026',''],
    ['Later','Buy running shoes','Optional','Inbox','','Quick','20 min','Errand','Low Effort',false,'Keep','CAP-00045',''],
    ['Later','Schedule hair appointment','Optional','Inbox',d(21),'Quick','10 min','Personal','Low Effort',false,'Keep','CAP-00038',''],
    ['Later','Get ergonomic chair','Useful','Inbox','','Standard','45 min','Work','Medium Effort',false,'Keep','CAP-00061',''],
    ['Later','Research meal prep ideas','Useful','Inbox','','Standard','45 min','Personal','Low Effort',false,'Keep','CAP-00004',''],
    ['Later','Set up automatic bill payments','Useful','Inbox','','Standard','30 min','Personal','Low Effort',false,'Keep','CAP-00006',''],
    ['Later','Clean out email inbox','Useful','Inbox','','Standard','45 min','Work','Medium Effort',false,'Keep','CAP-00007','1,200 unread'],
    ['Later','Set up body doubling session','Useful','Inbox',d(7),'Quick','15 min','Personal','Low Effort',false,'Keep','CAP-00068',''],
    ['Event','Dental appointment','Must Do','Scheduled',d(14),'Quick','60 min','Appointment','Low Effort',false,'Keep','CAP-00001',''],
    ['Event','Annual physical','Important','Scheduled',d(60),'Quick','90+ min','Appointment','Low Effort',false,'Keep','CAP-00013',''],
    ['Revisit','Send invoice to freelance client','Must Do','Completed',d(-14),'Quick','10 min','Work','Low Effort',false,'Keep','CAP-00052','Paid! $1200'],
    ['Revisit','Fix squeaky door hinge','Optional','Inbox','','Quick','15 min','Household','Low Effort',false,'Keep','CAP-00008','Need WD-40'],
    ['Now','Return Amazon package','Important','Ready',d(5),'Quick','20 min','Errand','Low Effort',false,'Keep','CAP-00011','Box in garage'],
    ['Next','Write performance self-review','Must Do','In Progress',d(30),'Deep Focus','90+ min','Work','High Effort',false,'Keep','CAP-00010',''],
    ['Later','Research productivity apps','Optional','Inbox','','Standard','30 min','Personal','Low Effort',false,'Keep','CAP-00058',''],
    ['Later','Set up routine for weekday mornings','Useful','Inbox','','Standard','45 min','Personal','Medium Effort',false,'Keep','',''],
    ['Later','Practice saying no to requests','Optional','Inbox','','Deep Focus','','Personal','Medium Effort',false,'Keep','CAP-00064',''],
    ['Later','Look up hyperfocus techniques','Useful','Inbox','','Standard','30 min','Personal','Low Effort',false,'Keep','CAP-00072',''],
    ['Later','Learn to use time blocking','Useful','Inbox','','Standard','45 min','Work','Medium Effort',false,'Keep','',''],
    ['Later','Create vision board for goals','Optional','Inbox','','Standard','60 min','Personal','Medium Effort',false,'Keep','',''],
    ['Now','Review and triage task list','Must Do','Ready',d(0),'Quick','15 min','Personal','Low Effort',false,'Keep','CAP-00028','Weekly habit'],
    ['Now','Set weekly priorities','Must Do','Ready',d(0),'Quick','15 min','Personal','Low Effort',false,'Keep','CAP-00030',''],
    ['Now','Text Jake about moving help','Important','Ready',d(3),'Quick','5 min','Personal','Low Effort',false,'Keep','CAP-00031',''],
    ['Now','Find recipe for dinner tonight','Important','Ready',d(0),'Quick','10 min','Personal','Low Effort',false,'Keep','CAP-00034',''],
    ['Now','Stretch for 10 minutes','Must Do','Ready',d(0),'Quick','10 min','Personal','Low Effort',false,'Keep','CAP-00071',''],
    ['Now','Meditate 5 minutes','Must Do','Ready',d(0),'Quick','5 min','Personal','Low Effort',false,'Keep','CAP-00020',''],
    ['Next','Order new noise-cancelling headphones','Useful','Ready','','Quick','15 min','Work','Low Effort',false,'Keep','CAP-00027',''],
    ['Next','Set up emergency fund auto-transfer','Important','Ready','','Standard','20 min','Personal','Low Effort',false,'Keep','CAP-00034',''],
    ['Later','Plan home office reorganization','Useful','Inbox','','Deep Focus','90+ min','Household','High Effort',true,'Break Down','CAP-00027',''],
    ['Later','Develop better bedtime routine','Useful','Inbox','','Standard','30 min','Personal','Medium Effort',false,'Keep','',''],
    ['Later','Research ADHD coaches online','Useful','Inbox','','Standard','45 min','Personal','Medium Effort',false,'Keep','',''],
    ['Later','Set up Pomodoro timer system','Useful','Inbox','','Quick','20 min','Work','Low Effort',false,'Keep','',''],
    ['Later','Batch cook meals for the week','Useful','Inbox','','Deep Focus','90+ min','Household','High Effort',false,'Keep','',''],
    ['Revisit','Call Mom back','Important','Completed',d(-4),'Quick','15 min','Personal','Low Effort',false,'Keep','CAP-00024','Called!'],
    ['Revisit','Check in with Jake about schedule','Useful','Inbox',d(7),'Quick','5 min','Personal','Low Effort',false,'Keep','',''],
    ['Event','Doctor follow-up appointment','Important','Scheduled',d(21),'Quick','60 min','Appointment','Low Effort',false,'Keep','','Book this week'],
    ['Event','Car registration deadline','Must Do','Scheduled',d(7),'Quick','30 min','Errand','Low Effort',false,'Keep','CAP-00003',''],
    ['Now','Read industry newsletter','Useful','Ready',d(0),'Quick','15 min','Work','Low Effort',false,'Keep','CAP-00048',''],
    ['Now','Brainstorm new project ideas','Important','Ready',d(0),'Standard','30 min','Work','Medium Effort',false,'Keep','CAP-00029',''],
    ['Now','Fix broken link on website','Important','In Progress',d(2),'Quick','15 min','Work','Low Effort',false,'Keep','CAP-00033',''],
    ['Later','Explore new hobby options','Optional','Inbox','','Standard','','Personal','Medium Effort',false,'Keep','','For downtime'],
    ['Later','Create a daily check-in habit','Useful','Inbox','','Quick','10 min','Personal','Low Effort',false,'Keep','',''],
    ['Later','Research noise machines for focus','Optional','Inbox','','Quick','15 min','Work','Low Effort',false,'Keep','',''],
    ['Later','Try walking meetings','Useful','Inbox','','Standard','','Work','Medium Effort',false,'Keep','',''],
    ['Later','Set up digital detox evening','Useful','Inbox','','Standard','','Personal','Low Effort',false,'Keep','',''],
    ['Later','Explore ADHD medication review','Important','Inbox','','Standard','','Appointment','Medium Effort',false,'Keep','','Talk to doctor'],
    ['Later','Learn about habit stacking','Useful','Inbox','','Standard','30 min','Personal','Low Effort',false,'Keep','',''],
    ['Later','Create a "done list" habit','Optional','Inbox','','Quick','10 min','Personal','Low Effort',false,'Keep','',''],
    ['Later','Research mindfulness for ADHD','Useful','Inbox','','Standard','30 min','Personal','Low Effort',false,'Keep','',''],
    ['Later','Try the 2-minute rule for small tasks','Useful','Inbox','','Quick','','Personal','Low Effort',false,'Keep','',''],
    ['Now','Send invoice for September work','Must Do','Ready',d(1),'Quick','15 min','Work','Low Effort',false,'Keep','',''],
    ['Now','Defrost something for dinner','Important','Ready',d(0),'Quick','2 min','Personal','Low Effort',false,'Keep','CAP-00070',''],
    ['Next','Check in with accountability partner','Important','Inbox',d(7),'Quick','15 min','Personal','Low Effort',false,'Keep','',''],
    ['Event','Weekly team standup','Must Do','Scheduled',d(1),'Quick','30 min','Event','Low Effort',false,'Keep','','Every Monday'],
    ['Event','Monthly budget review','Important','Scheduled',d(16),'Standard','60 min','Event','Medium Effort',false,'Keep','','Last Monday of month'],
    ['Now','Respond to urgent Slack messages','Must Do','In Progress',d(0),'Quick','15 min','Work','Low Effort',false,'Keep','',''],
    ['Now','Review calendar for upcoming week','Must Do','Ready',d(0),'Quick','10 min','Personal','Low Effort',false,'Keep','','Sunday habit'],
    ['Later','Plan quarterly goals','Important','Inbox',d(16),'Deep Focus','90+ min','Work','High Effort',false,'Keep','',''],
    ['Later','Research retirement savings options','Useful','Inbox','','Deep Focus','60 min','Personal','High Effort',false,'Keep','',''],
    ['Later','Review and update emergency fund goal','Useful','Inbox','','Standard','30 min','Personal','Medium Effort',false,'Keep','',''],
    ['Later','Create a reading list and schedule','Optional','Inbox','','Quick','20 min','Personal','Low Effort',false,'Keep','',''],
    ['Later','Evaluate workspace ergonomics','Useful','Inbox','','Standard','30 min','Work','Low Effort',false,'Keep','',''],
    ['Later','Try a new focus technique','Useful','Inbox','','Standard','30 min','Personal','Medium Effort',false,'Keep','',''],
    ['Later','Reach out to an old colleague','Optional','Inbox','','Quick','10 min','Work','Low Effort',false,'Keep','',''],
  ];

  const rows = DATA.map(r => {
    // [FocusStage, TaskName, Priority, Status, DueDate, EffortLevel, TimeEst, TaskType, EnergyMatch, Breakdown, RescueAction, Source, Notes]
    return ['', r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[12]];
  });
  vals.push({ range: `${S}!A8`, values: rows });

  await valuesBatchUpdate(id, vals, '04-mastertasks values');
  await batchUpdate(id, fmt, '04-mastertasks format');
  console.log('✅  Master Tasks done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
