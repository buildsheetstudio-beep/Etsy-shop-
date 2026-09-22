'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TT = sheetMap['✅ Task Tracker'];

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: TT, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: TT, dimension: 'ROWS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: TT, dimension: 'ROWS', startIndex: 3, endIndex: 65 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // Column widths: A(#30), B(Task 260), C(Cat 120), D(Event 180), E(Priority 90), F(Status 100), G(Due 110), H(Assigned 140), I(Notes 200), J(Done 60)
  const colW = [30, 260, 120, 180, 90, 100, 110, 140, 200, 60];
  colW.forEach((w,i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: TT, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title row merged
  reqs.push({ mergeCells: { range: gridRange(TT, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(TT, 0, 1, 0, 10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.antiqueGold),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Summary row 2 — labels
  reqs.push({ repeatCell: {
    range: gridRange(TT, 1, 2, 0, 10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.darkPlum), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Summary row 3 — values
  reqs.push({ repeatCell: {
    range: gridRange(TT, 2, 3, 0, 10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.paleRose),
      textFormat: { foregroundColor: hex(C.darkPlum), bold: true, fontSize: 11 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Header row 4
  reqs.push({ repeatCell: {
    range: gridRange(TT, 3, 4, 0, 10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlush),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Data rows alternating
  for (let r = 4; r < 65; r++) {
    const bg = r % 2 === 0 ? C.champagne : C.ivory;
    reqs.push({ repeatCell: {
      range: gridRange(TT, r, r+1, 0, 10),
      cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 10 }, verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // Center-align A, E, F, G, J
  [0, 4, 5, 6, 9].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(TT, 3, 65, c, c+1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat.horizontalAlignment' }});
  });

  // Date format G
  reqs.push({ repeatCell: {
    range: gridRange(TT, 4, 65, 6, 7),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});

  // Done col J — paleRose bg
  reqs.push({ repeatCell: { range: gridRange(TT, 4, 65, 9, 10), cell: { userEnteredFormat: { backgroundColor: hex(C.paleRose) } }, fields: 'userEnteredFormat.backgroundColor' }});

  // Borders
  reqs.push({ updateBorders: {
    range: gridRange(TT, 3, 65, 0, 10),
    innerHorizontal: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
  }});

  await batchUpdate(id, reqs, 'task-format');

  const data = [];
  data.push({ range: `'✅ Task Tracker'!A1`, values: [['✅ Task Tracker — Event Planning To-Dos']] });
  data.push({ range: `'✅ Task Tracker'!A2:J2`, values: [['Total','Done','In Progress','Not Started','Blocked','','','','','% Complete']] });
  data.push({ range: `'✅ Task Tracker'!A3:J3`, values: [[
    '=COUNTA(B5:B65)',
    '=COUNTIF(F5:F65,"Done")',
    '=COUNTIF(F5:F65,"In Progress")',
    '=COUNTIF(F5:F65,"Not Started")',
    '=COUNTIF(F5:F65,"Blocked")',
    '','','','',
    '=IFERROR(B3/A3,0)',
  ]] });
  // Percentage format J3
  reqs.push({ repeatCell: {
    range: gridRange(TT, 2, 3, 9, 10),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  await batchUpdate(id, reqs.slice(-1), 'task-pct');

  data.push({ range: `'✅ Task Tracker'!A4:J4`, values: [['#','Task','Category','Event','Priority','Status','Due Date','Assigned To','Notes','Done']] });

  const tasks = [
    [1,'Book venue and confirm contract','Venue','Tech Summit 2025','High','Done','2025-07-01','Sarah Johnson','Venue deposit paid',true],
    [2,'Send save-the-date emails','Marketing','Tech Summit 2025','High','Done','2025-07-15','Mark Davis','600 recipients',true],
    [3,'Confirm keynote speakers (x3)','Logistics','Tech Summit 2025','High','Done','2025-07-20','Sarah Johnson','All 3 confirmed',true],
    [4,'Set up Eventbrite registration page','Marketing','Tech Summit 2025','High','Done','2025-07-25','Lisa Chen','Link live',true],
    [5,'Arrange AV equipment & tech rehearsal','Audio/Visual','Tech Summit 2025','High','In Progress','2025-08-30','Mark Davis','AV team booked',false],
    [6,'Order conference swag bags (150 units)','Logistics','Tech Summit 2025','Medium','In Progress','2025-08-15','Lisa Chen','Supplier confirmed',false],
    [7,'Finalise catering menu','Catering','Tech Summit 2025','High','In Progress','2025-08-20','Sarah Johnson','Awaiting final headcount',false],
    [8,'Organise volunteer briefing','Guest Management','Tech Summit 2025','Medium','Not Started','2025-09-01','Mark Davis','',false],
    [9,'Print name badges and signage','Logistics','Tech Summit 2025','Medium','Not Started','2025-09-05','Lisa Chen','',false],
    [10,'Conduct post-event survey','Marketing','Tech Summit 2025','Low','Not Started','2025-09-20','Sarah Johnson','',false],
    [11,'Secure venue for Charity Gala','Venue','Charity Gala 2025','High','Done','2025-07-10','Sarah Johnson','Riverside Ballroom confirmed',true],
    [12,'Launch ticket sales','Marketing','Charity Gala 2025','High','Done','2025-07-20','Mark Davis','Sold 80% already',true],
    [13,'Hire auctioneer','Logistics','Charity Gala 2025','High','In Progress','2025-08-10','Sarah Johnson','2 candidates',false],
    [14,'Arrange live entertainment (band)','Entertainment','Charity Gala 2025','High','In Progress','2025-08-01','Mark Davis','Jazz quartet shortlisted',false],
    [15,'Design and print auction catalogue','Decor','Charity Gala 2025','Medium','Not Started','2025-09-01','Lisa Chen','',false],
    [16,'Arrange table centrepieces','Decor','Charity Gala 2025','Medium','Not Started','2025-09-15','Lisa Chen','',false],
    [17,'Confirm charity beneficiary speech','Logistics','Charity Gala 2025','Low','Not Started','2025-10-01','Sarah Johnson','',false],
    [18,'Book florist for gala','Vendor','Charity Gala 2025','Medium','Blocked','2025-08-20','Sarah Johnson','Budget approval needed',false],
    [19,'Create run-of-show document','Logistics','Charity Gala 2025','High','Not Started','2025-10-01','Mark Davis','',false],
    [20,'Finalise guest dietary requirements','Guest Management','Charity Gala 2025','Medium','Not Started','2025-10-10','Lisa Chen','',false],
    [21,'Launch Holiday Party invite','Marketing','Holiday Party 2025','High','Not Started','2025-10-01','Mark Davis','',false],
    [22,'Confirm catering for 80 pax','Catering','Holiday Party 2025','High','Not Started','2025-10-15','Sarah Johnson','',false],
    [23,'Book DJ/entertainment','Vendor','Holiday Party 2025','Medium','Not Started','2025-10-20','Mark Davis','',false],
    [24,'Arrange decor & theming','Decor','Holiday Party 2025','Medium','Not Started','2025-11-01','Lisa Chen','',false],
    [25,'Send RSVPs and collect dietary needs','Guest Management','Holiday Party 2025','Medium','Not Started','2025-11-15','Sarah Johnson','',false],
    [26,'Reserve venue for Investor Day','Venue','Investor Day 2025','High','Done','2025-08-01','Sarah Johnson','HQ boardroom + Zoom',true],
    [27,'Prepare investor presentation deck','Logistics','Investor Day 2025','High','In Progress','2025-09-15','Sarah Johnson','CFO to review',false],
    [28,'Send invites to investor list','Marketing','Investor Day 2025','High','Not Started','2025-09-01','Mark Davis','',false],
    [29,'Arrange catering (boardroom lunch)','Catering','Investor Day 2025','Medium','Not Started','2025-10-10','Lisa Chen','',false],
    [30,'Set up hybrid AV for Investor Day','Audio/Visual','Investor Day 2025','High','Not Started','2025-10-15','Mark Davis','',false],
    [31,'Book venue for Awards Night','Venue','Annual Awards Night','High','Done','2025-08-10','Sarah Johnson','Grand Hotel confirmed',true],
    [32,'Select award categories & winners','Logistics','Annual Awards Night','High','In Progress','2025-09-30','Sarah Johnson','HR to provide nominees',false],
    [33,'Design trophy/awards','Decor','Annual Awards Night','Medium','Not Started','2025-10-01','Lisa Chen','',false],
    [34,'Hire MC for awards night','Vendor','Annual Awards Night','High','Not Started','2025-10-15','Mark Davis','',false],
    [35,'Finalise seating plan for awards','Guest Management','Annual Awards Night','Medium','Not Started','2025-11-15','Sarah Johnson','',false],
  ];

  data.push({ range: `'✅ Task Tracker'!A5:J39`, values: tasks });

  await valuesBatchUpdate(id, data, 'task-values');
  console.log('Task Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
