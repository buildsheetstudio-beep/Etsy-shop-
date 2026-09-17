'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const ACT = sheetMap['📋 Action Steps'];

// Layout (GSheets 1-indexed):
// Row 1: title banner (A standalone, B:H merged)
// Row 2: Summary header
// Row 3: Summary stats
// Row 4: Column headers
// Rows 5+: Data (30 pre-filled)

const ACTIONS = [
  // Goal 1: Get Promoted to Senior Manager
  ['Get Promoted to Senior Manager','Enrol in leadership course online',2,'2025-01-12','High','Done',true,''],
  ['Get Promoted to Senior Manager','Complete leadership course modules 1–5',6,'2025-02-09','High','Done',true,''],
  ['Get Promoted to Senior Manager','Complete leadership course modules 6–10',10,'2025-03-09','High','Done',true,''],
  ['Get Promoted to Senior Manager','Volunteer to lead next team project',12,'2025-03-23','High','Done',true,''],
  ['Get Promoted to Senior Manager','Schedule 1:1 with manager to discuss goals',20,'2025-05-18','High','In Progress',false,''],
  ['Get Promoted to Senior Manager','Complete 360° feedback survey',28,'2025-07-13','Medium','Not Started',false,''],
  ['Get Promoted to Senior Manager','Draft promotion case document',36,'2025-09-07','High','Not Started',false,''],
  ['Get Promoted to Senior Manager','Present promotion case to manager',42,'2025-10-19','High','Not Started',false,''],
  // Goal 2: Run a Half Marathon
  ['Run a Half Marathon','Sign up for Couch to 5K app',1,'2025-01-05','High','Done',true,''],
  ['Run a Half Marathon','Complete Week 1–4 of C25K',4,'2025-01-26','High','Done',true,''],
  ['Run a Half Marathon','Complete Week 5–9 (5K run)',9,'2025-03-02','High','Done',true,''],
  ['Run a Half Marathon','Register for half marathon race',10,'2025-03-09','High','Done',true,''],
  ['Run a Half Marathon','Complete 10K training milestone',18,'2025-05-04','High','Done',true,''],
  ['Run a Half Marathon','Begin 12-week half marathon training plan',22,'2025-06-01','High','Done',true,''],
  ['Run a Half Marathon','Complete 18K long run',37,'2025-09-14','High','Not Started',false,''],
  ['Run a Half Marathon','Race day — complete half marathon',41,'2025-10-12','High','Not Started',false,''],
  // Goal 3: Save $20,000 Emergency Fund
  ['Save $20,000 Emergency Fund','Open Ubank high-interest savings account',2,'2025-01-12','High','Done',true,''],
  ['Save $20,000 Emergency Fund','Set up automatic $800/fortnight transfer',2,'2025-01-12','High','Done',true,''],
  ['Save $20,000 Emergency Fund','Review and cut 3 subscriptions',4,'2025-01-26','Medium','Done',true,''],
  ['Save $20,000 Emergency Fund','Reach $10,000 midpoint review',30,'2025-07-27','High','Not Started',false,''],
  ['Save $20,000 Emergency Fund','Final check — confirm $20,000 balance',52,'2025-12-28','High','Not Started',false,''],
  // Goal 4: Learn Spanish to B2
  ['Learn Spanish to B2 Level','Download Duolingo + set 20 min/day habit',1,'2025-01-05','Medium','Done',true,''],
  ['Learn Spanish to B2 Level','Enrol in online Spanish A2 course',6,'2025-02-09','Medium','In Progress',false,''],
  ['Learn Spanish to B2 Level','Start language exchange with native speaker',20,'2025-05-18','Medium','Not Started',false,''],
  ['Learn Spanish to B2 Level','Enrol in B1 level online course',28,'2025-07-13','Medium','Not Started',false,''],
  ['Learn Spanish to B2 Level','Take formal B2 practice test',44,'2025-11-02','High','Not Started',false,''],
  // Goal 5: Travel to Japan
  ['Travel to Japan','Book travel insurance',14,'2025-04-06','High','Done',true,''],
  ['Travel to Japan','Prepare packing list and finalise budget',36,'2025-09-07','Medium','Not Started',false,''],
  // Goal 6: Write and Publish a Blog
  ['Write and Publish a Blog','Register domain name — alexwrites.com',5,'2025-02-02','Medium','Done',true,''],
  ['Write and Publish a Blog','Draft first 3 blog post outlines',16,'2025-04-20','Low','Blocked',false,'Not finding time — need to block calendar'],
];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'📋 Action Steps'";
  const NCOLS = 8; // A-H

  // Column widths
  [200,300,80,110,100,120,70,220].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: ACT, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Row 0: title banner (A standalone, B:H merged)
  reqs.push({ repeatCell: {
    range: gridRange(ACT,0,1,0,1),
    cell: { userEnteredFormat: { backgroundColor: hex(C.midnightBlue) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ mergeCells: { range: gridRange(ACT,0,1,1,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(ACT,0,1,1,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midnightBlue),
      textFormat: { foregroundColor: hex(C.electricGold), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ACT, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!B1`, values: [['  📋  Action Steps — Weekly Task Tracker']] });

  // Row 1: summary section header (0-indexed 1 = GSheets row 2)
  reqs.push({ mergeCells: { range: gridRange(ACT,1,2,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(ACT,1,2,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepNavy),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A2`, values: [['ACTION STEPS SUMMARY']] });

  // Row 2: summary stats (0-indexed 2 = GSheets row 3)
  reqs.push({ repeatCell: {
    range: gridRange(ACT,2,3,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.paleGold),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A3`, values: [[
    'Total Steps:',
    '=COUNTA(B5:B200)',
    'Done:',
    '=COUNTIF(G5:G200,TRUE)',
    '% Complete:',
    '=IFERROR(TEXT(COUNTIF(G5:G200,TRUE)/COUNTA(B5:B200),"0%")&" "&SPARKLINE(COUNTIF(G5:G200,TRUE)/COUNTA(B5:B200),{"charttype","bar";"color1","#2ECC71";"max",1}),"0%")',
    'Overdue:',
    '=COUNTIFS(G5:G200,FALSE,D5:D200,"<"&TODAY())',
  ]]});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ACT, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Row 3: column headers (0-indexed 3 = GSheets row 4)
  const hdrs = ['GOAL','ACTION STEP','WEEK #','DUE DATE','PRIORITY','STATUS','DONE?','NOTES'];
  reqs.push({ repeatCell: {
    range: gridRange(ACT,3,4,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepNavy),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ACT, dimension: 'ROWS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A4`, values: [hdrs] });

  // Data rows (0-indexed 4+ = GSheets row 5+)
  for (let i = 0; i < ACTIONS.length; i++) {
    const ri  = 4 + i;
    const row = 5 + i;
    const bg  = i % 2 === 0 ? C.white : C.paleGold;
    reqs.push({ repeatCell: {
      range: gridRange(ACT, ri, ri+1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
    // Date col D format
    reqs.push({ repeatCell: {
      range: gridRange(ACT, ri, ri+1, 3, 4),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' }, backgroundColor: hex(bg) }},
      fields: 'userEnteredFormat(numberFormat,backgroundColor)',
    }});
    const [goal, step, week, date, pri, status, done, notes] = ACTIONS[i];
    vals.push({ range: `${TAB}!A${row}`, values: [[
      goal, step, week, date, pri, status, done ? 'TRUE' : 'FALSE', notes,
    ]]});
  }

  // Row heights for data rows
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ACT, dimension: 'ROWS', startIndex: 4, endIndex: 4+ACTIONS.length },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Borders
  reqs.push({ updateBorders: {
    range: gridRange(ACT, 3, 4+ACTIONS.length, 0, NCOLS),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  await batchUpdate(id, reqs, 'actions-format');
  await valuesBatchUpdate(id, vals, 'actions-values');
  console.log('Action Steps complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
