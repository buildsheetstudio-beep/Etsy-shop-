'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GL = sheetMap['🎯 Goals & Milestones'];

// Tab layout (0-indexed rows):
// Row 0: title banner (A0 standalone for col-A freeze, B0:K0 merged)
// Row 1: column headers
// Row 2: Goal 1 header  (GSheets row 3)
// Rows 3-7: Goal 1 milestones (GSheets rows 4-8)
// Row 8: Goal 2 header (GSheets row 9)
// ... every 6 rows
// Row 56: Goal 10 header (GSheets row 57)
// Rows 57-61: Goal 10 milestones

const NCOLS = 11; // A-K (indices 0-10)
const GOAL_ROWS_0IDX = [2, 8, 14, 20, 26, 32, 38, 44, 50, 56]; // 0-indexed goal header rows

const COL_HEADERS = ['GOAL / MILESTONE','LIFE AREA','WHY (MOTIVATION)','TARGET DATE','DAYS LEFT','STATUS','PRIORITY','% COMPLETE','PROGRESS BAR','SUCCESS METRIC','NOTES'];

const GOALS = [
  {
    title: 'Get Promoted to Senior Manager',
    area: 'Career & Work',
    why: 'Increase earning potential and lead a team',
    target: '2025-12-31',
    status: 'In Progress',
    priority: 'High',
    pct: 50,
    metric: "Achieve 'Exceeds Expectations' in performance review; receive formal promotion",
    notes: '',
    milestones: [
      { name: 'Complete leadership course',        date: '2025-03-31', status: 'Done',        done: true,  notes: '' },
      { name: 'Lead 2 major projects',             date: '2025-06-30', status: 'Done',        done: true,  notes: '' },
      { name: 'Request performance review',         date: '2025-09-30', status: 'Not Started', done: false, notes: '' },
      { name: 'Present promotion case to manager',  date: '2025-10-31', status: 'Not Started', done: false, notes: '' },
      { name: 'Receive promotion confirmation',     date: '2025-12-31', status: 'Not Started', done: false, notes: '' },
    ],
  },
  {
    title: 'Run a Half Marathon',
    area: 'Health & Fitness',
    why: 'Prove to myself I can do it — and improve my health',
    target: '2025-10-12',
    status: 'In Progress',
    priority: 'High',
    pct: 65,
    metric: 'Cross the finish line of an official half marathon (21.1km)',
    notes: '',
    milestones: [
      { name: 'Complete Couch to 5K program',      date: '2025-03-01', status: 'Done',        done: true,  notes: '' },
      { name: 'Run 10K without stopping',           date: '2025-05-01', status: 'Done',        done: true,  notes: '' },
      { name: 'Run 15K training run',               date: '2025-07-15', status: 'Done',        done: true,  notes: '' },
      { name: 'Complete final long run (18K)',        date: '2025-09-20', status: 'Not Started', done: false, notes: '' },
      { name: 'Complete half marathon race',         date: '2025-10-12', status: 'Not Started', done: false, notes: '' },
    ],
  },
  {
    title: 'Save $20,000 Emergency Fund',
    area: 'Finance & Wealth',
    why: 'Sleep better at night knowing I have a safety net',
    target: '2025-12-31',
    status: 'On Track',
    priority: 'High',
    pct: 40,
    metric: 'Balance of $20,000 in dedicated savings account by Dec 31',
    notes: '',
    milestones: [
      { name: 'Open dedicated savings account',   date: '2025-01-15', status: 'Done',        done: true,  notes: '' },
      { name: 'Save first $5,000',                date: '2025-04-30', status: 'Done',        done: true,  notes: '' },
      { name: 'Save $10,000 (halfway)',            date: '2025-07-31', status: 'Not Started', done: false, notes: '' },
      { name: 'Save $15,000',                      date: '2025-10-31', status: 'Not Started', done: false, notes: '' },
      { name: 'Reach $20,000 goal',                date: '2025-12-31', status: 'Not Started', done: false, notes: '' },
    ],
  },
  {
    title: 'Learn Spanish to B2 Level',
    area: 'Education & Skills',
    why: 'Travel to South America and connect with family heritage',
    target: '2025-11-30',
    status: 'In Progress',
    priority: 'Medium',
    pct: 30,
    metric: 'Pass official B2 DELE practice test with 70%+ score',
    notes: '',
    milestones: [
      { name: 'Complete Duolingo streak for 30 days', date: '2025-02-15', status: 'Done',        done: true,  notes: '' },
      { name: 'Finish beginner course (A2)',           date: '2025-05-31', status: 'In Progress', done: false, notes: '' },
      { name: 'Complete intermediate course (B1)',     date: '2025-08-31', status: 'Not Started', done: false, notes: '' },
      { name: 'Pass B1 practice test',                date: '2025-09-30', status: 'Not Started', done: false, notes: '' },
      { name: 'Complete B2 course + pass practice test',date: '2025-11-30', status: 'Not Started', done: false, notes: '' },
    ],
  },
  {
    title: 'Travel to Japan',
    area: 'Travel & Adventure',
    why: 'Bucket list trip — experience a completely different culture',
    target: '2025-09-20',
    status: 'On Track',
    priority: 'Medium',
    pct: 80,
    metric: 'Complete 14-day trip to Japan — Tokyo, Kyoto, Osaka',
    notes: '',
    milestones: [
      { name: 'Research itinerary and dates',          date: '2025-01-31', status: 'Done', done: true, notes: '' },
      { name: 'Book flights',                           date: '2025-02-28', status: 'Done', done: true, notes: '' },
      { name: 'Book accommodation',                    date: '2025-03-31', status: 'Done', done: true, notes: '' },
      { name: 'Plan daily activities + book experiences', date: '2025-07-31', status: 'Done', done: true, notes: '' },
      { name: 'Complete Japan trip',                   date: '2025-09-20', status: 'Not Started', done: false, notes: '' },
    ],
  },
  {
    title: 'Write and Publish a Blog',
    area: 'Creativity & Hobbies',
    why: 'Share my ideas and build a personal brand',
    target: '2025-12-31',
    status: 'At Risk',
    priority: 'Low',
    pct: 15,
    metric: 'Published blog with 20+ posts and 500 monthly readers',
    notes: 'Need to prioritise time for writing',
    milestones: [
      { name: 'Choose niche and domain name',  date: '2025-02-28', status: 'Done',        done: true,  notes: '' },
      { name: 'Set up website',                date: '2025-04-30', status: 'In Progress', done: false, notes: '' },
      { name: 'Publish first 5 posts',         date: '2025-07-31', status: 'Not Started', done: false, notes: '' },
      { name: 'Publish 20 posts total',        date: '2025-10-31', status: 'Not Started', done: false, notes: '' },
      { name: 'Reach 500 monthly readers',     date: '2025-12-31', status: 'Not Started', done: false, notes: '' },
    ],
  },
  null, null, null, null,  // Goals 7-10: blank
];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'🎯 Goals & Milestones'";

  // Column widths A-K
  [240,130,260,120,100,120,100,90,160,220,220].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: GL, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Row 0: title banner — A0 standalone (for col-A freeze), B0:K0 merged
  reqs.push({ repeatCell: {
    range: gridRange(GL,0,1,0,1),
    cell: { userEnteredFormat: { backgroundColor: hex(C.midnightBlue) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ mergeCells: { range: gridRange(GL,0,1,1,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(GL,0,1,1,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midnightBlue),
      textFormat: { foregroundColor: hex(C.electricGold), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GL, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!B1`, values: [['  🎯  Goals & Milestones — Alex\'s 2025 Goal Planner']] });

  // Row 1: column headers (0-indexed 1 = GSheets row 2)
  reqs.push({ repeatCell: {
    range: gridRange(GL,1,2,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepNavy),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GL, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A2`, values: [COL_HEADERS] });

  // Process each of 10 goals
  for (let gi = 0; gi < 10; gi++) {
    const headerRi = GOAL_ROWS_0IDX[gi];       // 0-indexed
    const headerGs = headerRi + 1;              // GSheets 1-indexed
    const goal = GOALS[gi];

    // Goal header row styling
    reqs.push({ repeatCell: {
      range: gridRange(GL, headerRi, headerRi+1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.midnightBlue),
        textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: GL, dimension: 'ROWS', startIndex: headerRi, endIndex: headerRi+1 },
      properties: { pixelSize: 32 }, fields: 'pixelSize',
    }});

    // E col (Days Left) and I col (Progress Bar) on goal header — formula/auto cell bg
    reqs.push({ repeatCell: {
      range: gridRange(GL, headerRi, headerRi+1, 4, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.deepNavy) }},
      fields: 'userEnteredFormat.backgroundColor',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(GL, headerRi, headerRi+1, 8, 9),
      cell: { userEnteredFormat: { backgroundColor: hex(C.deepNavy) }},
      fields: 'userEnteredFormat.backgroundColor',
    }});

    if (goal) {
      const daysLeft  = `=IFERROR(IF(D${headerGs}="","",D${headerGs}-TODAY()),"")`;
      const sparkline = `=SPARKLINE(H${headerGs}/100,{"charttype","bar";"color1",IF(H${headerGs}=100,"#2ECC71",IF(H${headerGs}>=50,"#F4D03F","#C9A800"));"max",1})`;
      vals.push({ range: `${TAB}!A${headerGs}`, values: [[
        goal.title, goal.area, goal.why, goal.target, daysLeft,
        goal.status, goal.priority, goal.pct, sparkline,
        goal.metric, goal.notes,
      ]]});
    }

    // 5 milestone rows
    for (let mi = 0; mi < 5; mi++) {
      const msRi = headerRi + 1 + mi;  // 0-indexed
      const msGs = msRi + 1;           // GSheets 1-indexed
      const ms   = goal ? goal.milestones[mi] : null;

      // Milestone row styling
      reqs.push({ repeatCell: {
        range: gridRange(GL, msRi, msRi+1, 0, NCOLS),
        cell: { userEnteredFormat: {
          backgroundColor: hex(C.milestoneRow),
          textFormat: { foregroundColor: hex(C.medSlate), fontSize: 10 },
          verticalAlignment: 'MIDDLE',
        }},
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      }});
      reqs.push({ updateDimensionProperties: {
        range: { sheetId: GL, dimension: 'ROWS', startIndex: msRi, endIndex: msRi+1 },
        properties: { pixelSize: 24 }, fields: 'pixelSize',
      }});
      // E col days left formula bg
      reqs.push({ repeatCell: {
        range: gridRange(GL, msRi, msRi+1, 4, 5),
        cell: { userEnteredFormat: { backgroundColor: hex(C.paleBlue) }},
        fields: 'userEnteredFormat.backgroundColor',
      }});

      const msDaysLeft = `=IFERROR(IF(D${msGs}="","",D${msGs}-TODAY()),"")`;
      const msName = ms ? `  ↳ ${ms.name}` : `  ↳ Milestone ${mi+1}`;

      if (ms) {
        vals.push({ range: `${TAB}!A${msGs}`, values: [[
          msName, '', '', ms.date, msDaysLeft,
          ms.status, '', ms.done ? 'TRUE' : 'FALSE', '', '', ms.notes,
        ]]});
      } else {
        vals.push({ range: `${TAB}!A${msGs}`, values: [[
          msName, '', '', '', msDaysLeft,
          '', '', 'FALSE', '', '', '',
        ]]});
      }
    }
  }

  // Borders on full data range
  reqs.push({ updateBorders: {
    range: gridRange(GL, 1, 62, 0, NCOLS),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Date column D format on goal header rows
  GOAL_ROWS_0IDX.forEach(ri => {
    reqs.push({ repeatCell: {
      range: gridRange(GL, ri, ri+1, 3, 4),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' }, backgroundColor: hex(C.midnightBlue) }},
      fields: 'userEnteredFormat(numberFormat,backgroundColor)',
    }});
  });
  // Days Left: right-align and number format on goal header rows
  GOAL_ROWS_0IDX.forEach(ri => {
    reqs.push({ repeatCell: {
      range: gridRange(GL, ri, ri+1, 4, 5),
      cell: { userEnteredFormat: {
        numberFormat: { type: 'NUMBER', pattern: '0' },
        horizontalAlignment: 'CENTER',
        backgroundColor: hex(C.deepNavy),
      }},
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment,backgroundColor)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(GL, ri, ri+1, 7, 8),
      cell: { userEnteredFormat: {
        numberFormat: { type: 'NUMBER', pattern: '0"%"' },
        horizontalAlignment: 'CENTER',
        backgroundColor: hex(C.midnightBlue),
      }},
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment,backgroundColor)',
    }});
  });

  await batchUpdate(id, reqs, 'goals-format');
  await valuesBatchUpdate(id, vals, 'goals-values');
  console.log('Goals & Milestones complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
