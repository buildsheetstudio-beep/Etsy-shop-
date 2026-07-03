'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const TT = sheetMap['✅ Task Tracker'];

(async () => {
  const values = [];
  const reqs = [];

  // ── Title row ────────────────────────────────────────────────────────────────
  values.push({ range: "'✅ Task Tracker'!A1", values: [['✅ All-in-One Moving Planner — Task Tracker']] });

  // ── Summary block (row 2) ─────────────────────────────────────────────────
  values.push({ range: "'✅ Task Tracker'!A2", values: [
    ['Total Tasks', 'Done', 'In Progress', 'Not Started', 'Blocked', '% Complete'],
  ]});
  values.push({ range: "'✅ Task Tracker'!A3", values: [
    [
      '=COUNTA(A6:A65)',
      '=COUNTIF(E6:E65,"Done")',
      '=COUNTIF(E6:E65,"In Progress")',
      '=COUNTIF(E6:E65,"Not Started")',
      '=COUNTIF(E6:E65,"Blocked")',
      '=IFERROR(B3/A3,0)',
    ],
  ]});

  // ── Column headers (row 5) ────────────────────────────────────────────────
  values.push({ range: "'✅ Task Tracker'!A5", values: [[
    '#', 'Category', 'Task', 'Priority', 'Status', 'Due Date', 'Assigned To', 'Notes', 'Done?', 'Week',
  ]]});

  // ── 40 pre-populated tasks ────────────────────────────────────────────────
  const tasks = [
    // 8 Weeks Out
    [1,  '8 Weeks Out', 'Confirm move date with all parties',              'High',   'Not Started', '2025-07-05', 'Both', '', false, 1],
    [2,  '8 Weeks Out', 'Research moving companies — get 3+ quotes',       'High',   'Not Started', '2025-07-05', 'Both', 'See Quotes tab', false, 1],
    [3,  '8 Weeks Out', 'Start decluttering — donate/sell unwanted items', 'High',   'Not Started', '2025-07-05', 'Both', '', false, 1],
    [4,  '8 Weeks Out', 'Research schools/childcare in new area',          'Medium', 'Not Started', '2025-07-05', 'Both', 'See Schools tab', false, 1],
    [5,  '8 Weeks Out', 'Notify employer of address change (HR)',          'Medium', 'Not Started', '2025-07-07', 'Self', '', false, 1],
    [6,  '8 Weeks Out', 'Research neighbourhoods in new city',             'Medium', 'Not Started', '2025-07-07', 'Both', 'See Neighbourhood tab', false, 1],
    [7,  '8 Weeks Out', 'Obtain medical/dental records',                   'High',   'Not Started', '2025-07-07', 'Both', '', false, 1],
    [8,  '8 Weeks Out', 'Organise insurance — home & contents',            'High',   'Not Started', '2025-07-07', 'Self', '', false, 1],
    // 6 Weeks Out
    [9,  '6 Weeks Out', 'Book moving company',                             'High',   'Not Started', '2025-07-19', 'Self', '', false, 3],
    [10, '6 Weeks Out', 'Begin collecting packing boxes',                  'Medium', 'Not Started', '2025-07-19', 'Both', '', false, 3],
    [11, '6 Weeks Out', 'Sort out storage unit if needed',                 'Medium', 'Not Started', '2025-07-19', 'Self', '', false, 3],
    [12, '6 Weeks Out', 'Notify children\'s school of withdrawal',         'High',   'Not Started', '2025-07-19', 'Both', '', false, 3],
    [13, '6 Weeks Out', 'Apply to new school / childcare',                 'High',   'Not Started', '2025-07-21', 'Both', 'See Schools tab', false, 3],
    [14, '6 Weeks Out', 'Transfer vehicle registration if interstate',     'Medium', 'Not Started', '2025-07-21', 'Self', '', false, 3],
    [15, '6 Weeks Out', 'Arrange pet transport or vet records',            'Medium', 'Not Started', '2025-07-21', 'Both', '', false, 3],
    [16, '6 Weeks Out', 'Set moving budget',                               'High',   'Not Started', '2025-07-21', 'Both', 'See Budget tab', false, 3],
    // 4 Weeks Out
    [17, '4 Weeks Out', 'Start packing non-essentials',                    'High',   'Not Started', '2025-08-02', 'Both', '', false, 5],
    [18, '4 Weeks Out', 'Confirm moving date & time with movers',          'High',   'Not Started', '2025-08-02', 'Self', '', false, 5],
    [19, '4 Weeks Out', 'Notify bank of address change',                   'High',   'Not Started', '2025-08-02', 'Self', 'See Address tab', false, 5],
    [20, '4 Weeks Out', 'Redirect mail at post office',                    'High',   'Not Started', '2025-08-02', 'Self', '', false, 5],
    [21, '4 Weeks Out', 'Contact utilities — old home disconnect dates',   'High',   'Not Started', '2025-08-04', 'Self', 'See Utilities tab', false, 5],
    [22, '4 Weeks Out', 'Contact utilities — new home connect dates',      'High',   'Not Started', '2025-08-04', 'Self', 'See Utilities tab', false, 5],
    [23, '4 Weeks Out', 'Update subscriptions & delivery addresses',       'Medium', 'Not Started', '2025-08-04', 'Both', 'See Address tab', false, 5],
    [24, '4 Weeks Out', 'Arrange cleaning of current home',                'Medium', 'Not Started', '2025-08-04', 'Both', '', false, 5],
    // 2 Weeks Out
    [25, '2 Weeks Out', 'Pack remaining rooms — except essentials',        'High',   'Not Started', '2025-08-16', 'Both', '', false, 7],
    [26, '2 Weeks Out', 'Confirm school enrolment paperwork',              'High',   'Not Started', '2025-08-16', 'Both', '', false, 7],
    [27, '2 Weeks Out', 'Return library books, equipment etc.',            'Low',    'Not Started', '2025-08-16', 'Both', '', false, 7],
    [28, '2 Weeks Out', 'Measure new home rooms vs. furniture',            'Medium', 'Not Started', '2025-08-18', 'Both', '', false, 7],
    [29, '2 Weeks Out', 'Plan furniture layout for new home',              'Medium', 'Not Started', '2025-08-18', 'Both', '', false, 7],
    [30, '2 Weeks Out', 'Defrost freezer',                                 'Low',    'Not Started', '2025-08-18', 'Self', '', false, 7],
    // Moving Week
    [31, 'Moving Week',  'Pack essentials bag (toiletries, chargers, etc.)','High',  'Not Started', '2025-08-25', 'Both', '', false, 8],
    [32, 'Moving Week',  'Label all boxes clearly by room',                'High',   'Not Started', '2025-08-25', 'Both', '', false, 8],
    [33, 'Moving Week',  'Disassemble large furniture',                    'High',   'Not Started', '2025-08-25', 'Both', '', false, 8],
    [34, 'Moving Week',  'Take meter readings — old home',                 'High',   'Not Started', '2025-08-28', 'Self', 'Gas, electric, water', false, 8],
    [35, 'Moving Week',  'Clean old home after moving out',                'High',   'Not Started', '2025-08-29', 'Both', '', false, 8],
    // Moving Day
    [36, 'Moving Day',   'Supervise movers — check inventory list',        'High',   'Not Started', '2025-08-30', 'Both', 'See Inventory tab', false, 9],
    [37, 'Moving Day',   'Take meter readings — new home',                 'High',   'Not Started', '2025-08-30', 'Self', 'Gas, electric, water', false, 9],
    [38, 'Moving Day',   'Photograph condition of new home on arrival',    'High',   'Not Started', '2025-08-30', 'Both', '', false, 9],
    // Post-Move
    [39, 'Post-Move',    'Update drivers licence with new address',        'High',   'Not Started', '2025-09-06', 'Self', '', false, 10],
    [40, 'Post-Move',    'Register on electoral roll at new address',      'High',   'Not Started', '2025-09-06', 'Self', '', false, 10],
  ];

  const taskRows = tasks.map(t => [t[0], t[1], t[2], t[3], t[4], t[5], t[6], t[7], t[8], t[9]]);
  values.push({ range: "'✅ Task Tracker'!A6", values: taskRows });

  await valuesBatchUpdate(id, values, 'task-values');

  // ── Merge title ───────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(TT, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });

  // ── Title format ──────────────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(TT, 0, 1, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepForest),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER',
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Summary header row 2 ──────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(TT, 1, 2, 0, 6),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.antiqueBrass),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Summary value row 3 ───────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(TT, 2, 3, 0, 5),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.ivory),
          textFormat: { foregroundColor: hex(C.nearBlack), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  // % Complete cell — format as percent
  reqs.push({
    repeatCell: {
      range: gridRange(TT, 2, 3, 5, 6),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.ivory),
          textFormat: { foregroundColor: hex(C.forestGreen), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER',
          numberFormat: { type: 'PERCENT', pattern: '0%' },
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,numberFormat)',
    },
  });

  // ── Column header row 5 ───────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(TT, 4, 5, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.mossGreen),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Zebra stripe data rows ────────────────────────────────────────────────
  for (let r = 0; r < 40; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(TT, 5 + r, 6 + r, 0, 10),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(C.nearBlack), fontSize: 10 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [35, 100, 280, 80, 90, 90, 90, 180, 60, 50];
  colWidths.forEach((w, i) => {
    reqs.push({
      updateDimensionProperties: {
        range: { sheetId: TT, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
        properties: { pixelSize: w },
        fields: 'pixelSize',
      },
    });
  });

  // ── Row heights ───────────────────────────────────────────────────────────
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: TT, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
      properties: { pixelSize: 40 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: TT, dimension: 'ROWS', startIndex: 1, endIndex: 3 },
      properties: { pixelSize: 30 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: TT, dimension: 'ROWS', startIndex: 4, endIndex: 5 },
      properties: { pixelSize: 28 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: TT, dimension: 'ROWS', startIndex: 5, endIndex: 45 },
      properties: { pixelSize: 24 },
      fields: 'pixelSize',
    },
  });

  // ── Freeze rows 1–5 ───────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: TT, gridProperties: { frozenRowCount: 5 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await batchUpdate(id, reqs, 'task-format');
  console.log('Task Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
