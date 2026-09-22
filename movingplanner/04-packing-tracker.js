'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const PT = sheetMap['📦 Packing Tracker'];

(async () => {
  const values = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  values.push({ range: "'📦 Packing Tracker'!A1", values: [['📦 All-in-One Moving Planner — Packing Tracker']] });

  // ── Column headers ────────────────────────────────────────────────────────
  values.push({ range: "'📦 Packing Tracker'!A2", values: [[
    'Box #', 'Room', 'Contents Description', 'Pack Priority', 'Status',
    'Fragile?', 'Weight (kg)', 'Condition', 'Notes',
  ]]});

  // ── 20 pre-populated boxes ────────────────────────────────────────────────
  const boxes = [
    [1,  'Kitchen',        'Pots, pans, baking trays',                'Pack Last',  'To Pack', true,  8,  'Good',      ''],
    [2,  'Kitchen',        'Plates, bowls, mugs (wrapped)',           'Pack Last',  'To Pack', true,  6,  'Good',      'Bubble wrap each item'],
    [3,  'Kitchen',        'Appliances — toaster, kettle, blender',   'Standard',   'To Pack', false, 9,  'Excellent', ''],
    [4,  'Kitchen',        'Dry pantry goods — tins, jars, spices',   'Pack Last',  'To Pack', false, 7,  'Good',      'Check expiry dates'],
    [5,  'Living Room',    'Books and magazines',                      'Pack First', 'To Pack', false, 12, 'Good',      ''],
    [6,  'Living Room',    'DVDs, games, remotes',                     'Pack First', 'To Pack', false, 5,  'Good',      ''],
    [7,  'Living Room',    'Ornaments and picture frames',             'Standard',   'To Pack', true,  4,  'Good',      'Wrap in tissue paper'],
    [8,  'Living Room',    'Throw blankets, cushions, small rugs',    'Standard',   'To Pack', false, 6,  'Good',      ''],
    [9,  'Master Bedroom', 'Clothing — winter/off-season',            'Pack First', 'To Pack', false, 10, 'Good',      ''],
    [10, 'Master Bedroom', 'Bedding — sheets, doona covers',          'Pack First', 'To Pack', false, 8,  'Good',      ''],
    [11, 'Master Bedroom', 'Shoes and accessories',                   'Standard',   'To Pack', false, 7,  'Good',      ''],
    [12, 'Master Bedroom', 'Toiletries and bathroom supplies',        'Pack Last',  'To Pack', false, 5,  'Good',      ''],
    [13, 'Bedroom 2',      'Children\'s toys — large',               'Standard',   'To Pack', false, 8,  'Good',      ''],
    [14, 'Bedroom 2',      'Children\'s toys — small / figurines',   'Standard',   'To Pack', true,  4,  'Good',      'Bag small pieces'],
    [15, 'Bedroom 2',      'Children\'s books and school supplies',   'Pack First', 'To Pack', false, 6,  'Good',      ''],
    [16, 'Home Office',    'Stationery and files',                    'Pack First', 'To Pack', false, 9,  'Good',      'Label by category'],
    [17, 'Home Office',    'Computer monitor and cables',             'Standard',   'To Pack', true,  8,  'Excellent', 'Original box if possible'],
    [18, 'Garage',         'Tools — hand tools',                      'Pack First', 'To Pack', false, 10, 'Good',      ''],
    [19, 'Garage',         'Sports equipment — bikes, helmets',       'Standard',   'To Pack', false, 14, 'Good',      'Deflate tyres'],
    [20, 'Storage',        'Holiday decorations and seasonal items',  'Pack First', 'To Pack', false, 8,  'Good',      ''],
  ];

  const boxRows = boxes.map(b => [b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7], b[8]]);
  values.push({ range: "'📦 Packing Tracker'!A3", values: boxRows });

  // ── Room Summary header (col H–M, rows 2–3) ───────────────────────────────
  values.push({ range: "'📦 Packing Tracker'!K2", values: [
    ['Room', 'Total Boxes', 'Packed', 'Loaded', 'Arrived', 'Unpacked'],
  ]});
  const summaryRooms = ['Kitchen','Living Room','Master Bedroom','Bedroom 2','Bedroom 3','Home Office','Garage','Storage'];
  const summaryRows = summaryRooms.map(room => [
    room,
    `=COUNTIF($B$3:$B$22,"${room}")`,
    `=COUNTIFS($B$3:$B$22,"${room}",$E$3:$E$22,"Packed")`,
    `=COUNTIFS($B$3:$B$22,"${room}",$E$3:$E$22,"Loaded")`,
    `=COUNTIFS($B$3:$B$22,"${room}",$E$3:$E$22,"Arrived")`,
    `=COUNTIFS($B$3:$B$22,"${room}",$E$3:$E$22,"Unpacked")`,
  ]);
  values.push({ range: "'📦 Packing Tracker'!K3", values: summaryRows });

  // ── Summary totals row ────────────────────────────────────────────────────
  values.push({ range: "'📦 Packing Tracker'!K11", values: [
    ['TOTAL', '=SUM(L3:L10)', '=SUM(M3:M10)', '=SUM(N3:N10)', '=SUM(O3:O10)', '=SUM(P3:P10)'],
  ]});

  // ── Pack Status summary (top-left mini block) ─────────────────────────────
  values.push({ range: "'📦 Packing Tracker'!K13", values: [['Packing Progress']] });
  values.push({ range: "'📦 Packing Tracker'!K14", values: [
    ['To Pack', 'Packed', 'Loaded', 'Arrived', 'Unpacked', '% Done'],
  ]});
  values.push({ range: "'📦 Packing Tracker'!K15", values: [[
    '=COUNTIF($E$3:$E$22,"To Pack")',
    '=COUNTIF($E$3:$E$22,"Packed")',
    '=COUNTIF($E$3:$E$22,"Loaded")',
    '=COUNTIF($E$3:$E$22,"Arrived")',
    '=COUNTIF($E$3:$E$22,"Unpacked")',
    '=IFERROR((P15+O15)/COUNTA($E$3:$E$22),0)',
  ]]});

  await valuesBatchUpdate(id, values, 'packing-values');

  // ── Merge title ───────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(PT, 0, 1, 0, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(PT, 0, 1, 0, 9),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.mossGreen),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Column headers row 2 (A–I) ────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(PT, 1, 2, 0, 9),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepForest),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Room Summary header row 2 (K–P) ──────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(PT, 1, 2, 10, 16),
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

  // ── Zebra stripe data rows ────────────────────────────────────────────────
  for (let r = 0; r < 20; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(PT, 2 + r, 3 + r, 0, 9),
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

  // ── Room summary data rows (K–P) ─────────────────────────────────────────
  for (let r = 0; r < 8; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(PT, 2 + r, 3 + r, 10, 16),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(C.nearBlack), fontSize: 10 },
            horizontalAlignment: 'CENTER',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
      },
    });
  }

  // ── Totals row ────────────────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(PT, 10, 11, 10, 16),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepForest),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Progress section header ───────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(PT, 12, 13, 10, 16), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(PT, 12, 13, 10, 16),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.mossGreen),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(PT, 13, 14, 10, 16),
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
  // % Done cell
  reqs.push({
    repeatCell: {
      range: gridRange(PT, 14, 15, 15, 16),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.ivory),
          textFormat: { foregroundColor: hex(C.forestGreen), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER',
          numberFormat: { type: 'PERCENT', pattern: '0%' },
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,numberFormat)',
    },
  });

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [45, 110, 260, 95, 80, 65, 80, 80, 160, 20, 110, 80, 70, 70, 70, 80];
  colWidths.forEach((w, i) => {
    reqs.push({
      updateDimensionProperties: {
        range: { sheetId: PT, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
        properties: { pixelSize: w },
        fields: 'pixelSize',
      },
    });
  });

  // ── Row heights ───────────────────────────────────────────────────────────
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: PT, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
      properties: { pixelSize: 40 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: PT, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
      properties: { pixelSize: 28 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: PT, dimension: 'ROWS', startIndex: 2, endIndex: 22 },
      properties: { pixelSize: 24 },
      fields: 'pixelSize',
    },
  });

  await batchUpdate(id, reqs, 'packing-format');
  console.log('Packing Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
