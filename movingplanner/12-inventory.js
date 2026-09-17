'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const INV = sheetMap['📊 Inventory List'];

(async () => {
  const values = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  values.push({ range: "'📊 Inventory List'!A1", values: [['📊 All-in-One Moving Planner — Inventory List']] });

  // ── Summary block ─────────────────────────────────────────────────────────
  values.push({ range: "'📊 Inventory List'!A2", values: [[
    'Total Items', 'Total Est. Value', 'Items by Room — Heaviest Room', 'Items Damaged/Poor',
  ]]});
  values.push({ range: "'📊 Inventory List'!A3", values: [[
    '=COUNTA(A6:A35)',
    '=SUM(H6:H35)',
    '=IFERROR(INDEX(B6:B35,MATCH(MAXIFS(COUNTIF(B6:B35,B6:B35),B6:B35,B6:B35),COUNTIF(B6:B35,B6:B35),0)),"—")',
    '=COUNTIF(I6:I35,"Poor")',
  ]]});

  // ── Column headers (row 5) ────────────────────────────────────────────────
  values.push({ range: "'📊 Inventory List'!A5", values: [[
    '#', 'Room', 'Item Description', 'Brand / Model', 'Serial Number',
    'Box #', 'Qty', 'Est. Value ($)', 'Condition', 'Notes',
  ]]});

  // ── 10 sample inventory items ─────────────────────────────────────────────
  const items = [
    [1,  'Living Room',    '65" LED Smart TV',              'Samsung',    'SN-778899',  '',   1,  1200, 'Excellent', 'Original box kept'],
    [2,  'Living Room',    'Three-seater sofa',             'IKEA',       '',           '',   1,  600,  'Good',      'Will fit new lounge'],
    [3,  'Living Room',    'Bookshelf — tall 5-shelf',      'IKEA',       '',           '',   1,  180,  'Good',      'Disassemble for move'],
    [4,  'Kitchen',        'Dishwasher',                    'Bosch',      'BSH-112233', '',   1,  900,  'Good',      'Stays with current property'],
    [5,  'Kitchen',        'KitchenAid Stand Mixer',        'KitchenAid', 'KA-445566',  '3',  1,  550,  'Excellent', ''],
    [6,  'Master Bedroom', 'Queen Bed Frame — timber',      'Bed Bath',   '',           '',   1,  700,  'Good',      'Disassemble'],
    [7,  'Master Bedroom', 'Wardrobe — 3 door',             'IKEA PAX',   '',           '',   1,  450,  'Good',      'Measure door frame first'],
    [8,  'Home Office',    'Desk — L-shaped',               'IKEA',       '',           '',   1,  320,  'Good',      'Disassemble for move'],
    [9,  'Home Office',    'Office Chair — ergonomic',      'Ergohuman',  'EH-887766',  '',   1,  480,  'Excellent', ''],
    [10, 'Garage',         'Lawnmower — petrol',            'Husqvarna',  'HQ-334455',  '',   1,  650,  'Good',      'Drain fuel before move'],
  ];

  const itemRows = items.map(item => [
    item[0], item[1], item[2], item[3], item[4],
    item[5], item[6], item[7], item[8], item[9],
  ]);
  values.push({ range: "'📊 Inventory List'!A6", values: itemRows });

  await valuesBatchUpdate(id, values, 'inventory-values');

  // ── Merge title ───────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(INV, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(INV, 0, 1, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.olive),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Summary header row 2 ──────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(INV, 1, 2, 0, 4),
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
      range: gridRange(INV, 2, 3, 0, 4),
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
  // Currency for total value
  reqs.push({
    repeatCell: {
      range: gridRange(INV, 2, 3, 1, 2),
      cell: {
        userEnteredFormat: {
          numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' },
        },
      },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // ── Column headers row 5 ──────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(INV, 4, 5, 0, 10),
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

  // ── Zebra stripe data rows ────────────────────────────────────────────────
  for (let r = 0; r < 10; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(INV, 5 + r, 6 + r, 0, 10),
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

  // Currency for Est. Value column (H = index 7)
  reqs.push({
    repeatCell: {
      range: gridRange(INV, 5, 15, 7, 8),
      cell: {
        userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } },
      },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [35, 110, 230, 110, 110, 55, 45, 90, 80, 200];
  colWidths.forEach((w, i) => {
    reqs.push({
      updateDimensionProperties: {
        range: { sheetId: INV, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
        properties: { pixelSize: w },
        fields: 'pixelSize',
      },
    });
  });

  // ── Row heights ───────────────────────────────────────────────────────────
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: INV, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
      properties: { pixelSize: 40 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: INV, dimension: 'ROWS', startIndex: 1, endIndex: 5 },
      properties: { pixelSize: 28 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: INV, dimension: 'ROWS', startIndex: 5, endIndex: 15 },
      properties: { pixelSize: 24 },
      fields: 'pixelSize',
    },
  });

  // ── Freeze rows 1–5 ───────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: INV, gridProperties: { frozenRowCount: 5 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await batchUpdate(id, reqs, 'inventory-format');
  console.log('Inventory List complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
