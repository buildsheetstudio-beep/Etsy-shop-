'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SEED = sheetMap['Seed Inventory'];

// 10 seed inventory entries
const SEED_DATA = [
  // Seed Name, Type, Season, Seed Source, Date Acquired, # Seeds, # Planted, Remaining, Germination Rate, Viability (years), Storage Location, Notes
  ['Tomato (Sun Gold F1)',  'Vegetable',     'Summer',    'Purchased', '2025-02-01', 25, 12, 13, '92%', 4, 'Refrigerator envelope', 'Open-pollinated; save next year'],
  ['Basil (Genovese)',      'Herb',          'Summer',    'Saved Seed','2024-09-01', 200,50, 150,'88%', 3, 'Glass jar, dark drawer',  'Saved from best plant in 2024'],
  ['Lettuce (Butterhead)',  'Leafy Green',   'Spring',    'Purchased', '2025-01-15', 100,30, 70, '85%', 3, 'Refrigerator envelope', 'Succession sow every 3 weeks'],
  ['Carrot (Nantes)',       'Root Vegetable','Spring',    'Purchased', '2024-12-10', 500,100,400,'78%', 3, 'Sealed tin, cool closet', 'Needs cold stratification'],
  ['Cucumber (Marketmore)', 'Vine',          'Summer',    'Gift/Trade','2025-03-01', 40, 8,  32, '90%', 5, 'Glass jar, dark drawer',  'From neighbor — great producer'],
  ['Kale (Lacinato)',       'Leafy Green',   'Fall',      'Purchased', '2025-03-15', 150,20, 130,'87%', 4, 'Refrigerator envelope', 'AKA Dinosaur kale; very hardy'],
  ['Radish (Cherry Belle)', 'Root Vegetable','Spring',    'Purchased', '2025-01-10', 300,60, 240,'95%', 4, 'Sealed tin, cool closet', 'Fast 25-day crop; succession sow'],
  ['Zucchini (Black Beauty)','Vegetable',    'Summer',    'Saved Seed','2024-09-15', 30, 4,  26, '92%', 5, 'Glass jar, dark drawer',  'Saved from hybrid; may vary'],
  ['Marigold (French)',     'Flower',        'Summer',    'Saved Seed','2024-10-01', 300,80, 220,'80%', 2, 'Paper envelope, drawer', 'Deadhead all season to collect'],
  ['Spinach (Bloomsdale)',  'Leafy Green',   'Spring',    'Purchased', '2024-11-01', 200,40, 160,'82%', 2, 'Refrigerator envelope', 'Short viability; refresh yearly'],
];

(async () => {
  const reqs = [];
  const values = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SEED, 0, 1, 0, 14), mergeType: 'MERGE_ALL' } });
  values.push({ range: 'Seed Inventory!A1', values: [['🌱 SEED INVENTORY — 2025 Season']] });
  reqs.push({
    repeatCell: {
      range: gridRange(SEED, 0, 1, 0, 14),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.oliveLight),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SEED, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 44 }, fields: 'pixelSize',
  }});

  // ── Header row ────────────────────────────────────────────────────────────
  const headers = [
    'Seed Name', 'Type', 'Season', 'Source', 'Date Acquired',
    '# Seeds', '# Planted', 'Remaining', 'Germ Rate', 'Viability (yrs)',
    'Storage Location', 'Notes',
  ];
  values.push({ range: 'Seed Inventory!A2', values: [headers] });
  reqs.push({
    repeatCell: {
      range: gridRange(SEED, 1, 2, 0, 12),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.olive),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SEED, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // ── Data rows ─────────────────────────────────────────────────────────────
  values.push({ range: 'Seed Inventory!A3', values: SEED_DATA });

  // Format data rows
  for (let i = 0; i < SEED_DATA.length; i++) {
    const bg = i % 2 === 0 ? C.cream2 : '#F0F7EC';
    reqs.push({
      repeatCell: {
        range: gridRange(SEED, 2 + i, 3 + i, 0, 12),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { fontSize: 10 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }

  // Bold seed name
  reqs.push({
    repeatCell: {
      range: gridRange(SEED, 2, 2 + SEED_DATA.length, 0, 1),
      cell: { userEnteredFormat: { textFormat: { bold: true } } },
      fields: 'userEnteredFormat.textFormat.bold',
    },
  });

  // Date format
  reqs.push({
    repeatCell: {
      range: gridRange(SEED, 2, 2 + SEED_DATA.length, 4, 5),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'MMM yyyy' }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
    },
  });

  // Center align numeric/category cols
  reqs.push({
    repeatCell: {
      range: gridRange(SEED, 2, 2 + SEED_DATA.length, 1, 11),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat.horizontalAlignment',
    },
  });

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SEED, dimension: 'ROWS', startIndex: 2, endIndex: 2 + SEED_DATA.length },
    properties: { pixelSize: 26 }, fields: 'pixelSize',
  }});

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [170, 110, 80, 90, 100, 70, 80, 80, 80, 90, 180, 200];
  colWidths.forEach((px, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SEED, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  // ── Borders ───────────────────────────────────────────────────────────────
  reqs.push({
    updateBorders: {
      range: gridRange(SEED, 1, 2 + SEED_DATA.length, 0, 12),
      innerHorizontal: { style: 'SOLID', color: hex(C.lightGray) },
      innerVertical: { style: 'SOLID', color: hex(C.lightGray) },
      bottom: { style: 'SOLID_MEDIUM', color: hex(C.olive) },
      top: { style: 'SOLID_MEDIUM', color: hex(C.olive) },
      left: { style: 'SOLID_MEDIUM', color: hex(C.olive) },
      right: { style: 'SOLID_MEDIUM', color: hex(C.olive) },
    },
  });

  // ── Summary block ─────────────────────────────────────────────────────────
  const sumRow = 2 + SEED_DATA.length + 2;
  reqs.push({ mergeCells: { range: gridRange(SEED, sumRow, sumRow + 1, 0, 6), mergeType: 'MERGE_ALL' } });
  values.push({ range: `Seed Inventory!A${sumRow + 1}`, values: [['📊 SEED INVENTORY SUMMARY']] });
  reqs.push({
    repeatCell: {
      range: gridRange(SEED, sumRow, sumRow + 1, 0, 6),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.terracotta),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  const lastRow = 2 + SEED_DATA.length;
  values.push({ range: `Seed Inventory!A${sumRow + 2}`, values: [[
    'Total Varieties:', `=COUNTA(A3:A${lastRow})`, '',
    'Total Seeds:', `=SUM(F3:F${lastRow})`, '',
  ]]});
  values.push({ range: `Seed Inventory!A${sumRow + 3}`, values: [[
    'Total Planted:', `=SUM(G3:G${lastRow})`, '',
    'Total Remaining:', `=SUM(H3:H${lastRow})`, '',
  ]]});
  values.push({ range: `Seed Inventory!A${sumRow + 4}`, values: [[
    'Saved Seeds:', `=COUNTIF(D3:D${lastRow},"Saved Seed")`, '',
    'Purchased:', `=COUNTIF(D3:D${lastRow},"Purchased")`, '',
  ]]});

  reqs.push({
    repeatCell: {
      range: gridRange(SEED, sumRow + 1, sumRow + 5, 0, 6),
      cell: { userEnteredFormat: { backgroundColor: hex(C.paleSage), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(SEED, sumRow + 1, sumRow + 5, 0, 1),
      cell: { userEnteredFormat: { textFormat: { bold: true } } },
      fields: 'userEnteredFormat.textFormat.bold',
    },
  });

  // ── Low seed alerts ───────────────────────────────────────────────────────
  const alertRow = sumRow + 6;
  reqs.push({ mergeCells: { range: gridRange(SEED, alertRow, alertRow + 1, 0, 6), mergeType: 'MERGE_ALL' } });
  values.push({ range: `Seed Inventory!A${alertRow + 1}`, values: [['⚠️ REORDER ALERTS — Seeds with < 20 remaining']] });
  reqs.push({
    repeatCell: {
      range: gridRange(SEED, alertRow, alertRow + 1, 0, 6),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.warningRed),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  values.push({ range: `Seed Inventory!A${alertRow + 2}`, values: [
    [`=IFERROR(FILTER(A3:A${lastRow},H3:H${lastRow}<20),"✅ All seeds well stocked!")`]
  ]});
  values.push({ range: `Seed Inventory!B${alertRow + 2}`, values: [
    [`=IFERROR(FILTER(H3:H${lastRow},H3:H${lastRow}<20),"")`]
  ]});

  reqs.push({
    repeatCell: {
      range: gridRange(SEED, alertRow + 1, alertRow + 2 + SEED_DATA.length, 0, 6),
      cell: { userEnteredFormat: { backgroundColor: hex(C.warningBg), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // ── Freeze ────────────────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: SEED, gridProperties: { frozenRowCount: 2 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await valuesBatchUpdate(id, values, 'seed-values');
  await batchUpdate(id, reqs, 'seed-format');
  console.log('Seed Inventory complete — 10 varieties');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
