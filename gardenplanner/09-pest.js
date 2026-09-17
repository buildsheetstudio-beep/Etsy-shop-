'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const PEST = sheetMap['Pest & Disease Log'];

// 6 pest/disease entries
const PEST_DATA = [
  // Affected Plant, Date Found, Type, Pest/Disease Name, Severity, Treatment Used, Treatment Date, Status, Notes
  ['Tomato (Cherry)',   '2025-05-15', 'Insect Pest',      'Aphids',               'Low',     'Neem oil spray',          '2025-05-16', 'Resolved',   'Found on undersides of leaves; repeated once'],
  ['Zucchini',         '2025-06-01', 'Fungal Disease',   'Powdery Mildew',       'Medium',  'Baking soda + copper spray','2025-06-02','Treating',   'Common in humid weather; remove affected leaves'],
  ['Tomato (Beefsteak)','2025-06-10', 'Bacterial Disease','Blossom End Rot',      'Medium',  'Calcium spray + mulch',    '2025-06-11','Treating',   'Related to inconsistent watering; not contagious'],
  ['Carrot',           '2025-06-12', 'Insect Pest',      'Carrot Rust Fly',      'Low',     'Row covers applied',       '2025-06-12','Monitoring', 'Preventive covers; no larvae found yet'],
  ['Kale',             '2025-06-18', 'Insect Pest',      'Cabbage White Caterpillar','High', 'Bt (Bacillus thuringiensis)','2025-06-19','Treating', 'Hand-picked 20+ caterpillars; Bt spray applied'],
  ['Cucumber',         '2025-06-25', 'Fungal Disease',   'Cucumber Mosaic Virus', 'High',   'Remove infected plants',   '2025-06-25','Treating',   'Spread by aphids; removed 2 plants; monitor others'],
];

(async () => {
  const reqs = [];
  const values = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(PEST, 0, 1, 0, 14), mergeType: 'MERGE_ALL' } });
  values.push({ range: 'Pest & Disease Log!A1', values: [['🐛 PEST & DISEASE LOG — Season 2025']] });
  reqs.push({
    repeatCell: {
      range: gridRange(PEST, 0, 1, 0, 14),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.warningRed),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: PEST, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 44 }, fields: 'pixelSize',
  }});

  // ── Header row ────────────────────────────────────────────────────────────
  const headers = [
    'Affected Plant', 'Date Found', 'Type', 'Pest/Disease Name',
    'Severity', 'Treatment Used', 'Treatment Date', 'Status', 'Notes',
  ];
  values.push({ range: 'Pest & Disease Log!A2', values: [headers] });
  reqs.push({
    repeatCell: {
      range: gridRange(PEST, 1, 2, 0, 9),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex('#922B21'),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: PEST, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});

  // ── Data rows ─────────────────────────────────────────────────────────────
  values.push({ range: 'Pest & Disease Log!A3', values: PEST_DATA });

  // Format data rows
  for (let i = 0; i < PEST_DATA.length; i++) {
    const bg = i % 2 === 0 ? '#FEF9F9' : C.warningBg;
    reqs.push({
      repeatCell: {
        range: gridRange(PEST, 2 + i, 3 + i, 0, 9),
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

  // Bold plant name
  reqs.push({
    repeatCell: {
      range: gridRange(PEST, 2, 2 + PEST_DATA.length, 0, 1),
      cell: { userEnteredFormat: { textFormat: { bold: true } } },
      fields: 'userEnteredFormat.textFormat.bold',
    },
  });

  // Date format
  [1, 6].forEach(col => {
    reqs.push({
      repeatCell: {
        range: gridRange(PEST, 2, 2 + PEST_DATA.length, col, col + 1),
        cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'MMM d, yyyy' } } },
        fields: 'userEnteredFormat.numberFormat',
      },
    });
  });

  // Center most cols
  reqs.push({
    repeatCell: {
      range: gridRange(PEST, 2, 2 + PEST_DATA.length, 1, 8),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat.horizontalAlignment',
    },
  });

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: PEST, dimension: 'ROWS', startIndex: 2, endIndex: 2 + PEST_DATA.length },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [160, 110, 130, 180, 90, 180, 110, 100, 220];
  colWidths.forEach((px, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: PEST, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  // ── Borders ───────────────────────────────────────────────────────────────
  reqs.push({
    updateBorders: {
      range: gridRange(PEST, 1, 2 + PEST_DATA.length, 0, 9),
      innerHorizontal: { style: 'SOLID', color: hex(C.warningBg) },
      innerVertical: { style: 'SOLID', color: hex(C.warningBg) },
      bottom: { style: 'SOLID_MEDIUM', color: hex(C.warningRed) },
      top: { style: 'SOLID_MEDIUM', color: hex(C.warningRed) },
      left: { style: 'SOLID_MEDIUM', color: hex(C.warningRed) },
      right: { style: 'SOLID_MEDIUM', color: hex(C.warningRed) },
    },
  });

  // ── Summary block ─────────────────────────────────────────────────────────
  const sumRow = 2 + PEST_DATA.length + 2;
  reqs.push({ mergeCells: { range: gridRange(PEST, sumRow, sumRow + 1, 0, 9), mergeType: 'MERGE_ALL' } });
  values.push({ range: `Pest & Disease Log!A${sumRow + 1}`, values: [['📊 PEST & DISEASE SUMMARY']] });
  reqs.push({
    repeatCell: {
      range: gridRange(PEST, sumRow, sumRow + 1, 0, 9),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.olive),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  const lastRow = 2 + PEST_DATA.length;
  values.push({ range: `Pest & Disease Log!A${sumRow + 2}`, values: [[
    'Total Incidents:', `=COUNTA(A3:A${lastRow})`, '',
    'Active (Treating):', `=COUNTIF(H3:H${lastRow},"Treating")`, '',
    'Resolved:', `=COUNTIF(H3:H${lastRow},"Resolved")`, '',
    'Monitoring:', `=COUNTIF(H3:H${lastRow},"Monitoring")`,
  ]]});
  values.push({ range: `Pest & Disease Log!A${sumRow + 3}`, values: [[
    'Insect Pests:', `=COUNTIF(C3:C${lastRow},"Insect Pest")`, '',
    'Fungal Diseases:', `=COUNTIF(C3:C${lastRow},"Fungal Disease")`, '',
    'Bacterial:', `=COUNTIF(C3:C${lastRow},"Bacterial Disease")`, '',
    'High Severity:', `=COUNTIF(E3:E${lastRow},"High")`,
  ]]});

  reqs.push({
    repeatCell: {
      range: gridRange(PEST, sumRow + 1, sumRow + 4, 0, 9),
      cell: { userEnteredFormat: { backgroundColor: hex(C.paleTerracotta), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(PEST, sumRow + 1, sumRow + 4, 0, 1),
      cell: { userEnteredFormat: { textFormat: { bold: true } } },
      fields: 'userEnteredFormat.textFormat.bold',
    },
  });

  // ── Chart helper table (N:O columns, col 13:14) ───────────────────────────
  const chartRow = sumRow + 5;
  reqs.push({ mergeCells: { range: gridRange(PEST, chartRow, chartRow + 1, 0, 4), mergeType: 'MERGE_ALL' } });
  values.push({ range: `Pest & Disease Log!A${chartRow + 1}`, values: [['Chart Helper — Pest by Type']] });
  reqs.push({
    repeatCell: {
      range: gridRange(PEST, chartRow, chartRow + 1, 0, 4),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.terracotta),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  const pestTypes = ['Insect Pest', 'Fungal Disease', 'Bacterial Disease', 'Viral Disease', 'Animal Pest', 'Other'];
  values.push({ range: `Pest & Disease Log!A${chartRow + 2}`, values: [['Type', 'Count']] });
  const typeRows = pestTypes.map(t => [t, `=COUNTIF(C3:C${lastRow},"${t}")`]);
  values.push({ range: `Pest & Disease Log!A${chartRow + 3}`, values: typeRows });

  reqs.push({
    repeatCell: {
      range: gridRange(PEST, chartRow + 1, chartRow + 2, 0, 2),
      cell: { userEnteredFormat: { backgroundColor: hex(C.olive), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(PEST, chartRow + 2, chartRow + 2 + pestTypes.length, 0, 2),
      cell: { userEnteredFormat: { backgroundColor: hex(C.cream2), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // ── Freeze ────────────────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: PEST, gridProperties: { frozenRowCount: 2 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await valuesBatchUpdate(id, values, 'pest-values');
  await batchUpdate(id, reqs, 'pest-format');
  console.log('Pest & Disease Log complete — 6 entries');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
