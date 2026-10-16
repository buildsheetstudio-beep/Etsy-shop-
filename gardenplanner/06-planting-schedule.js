'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const PS = sheetMap['Planting Schedule'];

// 15 planting rows with realistic 2025 dates
const PLANTING_DATA = [
  // Plant Name, Season, Indoors/Direct, Seed Date, Transplant Date, Zone, Qty, Status, Notes
  ['Tomato (Cherry)',   'Summer', 'Indoors',  '2025-03-01', '2025-05-01', 'Zone B — Tomatoes',   12, 'Germinated', 'Under grow lights 16hrs/day'],
  ['Tomato (Beefsteak)','Summer', 'Indoors',  '2025-03-01', '2025-05-08', 'Zone B — Tomatoes',   6,  'Germinated', 'Harden off 2 weeks before transplant'],
  ['Bell Pepper',       'Summer', 'Indoors',  '2025-02-15', '2025-05-15', 'Zone B — Tomatoes',   8,  'Transplanted','Start 10 weeks before last frost'],
  ['Jalapeño',          'Summer', 'Indoors',  '2025-02-15', '2025-05-15', 'Zone B — Tomatoes',   4,  'Growing',    'Same schedule as bell pepper'],
  ['Cucumber',          'Summer', 'Direct',   '2025-05-15', '',           'Zone E — Cucurbits',  8,  'Seeded',     'Direct sow after frost; thin to 12"'],
  ['Zucchini',          'Summer', 'Direct',   '2025-05-10', '',           'Zone E — Cucurbits',  4,  'Germinated', 'Plant 2 per hill; thin to 1'],
  ['Pumpkin',           'Summer', 'Direct',   '2025-05-20', '',           'Zone E — Cucurbits',  6,  'Seeded',     'Needs 90 days; plant late May'],
  ['Basil',             'Summer', 'Indoors',  '2025-04-01', '2025-05-20', 'Zone A — Herbs',      12, 'Transplanted','Pinch after 6 leaves appear'],
  ['Lettuce (Romaine)', 'Spring', 'Direct',   '2025-03-15', '',           'Zone D — Greens',     20, 'Harvested',  'Succession sow every 3 weeks'],
  ['Spinach',           'Spring', 'Direct',   '2025-03-10', '',           'Zone D — Greens',     15, 'Harvested',  'First sow of season — done'],
  ['Kale',              'Fall',   'Direct',   '2025-08-01', '',           'Zone D — Greens',     10, 'Planned',    'Second planting for fall harvest'],
  ['Carrot',            'Spring', 'Direct',   '2025-04-01', '',           'Zone C — Root Veg',   30, 'Growing',    'Thin to 2" when 2" tall'],
  ['Radish',            'Spring', 'Direct',   '2025-03-20', '',           'Zone C — Root Veg',   20, 'Harvested',  'Ready in 25 days — first to harvest'],
  ['Garlic',            'All Season','Direct', '2024-10-15', '',          'Zone C — Root Veg',   24, 'Growing',    'Fall planted; harvest July 2025'],
  ['Marigold',          'Summer', 'Indoors',  '2025-04-01', '2025-05-10', 'Zone F — Flowers',    20, 'Transplanted','Pest repellent; plant borders'],
];

(async () => {
  const reqs = [];
  const values = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(PS, 0, 1, 0, 16), mergeType: 'MERGE_ALL' } });
  values.push({ range: 'Planting Schedule!A1', values: [['📅 PLANTING SCHEDULE — Season 2025']] });
  reqs.push({
    repeatCell: {
      range: gridRange(PS, 0, 1, 0, 16),
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
    range: { sheetId: PS, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 44 }, fields: 'pixelSize',
  }});

  // ── Instructions row ──────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(PS, 1, 2, 0, 16), mergeType: 'MERGE_ALL' } });
  values.push({ range: 'Planting Schedule!A2', values: [['Track seed start dates, transplant dates, germination info, and succession planting. Germination & harvest days auto-calculate from Plant Library.']] });
  reqs.push({
    repeatCell: {
      range: gridRange(PS, 1, 2, 0, 16),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.paleSage),
          textFormat: { foregroundColor: hex(C.soil), italic: true, fontSize: 9 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Header row ────────────────────────────────────────────────────────────
  const headers = [
    'Plant Name', 'Season', 'Method', 'Seed/Start Date', 'Est. Germ Date',
    'Transplant Date', 'Est. Harvest Date', 'Zone', 'Qty Planted',
    'Days to Germ', 'Days to Harvest', 'Status', 'Succession #', 'Notes',
  ];
  values.push({ range: 'Planting Schedule!A3', values: [headers] });
  reqs.push({
    repeatCell: {
      range: gridRange(PS, 2, 3, 0, 14),
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
    range: { sheetId: PS, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // ── Data rows with formulas ───────────────────────────────────────────────
  for (let i = 0; i < PLANTING_DATA.length; i++) {
    const [plant, season, method, seedDate, transplantDate, zone, qty, status, notes] = PLANTING_DATA[i];
    const rowNum = i + 4; // 1-indexed row number

    // Col D = seed date, Col E = est germ date (D + VLOOKUP days to germ from Plant Library)
    // Col G = est harvest date (D + VLOOKUP days to harvest)
    // Col J = days to germ from Plant Library
    // Col K = days to harvest from Plant Library

    values.push({ range: `Planting Schedule!A${rowNum}`, values: [[
      plant, season, method, seedDate,
      `=IFERROR(IF(D${rowNum}<>"",D${rowNum}+VLOOKUP(A${rowNum},'Plant Library'!$A$3:$G$37,6,FALSE),""),"")`,
      transplantDate,
      `=IFERROR(IF(D${rowNum}<>"",D${rowNum}+VLOOKUP(A${rowNum},'Plant Library'!$A$3:$G$37,7,FALSE),""),"")`,
      zone, qty,
      `=IFERROR(VLOOKUP(A${rowNum},'Plant Library'!$A$3:$G$37,6,FALSE),"")`,
      `=IFERROR(VLOOKUP(A${rowNum},'Plant Library'!$A$3:$G$37,7,FALSE),"")`,
      status, 1, notes,
    ]]});
  }

  // ── Format data rows ──────────────────────────────────────────────────────
  for (let i = 0; i < PLANTING_DATA.length; i++) {
    const bg = i % 2 === 0 ? C.cream2 : C.parchment;
    reqs.push({
      repeatCell: {
        range: gridRange(PS, 3 + i, 4 + i, 0, 14),
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
      range: gridRange(PS, 3, 3 + PLANTING_DATA.length, 0, 1),
      cell: { userEnteredFormat: { textFormat: { bold: true } } },
      fields: 'userEnteredFormat.textFormat.bold',
    },
  });

  // Date format cols D, E, F, G (cols 3, 4, 5, 6)
  [3, 4, 5, 6].forEach(col => {
    reqs.push({
      repeatCell: {
        range: gridRange(PS, 3, 3 + PLANTING_DATA.length, col, col + 1),
        cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'MMM d, yyyy' } } },
        fields: 'userEnteredFormat.numberFormat',
      },
    });
  });

  // Center align most columns
  reqs.push({
    repeatCell: {
      range: gridRange(PS, 3, 3 + PLANTING_DATA.length, 1, 13),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat.horizontalAlignment',
    },
  });

  // Wrap notes
  reqs.push({
    repeatCell: {
      range: gridRange(PS, 3, 3 + PLANTING_DATA.length, 13, 14),
      cell: { userEnteredFormat: { wrapStrategy: 'WRAP', horizontalAlignment: 'LEFT' } },
      fields: 'userEnteredFormat(wrapStrategy,horizontalAlignment)',
    },
  });

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: PS, dimension: 'ROWS', startIndex: 3, endIndex: 3 + PLANTING_DATA.length },
    properties: { pixelSize: 26 }, fields: 'pixelSize',
  }});

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [150, 80, 80, 110, 110, 110, 110, 150, 80, 90, 90, 110, 90, 180];
  colWidths.forEach((px, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: PS, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  // ── Borders ───────────────────────────────────────────────────────────────
  reqs.push({
    updateBorders: {
      range: gridRange(PS, 2, 3 + PLANTING_DATA.length, 0, 14),
      innerHorizontal: { style: 'SOLID', color: hex(C.lightGray) },
      innerVertical: { style: 'SOLID', color: hex(C.lightGray) },
      bottom: { style: 'SOLID_MEDIUM', color: hex(C.olive) },
      top: { style: 'SOLID_MEDIUM', color: hex(C.olive) },
      left: { style: 'SOLID_MEDIUM', color: hex(C.olive) },
      right: { style: 'SOLID_MEDIUM', color: hex(C.olive) },
    },
  });

  // ── Succession planting summary block below ───────────────────────────────
  const sumRow = 3 + PLANTING_DATA.length + 2;
  reqs.push({ mergeCells: { range: gridRange(PS, sumRow, sumRow + 1, 0, 6), mergeType: 'MERGE_ALL' } });
  values.push({ range: `Planting Schedule!A${sumRow + 1}`, values: [['📊 SUCCESSION PLANTING SUMMARY']] });
  reqs.push({
    repeatCell: {
      range: gridRange(PS, sumRow, sumRow + 1, 0, 6),
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

  values.push({ range: `Planting Schedule!A${sumRow + 2}`, values: [['Total Plants Started:', `=COUNTA(A4:A${3 + PLANTING_DATA.length})`, '', 'Spring Plants:', `=COUNTIF(B4:B${3+PLANTING_DATA.length},"Spring")`, '', 'Summer Plants:', `=COUNTIF(B4:B${3+PLANTING_DATA.length},"Summer")`]] });
  values.push({ range: `Planting Schedule!A${sumRow + 3}`, values: [['Germinated:', `=COUNTIF(L4:L${3+PLANTING_DATA.length},"Germinated")`, '', 'Transplanted:', `=COUNTIF(L4:L${3+PLANTING_DATA.length},"Transplanted")`, '', 'Harvested:', `=COUNTIF(L4:L${3+PLANTING_DATA.length},"Harvested")`]] });

  reqs.push({
    repeatCell: {
      range: gridRange(PS, sumRow + 1, sumRow + 4, 0, 10),
      cell: { userEnteredFormat: { backgroundColor: hex(C.paleSage), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(PS, sumRow + 1, sumRow + 4, 0, 1),
      cell: { userEnteredFormat: { textFormat: { bold: true } } },
      fields: 'userEnteredFormat.textFormat.bold',
    },
  });

  // ── Freeze ────────────────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: PS, gridProperties: { frozenRowCount: 3 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await valuesBatchUpdate(id, values, 'planting-values');
  await batchUpdate(id, reqs, 'planting-format');
  console.log('Planting Schedule complete — 15 plants');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
