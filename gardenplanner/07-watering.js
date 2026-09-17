'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const WAT = sheetMap['Watering & Care Schedule'];

// 12 rows of watering/care data
const WATERING_DATA = [
  // Plant, Zone, Frequency, Water Amount (gal), Last Watered, Next Due, Fertilize Schedule, Notes
  ['Tomato (Cherry)',    'Zone B — Tomatoes',   'Every 2 Days', 0.5, '2025-06-28', '', 'Bi-weekly',    'Deep water; mulch to retain moisture'],
  ['Tomato (Beefsteak)','Zone B — Tomatoes',   'Every 2 Days', 0.5, '2025-06-28', '', 'Bi-weekly',    'Consistent moisture prevents blossom end rot'],
  ['Bell Pepper',       'Zone B — Tomatoes',   'Every 3 Days', 0.3, '2025-06-27', '', 'Monthly',      'Water at base; avoid wetting leaves'],
  ['Cucumber',          'Zone E — Cucurbits',  'Daily',        0.5, '2025-06-29', '', 'Weekly',       'Critical to keep soil moist during fruiting'],
  ['Zucchini',          'Zone E — Cucurbits',  'Every 2 Days', 0.5, '2025-06-28', '', 'Bi-weekly',   'Water deeply; avoid leaf wetness (powdery mildew)'],
  ['Basil',             'Zone A — Herbs',      'Every 2 Days', 0.2, '2025-06-28', '', 'Monthly',     'Needs consistent moisture; wilts quickly'],
  ['Lettuce (Romaine)', 'Zone D — Greens',     'Daily',        0.3, '2025-06-29', '', 'Bi-weekly',   'Shallow roots; water frequently in heat'],
  ['Carrot',            'Zone C — Root Veg',   'Every 3 Days', 0.4, '2025-06-27', '', 'Monthly',     'Deep water to encourage root development'],
  ['Garlic',            'Zone C — Root Veg',   'Weekly',       0.3, '2025-06-25', '', 'Once at planting','Reduce water as tops begin to fall over'],
  ['Kale',              'Zone D — Greens',     'Every 3 Days', 0.3, '2025-06-27', '', 'Monthly',     'Heat stress increases bitterness; keep moist'],
  ['Marigold',          'Zone F — Flowers',    'Twice a Week', 0.2, '2025-06-27', '', 'Monthly',     'Drought tolerant once established'],
  ['Pumpkin',           'Zone E — Cucurbits',  'Every 2 Days', 1.0, '2025-06-28', '', 'Bi-weekly',   'High water needs; use drip irrigation'],
];

// Days between watering per frequency
const FREQ_DAYS = {
  'Daily': 1, 'Every 2 Days': 2, 'Every 3 Days': 3,
  'Twice a Week': 3, 'Weekly': 7, 'Bi-weekly': 14,
};

(async () => {
  const reqs = [];
  const values = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(WAT, 0, 1, 0, 14), mergeType: 'MERGE_ALL' } });
  values.push({ range: 'Watering & Care Schedule!A1', values: [['💧 WATERING & CARE SCHEDULE — Season 2025']] });
  reqs.push({
    repeatCell: {
      range: gridRange(WAT, 0, 1, 0, 14),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex('#4A90D9'),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: WAT, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 44 }, fields: 'pixelSize',
  }});

  // ── Header row ────────────────────────────────────────────────────────────
  const headers = [
    'Plant Name', 'Zone', 'Frequency', 'Water (gal)', 'Last Watered',
    'Next Due', 'Days Until Due', 'Fertilize Schedule', 'Last Fertilized',
    'Notes',
  ];
  values.push({ range: 'Watering & Care Schedule!A2', values: [headers] });
  reqs.push({
    repeatCell: {
      range: gridRange(WAT, 1, 2, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex('#2E6DA4'),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: WAT, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});

  // ── Data rows ─────────────────────────────────────────────────────────────
  for (let i = 0; i < WATERING_DATA.length; i++) {
    const [plant, zone, freq, amount, lastWatered, , fertilize, notes] = WATERING_DATA[i];
    const rowNum = i + 3;
    const days = FREQ_DAYS[freq] || 3;

    values.push({ range: `Watering & Care Schedule!A${rowNum}`, values: [[
      plant, zone, freq, amount, lastWatered,
      `=IFERROR(IF(E${rowNum}<>"",E${rowNum}+${days},""),"")`,
      `=IFERROR(IF(F${rowNum}<>"",F${rowNum}-TODAY(),""),"")`,
      fertilize, '', notes,
    ]]});
  }

  // ── Format data rows ──────────────────────────────────────────────────────
  for (let i = 0; i < WATERING_DATA.length; i++) {
    const bg = i % 2 === 0 ? '#EBF5FB' : '#D6EEF8';
    reqs.push({
      repeatCell: {
        range: gridRange(WAT, 2 + i, 3 + i, 0, 10),
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

  // Bold plant names
  reqs.push({
    repeatCell: {
      range: gridRange(WAT, 2, 2 + WATERING_DATA.length, 0, 1),
      cell: { userEnteredFormat: { textFormat: { bold: true } } },
      fields: 'userEnteredFormat.textFormat.bold',
    },
  });

  // Date format cols E, F (4, 5)
  [4, 5].forEach(col => {
    reqs.push({
      repeatCell: {
        range: gridRange(WAT, 2, 2 + WATERING_DATA.length, col, col + 1),
        cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'MMM d, yyyy' } } },
        fields: 'userEnteredFormat.numberFormat',
      },
    });
  });

  // Center align numeric/date cols
  reqs.push({
    repeatCell: {
      range: gridRange(WAT, 2, 2 + WATERING_DATA.length, 2, 9),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat.horizontalAlignment',
    },
  });

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: WAT, dimension: 'ROWS', startIndex: 2, endIndex: 2 + WATERING_DATA.length },
    properties: { pixelSize: 26 }, fields: 'pixelSize',
  }});

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [150, 150, 110, 90, 110, 110, 90, 130, 110, 200];
  colWidths.forEach((px, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: WAT, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  // ── Borders ───────────────────────────────────────────────────────────────
  reqs.push({
    updateBorders: {
      range: gridRange(WAT, 1, 2 + WATERING_DATA.length, 0, 10),
      innerHorizontal: { style: 'SOLID', color: hex('#B3D7ED') },
      innerVertical: { style: 'SOLID', color: hex('#B3D7ED') },
      bottom: { style: 'SOLID_MEDIUM', color: hex('#2E6DA4') },
      top: { style: 'SOLID_MEDIUM', color: hex('#2E6DA4') },
      left: { style: 'SOLID_MEDIUM', color: hex('#2E6DA4') },
      right: { style: 'SOLID_MEDIUM', color: hex('#2E6DA4') },
    },
  });

  // ── Upcoming Care Tasks block ─────────────────────────────────────────────
  const blockRow = 2 + WATERING_DATA.length + 2;
  reqs.push({ mergeCells: { range: gridRange(WAT, blockRow, blockRow + 1, 0, 10), mergeType: 'MERGE_ALL' } });
  values.push({ range: `Watering & Care Schedule!A${blockRow + 1}`, values: [['🔔 UPCOMING CARE — Plants due for watering in next 3 days']] });
  reqs.push({
    repeatCell: {
      range: gridRange(WAT, blockRow, blockRow + 1, 0, 10),
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

  values.push({ range: `Watering & Care Schedule!A${blockRow + 2}`, values: [['Plant']] });
  values.push({ range: `Watering & Care Schedule!B${blockRow + 2}`, values: [['Next Due']] });
  values.push({ range: `Watering & Care Schedule!C${blockRow + 2}`, values: [['Days Until']] });
  reqs.push({
    repeatCell: {
      range: gridRange(WAT, blockRow + 1, blockRow + 2, 0, 3),
      cell: { userEnteredFormat: { backgroundColor: hex(C.paleSage), textFormat: { bold: true, fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // FILTER formula for plants due in ≤3 days
  values.push({ range: `Watering & Care Schedule!A${blockRow + 3}`, values: [
    [`=IFERROR(FILTER(A3:A${2+WATERING_DATA.length},G3:G${2+WATERING_DATA.length}<=3,G3:G${2+WATERING_DATA.length}>=0),"No plants due in 3 days")`]
  ]});
  values.push({ range: `Watering & Care Schedule!B${blockRow + 3}`, values: [
    [`=IFERROR(FILTER(F3:F${2+WATERING_DATA.length},G3:G${2+WATERING_DATA.length}<=3,G3:G${2+WATERING_DATA.length}>=0),"")`]
  ]});
  values.push({ range: `Watering & Care Schedule!C${blockRow + 3}`, values: [
    [`=IFERROR(FILTER(G3:G${2+WATERING_DATA.length},G3:G${2+WATERING_DATA.length}<=3,G3:G${2+WATERING_DATA.length}>=0),"")`]
  ]});

  reqs.push({
    repeatCell: {
      range: gridRange(WAT, blockRow + 2, blockRow + WATERING_DATA.length + 3, 0, 3),
      cell: { userEnteredFormat: { backgroundColor: hex(C.cream2), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // ── Freeze ────────────────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: WAT, gridProperties: { frozenRowCount: 2 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await valuesBatchUpdate(id, values, 'watering-values');
  await batchUpdate(id, reqs, 'watering-format');
  console.log('Watering & Care Schedule complete — 12 plants');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
