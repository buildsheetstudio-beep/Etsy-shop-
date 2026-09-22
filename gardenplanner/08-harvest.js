'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const HARV = sheetMap['Harvest & Yield Tracker'];

// 12 harvest entries
const HARVEST_DATA = [
  // Plant, Date, Amount, Unit, Value ($), Notes
  ['Radish',          '2025-04-14', 1.2,  'lbs',    1.80,  'First harvest of the season! Crisp and spicy'],
  ['Lettuce (Romaine)','2025-04-28', 2.5,  'lbs',    5.00,  'Cut outer leaves; plant still producing'],
  ['Spinach',         '2025-05-03', 1.8,  'lbs',    5.40,  'Before bolting; sweet and tender'],
  ['Lettuce (Romaine)','2025-05-15', 3.0,  'lbs',    6.00,  'Second cut; succession harvest'],
  ['Kale',            '2025-05-20', 1.5,  'lbs',    3.75,  'Baby kale for salads'],
  ['Garlic (Scapes)',  '2025-06-01', 0.8,  'bunch',  3.20,  'Cut scapes to redirect energy to bulb'],
  ['Basil',           '2025-06-10', 0.5,  'bunch',  2.50,  'First pinch; bushing out nicely'],
  ['Tomato (Cherry)', '2025-06-20', 2.0,  'lbs',    6.00,  'First tomatoes! Sweet Sun Gold variety'],
  ['Cucumber',        '2025-06-22', 3.5,  'lbs',    5.25,  'Pickling cucumbers; first flush'],
  ['Zucchini',        '2025-06-25', 4.0,  'lbs',    4.00,  'Harvest when 6-8" for best flavor'],
  ['Tomato (Beefsteak)','2025-06-28', 5.0, 'lbs',   10.00, 'Huge first beefsteak — 1.2 lbs each!'],
  ['Peas (Sugar Snap)', '2025-06-29', 1.5, 'lbs',   4.50, 'Sweet and crunchy; kids love these'],
];

(async () => {
  const reqs = [];
  const values = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(HARV, 0, 1, 0, 14), mergeType: 'MERGE_ALL' } });
  values.push({ range: 'Harvest & Yield Tracker!A1', values: [['🌾 HARVEST & YIELD TRACKER — Season 2025']] });
  reqs.push({
    repeatCell: {
      range: gridRange(HARV, 0, 1, 0, 14),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.wheat),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: HARV, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 44 }, fields: 'pixelSize',
  }});

  // ── Header row ────────────────────────────────────────────────────────────
  const headers = ['Plant Name', 'Harvest Date', 'Amount', 'Unit', 'Est. Value ($)', 'Notes'];
  values.push({ range: 'Harvest & Yield Tracker!A2', values: [headers] });
  reqs.push({
    repeatCell: {
      range: gridRange(HARV, 1, 2, 0, 6),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.harvestGold),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: HARV, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});

  // ── Data rows ─────────────────────────────────────────────────────────────
  values.push({ range: 'Harvest & Yield Tracker!A3', values: HARVEST_DATA });

  // ── Format data rows ──────────────────────────────────────────────────────
  for (let i = 0; i < HARVEST_DATA.length; i++) {
    const bg = i % 2 === 0 ? C.cream2 : C.paleWheat;
    reqs.push({
      repeatCell: {
        range: gridRange(HARV, 2 + i, 3 + i, 0, 6),
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

  // Bold plant name, date format, currency
  reqs.push({
    repeatCell: {
      range: gridRange(HARV, 2, 2 + HARVEST_DATA.length, 0, 1),
      cell: { userEnteredFormat: { textFormat: { bold: true } } },
      fields: 'userEnteredFormat.textFormat.bold',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(HARV, 2, 2 + HARVEST_DATA.length, 1, 2),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'MMM d, yyyy' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(HARV, 2, 2 + HARVEST_DATA.length, 4, 5),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(HARV, 2, 2 + HARVEST_DATA.length, 1, 5),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat.horizontalAlignment',
    },
  });

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: HARV, dimension: 'ROWS', startIndex: 2, endIndex: 2 + HARVEST_DATA.length },
    properties: { pixelSize: 26 }, fields: 'pixelSize',
  }});

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [160, 120, 80, 70, 110, 250];
  colWidths.forEach((px, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: HARV, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  // ── Borders ───────────────────────────────────────────────────────────────
  reqs.push({
    updateBorders: {
      range: gridRange(HARV, 1, 2 + HARVEST_DATA.length, 0, 6),
      innerHorizontal: { style: 'SOLID', color: hex(C.paleWheat) },
      innerVertical: { style: 'SOLID', color: hex(C.paleWheat) },
      bottom: { style: 'SOLID_MEDIUM', color: hex(C.wheat) },
      top: { style: 'SOLID_MEDIUM', color: hex(C.wheat) },
      left: { style: 'SOLID_MEDIUM', color: hex(C.wheat) },
      right: { style: 'SOLID_MEDIUM', color: hex(C.wheat) },
    },
  });

  // ── Summary block ─────────────────────────────────────────────────────────
  const sumRow = 2 + HARVEST_DATA.length + 2;
  reqs.push({ mergeCells: { range: gridRange(HARV, sumRow, sumRow + 1, 0, 6), mergeType: 'MERGE_ALL' } });
  values.push({ range: `Harvest & Yield Tracker!A${sumRow + 1}`, values: [['📊 HARVEST SUMMARY']] });
  reqs.push({
    repeatCell: {
      range: gridRange(HARV, sumRow, sumRow + 1, 0, 6),
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

  const lastDataRow = 2 + HARVEST_DATA.length;
  values.push({ range: `Harvest & Yield Tracker!A${sumRow + 2}`, values: [[
    'Total Harvests:', `=COUNTA(A3:A${lastDataRow})`,
    'Total Yield (lbs):', `=SUMIF(D3:D${lastDataRow},"lbs",C3:C${lastDataRow})`,
    'Total Value:', `=SUM(E3:E${lastDataRow})`,
  ]]});
  values.push({ range: `Harvest & Yield Tracker!A${sumRow + 3}`, values: [[
    'Most Harvested:', `=IFERROR(INDEX(A3:A${lastDataRow},MATCH(MAX(C3:C${lastDataRow}),C3:C${lastDataRow},0)),"")`,
    'Avg per Harvest:', `=IFERROR(AVERAGE(C3:C${lastDataRow}),"")`,
    'Highest Value:', `=IFERROR(MAX(E3:E${lastDataRow}),"")`,
  ]]});

  reqs.push({
    repeatCell: {
      range: gridRange(HARV, sumRow + 1, sumRow + 4, 0, 6),
      cell: { userEnteredFormat: { backgroundColor: hex(C.paleSage), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(HARV, sumRow + 1, sumRow + 4, 0, 1),
      cell: { userEnteredFormat: { textFormat: { bold: true } } },
      fields: 'userEnteredFormat.textFormat.bold',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(HARV, sumRow + 1, sumRow + 4, 4, 5),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // ── Monthly sparkline + chart helper table ────────────────────────────────
  const chartRow = sumRow + 5;
  reqs.push({ mergeCells: { range: gridRange(HARV, chartRow, chartRow + 1, 0, 6), mergeType: 'MERGE_ALL' } });
  values.push({ range: `Harvest & Yield Tracker!A${chartRow + 1}`, values: [['📈 MONTHLY HARVEST TOTALS (Chart Helper)']] });
  reqs.push({
    repeatCell: {
      range: gridRange(HARV, chartRow, chartRow + 1, 0, 6),
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

  values.push({ range: `Harvest & Yield Tracker!A${chartRow + 2}`, values: [['Month', 'Harvests', 'Total lbs', 'Total Value']] });
  reqs.push({
    repeatCell: {
      range: gridRange(HARV, chartRow + 1, chartRow + 2, 0, 4),
      cell: { userEnteredFormat: { backgroundColor: hex(C.olive), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  const months = ['April', 'May', 'June', 'July', 'August', 'September'];
  const monthNums = [4, 5, 6, 7, 8, 9];
  const monthRows = months.map((m, i) => [
    m,
    `=COUNTIFS(MONTH(B3:B${lastDataRow}),${monthNums[i]})`,
    `=IFERROR(SUMPRODUCT((MONTH(B3:B${lastDataRow})=${monthNums[i]})*(D3:D${lastDataRow}="lbs")*C3:C${lastDataRow}),0)`,
    `=IFERROR(SUMPRODUCT((MONTH(B3:B${lastDataRow})=${monthNums[i]})*E3:E${lastDataRow}),0)`,
  ]);
  values.push({ range: `Harvest & Yield Tracker!A${chartRow + 3}`, values: monthRows });

  reqs.push({
    repeatCell: {
      range: gridRange(HARV, chartRow + 2, chartRow + 2 + months.length, 0, 4),
      cell: { userEnteredFormat: { backgroundColor: hex(C.cream2), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // ── Freeze ────────────────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: HARV, gridProperties: { frozenRowCount: 2 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await valuesBatchUpdate(id, values, 'harvest-values');
  await batchUpdate(id, reqs, 'harvest-format');
  console.log('Harvest & Yield Tracker complete — 12 entries');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
