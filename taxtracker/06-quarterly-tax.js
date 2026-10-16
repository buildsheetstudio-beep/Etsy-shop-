'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const QTR = sheetMap['📅 Quarterly Tax'];

// Cols: A=Quarter | B=Due Date | C=Period | D=Est. Income | E=Tax Rate % | F=SE Tax | G=Income Tax | H=Total Due | I=Amount Paid | J=Paid?
// SE Tax = D * 0.9235 * 0.153 (self-employment tax)
// Income Tax = D * E (effective rate)
// Total = SE + Income Tax

const HEADERS = ['Quarter','Due Date','Income Period','Est. Income','Tax Rate %','SE Tax (15.3%)','Income Tax','Total Estimated','Amount Paid','Paid?'];
const WIDTHS  = [75, 90, 130, 110, 80, 110, 100, 120, 110, 60];

const quarterData = [
  ['Q1 2024', '2024-04-15', 'Jan 1 – Mar 31',  25200, 0.22],
  ['Q2 2024', '2024-06-15', 'Apr 1 – May 31',  24800, 0.22],
  ['Q3 2024', '2024-09-15', 'Jun 1 – Aug 31',  27100, 0.22],
  ['Q4 2024', '2025-01-15', 'Sep 1 – Dec 31',  28400, 0.24],
];

(async () => {
  const reqs = [];

  // Row 0: Title
  reqs.push({ mergeCells: { range: gridRange(QTR, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(QTR, 0, 1, 0, 10),
      cell: {
        userEnteredValue: { stringValue: '📅 Quarterly Estimated Tax Calculator' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepSapphire),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: QTR, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Row 1: subtitle / info bar
  reqs.push({ mergeCells: { range: gridRange(QTR, 1, 2, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(QTR, 1, 2, 0, 10),
      cell: {
        userEnteredValue: { stringValue: 'Self-employed? Pay quarterly taxes to avoid IRS penalties. SE Tax = Net Self-Employment × 92.35% × 15.3%' },
        userEnteredFormat: {
          backgroundColor: hex(C.lightAmber),
          textFormat: { foregroundColor: hex(C.charcoal), italic: true, fontSize: 9 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: QTR, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Row 2: Headers
  reqs.push({
    repeatCell: {
      range: gridRange(QTR, 2, 3, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.amber),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: QTR, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});

  // Data rows 3-6
  for (let r = 3; r < 7; r++) {
    const bg = r % 2 === 1 ? C.inputBg : C.altRow;
    reqs.push({
      repeatCell: {
        range: gridRange(QTR, r, r+1, 0, 10),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: QTR, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 30 }, fields: 'pixelSize',
    }});
  }

  // Formula columns (E, F, G, H) have formula bg
  for (const col of [5, 6, 7]) {
    reqs.push({ repeatCell: {
      range: gridRange(QTR, 3, 7, col, col+1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg) } },
      fields: 'userEnteredFormat',
    }});
  }

  // Totals row 7
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: QTR, dimension: 'ROWS', startIndex: 7, endIndex: 8 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});
  reqs.push({
    repeatCell: {
      range: gridRange(QTR, 7, 8, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.paleGold),
          textFormat: { bold: true, fontSize: 10 },
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat',
    },
  });

  // Summary panel rows 9-15 (A-D wide)
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: QTR, dimension: 'ROWS', startIndex: 8, endIndex: 9 },
    properties: { pixelSize: 14 }, fields: 'pixelSize',
  }});
  reqs.push({ mergeCells: { range: gridRange(QTR, 9, 10, 0, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(QTR, 9, 10, 0, 5),
      cell: {
        userEnteredValue: { stringValue: '📋 Year Summary' },
        userEnteredFormat: {
          backgroundColor: hex(C.charcoal),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: QTR, dimension: 'ROWS', startIndex: 9, endIndex: 10 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  const summaryRows = [
    ['Total Estimated Income', `=IFERROR(SUM(D3:D6),0)`],
    ['Total SE Tax',           `=IFERROR(SUM(F3:F6),0)`],
    ['Total Income Tax',       `=IFERROR(SUM(G3:G6),0)`],
    ['Total Tax Due',          `=IFERROR(SUM(H3:H6),0)`],
    ['Total Amount Paid',      `=IFERROR(SUM(I3:I6),0)`],
    ['Remaining Balance',      `=IFERROR(H10-H14,0)`],
  ];

  for (let r = 0; r < summaryRows.length; r++) {
    const row = r + 10;
    const bg = r % 2 === 0 ? C.lightAmber : C.inputBg;
    reqs.push({
      repeatCell: {
        range: gridRange(QTR, row, row+1, 0, 5),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: QTR, dimension: 'ROWS', startIndex: row, endIndex: row+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
  }

  // Col widths
  for (const [i, px] of WIDTHS.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: QTR, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  // Number formats
  reqs.push({ repeatCell: {
    range: gridRange(QTR, 3, 7, 1, 2),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(QTR, 3, 7, 3, 4),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(QTR, 3, 7, 4, 5),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  for (const col of [5, 6, 7, 8]) {
    reqs.push({ repeatCell: {
      range: gridRange(QTR, 3, 8, col, col+1),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  }
  reqs.push({ repeatCell: {
    range: gridRange(QTR, 3, 7, 9, 10),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  // Summary value column formats
  for (let r = 10; r < 16; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(QTR, r, r+1, 1, 2),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  }

  await batchUpdate(id, reqs, 'quarterly-format');

  // Values
  const data = [];
  data.push({ range: `'📅 Quarterly Tax'!A3:J3`, values: [HEADERS] });

  for (let i = 0; i < quarterData.length; i++) {
    const row = i + 4;
    const [quarter, dueDate, period, income, rate] = quarterData[i];
    const seFormula = `=IFERROR(D${row}*0.9235*0.153,0)`;
    const incomeFormula = `=IFERROR(D${row}*E${row},0)`;
    const totalFormula = `=IFERROR(F${row}+G${row},0)`;
    data.push({
      range: `'📅 Quarterly Tax'!A${row}:J${row}`,
      values: [[quarter, dueDate, period, income, rate, seFormula, incomeFormula, totalFormula, '', false]],
    });
  }

  // Totals row 8
  data.push({
    range: `'📅 Quarterly Tax'!A8:J8`,
    values: [['ANNUAL TOTALS', '', '', '=SUM(D4:D7)', '', '=SUM(F4:F7)', '=SUM(G4:G7)', '=SUM(H4:H7)', '=SUM(I4:I7)', '']],
  });

  // Summary labels and formulas (rows 10-15 = sheet rows 11-16)
  for (const [r, [label, formula]] of summaryRows.entries()) {
    const row = r + 11;
    data.push({ range: `'📅 Quarterly Tax'!A${row}:B${row}`, values: [[label, formula]] });
  }

  await valuesBatchUpdate(id, data, 'quarterly-values');
  console.log('Quarterly Tax complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
