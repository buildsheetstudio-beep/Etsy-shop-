'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CHK = sheetMap['✅ Doc Checklist'];

// Monthly Documentation Checklist
// Cols: A=Month | B=Receipts | C=Mileage Log | D=Bank/Card Stmts | E=Invoices Issued | F=Home Office Photos | G=Prior Year Return | H=% Complete

const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const DOC_ITEMS = ['Receipts','Mileage Log','Bank/Card Stmts','Invoices Issued','Home Office Docs','Prior Year Return'];
const HEADERS = ['Month', ...DOC_ITEMS, '% Complete'];

// Sample: varying monthly completion
const sampleChecks = [
  [true,  true,  true,  true,  false, false], // Jan - 4/6 = 67%
  [true,  true,  true,  false, false, false], // Feb - 3/6 = 50%
  [true,  true,  true,  true,  true,  false], // Mar - 5/6 = 83%
  [true,  true,  true,  true,  false, false], // Apr - 4/6 = 67%
  [true,  false, true,  true,  false, false], // May - 3/6 = 50%
  [true,  true,  true,  true,  true,  false], // Jun - 5/6 = 83%
  [true,  true,  true,  false, false, false], // Jul - 3/6 = 50%
  [true,  true,  true,  true,  true,  false], // Aug - 5/6 = 83%
  [true,  true,  true,  true,  false, false], // Sep - 4/6 = 67%
  [true,  true,  true,  true,  true,  false], // Oct - 5/6 = 83%
  [true,  true,  true,  true,  true,  true],  // Nov - 6/6 = 100%
  [false, false, false, false, false, false],  // Dec - 0/6 = 0%
];

(async () => {
  const reqs = [];

  // Row 0: Title
  reqs.push({ mergeCells: { range: gridRange(CHK, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(CHK, 0, 1, 0, 8),
      cell: {
        userEnteredValue: { stringValue: '✅ Monthly Documentation Checklist' },
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
    range: { sheetId: CHK, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Row 1: subtitle
  reqs.push({ mergeCells: { range: gridRange(CHK, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(CHK, 1, 2, 0, 8),
      cell: {
        userEnteredValue: { stringValue: 'Good record-keeping is your best defense in an audit. Check each item as you gather documentation each month.' },
        userEnteredFormat: {
          backgroundColor: hex(C.lightSage),
          textFormat: { foregroundColor: hex(C.charcoal), italic: true, fontSize: 9 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CHK, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 26 }, fields: 'pixelSize',
  }});

  // Row 2: Headers
  reqs.push({
    repeatCell: {
      range: gridRange(CHK, 2, 3, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.mutedSage),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CHK, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});

  // Data rows 3-14 (12 months)
  for (let r = 3; r < 15; r++) {
    const bg = r % 2 === 1 ? C.inputBg : C.altRow;
    reqs.push({
      repeatCell: {
        range: gridRange(CHK, r, r+1, 0, 8),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: CHK, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
  }

  // % Complete column: formula bg
  reqs.push({ repeatCell: {
    range: gridRange(CHK, 3, 15, 7, 8),
    cell: {
      userEnteredFormat: {
        backgroundColor: hex(C.formulaBg),
        numberFormat: { type: 'PERCENT', pattern: '0%' },
        horizontalAlignment: 'CENTER',
      },
    },
    fields: 'userEnteredFormat',
  }});

  // Checkbox cols (B-G = 1-6) center-aligned
  reqs.push({ repeatCell: {
    range: gridRange(CHK, 3, 15, 1, 7),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});

  // Totals row 15
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CHK, dimension: 'ROWS', startIndex: 15, endIndex: 16 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  reqs.push({
    repeatCell: {
      range: gridRange(CHK, 15, 16, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.paleGold),
          textFormat: { bold: true },
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat',
    },
  });

  // Col widths: A=100, B-G=100 each, H=90
  const colWidths = [110, 95, 95, 120, 110, 130, 120, 90];
  for (const [i, px] of colWidths.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: CHK, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  await batchUpdate(id, reqs, 'checklist-format');

  // Values
  const data = [];
  data.push({ range: `'✅ Doc Checklist'!A3:H3`, values: [HEADERS] });

  for (let m = 0; m < 12; m++) {
    const row = m + 4;
    const checks = sampleChecks[m];
    const pctFormula = `=IFERROR(COUNTIF(B${row}:G${row},TRUE)/6,0)`;
    data.push({
      range: `'✅ Doc Checklist'!A${row}:H${row}`,
      values: [[MONTHS[m], ...checks, pctFormula]],
    });
  }

  // Summary row 16
  const summaryFormulas = ['ANNUAL TOTALS'];
  for (let col = 1; col <= 6; col++) {
    const colLetter = String.fromCharCode(65 + col);
    summaryFormulas.push(`=COUNTIF(${colLetter}4:${colLetter}15,TRUE)/12`);
  }
  summaryFormulas.push(`=IFERROR(AVERAGE(H4:H15),0)`);
  data.push({ range: `'✅ Doc Checklist'!A16:H16`, values: [summaryFormulas] });

  await valuesBatchUpdate(id, data, 'checklist-values');

  // Format summary % cols
  const pctReqs = [];
  for (let col = 1; col <= 7; col++) {
    pctReqs.push({ repeatCell: {
      range: gridRange(CHK, 15, 16, col, col+1),
      cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  }
  await batchUpdate(id, pctReqs, 'checklist-summary-fmt');

  console.log('Doc Checklist complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
