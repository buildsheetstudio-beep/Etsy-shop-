'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const HOM = sheetMap['🏠 Home Office'];

// Home Office Calculator: Simplified vs Actual method comparison
// Sample: 200 sq ft office / 1,800 sq ft home
// Simplified: $5/sqft * 200 = $1,000
// Actual: home expenses * (200/1800) = ~$2,866

(async () => {
  const reqs = [];

  // Row 0: Title
  reqs.push({ mergeCells: { range: gridRange(HOM, 0, 1, 0, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(HOM, 0, 1, 0, 5),
      cell: {
        userEnteredValue: { stringValue: '🏠 Home Office Deduction Calculator' },
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
    range: { sheetId: HOM, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Section A: Inputs (rows 1-8)
  const sectionHeader = (r, text, bg) => ({
    repeatCell: {
      range: gridRange(HOM, r, r+1, 0, 5),
      cell: {
        userEnteredValue: { stringValue: text },
        userEnteredFormat: {
          backgroundColor: hex(bg),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  reqs.push(sectionHeader(1, '📐 Your Home Office Information', C.charcoal));
  reqs.push({ mergeCells: { range: gridRange(HOM, 1, 2, 0, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: HOM, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});

  // Input rows 2-7: label | value | unit | note
  const inputData = [
    ['Office Square Footage',    200,     'sq ft', 'Dedicated office space only'],
    ['Total Home Square Footage',1800,    'sq ft', 'Total home/apt area'],
    ['Annual Rent or Mortgage',  18000,   '$',     'Total paid in tax year'],
    ['Annual Utilities',         3600,    '$',     'Electric, gas, water, internet'],
    ['Annual Home Insurance',    1200,    '$',     'Homeowners or renters insurance'],
    ['Repairs & Maintenance',    600,     '$',     'Home repairs allocated to office'],
  ];

  for (let r = 0; r < inputData.length; r++) {
    const row = r + 2;
    const bg = r % 2 === 0 ? C.inputBg : C.altRow;
    reqs.push({
      repeatCell: {
        range: gridRange(HOM, row, row+1, 0, 5),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: HOM, dimension: 'ROWS', startIndex: row, endIndex: row+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
  }

  // Spacer row 8
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: HOM, dimension: 'ROWS', startIndex: 8, endIndex: 9 },
    properties: { pixelSize: 16 }, fields: 'pixelSize',
  }});

  // Section B: Method comparison header (row 9)
  reqs.push(sectionHeader(9, '📊 Method Comparison', C.charcoal));
  reqs.push({ mergeCells: { range: gridRange(HOM, 9, 10, 0, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: HOM, dimension: 'ROWS', startIndex: 9, endIndex: 10 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});

  // Column headers row 10
  reqs.push({
    repeatCell: {
      range: gridRange(HOM, 10, 11, 0, 5),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.warmGold),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat',
    },
  });

  // Simplified method block rows 11-15
  const simplifiedRows = [
    ['SIMPLIFIED METHOD', '', '', '', ''],
    ['IRS Rate per Sq Ft', '$5.00', '', 'Fixed IRS rate (2024)', ''],
    ['Max Eligible Sq Ft', '=MIN(B2,300)', 'sq ft', 'IRS cap: 300 sq ft', ''],
    ['Deduction = Rate × Sq Ft', '=5*B12', '', 'Total simplified deduction', ''],
    ['', '', '', '', ''],
  ];

  // Actual method block rows 16-24
  const actualRows = [
    ['ACTUAL METHOD', '', '', '', ''],
    ['Business Use %', '=IFERROR(B2/B3,0)', '', '=Office÷Total Home', ''],
    ['Deductible Rent/Mortgage', '=IFERROR(B4*B17,0)', '', '=Rent × Business%', ''],
    ['Deductible Utilities', '=IFERROR(B5*B17,0)', '', '=Utilities × Business%', ''],
    ['Deductible Insurance', '=IFERROR(B6*B17,0)', '', '=Insurance × Business%', ''],
    ['Deductible Repairs', '=IFERROR(B7*B17,0)', '', '=Repairs × Business%', ''],
    ['Total Actual Deduction', '=IFERROR(B18+B19+B20+B21,0)', '', 'Sum of all actual costs', ''],
    ['', '', '', '', ''],
  ];

  // Result block rows 24-26
  const resultRows = [
    ['🏆 RECOMMENDED METHOD', '=IF(B14>B23,"Simplified","Actual")', '', '', ''],
    ['💰 MAXIMUM DEDUCTION', '=MAX(B14,B23)', '', 'Higher of the two methods', ''],
    ['📊 You Save vs Other Method', '=ABS(B14-B23)', '', 'Additional savings by choosing right', ''],
  ];

  for (let r = 11; r < 11 + simplifiedRows.length + actualRows.length + resultRows.length; r++) {
    const bg = r % 2 === 0 ? C.inputBg : C.altRow;
    reqs.push({
      repeatCell: {
        range: gridRange(HOM, r, r+1, 0, 5),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: HOM, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
  }

  // Section titles bold in col A
  for (const r of [11, 16]) {
    reqs.push({
      repeatCell: {
        range: gridRange(HOM, r, r+1, 0, 1),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.deepSapphire),
            textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          },
        },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ mergeCells: { range: gridRange(HOM, r, r+1, 0, 5), mergeType: 'MERGE_ALL' } });
  }

  // Result rows: gold bg
  for (let r = 24; r < 27; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(HOM, r, r+1, 0, 5),
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
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: HOM, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 32 }, fields: 'pixelSize',
    }});
  }

  // Col widths: A=240, B=140, C=70, D=230, E=100
  for (const [i, px] of [[0,240],[1,140],[2,70],[3,230],[4,100]]) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: HOM, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  // Currency format for B col in expense rows
  for (const r of [4,5,6,7]) {
    reqs.push({ repeatCell: {
      range: gridRange(HOM, r, r+1, 1, 2),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  }
  // Business Use % format
  reqs.push({ repeatCell: {
    range: gridRange(HOM, 17, 18, 1, 2),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.00%' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  // Currency for deductibles and results
  for (const r of [13, 18, 19, 20, 21, 22, 25, 26]) {
    reqs.push({ repeatCell: {
      range: gridRange(HOM, r, r+1, 1, 2),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  }

  await batchUpdate(id, reqs, 'home-format');

  // Values
  const data = [];

  // Column headers
  data.push({ range: `'🏠 Home Office'!A11:E11`, values: [['Item', 'Value', 'Unit', 'Notes', '']] });

  // Input labels + values
  for (let r = 0; r < inputData.length; r++) {
    const row = r + 3;
    const [label, value, unit, note] = inputData[r];
    data.push({ range: `'🏠 Home Office'!A${row}:E${row}`, values: [[label, value, unit, note, '']] });
  }

  // Simplified rows (starting row 12)
  for (let r = 0; r < simplifiedRows.length; r++) {
    data.push({ range: `'🏠 Home Office'!A${r+12}:E${r+12}`, values: [simplifiedRows[r]] });
  }

  // Actual rows (starting row 17)
  for (let r = 0; r < actualRows.length; r++) {
    data.push({ range: `'🏠 Home Office'!A${r+17}:E${r+17}`, values: [actualRows[r]] });
  }

  // Result rows (starting row 25)
  for (let r = 0; r < resultRows.length; r++) {
    data.push({ range: `'🏠 Home Office'!A${r+25}:E${r+25}`, values: [resultRows[r]] });
  }

  await valuesBatchUpdate(id, data, 'home-values');
  console.log('Home Office complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
