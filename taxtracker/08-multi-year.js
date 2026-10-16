'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const MYR = sheetMap['📈 Multi-Year'];

// Multi-Year Comparison tab
// Cols: A=Year | B=Total Deductions | C=Meals (50%) | D=Mileage | E=Home Office | F=Other | G=Documentation %

const HEADERS = ['Tax Year','Total Deductions','Meals & Ent (Net)','Mileage Deductible','Home Office','Other Deductions','Doc Rate %'];
const WIDTHS  = [80, 130, 130, 130, 110, 130, 90];

const years = [2024, 2025, 2026];

// SUMPRODUCT formulas from Deduction Log filtered by year
// Log is '📝 Deduction Log', data A3:H200
// A=Date, B=Category, G=Net Deductible, F=Doc captured
const LOG = "'📝 Deduction Log'";

function yearFilter(year, col) {
  return `(YEAR(${LOG}!$A$3:$A$200)=${year})`;
}

function sumForYear(year, colLetter) {
  return `=IFERROR(SUMPRODUCT(${yearFilter(year)}*(${LOG}!$${colLetter}$3:$${colLetter}$200)),0)`;
}

function mealsForYear(year) {
  // Net deductible for Meals & Entertainment specifically
  return `=IFERROR(SUMPRODUCT(${yearFilter(year)}*(${LOG}!$B$3:$B$200="Meals & Entertainment")*(${LOG}!$G$3:$G$200)),0)`;
}

function mileageForYear(year) {
  // Pull from Mileage Log '🚗 Mileage Log'
  const MIL = "'🚗 Mileage Log'";
  return `=IFERROR(SUMPRODUCT((YEAR(${MIL}!$A$3:$A$200)=${year})*(${MIL}!$F$3:$F$200)),0)`;
}

function homeOfficeForYear(year) {
  return `=IFERROR(SUMPRODUCT(${yearFilter(year)}*(${LOG}!$B$3:$B$200="Home Office")*(${LOG}!$G$3:$G$200)),0)`;
}

function otherForYear(year) {
  // All net deductible except Meals & Entertainment and Home Office
  return `=IFERROR(SUMPRODUCT(${yearFilter(year)}*(${LOG}!$B$3:$B$200<>"Meals & Entertainment")*(${LOG}!$B$3:$B$200<>"Home Office")*(${LOG}!$G$3:$G$200)),0)`;
}

function docRateForYear(year) {
  // % of rows that have F=TRUE (documented) for the year
  return `=IFERROR(SUMPRODUCT(${yearFilter(year)}*(${LOG}!$F$3:$F$200=TRUE))/SUMPRODUCT(${yearFilter(year)}*(${LOG}!$A$3:$A$200<>"")),0)`;
}

(async () => {
  const reqs = [];

  // Row 0: Title
  reqs.push({ mergeCells: { range: gridRange(MYR, 0, 1, 0, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(MYR, 0, 1, 0, 7),
      cell: {
        userEnteredValue: { stringValue: '📈 Multi-Year Tax Comparison' },
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
    range: { sheetId: MYR, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Row 1: Headers
  reqs.push({
    repeatCell: {
      range: gridRange(MYR, 1, 2, 0, 7),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepSapphire),
          textFormat: { foregroundColor: hex(C.warmGold), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: MYR, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});

  // Data rows 2-4 (3 years)
  for (let r = 2; r < 5; r++) {
    const bg = r % 2 === 0 ? C.inputBg : C.altRow;
    reqs.push({
      repeatCell: {
        range: gridRange(MYR, r, r+1, 0, 7),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: MYR, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 30 }, fields: 'pixelSize',
    }});
  }

  // Formula cols: B-F currency, G percent
  reqs.push({ repeatCell: {
    range: gridRange(MYR, 2, 5, 1, 6),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' }, backgroundColor: hex(C.formulaBg) } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(MYR, 2, 5, 6, 7),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' }, backgroundColor: hex(C.formulaBg), horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(MYR, 2, 5, 0, 1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true } } },
    fields: 'userEnteredFormat',
  }});

  // Growth row 5
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: MYR, dimension: 'ROWS', startIndex: 5, endIndex: 6 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  reqs.push({
    repeatCell: {
      range: gridRange(MYR, 5, 6, 0, 7),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.lightSage),
          textFormat: { italic: true, fontSize: 9 },
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat',
    },
  });

  // Summary panel rows 7-11 (below data)
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: MYR, dimension: 'ROWS', startIndex: 6, endIndex: 7 },
    properties: { pixelSize: 14 }, fields: 'pixelSize',
  }});
  reqs.push({ mergeCells: { range: gridRange(MYR, 7, 8, 0, 4), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(MYR, 7, 8, 0, 4),
      cell: {
        userEnteredValue: { stringValue: '📊 3-Year Totals' },
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
    range: { sheetId: MYR, dimension: 'ROWS', startIndex: 7, endIndex: 8 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  const summaryItems = [
    ['Total All-Year Deductions', '=IFERROR(SUM(B3:B5),0)'],
    ['Best Year',                  '=IFERROR(INDEX(A3:A5,MATCH(MAX(B3:B5),B3:B5,0)),"")'],
    ['Avg Annual Deductions',      '=IFERROR(AVERAGE(B3:B5),0)'],
    ['Avg Documentation Rate',     '=IFERROR(AVERAGE(G3:G5),0)'],
  ];

  for (let r = 0; r < summaryItems.length; r++) {
    const row = r + 8;
    const bg = r % 2 === 0 ? C.lightAmber : C.inputBg;
    reqs.push({
      repeatCell: {
        range: gridRange(MYR, row, row+1, 0, 4),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: MYR, dimension: 'ROWS', startIndex: row, endIndex: row+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
  }

  // Col widths
  for (const [i, px] of WIDTHS.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: MYR, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  await batchUpdate(id, reqs, 'multiyear-format');

  // Values
  const data = [];
  data.push({ range: `'📈 Multi-Year'!A2:G2`, values: [HEADERS] });

  for (let i = 0; i < years.length; i++) {
    const year = years[i];
    const row = i + 3;
    data.push({
      range: `'📈 Multi-Year'!A${row}:G${row}`,
      values: [[
        year,
        `=IFERROR(SUMPRODUCT((YEAR('📝 Deduction Log'!$A$3:$A$200)=${year})*('📝 Deduction Log'!$G$3:$G$200)),0)`,
        mealsForYear(year),
        mileageForYear(year),
        homeOfficeForYear(year),
        otherForYear(year),
        docRateForYear(year),
      ]],
    });
  }

  // Year-over-year growth row 6
  data.push({
    range: `'📈 Multi-Year'!A6:G6`,
    values: [['YoY Growth', '=IFERROR(B5/B4-1,"")', '', '', '', '', '']],
  });

  // Summary values (rows 9-12)
  for (const [r, [label, formula]] of summaryItems.entries()) {
    const row = r + 9;
    data.push({ range: `'📈 Multi-Year'!A${row}:B${row}`, values: [[label, formula]] });
  }

  await valuesBatchUpdate(id, data, 'multiyear-values');

  // Format summary values
  const fmtReqs = [];
  for (let r = 9; r < 13; r++) {
    const pattern = r === 12 ? '0.0%' : (r === 10 ? '0' : '$#,##0.00');
    const type = r === 12 ? 'PERCENT' : (r === 10 ? 'NUMBER' : 'CURRENCY');
    fmtReqs.push({ repeatCell: {
      range: gridRange(MYR, r, r+1, 1, 2),
      cell: { userEnteredFormat: { numberFormat: { type, pattern } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  }
  // YoY row format
  fmtReqs.push({ repeatCell: {
    range: gridRange(MYR, 5, 6, 1, 2),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '+0.0%;-0.0%' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  await batchUpdate(id, fmtReqs, 'multiyear-numfmt');

  console.log('Multi-Year complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
