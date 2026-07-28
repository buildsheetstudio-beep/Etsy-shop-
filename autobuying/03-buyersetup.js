'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BS = sheetMap['Buyer Setup'];
const S = "'Buyer Setup'";

(async () => {
  const data = [];

  // ── Row 1: Title banner ──────────────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['BUYER SETUP & FINANCIAL PROFILE']] });
  data.push({ range: `${S}!A2`, values: [['Jordan Taylor — Auto Purchase Budget Planner']] });
  data.push({ range: `${S}!A3`, values: [['Complete all yellow input cells. Green cells calculate automatically.']] });

  // ── Section A: Buyer Inputs (col A=label, col B=value) ───────────────────
  data.push({ range: `${S}!A5`, values: [['BUYER PROFILE']] });

  const inputs = [
    ['Buyer Name',                           'Jordan Taylor'],
    ['Currency Symbol',                      '$'],
    ['Monthly Take-Home Pay',                5800],
    ['Other Monthly Income',                  400],
    ['Current Fixed Expenses (rent, utilities, subscriptions)', 2750],
    ['Current Debt Payments (student loans, credit cards)',      450],
    ['Current Transportation Cost (insurance, fuel, transit)',   520],
    ['Monthly Savings Goal',                  700],
    ['Emergency Fund Balance',              18000],
    ['Desired Emergency Fund Minimum',      12000],
    ['Maximum Comfortable Vehicle Payment',   750],
    ['Maximum Total Transportation Budget',  1150],
    ['Planned Down Payment',                 6000],
    ['Trade-In Value',                       8500],
    ['Trade-In Loan Balance',                3000],
    ['Expected Annual Miles',               12000],
    ['Expected Ownership Period (years)',        5],
    ['Preferred Purchase Method',         'Finance'],
    ['Target Purchase Date',              'Aug 2026'],
    ['Notes',                             'Comparing sedan, SUV, and EV options. Prefer reliable brand.'],
  ];

  inputs.forEach(([label, val], i) => {
    data.push({ range: `${S}!A${i+6}:B${i+6}`, values: [[label, val]] });
  });

  // ── Section B: Summary Cards ──────────────────────────────────────────────
  const cardRow = 30;
  data.push({ range: `${S}!A${cardRow}`, values: [['FINANCIAL SUMMARY']] });

  const cardDefs = [
    ['Total Monthly Income',          `=IFERROR(B8+B9,0)`],
    ['Monthly Obligations Before Vehicle', `=IFERROR(B10+B11,0)`],
    ['Available Monthly Cash',        `=IFERROR(B${cardRow+1}-B${cardRow+2}-B13,0)`],
    ['Maximum Vehicle Budget',        `=IFERROR(MIN(B16,B${cardRow+3}),0)`],
    ['Net Trade-In Equity',           `=IFERROR(B19-B20,0)`],
    ['Cash Remaining After Down Payment', `=IFERROR(B14-B18,0)`],
    ['Emergency Fund After Purchase', `=IFERROR(B14-B18,0)`],
    ['Recommended Payment Ceiling',   `=IFERROR(MIN(B16,B${cardRow+3}),0)`],
  ];

  cardDefs.forEach(([label, formula], i) => {
    data.push({ range: `${S}!A${cardRow+1+i}`, values: [[label]] });
    data.push({ range: `${S}!B${cardRow+1+i}`, values: [[formula]] });
  });

  // Affordability note below cards
  const noteRow = cardRow + cardDefs.length + 2;
  data.push({ range: `${S}!A${noteRow}`, values: [['Payment ceiling = lower of your entered maximum and available cash after obligations and savings goal.']] });
  data.push({ range: `${S}!A${noteRow+1}`, values: [['Planning guideline — not a lender requirement. Adjust inputs to reflect your situation.']] });

  // ── Disclaimer ────────────────────────────────────────────────────────────
  const discRow = noteRow + 3;
  data.push({ range: `${S}!A${discRow}`, values: [[
    'This workbook is an educational comparison tool only. It does not provide lending, financial, tax, insurance, legal, or purchasing advice. ' +
    'Vehicle prices, taxes, fees, interest rates, lease terms, fuel costs, insurance premiums, maintenance costs, depreciation, and resale values ' +
    'vary by location and individual circumstances. Replace all estimates with actual quotes and review any purchase or lease agreement carefully ' +
    'before making a decision.'
  ]] });

  await valuesBatchUpdate(id, data, 'buyer-setup-values');

  // ── Formatting ────────────────────────────────────────────────────────────
  const reqs = [];

  // Sheet background
  reqs.push({ repeatCell: { range: gridRange(BS,0,500,0,10), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' }});

  // Row 1 title
  reqs.push({ mergeCells: { range: gridRange(BS,0,1,0,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(BS,0,1,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize: 16 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE'
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:BS, dimension:'ROWS', startIndex:0, endIndex:1 }, properties: { pixelSize:48 }, fields:'pixelSize' }});

  // Row 2 subtitle
  reqs.push({ mergeCells: { range: gridRange(BS,1,2,0,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(BS,1,2,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize: 11 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE'
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Row 3 guidance
  reqs.push({ mergeCells: { range: gridRange(BS,2,3,0,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(BS,2,3,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.mutedBlue), textFormat: { italic:true, foregroundColor: hex(C.mainText), fontSize: 10 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE'
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Section A header (row 5, index 4)
  reqs.push({ mergeCells: { range: gridRange(BS,4,5,0,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(BS,4,5,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize: 11 },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE'
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:BS, dimension:'ROWS', startIndex:4, endIndex:5 }, properties: { pixelSize:30 }, fields:'pixelSize' }});

  // Input labels (col A, rows 6-25, index 5-24)
  reqs.push({ repeatCell: { range: gridRange(BS,5,25,0,1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.white), textFormat: { bold:false, foregroundColor: hex(C.mainText) },
    wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE'
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' }});

  // Input value cells (col B, rows 6-25, index 5-24) — pale yellow
  reqs.push({ repeatCell: { range: gridRange(BS,5,25,1,2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input), textFormat: { foregroundColor: hex(C.mainText) },
    horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE'
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Currency formats for monetary inputs
  const currFmt = { type:'CURRENCY', pattern:'$#,##0.00' };
  const currRows = [2,3,4,5,6,7,8,9,10,11,12,13,14]; // 0-indexed within inputs: take-home=2,other=3...
  // Rows 8,9,10,11,12,13,14,15,16,17,18 (A1 rows, index 7-17 for monetary ones)
  [7,8,9,10,11,12,13,14,18].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(BS,ri,ri+1,1,2), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields: 'userEnteredFormat.numberFormat' }});
  });

  // Section B header
  reqs.push({ mergeCells: { range: gridRange(BS,29,30,0,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(BS,29,30,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:11 },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE'
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:BS, dimension:'ROWS', startIndex:29, endIndex:30 }, properties: { pixelSize:30 }, fields:'pixelSize' }});

  // Summary card labels (col A, rows 31-38, index 30-37)
  reqs.push({ repeatCell: { range: gridRange(BS,30,38,0,1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.white), textFormat: { bold:false, foregroundColor: hex(C.mainText) },
    wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE'
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' }});

  // Summary formula cells (col B, rows 31-38, index 30-37) — pale blue-gray
  reqs.push({ repeatCell: { range: gridRange(BS,30,38,1,2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula), textFormat: { bold:true, foregroundColor: hex(C.primary) },
    horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE', numberFormat: currFmt
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,numberFormat)' }});

  // Net Trade-In equity row — may be negative; highlight negative in rust
  // (conditional formatting added in cf pass)

  // Planning notes (rows 39-40, index 38-39)
  reqs.push({ mergeCells: { range: gridRange(BS,38,39,0,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(BS,38,39,0,8), cell: { userEnteredFormat: {
    textFormat: { italic:true, foregroundColor: hex(C.secText), fontSize:9 }, wrapStrategy:'WRAP'
  }}, fields: 'userEnteredFormat(textFormat,wrapStrategy)' }});
  reqs.push({ mergeCells: { range: gridRange(BS,39,40,0,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(BS,39,40,0,8), cell: { userEnteredFormat: {
    textFormat: { italic:true, foregroundColor: hex(C.secondary), fontSize:9 }, wrapStrategy:'WRAP'
  }}, fields: 'userEnteredFormat(textFormat,wrapStrategy)' }});

  // Disclaimer box
  const discIdx = 41;
  reqs.push({ mergeCells: { range: gridRange(BS,discIdx,discIdx+3,0,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(BS,discIdx,discIdx+3,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex('#E8E6E1'),
    textFormat: { italic:true, foregroundColor: hex(C.secText), fontSize:9 },
    wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE',
    borders: {
      top: { style:'SOLID_MEDIUM', color: hex(C.slate) },
      bottom: { style:'SOLID_MEDIUM', color: hex(C.slate) },
      left: { style:'SOLID_MEDIUM', color: hex(C.slate) },
      right: { style:'SOLID_MEDIUM', color: hex(C.slate) },
    }
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment,borders)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:BS, dimension:'ROWS', startIndex:discIdx, endIndex:discIdx+3 }, properties: { pixelSize:36 }, fields:'pixelSize' }});

  // Column widths
  reqs.push({ updateDimensionProperties: { range: { sheetId:BS, dimension:'COLUMNS', startIndex:0, endIndex:1 }, properties: { pixelSize:340 }, fields:'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:BS, dimension:'COLUMNS', startIndex:1, endIndex:2 }, properties: { pixelSize:160 }, fields:'pixelSize' }});

  // Freeze rows 1:3 (index 0:3)
  reqs.push({ updateSheetProperties: { properties: { sheetId:BS, gridProperties: { frozenRowCount:3 } }, fields:'gridProperties.frozenRowCount' }});

  await batchUpdate(id, reqs, 'buyer-setup-format');
  console.log('✓ Buyer Setup');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
