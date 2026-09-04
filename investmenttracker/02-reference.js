'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Reference Data'];
const S = "'Reference Data'";

const SECTIONS = [
  { label: 'Owners', data: [['Daniel Walsh'],['Emily Walsh'],['Joint Household']] },
  { label: 'Account Types', data: [
    ['Taxable Brokerage'],['Traditional IRA'],['Roth IRA'],['401(k)'],['403(b)'],
    ['457(b)'],['SEP IRA'],['SIMPLE IRA'],['Solo 401(k)'],['HSA'],
    ['Education Account'],['Trust Account'],['Cash Management'],['Crypto Account'],['Other Investment Account'],
  ]},
  { label: 'Tax Treatments', data: [['Taxable'],['Tax-Deferred'],['Tax-Free / Roth'],['Mixed'],['Other']] },
  { label: 'Asset Classes', data: [
    ['U.S. Stocks'],['International Stocks'],['Emerging Markets'],['Bonds'],
    ['Real Estate'],['REITs'],['Cash'],['Money Market'],['Commodities'],
    ['Gold / Precious Metals'],['Cryptocurrency'],['Alternatives'],['Other'],
  ]},
  { label: 'Security Types', data: [
    ['Stock'],['ETF'],['Mutual Fund'],['Bond'],['REIT'],['Money Market Fund'],
    ['CD'],['Cryptocurrency'],['Cash'],['Other'],
  ]},
  { label: 'Transaction Types', data: [
    ['Buy'],['Sell'],['Contribution'],['Withdrawal'],['Dividend'],['Interest'],
    ['Capital Gain Distribution'],['Fee'],['Transfer In'],['Transfer Out'],
    ['Stock Split'],['Reinvestment'],['Tax Withholding'],['Adjustment'],
  ]},
  { label: 'Dividend Frequencies', data: [
    ['Monthly'],['Quarterly'],['Semiannual'],['Annual'],['Irregular'],['None'],
  ]},
  { label: 'Goal Types', data: [
    ['Portfolio Value'],['Annual Contribution'],['Monthly Contribution'],['Net Worth'],
    ['Dividend Income'],['Emergency Savings'],['Retirement Account Balance'],
    ['Asset Allocation'],['Financial Independence'],['Education Funding'],
    ['Home Purchase'],['Other'],
  ]},
  { label: 'Goal Statuses', data: [
    ['Not Started'],['In Progress'],['Achieved'],['Delayed'],['Reassess'],
  ]},
  { label: 'Watchlist Statuses', data: [
    ['Researching'],['Monitoring'],['Waiting for Price'],['Ready for Review'],['Purchased'],['Removed'],
  ]},
  { label: 'Risk Levels', data: [
    ['Low'],['Moderate'],['High'],['Speculative'],['Not Assessed'],
  ]},
  { label: 'Review Categories', data: [
    ['Contributions'],['Performance'],['Allocation'],['Dividends'],['Fees'],
    ['Risk'],['Goals'],['Tax Organization'],['Account Maintenance'],['Other'],
  ]},
  { label: 'Review Statuses', data: [
    ['Not Started'],['In Progress'],['Complete'],['Follow-Up Needed'],
  ]},
  { label: 'Contribution Types', data: [
    ['Recurring'],['One-Time'],['Employer Match'],['Rollover'],['Transfer'],['Other'],
  ]},
  { label: 'Net Worth Asset Categories', data: [
    ['Investment Accounts'],['Retirement Accounts'],['Cash & Savings'],
    ['Real Estate'],['Business Ownership'],['Vehicles'],['Other Assets'],
  ]},
  { label: 'Liability Categories', data: [
    ['Mortgage'],['Student Loans'],['Auto Loans'],['Credit Cards'],
    ['Personal Loans'],['Business Debt'],['Other Liabilities'],
  ]},
  { label: 'Currencies', data: [
    ['USD'],['CAD'],['GBP'],['EUR'],['AUD'],['NZD'],['JPY'],['CHF'],['Other'],
  ]},
  { label: 'Yes / No', data: [['Yes'],['No']] },
  { label: 'Scenario Labels', data: [
    ['Conservative'],['Expected'],['Optimistic'],['Custom 1'],['Custom 2'],
  ]},
  { label: 'Priority', data: [['High'],['Medium'],['Low'],['None']] },
  { label: 'Contribution Statuses', data: [
    ['Planned'],['Scheduled'],['Completed'],['Skipped'],['Adjusted'],
  ]},
  { label: 'Planning Types', data: [['Individual'],['Couple / Household']] },
];

const ASSET_CLASS_COLORS = [
  ['U.S. Stocks',          '#496A88'],
  ['International Stocks', '#7A6680'],
  ['Emerging Markets',     '#C98474'],
  ['Bonds',                '#8FA98C'],
  ['Real Estate',          '#C6A15B'],
  ['REITs',                '#C6A15B'],
  ['Cash',                 '#8FB9B7'],
  ['Money Market',         '#8FB9B7'],
  ['Commodities',          '#D0B36A'],
  ['Gold / Precious Metals','#D0B36A'],
  ['Cryptocurrency',       '#74758F'],
  ['Alternatives',         '#A59482'],
  ['Other',                '#A9A9A6'],
];

(async () => {
  const vals = [];
  const fmt  = [];

  // Title
  vals.push({ range: `${S}!A1`, values: [['ULTIMATE INVESTMENT TRACKER — REFERENCE DATA (DO NOT EDIT)']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,5), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 32 }, fields: 'pixelSize' }});

  // Column widths
  [[0,240],[1,200],[2,120],[3,200],[4,120]].forEach(([ci, px]) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,1200,0,6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // Build sections sequentially, tracking current row
  let row = 2; // 0-indexed; row 1 is title
  const ranges = {}; // store range info for other scripts

  for (const sec of SECTIONS) {
    // Section header
    vals.push({ range: `${S}!A${row+1}`, values: [[sec.label]] });
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, 0, 5), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, 0, 5), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrB),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: row, endIndex: row+1 },
      properties: { pixelSize: 22 }, fields: 'pixelSize' }});

    const dStart = row + 1;
    ranges[sec.label] = { start: dStart + 1, count: sec.data.length }; // 1-indexed for A notation
    vals.push({ range: `${S}!A${dStart+1}`, values: sec.data });
    fmt.push({ repeatCell: { range: gridRange(SID, dStart, dStart + sec.data.length, 0, 5), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE',
      borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
    }}, fields: 'userEnteredFormat' }});

    row = dStart + sec.data.length + 1; // +1 spacer
  }

  // Asset-class color map
  vals.push({ range: `${S}!A${row+1}`, values: [['Asset-Class Color Map']] });
  fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, 0, 5), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, 0, 5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: row, endIndex: row+1 },
    properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  const acRow = row + 1;
  vals.push({ range: `${S}!A${acRow+1}`, values: ASSET_CLASS_COLORS });
  for (let i = 0; i < ASSET_CLASS_COLORS.length; i++) {
    const colorHex = ASSET_CLASS_COLORS[i][1];
    fmt.push({ repeatCell: { range: gridRange(SID, acRow+i, acRow+i+1, 0, 2), cell: { userEnteredFormat: {
      backgroundColor: hex(colorHex),
      textFormat: { fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }

  // Freeze row 1
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 1 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '02-reference values');
  await batchUpdate(id, fmt, '02-reference format');

  // Print key ranges for other scripts
  console.log('✅ Reference Data done.');
  let r2 = 2;
  for (const sec of SECTIONS) {
    const dStart = r2 + 1;
    console.log(`  ${sec.label.padEnd(30)} A${dStart+1}:A${dStart+sec.data.length}`);
    r2 = dStart + sec.data.length + 1;
  }
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
