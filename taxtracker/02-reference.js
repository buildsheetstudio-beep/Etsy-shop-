'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

// Col A: Category | Col B: Limit% | Col C: Tax Year | Col D: IRS Rate
// Col E: Quarterly due dates | Col F: Doc checklist items

const categories = [
  ['Office Expenses',         1.0],
  ['Meals & Entertainment',   0.5],
  ['Travel',                  1.0],
  ['Vehicle/Mileage',         1.0],
  ['Home Office',             1.0],
  ['Professional Services',   1.0],
  ['Marketing & Advertising', 1.0],
  ['Insurance',               1.0],
  ['Education & Training',    1.0],
  ['Equipment & Technology',  1.0],
  ['Subscriptions & Software',1.0],
  ['Rent & Utilities',        1.0],
  ['Other',                   1.0],
];

const taxYears = [
  [2024, 0.67],
  [2025, 0.70],
  [2026, 0.725],
];

const quarterlyDates = ['Apr 15', 'Jun 15', 'Sep 15', 'Jan 15'];

const docItems = [
  'Receipts',
  'Mileage Log',
  'Bank/Card Statements',
  'Invoices Issued',
  'Home Office Photos/Measurements',
  'Prior Year Return',
];

(async () => {
  const reqs = [];

  // Header row (row 0)
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 0, 1, 0, 6),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepSapphire),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat',
    },
  });

  // Data rows background
  reqs.push({
    repeatCell: {
      range: gridRange(REF, 1, 20, 0, 6),
      cell: { userEnteredFormat: { backgroundColor: hex(C.warmIvory) } },
      fields: 'userEnteredFormat',
    },
  });

  // Col widths
  for (const [col, px] of [[0,180],[1,80],[2,80],[3,90],[4,100],[5,200]]) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: REF, dimension: 'COLUMNS', startIndex: col, endIndex: col+1 },
      properties: { pixelSize: px },
      fields: 'pixelSize',
    }});
  }

  await batchUpdate(id, reqs, 'reference-format');

  // Values
  const data = [];

  // Headers
  data.push({ range: `'📋 Reference Data'!A1`, values: [['Category','Limit %','Tax Year','IRS Rate $/mi','Q Due Date','Doc Item']] });

  // A+B: categories and limits
  data.push({ range: `'📋 Reference Data'!A2`, values: categories.map(([cat]) => [cat]) });
  data.push({ range: `'📋 Reference Data'!B2`, values: categories.map(([,lim]) => [lim]) });

  // C+D: tax years and rates
  data.push({ range: `'📋 Reference Data'!C2`, values: taxYears.map(([yr]) => [yr]) });
  data.push({ range: `'📋 Reference Data'!D2`, values: taxYears.map(([,rate]) => [rate]) });

  // E: quarterly due dates
  data.push({ range: `'📋 Reference Data'!E2`, values: quarterlyDates.map(d => [d]) });

  // F: doc checklist items
  data.push({ range: `'📋 Reference Data'!F2`, values: docItems.map(d => [d]) });

  await valuesBatchUpdate(id, data, 'reference-values');

  // Format B col as percent, D col as number with 3 dec
  const fmtReqs = [];
  fmtReqs.push({
    repeatCell: {
      range: gridRange(REF, 1, 15, 1, 2),
      cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });
  fmtReqs.push({
    repeatCell: {
      range: gridRange(REF, 1, 5, 3, 4),
      cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '$0.000' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });
  await batchUpdate(id, fmtReqs, 'reference-numfmt');

  console.log('Reference data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
