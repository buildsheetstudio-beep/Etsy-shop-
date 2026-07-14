'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const RET = sheetMap['🔄 Return & Exchange Tracker'];

// A Recipient (formula)   B Item           C Store         D Purchase Date
// E Return Window (formula, days)           F Return Deadline (formula: =D+E)
// G Reason                H Return/Exchange status (dropdown)
// I Refund Amount ($)     J Notes

const headers = ['Recipient','Item Returned','Store','Purchase Date','Return Window (days)','Return Deadline','Reason','Return / Exchange','Refund Amount ($)','Notes'];

// 10 sample rows across multiple stores; mix of statuses
// Today is 2026-07-14. Set deadlines to show the CF triggers:
//   - Row with deadline within 3 days (Pending Return) → red CF
//   - Row with deadline past → past-due (still Pending Return, no special CF)
//   - Rows with Returned/Exchanged → sage CF
const sampleRows = [
  // item, store, purchase_date, reason, status, refund, notes
  ['Duplicate cookbook (received same from Aunt Carol)',  'Amazon',   '2025-12-01', 'Duplicate gift',            'Returned',      18.99, 'Refund received Dec 8'],
  ['Wrong LEGO set — wanted Technic Race not City',      'Target',   '2025-12-05', 'Wrong item — exchanged in store', 'Exchanged', 0,     'Exchanged Dec 7 — correct set now wrapped'],
  ['Earbuds wrong model (SE not Pro)',                   'Best Buy', '2025-12-08', 'Wrong model — wanted Pro', 'Pending Return', 79.99, 'Must return by Dec 23 — receipt in wallet'],
  ['Wine basket — one bottle arrived leaking',           'Other',    '2025-12-06', 'Damaged on arrival',        'Returned',      62,    'Harry & David sent replacement basket'],
  ['Yoga mat wrong colour (shipped green, ordered purple)','Amazon',  '2025-12-10', 'Wrong colour shipped',      'Pending Return',37.49, 'Print return label from Amazon app'],
  ['Candle set — candles cracked in transit',            'Other',    '2025-12-11', 'Arrived damaged',           'Pending Return',14.25, 'Returning to nearest TJ Maxx with receipt'],
  ['Board game missing 3 key pieces',                   'Amazon',   '2025-12-07', 'Incomplete item',           'Exchanged',     0,     'Amazon sent replacement — no return needed'],
  ['Video game — disc scratched out of box',             'Walmart',  '2025-12-09', 'Defective disc',            'Returned',      59.99, 'Returned to Walmart Dec 12 — cash refund'],
  ['Cordless drill — battery won\'t charge',             'Other',    '2025-12-03', 'Defective product',         'Pending Return',74.99, 'Home Depot — need receipt; 90-day policy'],
  ['Artisan candle set wrong scent (got vanilla, wanted cedar)','Etsy','2025-12-12','Wrong scent sent',          'Pending Return',32,    'Etsy seller agreed to exchange; awaiting label'],
];

(async () => {
  const recipientFormula = (r) =>
    `=IFERROR(INDEX('🎁 Gift Recipient Tracker'!$A:$A,MATCH(B${r},'🎁 Gift Recipient Tracker'!$C:$C,0)),"")`;
  const windowFormula = (r) =>
    `=IFERROR(VLOOKUP(C${r},'📋 Reference Data'!$F:$G,2,FALSE),30)`;
  const deadlineFormula = (r) =>
    `=IFERROR(IF(D${r}="","",D${r}+E${r}),"")`;

  const valueRows = sampleRows.map((row, i) => {
    const r = i + 2;
    return [
      recipientFormula(r),
      row[0], row[1], row[2],
      windowFormula(r),
      deadlineFormula(r),
      row[3], row[4], row[5], row[6],
    ];
  });

  await valuesBatchUpdate(id, [
    { range: "'🔄 Return & Exchange Tracker'!A1", values: [headers] },
    { range: "'🔄 Return & Exchange Tracker'!A2", values: valueRows },
  ], 'returns-values');

  const reqs = [];
  const n = sampleRows.length;

  // Header
  reqs.push({
    repeatCell: {
      range: gridRange(RET, 0, 1, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.mutedSage),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });

  for (let r = 1; r <= n; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(RET, r, r + 1, 0, 10),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(r % 2 === 1 ? C.snowWhite : C.altRow),
            textFormat: { foregroundColor: hex(C.darkText), fontSize: 10 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }

  // Col A — formula italic
  reqs.push({ repeatCell: { range: gridRange(RET, 1, n + 1, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { italic: true, bold: true } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Col E — formula bg, center
  reqs.push({ repeatCell: { range: gridRange(RET, 1, n + 1, 4, 5), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { italic: true }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  // Col F — formula bg, date format
  reqs.push({ repeatCell: { range: gridRange(RET, 1, n + 1, 5, 6), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { italic: true }, numberFormat: { type: 'DATE', pattern: 'MMM d, yyyy' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat,horizontalAlignment)' } });

  // Col D — date format
  reqs.push({ repeatCell: { range: gridRange(RET, 1, n + 1, 3, 4), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'MMM d, yyyy' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(numberFormat,horizontalAlignment)' } });

  // Col I — currency
  reqs.push({ repeatCell: { range: gridRange(RET, 1, n + 1, 8, 9), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '"$"#,##0.00' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(numberFormat,horizontalAlignment)' } });

  [140, 220, 90, 110, 110, 120, 160, 120, 110, 200].forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: RET, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: RET, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 38 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'returns-format');
  console.log('Return & Exchange Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
