'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const RET = sheetMap['🔄 Return & Exchange Tracker'];

// Columns:
//   A(0): Recipient (formula)   B(1): Item Name        C(2): Store
//   D(3): Purchase Date         E(4): Return Deadline   F(5): Return Window
//   G(6): Reason                H(7): Return Status     I(8): Refund Amount ($)
//   J(9): Notes

const headers = ['Recipient','Item Returned','Store','Purchase Date','Return Deadline','Return Window (days)','Reason for Return','Return Status','Refund Amount ($)','Notes'];

// Recipient formula: look up item in GRT Gift Idea col (C) to find recipient
// =IFERROR(INDEX('🎁 Gift Recipient Tracker'!$A:$A,MATCH(B2,'🎁 Gift Recipient Tracker'!$C:$C,0)),"")
// Return deadline: =IFERROR(D2+F2,"")
// Return window: =IFERROR(VLOOKUP(C2,'📋 Reference Data'!$F:$G,2,FALSE),30)

const sampleData = [
  // [recipient_formula, item, store, purchase_date, deadline_formula, window_formula, reason, status, refund, notes]
  ['Mom','Duplicate cookbook','Amazon','2025-12-01','','','Received duplicate from another person','Completed',18.99,'Returned via UPS drop-off, refund received'],
  ['Jake (son)','Wrong LEGO set (got Technic City not Technic Race)','Target','2025-12-05','','','Wrong item — wanted different set','In Progress',62.99,'Exchange in progress at store'],
  ['Grandma Rose','Photo frame broken in shipping','Shutterfly','2025-12-03','','','Arrived damaged','Initiated',0,'Shutterfly sending replacement, no return needed'],
  ['Brother Mike','Earbuds — wrong model (bought SE not Pro)','Best Buy','2025-12-08','','','Wrong model purchased','Pending',79.99,'Need to return before Dec 23'],
  ['Aunt Carol','Wine gift basket — one bottle leaking','Harry & David','2025-12-06','','','Product defective','Completed',62,'Replacement basket sent at no charge'],
  ['Cousin Rachel','Yoga mat — color wrong','Amazon','2025-12-10','','','Wrong color shipped','Pending',37.49,'Print return label from Amazon'],
  ['Office raffle prize','Candle set — candles melted in shipping','TJ Maxx','2025-12-11','','','Arrived damaged','Pending',14.25,'Returning to store — need receipt'],
  ['Secret Santa draw','Board game missing pieces','Amazon','2025-12-07','','','Incomplete item','Initiated',22,'Amazon sending replacement game'],
  ['Baby nephew Leo','Wrong plush toy (Easter not Christmas theme)','Amazon','2025-12-09','','','Wrong item shipped','Pending',22,'Return window 30 days'],
  ['Neighbor Sue','Cookie tin broken lid','Homemade',  '','','','N/A — homemade','Denied',0,'Cannot return homemade items'],
];

(async () => {
  const numRows = sampleData.length;

  const recipientFormula = (r) =>
    `=IFERROR(INDEX('🎁 Gift Recipient Tracker'!$A:$A,MATCH(B${r},'🎁 Gift Recipient Tracker'!$C:$C,0)),"")`;
  const deadlineFormula  = (r) =>
    `=IFERROR(IF(D${r}="","",D${r}+F${r}),"")`;
  const windowFormula    = (r) =>
    `=IFERROR(VLOOKUP(C${r},'📋 Reference Data'!$F:$G,2,FALSE),30)`;

  const valueRows = sampleData.map((row, i) => {
    const r = i + 2;
    return [
      recipientFormula(r),
      row[1], row[2], row[3],
      deadlineFormula(r),
      windowFormula(r),
      row[6], row[7], row[8], row[9],
    ];
  });

  await valuesBatchUpdate(id, [
    { range: "'🔄 Return & Exchange Tracker'!A1", values: [headers] },
    { range: "'🔄 Return & Exchange Tracker'!A2", values: valueRows },
  ], 'returns-values');

  const reqs = [];

  // Header row
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

  // Data rows alternating
  for (let r = 1; r <= numRows; r++) {
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

  // Col A — formula col (italic)
  reqs.push({
    repeatCell: {
      range: gridRange(RET, 1, numRows + 1, 0, 1),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.formulaBg),
          textFormat: { italic: true, bold: true },
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // Col E — deadline date format + formula bg
  reqs.push({
    repeatCell: {
      range: gridRange(RET, 1, numRows + 1, 4, 5),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.formulaBg),
          textFormat: { italic: true },
          numberFormat: { type: 'DATE', pattern: 'MMM d, yyyy' },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat,horizontalAlignment)',
    },
  });

  // Col F — window formula bg
  reqs.push({
    repeatCell: {
      range: gridRange(RET, 1, numRows + 1, 5, 6),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.formulaBg),
          textFormat: { italic: true },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // Col D — purchase date format
  reqs.push({
    repeatCell: {
      range: gridRange(RET, 1, numRows + 1, 3, 4),
      cell: {
        userEnteredFormat: {
          numberFormat: { type: 'DATE', pattern: 'MMM d, yyyy' },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
    },
  });

  // Col I — currency
  reqs.push({
    repeatCell: {
      range: gridRange(RET, 1, numRows + 1, 8, 9),
      cell: {
        userEnteredFormat: {
          numberFormat: { type: 'NUMBER', pattern: '"$"#,##0.00' },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
    },
  });

  // Column widths
  [140, 180, 110, 110, 120, 120, 160, 110, 110, 180].forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: RET, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  reqs.push({ updateDimensionProperties: { range: { sheetId: RET, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'returns-format');
  console.log('Return & Exchange Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
