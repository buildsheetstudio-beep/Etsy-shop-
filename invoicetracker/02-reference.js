'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const RD = sheetMap['Reference Data'];
const S  = "'Reference Data'";

const lists = [
  ['Invoice Status',      ['Draft','Ready to Send','Sent','Viewed','Partially Paid','Paid','Overdue','Cancelled','Refunded']],
  ['Payment Status',      ['Unpaid','Deposit Paid','Partially Paid','Paid','Overpaid','Refunded','Cancelled']],
  ['Payment Methods',     ['Cash','Check','Bank Transfer','ACH','Credit Card','Debit Card','PayPal','Stripe','Venmo','Zelle','Other']],
  ['Payment Terms',       ['Due on Receipt','Net 7','Net 14','Net 15','Net 30','Net 45','Net 60','Custom']],
  ['Delivery Methods',    ['Email','Printed','Client Portal','In Person','Postal Mail','Other']],
  ['Item Types',          ['Product','Service','Hourly Work','Retainer','Deposit','Reimbursement','Shipping','Discount','Other']],
  ['Units',               ['Each','Hour','Day','Week','Month','Project','Session','Package','Mile','Custom']],
  ['Tax Treatment',       ['Taxable','Non-Taxable','Exempt','Tax Included']],
  ['Discount Types',      ['None','Percentage','Fixed Amount']],
  ['Currencies',          ['USD','CAD','GBP','EUR','AUD','NZD','Other']],
  ['Follow-Up Status',    ['Not Needed','Upcoming','Due Today','Overdue','Completed']],
  ['Yes/No',              ['Yes','No']],
  ['Payment Types',       ['Deposit','Partial Payment','Final Payment','Overpayment','Refund','Credit','Adjustment']],
];

(async () => {
  const data = [];
  lists.forEach(([header, vals], col) => {
    data.push({ range: `${S}!${colLetter(col)}1`, values: [[header]] });
    vals.forEach((v, row) => {
      data.push({ range: `${S}!${colLetter(col)}${row + 2}`, values: [[v]] });
    });
  });
  await valuesBatchUpdate(id, data, 'reference-values');

  const reqs = [];
  // Header row formatting
  reqs.push({ repeatCell: {
    range: gridRange(RD, 0, 1, 0, lists.length),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold: true, foregroundColor: hex(C.white) },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  // Freeze row 1
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: RD, gridProperties: { frozenRowCount: 1 } },
    fields: 'gridProperties.frozenRowCount',
  }});
  // Hide the tab
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: RD, hidden: true },
    fields: 'hidden',
  }});
  await batchUpdate(id, reqs, 'reference-format');
  console.log('✓ Reference Data');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });

function colLetter(idx) {
  let s = ''; idx += 1;
  while (idx > 0) { const r = (idx - 1) % 26; s = String.fromCharCode(65 + r) + s; idx = Math.floor((idx - 1) / 26); }
  return s;
}
