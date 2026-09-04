'use strict';
const { C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

const LISTS = [
  { header: 'RSVP Status',      items: ['Attending','Declined','Maybe','Awaiting Reply'] },
  { header: 'Dietary',          items: ['None','Vegetarian','Vegan','Gluten-Free','Nut Allergy','Dairy-Free','Halal','Kosher','Other'] },
  { header: 'Meal Choice',      items: ['Chicken','Fish','Beef','Vegetarian','Vegan','Kids Meal'] },
  { header: 'Gift Status',      items: ['Not Given','Given – Cash','Given – Registry','Given – Other','Thank You Sent'] },
  { header: 'Side',             items: ['Partner 1','Partner 2','Both','Friend of Both'] },
  { header: 'Vendor Category',  items: ['Venue','Catering','Photography','Videography','Florals','Music/DJ','Hair & Makeup','Cake/Desserts','Invitations','Other'] },
  { header: 'Payment Status',   items: ['Not Paid','Deposit Paid','Fully Paid','Refunded'] },
  { header: 'Timeline Type',    items: ['Setup','Ceremony','Reception','Catering','Entertainment','Photography','Vendor','Transportation','Cleanup','Buffer'] },
  { header: 'Table Shape',      items: ['Round','Rectangular','Square','Oval'] },
  { header: 'Toast Status',     items: ['Confirmed','Pending','Draft Written','Speech Ready','Delivered'] },
  { header: 'Shot Type',        items: ['Group','Portrait','Detail','Candid','Venue','Food & Drink','Ceremony','Reception'] },
  { header: 'Shot Status',      items: ['Must Have','Nice to Have','Completed'] },
  { header: 'Item Category',    items: ['Decor','Food','Drinks','Desserts','Favors','Supplies'] },
  { header: 'Task Category',    items: ['Venue','Catering','Invitations','Decor','Beauty','Day-Of','Other'] },
  { header: 'Task Status',      items: ['Not Started','In Progress','Complete','Delegated','Cancelled'] },
  { header: 'Priority',         items: ['High','Medium','Low'] },
];

(async () => {
  const reqs = [];

  // Header row format
  reqs.push({ repeatCell: {
    range: gridRange(REF, 0, 1, 0, LISTS.length),
    cell: {
      userEnteredFormat: {
        backgroundColor: hex(C.deepPlum),
        textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
        horizontalAlignment: 'CENTER',
      },
    },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});

  // Data rows alternating
  const maxRows = Math.max(...LISTS.map(l => l.items.length));
  for (let r = 0; r < maxRows; r++) {
    const bg = r % 2 === 0 ? C.ivoryCream : C.altRow;
    reqs.push({ repeatCell: {
      range: gridRange(REF, r + 1, r + 2, 0, LISTS.length),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(bg),
          textFormat: { foregroundColor: hex(C.bodyText), fontSize: 10 },
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    }});
  }

  // Column widths (80px each)
  LISTS.forEach((_, ci) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: REF, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
      properties: { pixelSize: 120 },
      fields: 'pixelSize',
    }});
  });

  await batchUpdate(id, reqs, 'reference-format');

  // Values
  const data = [];
  data.push({ range: `'📋 Reference Data'!A1:P1`, values: [LISTS.map(l => l.header)] });
  LISTS.forEach((list, ci) => {
    const col = String.fromCharCode(65 + ci);
    data.push({
      range: `'📋 Reference Data'!${col}2:${col}${list.items.length + 1}`,
      values: list.items.map(v => [v]),
    });
  });

  await valuesBatchUpdate(id, data, 'reference-values');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
