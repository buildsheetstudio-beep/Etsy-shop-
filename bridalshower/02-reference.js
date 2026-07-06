'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

const LISTS = {
  A: ['Relationship','Bride','MOH','Bridesmaid','Family (Bride\'s)','Family (Partner\'s)','Close Friend','Work Friend','Mutual Friend','Other'],
  B: ['RSVP Status','Invited','Confirmed','Declined','No Response','Waitlisted'],
  C: ['Dietary','None','Vegetarian','Vegan','Gluten Free','Halal','Kosher','Dairy Free','Nut Allergy','Other'],
  D: ['Game Type','Classic','Creative','Trivia','Team','Keepsake','Prize Game','Icebreaker'],
  E: ['Vendor Category','Venue','Caterer','Florist','Photographer','Baker','Stylist','Other'],
  F: ['Task Status','Not Started','In Progress','Done','Delegated'],
  G: ['Priority','High','Medium','Low'],
  H: ['Gift Category','Homewares','Clothing & Accessories','Gift Card','Experience','Food & Drink','Books & Stationery','Keepsake','Handmade','Other'],
  I: ['Wishlist Priority','Must Have','Would Love','Nice to Have'],
  J: ['Payment Method','Credit Card','Debit Card','Bank Transfer','Cash','PayPal','Gift','Other'],
  K: ['Budget Category','Venue Hire','Food & Catering','Cake & Desserts','Drinks','Decorations & Florals','Invitations & Stationery','Party Favours','Games & Prizes','Photography','Gift for Bride','Entertainment','Miscellaneous'],
  L: ['Table Name','Bride\'s Table','Table 1','Table 2','Table 3','Table 4','Table 5','Table 6','Not Yet Assigned'],
  M: ['Activity Type','Welcome','Food & Drinks','Games','Gift Opening','Speeches','Photo','Activity','Bride\'s Surprise','Farewell','Setup','Other'],
  N: ['Bridesmaid Role','Maid of Honour','Bridesmaid','Co-Host','Helper'],
  O: ['Task Category','Planning','Logistics','Day-Of','Shopping','Coordination','Communication','Other'],
  P: ['Availability','Yes','No','Pending'],
  Q: ['Rating','1 — Poor','2 — Fair','3 — Good','4 — Very Good','5 — Excellent'],
};

(async () => {
  const reqs = [];

  Object.keys(LISTS).forEach((col, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: REF, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: 180 }, fields: 'pixelSize',
    }});
  });

  reqs.push({ repeatCell: {
    range: gridRange(REF, 0, 1, 0, Object.keys(LISTS).length),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});

  reqs.push({ repeatCell: {
    range: gridRange(REF, 1, 22, 0, Object.keys(LISTS).length),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.parchment),
      textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});

  await batchUpdate(id, reqs, 'ref-format');

  const vals = [];
  Object.entries(LISTS).forEach(([col, items]) => {
    vals.push({ range: `'📋 Reference Data'!${col}1`, values: items.map(v => [v]) });
  });
  await valuesBatchUpdate(id, vals, 'ref-values');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
