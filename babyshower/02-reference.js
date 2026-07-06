'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

// Each column: one dropdown list
// A: Timeframes | B: Priorities | C: Task Statuses | D: Relationships
// E: RSVP Statuses | F: Dietary Options | G: Activity Types | H: Décor Categories
// I: Décor Statuses | J: Food Types | K: Dietary Tags | L: Food Supplier
// M: Food Statuses | N: Game Types | O: Game Statuses | P: Gift Categories
// Q: Budget Categories

const LISTS = {
  A: ['Timeframe','8+ Weeks Out','6 Weeks Out','4 Weeks Out','2 Weeks Out','1 Week Out','Day Before','Day Of','After Party'],
  B: ['Priority','High','Medium','Low'],
  C: ['Task Status','Not Started','In Progress','Done','Delegated'],
  D: ['Relationship','Mum-to-Be','Family (Mum\'s side)','Family (Partner\'s side)','Close Friend','Work Friend','Neighbour','Partner\'s Friend','Child','Other'],
  E: ['RSVP Status','Invited','Confirmed','Declined','No Response','Waitlisted'],
  F: ['Dietary','None','Vegetarian','Vegan','Gluten Free','Halal','Kosher','Dairy Free','Nut Allergy','Other'],
  G: ['Activity Type','Welcome','Food & Drinks','Games','Gift Opening','Speeches','Photo','Activity','Setup','Farewell','Other'],
  H: ['Décor Category','Florals','Balloons','Signage','Table Setting','Backdrop','Favours','Lighting','Tableware','Misc'],
  I: ['Décor Status','Needed','Ordered','Arrived','Set Up'],
  J: ['Food Type','Appetiser','Main','Dessert','Cake','Drinks','Non-Alcoholic','Snack','Other'],
  K: ['Dietary Tag','Standard','Vegetarian','Vegan','Gluten Free','Dairy Free','Nut Free','Contains Nuts'],
  L: ['Food Supplier','Host','Co-Host','Caterer','Bakery','Guest Contribution','Restaurant'],
  M: ['Food Status','Planned','Ordered','Confirmed','Prepared','Ready'],
  N: ['Game Type','Classic','Creative','Interactive','Trivia','Team','Prize Game','Keepsake'],
  O: ['Game Status','Planned','Prepped','Done'],
  P: ['Gift Category','Clothing','Nursery','Feeding','Toys & Books','Gift Card','Bath & Care','Keepsake','Handmade','Other'],
  Q: ['Budget Category','Venue','Food & Catering','Cake','Drinks','Decorations & Florals','Invitations & Stationery','Party Favours','Games & Prizes','Photography','Gift for Mum-to-Be','Miscellaneous','Buffer'],
};

(async () => {
  const reqs = [];

  // Column widths
  Object.keys(LISTS).forEach((col, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: REF, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: 170 },
      fields: 'pixelSize',
    }});
  });

  // Header row style
  reqs.push({ repeatCell: {
    range: gridRange(REF, 0, 1, 0, Object.keys(LISTS).length),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});

  // Data cells
  reqs.push({ repeatCell: {
    range: gridRange(REF, 1, 20, 0, Object.keys(LISTS).length),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.petalPink),
      textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});

  await batchUpdate(id, reqs, 'ref-format');

  // Values
  const vals = [];
  Object.entries(LISTS).forEach(([col, items]) => {
    vals.push({ range: `'📋 Reference Data'!${col}1`, values: items.map(v => [v]) });
  });
  await valuesBatchUpdate(id, vals, 'ref-values');

  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
