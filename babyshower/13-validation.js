'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const CK = sheetMap['✅ Party Checklist'];
const GL = sheetMap['👥 Guest List & RSVPs'];
const BU = sheetMap['💰 Budget & Expenses'];
const IT = sheetMap['🗓️ Party Itinerary'];
const DC = sheetMap['🎀 Décor, Food & Drinks'];
const GA = sheetMap['🎮 Games & Activities'];
const GF = sheetMap['🎁 Gift Tracker'];
const REF = sheetMap['📋 Reference Data'];

function dropdownList(values) {
  return {
    condition: { type: 'ONE_OF_LIST', values: values.map(v => ({ userEnteredValue: v })) },
    showCustomUi: true,
    strict: false,
  };
}

function refRange(col, count) {
  // Reference column (0-indexed), row 2 onwards (skip header row 1)
  const colIdx = col.charCodeAt(0) - 65;
  return `'📋 Reference Data'!${col}2:${col}${count + 1}`;
}

(async () => {
  const reqs = [];

  // ── CHECKLIST ──
  // Col B: Timeframe (A column of Reference = index 0)
  reqs.push({ setDataValidation: {
    range: gridRange(CK, 3, 48, 1, 2),
    rule: dropdownList(['8+ Weeks Out','6 Weeks Out','4 Weeks Out','2 Weeks Out','1 Week Out','Day Before','Day Of','After Party']),
  }});
  // Col D: Priority
  reqs.push({ setDataValidation: {
    range: gridRange(CK, 3, 48, 3, 4),
    rule: dropdownList(['High','Medium','Low']),
  }});
  // Col G: Status
  reqs.push({ setDataValidation: {
    range: gridRange(CK, 3, 48, 6, 7),
    rule: dropdownList(['Not Started','In Progress','Done','Delegated']),
  }});

  // ── GUEST LIST ──
  // Col B: Relationship
  reqs.push({ setDataValidation: {
    range: gridRange(GL, 3, 23, 1, 2),
    rule: dropdownList(['Mum-to-Be','Family (Mum\'s side)','Family (Partner\'s side)','Close Friend','Work Friend','Neighbour','Partner\'s Friend','Child','Other']),
  }});
  // Col E: RSVP Status
  reqs.push({ setDataValidation: {
    range: gridRange(GL, 3, 23, 4, 5),
    rule: dropdownList(['Invited','Confirmed','Declined','No Response','Waitlisted']),
  }});
  // Col G: Dietary
  reqs.push({ setDataValidation: {
    range: gridRange(GL, 3, 23, 6, 7),
    rule: dropdownList(['None','Vegetarian','Vegan','Gluten Free','Halal','Kosher','Dairy Free','Nut Allergy','Other']),
  }});

  // ── BUDGET ──
  // Col B: Category (allocation table rows 4-15)
  reqs.push({ setDataValidation: {
    range: gridRange(BU, 3, 15, 1, 2),
    rule: dropdownList(['Venue','Food & Catering','Cake','Drinks','Decorations & Florals','Invitations & Stationery','Party Favours','Games & Prizes','Photography','Gift for Mum-to-Be','Miscellaneous','Buffer']),
  }});
  // Col B: Category in expense log (rows 18-32)
  reqs.push({ setDataValidation: {
    range: gridRange(BU, 17, 32, 1, 2),
    rule: dropdownList(['Venue','Food & Catering','Cake','Drinks','Decorations & Florals','Invitations & Stationery','Party Favours','Games & Prizes','Photography','Gift for Mum-to-Be','Miscellaneous','Buffer']),
  }});

  // ── ITINERARY ──
  // Col E: Activity Type
  reqs.push({ setDataValidation: {
    range: gridRange(IT, 3, 17, 4, 5),
    rule: dropdownList(['Welcome','Food & Drinks','Games','Gift Opening','Speeches','Photo','Activity','Setup','Farewell','Other']),
  }});

  // ── DÉCOR ──
  // Col B: Décor Category (rows 4-17)
  reqs.push({ setDataValidation: {
    range: gridRange(DC, 3, 17, 1, 2),
    rule: dropdownList(['Florals','Balloons','Signage','Table Setting','Backdrop','Favours','Lighting','Tableware','Misc']),
  }});
  // Col G: Décor Status (rows 4-17)
  reqs.push({ setDataValidation: {
    range: gridRange(DC, 3, 17, 6, 7),
    rule: dropdownList(['Needed','Ordered','Arrived','Set Up','Done']),
  }});
  // Col B: Food Type (rows 21-35)
  reqs.push({ setDataValidation: {
    range: gridRange(DC, 20, 35, 1, 2),
    rule: dropdownList(['Appetiser','Main','Dessert','Cake','Drinks','Non-Alcoholic','Snack','Other']),
  }});
  // Col C: Dietary Tag (rows 21-35)
  reqs.push({ setDataValidation: {
    range: gridRange(DC, 20, 35, 2, 3),
    rule: dropdownList(['Standard','Vegetarian','Vegan','Gluten Free','Dairy Free','Nut Free','Contains Nuts']),
  }});
  // Col G: Food Status (rows 21-35)
  reqs.push({ setDataValidation: {
    range: gridRange(DC, 20, 35, 6, 7),
    rule: dropdownList(['Planned','Ordered','Confirmed','Prepared','Ready']),
  }});

  // ── GAMES ──
  // Col B: Game Type
  reqs.push({ setDataValidation: {
    range: gridRange(GA, 2, 14, 1, 2),
    rule: dropdownList(['Classic','Creative','Interactive','Trivia','Team','Prize Game','Keepsake']),
  }});
  // Col I: Game Status
  reqs.push({ setDataValidation: {
    range: gridRange(GA, 2, 14, 8, 9),
    rule: dropdownList(['Planned','Prepped','Done']),
  }});

  // ── GIFT TRACKER ──
  // Col D: Gift Category
  reqs.push({ setDataValidation: {
    range: gridRange(GF, 3, 23, 3, 4),
    rule: dropdownList(['Clothing','Nursery','Feeding','Toys & Books','Gift Card','Bath & Care','Keepsake','Handmade','Other']),
  }});

  await batchUpdate(id, reqs, 'validation');
  console.log('Data validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
