'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

const LISTS = {
  A: { header: 'Timeframe', items: ['3+ Months Out','2 Months Out','6 Weeks Out','4 Weeks Out','2 Weeks Out','1 Week Out','Day Before','Day Of','After Reunion'] },
  B: { header: 'Task Status', items: ['Not Started','In Progress','Done','Delegated'] },
  C: { header: 'Priority', items: ['High','Medium','Low'] },
  D: { header: 'Age Group', items: ['Baby (0–2)','Child (3–12)','Teen (13–17)','Adult (18–59)','Senior (60+)'] },
  E: { header: 'RSVP Status', items: ['Invited','Confirmed','Declined','No Response','Waitlisted'] },
  F: { header: 'Dietary', items: ['None','Vegetarian','Vegan','Gluten Free','Halal','Kosher','Dairy Free','Nut Allergy','Other'] },
  G: { header: 'Food Type', items: ['Appetiser','Main','Salad & Sides','Dessert','Snack','Drinks','Non-Alcoholic','Condiment','Other'] },
  H: { header: "Who's Bringing", items: ['Organiser','Potluck','Catering','Purchase','Guest Contribution'] },
  I: { header: 'Accommodation Type', items: ['Cabin Room','Whole Cabin','Hotel Room','Tent','Caravan','Private Home','Other'] },
  J: { header: 'Booking Status', items: ['Not Booked','Deposit Paid','Fully Booked','Cancelled'] },
  K: { header: 'Transport', items: ['Driving','Flying','Bus','Train','Other'] },
  L: { header: 'Activity Status', items: ['Planned','Supplies Bought','Ready','Done'] },
  M: { header: 'Activity Ages', items: ['All Ages','Kids Only','Adults Only','Teens + Adults','Family Teams'] },
  N: { header: 'Activity Type', items: ['Outdoor','Indoor','Games','Creative','Swimming','Trivia','Team','For Kids','All Ages'] },
  O: { header: 'Itinerary Type', items: ['Welcome','Meal','Activity','Games','Speech','Photo','Free Time','Travel','Setup','Farewell','Other'] },
  P: { header: 'Day', items: ['Day 1','Day 2','Day 3','Pre-Event','Post-Event'] },
  Q: { header: 'Meal Slot', items: ['Day 1 Lunch','Day 1 Dinner','Day 1 Snacks','Day 1 Drinks','Day 2 Breakfast','Day 2 Lunch','Day 2 Dinner','Day 2 Drinks','All Day','Other'] },
  R: { header: 'Food Status', items: ['Planned','Assigned','Confirmed','Purchased','Ready'] },
  S: { header: 'Dietary Tags', items: ['Standard','Vegetarian','Vegan','Gluten Free','Dairy Free','Nut Free','Contains Nuts'] },
  T: { header: 'Budget Category', items: ['Venue','Accommodation','Catering & Food','Drinks & Ice','Activities & Games','Decorations','Printed Materials & Signage','Transport & Shuttle','Photography','Keepsakes & Gifts','Miscellaneous','Buffer'] },
  U: { header: 'Family Branch', items: ['Henderson Side','Thompson Side','Wilson Side','Smith Side','Out-of-Town','Family Friends'] },
};

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'📋 Reference Data'";

  // Column widths
  Object.keys(LISTS).forEach((col, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: REF, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: 160 }, fields: 'pixelSize',
    }});
  });

  // Title banner
  reqs.push({ mergeCells: { range: gridRange(REF,0,1,0,21), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(REF,0,1,0,21),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: REF, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A1`, values: [['📋 Reference Data — Family Reunion Planner']] });

  // Header row (row 1, 0-indexed)
  reqs.push({ repeatCell: {
    range: gridRange(REF,1,2,0,21),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: REF, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Data styling
  reqs.push({ repeatCell: {
    range: gridRange(REF,2,60,0,21),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.warmCream),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});

  // Write headers and data
  const headers = Object.values(LISTS).map(l => l.header);
  vals.push({ range: `${TAB}!A2`, values: [headers] });

  Object.entries(LISTS).forEach(([col, { items }]) => {
    vals.push({ range: `${TAB}!${col}3`, values: items.map(v => [v]) });
  });

  await batchUpdate(id, reqs, 'ref-format');
  await valuesBatchUpdate(id, vals, 'ref-values');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
