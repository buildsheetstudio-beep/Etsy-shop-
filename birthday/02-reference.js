'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const RD = sheetMap['Reference Data'];
const S = "'Reference Data'";

const LISTS = {
  A: { header:'RSVP Statuses', col:'A', values:['Not Invited','Invitation Ready','Invitation Sent','Follow-Up Needed','Confirmed','Declined','Maybe','No Response'] },
  B: { header:'Guest Types', col:'C', values:['Child','Parent / Guardian','Sibling','Family','Friend','Classmate','Neighbor','Other'] },
  C: { header:'Invitation Methods', col:'E', values:['Printed Invitation','Email','Text Message','Online Invitation','Phone Call','Social Media','In Person'] },
  D: { header:'Expense Categories', col:'G', values:['Venue','Invitations','Decorations','Food','Drinks','Cake & Desserts','Entertainment','Games & Activities','Party Favors','Supplies','Clothing','Photography','Transportation','Setup & Cleanup','Miscellaneous'] },
  E: { header:'Payment Statuses', col:'I', values:['Planned','Deposit Paid','Partially Paid','Paid','Refunded','Cancelled'] },
  F: { header:'Food Types', col:'K', values:['Appetizer','Main Dish','Side','Snack','Dessert','Cake','Beverage','Adult Food','Kids Food','Other'] },
  G: { header:'Dietary Restrictions', col:'M', values:['None','Vegetarian','Vegan','Gluten-Free','Dairy-Free','Nut-Free','Egg-Free','Halal','Kosher','Other','Unknown'] },
  H: { header:'Shopping Categories', col:'O', values:['Food','Drinks','Cake & Desserts','Decorations','Tableware','Games','Activities','Favors','Cleaning','Safety','Miscellaneous'] },
  I: { header:'Task Phases', col:'Q', values:['8+ Weeks Before','6 Weeks Before','4 Weeks Before','3 Weeks Before','2 Weeks Before','1 Week Before','3 Days Before','1 Day Before','Party Day','After Party'] },
  J: { header:'Task Statuses', col:'S', values:['Not Started','In Progress','Waiting','Complete','Cancelled'] },
  K: { header:'Priority Levels', col:'U', values:['Low','Medium','High','Urgent'] },
  L: { header:'Yes / No', col:'W', values:['Yes','No'] },
  M: { header:'Party Types', col:'Y', values:['Home Party','Venue Party','Outdoor Party','Sleepover Party','Activity Party','Pool Party','Virtual Party'] },
  N: { header:'Child or Adult', col:'A', startRow:12, values:['Child','Adult'] },
  O: { header:'Months', col:'P', startRow:2, values:['January','February','March','April','May','June','July','August','September','October','November','December'] },
};

(async () => {
  const data = [];

  // Main header row
  data.push({ range:`${S}!A1`, values:[['REFERENCE DATA — DROPDOWN LISTS']] });

  // Write all lists
  for (const key of Object.keys(LISTS)) {
    const { header, col, values, startRow } = LISTS[key];
    const sr = startRow || 1;
    data.push({ range:`${S}!${col}${sr}`, values:[[header]] });
    values.forEach((v, i) => {
      data.push({ range:`${S}!${col}${sr+1+i}`, values:[[v]] });
    });
  }

  await valuesBatchUpdate(id, data, 'reference-values');

  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(RD,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };

  // Title row
  fmt(0,1,0,25,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:12, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Section headers (row 1 of each list)
  const headerCols = [0,2,4,6,8,10,12,14,16,18,20,22,24,0,15];
  const headerRows = [1,1,1,1,1,1,1,1,1,1,1,1,1,12,2];
  headerCols.forEach((c,i) => {
    fmt(headerRows[i]-1,headerRows[i],c,c+1,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  });

  reqs.push({ updateDimensionProperties:{ range:{ sheetId:RD, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:28 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:RD, dimension:'ROWS', startIndex:1, endIndex:30 }, properties:{ pixelSize:20 }, fields:'pixelSize' } });

  // Freeze row 1
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:RD, gridProperties:{ frozenRowCount:1 } }, fields:'gridProperties.frozenRowCount' } });

  // Hide Reference Data tab
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:RD, hidden:true }, fields:'hidden' } });

  await batchUpdate(id, reqs, 'reference-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
