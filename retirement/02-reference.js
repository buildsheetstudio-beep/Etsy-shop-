'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const RD = sheetMap['Reference Data'];
const S = "'Reference Data'";

const LISTS = {
  A: { header:'RSVP Statuses', col:'A', values:['Not Invited','Invitation Ready','Invitation Sent','Follow-Up Needed','Confirmed','Declined','Maybe','No Response'] },
  B: { header:'Guest Types', col:'C', values:['Coworker','Manager','Direct Report','Family','Friend','Neighbor','Client','Community Member','Other'] },
  C: { header:'Invitation Methods', col:'E', values:['Printed Invitation','Email','Text Message','Online Invitation','Phone Call','In Person','Workplace Announcement'] },
  D: { header:'Expense Categories', col:'G', values:['Venue','Invitations','Decorations','Food','Drinks','Cake & Desserts','Entertainment','Music','Audio / Visual','Gifts','Awards','Photography','Supplies','Transportation','Setup & Cleanup','Miscellaneous'] },
  E: { header:'Payment Statuses', col:'I', values:['Planned','Deposit Paid','Partially Paid','Paid','Refunded','Cancelled'] },
  F: { header:'Entertainment Types', col:'K', values:['Live Music','DJ','Playlist','Slideshow','Video Tribute','Speaker','Toast','Games','Photo Booth','Performance','Open Mic','Other'] },
  G: { header:'Music Categories', col:'M', values:['Arrival Music','Background Music','Dinner Music','Tribute Song','Dance Music','Closing Song','Retiree Favorite','Other'] },
  H: { header:'Thank-You Statuses', col:'O', values:['Not Needed','To Write','Drafted','Ready to Send','Sent','Delivered'] },
  I: { header:'Task Phases', col:'Q', values:['8+ Weeks Before','6 Weeks Before','4 Weeks Before','3 Weeks Before','2 Weeks Before','1 Week Before','3 Days Before','1 Day Before','Party Day','After Party'] },
  J: { header:'Task Statuses', col:'S', values:['Not Started','In Progress','Waiting','Complete','Cancelled'] },
  K: { header:'Priority Levels', col:'U', values:['Low','Medium','High','Urgent'] },
  L: { header:'Dietary Restrictions', col:'W', values:['None','Vegetarian','Vegan','Gluten-Free','Dairy-Free','Nut-Free','Halal','Kosher','Other','Unknown'] },
  M: { header:'Yes / No', col:'Y', values:['Yes','No'] },
};

(async () => {
  const data = [];
  data.push({ range:`${S}!A1`, values:[['REFERENCE DATA — DROPDOWN LISTS']] });

  for (const [, list] of Object.entries(LISTS)) {
    const startRow = (list.startRow || 1);
    const rows = [[list.header], ...list.values.map(v => [v])];
    data.push({ range:`${S}!${list.col}${startRow}`, values: rows });
  }

  await valuesBatchUpdate(id, data);

  const reqs = [];

  // Header row styling
  reqs.push({ repeatCell: {
    range: gridRange(RD, 0, 1, 0, 26),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 11 }
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)'
  }});

  // Column header formatting for each list
  const colIndices = { A:0, C:2, E:4, G:6, I:8, K:10, M:12, O:14, Q:16, S:18, U:20, W:22, Y:24 };
  for (const [, list] of Object.entries(LISTS)) {
    const ci = colIndices[list.col];
    const startRow = (list.startRow || 0);
    reqs.push({ repeatCell: {
      range: gridRange(RD, startRow, startRow+1, ci, ci+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.secondary),
        textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 }
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat)'
    }});
  }

  // Freeze row 1
  reqs.push({ updateSheetProperties: { properties: { sheetId: RD, gridProperties: { frozenRowCount: 1 } }, fields: 'gridProperties.frozenRowCount' } });

  // Hide the tab
  reqs.push({ updateSheetProperties: { properties: { sheetId: RD, hidden: true }, fields: 'hidden' } });

  await batchUpdate(id, reqs);
  console.log('Reference Data complete');
})();
