'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'📋 Reference Data'!A1", values: [['Life Category'],['Personal Growth'],['Health & Wellness'],['Career & Business'],['Home & Household'],['Relationships & Social'],['Finance']] },
    { range: "'📋 Reference Data'!B1", values: [['Task Priority'],['High'],['Medium'],['Low']] },
    { range: "'📋 Reference Data'!C1", values: [['Task Status'],['To Do'],['In Progress'],['Done']] },
    { range: "'📋 Reference Data'!D1", values: [['Budget Category'],['Housing'],['Food'],['Transportation'],['Utilities'],['Entertainment'],['Health'],['Savings'],['Debt Payments'],['Personal'],['Other']] },
    { range: "'📋 Reference Data'!E1", values: [['Mood'],['Great'],['Good'],['Okay'],['Low'],['Rough']] },
    { range: "'📋 Reference Data'!F1", values: [['Meal Type'],['Breakfast'],['Lunch'],['Dinner'],['Snack']] },
  ], 'ref-values');

  const reqs = [];
  reqs.push({ repeatCell: { range: gridRange(REF,0,1,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.raspberry), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  reqs.push({ repeatCell: { range: gridRange(REF,1,12,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.offWhite), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  [[0,140],[1,110],[2,100],[3,140],[4,80],[5,90]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: REF, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  await batchUpdate(id, reqs, 'ref-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
