'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TAB = sheetMap['Chore & Allowance Tracker'];
const RWD = sheetMap['Reward & Prize Redemption'];

function cfRule(ranges, rule) {
  return { addConditionalFormatRule: { rule: { ranges, ...rule }, index: 0 } };
}

(async () => {
  const reqs = [];

  // Chore Log (rows 25-200, 0-indexed 24-199): sage fill when Paid Out (col F=5) is TRUE
  reqs.push(cfRule(
    [gridRange(TAB, 24, 200, 0, 6)],
    {
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F25=TRUE' }] },
        format: { backgroundColor: hex(C.sage) }
      }
    }
  ));

  // Chore Log: amber fill when Paid Out is FALSE AND date is more than 7 days ago
  reqs.push(cfRule(
    [gridRange(TAB, 24, 200, 0, 6)],
    {
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($F25=FALSE,$A25<TODAY()-7,$A25<>"")' }] },
        format: { backgroundColor: hex(C.amber) }
      }
    }
  ));

  // Reward & Prize Redemption: red text if Remaining Balance (col E=4) is negative
  reqs.push(cfRule(
    [gridRange(RWD, 1, 50, 4, 5)],
    {
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E2<0' }] },
        format: { textFormat: { foregroundColor: hex(C.rust), bold: true } }
      }
    }
  ));

  await batchUpdate(id, reqs, 'cf-rules');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
