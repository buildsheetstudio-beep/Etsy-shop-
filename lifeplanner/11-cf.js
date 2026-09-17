'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const TSK = sheetMap['✅ Task Manager'];
const BUD = sheetMap['💰 Budget & Expense Tracker'];
const HAB = sheetMap['🔁 Habit Tracker'];
const WHL = sheetMap['⚖️ Life Balance Wheel'];

function cfRule(ranges, rule) {
  return { addConditionalFormatRule: { rule: { ranges, ...rule }, index: 0 } };
}

(async () => {
  const reqs = [];

  // Task Manager: sage bg for Done (col D = "Done")
  reqs.push(cfRule(
    [gridRange(TSK, 1, 31, 0, 6)],
    {
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$D2="Done"' }] },
        format: { backgroundColor: hex(C.sage) }
      }
    }
  ));

  // Task Manager: amber bg for In Progress
  reqs.push(cfRule(
    [gridRange(TSK, 1, 31, 0, 6)],
    {
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$D2="In Progress"' }] },
        format: { backgroundColor: hex(C.amber) }
      }
    }
  ));

  // Task Manager: rust text for overdue (Due Date < TODAY and Status != Done)
  reqs.push(cfRule(
    [gridRange(TSK, 1, 31, 0, 6)],
    {
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($E2<TODAY(),$E2<>"",$D2<>"Done")' }] },
        format: { textFormat: { foregroundColor: hex(C.rust), bold: true } }
      }
    }
  ));

  // Budget bills block (cols H–K = indices 7–11): sage if Paid=TRUE (col K = index 10)
  reqs.push(cfRule(
    [gridRange(BUD, 1, 9, 7, 11)],
    {
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$K2=TRUE' }] },
        format: { backgroundColor: hex(C.sage) }
      }
    }
  ));

  // Budget bills block: rust if Paid=FALSE and Due Date < TODAY (col J = index 9)
  reqs.push(cfRule(
    [gridRange(BUD, 1, 9, 7, 11)],
    {
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($K2=FALSE,$J2<TODAY(),$J2<>"")' }] },
        format: { backgroundColor: hex(C.rust), textFormat: { foregroundColor: hex(C.white) } }
      }
    }
  ));

  // Habit Tracker: sage fill when Completed=TRUE (col C = index 2)
  reqs.push(cfRule(
    [gridRange(HAB, 1, 41, 0, 4)],
    {
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$C2=TRUE' }] },
        format: { backgroundColor: hex(C.sage) }
      }
    }
  ));

  // Life Balance Wheel: amber if Goal - Current >= 3 (col B=current, col C=goal)
  reqs.push(cfRule(
    [gridRange(WHL, 1, 7, 0, 4)],
    {
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=($C2-$B2)>=3' }] },
        format: { backgroundColor: hex(C.amber) }
      }
    }
  ));

  await batchUpdate(id, reqs, 'cf-rules');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
