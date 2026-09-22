'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const REF = sheetMap['📋 Reference Data'];
const TSK = sheetMap['✅ Task Manager'];
const BUD = sheetMap['💰 Budget & Expense Tracker'];
const HAB = sheetMap['🔁 Habit Tracker'];
const WHL = sheetMap['⚖️ Life Balance Wheel'];
const RFL = sheetMap['📝 Weekly & Monthly Reflection'];
const JRN = sheetMap['📓 Journal & Gratitude Log'];

function dropdown(sheetId, r1, r2, c1, c2, formula) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r1, r2, c1, c2),
      rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: formula }] },
        strict: true,
        showCustomUi: true
      }
    }
  };
}

function inlineDropdown(sheetId, r1, r2, c1, c2, values) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r1, r2, c1, c2),
      rule: {
        condition: { type: 'ONE_OF_LIST', values: values.map(v => ({ userEnteredValue: v })) },
        strict: true,
        showCustomUi: true
      }
    }
  };
}

function checkbox(sheetId, r1, r2, c1, c2) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r1, r2, c1, c2),
      rule: {
        condition: { type: 'BOOLEAN' },
        strict: true,
        showCustomUi: true
      }
    }
  };
}

(async () => {
  const reqs = [];

  // Task Manager: Life Category (col B=1), Priority (col C=2), Status (col D=3)
  reqs.push(dropdown(TSK, 1, 31, 1, 2, "='📋 Reference Data'!$A$2:$A$7"));
  reqs.push(dropdown(TSK, 1, 31, 2, 3, "='📋 Reference Data'!$B$2:$B$4"));
  reqs.push(dropdown(TSK, 1, 31, 3, 4, "='📋 Reference Data'!$C$2:$C$4"));

  // Budget transactions: Category (col B=1), Type (col D=3 = Income/Expense)
  reqs.push(dropdown(BUD, 1, 31, 1, 2, "='📋 Reference Data'!$D$2:$D$11"));
  reqs.push(inlineDropdown(BUD, 1, 31, 3, 4, ['Income', 'Expense']));

  // Budget bills: Paid checkbox (col K = index 10)
  reqs.push(checkbox(BUD, 1, 9, 10, 11));

  // Habit Tracker: Completed checkbox (col C = index 2)
  reqs.push(checkbox(HAB, 1, 41, 2, 3));

  // Life Balance Wheel: Life Category (col A=0)
  reqs.push(dropdown(WHL, 1, 7, 0, 1, "='📋 Reference Data'!$A$2:$A$7"));

  // Weekly & Monthly Reflection: Period Type (col B=1)
  reqs.push(inlineDropdown(RFL, 1, 9, 1, 2, ['Weekly', 'Monthly']));

  // Journal & Gratitude Log: Mood (col F=5)
  reqs.push(dropdown(JRN, 1, 15, 5, 6, "='📋 Reference Data'!$E$2:$E$6"));

  await batchUpdate(id, reqs, 'validation');
  console.log('Validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
