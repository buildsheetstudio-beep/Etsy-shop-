'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH = sheetMap['Life Dashboard'];
const GOALS = sheetMap['Annual Goals'];
const QPLAN = sheetMap['Quarterly Planner'];
const MONTHLY = sheetMap['Monthly Milestones'];
const HABITS = sheetMap['Habit Tracker'];
const REF   = sheetMap['Reference'];

function dropdownRange(sheetId, row0, row1, col0, col1, refRange) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, row0, row1, col0, col1),
      rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: '=' + refRange }] },
        strict: true,
        showCustomUi: true,
      },
    },
  };
}

function numberBetween(sheetId, row0, row1, col0, col1, min, max) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, row0, row1, col0, col1),
      rule: {
        condition: { type: 'NUMBER_BETWEEN', values: [{ userEnteredValue: String(min) }, { userEnteredValue: String(max) }] },
        strict: true,
        showCustomUi: false,
        inputMessage: `Enter a value between ${min} and ${max}`,
      },
    },
  };
}

function checkbox(sheetId, row0, row1, col0, col1) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, row0, row1, col0, col1),
      rule: { condition: { type: 'BOOLEAN' }, strict: true, showCustomUi: true },
    },
  };
}

(async () => {
  const reqs = [];

  // ── Annual Goals: Life Area col (C, col 2) ────────────────────────────────
  const goalRows = [[4,7],[9,12],[15,18],[21,24],[27,30],[33,36]];
  goalRows.forEach(([s, e]) => {
    reqs.push(dropdownRange(GOALS, s, e, 2, 3, "Reference!$A$2:$A$8"));
  });

  // ── Annual Goals: Priority col (D, col 3) ─────────────────────────────────
  goalRows.forEach(([s, e]) => {
    reqs.push(dropdownRange(GOALS, s, e, 3, 4, "Reference!$I$2:$I$4"));
  });

  // ── Annual Goals: Status col (E, col 4) ───────────────────────────────────
  goalRows.forEach(([s, e]) => {
    reqs.push(dropdownRange(GOALS, s, e, 4, 5, "Reference!$C$2:$C$7"));
  });

  // ── Monthly: Life Area col (C, col 2) rows 5-10 ───────────────────────────
  reqs.push(dropdownRange(MONTHLY, 4, 10, 2, 3, "Reference!$A$2:$A$8"));
  // Monthly: Status col (D, col 3) rows 5-10
  reqs.push(dropdownRange(MONTHLY, 4, 10, 3, 4, "Reference!$C$2:$C$7"));

  // ── Quarterly: Quarter dropdown (B2) ─────────────────────────────────────
  reqs.push(dropdownRange(QPLAN, 1, 2, 1, 2, "Reference!$E$2:$E$5"));

  // ── Habit Tracker: Life Area col (C, col 2) rows 4-15 ────────────────────
  reqs.push(dropdownRange(HABITS, 3, 15, 2, 3, "Reference!$A$2:$A$8"));
  // Frequency col (D, col 3)
  reqs.push(dropdownRange(HABITS, 3, 15, 3, 4, "Reference!$K$2:$K$7"));

  // ── Habit Tracker: Checkboxes I4:AM15 (days 1-31) ────────────────────────
  reqs.push(checkbox(HABITS, 3, 15, 8, 39));

  // ── Life Dashboard: Wheel of Life score col I (col 8) rows 9-14 ──────────
  reqs.push(numberBetween(DASH, 8, 14, 8, 9, 1, 10));

  await batchUpdate(id, reqs, 'validation');

  console.log('Data validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
