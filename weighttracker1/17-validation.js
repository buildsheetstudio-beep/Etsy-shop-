'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID_WEIGHT    = sheetMap['Weight Progress'];
const SID_NUTRITION = sheetMap['Nutrition Log'];
const SID_WORKOUT   = sheetMap['Workout Tracker'];
const SID_HABITS    = sheetMap['Habit Tracker'];
const SID_GOALS     = sheetMap['Goals & Milestones'];
const SID_CHECKIN   = sheetMap['Weekly Check-In'];
const SID_MEAL      = sheetMap['Weekly Meal Planner'];
const SID_GROC      = sheetMap['Grocery List'];
const SID_REF       = sheetMap['Reference Data'];

function oneOfRange(sheetId, col, r1, r2, rangeA1) {
  return { setDataValidation: { range: gridRange(sheetId, r1, r2, col, col + 1), rule: {
    condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: rangeA1 }] },
    strict: false, showCustomUi: true,
  }}};
}

function oneOfList(sheetId, col, r1, r2, values) {
  return { setDataValidation: { range: gridRange(sheetId, r1, r2, col, col + 1), rule: {
    condition: { type: 'ONE_OF_LIST', values: values.map(v => ({ userEnteredValue: v })) },
    strict: false, showCustomUi: true,
  }}};
}

function checkbox(sheetId, col, r1, r2) {
  return { setDataValidation: { range: gridRange(sheetId, r1, r2, col, col + 1), rule: {
    condition: { type: 'BOOLEAN' }, strict: true,
  }}};
}

function numRange(sheetId, col, r1, r2, min, max) {
  return { setDataValidation: { range: gridRange(sheetId, r1, r2, col, col + 1), rule: {
    condition: { type: 'NUMBER_BETWEEN', values: [{ userEnteredValue: String(min) }, { userEnteredValue: String(max) }] },
    strict: false, showCustomUi: false,
  }}};
}

(async () => {
  const req = [];

  // ── Weight Progress: Weight Units dropdown (col H = index 7)
  req.push(oneOfRange(SID_WEIGHT, 7, 7, 5000, "='Reference Data'!H2:H3"));

  // ── Nutrition Log: Meal Type dropdown (col C = index 2)
  req.push(oneOfRange(SID_NUTRITION, 2, 7, 5000, "='Reference Data'!B2:B8"));

  // ── Workout Tracker: Type dropdown (col C = index 2)
  req.push(oneOfRange(SID_WORKOUT, 2, 7, 5000, "='Reference Data'!A2:A9"));

  // ── Workout Tracker: Intensity dropdown (col F = index 5)
  req.push(oneOfRange(SID_WORKOUT, 5, 7, 5000, "='Reference Data'!G2:G7"));

  // ── Habit Tracker: Category dropdown for habit definitions (col C = index 2, rows 4-13 = ri 3-13)
  req.push(oneOfRange(SID_HABITS, 2, 3, 13, "='Reference Data'!C2:C9"));

  // ── Habit Log: Done checkbox (col E = index 4, rows 20+ = ri 19+)
  req.push(checkbox(SID_HABITS, 4, 19, 5000));

  // ── Goals: Goal Type dropdown (col C = index 2)
  req.push(oneOfRange(SID_GOALS, 2, 7, 500, "='Reference Data'!D2:D7"));

  // ── Goals: Status dropdown (col D = index 3)
  req.push(oneOfRange(SID_GOALS, 3, 7, 500, "='Reference Data'!E2:E5"));

  // ── Weekly Check-In: Energy 1-5 (col D = index 3)
  req.push(numRange(SID_CHECKIN, 3, 7, 500, 1, 5));

  // ── Weekly Check-In: Sleep Quality 1-5 (col E = index 4)
  req.push(numRange(SID_CHECKIN, 4, 7, 500, 1, 5));

  // ── Weekly Check-In: Stress 1-5 (col F = index 5)
  req.push(numRange(SID_CHECKIN, 5, 7, 500, 1, 5));

  // ── Weekly Meal Planner: Meal Slot dropdown (col C = index 2)
  req.push(oneOfRange(SID_MEAL, 2, 7, 5000, "='Reference Data'!K2:K7"));

  // ── Weekly Meal Planner: Day dropdown (col B = index 1)
  req.push(oneOfRange(SID_MEAL, 1, 7, 5000, "='Reference Data'!J2:J8"));

  // ── Weekly Meal Planner: Made It? checkbox (col G = index 6)
  req.push(checkbox(SID_MEAL, 6, 7, 5000));

  // ── Grocery List: Category dropdown (col C = index 2)
  req.push(oneOfRange(SID_GROC, 2, 7, 500, "='Reference Data'!I2:I13"));

  // ── Grocery List: Needed? checkbox (col G = index 6)
  req.push(checkbox(SID_GROC, 6, 7, 500));

  // ── Grocery List: Got It? checkbox (col H = index 7)
  req.push(checkbox(SID_GROC, 7, 7, 500));

  await batchUpdate(id, req, '17-validation');
  console.log('✅  Data validation done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
