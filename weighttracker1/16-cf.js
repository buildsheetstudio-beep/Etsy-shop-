'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

// Sheet IDs
const SID_WEIGHT    = sheetMap['Weight Progress'];
const SID_MEAS      = sheetMap['Body Measurements'];
const SID_NUTRITION = sheetMap['Nutrition Log'];
const SID_WORKOUT   = sheetMap['Workout Tracker'];
const SID_HABITS    = sheetMap['Habit Tracker'];
const SID_GOALS     = sheetMap['Goals & Milestones'];
const SID_CHECKIN   = sheetMap['Weekly Check-In'];

function cfRule(ranges, rule, fmt) {
  return { addConditionalFormatRule: { rule: { ranges, booleanRule: { condition: rule, format: fmt } }, index: 0 } };
}

(async () => {
  const fmt = [];

  // ── Weight Progress: neutral banding for change col G (no red/green for gain/loss)
  // G = change vs prior — use neutral teal tint for any non-zero change (no directional color)
  fmt.push(cfRule(
    [gridRange(SID_WEIGHT, 7, 5000, 6, 7)],
    { type: 'NOT_BLANK' },
    { backgroundColor: hex(C.formula) }
  ));

  // ── Workout Tracker: intensity color coding (col F)
  // Intensity: Very Light, Light, Moderate, Hard, Very Hard, Max Effort
  const intensityRules = [
    { val: 'Very Light', color: C.planned },
    { val: 'Light',      color: C.secondaryTint },
    { val: 'Moderate',   color: C.active },
    { val: 'Hard',       color: C.partial },
    { val: 'Very Hard',  color: C.warning },
    { val: 'Max Effort', color: C.attention },
  ];
  intensityRules.forEach(({ val, color }) => {
    fmt.push(cfRule(
      [gridRange(SID_WORKOUT, 7, 5000, 5, 6)],
      { type: 'TEXT_EQ', values: [{ userEnteredValue: val }] },
      { backgroundColor: hex(color) }
    ));
  });

  // ── Goals & Milestones: status color coding (col D = index 3)
  const goalStatusRules = [
    { val: 'Achieved', color: C.achieved },
    { val: 'Active',   color: C.active },
    { val: 'Paused',   color: C.paused },
    { val: 'Abandoned',color: C.neutral },
  ];
  goalStatusRules.forEach(({ val, color }) => {
    fmt.push(cfRule(
      [gridRange(SID_GOALS, 7, 500, 0, 11)],
      { type: 'TEXT_EQ', values: [{ userEnteredValue: val }] },
      { backgroundColor: hex(color) }
    ));
  });

  // ── Goals: progress % col J highlight when >= 100%
  fmt.push(cfRule(
    [gridRange(SID_GOALS, 7, 500, 9, 10)],
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '1' }] },
    { backgroundColor: hex(C.achieved), textFormat: { bold: true } }
  ));

  // ── Weekly Check-In: highlight energy = 5 (excellent) in col D (index 3)
  fmt.push(cfRule(
    [gridRange(SID_CHECKIN, 7, 500, 3, 4)],
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '5' }] },
    { backgroundColor: hex(C.achieved) }
  ));

  // ── Weekly Check-In: flag stress = 4 or 5 in col F (index 5)
  fmt.push(cfRule(
    [gridRange(SID_CHECKIN, 7, 500, 5, 6)],
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '4' }] },
    { backgroundColor: hex(C.partial) }
  ));

  // ── Habit Tracker: habit def rows — highlight if "Done Today" (col G = index 6) = 1
  fmt.push(cfRule(
    [gridRange(SID_HABITS, 3, 13, 0, 7)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G4=1' }] },
    { backgroundColor: hex(C.achieved) }
  ));

  // ── Habit Log: Done col E (index 4) — checked rows get soft highlight
  fmt.push(cfRule(
    [gridRange(SID_HABITS, 19, 5000, 0, 6)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E20=TRUE' }] },
    { backgroundColor: hex(C.achieved) }
  ));

  // ── Nutrition Log: high-sodium flag (col M = index 12, >2000mg)
  fmt.push(cfRule(
    [gridRange(SID_NUTRITION, 7, 5000, 12, 13)],
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '2000' }] },
    { backgroundColor: hex(C.partial), textFormat: { bold: true } }
  ));

  // ── Meal Planner: "Made It?" checkboxes — highlight whole row when TRUE
  const SID_MEAL = sheetMap['Weekly Meal Planner'];
  fmt.push(cfRule(
    [gridRange(SID_MEAL, 7, 5000, 0, 7)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G8=TRUE' }] },
    { backgroundColor: hex(C.achieved) }
  ));

  // ── Grocery List: "Got It?" checkboxes — strikethrough style when checked
  const SID_GROC = sheetMap['Grocery List'];
  fmt.push(cfRule(
    [gridRange(SID_GROC, 7, 500, 0, 9)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$H8=TRUE' }] },
    { textFormat: { strikethrough: true }, backgroundColor: hex(C.formula) }
  ));

  await batchUpdate(id, fmt, '16-cf');
  console.log('✅  Conditional formatting done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
