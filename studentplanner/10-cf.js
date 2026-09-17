'use strict';
const { batchUpdate, hex, COLOR, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID_DASH   = sheetMap['Dashboard'];
const SID_ASSIGN = sheetMap['Assignments'];
const SID_WEEK   = sheetMap['Weekly Planner'];
const SID_GRADES = sheetMap['Grades & GPA'];
const SID_BUDGET = sheetMap['Student Budget'];
const SID_EXAM   = sheetMap['Exam Prep'];

function cfRule(ranges, condition, format) {
  return { addConditionalFormatRule: { rule: { ranges, booleanRule: { condition, format } }, index: 0 } };
}

function bg(colorHex) { return { backgroundColor: hex(colorHex) }; }
function bgText(bgHex, fgHex, bold = false) {
  return { backgroundColor: hex(bgHex), textFormat: { foregroundColor: hex(fgHex), bold } };
}
function strikethrough(bgHex) {
  return { backgroundColor: hex(bgHex), textFormat: { strikethrough: true, foregroundColor: hex(COLOR.midGrey) } };
}

(async () => {
  const reqs = [];

  // ── ASSIGNMENTS TAB ──────────────────────────────────────────────────────────

  // 1. Overdue (status Not Started/In Progress AND due date < today) — whole row red
  reqs.push(cfRule(
    [gridRange(SID_ASSIGN, 4, 204, 0, 10)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND(OR($H5="Not Started",$H5="In Progress"),$E5<TODAY(),$E5<>"")' }] },
    bgText(COLOR.paleRed, COLOR.alertRed),
  ));

  // 2. Due within 3 days — amber highlight
  reqs.push(cfRule(
    [gridRange(SID_ASSIGN, 4, 204, 0, 10)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND(OR($H5="Not Started",$H5="In Progress"),$E5>=TODAY(),$E5<=TODAY()+3,$E5<>"")' }] },
    bgText(COLOR.paleGold, COLOR.amber),
  ));

  // 3. Complete / Submitted — green + strikethrough on title col
  reqs.push(cfRule(
    [gridRange(SID_ASSIGN, 4, 204, 1, 2)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=OR($H5="Complete",$H5="Submitted")' }] },
    strikethrough(COLOR.paleMint),
  ));
  reqs.push(cfRule(
    [gridRange(SID_ASSIGN, 4, 204, 0, 10)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=OR($H5="Complete",$H5="Submitted")' }] },
    bgText(COLOR.paleMint, COLOR.emerald),
  ));

  // 4. Status cell colours — High priority
  reqs.push(cfRule(
    [gridRange(SID_ASSIGN, 4, 204, 5, 6)],
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'High' }] },
    bgText(COLOR.paleRed, COLOR.alertRed, true),
  ));
  reqs.push(cfRule(
    [gridRange(SID_ASSIGN, 4, 204, 5, 6)],
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Medium' }] },
    bgText(COLOR.paleGold, COLOR.amber, false),
  ));
  reqs.push(cfRule(
    [gridRange(SID_ASSIGN, 4, 204, 5, 6)],
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Low' }] },
    bgText(COLOR.paleMint, COLOR.emerald, false),
  ));

  // 5. Course cell colours
  const courseColors = [
    ['English 101',   COLOR.paleSkyBlue,  COLOR.darkBlue],
    ['Calculus II',   COLOR.paleMint,     COLOR.emerald],
    ['World History', COLOR.paleLavender, COLOR.violet],
    ['Biology 201',   COLOR.palePeach,    COLOR.orange],
    ['Economics 101', COLOR.paleGold,     COLOR.amber],
  ];
  courseColors.forEach(([course, bgC, fgC]) => {
    reqs.push(cfRule(
      [gridRange(SID_ASSIGN, 4, 204, 2, 3)],
      { type: 'TEXT_EQ', values: [{ userEnteredValue: course }] },
      bgText(bgC, fgC),
    ));
  });

  // ── WEEKLY PLANNER TAB ──────────────────────────────────────────────────────

  // Course cells in planner grid
  courseColors.forEach(([course, bgC]) => {
    reqs.push(cfRule(
      [gridRange(SID_WEEK, 4, 37, 1, 8)],
      { type: 'TEXT_EQ', values: [{ userEnteredValue: course }] },
      bg(bgC),
    ));
  });

  // Activity types
  const activityColors = [
    ['Study',          COLOR.paleSkyBlue],
    ['Group Study',    COLOR.paleLavender],
    ['Review Session', COLOR.paleMint],
    ['Office Hours',   COLOR.paleGold],
    ['Exercise',       COLOR.palePeach],
    ['Personal',       COLOR.softGrey],
    ['Break',          COLOR.softGrey],
    ['Reading',        COLOR.paleSkyBlue],
    ['Lab',            COLOR.palePeach],
  ];
  activityColors.forEach(([act, bgC]) => {
    reqs.push(cfRule(
      [gridRange(SID_WEEK, 4, 37, 1, 8)],
      { type: 'TEXT_EQ', values: [{ userEnteredValue: act }] },
      bg(bgC),
    ));
  });

  // ── GRADES & GPA TAB ────────────────────────────────────────────────────────

  // Letter grade colours on col Q (16, row 6-11)
  const gradeColorMap = [
    ['A',  COLOR.paleMint,     COLOR.emerald],
    ['A-', COLOR.paleMint,     COLOR.emerald],
    ['B+', COLOR.paleSkyBlue,  COLOR.darkBlue],
    ['B',  COLOR.paleSkyBlue,  COLOR.darkBlue],
    ['B-', COLOR.paleSkyBlue,  COLOR.darkBlue],
    ['C+', COLOR.paleGold,     COLOR.amber],
    ['C',  COLOR.paleGold,     COLOR.amber],
    ['C-', COLOR.paleGold,     COLOR.amber],
    ['D+', COLOR.palePeach,    COLOR.orange],
    ['D',  COLOR.palePeach,    COLOR.orange],
    ['F',  COLOR.paleRed,      COLOR.alertRed],
  ];
  gradeColorMap.forEach(([grade, bgC, fgC]) => {
    reqs.push(cfRule(
      [gridRange(SID_GRADES, 6, 11, 16, 17)],
      { type: 'TEXT_EQ', values: [{ userEnteredValue: grade }] },
      bgText(bgC, fgC, true),
    ));
  });

  // Current % colour gradient on col O (14)
  reqs.push(cfRule(
    [gridRange(SID_GRADES, 6, 11, 14, 15)],
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '90' }] },
    bg(COLOR.paleMint),
  ));
  reqs.push(cfRule(
    [gridRange(SID_GRADES, 6, 11, 14, 15)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND(O7>=80,O7<90)' }] },
    bg(COLOR.paleSkyBlue),
  ));
  reqs.push(cfRule(
    [gridRange(SID_GRADES, 6, 11, 14, 15)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND(O7>=70,O7<80)' }] },
    bg(COLOR.paleGold),
  ));
  reqs.push(cfRule(
    [gridRange(SID_GRADES, 6, 11, 14, 15)],
    { type: 'NUMBER_LESS', values: [{ userEnteredValue: '70' }] },
    bg(COLOR.paleRed),
  ));

  // ── STUDENT BUDGET TAB ──────────────────────────────────────────────────────

  // Over budget: actual > budgeted in data rows (expenses rows 14-26)
  reqs.push(cfRule(
    [gridRange(SID_BUDGET, 13, 27, 0, 6)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$D14>$C14' }] },
    bgText(COLOR.paleRed, COLOR.alertRed),
  ));

  // Under budget: actual < budgeted
  reqs.push(cfRule(
    [gridRange(SID_BUDGET, 13, 27, 0, 6)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($D14<$C14,$C14>0)' }] },
    bgText(COLOR.paleMint, COLOR.emerald),
  ));

  // On track: actual = budgeted
  reqs.push(cfRule(
    [gridRange(SID_BUDGET, 13, 27, 0, 6)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$D14=$C14' }] },
    bgText(COLOR.paleSkyBlue, COLOR.darkBlue),
  ));

  // ── EXAM PREP TAB ───────────────────────────────────────────────────────────

  // Exams in < 7 days — red alert
  reqs.push(cfRule(
    [gridRange(SID_EXAM, 5, 10, 0, 10)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND(ISNUMBER($F6),$F6<=7,$F6>0)' }] },
    bgText(COLOR.paleRed, COLOR.alertRed),
  ));

  // Exams in 8-14 days — amber
  reqs.push(cfRule(
    [gridRange(SID_EXAM, 5, 10, 0, 10)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND(ISNUMBER($F6),$F6>=8,$F6<=14)' }] },
    bgText(COLOR.paleGold, COLOR.amber),
  ));

  // Confidence low (1-2) — red
  reqs.push(cfRule(
    [gridRange(SID_EXAM, 5, 10, 4, 5)],
    { type: 'NUMBER_LESS_THAN_EQ', values: [{ userEnteredValue: '2' }] },
    bgText(COLOR.paleRed, COLOR.alertRed, true),
  ));

  // Confidence medium (3) — amber
  reqs.push(cfRule(
    [gridRange(SID_EXAM, 5, 10, 4, 5)],
    { type: 'NUMBER_EQ', values: [{ userEnteredValue: '3' }] },
    bgText(COLOR.paleGold, COLOR.amber, false),
  ));

  // Confidence high (4-5) — green
  reqs.push(cfRule(
    [gridRange(SID_EXAM, 5, 10, 4, 5)],
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '4' }] },
    bgText(COLOR.paleMint, COLOR.emerald, false),
  ));

  // Checklist: Done checkbox (col H = 7) — green strikethrough on task col
  // Checklist rows start at various positions; apply broadly to rows 12-60
  reqs.push(cfRule(
    [gridRange(SID_EXAM, 11, 65, 1, 2)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$H12=TRUE' }] },
    strikethrough(COLOR.paleMint),
  ));

  await batchUpdate(id, reqs, 'conditional-formatting');
  console.log('Conditional formatting complete —', reqs.length, 'rules applied');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
