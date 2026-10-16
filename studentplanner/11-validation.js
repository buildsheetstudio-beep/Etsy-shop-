'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID_ASSIGN = sheetMap['Assignments'];
const SID_GRADES = sheetMap['Grades & GPA'];
const SID_BUDGET = sheetMap['Student Budget'];
const SID_EXAM   = sheetMap['Exam Prep'];
const SID_REF    = sheetMap['Reference'];

function dropdown(sheetId, r0, r1, c0, c1, refRange, strict = true) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r0, r1, c0, c1),
      rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${refRange}` }] },
        strict,
        showCustomUi: true,
      },
    },
  };
}

function inlineList(sheetId, r0, r1, c0, c1, options, strict = true) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r0, r1, c0, c1),
      rule: {
        condition: { type: 'ONE_OF_LIST', values: options.map(v => ({ userEnteredValue: v })) },
        strict,
        showCustomUi: true,
      },
    },
  };
}

function checkbox(sheetId, r0, r1, c0, c1) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r0, r1, c0, c1),
      rule: { condition: { type: 'BOOLEAN' }, strict: true, showCustomUi: true },
    },
  };
}

function numberBetween(sheetId, r0, r1, c0, c1, min, max) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r0, r1, c0, c1),
      rule: {
        condition: { type: 'NUMBER_BETWEEN', values: [{ userEnteredValue: String(min) }, { userEnteredValue: String(max) }] },
        strict: false, showCustomUi: false,
      },
    },
  };
}

(async () => {
  const reqs = [];

  // ── ASSIGNMENTS TAB ──────────────────────────────────────────────────────────
  // Col C (2): Course — Reference!A2:A6
  reqs.push(dropdown(SID_ASSIGN, 4, 204, 2, 3, `'Reference'!$A$2:$A$6`));
  // Col D (3): Type — Reference!B2:B11
  reqs.push(dropdown(SID_ASSIGN, 4, 204, 3, 4, `'Reference'!$B$2:$B$11`));
  // Col F (5): Priority — Reference!D2:D4
  reqs.push(dropdown(SID_ASSIGN, 4, 204, 5, 6, `'Reference'!$D$2:$D$4`));
  // Col H (7): Status — Reference!C2:C6
  reqs.push(dropdown(SID_ASSIGN, 4, 204, 7, 8, `'Reference'!$C$2:$C$6`));

  // ── EXAM PREP TAB ────────────────────────────────────────────────────────────
  // Col A (0): Course — Reference!A2:A6
  reqs.push(dropdown(SID_EXAM, 5, 10, 0, 1, `'Reference'!$A$2:$A$6`));
  // Col E (4): Confidence 1-5
  reqs.push(numberBetween(SID_EXAM, 5, 10, 4, 5, 1, 5));
  // Exam Prep checklist Done column (col H = 7): rows 12-65
  reqs.push(checkbox(SID_EXAM, 12, 65, 7, 8));

  // ── GRADES & GPA TAB ─────────────────────────────────────────────────────────
  // Score cols D, F, H, J, L (cols 3,5,7,9,11): NUMBER_BETWEEN 0-100
  [3, 5, 7, 9, 11].forEach(ci => {
    reqs.push(numberBetween(SID_GRADES, 6, 11, ci, ci + 1, 0, 100));
  });
  // Credits col B (1): NUMBER_BETWEEN 1-6
  reqs.push(numberBetween(SID_GRADES, 6, 11, 1, 2, 1, 6));

  // ── STUDENT BUDGET TAB ───────────────────────────────────────────────────────
  // Col A (0): Category — Reference!H2:H12
  reqs.push(dropdown(SID_BUDGET, 5, 36, 0, 1, `'Reference'!$H$2:$H$12`));
  // Col E (4): Frequency — inline list
  reqs.push(inlineList(SID_BUDGET, 5, 36, 4, 5, ['Monthly', 'Weekly', 'Bi-weekly', 'Semester', 'Annual', 'One-time']));

  await batchUpdate(id, reqs, 'data-validation');
  console.log('Data validation complete —', reqs.length, 'rules applied');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
