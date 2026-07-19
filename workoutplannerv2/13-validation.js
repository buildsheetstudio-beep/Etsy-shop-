'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const WL  = sheetMap['💪 Workout Log'];
const WDL = sheetMap['📝 Weekly Detail Log'];
const SS  = sheetMap['⚡ Superset & Circuit Builder'];
const PR  = sheetMap['🏆 Personal Records (PR) Tracker'];
const CRD = sheetMap['❤️ Cardio & Heart Rate Zone Tracker'];

function dropdown(range, refRange, strict = true) {
  return {
    setDataValidation: {
      range,
      rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: refRange }] },
        showCustomUi: true,
        strict,
      },
    },
  };
}

function checkbox(range) {
  return {
    setDataValidation: {
      range,
      rule: { condition: { type: 'BOOLEAN' }, strict: true, showCustomUi: true },
    },
  };
}

(async () => {
  const reqs = [];

  // ── Workout Log ───────────────────────────────────────────────────────────
  // Col B (idx 1): Muscle Group — Reference!A2:A10
  reqs.push(dropdown(gridRange(WL,1,31,1,2), "='📋 Reference Data'!$A$2:$A$10"));
  // Col C (idx 2): Workout Type — Reference!B2:B7
  reqs.push(dropdown(gridRange(WL,1,31,2,3), "='📋 Reference Data'!$B$2:$B$7"));
  // Col F (idx 5): Completed — checkbox
  reqs.push(checkbox(gridRange(WL,1,31,5,6)));

  // ── Weekly Detail Log ─────────────────────────────────────────────────────
  // Col D (idx 3): Muscle Group — Reference!A2:A10
  reqs.push(dropdown(gridRange(WDL,1,26,3,4), "='📋 Reference Data'!$A$2:$A$10"));

  // ── Superset & Circuit Builder ────────────────────────────────────────────
  // Col G (idx 6): Muscle Group Focus — Reference!A2:A10
  reqs.push(dropdown(gridRange(SS,1,21,6,7), "='📋 Reference Data'!$A$2:$A$10"));

  // ── PR Tracker ────────────────────────────────────────────────────────────
  // Col C (idx 2): PR Type — Reference!D2:D5
  reqs.push(dropdown(gridRange(PR,1,14,2,3), "='📋 Reference Data'!$D$2:$D$5"));

  // ── Cardio & HR Zone ─────────────────────────────────────────────────────
  // Col B (idx 1): Activity — Reference!C2:C7
  reqs.push(dropdown(gridRange(CRD,1,13,1,2), "='📋 Reference Data'!$C$2:$C$7"));

  await batchUpdate(id, reqs, 'validation');
  console.log('Data validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
