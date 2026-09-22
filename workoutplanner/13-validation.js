'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const WKL = sheetMap['📝 Workout Log'];
const PTL = sheetMap['📚 Program Templates'];
const CAR = sheetMap['🏃 Cardio Tracker'];
const REC = sheetMap['😴 Rest & Recovery'];

function dropdown(range, refRange) {
  return {
    setDataValidation: {
      range,
      rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: refRange }] },
        showCustomUi: true,
        strict: false,
      },
    },
  };
}

function checkbox(range) {
  return {
    setDataValidation: {
      range,
      rule: {
        condition: { type: 'BOOLEAN' },
        strict: true,
        showCustomUi: true,
      },
    },
  };
}

(async () => {
  const reqs = [];

  // ── Workout Log ──────────────────────────────────────────────────────────
  // Col C (idx 2): Muscle Group
  reqs.push(dropdown(gridRange(WKL, 2, 32, 2, 3), "='📋 Reference Data'!$A$2:$A$11"));
  // Col E (idx 4): Workout Type
  reqs.push(dropdown(gridRange(WKL, 2, 32, 4, 5), "='📋 Reference Data'!$B$2:$B$11"));
  // Col J (idx 9): Completed checkbox
  reqs.push(checkbox(gridRange(WKL, 2, 32, 9, 10)));

  // ── Program Templates ────────────────────────────────────────────────────
  // Col C (idx 2): Muscle Group
  reqs.push(dropdown(gridRange(PTL, 2, 17, 2, 3), "='📋 Reference Data'!$A$2:$A$11"));
  // Col A (idx 0): Program Name
  reqs.push(dropdown(gridRange(PTL, 2, 17, 0, 1), "='📋 Reference Data'!$C$2:$C$11"));

  // ── Cardio Tracker ────────────────────────────────────────────────────────
  // Col B (idx 1): Cardio Type
  reqs.push(dropdown(gridRange(CAR, 2, 12, 1, 2), "='📋 Reference Data'!$D$2:$D$7"));

  // ── Rest & Recovery ───────────────────────────────────────────────────────
  // Col C (idx 2): Sleep Quality
  reqs.push(dropdown(gridRange(REC, 2, 32, 2, 3), "='📋 Reference Data'!$E$2:$E$5"));
  // Col D (idx 3): Soreness Level
  reqs.push(dropdown(gridRange(REC, 2, 32, 3, 4), "='📋 Reference Data'!$F$2:$F$5"));
  // Col F (idx 5): Rest Day checkbox
  reqs.push(checkbox(gridRange(REC, 2, 32, 5, 6)));

  await batchUpdate(id, reqs, 'validation');
  console.log('Data validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
