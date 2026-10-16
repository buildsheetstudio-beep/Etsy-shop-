'use strict';
const { batchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const WKL = sheetMap['📝 Workout Log'];
const REC = sheetMap['😴 Rest & Recovery'];
const PRT = sheetMap['🏆 PR Tracker'];
const CAR = sheetMap['🏃 Cardio Tracker'];

function cf(ranges, rule) {
  return { addConditionalFormatRule: { rule: { ranges, booleanRule: rule }, index: 0 } };
}

(async () => {
  const reqs = [];

  // ── 1. Workout Log: Completed = TRUE → light green row ──────────────────
  reqs.push(cf([gridRange(WKL, 2, 32, 0, 11)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$J3=TRUE' }] },
    format: { backgroundColor: hex(C.lightGreen) },
  }));

  // ── 2. Workout Log: Planned (date exists, not completed) → light amber ───
  reqs.push(cf([gridRange(WKL, 2, 32, 0, 11)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($A3<>"",$J3=FALSE)' }] },
    format: { backgroundColor: hex(C.lightAmber) },
  }));

  // ── 3. Recovery: Sleep < 6hrs → light coral ───────────────────────────────
  reqs.push(cf([gridRange(REC, 2, 32, 1, 2)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$B3<6' }] },
    format: { backgroundColor: hex(C.lightCoral), textFormat: { bold: true } },
  }));

  // ── 4. Recovery: Sleep >= 8hrs → light green ─────────────────────────────
  reqs.push(cf([gridRange(REC, 2, 32, 1, 2)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$B3>=8' }] },
    format: { backgroundColor: hex(C.lightGreen) },
  }));

  // ── 5. Recovery: Sleep Quality = "Excellent" → green cell ────────────────
  reqs.push(cf([gridRange(REC, 2, 32, 2, 3)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$C3="Excellent"' }] },
    format: { backgroundColor: hex(C.lightGreen), textFormat: { foregroundColor: hex(C.teal) } },
  }));

  // ── 6. Recovery: Sleep Quality = "Poor" → coral cell ─────────────────────
  reqs.push(cf([gridRange(REC, 2, 32, 2, 3)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$C3="Poor"' }] },
    format: { backgroundColor: hex(C.lightCoral), textFormat: { foregroundColor: hex(C.coral) } },
  }));

  // ── 7. Recovery: Soreness = "Severe" → coral cell ────────────────────────
  reqs.push(cf([gridRange(REC, 2, 32, 3, 4)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$D3="Severe"' }] },
    format: { backgroundColor: hex(C.lightCoral), textFormat: { foregroundColor: hex(C.coral), bold: true } },
  }));

  // ── 8. Recovery: Soreness = "None" → light green cell ────────────────────
  reqs.push(cf([gridRange(REC, 2, 32, 3, 4)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$D3="None"' }] },
    format: { backgroundColor: hex(C.lightGreen) },
  }));

  // ── 9. Recovery: Rest Day = TRUE → full row light blue ───────────────────
  reqs.push(cf([gridRange(REC, 2, 32, 0, 7)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F3=TRUE' }] },
    format: { backgroundColor: hex(C.lightBlue) },
  }));

  // ── 10. PR Tracker: PR > 0 (achieved) → light green ──────────────────────
  reqs.push(cf([gridRange(PRT, 2, 14, 1, 2)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$B3>0' }] },
    format: { backgroundColor: hex(C.lightGreen), textFormat: { foregroundColor: hex(C.teal), bold: true } },
  }));

  // ── 11. PR Tracker: PR >= Target → gold/amber highlight ───────────────────
  reqs.push(cf([gridRange(PRT, 2, 14, 0, 6)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($B3>0,$E3>0,$B3>=$E3)' }] },
    format: { backgroundColor: hex(C.lightAmber), textFormat: { foregroundColor: hex(C.amber), bold: true } },
  }));

  // ── 12. Cardio Tracker: Pace < 9 min/mi (fast run) → light green ─────────
  reqs.push(cf([gridRange(CAR, 2, 12, 4, 5)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($E3>0,$E3<9,$B3="Running")' }] },
    format: { backgroundColor: hex(C.lightGreen), textFormat: { foregroundColor: hex(C.teal) } },
  }));

  await batchUpdate(id, reqs, 'cf');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
