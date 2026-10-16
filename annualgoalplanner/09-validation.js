'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GL  = sheetMap['🎯 Goals & Milestones'];
const ACT = sheetMap['📋 Action Steps'];
const WKL = sheetMap['📅 Weekly Check-In & Reflection'];
const REF = sheetMap['📋 Reference Data'];

function dv(sheetId, r1, r2, c1, c2, condition, strict = true) {
  return { setDataValidation: {
    range: { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 },
    rule: { condition, showCustomUi: true, strict },
  }};
}

function rangeCondition(sheetId, r1, r2, c1, c2) {
  return {
    type: 'ONE_OF_RANGE',
    values: [{ userEnteredValue: `='📋 Reference Data'!${colLetter(c1)}${r1+1}:${colLetter(c2-1)}${r2}` }],
  };
}

function colLetter(i) {
  return String.fromCharCode(65 + i);
}

function refRange(col, rows) {
  // col: 'A'-'I', rows: e.g. '3:12'
  return { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='📋 Reference Data'!${col}${rows}` }] };
}

function boolCond() {
  return { type: 'BOOLEAN' };
}

function numBetween(min, max) {
  return {
    type: 'NUMBER_BETWEEN',
    values: [{ userEnteredValue: String(min) }, { userEnteredValue: String(max) }],
  };
}

(async () => {
  const reqs = [];

  // ── GOALS & MILESTONES tab ──
  // Goal header rows (0-indexed): 2,8,14,20,26,32,38,44,50,56 — 10 rows
  // Milestone rows: all others in 3-61 range
  // For simplicity, apply to the entire data block and let Reference drive valid values

  // Col B (Life Area) — goal headers only: rows 2-61 col B (index 1) — apply to all, harmless on milestone rows
  reqs.push(dv(GL, 2, 62, 1, 2, refRange('A', '3:12')));

  // Col F (Goal Status) — goal header rows
  reqs.push(dv(GL, 2, 62, 5, 6, refRange('B', '3:9')));

  // Col G (Priority) — goal header rows
  reqs.push(dv(GL, 2, 62, 6, 7, refRange('C', '3:5')));

  // Col H on goal header rows: NUMBER_BETWEEN 0 100
  // Apply to full H column data range; milestone H values are checkbox, but validation on goal rows
  // Use number_between — milestones will show error if value entered, which is acceptable
  const GOAL_ROWS_0IDX = [2,8,14,20,26,32,38,44,50,56];
  GOAL_ROWS_0IDX.forEach(ri => {
    reqs.push(dv(GL, ri, ri+1, 7, 8, numBetween(0, 100)));
  });

  // Col F on milestone rows (status): Ref D col (Milestone Status)
  // Milestone rows are every non-goal row in 2-62 range
  // Apply to rows 3-62 (0-indexed 2-62) col F broadly; goal rows already have Goal Status
  // Better: apply to milestone sub-rows only
  for (let gi = 0; gi < 10; gi++) {
    const headerRi = GOAL_ROWS_0IDX[gi];
    for (let mi = 1; mi <= 5; mi++) {
      const msRi = headerRi + mi;
      reqs.push(dv(GL, msRi, msRi+1, 5, 6, refRange('D', '3:6')));
    }
  }

  // Col H (Done checkbox) on milestone rows
  GOAL_ROWS_0IDX.forEach(headerRi => {
    for (let mi = 1; mi <= 5; mi++) {
      const msRi = headerRi + mi;
      reqs.push(dv(GL, msRi, msRi+1, 7, 8, boolCond()));
    }
  });

  // ── ACTION STEPS tab ──
  // Data rows 4-34 (0-indexed; GSheets 5-35) — 30 rows
  // Col A (Goal name): Ref I col (I3:I12)
  reqs.push(dv(ACT, 4, 34, 0, 1, refRange('I', '3:12')));

  // Col E (Priority): Ref C col
  reqs.push(dv(ACT, 4, 34, 4, 5, refRange('C', '3:5')));

  // Col F (Action Status): Ref E col
  reqs.push(dv(ACT, 4, 34, 5, 6, refRange('E', '3:7')));

  // Col G (Done?): checkbox
  reqs.push(dv(ACT, 4, 34, 6, 7, boolCond()));

  // ── WEEKLY CHECK-IN tab ──
  // Data rows 2-53 (0-indexed; GSheets 3-54) — 52 rows
  // Col H (Energy 1-5): Ref F col
  reqs.push(dv(WKL, 2, 54, 7, 8, refRange('F', '3:7')));

  // Col I (Motivation 1-5): Ref F col
  reqs.push(dv(WKL, 2, 54, 8, 9, refRange('F', '3:7')));

  // Col J (Week Rating 1-5): Ref F col
  reqs.push(dv(WKL, 2, 54, 9, 10, refRange('F', '3:7')));

  await batchUpdate(id, reqs, 'validation');
  console.log('Data Validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
