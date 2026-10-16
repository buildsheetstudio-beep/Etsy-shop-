'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const REF  = sheetMap['📋 Reference Data'];
const EXP  = sheetMap['🧾 Expense Log'];
const HSA  = sheetMap['💳 HSA/FSA Tracker'];
const INS  = sheetMap['📋 Insurance Claims'];
const PRV  = sheetMap['🏥 Provider Log'];
const RX   = sheetMap['💊 Prescription Tracker'];

function refRange(startRow, endRow, col) {
  return { sheetId: REF, startRowIndex: startRow, endRowIndex: endRow, startColumnIndex: col, endColumnIndex: col + 1 };
}

function dvRange(rule, ranges) {
  return {
    setDataValidation: {
      range: ranges,
      rule,
    },
  };
}

function oneOfRange(sheetId, startRow, endRow, col) {
  return {
    condition: {
      type: 'ONE_OF_RANGE',
      values: [{ userEnteredValue: `='📋 Reference Data'!$${String.fromCharCode(65 + col)}$${startRow + 1}:$${String.fromCharCode(65 + col)}$${endRow}` }],
    },
    strict: true,
    showCustomUi: true,
  };
}

function oneOfList(items) {
  return {
    condition: {
      type: 'ONE_OF_LIST',
      values: items.map(v => ({ userEnteredValue: v })),
    },
    strict: true,
    showCustomUi: true,
  };
}

function boolCheck() {
  return { condition: { type: 'BOOLEAN' }, strict: true };
}

(async () => {
  const reqs = [];

  // ── Expense Log (EXP) ───────────────────────────────────────────────────
  // Col B (index 1): Participant — Ref A2:A5 (rows 1-4, col 0)
  reqs.push(dvRange(
    oneOfRange(REF, 1, 5, 0),
    gridRange(EXP, 2, 500, 1, 2),
  ));
  // Col C (index 2): Category — Ref B2:B12 (rows 1-11, col 1)
  reqs.push(dvRange(
    oneOfRange(REF, 1, 12, 1),
    gridRange(EXP, 2, 500, 2, 3),
  ));
  // Col F (index 5): Payment Method — Ref C2:C8 (rows 1-7, col 2)
  reqs.push(dvRange(
    oneOfRange(REF, 1, 8, 2),
    gridRange(EXP, 2, 500, 5, 6),
  ));
  // Col H (index 7): Claim Status — Ref D2:D7 (rows 1-6, col 3)
  reqs.push(dvRange(
    oneOfRange(REF, 1, 7, 3),
    gridRange(EXP, 2, 500, 7, 8),
  ));

  // ── HSA/FSA Tracker (HSA) ───────────────────────────────────────────────
  // Col A (index 0): Account Type — ONE_OF_LIST
  reqs.push(dvRange(
    oneOfList(['HSA Self-Only', 'HSA Family', 'FSA']),
    gridRange(HSA, 2, 50, 0, 1),
  ));
  // Col B (index 1): Tax Year — Ref E2:E4 (rows 1-3, col 4)
  reqs.push(dvRange(
    oneOfRange(REF, 1, 4, 4),
    gridRange(HSA, 2, 50, 1, 2),
  ));

  // ── Insurance Claims (INS) ──────────────────────────────────────────────
  // Col D (index 3): Participant — Ref A2:A5
  reqs.push(dvRange(
    oneOfRange(REF, 1, 5, 0),
    gridRange(INS, 2, 500, 3, 4),
  ));
  // Col H (index 7): Status — Ref I2:I6 (rows 1-5, col 8)
  reqs.push(dvRange(
    oneOfRange(REF, 1, 6, 8),
    gridRange(INS, 2, 500, 7, 8),
  ));

  // ── Provider Log (PRV) ──────────────────────────────────────────────────
  // Col D (index 3): Participant — Ref A2:A5
  reqs.push(dvRange(
    oneOfRange(REF, 1, 5, 0),
    gridRange(PRV, 2, 500, 3, 4),
  ));
  // Col F (index 5): Follow-Up Needed — BOOLEAN checkbox
  reqs.push(dvRange(
    boolCheck(),
    gridRange(PRV, 2, 500, 5, 6),
  ));

  // ── Prescription Tracker (RX) ───────────────────────────────────────────
  // Col B (index 1): Participant — Ref A2:A5
  reqs.push(dvRange(
    oneOfRange(REF, 1, 5, 0),
    gridRange(RX, 2, 500, 1, 2),
  ));
  // Col J (index 9): Insurance Covered — BOOLEAN checkbox
  reqs.push(dvRange(
    boolCheck(),
    gridRange(RX, 2, 500, 9, 10),
  ));

  await batchUpdate(id, reqs, 'validation');
  console.log('Data validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
