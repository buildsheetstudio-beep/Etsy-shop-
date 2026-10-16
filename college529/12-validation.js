'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID_BS = sheetMap['Beneficiary Setup'];
const SID_CA = sheetMap['529 Accounts'];
const SID_CL = sheetMap['Contribution Log'];
const SID_GM = sheetMap['Goals & Milestones'];

const REF = "'Reference Data'";

// ONE_OF_RANGE dropdown helper
const dv = (sheetId, r1, r2, c1, c2, refRange) => ({
  setDataValidation: {
    range: gridRange(sheetId, r1, r2, c1, c2),
    rule: {
      condition: {
        type: 'ONE_OF_RANGE',
        values: [{ userEnteredValue: `=${REF}!${refRange}` }]
      },
      showCustomUi: true,
      strict: false
    }
  }
});

// Checkbox (BooleanCondition) helper
const cb = (sheetId, r1, r2, c1, c2) => ({
  setDataValidation: {
    range: gridRange(sheetId, r1, r2, c1, c2),
    rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true }
  }
});

(async () => {
  const fmt = [];

  // ── Beneficiary Setup ─────────────────────────────────────────────────────
  // C (2): Family Role  →  Reference Data col A (7 items, rows 3-9)
  fmt.push(dv(SID_BS, 7, 507, 2, 3, '$A$3:$A$9'));
  // J (9): College Type  →  Reference Data col H (7 items, rows 3-9)
  fmt.push(dv(SID_BS, 7, 507, 9, 10, '$H$3:$H$9'));
  // Q (16): Beneficiary Status  →  Reference Data col B (7 items, rows 3-9)
  fmt.push(dv(SID_BS, 7, 507, 16, 17, '$B$3:$B$9'));
  // Control J6 (row 5, col 9): Growth Scenario  →  Reference Data col J (3 items)
  fmt.push(dv(SID_BS, 5, 6, 9, 10, '$J$3:$J$5'));
  // Control N6 (row 5, col 13): Paycheck Frequency  →  Reference Data col G (7 items)
  fmt.push(dv(SID_BS, 5, 6, 13, 14, '$G$3:$G$9'));

  // ── 529 Accounts ─────────────────────────────────────────────────────────
  // D (3): Account Owner  →  Reference Data col D (6 items, rows 3-8)
  fmt.push(dv(SID_CA, 7, 1007, 3, 4, '$D$3:$D$8'));
  // F (5): Account Type  →  Reference Data col C (8 items, rows 3-10)
  fmt.push(dv(SID_CA, 7, 1007, 5, 6, '$C$3:$C$10'));
  // G (6): Account Status  →  Reference Data col O (5 items, rows 3-7)
  fmt.push(dv(SID_CA, 7, 1007, 6, 7, '$O$3:$O$7'));
  // W (22): Tax Advantaged? — checkbox
  fmt.push(cb(SID_CA, 7, 1007, 22, 23));

  // ── Contribution Log ─────────────────────────────────────────────────────
  // F (5): Transaction Type  →  Reference Data col F (7 items, rows 3-9)
  fmt.push(dv(SID_CL, 5, 5005, 5, 6, '$F$3:$F$9'));
  // G (6): Source  →  Reference Data col E (9 items, rows 3-11)
  fmt.push(dv(SID_CL, 5, 5005, 6, 7, '$E$3:$E$11'));
  // I (8): Frequency  →  Reference Data col G (7 items, rows 3-9)
  fmt.push(dv(SID_CL, 5, 5005, 8, 9, '$G$3:$G$9'));

  // ── Goals & Milestones ────────────────────────────────────────────────────
  // C (2): Milestone Type  →  Reference Data col L (10 items, rows 3-12)
  fmt.push(dv(SID_GM, 7, 1600, 2, 3, '$L$3:$L$12'));
  // G (6): Status  →  Reference Data col K (8 items, rows 3-10)
  fmt.push(dv(SID_GM, 7, 1600, 6, 7, '$K$3:$K$10'));

  await batchUpdate(id, fmt, '12-validation');
  console.log('12-validation done ✓');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
