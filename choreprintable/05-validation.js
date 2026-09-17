'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TAB = sheetMap['Weekly Chore Chart'];

// Task desc left cols (C,F,I,L,O,R,U) — 0-indexed: 2,5,8,11,14,17,20
const TASK_COLS = [2,5,8,11,14,17,20];
// Checkbox cols (B,E,H,K,N,Q,T) — 0-indexed: 1,4,7,10,13,16,19
const CHK_COLS  = [1,4,7,10,13,16,19];

(async () => {
  const reqs = [];

  // ── Checkboxes for all 7 days × rows 14:37 ───────────────────────────────
  for (const ci of CHK_COLS) {
    reqs.push({ setDataValidation: {
      range: gridRange(TAB, 13, 37, ci, ci+1),  // rows 14-37 (0-indexed 13-36)
      rule: {
        condition: { type: 'BOOLEAN' },
        strict: true,
        showCustomUi: true,
      }
    }});
  }

  // ── Chore dropdowns for task desc cells (C,F,I,L,O,R,U), rows 14:37 ──────
  for (const ci of TASK_COLS) {
    reqs.push({ setDataValidation: {
      range: gridRange(TAB, 13, 37, ci, ci+1),
      rule: {
        condition: {
          type: 'ONE_OF_RANGE',
          values: [{ userEnteredValue: "='Reference Data'!$F$2:$F$31" }]
        },
        strict: false,       // allow custom entries
        showCustomUi: true,
      }
    }});
  }

  // ── Date validation for week start input G6:J7 (0-indexed rows 5-6, cols 6-9) ──
  reqs.push({ setDataValidation: {
    range: gridRange(TAB, 5, 7, 6, 10),
    rule: {
      condition: { type: 'DATE_IS_VALID' },
      strict: false,
      showCustomUi: true,
    }
  }});

  await batchUpdate(id, reqs, 'validation');
  console.log('Validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
