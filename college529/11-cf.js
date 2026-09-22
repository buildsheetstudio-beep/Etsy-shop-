'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID_BS = sheetMap['Beneficiary Setup'];
const SID_CA = sheetMap['529 Accounts'];
const SID_CL = sheetMap['Contribution Log'];
const SID_CP = sheetMap['Contribution Planner'];
const SID_GM = sheetMap['Goals & Milestones'];

// Build a TEXT_EQ addConditionalFormatRule request
const cfEq = (sheetId, r1, r2, c1, c2, text, bgColor) => ({
  addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(sheetId, r1, r2, c1, c2)],
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] },
        format: { backgroundColor: hex(bgColor) }
      }
    },
    index: 0
  }
});

(async () => {
  const fmt = [];

  // ── Beneficiary Setup: Q col (16) — Beneficiary Status ───────────────────
  [
    ['Enrolled',        C.success   ],
    ['Graduated',       C.success   ],
    ['Near Enrollment', C.info      ],
    ['Saving',          C.eucalTint ],
    ['Planning',        C.aubergTint],
    ['Paused',          C.altRow    ],
    ['Archived',        C.altRow    ],
  ].forEach(([text, bg]) => fmt.push(cfEq(SID_BS, 7, 507, 16, 17, text, bg)));

  // ── 529 Accounts: G col (6) — Account Status ─────────────────────────────
  [
    ['Active',      C.eucalTint ],
    ['Paused',      C.warning   ],
    ['Closed',      C.altRow    ],
    ['Transferred', C.aubergTint],
    ['Archived',    C.altRow    ],
  ].forEach(([text, bg]) => fmt.push(cfEq(SID_CA, 7, 1007, 6, 7, text, bg)));

  // ── Contribution Log: F col (5) — Transaction Type ───────────────────────
  [
    ['Contribution',        C.eucalTint  ],
    ['Withdrawal',          C.copperTint ],
    ['Transfer In',         C.aubergTint ],
    ['Transfer Out',        C.aubergTint ],
    ['Adjustment Increase', C.eucalTint  ],
    ['Adjustment Decrease', C.copperTint ],
  ].forEach(([text, bg]) => fmt.push(cfEq(SID_CL, 5, 5005, 5, 6, text, bg)));

  // ── Goals & Milestones: G col (6) — Goal Status ──────────────────────────
  [
    ['Achieved',     C.success   ],
    ['Active',       C.info      ],
    ['On Track',     C.eucalTint ],
    ['Ahead of Plan',C.eucalTint ],
    ['Behind Plan',  C.warning   ],
    ['Not Started',  C.altRow    ],
    ['Paused',       C.altRow    ],
    ['Delayed',      C.warning   ],
    ['Reassess',     C.copperTint],
  ].forEach(([text, bg]) => fmt.push(cfEq(SID_GM, 7, 1600, 6, 7, text, bg)));

  // ── Contribution Planner: J col (9) — Planner Status ─────────────────────
  [
    ['Goal Reached!',     C.success   ],
    ['Ahead of Plan',     C.eucalTint ],
    ['On Track',          C.eucalTint ],
    ['Behind Plan',       C.warning   ],
    ['Insufficient Data', C.altRow    ],
  ].forEach(([text, bg]) => fmt.push(cfEq(SID_CP, 7, 12, 9, 10, text, bg)));

  await batchUpdate(id, fmt, '11-cf');
  console.log('11-cf done ✓');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
