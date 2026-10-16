'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TAB = sheetMap['Chore & Allowance Tracker'];
const RWD = sheetMap['Reward & Prize Redemption'];
const REF = sheetMap['Reference Data'];

(async () => {
  const reqs = [];

  // ── Rate table Type column: H3:H8 → ONE_OF_RANGE from Reference!A2:A3 ────
  reqs.push({ setDataValidation: {
    range: gridRange(TAB, 2, 8, 7, 8),  // H3:H8 (0-indexed rows 2-7, col 7)
    rule: {
      condition: {
        type: 'ONE_OF_RANGE',
        values: [{ userEnteredValue: "='Reference Data'!$A$2:$A$3" }]
      },
      strict: true,
      showCustomUi: true,
    }
  } });

  // ── Reward tab Family Member col A2:A100 → from Chore tab I3:I12 ─────────
  reqs.push({ setDataValidation: {
    range: gridRange(RWD, 1, 100, 0, 1),  // A2:A100
    rule: {
      condition: {
        type: 'ONE_OF_RANGE',
        values: [{ userEnteredValue: "='Chore & Allowance Tracker'!$I$3:$I$12" }]
      },
      strict: false,
      showCustomUi: true,
    }
  } });

  // ── CF: Remaining Balance (col E) < 0 → red text on Reward tab ───────────
  // Conditional format rule on E2:E100
  reqs.push({ addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(RWD, 1, 100, 4, 5)],  // E2:E100
      booleanRule: {
        condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
        format: {
          textFormat: { foregroundColor: hex(C.rust), bold: true },
        }
      }
    },
    index: 0
  } });

  await batchUpdate(id, reqs, 'validation');
  console.log('Validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
