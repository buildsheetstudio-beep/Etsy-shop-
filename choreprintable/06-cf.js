'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TAB = sheetMap['Weekly Chore Chart'];

// Pairs: [chkCol(0-idx), descCol(0-idx), descEndCol(0-idx)] for each day
// Mon: B=1 → C:D(2:4), Tue: E=4 → F:G(5:7), Wed: H=7 → I:J(8:10)
// Thu: K=10 → L:M(11:13), Fri: N=13 → O:P(14:16)
// Sat: Q=16 → R:S(17:19), Sun: T=19 → U:V(20:22)
const DAY_PAIRS = [
  { chk: 1, desc: 2, descEnd: 4, chkLetter: 'B' },
  { chk: 4, desc: 5, descEnd: 7, chkLetter: 'E' },
  { chk: 7, desc: 8, descEnd: 10, chkLetter: 'H' },
  { chk: 10, desc: 11, descEnd: 13, chkLetter: 'K' },
  { chk: 13, desc: 14, descEnd: 16, chkLetter: 'N' },
  { chk: 16, desc: 17, descEnd: 19, chkLetter: 'Q' },
  { chk: 19, desc: 20, descEnd: 22, chkLetter: 'T' },
];

(async () => {
  const reqs = [];
  let idx = 0;

  // Rule 1 — Weekend tint (lowest priority, added first)
  // Saturday Q:S, Sunday T:V — rows 14:37
  for (const col of [[16,19],[19,22]]) {
    reqs.push({ addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(TAB, 13, 37, col[0], col[1])],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=TRUE' }] },
          format: { backgroundColor: hex(C.weekend) }
        }
      },
      index: idx++
    }});
  }

  // Rule 2 — Completed task row (higher priority — added after weekend tint)
  // For each day: format desc+spacer cols when checkbox is TRUE
  for (const { chk, desc, descEnd, chkLetter } of DAY_PAIRS) {
    // Format the desc cells when the checkbox in same row is TRUE
    reqs.push({ addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(TAB, 13, 37, desc, descEnd)],
        booleanRule: {
          condition: {
            type: 'CUSTOM_FORMULA',
            values: [{ userEnteredValue: `=$${chkLetter}14=TRUE` }]
          },
          format: {
            backgroundColor: hex(C.completed),
            textFormat: { foregroundColor: hex(C.completedText) }
          }
        }
      },
      index: idx++
    }});
    // Format the checkbox cell itself
    reqs.push({ addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(TAB, 13, 37, chk, chk+1)],
        booleanRule: {
          condition: {
            type: 'CUSTOM_FORMULA',
            values: [{ userEnteredValue: `=$${chkLetter}14=TRUE` }]
          },
          format: { backgroundColor: hex(C.completed) }
        }
      },
      index: idx++
    }});
  }

  await batchUpdate(id, reqs, 'cf');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
