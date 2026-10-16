'use strict';
const { batchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const EXP  = sheetMap['🧾 Expense Log'];
const INS  = sheetMap['📋 Insurance Claims'];
const HSA  = sheetMap['💳 HSA/FSA Tracker'];
const RX   = sheetMap['💊 Prescription Tracker'];
const DASH = sheetMap['📊 Dashboard'];

function cfRule(ranges, formula, bg, textColor) {
  const rule = {
    addConditionalFormatRule: {
      rule: {
        ranges,
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: formula }] },
          format: { backgroundColor: hex(bg) },
        },
      },
      index: 0,
    },
  };
  if (textColor) {
    rule.addConditionalFormatRule.rule.booleanRule.format.textFormat = { foregroundColor: hex(textColor) };
  }
  return rule;
}

(async () => {
  const reqs = [];

  // ── Expense Log col H (index 7): Status CF on full row ──────────────────
  // Paid → light sage
  reqs.push(cfRule(
    [gridRange(EXP, 2, 200, 0, 9)],
    '=$H3="Paid"',
    C.lightSage,
  ));
  // Submitted → light amber
  reqs.push(cfRule(
    [gridRange(EXP, 2, 200, 0, 9)],
    '=$H3="Submitted"',
    C.lightAmber,
  ));
  // Processing → light amber
  reqs.push(cfRule(
    [gridRange(EXP, 2, 200, 0, 9)],
    '=$H3="Processing"',
    C.lightAmber,
  ));
  // Approved → light sage (lighter shade)
  reqs.push(cfRule(
    [gridRange(EXP, 2, 200, 0, 9)],
    '=$H3="Approved"',
    C.lightSage,
  ));
  // Denied → light red
  reqs.push(cfRule(
    [gridRange(EXP, 2, 200, 0, 9)],
    '=$H3="Denied"',
    C.lightRed,
    C.brickRed,
  ));
  // Not Submitted → light gray
  reqs.push(cfRule(
    [gridRange(EXP, 2, 200, 0, 9)],
    '=$H3="Not Submitted"',
    C.lightGray,
  ));

  // ── Insurance Claims col H (index 7): same status CF on full row ─────────
  reqs.push(cfRule(
    [gridRange(INS, 2, 200, 0, 10)],
    '=$H3="Paid"',
    C.lightSage,
  ));
  reqs.push(cfRule(
    [gridRange(INS, 2, 200, 0, 10)],
    '=$H3="Submitted"',
    C.lightAmber,
  ));
  reqs.push(cfRule(
    [gridRange(INS, 2, 200, 0, 10)],
    '=$H3="Processing"',
    C.lightAmber,
  ));
  reqs.push(cfRule(
    [gridRange(INS, 2, 200, 0, 10)],
    '=$H3="Approved"',
    C.lightSage,
  ));
  reqs.push(cfRule(
    [gridRange(INS, 2, 200, 0, 10)],
    '=$H3="Denied"',
    C.lightRed,
    C.brickRed,
  ));

  // ── HSA/FSA Tracker col H (index 7): negative balance → red ────────────
  reqs.push(cfRule(
    [gridRange(HSA, 2, 20, 7, 8)],
    '=$H3<0',
    C.lightRed,
    C.brickRed,
  ));
  // Also highlight full row if balance is negative
  reqs.push(cfRule(
    [gridRange(HSA, 2, 20, 0, 9)],
    '=$H3<0',
    C.lightRed,
  ));

  // ── Prescription Tracker col G (Next Refill Due): CF ───────────────────
  // Overdue: G < TODAY() → red
  reqs.push(cfRule(
    [gridRange(RX, 2, 200, 0, 10)],
    '=AND($G3<>"", $G3<TODAY())',
    C.lightRed,
    C.brickRed,
  ));
  // Within 7 days: G - TODAY() <= 7 AND G >= TODAY() → amber
  reqs.push(cfRule(
    [gridRange(RX, 2, 200, 0, 10)],
    '=AND($G3<>"", $G3>=TODAY(), $G3-TODAY()<=7)',
    C.lightAmber,
  ));

  // ── Dashboard OOP progress stat (E5 = index row 4, col 4): red if ≥ OOP max ──
  // G5 cell = row 4, col 6 (0-indexed): % progress
  reqs.push(cfRule(
    [gridRange(DASH, 4, 5, 6, 7)],
    '=AND(F3>0,E5>=F3)',
    C.lightRed,
    C.brickRed,
  ));
  // Also flag the OOP value cell (E5 = row 4, col 4)
  reqs.push(cfRule(
    [gridRange(DASH, 4, 5, 4, 5)],
    '=AND(F3>0,E5>=F3)',
    C.lightRed,
    C.brickRed,
  ));

  await batchUpdate(id, reqs, 'cf-rules');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
