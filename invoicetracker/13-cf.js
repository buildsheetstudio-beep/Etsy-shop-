'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const IL = sheetMap['Invoice Log'];
const PT = sheetMap['Payment Tracker'];
const IG = sheetMap['Invoice Generator'];
const DB = sheetMap['Dashboard'];

function cfTextEq(sheetId, r1, r2, c1, c2, text, bgHex, textHex) {
  return { addConditionalFormatRule: { rule: {
    ranges: [gridRange(sheetId, r1, r2, c1, c2)],
    booleanRule: {
      condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] },
      format: { backgroundColor: hex(bgHex), textFormat: { foregroundColor: hex(textHex || C.mainText) } },
    },
  }, index: 0 }};
}

function cfFormula(sheetId, r1, r2, c1, c2, formula, bgHex, textHex) {
  return { addConditionalFormatRule: { rule: {
    ranges: [gridRange(sheetId, r1, r2, c1, c2)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: formula }] },
      format: { backgroundColor: hex(bgHex), textFormat: { foregroundColor: hex(textHex || C.mainText) } },
    },
  }, index: 0 }};
}

(async () => {
  const reqs = [];

  // ── Invoice Log: Invoice Status (col T = index 19) ────────────────────────
  cfTextEq(IL, 5, 1005, 19, 20, 'Paid',           C.success,   '#1B3C1E');
  cfTextEq(IL, 5, 1005, 19, 20, 'Ready to Send',  C.mutedBlue, '#12293D');
  cfTextEq(IL, 5, 1005, 19, 20, 'Sent',           C.mutedBlue, '#12293D');
  cfTextEq(IL, 5, 1005, 19, 20, 'Viewed',         C.mutedBlue, '#12293D');
  cfTextEq(IL, 5, 1005, 19, 20, 'Partially Paid', C.warning,   C.mainText);
  cfTextEq(IL, 5, 1005, 19, 20, 'Overdue',        C.attention, C.white);
  cfTextEq(IL, 5, 1005, 19, 20, 'Draft',          C.neutral,   C.white);
  cfTextEq(IL, 5, 1005, 19, 20, 'Cancelled',      C.neutral,   C.white);
  cfTextEq(IL, 5, 1005, 19, 20, 'Refunded',       C.attention, C.white);
  [cfTextEq, cfTextEq, cfTextEq, cfTextEq, cfTextEq, cfTextEq, cfTextEq, cfTextEq, cfTextEq].forEach((f,i) => {});
  // Push all rules
  const invStatusRules = [
    ['Paid',           C.success,   '#1B3C1E'],
    ['Ready to Send',  C.mutedBlue, '#12293D'],
    ['Sent',           C.mutedBlue, '#12293D'],
    ['Viewed',         C.mutedBlue, '#12293D'],
    ['Partially Paid', C.warning,   C.mainText],
    ['Overdue',        C.attention, C.white],
    ['Draft',          C.neutral,   C.white],
    ['Cancelled',      C.neutral,   C.white],
    ['Refunded',       C.attention, C.white],
  ];
  invStatusRules.forEach(([text, bg, fg]) => {
    reqs.push(cfTextEq(IL, 5, 1005, 19, 20, text, bg, fg));
  });

  // ── Invoice Log: Payment Status (col S = index 18) ────────────────────────
  const payStatusRules = [
    ['Paid',           C.success,   '#1B3C1E'],
    ['Deposit Paid',   C.warning,   C.mainText],
    ['Partially Paid', C.warning,   C.mainText],
    ['Overpaid',       C.mutedBlue, '#12293D'],
    ['Unpaid',         C.neutral,   C.white],
    ['Refunded',       C.attention, C.white],
    ['Cancelled',      C.neutral,   C.white],
  ];
  payStatusRules.forEach(([text, bg, fg]) => {
    reqs.push(cfTextEq(IL, 5, 1005, 18, 19, text, bg, fg));
  });

  // ── Invoice Log: Follow-Up Status (col X = index 23) ─────────────────────
  const followRules = [
    ['Not Needed',  C.success,  '#1B3C1E'],
    ['Upcoming',    C.warning,  C.mainText],
    ['Due Today',   C.attention,C.white],
    ['Overdue',     C.attention,C.white],
    ['Completed',   C.success,  '#1B3C1E'],
  ];
  followRules.forEach(([text, bg, fg]) => {
    reqs.push(cfTextEq(IL, 5, 1005, 23, 24, text, bg, fg));
  });

  // ── Invoice Log: Overdue rows (entire row gray out when Cancelled/Refunded) ─
  reqs.push(cfFormula(IL, 5, 1005, 0, 28, `=$T6="Cancelled"`, C.neutral, C.white));
  reqs.push(cfFormula(IL, 5, 1005, 0, 28, `=$T6="Paid"`,      '#F0F5F1', '#1B3C1E'));
  reqs.push(cfFormula(IL, 5, 1005, 0, 28, `=$T6="Overdue"`,   '#FDF0EE', C.mainText));

  // ── Invoice Log: Balance Due negative guard (col R = index 17) ───────────
  reqs.push(cfFormula(IL, 5, 1005, 17, 18, `=$R6>0`, C.attention, C.white));
  reqs.push(cfFormula(IL, 5, 1005, 17, 18, `=$R6=0`, C.success, '#1B3C1E'));

  // ── Payment Tracker: Refund? flag (col L = index 11) ─────────────────────
  reqs.push(cfFormula(PT, 5, 1505, 0, 17, `=$L6=TRUE`, '#FDF0EE', C.mainText));

  // ── Payment Tracker: Net Payment negative (col N = index 13) ─────────────
  reqs.push(cfFormula(PT, 5, 1505, 13, 14, `=$N6<0`, C.attention, C.white));
  reqs.push(cfFormula(PT, 5, 1505, 13, 14, `=$N6>0`, C.success, '#1B3C1E'));

  // ── Invoice Generator: Ready-to-Log check (row 63, col N = index 13) ─────
  reqs.push(cfTextEq(IG, 62, 63, 13, 14, 'Ready',              C.success,   '#1B3C1E'));
  reqs.push(cfTextEq(IG, 62, 63, 13, 14, 'Missing Information', C.attention, C.white));
  reqs.push(cfTextEq(IG, 62, 63, 13, 14, 'Review Required',    C.warning,   C.mainText));

  // ── Dashboard: Overdue balance card (row 9, col H = index 7, 2 wide) ─────
  reqs.push(cfFormula(DB, 8, 10, 6, 8, `=$H9>0`, C.attention, C.white));

  await batchUpdate(id, reqs, 'conditional-formatting');
  console.log('✓ Conditional Formatting');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
