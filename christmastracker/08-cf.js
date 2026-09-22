'use strict';
const { batchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const GRT = sheetMap['🎁 Gift Recipient Tracker'];
const WSH = sheetMap['💡 Gift Idea Wishlist'];
const CDL = sheetMap['💌 Holiday Card & Mailing List'];
const RET = sheetMap['🔄 Return & Exchange Tracker'];

function cfRule(ranges, condition, format, index = 0) {
  return {
    addConditionalFormatRule: {
      rule: { ranges, booleanRule: { condition, format } },
      index,
    },
  };
}

(async () => {
  const reqs = [];

  // ── 1. Gift Recipient Tracker: Diff column (G, idx 6) ────────────────────
  // Red when negative (over budget: actual > budget → E-F < 0)
  reqs.push(cfRule(
    [gridRange(GRT, 1, 26, 6, 7)],
    { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
    { backgroundColor: hex(C.lightRed), textFormat: { foregroundColor: hex(C.trueRed), bold: true } },
  ));
  // Sage when zero or positive (under/on budget)
  reqs.push(cfRule(
    [gridRange(GRT, 1, 26, 6, 7)],
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0' }] },
    { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.deepGreen), bold: true } },
  ));

  // ── 2. Gift Recipient Tracker: Bought/Wrapped/Sent checkboxes ────────────
  // Single CF rule on H2:J26 with relative formula — shifts column per cell
  reqs.push(cfRule(
    [gridRange(GRT, 1, 26, 7, 10)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=H2=TRUE' }] },
    { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.deepGreen) } },
  ));

  // ── 3. Holiday Card & Mailing List: Card Status (col C, idx 2) ───────────
  // Sage = Received
  reqs.push(cfRule(
    [gridRange(CDL, 1, 21, 0, 5)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$C2="Received"' }] },
    { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.deepGreen) } },
  ));
  // Amber = Sent
  reqs.push(cfRule(
    [gridRange(CDL, 1, 21, 0, 5)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$C2="Sent"' }] },
    { backgroundColor: hex(C.lightAmber), textFormat: { foregroundColor: hex(C.darkText) } },
  ));
  // Light gray = Not Sent
  reqs.push(cfRule(
    [gridRange(CDL, 1, 21, 0, 5)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$C2="Not Sent"' }] },
    { backgroundColor: hex(C.lightGray), textFormat: { foregroundColor: hex(C.mediumGray) } },
  ));

  // ── 4. Return & Exchange Tracker ─────────────────────────────────────────
  // Red: deadline within 3 days (inclusive today) AND still "Pending Return"
  reqs.push(cfRule(
    [gridRange(RET, 1, 11, 0, 10)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($F2<>"",TODAY()<=$F2,$F2-TODAY()<=3,$H2="Pending Return")' }] },
    { backgroundColor: hex(C.lightRed), textFormat: { foregroundColor: hex(C.trueRed), bold: true } },
  ));
  // Sage: Returned or Exchanged
  reqs.push(cfRule(
    [gridRange(RET, 1, 11, 0, 10)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=OR($H2="Returned",$H2="Exchanged")' }] },
    { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.deepGreen) } },
  ));

  // ── 5. Gift Idea Wishlist: Priority=High AND not yet Selected ─────────────
  // Gold highlight — surfaces top-priority items still needing a decision
  reqs.push(cfRule(
    [gridRange(WSH, 1, 21, 0, 7)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($D2="High",$G2=FALSE)' }] },
    { backgroundColor: hex(C.lightGold), textFormat: { foregroundColor: hex(C.darkText), bold: true } },
  ));

  await batchUpdate(id, reqs, 'cf');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
