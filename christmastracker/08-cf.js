'use strict';
const { batchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const GRT = sheetMap['🎁 Gift Recipient Tracker'];
const WSH = sheetMap['💡 Gift Idea Wishlist'];
const CDL = sheetMap['💌 Holiday Card & Mailing List'];
const RET = sheetMap['🔄 Return & Exchange Tracker'];

function cfRule(ranges, condition, format) {
  return {
    addConditionalFormatRule: {
      rule: {
        ranges,
        booleanRule: { condition, format },
      },
      index: 0,
    },
  };
}

(async () => {
  const reqs = [];

  // ── Gift Recipient Tracker ────────────────────────────────────────────────

  // 1. Row is fully bought + wrapped + delivered → light sage green (celebration!)
  reqs.push(cfRule(
    [gridRange(GRT, 1, 26, 0, 13)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($H2=TRUE,$I2=TRUE,$J2=TRUE)' }] },
    { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.deepGreen) } },
  ));

  // 2. Bought but not yet wrapped → light amber
  reqs.push(cfRule(
    [gridRange(GRT, 1, 26, 0, 13)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($H2=TRUE,$I2=FALSE)' }] },
    { backgroundColor: hex(C.lightAmber), textFormat: { foregroundColor: hex(C.darkText) } },
  ));

  // 3. Not yet bought → light red (action needed)
  reqs.push(cfRule(
    [gridRange(GRT, 1, 26, 0, 13)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$H2=FALSE' }] },
    { backgroundColor: hex(C.lightRed), textFormat: { foregroundColor: hex(C.trueRed) } },
  ));

  // 4. Over budget (Diff col G > 0) → red bold on that cell
  reqs.push(cfRule(
    [gridRange(GRT, 1, 26, 6, 7)],
    { type: 'NUMBER_GREATER', values: [{ userEnteredValue: '0' }] },
    { backgroundColor: hex(C.lightRed), textFormat: { foregroundColor: hex(C.trueRed), bold: true } },
  ));

  // 5. Under budget (Diff col G < 0) → green
  reqs.push(cfRule(
    [gridRange(GRT, 1, 26, 6, 7)],
    { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
    { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.deepGreen), bold: true } },
  ));

  // ── Gift Idea Wishlist ────────────────────────────────────────────────────

  // 6. Selected = TRUE → warm gold highlight
  reqs.push(cfRule(
    [gridRange(WSH, 1, 21, 0, 7)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G2=TRUE' }] },
    { backgroundColor: hex(C.lightGold), textFormat: { foregroundColor: hex(C.darkText), bold: true } },
  ));

  // ── Holiday Card & Mailing List ───────────────────────────────────────────

  // 7. Card Received = TRUE → sage green
  reqs.push(cfRule(
    [gridRange(CDL, 1, 21, 0, 5)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$D2=TRUE' }] },
    { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.deepGreen) } },
  ));

  // 8. Card Status = "Sent" (but not received) → light amber
  reqs.push(cfRule(
    [gridRange(CDL, 1, 21, 0, 5)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($C2="Sent",$D2=FALSE)' }] },
    { backgroundColor: hex(C.lightAmber), textFormat: { foregroundColor: hex(C.darkText) } },
  ));

  // ── Return & Exchange Tracker ─────────────────────────────────────────────

  // 9. Return status = "Completed" → light sage
  reqs.push(cfRule(
    [gridRange(RET, 1, 11, 0, 10)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$H2="Completed"' }] },
    { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.deepGreen) } },
  ));

  // 10. Return status = "Pending" → light red
  reqs.push(cfRule(
    [gridRange(RET, 1, 11, 0, 10)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$H2="Pending"' }] },
    { backgroundColor: hex(C.lightRed), textFormat: { foregroundColor: hex(C.trueRed) } },
  ));

  await batchUpdate(id, reqs, 'cf');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
