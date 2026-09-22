'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const BGT = sheetMap['💰 Budget & Expense Tracker'];
const VLT = sheetMap['🔐 Travel Document & Confirmation Vault'];
const WIS = sheetMap['✨ Activity & Excursion Wishlist'];
const CHK = sheetMap['✅ Research & Reservation Checklist'];
const PKG = sheetMap['🧳 Packing List'];
const ITN = sheetMap['🗓️ Multi-City Day-by-Day Itinerary'];

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

  // ── 1. Budget: Diff col E — under budget (positive) → sage, over budget (negative) → rust ──
  reqs.push(cfRule(
    [gridRange(BGT, 1, 26, 4, 5)],
    { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
    { backgroundColor: hex(C.lightRust), textFormat: { foregroundColor: hex(C.rustRed), bold: true } },
  ));
  reqs.push(cfRule(
    [gridRange(BGT, 1, 26, 4, 5)],
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0' }] },
    { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.mutedSage), bold: true } },
  ));

  // ── 2. Document Vault: Status col E — Confirmed → sage, Researching → amber, Cancelled → rust ──
  reqs.push(cfRule(
    [gridRange(VLT, 1, 16, 0, 9)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E2="Confirmed"' }] },
    { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.mutedSage) } },
  ));
  reqs.push(cfRule(
    [gridRange(VLT, 1, 16, 0, 9)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E2="Researching"' }] },
    { backgroundColor: hex(C.lightAmber), textFormat: { foregroundColor: hex(C.amber) } },
  ));
  reqs.push(cfRule(
    [gridRange(VLT, 1, 16, 0, 9)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E2="Cancelled"' }] },
    { backgroundColor: hex(C.lightRust), textFormat: { foregroundColor: hex(C.rustRed) } },
  ));
  reqs.push(cfRule(
    [gridRange(VLT, 1, 16, 0, 9)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E2="Pending"' }] },
    { backgroundColor: hex(C.lightAmber), textFormat: { foregroundColor: hex(C.warmCharcoal) } },
  ));

  // ── 3. Activity Wishlist: Booking Status col F — Confirmed → sage, Researching/Pending → amber, Cancelled → rust ──
  reqs.push(cfRule(
    [gridRange(WIS, 1, 21, 0, 8)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F2="Confirmed"' }] },
    { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.mutedSage) } },
  ));
  reqs.push(cfRule(
    [gridRange(WIS, 1, 21, 0, 8)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=OR($F2="Researching",$F2="Pending")' }] },
    { backgroundColor: hex(C.lightAmber), textFormat: { foregroundColor: hex(C.amber) } },
  ));
  reqs.push(cfRule(
    [gridRange(WIS, 1, 21, 0, 8)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F2="Cancelled"' }] },
    { backgroundColor: hex(C.lightRust), textFormat: { foregroundColor: hex(C.rustRed) } },
  ));
  // Priority = High + not Confirmed → highlight
  reqs.push(cfRule(
    [gridRange(WIS, 1, 21, 0, 8)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($E2="High",$F2<>"Confirmed")' }] },
    { textFormat: { bold: true } },
  ));

  // ── 4. Checklist: Done checkbox col E → strike-through sage ──
  reqs.push(cfRule(
    [gridRange(CHK, 1, 26, 0, 6)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E2=TRUE' }] },
    { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.mutedSage), strikethrough: true } },
  ));

  // ── 5. Packing: Packed checkbox col D → sage ──
  reqs.push(cfRule(
    [gridRange(PKG, 1, 34, 0, 5)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$D2=TRUE' }] },
    { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.mutedSage) } },
  ));

  // ── 6. Itinerary: Booked checkbox col G → teal highlight ──
  reqs.push(cfRule(
    [gridRange(ITN, 1, 31, 0, 8)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G2=TRUE' }] },
    { backgroundColor: hex(C.lightTeal), textFormat: { foregroundColor: hex(C.deepTeal) } },
  ));

  await batchUpdate(id, reqs, 'cf');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
