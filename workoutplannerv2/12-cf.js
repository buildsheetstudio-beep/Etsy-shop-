'use strict';
const { batchUpdate, gridRange, hex, C, MG_COLORS, ZONE_COLORS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const WL  = sheetMap['💪 Workout Log'];
const CAL = sheetMap['📅 Monthly Calendar View'];
const CRD = sheetMap['❤️ Cardio & Heart Rate Zone Tracker'];
const PR  = sheetMap['🏆 Personal Records (PR) Tracker'];
const STK = sheetMap['🔥 Streak & Consistency Tracker'];

function cfRule(ranges, condition, format, index = 0) {
  return { addConditionalFormatRule: { rule: { ranges, booleanRule: { condition, format } }, index } };
}

(async () => {
  const reqs = [];

  // ── 1. Workout Log: Completed col F (idx 5) ───────────────────────────────
  // Sage fill on completed rows
  reqs.push(cfRule(
    [gridRange(WL, 1, 31, 0, 7)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F2=TRUE' }] },
    { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.sage) } },
  ));
  // Red text when NOT completed AND date has passed (A2 < TODAY())
  reqs.push(cfRule(
    [gridRange(WL, 1, 31, 0, 7)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($F2=FALSE,$A2<TODAY(),$A2<>"")' }] },
    { textFormat: { foregroundColor: hex(C.brick), bold: true } },
  ));
  // Amber/gray when NOT completed AND date in future
  reqs.push(cfRule(
    [gridRange(WL, 1, 31, 0, 7)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($F2=FALSE,$A2>=TODAY(),$A2<>"")' }] },
    { backgroundColor: hex(C.lightAmber) },
  ));

  // ── 2. Monthly Calendar: Rest day formula rows ─────────────────────────────
  // Rows idx 4,6,8,10,12 (formula rows) — light gray when "Rest" is shown
  [4, 6, 8, 10, 12].forEach(ri => {
    reqs.push(cfRule(
      [gridRange(CAL, ri, ri+1, 0, 7)],
      { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=ISNUMBER(SEARCH("Rest",A5))' }] },
      { backgroundColor: hex(MG_COLORS['Rest']) },
    ));
    // Non-empty and NOT rest → light lime
    reqs.push(cfRule(
      [gridRange(CAL, ri, ri+1, 0, 7)],
      { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND(A5<>"",NOT(ISNUMBER(SEARCH("Rest",A5))))' }] },
      { backgroundColor: hex('#E8F5E9') },
    ));
  });

  // ── 3. Cardio Tracker: Zone col F (idx 5) — cool-to-warm gradient ─────────
  // Zone 1 (blue)
  reqs.push(cfRule(
    [gridRange(CRD, 1, 13, 5, 6)],
    { type: 'TEXT_CONTAINS', values: [{ userEnteredValue: 'Zone 1' }] },
    { backgroundColor: hex(ZONE_COLORS['Zone 1']), textFormat: { foregroundColor: hex('#1A5276') } },
  ));
  // Zone 2 (teal)
  reqs.push(cfRule(
    [gridRange(CRD, 1, 13, 5, 6)],
    { type: 'TEXT_CONTAINS', values: [{ userEnteredValue: 'Zone 2' }] },
    { backgroundColor: hex(ZONE_COLORS['Zone 2']), textFormat: { foregroundColor: hex('#1E8449') } },
  ));
  // Zone 3 (yellow)
  reqs.push(cfRule(
    [gridRange(CRD, 1, 13, 5, 6)],
    { type: 'TEXT_CONTAINS', values: [{ userEnteredValue: 'Zone 3' }] },
    { backgroundColor: hex(ZONE_COLORS['Zone 3']), textFormat: { foregroundColor: hex('#9A7D0A') } },
  ));
  // Zone 4 (orange)
  reqs.push(cfRule(
    [gridRange(CRD, 1, 13, 5, 6)],
    { type: 'TEXT_CONTAINS', values: [{ userEnteredValue: 'Zone 4' }] },
    { backgroundColor: hex(ZONE_COLORS['Zone 4']), textFormat: { foregroundColor: hex('#784212') } },
  ));
  // Zone 5 (red)
  reqs.push(cfRule(
    [gridRange(CRD, 1, 13, 5, 6)],
    { type: 'TEXT_CONTAINS', values: [{ userEnteredValue: 'Zone 5' }] },
    { backgroundColor: hex(ZONE_COLORS['Zone 5']), textFormat: { foregroundColor: hex('#922B21') } },
  ));

  // ── 4. PR Tracker: highlight row when Current Best = Value (col G = col D) ─
  // Gold highlight on current best rows
  reqs.push(cfRule(
    [gridRange(PR, 1, 14, 0, 8)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($G2<>"",$D2=$G2)' }] },
    { backgroundColor: hex(C.lightGold), textFormat: { bold: true, foregroundColor: hex('#7D6608') } },
  ));

  // ── 5. Streak & Consistency: Running Streak col C ─────────────────────────
  // Sage on streak ≥ 7
  reqs.push(cfRule(
    [gridRange(STK, 1, 31, 2, 3)],
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '7' }] },
    { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.sage), bold: true } },
  ));
  // Amber on reset (= 0) — broken streak marker (gentle, not shaming)
  reqs.push(cfRule(
    [gridRange(STK, 1, 31, 2, 3)],
    { type: 'NUMBER_EQ', values: [{ userEnteredValue: '0' }] },
    { backgroundColor: hex(C.lightAmber), textFormat: { foregroundColor: hex(C.amber) } },
  ));
  // Valid Day col B — sage when TRUE
  reqs.push(cfRule(
    [gridRange(STK, 1, 31, 1, 2)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$B2=TRUE' }] },
    { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.sage) } },
  ));

  await batchUpdate(id, reqs, 'cf');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
