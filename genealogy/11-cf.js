'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SC  = sheetMap['📜 Source & Citation Tracker'];
const FM  = sheetMap['👨‍👩‍👧‍👦 Family Members'];
const TV  = sheetMap['🌳 Family Tree / Relationship Viewer'];
const LR  = sheetMap['📞 Living Relatives Directory'];

function cfRule(ranges, condition, format, index = 0) {
  return { addConditionalFormatRule: { rule: { ranges, booleanRule: { condition, format } }, index } };
}

(async () => {
  const reqs = [];

  // ── 1. Source & Citation Tracker: Confidence Level col E (idx 4) ──────
  // Confirmed → sage
  reqs.push(cfRule(
    [gridRange(SC, 1, 26, 0, 7)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E2="Confirmed"' }] },
    { backgroundColor: hex(C.sage), textFormat: { foregroundColor: hex('#1A3A1A') } },
  ));
  // Probable → amber
  reqs.push(cfRule(
    [gridRange(SC, 1, 26, 0, 7)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E2="Probable"' }] },
    { backgroundColor: hex(C.amber), textFormat: { foregroundColor: hex('#5A3800') } },
  ));
  // Uncertain → neutral gray
  reqs.push(cfRule(
    [gridRange(SC, 1, 26, 0, 7)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E2="Uncertain"' }] },
    { backgroundColor: hex(C.uncertain), textFormat: { foregroundColor: hex('#555555') } },
  ));

  // ── 2. Family Tree Viewer: highlight currently viewed person in Family Members ─
  // We apply a pale gold highlight on the TV sheet itself for the active record
  reqs.push(cfRule(
    [gridRange(TV, 3, 15, 1, 2)],
    { type: 'NOT_BLANK', values: [] },
    { backgroundColor: hex(C.lightGold) },
  ));

  // ── 3. Living Relatives: amber if Last Contact Date > 12 months ago ───
  // Col F (idx 5) = Last Contact Date; TODAY()-F2 > 365
  reqs.push(cfRule(
    [gridRange(LR, 1, 16, 0, 7)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($F2<>"",TODAY()-$F2>365)' }] },
    { backgroundColor: hex(C.amber), textFormat: { foregroundColor: hex('#5A3800') } },
  ));

  await batchUpdate(id, reqs, 'cf');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
