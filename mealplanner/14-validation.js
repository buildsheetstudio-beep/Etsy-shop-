'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const WMP = sheetMap['📅 Weekly Meal Plan'];
const RCP = sheetMap['📖 Recipe Book'];
const GRL = sheetMap['🛒 Grocery List'];
const FIN = sheetMap['🥗 Food Inventory'];
const MPR = sheetMap['🍳 Meal Prep Schedule'];
const GBD = sheetMap['💰 Grocery Budget'];
const DRT = sheetMap['🚫 Dietary Restrictions'];

// Reference column indices (0-based):
// A=0 MealTypes     rows 2-6  (5 items)
// B=1 Cuisines      rows 2-11 (10 items)
// C=2 Difficulty    rows 2-5  (4 items)
// D=3 AllergyTags   rows 2-9  (8 items)
// E=4 GroceryCats   rows 2-11 (10 items)
// F=5 Stores        rows 2-7  (6 items)
// G=6 StorageMethods rows 2-4 (3 items)
// H=7 Restrictions  rows 2-9  (8 items)

function oneOfRange(col, startRow, endRow) {
  const colLetter = String.fromCharCode(65 + col);
  return {
    condition: {
      type: 'ONE_OF_RANGE',
      values: [{ userEnteredValue: `='📋 Reference Data'!$${colLetter}$${startRow}:$${colLetter}$${endRow}` }],
    },
    strict: true, showCustomUi: true,
  };
}

function oneOfList(items) {
  return {
    condition: { type: 'ONE_OF_LIST', values: items.map(v => ({ userEnteredValue: v })) },
    strict: true, showCustomUi: true,
  };
}

function boolCheck() {
  return { condition: { type: 'BOOLEAN' }, strict: true };
}

function dv(rule, range) {
  return { setDataValidation: { range, rule } };
}

(async () => {
  const reqs = [];

  // ── Weekly Meal Plan ──────────────────────────────────────────────────────
  // Col B (1): Meal Type → Ref A col, rows 2-6
  reqs.push(dv(oneOfRange(0, 2, 6), gridRange(WMP, 2, 23, 1, 2)));
  // Col F (5): Prepped in Advance → checkbox
  reqs.push(dv(boolCheck(), gridRange(WMP, 2, 23, 5, 6)));

  // ── Recipe Book ───────────────────────────────────────────────────────────
  // Col C (2): Cuisine → Ref B col, rows 2-11
  reqs.push(dv(oneOfRange(1, 2, 11), gridRange(RCP, 2, 22, 2, 3)));
  // Col D (3): Meal Type → Ref A col, rows 2-6
  reqs.push(dv(oneOfRange(0, 2, 6), gridRange(RCP, 2, 22, 3, 4)));
  // Col E (4): Difficulty → Ref C col, rows 2-5
  reqs.push(dv(oneOfRange(2, 2, 5), gridRange(RCP, 2, 22, 4, 5)));

  // ── Grocery List ──────────────────────────────────────────────────────────
  // Col B (1): Category → Ref E col, rows 2-11
  reqs.push(dv(oneOfRange(4, 2, 11), gridRange(GRL, 2, 27, 1, 2)));
  // Col F (5): Store → Ref F col, rows 2-7
  reqs.push(dv(oneOfRange(5, 2, 7), gridRange(GRL, 2, 27, 5, 6)));
  // Col G (6): Purchased → checkbox
  reqs.push(dv(boolCheck(), gridRange(GRL, 2, 27, 6, 7)));

  // ── Food Inventory ────────────────────────────────────────────────────────
  // Col B (1): Category → Ref E col, rows 2-11
  reqs.push(dv(oneOfRange(4, 2, 11), gridRange(FIN, 2, 22, 1, 2)));

  // ── Meal Prep Schedule ────────────────────────────────────────────────────
  // Col D (3): Storage Method → Ref G col, rows 2-4
  reqs.push(dv(oneOfRange(6, 2, 4), gridRange(MPR, 2, 12, 3, 4)));

  // ── Dietary Restrictions ──────────────────────────────────────────────────
  // Col B (1): Restriction → Ref H col, rows 2-9
  reqs.push(dv(oneOfRange(7, 2, 9), gridRange(DRT, 2, 8, 1, 2)));

  await batchUpdate(id, reqs, 'validation');
  console.log('Data validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
