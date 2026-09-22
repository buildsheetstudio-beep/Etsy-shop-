'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const FS = sheetMap['Farm Setup'];
const IL = sheetMap['Income Log'];
const EL = sheetMap['Expense Log'];
const SIL = "'Income Log'";
const SEL = "'Expense Log'";
const SFS = "'Farm Setup'";

function clearValidation(sheetId, r1, r2, c1, c2) {
  return { setDataValidation: { range: gridRange(sheetId, r1, r2, c1, c2) } };
}

function numberValidation(sheetId, r1, r2, c1, c2) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r1, r2, c1, c2),
      rule: {
        condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0' }] },
        showCustomUi: true,
        strict: false,
      },
    },
  };
}

(async () => {
  const batchReqs = [];
  const valueData = [];

  // ── 1. FARM SETUP — Active? column (C16:C30) → checkbox validation ──────────
  // Add BOOLEAN data validation so cells render as checkboxes
  batchReqs.push({
    setDataValidation: {
      range: gridRange(FS, 15, 30, 2, 3), // rows 16-30, col C
      rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
    },
  });
  // Re-write `true` values — with checkbox validation applied they render as checked boxes
  for (let r = 16; r <= 30; r++) {
    valueData.push({ range: `${SFS}!C${r}`, values: [[true]] });
  }

  // ── 2. INCOME LOG — Fix misapplied validation ────────────────────────────────
  // Column J (index 9) = Quantity: had a "Payment Method" dropdown applied — WRONG
  // Column K (index 10) = Unit (text): had "number >= 0" validation — WRONG
  // Fix: clear both, add number >= 0 to J (Quantity) and L (Gross Income, index 11)
  batchReqs.push(clearValidation(IL, 5, 505, 9, 10));  // remove wrong dropdown from J
  batchReqs.push(clearValidation(IL, 5, 505, 10, 11)); // remove wrong number rule from K
  batchReqs.push(numberValidation(IL, 5, 505, 9, 10)); // J = Quantity, number >= 0
  batchReqs.push(numberValidation(IL, 5, 505, 11, 12)); // L = Gross Income, number >= 0

  // Re-push Year/Month/Quarter formulas for the 30 sample rows (6-35)
  // These were overwritten by null in the original sample data push (B:N range included C,D,E as null)
  for (let r = 6; r <= 35; r++) {
    valueData.push({ range: `${SIL}!C${r}`, values: [[`=IFERROR(YEAR(B${r}),"")`]] });
    valueData.push({ range: `${SIL}!D${r}`, values: [[`=IFERROR(TEXT(B${r},"mmmm"),"")`]] });
    valueData.push({ range: `${SIL}!E${r}`, values: [[`=IFERROR(IF(MONTH(B${r})<=3,"Q1",IF(MONTH(B${r})<=6,"Q2",IF(MONTH(B${r})<=9,"Q3","Q4"))),"")`]] });
  }

  // ── 3. EXPENSE LOG — Re-push all formula columns for the 40 sample rows ─────
  // The original B:Q sample data push included null for C, D, E, M, P positions,
  // which overwrote the formulas written in the loop, leaving those columns blank.
  for (let r = 6; r <= 45; r++) {
    valueData.push({ range: `${SEL}!C${r}`, values: [[`=IFERROR(YEAR(B${r}),"")`]] });
    valueData.push({ range: `${SEL}!D${r}`, values: [[`=IFERROR(TEXT(B${r},"mmmm"),"")`]] });
    valueData.push({ range: `${SEL}!E${r}`, values: [[`=IFERROR(IF(MONTH(B${r})<=3,"Q1",IF(MONTH(B${r})<=6,"Q2",IF(MONTH(B${r})<=9,"Q3","Q4"))),"")`]] });
    valueData.push({ range: `${SEL}!M${r}`, values: [[`=IFERROR(K${r}+L${r},"")`]] });
    valueData.push({ range: `${SEL}!P${r}`, values: [[`=IFERROR(VLOOKUP(G${r},'Reference Data'!$P$2:$Q$31,2,FALSE),"")`]] });
  }

  await batchUpdate(id, batchReqs, 'fix-validation');
  await valuesBatchUpdate(id, valueData, 'fix-values');
  console.log('All fixes applied successfully');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
