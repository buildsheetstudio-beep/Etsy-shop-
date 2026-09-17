'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID_MP  = sheetMap['Weekly Meal Planner'];
const SID_MCT = sheetMap['Meal Cost Tracker'];
const SID_DB  = sheetMap['Meal Planning Dashboard'];

const MP_S = "'Weekly Meal Planner'";
const DAYS = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
const WK_MP = "'Meal Planning Dashboard'!$A$4";  // week start for MP grid formulas
const WK_DB = '$A$4';                             // week start in Dashboard itself

(async () => {
  // ── Fix 1a: Insert row in Meal Planner to create Dinner row ─────────────────
  // Row 10 (0-indexed = 1-indexed row 11) was overwritten by column headers.
  // Inserting a blank row here pushes column headers to row 12 and data to row 13.
  // Google Sheets auto-updates all cross-sheet formula references past this point.
  await batchUpdate(id, [{
    insertDimension: {
      range: { sheetId: SID_MP, dimension: 'ROWS', startIndex: 10, endIndex: 11 },
      inheritFromBefore: false,
    },
  }], 'mp-insert-dinner-row');

  const fmtReqs = [];
  const valData = [];

  // ── Fix 1b: Format Dinner row at new 0-indexed row 10 ───────────────────────
  fmtReqs.push({ mergeCells: { range: gridRange(SID_MP, 10, 11, 0, 2), mergeType: 'MERGE_ALL' } });
  fmtReqs.push({
    repeatCell: {
      range: gridRange(SID_MP, 10, 11, 0, 2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.powder),
        textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  valData.push({ range: `${MP_S}!A11`, values: [['Dinner / Dessert']] });

  DAYS.forEach((_, di) => {
    const col = 2 + di * 2;
    fmtReqs.push({ mergeCells: { range: gridRange(SID_MP, 10, 11, col, col+2), mergeType: 'MERGE_ALL' } });
    fmtReqs.push({
      repeatCell: {
        range: gridRange(SID_MP, 10, 11, col, col+2),
        cell: { userEnteredFormat: {
          backgroundColor: hex(C.powder),
          textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
        } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
      },
    });
  });

  // ── Fix 1c: Replace broken INDEX/MATCH with FILTER for all 3 grid rows ──────
  // After insert, Meal Planner data is at rows 13-512.
  // WK_MP references Dashboard $A$4 (Monday Aug 18) as the week start.
  const gridRows = [
    { row: 8,  mealType: 'Breakfast' },   // 0-indexed row 8 = 1-indexed row 9
    { row: 9,  mealType: 'Lunch'     },   // 0-indexed row 9 = 1-indexed row 10
    { row: 10, mealType: 'Dinner'    },   // 0-indexed row 10 = 1-indexed row 11 (new)
  ];
  gridRows.forEach(({ row, mealType }) => {
    DAYS.forEach((_, di) => {
      const col = 2 + di * 2;
      const colLetter = String.fromCharCode(65 + col);
      const formula = `=IFERROR(INDEX(FILTER($G$13:$G$512,$C$13:$C$512=(${WK_MP}+${di}),$E$13:$E$512="${mealType}"),1),"—")`;
      valData.push({ range: `${MP_S}!${colLetter}${row + 1}`, values: [[formula]] });
    });
  });

  // ── Fix 2: Unique Recipes formula — prevent div/0 on empty cells ─────────────
  // After insert, Google Sheets auto-updated the range to F13:F512, but we still
  // need the inner IFERROR to guard against empty-cell division by zero.
  valData.push({
    range: `${MP_S}!F4`,
    values: [['=IFERROR(SUMPRODUCT(IFERROR(1/COUNTIFS(F13:F512,F13:F512,F13:F512,"REC-*"),0)),0)']],
  });

  // ── Fix 3: Budget Variance in Meal Cost Tracker (MCT!M10) ────────────────────
  // Was: =IFERROR(G8-B8,0) — Estimated Grocery Cost minus Planned Meal Cost (wrong)
  // Fix: Weekly Grocery Budget (I10) minus Weekly Planned Meal Cost (A8)
  valData.push({ range: `'Meal Cost Tracker'!M10`, values: [['=IFERROR(I10-A8,0)']] });

  // ── Fix 4: Budget Remaining in Dashboard (DB!M10) ────────────────────────────
  // Was: =IFERROR('Automated Grocery List'!G10,0) — referenced a pantry count cell
  // Fix: Household Setup budget minus Meal Cost Tracker weekly planned cost
  valData.push({
    range: `'Meal Planning Dashboard'!M10`,
    values: [["=IFERROR('Household Setup'!F8-'Meal Cost Tracker'!A8,0)"]],
  });

  // ── Fix 5: Dashboard Weekly Meal Snapshot — replace INDEX/MATCH with FILTER ──
  // SNAP_D0 = 19 (0-indexed). After MP row insert, data is at MP rows 13-512.
  // Also fix mealType for Snack row: data uses 'Snack', not 'Afternoon Snack'.
  const MP_REF = "'Weekly Meal Planner'";
  const snapshotMeals = [
    { ri: 0, mealType: 'Breakfast' },
    { ri: 1, mealType: 'Lunch'     },
    { ri: 2, mealType: 'Dinner'    },
    { ri: 3, mealType: 'Snack'     },
  ];
  snapshotMeals.forEach(({ ri, mealType }) => {
    const row = 19 + ri;  // 0-indexed
    DAYS.forEach((_, di) => {
      const col = 2 + di * 2;
      const colLetter = String.fromCharCode(65 + col);
      const formula = `=IFERROR(INDEX(FILTER(${MP_REF}!$G$13:$G$512,${MP_REF}!$C$13:$C$512=(${WK_DB}+${di}),${MP_REF}!$E$13:$E$512="${mealType}"),1),"—")`;
      valData.push({ range: `'Meal Planning Dashboard'!${colLetter}${row + 1}`, values: [[formula]] });
    });
  });

  await batchUpdate(id, fmtReqs, 'fixes-fmt');
  await valuesBatchUpdate(id, valData, 'fixes-vals');
  console.log('✓ All 5 fixes applied');
})().catch(e => { console.error(e.message || e); process.exit(1); });
