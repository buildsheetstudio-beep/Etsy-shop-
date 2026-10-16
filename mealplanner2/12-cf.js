'use strict';
const { sheets, hex, batchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID_RB   = sheetMap['Recipe Book'];
const SID_PI   = sheetMap['Pantry Inventory'];
const SID_MP   = sheetMap['Weekly Meal Planner'];
const SID_GL   = sheetMap['Automated Grocery List'];
const SID_RT   = sheetMap['Recipe Tags & Categories'];

// Helper: text-equality CF rule on a list of ranges
function cfText(ranges, value, bgHex, fgHex) {
  return {
    addConditionalFormatRule: {
      rule: {
        ranges,
        booleanRule: {
          condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: value }] },
          format: {
            backgroundColor: hex(bgHex),
            ...(fgHex ? { textFormat: { foregroundColor: hex(fgHex) } } : {}),
          },
        },
      },
      index: 0,
    },
  };
}

// Helper: custom-formula CF rule
function cfFormula(ranges, formula, bgHex, fgHex) {
  return {
    addConditionalFormatRule: {
      rule: {
        ranges,
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: formula }] },
          format: {
            backgroundColor: hex(bgHex),
            ...(fgHex ? { textFormat: { foregroundColor: hex(fgHex) } } : {}),
          },
        },
      },
      index: 0,
    },
  };
}

(async () => {
  const fmt = [];

  // Push order: lowest priority first, highest priority last.
  // Each addConditionalFormatRule with index:0 inserts at the top of the sheet's
  // CF rule list, pushing previous rules down. So the LAST rule we push ends up
  // at index 0 (highest priority) on the sheet.

  // ── MEAL TYPE COLORS ──────────────────────────────────────────────────────
  // Recipe Book col D (index 3, data rows 6-1005)
  // Meal Planner col E (index 4, data rows 12-101)
  // Recipe Tags browser col D (index 3, data rows 17-66)
  const mealTypeTargets = [
    [SID_RB, 3,  5, 1005, 1,  25],  // col D only
    [SID_MP, 4, 11,  101, 1,  20],  // col E only
    [SID_RT, 3, 16,   66, 1,  16],  // col D only
  ];
  const mealTypeColors = [
    ['Breakfast', C.peach],
    ['Lunch',     C.mint],
    ['Dinner',    C.lavender],
    ['Snack',     C.butter],
    ['Dessert',   C.blush],
  ];
  // Meal types are mutually exclusive — push in any order
  mealTypeTargets.forEach(([sid, col, r1, r2]) => {
    mealTypeColors.forEach(([label, color]) => {
      fmt.push(cfText([gridRange(sid, r1, r2, col, col + 1)], label, color));
    });
  });

  // ── RECIPE BOOK STATUS (col N = index 13) ────────────────────────────────
  const rbStatus = [gridRange(SID_RB, 5, 1005, 13, 14)];
  // Low-priority → high-priority (last pushed wins at index 0)
  fmt.push(cfText(rbStatus, 'Archived',  C.gray));
  fmt.push(cfText(rbStatus, 'Draft',     C.info));
  fmt.push(cfText(rbStatus, 'Active',    C.success));
  fmt.push(cfText(rbStatus, 'Inactive',  C.attention, C.white));
  fmt.push(cfText(rbStatus, 'Favorite',  C.peach));  // highest: peach wins

  // ── RECIPE BOOK: FAVORITE row tint (col M = index 12 = TRUE) ────────────
  // Tints the full row with butter when Favorite? is checked
  fmt.push(cfFormula(
    [gridRange(SID_RB, 5, 1005, 0, 25)],
    '=$M6=TRUE',
    C.butter
  ));

  // ── RECIPE TAGS STATUS (col P = index 15, browser rows 17-66) ────────────
  const rtStatus = [gridRange(SID_RT, 16, 66, 15, 16)];
  fmt.push(cfText(rtStatus, 'Archived',  C.gray));
  fmt.push(cfText(rtStatus, 'Draft',     C.info));
  fmt.push(cfText(rtStatus, 'Active',    C.success));
  fmt.push(cfText(rtStatus, 'Inactive',  C.attention, C.white));
  fmt.push(cfText(rtStatus, 'Favorite',  C.peach));

  // ── PANTRY STATUS (full row coloring by col G = index 6) ─────────────────
  // Only highlight problem states; leave "In Stock" as default alternating rows.
  const piRows = [gridRange(SID_PI, 5, 500, 0, 15)];
  fmt.push(cfFormula(piRows, '=$G6="Low Stock"',     C.butter));        // amber tint
  fmt.push(cfFormula(piRows, '=$G6="Expiring Soon"', C.peach));         // orange tint
  fmt.push(cfFormula(piRows, '=$G6="Out of Stock"',  C.attention, C.white)); // red + white text

  // ── GROCERY LIST STATUS (col L = index 11, data rows 12-70) ─────────────
  const glStatus = [gridRange(SID_GL, 11, 71, 11, 12)];
  fmt.push(cfText(glStatus, 'Already Have', C.success));
  fmt.push(cfText(glStatus, 'Optional',     C.lavender));
  fmt.push(cfText(glStatus, 'Review Units', C.info));
  fmt.push(cfText(glStatus, 'Needed',       C.attention, C.white)); // highest: red for "Needed"

  // ── GROCERY LIST PURCHASED? (col M = index 12 = TRUE → gray out full row) ─
  // This has the highest priority of all grocery rules — pushed last.
  fmt.push(cfFormula(
    [gridRange(SID_GL, 11, 71, 0, 17)],
    '=$M12=TRUE',
    C.altRow,
    C.secText
  ));

  await batchUpdate(id, fmt, 'cf');
  console.log('✓ Conditional formatting complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
