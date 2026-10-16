'use strict';
const { batchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const LOG_SID  = sheetMap['Daily Workout Log'];
const WEEK_SID = sheetMap['Weekly Planner'];

(async () => {
  const fmt = [];

  // Helper to build a CUSTOM_FORMULA boolean-rule CF request
  const cf = (sheetId, r1, r2, c1, c2, formula, bg, idx) => ({
    addConditionalFormatRule: {
      rule: {
        ranges: [{ sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 }],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: formula }] },
          format: { backgroundColor: hex(bg) },
        },
      },
      index: idx,
    },
  });

  // ── Daily Workout Log: row colors based on Status (col G = index 6) ─────────
  // Data rows ri=7..4999 (row 8..5000), all 30 cols
  // Status values from Reference Data col E:
  //   Planned | Completed | Partially Completed | Skipped | Rest Day | Rescheduled | Cancelled
  // Rules evaluated top-down; index 0 = highest priority
  fmt.push(cf(LOG_SID, 7, 5000, 0, 30, '=$G8="Completed"',           C.completed, 0));
  fmt.push(cf(LOG_SID, 7, 5000, 0, 30, '=$G8="Partially Completed"', C.partial,   1));
  fmt.push(cf(LOG_SID, 7, 5000, 0, 30, '=$G8="Planned"',             C.planned,   2));
  fmt.push(cf(LOG_SID, 7, 5000, 0, 30, '=$G8="Skipped"',             C.skipped,   3));
  fmt.push(cf(LOG_SID, 7, 5000, 0, 30, '=$G8="Rest Day"',            C.restDay,   4));
  fmt.push(cf(LOG_SID, 7, 5000, 0, 30, '=$G8="Rescheduled"',         C.planned,   5));

  // ── Weekly Planner: row colors based on Status (col E = index 4) ──────────
  // Data rows ri=3..9 (Mon-Sun = rows 4-10), cols 0-10
  fmt.push(cf(WEEK_SID, 3, 10, 0, 11, '=$E4="Completed"',           C.completed, 0));
  fmt.push(cf(WEEK_SID, 3, 10, 0, 11, '=$E4="Partially Completed"', C.partial,   1));
  fmt.push(cf(WEEK_SID, 3, 10, 0, 11, '=$E4="Planned"',             C.planned,   2));
  fmt.push(cf(WEEK_SID, 3, 10, 0, 11, '=$E4="Skipped"',             C.skipped,   3));
  fmt.push(cf(WEEK_SID, 3, 10, 0, 11, '=$E4="Rest Day"',            C.restDay,   4));

  await batchUpdate(id, fmt, '11-cf');
  console.log('✅  Conditional Formatting done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
