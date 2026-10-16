'use strict';
const { batchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH = sheetMap['Life Dashboard'];
const GOALS = sheetMap['Annual Goals'];
const MONTHLY = sheetMap['Monthly Milestones'];
const HABITS = sheetMap['Habit Tracker'];

function cfRule(ranges, condition, format) {
  return { addConditionalFormatRule: { rule: { ranges, booleanRule: { condition, format } }, index: 0 } };
}

(async () => {
  const reqs = [];

  // ── Annual Goals: Status column (E, col 4) — rows 5-37 ───────────────────
  const goalStatusRows = [
    [4,7],[9,12],[15,18],[21,24],[27,30],[33,36]
  ].flatMap(([s,e]) => {
    const rows = [];
    for (let r = s; r < e; r++) rows.push(r);
    return rows;
  });

  const goalRanges = goalStatusRows.map(r => gridRange(GOALS, r, r + 1, 4, 5));

  // Complete (green)
  reqs.push(cfRule(goalRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Complete' }] },
    { backgroundColor: hex(C.completeBg), textFormat: { foregroundColor: hex(C.complete), bold: true } }
  ));
  // In Progress (blue)
  reqs.push(cfRule(goalRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'In Progress' }] },
    { backgroundColor: hex(C.inProgressBg), textFormat: { foregroundColor: hex(C.inProgress) } }
  ));
  // Not Started (gray)
  reqs.push(cfRule(goalRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Not Started' }] },
    { backgroundColor: hex(C.notStartedBg), textFormat: { foregroundColor: hex(C.notStarted) } }
  ));
  // On Hold (gold)
  reqs.push(cfRule(goalRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'On Hold' }] },
    { backgroundColor: hex(C.onHoldBg), textFormat: { foregroundColor: hex(C.onHold) } }
  ));

  // ── Annual Goals: % Done column (F, col 5) ────────────────────────────────
  const pctRanges = goalStatusRows.map(r => gridRange(GOALS, r, r + 1, 5, 6));
  // >= 100%
  reqs.push(cfRule(pctRanges,
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '1' }] },
    { backgroundColor: hex(C.completeBg), textFormat: { foregroundColor: hex(C.complete), bold: true } }
  ));
  // >= 50%
  reqs.push(cfRule(pctRanges,
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0.5' }] },
    { backgroundColor: hex(C.inProgressBg), textFormat: { foregroundColor: hex(C.inProgress) } }
  ));
  // < 25%
  reqs.push(cfRule(pctRanges,
    { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0.25' }] },
    { backgroundColor: hex(C.notStartedBg), textFormat: { foregroundColor: hex(C.notStarted) } }
  ));

  // ── Monthly Milestones: Status col (D, col 3) rows 5-10 ───────────────────
  const monthlyStatusRanges = [gridRange(MONTHLY, 4, 10, 3, 4)];
  reqs.push(cfRule(monthlyStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Complete' }] },
    { backgroundColor: hex(C.completeBg), textFormat: { foregroundColor: hex(C.complete), bold: true } }
  ));
  reqs.push(cfRule(monthlyStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'In Progress' }] },
    { backgroundColor: hex(C.inProgressBg), textFormat: { foregroundColor: hex(C.inProgress) } }
  ));
  reqs.push(cfRule(monthlyStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Not Started' }] },
    { backgroundColor: hex(C.notStartedBg), textFormat: { foregroundColor: hex(C.notStarted) } }
  ));

  // ── Monthly Milestones: % col (E, col 4) rows 5-10 ────────────────────────
  const monthlyPctRanges = [gridRange(MONTHLY, 4, 10, 4, 5)];
  reqs.push(cfRule(monthlyPctRanges,
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '1' }] },
    { backgroundColor: hex(C.completeBg), textFormat: { foregroundColor: hex(C.complete), bold: true } }
  ));
  reqs.push(cfRule(monthlyPctRanges,
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0.5' }] },
    { backgroundColor: hex(C.inProgressBg), textFormat: { foregroundColor: hex(C.inProgress) } }
  ));

  // ── Habit Tracker: Checkbox TRUE = green, FALSE = subtle ─────────────────
  // Days 1-31 (cols I-AM = indices 8-38) for habits rows 4-15 (indices 3-14)
  const habitDayRanges = [gridRange(HABITS, 3, 15, 8, 39)];
  reqs.push(cfRule(habitDayRanges,
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$I4=TRUE' }] },
    { backgroundColor: hex(C.complete) }
  ));
  // FALSE explicitly (unchecked but not empty)
  reqs.push(cfRule(habitDayRanges,
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$I4=FALSE' }] },
    { backgroundColor: hex(C.notStartedBg) }
  ));

  // ── Habit % col (G, col 6) rows 4-15 ─────────────────────────────────────
  const habitPctRanges = [gridRange(HABITS, 3, 15, 6, 7)];
  reqs.push(cfRule(habitPctRanges,
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0.9' }] },
    { backgroundColor: hex(C.completeBg), textFormat: { foregroundColor: hex(C.complete), bold: true } }
  ));
  reqs.push(cfRule(habitPctRanges,
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0.6' }] },
    { backgroundColor: hex(C.inProgressBg), textFormat: { foregroundColor: hex(C.inProgress) } }
  ));
  reqs.push(cfRule(habitPctRanges,
    { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0.4' }] },
    { backgroundColor: hex(C.paleBlush), textFormat: { foregroundColor: hex(C.blush) } }
  ));

  // ── Life Dashboard: Wheel of Life score col (I, col 8) rows 9-14 ──────────
  const dashScoreRanges = [gridRange(DASH, 8, 14, 8, 9)];
  reqs.push(cfRule(dashScoreRanges,
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '8' }] },
    { backgroundColor: hex(C.completeBg), textFormat: { foregroundColor: hex(C.complete), bold: true } }
  ));
  reqs.push(cfRule(dashScoreRanges,
    { type: 'NUMBER_LESS', values: [{ userEnteredValue: '5' }] },
    { backgroundColor: hex(C.paleBlush), textFormat: { foregroundColor: hex(C.blush) } }
  ));

  await batchUpdate(id, reqs, 'cf-rules');

  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
