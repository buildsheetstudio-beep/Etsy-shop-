'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH   = sheetMap['Life Dashboard'];
const GOALS  = sheetMap['Annual Goals'];
const QPLAN  = sheetMap['Quarterly Planner'];
const MONTHLY = sheetMap['Monthly Milestones'];
const HABITS = sheetMap['Habit Tracker'];
const VISION = sheetMap['Vision & Affirmations'];
const REF    = sheetMap['Reference'];

(async () => {
  const reqs = [];

  // Freeze rows only (no col freezes to avoid merge conflicts), except Habit Tracker
  const freezeRowsOnly = [
    [DASH,   4],  // rows 1-4 frozen
    [GOALS,  4],  // rows 1-4 frozen
    [QPLAN,  4],  // rows 1-4 frozen
    [MONTHLY,3],  // rows 1-3 frozen
    [VISION, 2],  // rows 1-2 frozen
    [REF,    1],  // row 1 frozen
  ];

  freezeRowsOnly.forEach(([sid, rows]) => {
    reqs.push({
      updateSheetProperties: {
        properties: { sheetId: sid, gridProperties: { frozenRowCount: rows, frozenColumnCount: 0 } },
        fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
      },
    });
  });

  // Habit Tracker: 3 rows + 8 cols frozen (critical for 31-day grid scrolling)
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: HABITS, gridProperties: { frozenRowCount: 3, frozenColumnCount: 8 } },
      fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
    },
  });

  // Hide Reference tab
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: REF, hidden: true },
      fields: 'hidden',
    },
  });

  await batchUpdate(id, reqs, 'freezes');

  console.log('Freezes + Reference hidden complete');
  console.log('\n✅ Ultimate Goal Planner COMPLETE!');
  console.log('URL:', JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json')).url);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
