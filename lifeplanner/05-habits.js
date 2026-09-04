'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const HAB = sheetMap['🔁 Habit Tracker'];

const HEADERS = ['Date','Habit Name','Completed'];
const habits = ['Morning Meditation','Exercise / Move','Read 20 Minutes','Drink 8 Glasses of Water','Gratitude Journal'];
const dates = ['2025-01-06','2025-01-07','2025-01-08','2025-01-09','2025-01-10','2025-01-11','2025-01-12','2025-01-13'];
const completions = {
  'Morning Meditation':       [true, true, false,true, true, false,true, true ],
  'Exercise / Move':          [true, false,true, true, false,true, true, false],
  'Read 20 Minutes':          [true, true, true, false,true, true, true, true ],
  'Drink 8 Glasses of Water': [false,true, true, true, true, false,true, true ],
  'Gratitude Journal':        [true, true, false,false,true, true, false,true ],
};
const ROWS = [];
dates.forEach((date, di) => { habits.forEach(habit => { ROWS.push([date, habit, completions[habit][di]]); }); });

(async () => {
  if (ROWS.length !== 40) throw new Error(`Expected 40 rows, got ${ROWS.length}`);
  const data = [
    { range: "'🔁 Habit Tracker'!A1:C1", values: [HEADERS] },
    { range: "'🔁 Habit Tracker'!A2:C41", values: ROWS },
  ];
  await valuesBatchUpdate(id, data, 'hab-values');

  const reqs = [];
  reqs.push({ repeatCell: { range: gridRange(HAB,0,1,0,3), cell: { userEnteredFormat: { backgroundColor: hex(C.raspberry), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: HAB, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  for (let r = 1; r <= 40; r++) {
    reqs.push({ repeatCell: { range: gridRange(HAB,r,r+1,0,3), cell: { userEnteredFormat: { backgroundColor: hex(r%2===1?C.altRow:C.white), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  }
  reqs.push({ repeatCell: { range: gridRange(HAB,1,41,2,3), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });
  [[0,100],[1,180],[2,80]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: HAB, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  await batchUpdate(id, reqs, 'hab-format');
  console.log('Habit Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
