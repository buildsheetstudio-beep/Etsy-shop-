'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

// All tabs have full-width row-0 title merges → column freezes rejected by Sheets API.
// Rows-only freezes on all tabs.
const freezes = [
  { sheetId: sheetMap['Dashboard'],      rows: 3, cols: 0 }, // title + subtitle + KPI labels
  { sheetId: sheetMap['Assignments'],    rows: 4, cols: 0 }, // title + subtitle + stats + header
  { sheetId: sheetMap['Weekly Planner'], rows: 4, cols: 0 }, // title + week label + stats + day headers
  { sheetId: sheetMap['Grades & GPA'],   rows: 6, cols: 0 }, // title + subtitle + gpa rows + component header
  { sheetId: sheetMap['Student Budget'], rows: 4, cols: 0 }, // title + subtitle + kpi + (first section header)
  { sheetId: sheetMap['Exam Prep'],      rows: 4, cols: 0 }, // title + subtitle + stats + exam section header
  { sheetId: sheetMap['Reference'],      rows: 1, cols: 0 }, // just the header row
];

(async () => {
  const reqs = freezes.map(f => ({
    updateSheetProperties: {
      properties: {
        sheetId: f.sheetId,
        gridProperties: { frozenRowCount: f.rows, frozenColumnCount: f.cols },
      },
      fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
    },
  }));

  // Hide Reference Data tab
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: sheetMap['Reference'], hidden: true },
      fields: 'hidden',
    },
  });

  await batchUpdate(id, reqs, 'freezes-and-hide');
  console.log('Freezes applied, Reference hidden');
  freezes.forEach(f => {
    const name = Object.entries(sheetMap).find(([,v]) => v === f.sheetId)?.[0];
    console.log(` ${name}: frozen ${f.rows} rows`);
  });
})().catch(e => { console.error(e.errors || e); process.exit(1); });
