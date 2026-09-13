'use strict';
const { getSheets } = require('./lib');
const fs = require('fs');

const SHEETS = [
  'Reference Data',
  'Weight Progress',
  'Body Measurements',
  'Nutrition Log',
  'Macro Tracker',
  'Workout Tracker',
  'Weekly Meal Planner',
  'Grocery List',
  'Habit Tracker',
  'Goals & Milestones',
  'Weekly Check-In',
  'Non-Scale Victories',
  'Progress & Trends',
  'Weight Dashboard',
];

(async () => {
  const sheets = await getSheets();
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Weight Management & Lifestyle Tracker' },
      sheets: SHEETS.map((title, i) => ({
        properties: { sheetId: i, title, index: i },
      })),
    },
  });

  const id = res.data.spreadsheetId;

  // Hide Reference Data tab
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ updateSheetProperties: {
      properties: { sheetId: 0, hidden: true },
      fields: 'hidden',
    }}]},
  });

  const sheetMap = {};
  SHEETS.forEach((name, i) => { sheetMap[name] = i; });

  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({ id, sheetMap }, null, 2));
  console.log(`✅  Created: https://docs.google.com/spreadsheets/d/${id}/edit`);
  console.log('    spreadsheet.json written.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
