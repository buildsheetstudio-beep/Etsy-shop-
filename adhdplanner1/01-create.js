'use strict';
const { getSheets } = require('./lib');
const fs = require('fs');
const path = require('path');

const SHEETS = [
  { title: 'Reference Data',       sheetId: 0 },
  { title: 'Brain Dump',           sheetId: 1 },
  { title: 'Master Tasks',         sheetId: 2 },
  { title: 'Task Breakdown',       sheetId: 3 },
  { title: 'Daily Focus',          sheetId: 4 },
  { title: 'Weekly Planner',       sheetId: 5 },
  { title: 'Monthly Calendar',     sheetId: 6 },
  { title: 'Annual Calendar',      sheetId: 7 },
  { title: 'ADHD Focus Dashboard', sheetId: 8 },
];

(async () => {
  const sheets = await getSheets();

  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate ADHD Focus Flow Planner' },
      sheets: SHEETS.map(s => ({
        properties: {
          sheetId: s.sheetId,
          title: s.title,
          tabColor: { red: 0.208, green: 0.251, blue: 0.278 },
        },
      })),
    },
  });

  const id = res.data.spreadsheetId;
  const url = res.data.spreadsheetUrl;
  console.log(`✅  Created: ${url}`);

  // Hide Reference Data
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [
      { updateSheetProperties: { properties: { sheetId: 0, hidden: true }, fields: 'hidden' } },
      // Move ADHD Focus Dashboard to first visible position (after Reference Data)
      { updateSheetProperties: { properties: { sheetId: 8, index: 1 }, fields: 'index' } },
    ]},
  });

  const sheetMap = {};
  SHEETS.forEach(s => { sheetMap[s.title] = s.sheetId; });

  fs.writeFileSync(path.join(__dirname, 'spreadsheet.json'), JSON.stringify({ id, sheetMap }, null, 2));
  console.log('    spreadsheet.json written.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
