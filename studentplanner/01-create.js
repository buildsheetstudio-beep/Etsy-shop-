'use strict';
const { sheets } = require('./lib');
const fs = require('fs');

const TABS = [
  { title: 'Dashboard',        color: { red: 0.231, green: 0.510, blue: 0.965 } }, // #3B82F6
  { title: 'Assignments',      color: { red: 0.376, green: 0.647, blue: 0.980 } }, // #60A5FA
  { title: 'Weekly Planner',   color: { red: 0.204, green: 0.827, blue: 0.600 } }, // #34D399
  { title: 'Grades & GPA',     color: { red: 0.655, green: 0.545, blue: 0.980 } }, // #A78BFA
  { title: 'Student Budget',   color: { red: 0.984, green: 0.749, blue: 0.141 } }, // #FBBF24
  { title: 'Exam Prep',        color: { red: 0.984, green: 0.573, blue: 0.235 } }, // #FB923C
  { title: 'Reference',        color: { red: 0.580, green: 0.639, blue: 0.722 } }, // #94A3B8
];

(async () => {
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Student Planner' },
      sheets: TABS.map((t, i) => ({
        properties: {
          sheetId: i,
          title: t.title,
          index: i,
          tabColor: t.color,
          gridProperties: { rowCount: i === 2 ? 50 : (i === 0 ? 60 : 220), columnCount: 26 },
        },
      })),
    },
  });

  const spreadsheetId = res.data.spreadsheetId;
  const url = res.data.spreadsheetUrl;

  // Map tab titles to sheetIds (Google assigns them, but we pass sheetId=0,1,2... in create request)
  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });

  console.log('Spreadsheet created:', spreadsheetId);
  console.log('URL:', url);
  console.log('Sheet IDs:');
  Object.entries(sheetMap).forEach(([t, id]) => console.log(' ', id, t));

  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({ id: spreadsheetId, url, sheetMap }, null, 2));
  console.log('spreadsheet.json written');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
