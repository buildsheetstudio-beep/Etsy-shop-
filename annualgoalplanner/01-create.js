'use strict';
const { getAuth, google } = require('./lib');
const fs = require('fs');
const path = require('path');

function rgb(h) {
  return { red: parseInt(h.slice(1,3),16)/255, green: parseInt(h.slice(3,5),16)/255, blue: parseInt(h.slice(5,7),16)/255 };
}

const TABS = [
  { title: '🏆 Goal Dashboard',               color: '#0A0F3D' },
  { title: '🎯 Goals & Milestones',           color: '#1A2060' },
  { title: '📋 Action Steps',                 color: '#2D3A8C' },
  { title: '📅 Weekly Check-In & Reflection', color: '#3D4FA6' },
  { title: '📋 Reference Data',               color: '#888888' },
];

(async () => {
  const auth   = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Annual Goal Planner' },
      sheets: TABS.map((t, i) => ({
        properties: {
          sheetId: i,
          index:   i,
          title:   t.title,
          tabColorStyle: { rgbColor: rgb(t.color) },
          gridProperties: { rowCount: 200, columnCount: 26 },
        },
      })),
    },
  });

  const id  = res.data.spreadsheetId;
  const url = res.data.spreadsheetUrl;
  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });

  fs.writeFileSync(path.join(__dirname,'spreadsheet.json'), JSON.stringify({ id, url, sheetMap }, null, 2));
  console.log('Created:', url);
  console.log('ID:', id);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
