'use strict';
const { getAuth } = require('./lib');
const { google } = require('googleapis');
const fs = require('fs');

function rgb(h) {
  return {
    red:   parseInt(h.slice(1, 3), 16) / 255,
    green: parseInt(h.slice(3, 5), 16) / 255,
    blue:  parseInt(h.slice(5, 7), 16) / 255,
  };
}

const TABS = [
  { title: '📊 Dashboard',         color: '#2454A6' },
  { title: '📝 Workout Log',        color: '#7CB342' },
  { title: '🏆 PR Tracker',         color: '#FFC107' },
  { title: '📏 Body Measurements',  color: '#E8534A' },
  { title: '📚 Program Templates',  color: '#607D8B' },
  { title: '🏃 Cardio Tracker',     color: '#00897B' },
  { title: '😴 Rest & Recovery',    color: '#7B1FA2' },
  { title: '📅 Monthly Calendar',   color: '#1565C0' },
  { title: '📋 Reference Data',     color: '#90A4AE' },
];

(async () => {
  const sheets = google.sheets({ version: 'v4', auth: getAuth() });
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: {
        title: '🏋️ Ultimate Workout Planner',
        defaultFormat: { backgroundColor: { red: 0.98, green: 0.98, blue: 0.98 } },
      },
      sheets: TABS.map((t, i) => ({
        properties: {
          sheetId: i,
          title: t.title,
          index: i,
          tabColorStyle: { rgbColor: rgb(t.color) },
        },
      })),
    },
  });

  const id = res.data.spreadsheetId;
  const sheetMap = {};
  for (const s of res.data.sheets) {
    sheetMap[s.properties.title] = s.properties.sheetId;
  }
  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({ id, sheetMap }, null, 2));
  console.log('Created spreadsheet:', id);
  console.log('Sheet map:', JSON.stringify(sheetMap, null, 2));
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
