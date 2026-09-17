'use strict';
const { getAuth, google } = require('./lib');
const fs = require('fs');

const TABS = [
  { title: '🏡 Dashboard',                      color: '#3A5A2A' },
  { title: '✅ Planning Checklist',              color: '#4A7A3A' },
  { title: '👨‍👩‍👧‍👦 Guest List & RSVPs',              color: '#5A8A4A' },
  { title: '💰 Budget & Expenses',              color: '#7A6A3A' },
  { title: '🗓️ Event Itinerary & Activities',   color: '#3A5A4A' },
  { title: '🍖 Food, Drinks & Potluck',         color: '#8B5A2A' },
  { title: '🏠 Accommodation & Travel',         color: '#5A7A8B' },
  { title: '📋 Reference Data',                 color: '#888888' },
];

(async () => {
  const auth   = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  const drive  = google.drive({ version: 'v3', auth });

  const res = await sheets.spreadsheets.create({ requestBody: {
    properties: { title: 'Ultimate Family Reunion Planner — BuildSheetStudio' },
    sheets: TABS.map((t, i) => ({
      properties: {
        sheetId: i, title: t.title, index: i,
        tabColorStyle: { rgbColor: { red: parseInt(t.color.slice(1,3),16)/255, green: parseInt(t.color.slice(3,5),16)/255, blue: parseInt(t.color.slice(5,7),16)/255 } },
        gridProperties: { rowCount: 300, columnCount: 26 },
      },
    })),
  }});

  const id = res.data.spreadsheetId;
  await drive.permissions.create({ fileId: id, requestBody: { role: 'reader', type: 'anyone' } });

  const sheetMap = {};
  TABS.forEach((t, i) => { sheetMap[t.title] = i; });
  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({ id, sheetMap }, null, 2));
  console.log('Created:', `https://docs.google.com/spreadsheets/d/${id}/edit`);
})().catch(e => { console.error(e.message); process.exit(1); });
