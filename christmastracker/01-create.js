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
  { title: '📋 Reference Data',              color: '#9E9E9E' },
  { title: '📊 Dashboard',                   color: '#8C1F2B' },
  { title: '🎁 Gift Recipient Tracker',      color: '#C9A227' },
  { title: '💡 Gift Idea Wishlist',          color: '#2D5016' },
  { title: '💌 Holiday Card & Mailing List', color: '#B23A3A' },
  { title: '🔄 Return & Exchange Tracker',   color: '#7A9B7A' },
  { title: '🌙 Dashboard — Dark Theme',      color: '#1F2B22' },
];

(async () => {
  const sheets = google.sheets({ version: 'v4', auth: getAuth() });
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: {
        title: '🎄 Ultimate Christmas Gift Tracker',
        defaultFormat: { backgroundColor: { red: 0.969, green: 0.973, blue: 0.969 } },
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
