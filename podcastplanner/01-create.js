'use strict';
const { getAuth, batchUpdate, C, hex } = require('./lib');
const { google } = require('googleapis');
const fs = require('fs');

const TABS = [
  { title: '📋 Reference Data',           color: C.gray },
  { title: '📊 Dashboard',                color: C.deepCharcoal },
  { title: '🎙 Content Calendar',          color: C.burntOrange },
  { title: '👥 Guest Outreach',            color: C.mutedSage },
  { title: '💰 Sponsorship Tracker',       color: C.amber },
  { title: '📈 Episode Analytics',         color: C.deepCharcoal },
  { title: '📱 Social Promotion',          color: C.burntOrange },
  { title: '🎧 Equipment Checklist',       color: C.mutedSage },
  { title: '📅 Monthly Calendar',          color: C.amber },
];

(async () => {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Podcast Planner' },
      sheets: TABS.map((t, i) => ({
        properties: {
          sheetId: i,
          title: t.title,
          index: i,
          tabColor: hex(t.color),
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
  console.log('Spreadsheet ID:', id);

  // Set warm cream background on all sheets
  const reqs = TABS.map((_, i) => ({
    repeatCell: {
      range: { sheetId: i, startRowIndex: 0, endRowIndex: 1000, startColumnIndex: 0, endColumnIndex: 26 },
      cell: { userEnteredFormat: { backgroundColor: hex(C.warmCream) } },
      fields: 'userEnteredFormat.backgroundColor',
    },
  }));
  await batchUpdate(id, reqs, 'bg');
  console.log('Created:', id);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
