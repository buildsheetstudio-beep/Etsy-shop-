'use strict';
const { hex, C, batchUpdate } = require('./lib');
const fs = require('fs');
const { google } = require('googleapis');

const SECRET = JSON.parse(fs.readFileSync(__dirname + '/client_secret.json'));
const creds = SECRET.installed || SECRET.web;
const oAuth2 = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
oAuth2.setCredentials(JSON.parse(fs.readFileSync(__dirname + '/tokens.json')));
const sheets = google.sheets({ version: 'v4', auth: oAuth2 });

const TABS = [
  { title: '📋 Reference Data',              color: '#808080' },
  { title: '✨ Dashboard',                   color: '#A13860' },
  { title: '📅 Monthly Calendar View',       color: '#7FB8A4' },
  { title: '💰 Budget & Expense Tracker',    color: '#D9A548' },
  { title: '✅ Task Manager',                color: '#7A9B7A' },
  { title: '🔁 Habit Tracker',              color: '#A13860' },
  { title: '⚖️ Life Balance Wheel',          color: '#7FB8A4' },
  { title: '📝 Weekly & Monthly Reflection', color: '#D9A548' },
  { title: '📓 Journal & Gratitude Log',     color: '#7A9B7A' },
];

(async () => {
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Life Planner' },
      sheets: TABS.map((t, i) => ({
        properties: {
          sheetId: i, index: i, title: t.title,
          tabColorStyle: { rgbColor: hex(t.color) },
          gridProperties: { rowCount: 200, columnCount: 26 },
        },
      })),
    },
  });
  const id = res.data.spreadsheetId;
  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });
  const bgReqs = TABS.map((t, i) => ({
    repeatCell: {
      range: { sheetId: i, startRowIndex: 0, endRowIndex: 200, startColumnIndex: 0, endColumnIndex: 26 },
      cell: { userEnteredFormat: { backgroundColor: hex(C.offWhite) } },
      fields: 'userEnteredFormat.backgroundColor',
    },
  }));
  await batchUpdate(id, bgReqs, 'bg');
  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({ id, sheetMap }, null, 2));
  console.log('Created:', id);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
