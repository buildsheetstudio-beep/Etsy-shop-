'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const lib = require('./lib');

const SECRET = JSON.parse(fs.readFileSync(__dirname + '/client_secret.json'));
const creds = SECRET.installed || SECRET.web;
const oAuth2 = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
oAuth2.setCredentials(JSON.parse(fs.readFileSync(__dirname + '/tokens.json')));
const sheets = google.sheets({ version: 'v4', auth: oAuth2 });

const TABS = [
  { title: '📋 Reference Data',                     color: '#9E9E9E' },
  { title: '📊 Dashboard',                           color: '#6B4A34' },
  { title: '👨‍👩‍👧‍👦 Family Members',                   color: '#B8925A' },
  { title: '🌳 Family Tree / Relationship Viewer',   color: '#7A9B7A' },
  { title: '📸 Photos & Notes',                      color: '#5C6BC0' },
  { title: '📜 Source & Citation Tracker',           color: '#8D6E63' },
  { title: '🧬 DNA & Ethnicity Estimate',            color: '#4A7A9B' },
  { title: '🚢 Immigration & Migration Tracker',     color: '#2E7D5E' },
  { title: '📞 Living Relatives Directory',          color: '#6A4C93' },
];

(async () => {
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Genealogy & Family Tree Tracker' },
      sheets: TABS.map((t, i) => ({
        properties: {
          sheetId: i,
          title: t.title,
          index: i,
          tabColorStyle: { rgbColor: { red: parseInt(t.color.slice(1,3),16)/255, green: parseInt(t.color.slice(3,5),16)/255, blue: parseInt(t.color.slice(5,7),16)/255 } },
        },
      })),
    },
  });

  const id = res.data.spreadsheetId;
  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });

  // Global background: parchment
  const reqs = [];
  TABS.forEach((t, i) => {
    reqs.push({
      repeatCell: {
        range: { sheetId: i, startRowIndex: 0, endRowIndex: 1000, startColumnIndex: 0, endColumnIndex: 26 },
        cell: { userEnteredFormat: { backgroundColor: { red: 0.957, green: 0.925, blue: 0.867 } } },
        fields: 'userEnteredFormat(backgroundColor)',
      },
    });
  });
  await sheets.spreadsheets.batchUpdate({ spreadsheetId: id, requestBody: { requests: reqs } });

  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({ id, sheetMap }, null, 2));
  console.log('Created:', id);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
