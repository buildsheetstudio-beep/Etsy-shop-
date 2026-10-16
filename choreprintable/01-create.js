'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const { batchUpdate, gridRange, hex, C } = require('./lib');

function getAuth() {
  const secret = JSON.parse(fs.readFileSync(__dirname + '/client_secret.json'));
  const { client_id, client_secret, redirect_uris } = secret.installed;
  const auth = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  auth.setCredentials(JSON.parse(fs.readFileSync(__dirname + '/tokens.json')));
  return auth;
}

(async () => {
  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Printable Family Chore Chart — Weekly Household Responsibility Planner' },
      sheets: [
        { properties: { title: 'Reference Data',     sheetId: 0, index: 0 } },
        { properties: { title: 'Weekly Chore Chart', sheetId: 1, index: 1 } },
      ]
    }
  });

  const id = res.data.spreadsheetId;
  const sheetMap = {};
  for (const s of res.data.sheets) sheetMap[s.properties.title] = s.properties.sheetId;

  // Baseline background
  const reqs = [];
  for (const sid of Object.values(sheetMap)) {
    reqs.push({ repeatCell: {
      range: gridRange(sid, 0, 1000, 0, 26),
      cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
      fields: 'userEnteredFormat.backgroundColor'
    } });
  }
  await batchUpdate(id, reqs, 'bg');

  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({ id, sheetMap }, null, 2));
  console.log('Created:', id);
  console.log('Sheets:', sheetMap);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
