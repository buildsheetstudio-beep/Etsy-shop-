'use strict';
const { google } = require('googleapis');
const fs = require('fs');

function getAuth() {
  const s = JSON.parse(fs.readFileSync(__dirname + '/client_secret.json'));
  const { client_id, client_secret, redirect_uris } = s.installed;
  const auth = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  auth.setCredentials(JSON.parse(fs.readFileSync(__dirname + '/tokens.json')));
  return auth;
}

const TABS = [
  { name: 'Reference Data',        id: 0 },
  { name: 'Team Setup',            id: 1 },
  { name: 'Master Task Log',       id: 2 },
  { name: 'Recurring Task Planner',id: 3 },
  { name: 'Weekly Calendar',       id: 4 },
  { name: 'Monthly Calendar',      id: 5 },
  { name: 'Project Progress',      id: 6 },
  { name: 'Team Dashboard',        id: 7 },
];

(async () => {
  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Team Task Tracker' },
      sheets: TABS.map(t => ({
        properties: { sheetId: t.id, title: t.name, index: t.id }
      }))
    }
  });

  const id = res.data.spreadsheetId;
  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });

  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({ id, sheetMap }, null, 2));
  console.log('Created:', id);
  console.log('URL: https://docs.google.com/spreadsheets/d/' + id);
})().catch(e => { console.error(e.message); process.exit(1); });
