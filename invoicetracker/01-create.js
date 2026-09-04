'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const TOKEN_PATH  = path.join(__dirname, 'tokens.json');
const SECRET_PATH = path.join(__dirname, 'client_secret.json');

function getAuth() {
  const { client_secret, client_id, redirect_uris } =
    JSON.parse(fs.readFileSync(SECRET_PATH)).installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
  return oAuth2Client;
}

(async () => {
  const sheets = google.sheets({ version: 'v4', auth: getAuth() });

  const tabs = [
    'Reference Data',
    'Business Setup',
    'Client CRM',
    'Products & Services',
    'Invoice Generator',
    'Invoice Log',
    'Payment Tracker',
    'Invoice — Classic',
    'Invoice — Modern',
    'Invoice — Elegant',
    'Invoice — Compact',
    'Dashboard',
  ];

  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Printable Invoice Tracker & Generator' },
      sheets: tabs.map((title, i) => ({
        properties: {
          sheetId: i,
          title,
          index: i,
          gridProperties: { rowCount: 1200, columnCount: 30 },
        },
      })),
    },
  });

  const id = res.data.spreadsheetId;
  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });

  fs.writeFileSync(path.join(__dirname, 'spreadsheet.json'), JSON.stringify({ id, sheetMap }, null, 2));
  console.log('✓ Created spreadsheet:', id);
  console.log('  URL: https://docs.google.com/spreadsheets/d/' + id);
  console.log('  Tabs:', tabs.join(', '));
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
