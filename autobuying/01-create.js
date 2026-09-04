'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const TOKEN_PATH = path.join(__dirname, 'tokens.json');
const SECRET_PATH = path.join(__dirname, 'client_secret.json');

function getAuth() {
  const credentials = JSON.parse(fs.readFileSync(SECRET_PATH));
  const { client_id, client_secret, redirect_uris } = credentials.installed || credentials.web;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
  return oAuth2Client;
}

const TABS = [
  { title: 'Reference Data',          sheetId: 0 },
  { title: 'Buyer Setup',             sheetId: 1 },
  { title: 'Vehicle Details & Quotes',sheetId: 2 },
  { title: 'Auto Loan Calculator',    sheetId: 3 },
  { title: 'Monthly Budget Impact',   sheetId: 4 },
  { title: 'Total Cost of Ownership', sheetId: 5 },
  { title: 'Lease vs Buy',            sheetId: 6 },
  { title: 'Vehicle Comparison',      sheetId: 7 },
  { title: 'Dashboard',               sheetId: 8 },
];

(async () => {
  const sheets = google.sheets({ version: 'v4', auth: getAuth() });

  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Auto Buying & Budget Comparison' },
      sheets: TABS.map(({ title, sheetId }) => ({
        properties: {
          sheetId,
          title,
          gridProperties: { rowCount: 500, columnCount: 35 }
        }
      }))
    }
  });

  const spreadsheetId = res.data.spreadsheetId;
  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });

  fs.writeFileSync(
    path.join(__dirname, 'spreadsheet.json'),
    JSON.stringify({ id: spreadsheetId, sheetMap }, null, 2)
  );

  console.log('Created:', spreadsheetId);
  console.log('URL: https://docs.google.com/spreadsheets/d/' + spreadsheetId);
  console.log('Sheets:', JSON.stringify(sheetMap, null, 2));
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
