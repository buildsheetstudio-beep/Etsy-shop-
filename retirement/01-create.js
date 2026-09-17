'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SECRET_PATH = path.join(__dirname, 'client_secret.json');
const TOKEN_PATH  = path.join(__dirname, 'tokens.json');

function getAuth() {
  const credentials = JSON.parse(fs.readFileSync(SECRET_PATH));
  const { client_id, client_secret, redirect_uris } = credentials.installed || credentials.web;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
  return oAuth2Client;
}

(async () => {
  const sheets = google.sheets({ version: 'v4', auth: getAuth() });

  const tabNames = [
    'Reference Data',
    'Party Setup',
    'Guest List & RSVP',
    'Invitation Tracker',
    'Budget & Expenses',
    'Entertainment & Music',
    'Thank-You Cards',
    'Prep & Party Day',
    'Dashboard',
  ];

  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Retirement Party Planner' },
      sheets: tabNames.map((title, i) => ({
        properties: { sheetId: i, title, index: i }
      }))
    }
  });

  const id = res.data.spreadsheetId;
  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });

  fs.writeFileSync(path.join(__dirname, 'spreadsheet.json'), JSON.stringify({ id, sheetMap }, null, 2));
  console.log('ID:', id);
  console.log('URL: https://docs.google.com/spreadsheets/d/' + id);
  console.log('SheetMap:', JSON.stringify(sheetMap));
})();
