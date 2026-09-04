'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
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
  { title:'Reference Data',     sheetId:0 },
  { title:'Party Setup',        sheetId:1 },
  { title:'Guest List & RSVP',  sheetId:2 },
  { title:'Invitation Tracker', sheetId:3 },
  { title:'Budget & Expenses',  sheetId:4 },
  { title:'Food, Drinks & Cake',sheetId:5 },
  { title:'Shopping List',      sheetId:6 },
  { title:'Prep, Setup & Cleanup', sheetId:7 },
  { title:'Dashboard',          sheetId:8 },
];

(async () => {
  const sheets = google.sheets({ version:'v4', auth:getAuth() });

  const res = await sheets.spreadsheets.create({
    requestBody:{
      properties:{ title:'Ultimate Kids Birthday Party Planner' },
      sheets: TABS.map(({ title, sheetId }) => ({
        properties:{ sheetId, title, gridProperties:{ rowCount:300, columnCount:26 } }
      }))
    }
  });

  const spreadsheetId = res.data.spreadsheetId;
  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });

  fs.writeFileSync(path.join(__dirname,'spreadsheet.json'), JSON.stringify({ id:spreadsheetId, sheetMap }, null, 2));
  console.log('Created:', spreadsheetId);
  console.log('URL: https://docs.google.com/spreadsheets/d/' + spreadsheetId);
  console.log('Sheets:', JSON.stringify(sheetMap, null, 2));
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
