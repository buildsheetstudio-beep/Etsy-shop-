'use strict';
/**
 * Step 1: Create the spreadsheet with all 6 sheets (sheetId 0-5).
 * Saves actual IDs to budget-config.json for use by subsequent scripts.
 */
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const AUTH_DIR = path.join(__dirname, '..');

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(AUTH_DIR, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const auth = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  auth.setCredentials(JSON.parse(fs.readFileSync(path.join(AUTH_DIR, 'tokens.json'))));
  try {
    const r = await auth.refreshAccessToken();
    fs.writeFileSync(path.join(AUTH_DIR, 'tokens.json'), JSON.stringify(r.credentials, null, 2));
  } catch (e) { /* already fresh */ }
  return auth;
}

(async () => {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  console.log('Creating spreadsheet…');
  const resp = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate 50/30/20 Budget Planner — BuildSheetStudio' },
      sheets: [
        { properties: { sheetId: 0, index: 0, title: 'Reference Data',    sheetType: 'GRID', gridProperties: { rowCount: 200, columnCount: 26 } } },
        { properties: { sheetId: 1, index: 1, title: 'Start Here',         sheetType: 'GRID', gridProperties: { rowCount: 100, columnCount: 15 } } },
        { properties: { sheetId: 2, index: 2, title: 'Transactions',       sheetType: 'GRID', gridProperties: { rowCount: 1050, columnCount: 32 } } },
        { properties: { sheetId: 3, index: 3, title: 'Bills & Goals',      sheetType: 'GRID', gridProperties: { rowCount: 450, columnCount: 14 } } },
        { properties: { sheetId: 4, index: 4, title: 'Yearly Insights',    sheetType: 'GRID', gridProperties: { rowCount: 200, columnCount: 20 } } },
        { properties: { sheetId: 5, index: 5, title: 'Budget Dashboard',   sheetType: 'GRID', gridProperties: { rowCount: 120, columnCount: 16 } } },
      ],
    },
  });

  const ss = resp.data;
  const id = ss.spreadsheetId;
  const url = `https://docs.google.com/spreadsheets/d/${id}/edit`;

  // Retrieve actual assigned sheet IDs
  const config = { spreadsheetId: id, url, sheets: {} };
  ss.sheets.forEach(s => {
    const p = s.properties;
    config.sheets[p.title] = p.sheetId;
    console.log(`  Sheet "${p.title}": id=${p.sheetId} index=${p.index}`);
  });

  fs.writeFileSync(path.join(__dirname, 'budget-config.json'), JSON.stringify(config, null, 2));
  console.log('\n✅ Spreadsheet created:', url);
  console.log('   Config saved to budget-config.json');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
