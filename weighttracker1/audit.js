'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1YSrKj6U3Zb_al3g7j9e-80qFRGpQVwD4JLe4gh1Ah2k';

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}
async function getSheets() { return google.sheets({ version: 'v4', auth: await getAuth() }); }

(async () => {
  const sheets = await getSheets();

  const ranges = [
    // Dashboard — all formula rows
    "'🎉 Dashboard'!A1:P40",
    // Events Log — headers + first data row + formula row
    "'📋 Events Log'!A1:Q5",
    "'📋 Events Log'!A3:Q65",
    // Task Tracker — headers + stats row + data sample
    "'✅ Task Tracker'!A1:J6",
    // Budget — full structure
    "'💰 Event Budget'!A1:H25",
    // Guest List — headers + stats
    "'👥 Guest List & RSVP'!A1:T6",
    // Vendor — headers + stats
    "'🏪 Vendor & Supplier Tracker'!A1:L6",
    // Analytics — headers + stats
    "'📊 Event Analytics'!A1:M6",
    // Run-of-Show — top section
    "'⏱️ Run-of-Show'!A1:J5",
  ];

  const resp = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges,
    valueRenderOption: 'FORMULA',
  });

  for (const vr of resp.data.valueRanges) {
    console.log(`\n====== ${vr.range} ======`);
    (vr.values || []).forEach((r, i) => {
      if (r.some(v => v !== '')) console.log(`  Row ${i+1}: ${JSON.stringify(r)}`);
    });
  }

  // Also check displayed values for dashboard
  const dashDisp = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'🎉 Dashboard'!A1:P35",
    valueRenderOption: 'FORMATTED_VALUE',
  });
  console.log('\n====== DASHBOARD displayed values ======');
  (dashDisp.data.values || []).forEach((r, i) => {
    if (r.some(v => v !== '')) console.log(`  Row ${i+1}: ${JSON.stringify(r)}`);
  });

  // Check Event Analytics stats row
  const anaDisp = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'📊 Event Analytics'!A1:M12",
    valueRenderOption: 'FORMATTED_VALUE',
  });
  console.log('\n====== EVENT ANALYTICS displayed ======');
  (anaDisp.data.values || []).forEach((r, i) => {
    if (r.some(v => v !== '')) console.log(`  Row ${i+1}: ${JSON.stringify(r)}`);
  });

  // Check Budget displayed
  const budDisp = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'💰 Event Budget'!A1:H26",
    valueRenderOption: 'FORMATTED_VALUE',
  });
  console.log('\n====== BUDGET displayed ======');
  (budDisp.data.values || []).forEach((r, i) => {
    if (r.some(v => v !== '')) console.log(`  Row ${i+1}: ${JSON.stringify(r)}`);
  });
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
