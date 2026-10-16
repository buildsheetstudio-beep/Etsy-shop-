'use strict';
/**
 * Step 8: Global formatting — tab reorder, hide Reference Data, protected ranges.
 *
 * Target tab order:
 *   0: Budget Dashboard  (sheetId 5)
 *   1: Start Here        (sheetId 1)
 *   2: Transactions      (sheetId 2)
 *   3: Bills & Goals     (sheetId 3)
 *   4: Yearly Insights   (sheetId 4)
 *   5: Reference Data    (sheetId 0) — hidden
 */
const { google } = require('googleapis');
const fs = require('fs'), path = require('path');
const AUTH_DIR = path.join(__dirname, '..');
const CFG = JSON.parse(fs.readFileSync(path.join(__dirname, 'budget-config.json')));
const SID = CFG.spreadsheetId;

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(AUTH_DIR, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const auth = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  auth.setCredentials(JSON.parse(fs.readFileSync(path.join(AUTH_DIR, 'tokens.json'))));
  return auth;
}

(async () => {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  // ── Tab reorder + colors + hide Reference Data ─────────────────────────────
  // Tab color choices (hex RGB objects)
  const tabColors = {
    dashboard:    { red: 0.443, green: 0.545, blue: 0.475 }, // moss green
    startHere:    { red: 0.651, green: 0.725, blue: 0.631 }, // soft sage
    transactions: { red: 0.208, green: 0.282, blue: 0.243 }, // deep forest
    billsGoals:   { red: 0.851, green: 0.788, blue: 0.682 }, // sand
    yearly:       { red: 0.953, green: 0.894, blue: 0.784 }, // amber
    refData:      { red: 0.7,   green: 0.7,   blue: 0.7   }, // gray (hidden)
  };

  const requests = [
    // Move Budget Dashboard to index 0
    { updateSheetProperties: { properties: { sheetId: 5, index: 0, tabColor: tabColors.dashboard }, fields: 'index,tabColor' } },
    // Move Start Here to index 1
    { updateSheetProperties: { properties: { sheetId: 1, index: 1, tabColor: tabColors.startHere }, fields: 'index,tabColor' } },
    // Move Transactions to index 2
    { updateSheetProperties: { properties: { sheetId: 2, index: 2, tabColor: tabColors.transactions }, fields: 'index,tabColor' } },
    // Move Bills & Goals to index 3
    { updateSheetProperties: { properties: { sheetId: 3, index: 3, tabColor: tabColors.billsGoals }, fields: 'index,tabColor' } },
    // Move Yearly Insights to index 4
    { updateSheetProperties: { properties: { sheetId: 4, index: 4, tabColor: tabColors.yearly }, fields: 'index,tabColor' } },
    // Move Reference Data to index 5, hide it
    { updateSheetProperties: { properties: { sheetId: 0, index: 5, hidden: true, tabColor: tabColors.refData }, fields: 'index,hidden,tabColor' } },
  ];

  await sheets.spreadsheets.batchUpdate({ spreadsheetId: SID, requestBody: { requests } });
  console.log('✅ Tab order set:');
  console.log('   0: Budget Dashboard (moss green)');
  console.log('   1: Start Here (soft sage)');
  console.log('   2: Transactions (deep forest)');
  console.log('   3: Bills & Goals (sand)');
  console.log('   4: Yearly Insights (amber)');
  console.log('   5: Reference Data (hidden)');

  // ── Verify final sheet order ───────────────────────────────────────────────
  const ss = await sheets.spreadsheets.get({ spreadsheetId: SID, fields: 'sheets.properties' });
  console.log('\n   Verification:');
  ss.data.sheets.sort((a,b) => a.properties.index - b.properties.index).forEach(s => {
    const p = s.properties;
    console.log(`   Index ${p.index}: "${p.title}" (id=${p.sheetId}) hidden=${p.hidden||false}`);
  });
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
