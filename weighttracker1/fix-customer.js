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

  // ── Get sheet IDs ─────────────────────────────────────────────────────────
  console.log('Reading spreadsheet metadata...');
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets(properties(sheetId,title))',
  });
  const sheetMap = {};
  for (const s of meta.data.sheets) {
    sheetMap[s.properties.title] = s.properties.sheetId;
    console.log(`  "${s.properties.title}" → ${s.properties.sheetId}`);
  }

  // Find tab names (may have emoji prefixes)
  const tabDashboard     = Object.keys(sheetMap).find(t => t.includes('Dashboard'));
  const tabEventsLog     = Object.keys(sheetMap).find(t => t.includes('Events Log'));
  const tabGuestList     = Object.keys(sheetMap).find(t => t.includes('Guest'));
  const tabBudget        = Object.keys(sheetMap).find(t => t.includes('Budget'));
  const tabTaskTracker   = Object.keys(sheetMap).find(t => t.includes('Task'));
  const tabVendorTracker = Object.keys(sheetMap).find(t => t.includes('Vendor'));
  const tabAnalytics     = Object.keys(sheetMap).find(t => t.includes('Analytics'));

  console.log(`\nResolved tabs:`);
  console.log(`  Dashboard:     "${tabDashboard}"`);
  console.log(`  Events Log:    "${tabEventsLog}"`);
  console.log(`  Guest List:    "${tabGuestList}"`);
  console.log(`  Budget:        "${tabBudget}"`);
  console.log(`  Task Tracker:  "${tabTaskTracker}"`);
  console.log(`  Vendor:        "${tabVendorTracker}"`);
  console.log(`  Analytics:     "${tabAnalytics}"`);

  // ── FIX 1: Dashboard B8 — correct FILTER range from B3:B65 → B3:Q65 ──────
  // B8 currently: FILTER('📋 Events Log'!B3:B65,...) — 1 column, SORT col 4 invalid
  // B9-B17 pattern: FILTER('📋 Events Log'!B3:Q65,'📋 Events Log'!E$3:E$65>=TODAY(),...)
  console.log('\nFix 1: Dashboard B8 formula...');
  const b8Formula =
    `=IFERROR(INDEX(SORT(FILTER('${tabEventsLog}'!B$3:Q$65,'${tabEventsLog}'!E$3:E$65>=TODAY(),'${tabEventsLog}'!E$3:E$65<=TODAY()+90),4,1),1,1),"")`;
  console.log(`  New B8: ${b8Formula}`);

  // ── FIX 2: Dashboard K4 — extend budget range from E4:E15 → E4:E200 ───────
  console.log('Fix 2: Dashboard K4 formula...');
  const k4Formula = `=SUM('${tabBudget}'!E4:E200)`;
  console.log(`  New K4: ${k4Formula}`);

  // ── FIX 3: Events Log col I (Confirmed Guests) — fix COUNTIFS refs ─────────
  // Wrong: COUNTIFS('👥 Guest List & RSVP'!C$3:C$65, B3, '👥 Guest List & RSVP'!F$3:F$65, "Confirmed")
  //         → col C = Last Name, rows 3-4 = headers/stats
  // Right: COUNTIFS('👥 Guest List & RSVP'!D$5:D$65, B3, '👥 Guest List & RSVP'!F$5:F$65, "Confirmed")
  //         → col D = Event Name, rows 5+ = data
  // Apply to I3:I65 (63 rows)
  console.log('Fix 3: Events Log col I (Confirmed Guests) formulas...');
  const iFormulas = [];
  for (let row = 3; row <= 65; row++) {
    iFormulas.push([
      `=IFERROR(COUNTIFS('${tabGuestList}'!D$5:D$65,B${row},'${tabGuestList}'!F$5:F$65,"Confirmed"),0)`
    ]);
  }

  // ── CLEAR OLD SAMPLE DATA ─────────────────────────────────────────────────
  // Task Tracker: data in rows 5-65
  // Guest List: data in rows 5-65
  // Budget allocation rows: 4-15; expense log rows 19-200
  // Vendor Tracker: data in rows 5-65
  // Analytics: data in rows 5-35

  console.log('\nBuilding clear + formula batches...');

  const valueUpdates = [
    // Fix 1: Dashboard B8
    { range: `'${tabDashboard}'!B8`, values: [[b8Formula]] },
    // Fix 2: Dashboard K4
    { range: `'${tabDashboard}'!K4`, values: [[k4Formula]] },
    // Fix 3: Events Log I3:I65 confirmed guests
    { range: `'${tabEventsLog}'!I3:I65`, values: iFormulas },
  ];

  // Clear Task Tracker data rows (keep headers)
  if (tabTaskTracker) {
    console.log(`  Queuing Task Tracker clear (rows 5-65)...`);
    valueUpdates.push({ range: `'${tabTaskTracker}'!A5:Z65`, values: Array(61).fill(Array(26).fill('')) });
  }

  // Clear Guest List data rows (keep headers/stats in rows 1-4)
  if (tabGuestList) {
    console.log(`  Queuing Guest List clear (rows 5-65)...`);
    valueUpdates.push({ range: `'${tabGuestList}'!A5:Z65`, values: Array(61).fill(Array(26).fill('')) });
  }

  // Clear Budget allocation rows AND expense log rows (preserve header structure)
  if (tabBudget) {
    console.log(`  Queuing Budget clear (allocation rows 4-15, expense log rows 19-200)...`);
    valueUpdates.push({ range: `'${tabBudget}'!A4:G15`, values: Array(12).fill(Array(7).fill('')) });
    valueUpdates.push({ range: `'${tabBudget}'!A19:G200`, values: Array(182).fill(Array(7).fill('')) });
  }

  // Clear Vendor Tracker data rows
  if (tabVendorTracker) {
    console.log(`  Queuing Vendor Tracker clear (rows 5-65)...`);
    valueUpdates.push({ range: `'${tabVendorTracker}'!A5:Z65`, values: Array(61).fill(Array(26).fill('')) });
  }

  // Clear Analytics data rows
  if (tabAnalytics) {
    console.log(`  Queuing Analytics clear (rows 5-35)...`);
    valueUpdates.push({ range: `'${tabAnalytics}'!A5:Z35`, values: Array(31).fill(Array(26).fill('')) });
  }

  // ── Execute value updates in chunks ──────────────────────────────────────
  const VCHUNK = 20;
  console.log(`\nApplying ${valueUpdates.length} value updates...`);
  for (let i = 0; i < valueUpdates.length; i += VCHUNK) {
    const chunk = valueUpdates.slice(i, i + VCHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: 'USER_ENTERED', data: chunk },
    });
    console.log(`  Chunk ${Math.floor(i / VCHUNK) + 1}/${Math.ceil(valueUpdates.length / VCHUNK)} ✓`);
  }

  // ── Verify fixes ──────────────────────────────────────────────────────────
  console.log('\nVerifying fixes...');
  const check = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: [
      `'${tabDashboard}'!B8`,
      `'${tabDashboard}'!K4`,
      `'${tabEventsLog}'!I3`,
    ],
    valueRenderOption: 'FORMULA',
  });
  for (const vr of check.data.valueRanges) {
    const val = (vr.values || [['']])[0][0];
    console.log(`  ${vr.range}: ${val}`);
  }

  console.log('\n✅  All fixes applied:');
  console.log('    Dashboard B8      — FILTER range corrected (B3:Q65, anchored rows)');
  console.log('    Dashboard K4      — Budget SUM extended to E4:E200');
  console.log('    Events Log I3:I65 — Confirmed Guests COUNTIFS now refs col D rows 5+');
  console.log('    Task Tracker      — Sample data cleared (rows 5-65)');
  console.log('    Guest List        — Sample data cleared (rows 5-65)');
  console.log('    Budget            — Sample data cleared (alloc rows 4-15, expense rows 19-200)');
  console.log('    Vendor Tracker    — Sample data cleared (rows 5-65)');
  console.log('    Analytics         — Sample data cleared (rows 5-35)');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
