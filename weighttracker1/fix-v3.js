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
  const valueUpdates = [];

  // ── 1. Clear orphan row-numbers in Events Log A11:A14 ────────────────────
  // Events 7 & 8 are in rows 9-10 (valid). Rows 11-14 have A=9,10,11,12 with no event data.
  valueUpdates.push({
    range: "'📋 Events Log'!A11:A14",
    values: [[''], [''], [''], ['']],
  });
  console.log('1. Events Log A11:A14 — orphan row numbers queued to clear');

  // ── 2. RSVP Summary: fix S:V formulas for rows 3-10 ──────────────────────
  // Rows 3-4 had hardcoded "Tech Summit 2025" / "Charity Gala 2025".
  // Rows 5-10 had event labels in R col but no count formulas at all.
  // Fix: reference R column dynamically so any label change auto-updates counts.
  const rsvpData = [];
  for (let row = 3; row <= 10; row++) {
    rsvpData.push([
      // S = Invited (total with that event name)
      `=IFERROR(COUNTIF('👥 Guest List & RSVP'!D$5:D$65,R${row}),0)`,
      // T = Confirmed
      `=IFERROR(COUNTIFS('👥 Guest List & RSVP'!D$5:D$65,R${row},'👥 Guest List & RSVP'!F$5:F$65,"Confirmed"),0)`,
      // U = Declined
      `=IFERROR(COUNTIFS('👥 Guest List & RSVP'!D$5:D$65,R${row},'👥 Guest List & RSVP'!F$5:F$65,"Declined"),0)`,
      // V = Pending
      `=IFERROR(COUNTIFS('👥 Guest List & RSVP'!D$5:D$65,R${row},'👥 Guest List & RSVP'!F$5:F$65,"Pending"),0)`,
    ]);
  }
  valueUpdates.push({ range: "'👥 Guest List & RSVP'!S3:V10", values: rsvpData });
  console.log('2. RSVP Summary S3:V10 — dynamic formulas queued for rows 3-10');

  // ── 3. Analytics: add satisfaction scores for 3 completed events ──────────
  // Tech Summit (row 5) = 4.6, Team Building (row 8) = 4.4, Marketing Workshop (row 10) = 4.7
  valueUpdates.push({ range: "'📊 Event Analytics'!M5", values: [[4.6]] });
  valueUpdates.push({ range: "'📊 Event Analytics'!M8", values: [[4.4]] });
  valueUpdates.push({ range: "'📊 Event Analytics'!M10", values: [[4.7]] });
  console.log('3. Analytics M5/M8/M10 — satisfaction scores queued');

  // ── 4. Analytics row 3: fix Avg Rating and Avg Satisfaction to exclude zeros
  // F3 = Avg Rating, G3 = Avg Satisfaction
  valueUpdates.push({
    range: "'📊 Event Analytics'!F3",
    values: [['=IFERROR(AVERAGEIF(L5:L35,">"&0),0)']],
  });
  valueUpdates.push({
    range: "'📊 Event Analytics'!G3",
    values: [['=IFERROR(AVERAGEIF(M5:M35,">"&0),0)']],
  });
  console.log('4. Analytics F3/G3 — AVERAGEIF formulas queued');

  // ── 5. Dashboard N16/N17: fix Avg Rating and Avg Satisfaction ────────────
  // These mirror the Analytics formulas but reference the Analytics sheet.
  valueUpdates.push({
    range: "'🎉 Dashboard'!N16",
    values: [['=IFERROR(AVERAGEIF(\'📊 Event Analytics\'!L5:L35,">"&0),0)']],
  });
  valueUpdates.push({
    range: "'🎉 Dashboard'!N17",
    values: [['=IFERROR(AVERAGEIF(\'📊 Event Analytics\'!M5:M35,">"&0),0)']],
  });
  console.log('5. Dashboard N16/N17 — AVERAGEIF formulas queued');

  // ── 6. Budget: add 3 missing expense log entries ──────────────────────────
  // These cover actuals already entered manually in F6/F7/F15 but absent from the log.
  // After F4:F15 become SUMIFS, these entries will drive those actuals automatically.
  valueUpdates.push({
    range: "'💰 Event Budget'!A26:H28",
    values: [
      [7, 'Tech Summit 2026',    'AV & Tech', 'TechPro AV Services', '9/15/2026', 2500, 'AV equipment + technician', true],
      [8, 'Tech Summit 2026',    'Marketing', 'Creative Agency',      '9/1/2026',  1200, 'Marketing materials',       true],
      [9, 'Marketing Workshop',  'Catering',  'Café Solutions',       '7/12/2026', 250,  'Workshop catering',         true],
    ],
  });
  console.log('6. Budget expense log rows 26-28 — 3 missing transactions queued');

  // ── 7. Budget F4:F15: replace manual actuals with SUMIFS from expense log
  // Each cell sums expense log amounts where Event Name + Category match this row.
  const budgetActualFormulas = Array.from({ length: 12 }, (_, i) => {
    const r = i + 4; // rows 4-15
    return [`=IFERROR(SUMIFS($F$20:$F$100,$B$20:$B$100,B${r},$C$20:$C$100,C${r}),0)`];
  });
  valueUpdates.push({ range: "'💰 Event Budget'!F4:F15", values: budgetActualFormulas });
  console.log('7. Budget F4:F15 — SUMIFS formulas queued');

  // ── Execute all value updates ─────────────────────────────────────────────
  console.log('\nApplying updates...');
  const CHUNK = 10;
  for (let i = 0; i < valueUpdates.length; i += CHUNK) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: 'USER_ENTERED', data: valueUpdates.slice(i, i + CHUNK) },
    });
    console.log(`  Chunk ${Math.floor(i / CHUNK) + 1} ✓`);
  }

  // ── Verify ────────────────────────────────────────────────────────────────
  console.log('\nVerifying...');
  const chk = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: [
      "'🎉 Dashboard'!N16:N17",
      "'📊 Event Analytics'!F3:G3",
      "'📊 Event Analytics'!M5:M12",
      "'👥 Guest List & RSVP'!R3:V5",
      "'💰 Event Budget'!F4:F15",
      "'💰 Event Budget'!A26:B28",
      "'📋 Events Log'!A9:A14",
    ],
    valueRenderOption: 'FORMATTED_VALUE',
  });
  for (const vr of chk.data.valueRanges) {
    console.log(`  ${vr.range}: ${JSON.stringify(vr.values)}`);
  }

  console.log('\n✅  Auto-update fixes complete:');
  console.log('    Events Log A11:A14       — orphan row numbers cleared');
  console.log('    RSVP Summary S3:V10      — all 8 events now use dynamic R-column references');
  console.log('    Analytics M5/M8/M10      — satisfaction scores added for completed events');
  console.log('    Analytics F3/G3          — AVERAGEIF now excludes zero-rated future events');
  console.log('    Dashboard N16/N17        — AVERAGEIF now excludes zero-rated future events');
  console.log('    Budget expense log       — 3 missing transactions added (rows 26-28)');
  console.log('    Budget F4:F15            — SUMIFS now auto-pulls actuals from expense log');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
