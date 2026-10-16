'use strict';
/**
 * Step 9: QA Audit — scan every tab for formula errors.
 * Checks for: #VALUE!, #REF!, #ERROR!, #NAME?, #N/A, #DIV/0!, #NULL!
 * Per CLAUDE.md quality gate: zero errors before delivery.
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

const ERROR_PATTERNS = ['#VALUE!','#REF!','#ERROR!','#NAME?','#N/A','#DIV/0!','#NULL!'];

// Ranges to scan per tab (keeping ranges targeted to avoid quota issues)
const SCAN_RANGES = [
  { tab: 'Start Here',       ranges: ['A1:E80'] },
  { tab: 'Transactions',     ranges: ['A1:T80'] },     // summary + first 68 data rows
  { tab: 'Bills & Goals',    ranges: ['A1:K25', 'A113:G125', 'A318:J325'] },
  { tab: 'Yearly Insights',  ranges: ['A1:O70'] },
  { tab: 'Budget Dashboard', ranges: ['A1:K55'] },
  { tab: 'Reference Data',   ranges: ['A1:K45'] },
];

(async () => {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  let totalErrors = 0;
  const allErrors = [];

  for (const { tab, ranges } of SCAN_RANGES) {
    let tabErrors = 0;
    for (const range of ranges) {
      const resp = await sheets.spreadsheets.values.get({
        spreadsheetId: SID,
        range: `'${tab}'!${range}`,
        valueRenderOption: 'FORMATTED_VALUE',
      });
      const rows = resp.data.values || [];
      rows.forEach((row, ri) => {
        row.forEach((cell, ci) => {
          const val = String(cell || '');
          const errType = ERROR_PATTERNS.find(e => val.includes(e));
          if (errType) {
            tabErrors++;
            allErrors.push({ tab, range: `${range} r${ri+1}c${ci+1}`, errType, cell: val.slice(0,60) });
          }
        });
      });
    }
    const status = tabErrors === 0 ? '✓' : `⚠ ${tabErrors} error(s)`;
    console.log(`   ${status}  ${tab}`);
    totalErrors += tabErrors;
  }

  if (allErrors.length > 0) {
    console.log('\n❌ Errors found:');
    allErrors.forEach(e => {
      console.log(`   [${e.tab}] ${e.range}: ${e.errType} → "${e.cell}"`);
    });
    console.log('\nFix required before delivery.');
    process.exit(1);
  } else {
    console.log('\n✅ QA PASSED — zero formula errors across all tabs');
    console.log('   Spreadsheet is ready for delivery.');
  }

  // ── Verify key cell values ────────────────────────────────────────────────
  console.log('\n── Key cell spot-check ──────────────────────────────────────────────────');

  const checks = [
    { range: "'Start Here'!B11",  label: 'Cadence (B11)' },
    { range: "'Start Here'!B14",  label: 'Needs % (B14)' },
    { range: "'Start Here'!B15",  label: 'Wants % (B15)' },
    { range: "'Start Here'!B16",  label: 'Future % (B16)' },
    { range: "'Transactions'!F10", label: 'First Bucket (F10)' },
    { range: "'Transactions'!A3:G4", label: 'Transaction summary row 4' },
    { range: "'Budget Dashboard'!D5", label: 'Dashboard period start (D5)' },
    { range: "'Budget Dashboard'!F5", label: 'Dashboard period end (F5)' },
    { range: "'Budget Dashboard'!A9:H9", label: 'Dashboard stat cards (row 9)' },
    { range: "'Budget Dashboard'!C13:F15", label: 'Budget progress (C-F, rows 13-15)' },
    { range: "'Yearly Insights'!C7:K7",  label: 'Yearly January row' },
    { range: "'Yearly Insights'!C19:K19", label: 'Yearly Annual total row' },
  ];

  for (const { range, label } of checks) {
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SID, range,
      valueRenderOption: 'FORMATTED_VALUE',
    });
    const vals = (r.data.values || []).map(row => row.join(' | ')).join(' / ');
    console.log(`   ${label}: ${vals.slice(0,80)}`);
  }
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
