'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1bvkKXEI2xpKpyaiOLBBr0LoSv9EwOYlE-DcLfcvwrjk';

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}
async function getSheets() { return google.sheets({ version: 'v4', auth: await getAuth() }); }

function colIndexToLetter(idx) {
  let result = '', i = idx + 1;
  while (i > 0) { const rem=(i-1)%26; result=String.fromCharCode(65+rem)+result; i=Math.floor((i-1)/26); }
  return result;
}

(async () => {
  const sheets = await getSheets();

  // ── A. Progress & Performance — full formula scan for hardcoded years ────
  console.log('── A. PROGRESS & PERFORMANCE — full scan ───────────────');
  const ppData = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Progress & Performance'!A1:T120",
    valueRenderOption: 'FORMULA',
  });
  const ppRows = ppData.data.values || [];
  const hardcodedYear = [];
  const maxifsCells = [];
  ppRows.forEach((row, r) => {
    row.forEach((cell, c) => {
      if (typeof cell !== 'string' || !cell.startsWith('=')) return;
      if (/=202[0-9]/.test(cell) || /=20[1-9][0-9]/.test(cell)) {
        hardcodedYear.push({ cell: `${colIndexToLetter(c)}${r+1}`, formula: cell.slice(0, 120) });
      }
      if (cell.toUpperCase().includes('MAXIFS(')) {
        maxifsCells.push({ cell: `${colIndexToLetter(c)}${r+1}`, formula: cell.slice(0, 100) });
      }
    });
  });
  console.log(`  Hardcoded year cells: ${hardcodedYear.length}`);
  hardcodedYear.slice(0, 8).forEach(({ cell, formula }) => console.log(`    ${cell}: ${formula}`));
  console.log(`\n  MAXIFS cells: ${maxifsCells.length}`);
  maxifsCells.slice(0, 4).forEach(({ cell, formula }) => console.log(`    ${cell}: ${formula}`));

  // Is there a year selector cell?
  const ppVals = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Progress & Performance'!A1:T10",
    valueRenderOption: 'FORMATTED_VALUE',
  });
  console.log('\n  Top rows (displayed values):');
  (ppVals.data.values || []).slice(0,10).forEach((row, r) => {
    if (row.some(c => c)) console.log(`    row ${r+1}: ${row.map(c=>JSON.stringify(c)).join('  ')}`);
  });

  // ── B. Fitness Dashboard — map ALL formulas to find year pattern ─────────
  console.log('\n── B. FITNESS DASHBOARD — year pattern + structure ────');
  const dashData = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Fitness Dashboard'!A1:T60",
    valueRenderOption: 'FORMULA',
  });
  const dashRows = dashData.data.values || [];
  const dashHardYear = [];
  const dashMaxifs = [];
  dashRows.forEach((row, r) => {
    row.forEach((cell, c) => {
      if (typeof cell !== 'string' || !cell.startsWith('=')) return;
      if (/=202[0-9]/.test(cell) || (cell.includes('2025') && !cell.includes('$B$3') && !cell.includes('$B$2'))) {
        dashHardYear.push({ cell: `${colIndexToLetter(c)}${r+1}`, formula: cell.slice(0,100) });
      }
      if (cell.toUpperCase().includes('MAXIFS(') && cell.includes('"EX-')) {
        dashMaxifs.push({ cell: `${colIndexToLetter(c)}${r+1}`, formula: cell.slice(0,100) });
      }
    });
  });
  const dashVals = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Fitness Dashboard'!A1:T15",
    valueRenderOption: 'FORMATTED_VALUE',
  });
  console.log('  Displayed header area:');
  (dashVals.data.values || []).slice(0,10).forEach((row, r) => {
    if (row.some(c => c)) console.log(`    row ${r+1}: ${row.map(c=>String(c||'').slice(0,20)).filter(c=>c).join('  ')}`);
  });
  console.log(`\n  Hardcoded year formulas: ${dashHardYear.length}`);
  dashHardYear.slice(0, 3).forEach(({ cell, formula }) => console.log(`    ${cell}: ${formula}`));
  console.log(`\n  MAXIFS with hardcoded EX- ID: ${dashMaxifs.length}`);
  dashMaxifs.slice(0, 6).forEach(({ cell, formula }) => console.log(`    ${cell}: ${formula}`));

  // ── C. Check all unique EX- IDs hardcoded in Dashboard ───────────────────
  console.log('\n── C. UNIQUE HARDCODED EXERCISE IDs IN DASHBOARD ──────');
  const exIds = new Set();
  dashRows.forEach(row => {
    row.forEach(cell => {
      if (typeof cell !== 'string') return;
      const matches = cell.match(/"EX-\d+"/g) || [];
      matches.forEach(m => exIds.add(m));
    });
  });
  console.log(`  Unique EX- IDs referenced: ${[...exIds].sort().join(', ')}`);

  // ── D. Yearly Fitness Planner — hardcoded year check ────────────────────
  console.log('\n── D. YEARLY FITNESS PLANNER — year check ─────────────');
  const yfpData = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Yearly Fitness Planner'!A1:T30",
    valueRenderOption: 'FORMULA',
  });
  const yfpHardYear = [];
  (yfpData.data.values || []).forEach((row, r) => {
    row.forEach((cell, c) => {
      if (typeof cell !== 'string' || !cell.startsWith('=')) return;
      if (cell.includes('$B$2')) return; // dynamic year ref — OK
      if (/=202[0-9]/.test(cell) || cell.match(/\(202[0-9]\)/)) {
        yfpHardYear.push({ cell: `${colIndexToLetter(c)}${r+1}`, formula: cell.slice(0,100) });
      }
    });
  });
  const yfpVals = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Yearly Fitness Planner'!A1:T8",
    valueRenderOption: 'FORMATTED_VALUE',
  });
  console.log('  Header area:');
  (yfpVals.data.values || []).slice(0,6).forEach((row,r) => {
    if (row.some(c=>c)) console.log(`    row ${r+1}: ${row.map(c=>String(c||'').slice(0,20)).filter(c=>c).join('  ')}`);
  });
  console.log(`  Hardcoded year formulas: ${yfpHardYear.length}`);
  yfpHardYear.slice(0,3).forEach(({cell,formula}) => console.log(`    ${cell}: ${formula}`));

  // ── E. Personal Records — MAXIFS pattern ─────────────────────────────────
  console.log('\n── E. PERSONAL RECORDS — MAXIFS full count ────────────');
  const prData = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Personal Records'!A1:P600",
    valueRenderOption: 'FORMULA',
  });
  let prMaxifs = 0, prRows2 = 0;
  (prData.data.values || []).forEach(row => {
    let hasData = false;
    row.forEach(cell => {
      if (typeof cell === 'string' && cell.toUpperCase().includes('MAXIFS(')) { prMaxifs++; hasData=true; }
    });
    if (hasData) prRows2++;
  });
  console.log(`  MAXIFS cells: ${prMaxifs} across ${prRows2} data rows`);
  // Sample one full row of formulas
  const prSample = (prData.data.values || [])[3] || [];
  console.log('  Row 4 formulas:');
  prSample.forEach((cell, c) => {
    if (typeof cell === 'string' && cell.startsWith('=')) {
      console.log(`    ${colIndexToLetter(c)}4: ${cell.slice(0,110)}`);
    }
  });

  // ── F. Daily Workout Log — check column structure ─────────────────────────
  console.log('\n── F. DAILY WORKOUT LOG — column headers + DV scope ────');
  const dwlHeaders = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Daily Workout Log'!A1:AD8",
    valueRenderOption: 'FORMATTED_VALUE',
  });
  (dwlHeaders.data.values || []).slice(0,7).forEach((row,r) => {
    if (row.some(c=>c)) console.log(`  row ${r+1}: ${row.map((c,i)=>`${colIndexToLetter(i)}:${String(c||'').slice(0,15)}`).filter(c=>!c.endsWith(':')).join('  ')}`);
  });

  console.log('\n✅  Targeted scan complete.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
