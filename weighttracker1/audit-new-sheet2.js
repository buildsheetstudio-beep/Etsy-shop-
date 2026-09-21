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
  while (i > 0) { const rem = (i-1)%26; result = String.fromCharCode(65+rem)+result; i = Math.floor((i-1)/26); }
  return result;
}

// Pull distinct formula patterns from a range (de-duped by structure, ignoring cell refs)
function extractPatterns(rows, limit = 6) {
  const seen = new Set();
  const patterns = [];
  for (const row of rows) {
    for (const cell of row) {
      if (typeof cell !== 'string' || !cell.startsWith('=')) continue;
      // Normalise cell refs to placeholders for dedup
      const key = cell.replace(/\$?[A-Z]{1,3}\$?\d+/g, 'R').replace(/\d+/g,'N').slice(0, 60);
      if (!seen.has(key)) {
        seen.add(key);
        patterns.push(cell.slice(0, 120));
        if (patterns.length >= limit) return patterns;
      }
    }
  }
  return patterns;
}

(async () => {
  const sheets = await getSheets();

  // ─── A. Sample formulas from each key tab ────────────────────────────────
  const tabs = [
    { name: 'Fitness Dashboard',       range: 'A1:T50'    },
    { name: 'Workout Setup',           range: 'A1:P30'    },
    { name: 'Daily Workout Log',       range: 'A1:AD20'   },
    { name: 'Weekly Planner',          range: 'A1:T30'    },
    { name: 'Monthly Calendar',        range: 'A1:P30'    },
    { name: 'Yearly Fitness Planner',  range: 'A1:T30'    },
    { name: 'Progress & Performance',  range: 'A1:T50'    },
    { name: 'Personal Records',        range: 'A1:P20'    },
  ];

  console.log('═══════════════════════════════════════════════════════════');
  console.log('DEEP AUDIT — Formula samples & DV deep scan');
  console.log('═══════════════════════════════════════════════════════════\n');

  console.log('── A. FORMULA SAMPLES (per tab) ────────────────────────');
  for (const { name, range } of tabs) {
    const data = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${name}'!${range}`,
      valueRenderOption: 'FORMULA',
    });
    const rows = data.data.values || [];
    const patterns = extractPatterns(rows);
    console.log(`\n  "${name}":`);
    patterns.forEach(p => console.log(`    ${p}`));
    if (patterns.length === 0) console.log('    (no formulas in sample range)');
  }

  // ─── B. IFS/MAXIFS deep sample ────────────────────────────────────────────
  console.log('\n\n── B. IFS / MAXIFS SAMPLES ─────────────────────────────');
  const ifsTabs = [
    { name: 'Fitness Dashboard',      range: 'A1:T60'  },
    { name: 'Progress & Performance', range: 'A1:T60'  },
    { name: 'Personal Records',       range: 'A1:P100' },
  ];

  for (const { name, range } of ifsTabs) {
    const data = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${name}'!${range}`,
      valueRenderOption: 'FORMULA',
    });
    const samples = [];
    (data.data.values || []).forEach((row, r) => {
      row.forEach((cell, c) => {
        if (typeof cell !== 'string') return;
        const up = cell.toUpperCase();
        if ((up.includes('IFS(') || up.includes('MAXIFS(')) && samples.length < 3) {
          samples.push(`  ${colIndexToLetter(c)}${r+1}: ${cell.slice(0,120)}`);
        }
      });
    });
    console.log(`\n  "${name}":`);
    samples.forEach(s => console.log(s));
    if (samples.length === 0) console.log('  (none in sample range)');
  }

  // ─── C. Cross-sheet reference format check ───────────────────────────────
  console.log('\n\n── C. CROSS-SHEET REFERENCE CHECK ─────────────────────');
  // Look at Fitness Dashboard formulas that reference other tabs
  const dashData = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Fitness Dashboard'!A1:T100",
    valueRenderOption: 'FORMULA',
  });
  const crossSheetRefs = new Set();
  (dashData.data.values || []).forEach(row => {
    row.forEach(cell => {
      if (typeof cell !== 'string' || !cell.startsWith('=')) return;
      const matches = cell.match(/'[^']+'/g) || [];
      matches.forEach(m => crossSheetRefs.add(m));
    });
  });
  if (crossSheetRefs.size) {
    console.log('  Sheet names referenced in Fitness Dashboard:');
    crossSheetRefs.forEach(r => console.log(`    ${r}`));
  } else {
    console.log('  No cross-sheet references in Fitness Dashboard sample');
  }

  // ─── D. Full DV scan (rows 1–600 on data tabs) ───────────────────────────
  console.log('\n\n── D. DATA VALIDATION — FULL COLUMN SCAN ───────────────');
  const dvTabs = ['Workout Setup','Daily Workout Log','Personal Records'];
  for (const tabName of dvTabs) {
    const dvData = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
      ranges: [`'${tabName}'!A1:AD20`],
      includeGridData: true,
      fields: 'sheets(data(rowData(values(dataValidation))))',
    });
    const dvMap = {};
    (dvData.data.sheets[0]?.data[0]?.rowData || []).forEach((row, r) => {
      (row.values || []).forEach((cell, c) => {
        if (!cell.dataValidation) return;
        const type = cell.dataValidation.condition?.type || '?';
        const src  = (cell.dataValidation.condition?.values || []).map(v => v.userEnteredValue).join('|').slice(0, 50);
        const key  = `${type}::${src}`;
        if (!dvMap[key]) dvMap[key] = { type, src, cols: new Set() };
        dvMap[key].cols.add(colIndexToLetter(c));
      });
    });
    console.log(`\n  "${tabName}":`);
    Object.values(dvMap).forEach(({ type, src, cols }) => {
      console.log(`    cols [${[...cols].join(',')}]: ${type}  ${src}`);
    });
    if (Object.keys(dvMap).length === 0) console.log('    (no DV in A1:AD20)');
  }

  // ─── E. Reference Data content ───────────────────────────────────────────
  console.log('\n\n── E. REFERENCE DATA CONTENT ───────────────────────────');
  const refData = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Reference Data'!A1:T25",
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const refRows = refData.data.values || [];
  refRows.slice(0, 20).forEach((row, r) => {
    if (row.some(c => c !== '')) console.log(`  row ${r+1}: ${row.filter(c=>c).map(c=>JSON.stringify(c)).join('  ')}`);
  });

  // ─── F. Monthly Calendar formula check ───────────────────────────────────
  console.log('\n\n── F. MONTHLY CALENDAR — FULL FORMULA SCAN ─────────────');
  const mcData = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Monthly Calendar'!A1:P100",
    valueRenderOption: 'FORMULA',
  });
  const mcRows = mcData.data.values || [];
  const mcFns = {};
  let mcTotal = 0;
  const mcSamples = [];
  mcRows.forEach((row, r) => {
    row.forEach((cell, c) => {
      if (typeof cell !== 'string' || !cell.startsWith('=')) return;
      mcTotal++;
      const up = cell.toUpperCase();
      ['ARRAYFORMULA','LET','TEXTJOIN','IF','IFERROR','TEXT','DATE','WEEKDAY','MONTH'].forEach(fn => {
        if (up.includes(fn+'(')) mcFns[fn] = (mcFns[fn]||0)+1;
      });
      if (mcSamples.length < 4) mcSamples.push(`${colIndexToLetter(c)}${r+1}: ${cell.slice(0,120)}`);
    });
  });
  console.log(`  Total formula cells: ${mcTotal}`);
  console.log(`  Functions: ${JSON.stringify(mcFns)}`);
  console.log('  Samples:');
  mcSamples.forEach(s => console.log('    ' + s));

  console.log('\n\n✅  Deep audit complete.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
