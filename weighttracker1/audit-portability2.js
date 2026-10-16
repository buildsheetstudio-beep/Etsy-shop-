'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const SPREADSHEET_ID = '1CCo2yKNdKwNiuzK0E4CuNVzVCP27M72h0d4h81IMEaI';
async function getAuth() {
  const s = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const c = s.installed || s.web;
  const o = new google.auth.OAuth2(c.client_id, c.client_secret, c.redirect_uris[0]);
  o.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return o;
}
(async () => {
  const sheets = google.sheets({ version: 'v4', auth: await getAuth() });

  // 1. Named ranges with their actual references
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'namedRanges,sheets.properties',
  });
  const sheetMap = {};
  meta.data.sheets.forEach(s => { sheetMap[s.properties.sheetId] = s.properties.title; });

  console.log('=== NAMED RANGES (full detail) ===');
  (meta.data.namedRanges || []).forEach(nr => {
    const r = nr.range;
    const sheet = sheetMap[r.sheetId] || `sheetId=${r.sheetId}`;
    console.log(`  "${nr.name}": ${sheet} R${r.startRowIndex+1}:C${r.startColumnIndex+1} → R${r.endRowIndex}:C${r.endColumnIndex}`);
  });

  // 2. Sample formulas from key tabs (first 5 unique formula patterns per tab)
  const TABS = ['Event Dashboard', 'Monthly Calendar', 'Master Event Log', 'Event Details', 'Event Tasks'];
  for (const tab of TABS) {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${tab}'!A1:Z200`,
      valueRenderOption: 'FORMULA',
    });
    const seen = new Set();
    const samples = [];
    (res.data.values || []).forEach((row, ri) => {
      row.forEach((cell, ci) => {
        if (typeof cell === 'string' && cell.startsWith('=')) {
          const key = cell.slice(0, 50);
          if (!seen.has(key)) {
            seen.add(key);
            const col = String.fromCharCode(65 + ci);
            samples.push(`    ${col}${ri+1}: ${cell.slice(0, 120)}`);
          }
        }
      });
    });
    console.log(`\n=== "${tab}" — unique formula patterns (first 12) ===`);
    samples.slice(0, 12).forEach(s => console.log(s));
  }

  // 3. CF rules on ALL sheets (not just Weekly)
  const cfMeta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets(properties(title),conditionalFormats)',
  });
  console.log('\n=== CONDITIONAL FORMAT RULES (all sheets) ===');
  cfMeta.data.sheets.forEach(sheet => {
    const cfs = sheet.conditionalFormats || [];
    if (cfs.length > 0) {
      console.log(`\n  "${sheet.properties.title}": ${cfs.length} CF rule(s)`);
      cfs.slice(0, 3).forEach((cf, i) => {
        const rule = cf.booleanRule;
        if (rule && rule.condition) {
          const val = rule.condition.values && rule.condition.values[0];
          console.log(`    CF[${i}]: type=${rule.condition.type} formula=${val ? val.userEnteredValue.slice(0, 80) : 'n/a'}`);
        }
      });
      if (cfs.length > 3) console.log(`    ... and ${cfs.length - 3} more`);
    }
  });

  // 4. Data validation rules (do any reference named ranges or INDIRECT?)
  const dvMeta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets(properties(title),data(rowData(values(dataValidation))))',
  });
  console.log('\n=== DATA VALIDATION PATTERNS ===');
  dvMeta.data.sheets.forEach(sheet => {
    const dvTypes = new Set();
    const dvSources = [];
    (sheet.data || []).forEach(grid => {
      (grid.rowData || []).forEach(row => {
        (row.values || []).forEach(cell => {
          const dv = cell.dataValidation;
          if (!dv) return;
          dvTypes.add(dv.condition.type);
          if (dv.condition.values) {
            dv.condition.values.forEach(v => {
              if (v.userEnteredValue && !dvSources.includes(v.userEnteredValue)) {
                dvSources.push(v.userEnteredValue);
              }
            });
          }
        });
      });
    });
    if (dvTypes.size > 0) {
      console.log(`  "${sheet.properties.title}": types=[${[...dvTypes].join(',')}]`);
      dvSources.slice(0, 5).forEach(s => console.log(`    source: ${s.slice(0, 80)}`));
    }
  });
})().catch(e => { console.error(e.message || e); process.exit(1); });
