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

  // ── 1. Named ranges: sheetId=undefined actually means sheetId=0 in protobuf ──
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'namedRanges,sheets.properties',
  });
  const sheetById = {};
  meta.data.sheets.forEach(s => { sheetById[s.properties.sheetId] = s.properties.title; });
  console.log('=== NAMED RANGES — resolved sheet anchor ===');
  (meta.data.namedRanges || []).forEach(nr => {
    const sid = nr.range.sheetId; // undefined = 0 (protobuf default omission)
    const resolvedId = sid === undefined ? 0 : sid;
    const sheetTitle = sheetById[resolvedId] || `unknown(${resolvedId})`;
    const startCol = String.fromCharCode(65 + nr.range.startColumnIndex);
    const endCol   = String.fromCharCode(65 + nr.range.endColumnIndex - 1);
    console.log(`  "${nr.name}" → "${sheetTitle}" ${startCol}${nr.range.startRowIndex+1}:${endCol}${nr.range.endRowIndex} ✓`);
  });

  // ── 2. Data validation: verify event ID columns use live range ──────────────
  const dvMeta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ["'Guests & Attendees'!A6:A8", "'Event Tasks'!A6:A8"],
    includeGridData: true,
    fields: 'sheets(properties(title),data(rowData(values(dataValidation))))',
  });
  console.log('\n=== EVENT ID DROPDOWN VALIDATION TYPE ===');
  dvMeta.data.sheets.forEach(sh => {
    const cell = sh.data[0].rowData[0].values[0];
    const dv = cell && cell.dataValidation;
    if (dv) {
      const src = (dv.condition.values || [])[0]?.userEnteredValue || '';
      console.log(`  "${sh.properties.title}" col A: type=${dv.condition.type} source="${src}"`);
    }
  });

  // ── 3. Check all other DV types across every tab ────────────────────────────
  const allDvMeta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets(properties(title),data(rowData(values(dataValidation))))',
  });
  console.log('\n=== ALL DROPDOWN TYPES ACROSS EVERY TAB ===');
  allDvMeta.data.sheets.forEach(sheet => {
    const dvSummary = {};
    (sheet.data || []).forEach(grid => {
      (grid.rowData || []).forEach(row => {
        (row.values || []).forEach(cell => {
          const dv = cell.dataValidation;
          if (!dv) return;
          const t = dv.condition.type;
          dvSummary[t] = (dvSummary[t] || 0) + 1;
        });
      });
    });
    if (Object.keys(dvSummary).length > 0) {
      const parts = Object.entries(dvSummary).map(([t, n]) => `${t}×${n}`).join(', ');
      // Flag any remaining hardcoded EVT- lists
      let hasHardcodedEvt = false;
      (sheet.data || []).forEach(grid => {
        (grid.rowData || []).forEach(row => {
          (row.values || []).forEach(cell => {
            const dv = cell.dataValidation;
            if (!dv || dv.condition.type !== 'ONE_OF_LIST') return;
            const vals = (dv.condition.values || []).map(v => v.userEnteredValue || '');
            if (vals.some(v => /^EVT-\d/.test(v))) hasHardcodedEvt = true;
          });
        });
      });
      const flag = hasHardcodedEvt ? ' ⚠️ HARDCODED EVT IDs REMAIN' : '';
      console.log(`  "${sheet.properties.title}": ${parts}${flag}`);
    }
  });

  // ── 4. Scan every tab for formula errors (#VALUE, #REF, #NAME, #ERROR) ───────
  console.log('\n=== FORMULA ERROR SCAN ===');
  const allTabs = meta.data.sheets.map(s => s.properties.title);
  let totalErrors = 0;
  for (const tab of allTabs) {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${tab}'!A1:Z500`,
      valueRenderOption: 'FORMATTED_VALUE',
    });
    const errors = [];
    (res.data.values || []).forEach((row, ri) => {
      row.forEach((cell, ci) => {
        if (typeof cell === 'string' && /^#(VALUE|REF|NAME|ERROR|DIV\/0|N\/A|NULL)/.test(cell)) {
          errors.push(`${String.fromCharCode(65+ci)}${ri+1}=${cell}`);
        }
      });
    });
    if (errors.length > 0) {
      console.log(`  ❌ "${tab}": ${errors.slice(0, 10).join(', ')}${errors.length > 10 ? ` +${errors.length-10} more` : ''}`);
      totalErrors += errors.length;
    } else {
      console.log(`  ✓  "${tab}": clean`);
    }
  }
  console.log(`\n  Total formula errors: ${totalErrors}`);

  // ── 5. WEEKNUM compatibility check ──────────────────────────────────────────
  console.log('\n=== WEEKNUM CELLS (check for GS vs Excel default mode) ===');
  const wnCells = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ["'Weekly Schedule'!H4", "'Weekly Schedule'!A34"],
    valueRenderOption: 'FORMULA',
  });
  wnCells.data.valueRanges.forEach(vr => {
    const f = (vr.values || [['']])[0][0];
    if (f) {
      const hasMode = /WEEKNUM\s*\([^,)]+,\s*\d/.test(f);
      console.log(`  ${vr.range}: ${hasMode ? 'has explicit mode ✓' : 'NO explicit mode — defaults differ between GS/Excel ⚠️'}`);
      console.log(`    ${f.slice(0, 100)}`);
    }
  });

  // ── 6. Cross-sheet reference style check ────────────────────────────────────
  console.log('\n=== CROSS-SHEET REF STYLE (GS vs Excel syntax) ===');
  // GS/Excel both use 'Sheet Name'!A1 format — should be identical
  const sampleCrossRef = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Event Dashboard'!A4",
    valueRenderOption: 'FORMULA',
  });
  const formula = (sampleCrossRef.data.values || [['']])[0][0];
  console.log(`  Sample cross-sheet formula: ${formula}`);
  const usesEmoji = /[\u{1F300}-\u{1FFFF}]/u.test(formula);
  console.log(`  Sheet names contain emoji: ${usesEmoji ? 'YES ⚠️ (may break in Excel)' : 'No ✓'}`);

  // ── 7. Conditional format summary (GS copy safety) ──────────────────────────
  const cfMeta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets(properties(title),conditionalFormats)',
  });
  console.log('\n=== CONDITIONAL FORMAT SUMMARY ===');
  cfMeta.data.sheets.forEach(sh => {
    const cfs = sh.conditionalFormats || [];
    if (cfs.length === 0) return;
    const custom = cfs.filter(cf => cf.booleanRule?.condition?.type === 'CUSTOM_FORMULA').length;
    const simple  = cfs.length - custom;
    const hasIndirect = cfs.some(cf => {
      const v = cf.booleanRule?.condition?.values?.[0]?.userEnteredValue || '';
      return v.includes('INDIRECT');
    });
    const gsOnly = hasIndirect ? '(INDIRECT in CF — GS-only cross-sheet trick)' : '';
    console.log(`  "${sh.properties.title}": ${cfs.length} rules (${custom} custom formula, ${simple} simple) ${gsOnly}`);
  });

  // ── 8. Banner + note verification ───────────────────────────────────────────
  console.log('\n=== PORTABILITY NOTES IN SPREADSHEET ===');
  const bannerCheck = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ["'Weekly Schedule'!A3", "'Monthly Calendar'!E3", "'Event Dashboard'!A2"],
    valueRenderOption: 'FORMATTED_VALUE',
  });
  bannerCheck.data.valueRanges.forEach(vr => {
    const val = (vr.values || [['']])[0][0] || '';
    console.log(`  ${vr.range}: "${val.slice(0, 65)}..."`);
  });
  // Check cell note on Event Dashboard A2
  const noteCheck = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ["'Event Dashboard'!A2"],
    includeGridData: true,
    fields: 'sheets(data(rowData(values(note))))',
  });
  const note = noteCheck.data.sheets[0]?.data[0]?.rowData[0]?.values[0]?.note || '';
  console.log(`  Event Dashboard A2 cell note: "${note.slice(0, 70)}..."`);

})().catch(e => { console.error(e.message || e); process.exit(1); });
