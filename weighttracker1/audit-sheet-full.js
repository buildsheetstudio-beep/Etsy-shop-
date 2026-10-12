'use strict';
/**
 * Full spreadsheet audit — formula errors, portability, hardcoded values,
 * data validation, conditional formatting, named ranges, cross-sheet refs.
 *
 * Usage: SPREADSHEET_ID=<id> node audit-sheet-full.js
 *   or set SPREADSHEET_ID at top of file.
 */
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = process.env.SPREADSHEET_ID || '1tmklWErTpmbkoP3mVAYds6gbN1BYauwPvRTWfOmWwZI';

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}
async function getSheets() { return google.sheets({ version: 'v4', auth: await getAuth() }); }

function colLetter(idx) {
  let r = '', i = idx + 1;
  while (i > 0) { const rem = (i-1)%26; r = String.fromCharCode(65+rem)+r; i = Math.floor((i-1)/26); }
  return r;
}

// Google Sheets-only functions (not in Excel)
const GS_ONLY   = ['ARRAYFORMULA','GOOGLEFINANCE','IMPORTRANGE','QUERY','REGEXMATCH',
                   'REGEXEXTRACT','REGEXREPLACE','SPARKLINE','SPLIT','FLATTEN','SORTN'];
// Functions requiring Excel 365 / Excel 2019+
const EXCEL_NEW = ['LET','LAMBDA','XLOOKUP','XMATCH','FILTER','SORT','UNIQUE','SEQUENCE',
                   'BYROW','BYCOL','REDUCE','SCAN','MAKEARRAY','TEXTJOIN','IFS','MAXIFS','MINIFS','SWITCH'];
const ERROR_VALS = ['#VALUE!','#REF!','#ERROR!','#NAME?','#N/A','#DIV/0!','#NULL!','#NUM!'];
const YEAR_RE    = /[^A-Z](202[0-9]|20[3-9][0-9])/;  // hardcoded year literals in formulas

(async () => {
  const sheets = await getSheets();

  // ── 1. Metadata ────────────────────────────────────────────────────────────
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'properties,sheets(properties,protectedRanges),namedRanges',
  });
  const title    = meta.data.properties.title;
  const tabList  = meta.data.sheets.map(s => s.properties);
  const named    = meta.data.namedRanges || [];

  console.log('═'.repeat(70));
  console.log('FULL AUDIT:', title);
  console.log('Spreadsheet ID:', SPREADSHEET_ID);
  console.log('═'.repeat(70));

  console.log('\n── 1. TABS ──────────────────────────────────────────────────────────');
  tabList.forEach(p => {
    const tag = p.hidden ? ' [HIDDEN]' : '';
    console.log(`  ${String(p.sheetId).padStart(9)}  idx=${p.index}  "${p.title}"${tag}  (${p.gridProperties?.rowCount}r × ${p.gridProperties?.columnCount}c)`);
  });
  console.log(`  Total: ${tabList.length} tabs`);

  console.log('\n── 2. NAMED RANGES ─────────────────────────────────────────────────');
  if (named.length === 0) {
    console.log('  None');
  } else {
    named.forEach(nr => {
      const r = nr.range;
      const sh = tabList.find(s => s.sheetId === (r.sheetId ?? 0));
      console.log(`  "${nr.name}" → "${sh?.title||'?'}" rows ${r.startRowIndex}:${r.endRowIndex} cols ${r.startColumnIndex}:${r.endColumnIndex}`);
    });
  }

  // ── 2. Per-tab formula + error scan ────────────────────────────────────────
  console.log('\n── 3. FORMULA / FUNCTION / ERROR SCAN ──────────────────────────────');

  const allErrors      = [];
  const allHardYear    = [];
  const allGsFns       = {};
  const allNewFns      = {};
  const allArrayForm   = [];

  for (const sp of tabList) {
    const tab    = sp.title;
    const maxRow = Math.min(sp.gridProperties?.rowCount || 1000, 700);
    const maxCol = Math.min(sp.gridProperties?.columnCount || 26, 50);
    const lastCol = colLetter(maxCol - 1);

    let formData, dispData;
    try {
      [formData, dispData] = await Promise.all([
        sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `'${tab}'!A1:${lastCol}${maxRow}`, valueRenderOption: 'FORMULA' }),
        sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `'${tab}'!A1:${lastCol}${maxRow}`, valueRenderOption: 'FORMATTED_VALUE' }),
      ]);
    } catch { continue; }

    const formRows = formData.data.values || [];
    const dispRows = dispData.data.values || [];

    let tabFormulas = 0, tabErrors = 0;
    const tabGsFns = {}, tabNewFns = {};
    let tabArrayFormula = 0;

    formRows.forEach((row, r) => {
      row.forEach((cell, c) => {
        if (typeof cell !== 'string' || !cell.startsWith('=')) return;
        tabFormulas++;
        const up = cell.toUpperCase();

        // GS-only
        GS_ONLY.forEach(fn => { if (up.includes(fn+'(')) { tabGsFns[fn]=(tabGsFns[fn]||0)+1; allGsFns[fn]=(allGsFns[fn]||0)+1; }});
        // Excel 365/2019+
        EXCEL_NEW.forEach(fn => { if (up.includes(fn+'(')) { tabNewFns[fn]=(tabNewFns[fn]||0)+1; allNewFns[fn]=(allNewFns[fn]||0)+1; }});
        // ARRAYFORMULA (GS only, breaks in Excel)
        if (up.includes('ARRAYFORMULA(')) { tabArrayFormula++; allArrayFormula.push(`${tab}!${colLetter(c)}${r+1}`); }
        // Hardcoded years in formulas
        if (YEAR_RE.test(cell)) allHardYear.push({ tab, cell: `${colLetter(c)}${r+1}`, formula: cell.slice(0,100) });
      });
    });

    dispRows.forEach((row, r) => {
      row.forEach((cell, c) => {
        if (typeof cell === 'string' && ERROR_VALS.some(e => cell.startsWith(e))) {
          tabErrors++;
          allErrors.push({ tab, cell: `${colLetter(c)}${r+1}`, value: cell });
        }
      });
    });

    const gStr  = Object.entries(tabGsFns).map(([k,v])=>`${k}(${v})`).join(', ') || '';
    const nStr  = Object.entries(tabNewFns).map(([k,v])=>`${k}(${v})`).join(', ') || '';
    const eTag  = tabErrors > 0 ? `  ⚠ ${tabErrors} ERRORS` : '  ✓';
    const afTag = tabArrayFormula > 0 ? `  ⚠ ARRAYFORMULA(${tabArrayFormula})` : '';
    console.log(`  "${tab}": ${tabFormulas} formulas${eTag}${afTag}`);
    if (gStr)  console.log(`    GS-only: ${gStr}`);
    if (nStr)  console.log(`    Excel 365/2019+: ${nStr}`);
  }

  // ── 3. Error details ───────────────────────────────────────────────────────
  console.log('\n── 4. ERROR CELLS ──────────────────────────────────────────────────');
  if (allErrors.length === 0) {
    console.log('  ✓ Zero formula errors across all tabs');
  } else {
    allErrors.slice(0, 50).forEach(e => console.log(`  ⚠ ${e.tab}!${e.cell}: ${e.value}`));
    if (allErrors.length > 50) console.log(`  ... +${allErrors.length-50} more`);
  }

  // ── 4. Hardcoded years ─────────────────────────────────────────────────────
  console.log('\n── 5. HARDCODED YEAR LITERALS IN FORMULAS ───────────────────────────');
  if (allHardYear.length === 0) {
    console.log('  ✓ None found');
  } else {
    allHardYear.slice(0, 20).forEach(({ tab, cell, formula }) =>
      console.log(`  ⚠ ${tab}!${cell}: ${formula}`));
    if (allHardYear.length > 20) console.log(`  ... +${allHardYear.length-20} more`);
  }

  // ── 5. ARRAYFORMULA ────────────────────────────────────────────────────────
  console.log('\n── 6. ARRAYFORMULA USAGE (GS-only, breaks in Excel) ─────────────────');
  if (allArrayFormula.length === 0) {
    console.log('  ✓ None found');
  } else {
    allArrayFormula.slice(0, 20).forEach(c => console.log(`  ⚠ ${c}`));
    if (allArrayFormula.length > 20) console.log(`  ... +${allArrayFormula.length-20} more`);
  }

  // ── 6. Portability summary ─────────────────────────────────────────────────
  console.log('\n── 7. PORTABILITY SUMMARY ───────────────────────────────────────────');
  console.log('  Google Sheets-only functions:');
  if (Object.keys(allGsFns).length === 0) console.log('    ✓ None');
  Object.entries(allGsFns).forEach(([fn,cnt]) => console.log(`    ⚠ ${fn}: ${cnt} cells`));
  console.log('  Excel 365 / 2019+ only functions:');
  if (Object.keys(allNewFns).length === 0) console.log('    ✓ None');
  Object.entries(allNewFns).forEach(([fn,cnt]) => console.log(`    ⚠ ${fn}: ${cnt} cells — requires Excel 365 or 2019+`));

  // ── 7. Data validation ────────────────────────────────────────────────────
  console.log('\n── 8. DATA VALIDATION ───────────────────────────────────────────────');
  const cfMeta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets(properties(sheetId,title),conditionalFormats)',
  });
  for (const sp of tabList) {
    let dvData;
    try {
      dvData = await sheets.spreadsheets.get({
        spreadsheetId: SPREADSHEET_ID,
        ranges: [`'${sp.title}'!A1:Z30`],
        includeGridData: true,
        fields: 'sheets(data(rowData(values(dataValidation))))',
      });
    } catch { continue; }
    const dvMap = {};
    (dvData.data.sheets[0]?.data[0]?.rowData || []).forEach((row, r) => {
      (row.values || []).forEach((cell, c) => {
        if (!cell.dataValidation) return;
        const type = cell.dataValidation.condition?.type || '?';
        const vals = (cell.dataValidation.condition?.values||[]).map(v=>v.userEnteredValue).join('|').slice(0,60);
        const key = `${type}::${vals}`;
        if (!dvMap[key]) dvMap[key] = { type, vals, count: 0, sample: `${colLetter(c)}${r+1}` };
        dvMap[key].count++;
      });
    });
    if (Object.keys(dvMap).length) {
      console.log(`  "${sp.title}":`);
      Object.values(dvMap).forEach(({ type, vals, count, sample }) =>
        console.log(`    ${type} (${count} cells, e.g. ${sample}): ${vals.slice(0,60)}`));
    }
  }

  // ── 8. Conditional formatting ──────────────────────────────────────────────
  console.log('\n── 9. CONDITIONAL FORMATTING ────────────────────────────────────────');
  cfMeta.data.sheets.forEach(sh => {
    const cf = sh.conditionalFormats || [];
    if (!cf.length) return;
    let hasIndirect = false, hasCrossSheet = false;
    cf.forEach(rule => {
      const f = rule.booleanRule?.condition?.values?.[0]?.userEnteredValue || '';
      if (f.toUpperCase().includes('INDIRECT')) hasIndirect = true;
      if (f.includes("'") || f.includes('!')) hasCrossSheet = true;
    });
    const tags = [hasIndirect ? '⚠INDIRECT (GS-only in CF)' : '', hasCrossSheet ? '⚠cross-sheet (breaks in Excel)' : ''].filter(Boolean).join(' ');
    console.log(`  "${sh.properties.title}": ${cf.length} rules${tags ? '  '+tags : '  ✓'}`);
  });

  // ── 9. Deep formula samples per tab ──────────────────────────────────────
  console.log('\n── 10. FORMULA SAMPLES (first distinct pattern per tab) ─────────────');
  for (const sp of tabList) {
    const tab = sp.title;
    const maxRow = Math.min(sp.gridProperties?.rowCount || 1000, 50);
    const maxCol = Math.min(sp.gridProperties?.columnCount || 26, 20);
    let fData;
    try {
      fData = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${tab}'!A1:${colLetter(maxCol-1)}${maxRow}`,
        valueRenderOption: 'FORMULA',
      });
    } catch { continue; }
    const seen = new Set();
    const samples = [];
    (fData.data.values || []).forEach(row => {
      row.forEach(cell => {
        if (typeof cell !== 'string' || !cell.startsWith('=')) return;
        const key = cell.replace(/\$?[A-Z]{1,3}\$?\d+/g,'R').replace(/\d+/g,'N').slice(0,60);
        if (!seen.has(key)) { seen.add(key); samples.push(cell.slice(0,110)); }
        if (samples.length >= 5) return;
      });
    });
    if (samples.length) {
      console.log(`\n  "${tab}":`);
      samples.forEach(s => console.log(`    ${s}`));
    }
  }

  // ── 10. Check for TEXTJOIN(IF()) anti-pattern ─────────────────────────────
  console.log('\n── 11. TEXTJOIN(IF()) ANTI-PATTERN CHECK ────────────────────────────');
  let textjoinIf = 0;
  for (const sp of tabList) {
    const tab = sp.title;
    const maxRow = Math.min(sp.gridProperties?.rowCount || 1000, 700);
    const maxCol = Math.min(sp.gridProperties?.columnCount || 26, 50);
    let fData;
    try {
      fData = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${tab}'!A1:${colLetter(maxCol-1)}${maxRow}`,
        valueRenderOption: 'FORMULA',
      });
    } catch { continue; }
    (fData.data.values || []).forEach(row => {
      row.forEach(cell => {
        if (typeof cell === 'string' && cell.toUpperCase().includes('TEXTJOIN') && cell.includes(',IF(')) textjoinIf++;
      });
    });
  }
  if (textjoinIf > 0) {
    console.log(`  ⚠ ${textjoinIf} cells use TEXTJOIN(IF()) — needs ARRAYFORMULA in GS or FILTER in GS+Excel`);
  } else {
    console.log('  ✓ No TEXTJOIN(IF()) patterns found');
  }

  // ── Final summary ─────────────────────────────────────────────────────────
  console.log('\n' + '═'.repeat(70));
  console.log('SUMMARY');
  console.log('═'.repeat(70));
  console.log(`  Total tabs:           ${tabList.length}`);
  console.log(`  Formula errors:       ${allErrors.length} ${allErrors.length ? '⚠ NEEDS FIX' : '✓'}`);
  console.log(`  Hardcoded years:      ${allHardYear.length} ${allHardYear.length ? '⚠ will go stale' : '✓'}`);
  console.log(`  ARRAYFORMULA cells:   ${allArrayFormula.length} ${allArrayFormula.length ? '⚠ GS-only' : '✓'}`);
  console.log(`  TEXTJOIN(IF()) cells: ${textjoinIf} ${textjoinIf ? '⚠ broken w/o ARRAYFORMULA' : '✓'}`);
  console.log(`  GS-only functions:    ${Object.keys(allGsFns).join(', ')||'none'}`);
  console.log(`  Excel 365+ functions: ${Object.keys(allNewFns).join(', ')||'none'}`);
  console.log(`  Named ranges:         ${named.length}`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
