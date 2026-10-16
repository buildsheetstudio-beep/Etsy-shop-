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

// ── Google-only / version-gated functions to flag ─────────────────────────────
const GS_ONLY   = ['ARRAYFORMULA','GOOGLEFINANCE','IMPORTRANGE','QUERY','REGEXMATCH','REGEXEXTRACT','REGEXREPLACE','SPARKLINE','TO_DATE','SPLIT','JOIN','FLATTEN','UNIQUE','SORTN','IFERROR'];
const EXCEL_NEW = ['LET','LAMBDA','XLOOKUP','XMATCH','FILTER','SORT','UNIQUE','SEQUENCE','BYROW','BYCOL','REDUCE','SCAN','MAKEARRAY','TEXTJOIN','IFS','MAXIFS','MINIFS','SWITCH'];
const ERROR_VALS = ['#VALUE!','#REF!','#ERROR!','#NAME?','#N/A','#DIV/0!','#NULL!','#NUM!'];

function scanFormula(formula) {
  const up = formula.toUpperCase();
  const gsHits   = GS_ONLY.filter(fn => up.includes(fn + '('));
  const newHits  = EXCEL_NEW.filter(fn => up.includes(fn + '('));
  return { gsHits, newHits };
}

(async () => {
  const sheets = await getSheets();

  // ── 1. Spreadsheet metadata ──────────────────────────────────────────────
  console.log('═══════════════════════════════════════════════════════════');
  console.log('AUDIT — Spreadsheet:', SPREADSHEET_ID);
  console.log('═══════════════════════════════════════════════════════════\n');

  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'properties,sheets(properties,protectedRanges),namedRanges',
  });

  const sheetList = meta.data.sheets.map(s => s.properties);
  const namedRanges = meta.data.namedRanges || [];

  console.log('── 1. SHEET TABS ────────────────────────────────────────');
  sheetList.forEach(p => {
    const hidden = p.hidden ? ' [HIDDEN]' : '';
    console.log(`  sheetId=${String(p.sheetId).padStart(10)}  index=${p.index}  "${p.title}"${hidden}  (${p.gridProperties?.rowCount}r × ${p.gridProperties?.columnCount}c)`);
  });
  console.log(`  Total: ${sheetList.length} tabs\n`);

  console.log('── 2. NAMED RANGES ─────────────────────────────────────');
  if (namedRanges.length === 0) {
    console.log('  None defined');
  } else {
    namedRanges.forEach(nr => {
      const r = nr.range;
      const sid = r.sheetId ?? 0;
      const sheet = sheetList.find(s => s.sheetId === sid);
      console.log(`  "${nr.name}"  →  sheetId=${sid} ("${sheet?.title || '?'}")  rows ${r.startRowIndex}:${r.endRowIndex}  cols ${r.startColumnIndex}:${r.endColumnIndex}`);
    });
  }
  console.log();

  // ── 2. Per-tab formula + error scan ─────────────────────────────────────
  console.log('── 3. FORMULA / ERROR SCAN (per tab) ───────────────────');

  const gsFunctionCount = {};   // fnName → total occurrences across all cells
  const newFunctionCount = {};
  const errorCells = [];        // { tab, cell, value }
  const formulaCells = [];      // total formula cells per tab

  for (const sp of sheetList) {
    const tabName = sp.title;
    const maxRow = Math.min(sp.gridProperties?.rowCount || 1000, 600);
    const maxCol = Math.min(sp.gridProperties?.columnCount || 26, 40);
    const lastColLetter = colIndexToLetter(maxCol - 1);

    let tabFormulas = 0;
    let tabErrors = 0;
    const tabGsFns = {};
    const tabNewFns = {};

    // Read formulas
    let formData;
    try {
      formData = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${tabName}'!A1:${lastColLetter}${maxRow}`,
        valueRenderOption: 'FORMULA',
      });
    } catch { continue; }

    const formRows = formData.data.values || [];

    // Read displayed values (for error detection)
    let dispData;
    try {
      dispData = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${tabName}'!A1:${lastColLetter}${maxRow}`,
        valueRenderOption: 'FORMATTED_VALUE',
      });
    } catch { continue; }

    const dispRows = dispData.data.values || [];

    formRows.forEach((row, r) => {
      row.forEach((cell, c) => {
        if (typeof cell === 'string' && cell.startsWith('=')) {
          tabFormulas++;
          const { gsHits, newHits } = scanFormula(cell);
          gsHits.forEach(fn => { tabGsFns[fn] = (tabGsFns[fn]||0)+1; gsFunctionCount[fn] = (gsFunctionCount[fn]||0)+1; });
          newHits.forEach(fn => { tabNewFns[fn] = (tabNewFns[fn]||0)+1; newFunctionCount[fn] = (newFunctionCount[fn]||0)+1; });
        }
      });
    });

    // Error scan on displayed values
    dispRows.forEach((row, r) => {
      row.forEach((cell, c) => {
        if (typeof cell === 'string' && ERROR_VALS.some(e => cell.startsWith(e))) {
          tabErrors++;
          errorCells.push({ tab: tabName, cell: `${colIndexToLetter(c)}${r+1}`, value: cell });
        }
      });
    });

    const gsStr  = Object.entries(tabGsFns).map(([k,v]) => `${k}(${v})`).join(', ') || 'none';
    const newStr = Object.entries(tabNewFns).map(([k,v]) => `${k}(${v})`).join(', ') || 'none';
    const errStr = tabErrors > 0 ? ` ⚠ ${tabErrors} ERRORS` : ' ✓ no errors';
    console.log(`  "${tabName}": ${tabFormulas} formulas${errStr}`);
    if (Object.keys(tabGsFns).length) console.log(`    GS-only: ${gsStr}`);
    if (Object.keys(tabNewFns).length) console.log(`    Version-gated: ${newStr}`);
    formulaCells.push({ tab: tabName, count: tabFormulas });
  }

  console.log('\n── 4. ERROR CELLS DETAIL ───────────────────────────────');
  if (errorCells.length === 0) {
    console.log('  ✓ Zero formula errors across all tabs');
  } else {
    errorCells.slice(0, 40).forEach(e => console.log(`  ⚠ ${e.tab}!${e.cell}: ${e.value}`));
    if (errorCells.length > 40) console.log(`  ... and ${errorCells.length - 40} more`);
  }
  console.log();

  // ── 3. Data validation ───────────────────────────────────────────────────
  console.log('── 5. DATA VALIDATION (sample A6 on each tab) ──────────');
  for (const sp of sheetList) {
    let dvData;
    try {
      dvData = await sheets.spreadsheets.get({
        spreadsheetId: SPREADSHEET_ID,
        ranges: [`'${sp.title}'!A1:Z20`],
        includeGridData: true,
        fields: 'sheets(data(rowData(values(dataValidation))))',
      });
    } catch { continue; }

    const dvCells = [];
    (dvData.data.sheets[0]?.data[0]?.rowData || []).forEach((row, r) => {
      (row.values || []).forEach((cell, c) => {
        if (cell.dataValidation) {
          dvCells.push({ cell: `${colIndexToLetter(c)}${r+1}`, dv: cell.dataValidation });
        }
      });
    });

    if (dvCells.length) {
      console.log(`  "${sp.title}":`);
      dvCells.slice(0, 6).forEach(({ cell, dv }) => {
        const type = dv.condition?.type || '?';
        const vals = (dv.condition?.values || []).map(v => v.userEnteredValue).join(' | ').slice(0, 60);
        console.log(`    ${cell}: ${type}  ${vals}`);
      });
      if (dvCells.length > 6) console.log(`    ... +${dvCells.length - 6} more validation cells`);
    }
  }
  console.log();

  // ── 4. Conditional formatting summary ────────────────────────────────────
  console.log('── 6. CONDITIONAL FORMATTING ───────────────────────────');
  const cfMeta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets(properties(sheetId,title),conditionalFormats)',
  });
  cfMeta.data.sheets.forEach(sh => {
    const cf = sh.conditionalFormats || [];
    if (cf.length === 0) return;
    console.log(`  "${sh.properties.title}": ${cf.length} rules`);
    cf.slice(0, 4).forEach((rule, i) => {
      if (rule.booleanRule) {
        const cond = rule.booleanRule.condition;
        const formula = (cond.values?.[0]?.userEnteredValue || '').slice(0, 80);
        const hasIndirect = formula.toUpperCase().includes('INDIRECT');
        const crossSheet = formula.includes("'") || formula.includes('!');
        console.log(`    [${i}] ${cond.type}: ${formula} ${hasIndirect ? '⚠INDIRECT' : ''} ${crossSheet ? '⚠cross-sheet' : ''}`);
      } else if (rule.gradientRule) {
        console.log(`    [${i}] gradientRule`);
      }
    });
    if (cf.length > 4) console.log(`    ... +${cf.length - 4} more rules`);
  });
  console.log();

  // ── 5. Summary ────────────────────────────────────────────────────────────
  console.log('── 7. PORTABILITY SUMMARY ──────────────────────────────');
  console.log('  Google Sheets only functions detected:');
  Object.entries(gsFunctionCount).forEach(([fn,cnt]) => console.log(`    ${fn}: ${cnt} cells`));
  if (Object.keys(gsFunctionCount).length === 0) console.log('    None');
  console.log('  Version-gated functions (Excel 365 / 2019+):');
  Object.entries(newFunctionCount).forEach(([fn,cnt]) => console.log(`    ${fn}: ${cnt} cells`));
  if (Object.keys(newFunctionCount).length === 0) console.log('    None');
  console.log(`\n  Total error cells: ${errorCells.length}`);
  console.log(`  Total named ranges: ${namedRanges.length}`);

})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });

function colIndexToLetter(idx) {
  let result = '';
  let i = idx + 1;
  while (i > 0) {
    const rem = (i - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    i = Math.floor((i - 1) / 26);
  }
  return result;
}
