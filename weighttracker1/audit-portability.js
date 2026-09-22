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

// Functions that exist in Google Sheets but NOT in Excel
const GOOGLE_ONLY = [
  'QUERY','IMPORTRANGE','IMPORTDATA','IMPORTHTML','IMPORTFEED','IMPORTXML',
  'SPARKLINE','IMAGE','GOOGLETRANSLATE','DETECTLANGUAGE',
  'ARRAYFORMULA','FLATTEN','MAKEARRAY',
  'REGEXMATCH','REGEXEXTRACT','REGEXREPLACE',
  'TO_DATE','TO_TEXT','TO_PERCENT','TO_DOLLARS','TO_PURE_NUMBER',
  'EPOCH','EPOCHTODATE','DATEVALUE',
  'SPLIT','JOIN',  // JOIN is GS-specific; SPLIT exists in Excel 365
  'ISBETWEEN',
];

// Functions in Excel but with different behavior or version requirements
const EXCEL_COMPAT_WARN = [
  'TEXTJOIN',  // Excel 2019/365 only
  'IFS',       // Excel 2019/365 only  
  'SWITCH',    // Excel 2019/365 only
  'MAXIFS',    // Excel 2019/365 only
  'MINIFS',    // Excel 2019/365 only
  'UNIQUE',    // Excel 365 only
  'SORT','SORTBY','FILTER','SEQUENCE','RANDARRAY', // Excel 365 dynamic arrays only
  'XLOOKUP',   // Excel 365 only
  'LET',       // Excel 365 only
  'WEEKNUM',   // exists in Excel but GS uses ISO 8601 by default (mode 2 vs Excel's mode 1)
  'ISOWEEKNUM', // Excel 2013+ only
];

function scanFormula(formula) {
  const issues = [];
  const upper = formula.toUpperCase();

  for (const fn of GOOGLE_ONLY) {
    const re = new RegExp(`\\b${fn}\\s*\\(`, 'i');
    if (re.test(formula)) issues.push({ type: 'google_only', fn });
  }

  for (const fn of EXCEL_COMPAT_WARN) {
    const re = new RegExp(`\\b${fn}\\s*\\(`, 'i');
    if (re.test(formula)) issues.push({ type: 'excel_warn', fn });
  }

  // INDIRECT with hardcoded sheet names — works in copies IF sheet names unchanged
  if (/INDIRECT\s*\(/i.test(formula)) {
    issues.push({ type: 'indirect', formula: formula.slice(0, 80) });
  }

  // ARRAYFORMULA wrapping other formulas
  if (/ARRAYFORMULA\s*\(/i.test(formula)) {
    issues.push({ type: 'google_only', fn: 'ARRAYFORMULA' });
  }

  return issues;
}

(async () => {
  const sheets = google.sheets({ version: 'v4', auth: await getAuth() });

  // 1. Get all sheet names
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets.properties,namedRanges',
  });
  const sheetList = meta.data.sheets.map(s => ({
    id: s.properties.sheetId,
    title: s.properties.title,
    hidden: s.properties.hidden || false,
    rowCount: s.properties.gridProperties.rowCount,
    colCount: s.properties.gridProperties.columnCount,
  }));
  console.log('=== SHEETS ===');
  sheetList.forEach(s => console.log(`  [${s.id}] "${s.title}" ${s.hidden ? '(hidden)' : ''} ${s.rowCount}r × ${s.colCount}c`));

  const namedRanges = meta.data.namedRanges || [];
  console.log(`\n=== NAMED RANGES (${namedRanges.length}) ===`);
  namedRanges.forEach(nr => console.log(`  "${nr.name}" → sheetId=${nr.range.sheetId}`));

  // 2. Scan every sheet for formulas
  const allIssues = {
    google_only: {},   // fn → [{sheet, cell, formula}]
    excel_warn:  {},   // fn → [{sheet, cell, formula}]
    indirect:    [],   // [{sheet, cell, formula}]
    char10:      [],   // CHAR(10) line-break cells
    cross_sheet: [],   // formulas referencing other sheets by name
  };

  for (const sheet of sheetList) {
    const range = `'${sheet.title}'!A1:Z500`;
    let values;
    try {
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range,
        valueRenderOption: 'FORMULA',
      });
      values = res.data.values || [];
    } catch (e) {
      console.error(`  Could not read "${sheet.title}": ${e.message}`);
      continue;
    }

    let sheetFormulaCount = 0;
    values.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        if (typeof cell !== 'string' || !cell.startsWith('=')) return;
        sheetFormulaCount++;
        const colLetter = ci < 26 ? String.fromCharCode(65 + ci) : `col${ci}`;
        const cellRef = `${colLetter}${ri + 1}`;

        const issues = scanFormula(cell);
        for (const issue of issues) {
          if (issue.type === 'google_only') {
            if (!allIssues.google_only[issue.fn]) allIssues.google_only[issue.fn] = [];
            allIssues.google_only[issue.fn].push({ sheet: sheet.title, cell: cellRef });
          } else if (issue.type === 'excel_warn') {
            if (!allIssues.excel_warn[issue.fn]) allIssues.excel_warn[issue.fn] = [];
            allIssues.excel_warn[issue.fn].push({ sheet: sheet.title, cell: cellRef });
          } else if (issue.type === 'indirect') {
            allIssues.indirect.push({ sheet: sheet.title, cell: cellRef, formula: issue.formula });
          }
        }

        // CHAR(10) check
        if (/CHAR\s*\(\s*10\s*\)/i.test(cell)) {
          allIssues.char10.push({ sheet: sheet.title, cell: cellRef });
        }

        // Cross-sheet references (contains '!  and a sheet name)
        if (cell.includes('!') && !cell.includes('INDIRECT')) {
          allIssues.cross_sheet.push({ sheet: sheet.title, cell: cellRef, snippet: cell.slice(0, 60) });
        }
      });
    });
    console.log(`\n  "${sheet.title}": ${sheetFormulaCount} formula cells scanned`);
  }

  // 3. Report
  console.log('\n\n══════════════════════════════════════════════════');
  console.log('PORTABILITY ANALYSIS RESULTS');
  console.log('══════════════════════════════════════════════════');

  console.log('\n❌ GOOGLE-ONLY FUNCTIONS (will break in Excel):');
  if (Object.keys(allIssues.google_only).length === 0) {
    console.log('  None found');
  } else {
    for (const [fn, cells] of Object.entries(allIssues.google_only)) {
      const uniq = [...new Set(cells.map(c => `${c.sheet}!${c.cell}`))];
      console.log(`  ${fn}: ${uniq.length} cell(s) — e.g. ${uniq.slice(0,3).join(', ')}`);
    }
  }

  console.log('\n⚠️  EXCEL VERSION-DEPENDENT FUNCTIONS (Excel 2019/365 only):');
  if (Object.keys(allIssues.excel_warn).length === 0) {
    console.log('  None found');
  } else {
    for (const [fn, cells] of Object.entries(allIssues.excel_warn)) {
      const uniq = [...new Set(cells.map(c => `${c.sheet}!${c.cell}`))];
      console.log(`  ${fn}: ${uniq.length} cell(s) — e.g. ${uniq.slice(0,3).join(', ')}`);
    }
  }

  console.log('\n🔗 INDIRECT() USAGE (requires sheet names to stay unchanged):');
  if (allIssues.indirect.length === 0) {
    console.log('  None found');
  } else {
    const bySheet = {};
    allIssues.indirect.forEach(i => {
      bySheet[i.sheet] = (bySheet[i.sheet] || 0) + 1;
    });
    for (const [sheet, count] of Object.entries(bySheet)) {
      console.log(`  "${sheet}": ${count} INDIRECT cell(s)`);
    }
    // Show unique formula patterns
    const patterns = [...new Set(allIssues.indirect.map(i => i.formula))];
    console.log('  Unique patterns (first 5):');
    patterns.slice(0, 5).forEach(p => console.log(`    ${p}`));
  }

  console.log('\n↩️  CHAR(10) LINE BREAKS (wrapping may differ):');
  if (allIssues.char10.length === 0) {
    console.log('  None found');
  } else {
    const bySheet2 = {};
    allIssues.char10.forEach(i => { bySheet2[i.sheet] = (bySheet2[i.sheet] || 0) + 1; });
    for (const [sheet, count] of Object.entries(bySheet2)) {
      console.log(`  "${sheet}": ${count} cell(s)`);
    }
  }

  console.log('\n🔀 CROSS-SHEET REFERENCES (should survive GS copy, may break in Excel):');
  const crossBySheet = {};
  allIssues.cross_sheet.forEach(i => {
    crossBySheet[i.sheet] = (crossBySheet[i.sheet] || 0) + 1;
  });
  for (const [sheet, count] of Object.entries(crossBySheet)) {
    console.log(`  "${sheet}": ${count} cross-sheet formula(s)`);
  }

  console.log('\n══════════════════════════════════════════════════');
  console.log('GOOGLE SHEETS COPY SAFETY:');
  const gsIssues = allIssues.indirect.length;
  console.log(`  INDIRECT cells: ${gsIssues} — safe ONLY if sheet names are not renamed`);
  console.log('  Cross-sheet refs: safe — GS preserves sheet names in copies');
  console.log('  Named ranges: preserved in GS copies');

  console.log('\nMICROSOFT EXCEL COMPATIBILITY:');
  const googleOnlyCount = Object.values(allIssues.google_only).flat().length;
  const warnCount = Object.values(allIssues.excel_warn).flat().length;
  console.log(`  Google-only functions: ${googleOnlyCount} cells — WILL BREAK`);
  console.log(`  Version-gated functions: ${warnCount} cells — require Excel 2019+/365`);
  console.log(`  INDIRECT: Excel has INDIRECT but syntax for sheet names differs (!R1C1 vs A1)`);
  console.log('══════════════════════════════════════════════════');
})().catch(e => { console.error(e.message || e); process.exit(1); });
