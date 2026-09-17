'use strict';
// Reads the live Expenses tab, detects the broken YTD formulas in G3:I3,
// and rewrites them with SUMPRODUCT-based YTD calculations.

const { google } = require('googleapis');
const { getAuth, valuesBatchUpdate } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

(async () => {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  // ── Step 1: find the Expenses tab ─────────────────────────────────────────
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: id,
    fields: 'sheets(properties(sheetId,title))',
  });
  const allSheets = meta.data.sheets || [];
  const expSheet = allSheets.find(s =>
    /expense/i.test(s.properties.title)
  );
  if (!expSheet) {
    console.error('❌  No tab whose title contains "Expense" was found.');
    console.log('Tabs found:', allSheets.map(s => s.properties.title).join(', '));
    process.exit(1);
  }
  const tabTitle = expSheet.properties.title;
  const S = `'${tabTitle}'`;
  console.log(`✅  Found tab: ${tabTitle}`);

  // ── Step 2: read current G3:I3 formulas ───────────────────────────────────
  const currentRes = await sheets.spreadsheets.values.get({
    spreadsheetId: id,
    range: `${S}!G3:I3`,
    valueRenderOption: 'FORMULA',
  });
  const currentFormulas = (currentRes.data.values || [[]])[0] || [];
  console.log('\nCurrent G3:I3 formulas:');
  ['G3','H3','I3'].forEach((cell, i) => {
    console.log(`  ${cell}: ${currentFormulas[i] ?? '(empty)'}`);
  });

  // ── Step 3: read headers to detect date and amount columns ────────────────
  // Try rows 4, 5, 6 for headers (different products use different row offsets)
  let headerRow = null;
  let headerRowNum = null;
  for (const tryRow of [5, 6, 4]) {
    const hRes = await sheets.spreadsheets.values.get({
      spreadsheetId: id,
      range: `${S}!A${tryRow}:Z${tryRow}`,
      valueRenderOption: 'FORMATTED_VALUE',
    });
    const vals = (hRes.data.values || [[]])[0] || [];
    // A valid header row has multiple non-empty cells and at least one date/amount keyword
    const nonEmpty = vals.filter(v => v && String(v).trim()).length;
    const hasDateWord  = vals.some(v => /date|day|when/i.test(v));
    const hasAmountWord = vals.some(v => /amount|cost|price|total|spend|spent|\$/i.test(v));
    if (nonEmpty >= 3 && (hasDateWord || hasAmountWord)) {
      headerRow = vals;
      headerRowNum = tryRow;
      break;
    }
  }

  if (!headerRow) {
    console.log('\n⚠️  Could not auto-detect header row. Trying A5:Z5 as default.');
    const fallback = await sheets.spreadsheets.values.get({
      spreadsheetId: id, range: `${S}!A5:Z5`, valueRenderOption: 'FORMATTED_VALUE',
    });
    headerRow = (fallback.data.values || [[]])[0] || [];
    headerRowNum = 5;
  }

  console.log(`\nHeader row (row ${headerRowNum}):`, headerRow.map((h, i) =>
    `${String.fromCharCode(65+i)}=${h}`
  ).join(' | '));

  // Detect date column (A is the most common default)
  const dateColIdx = headerRow.findIndex(h => /date|day|when/i.test(h));
  const dateLetter = String.fromCharCode(65 + (dateColIdx >= 0 ? dateColIdx : 0));

  // Detect amount column — look for "amount", "cost", "price", "total", "$"
  const amtColIdx = headerRow.findIndex(h => /amount|cost|price|total|spend|spent|\$/i.test(h));
  const amtLetter = amtColIdx >= 0 ? String.fromCharCode(65 + amtColIdx) : null;

  console.log(`\n  Date column detected: ${dateLetter} (index ${dateColIdx >= 0 ? dateColIdx : 0})`);
  console.log(`  Amount column detected: ${amtLetter ?? '(not found — will need manual check)'} (index ${amtColIdx})`);

  if (!amtLetter) {
    console.error('\n❌  Could not detect an amount column. Please inspect the header row above and edit this script.');
    process.exit(1);
  }

  // ── Step 4: read G3 row (row 3) label row to understand card structure ────
  const labelRes = await sheets.spreadsheets.values.get({
    spreadsheetId: id,
    range: `${S}!A3:Z3`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const labelRow = (labelRes.data.values || [[]])[0] || [];
  console.log('\nRow 3 (label/header row):', labelRow.map((v, i) =>
    `${String.fromCharCode(65+i)}=${v || ''}`
  ).filter(s => !s.endsWith('=')).join(' | '));

  // ── Step 5: read G4:I4 to check if the values are in row 4 ───────────────
  const valRow4Res = await sheets.spreadsheets.values.get({
    spreadsheetId: id, range: `${S}!G4:I4`, valueRenderOption: 'FORMULA',
  });
  const row4Formulas = (valRow4Res.data.values || [[]])[0] || [];
  console.log('\nG4:I4 formulas (in case values live in row 4):');
  ['G4','H4','I4'].forEach((cell, i) => {
    console.log(`  ${cell}: ${row4Formulas[i] ?? '(empty)'}`);
  });

  // ── Step 6: determine which row actually has the YTD value cells ──────────
  // If G3 already looks like a formula (starts with =) → values are in row 3
  // If G3 looks like a label string → values are in row 4
  const g3IsFormula = String(currentFormulas[0] || '').startsWith('=');
  const g4IsFormula = String(row4Formulas[0] || '').startsWith('=');
  const valueRowNum = (g3IsFormula || (!g4IsFormula && currentFormulas[0])) ? 3 : 4;
  console.log(`\n  YTD values appear to be in row ${valueRowNum}`);

  // ── Step 7: determine category filter for each card ───────────────────────
  // Read G3:I4 labels for card context
  const cardLabels = valueRowNum === 3 ? labelRow.slice(6, 9) : labelRow.slice(6, 9);
  // G=index6, H=index7, I=index8
  const gLabel = labelRow[6] || '';
  const hLabel = labelRow[7] || '';
  const iLabel = labelRow[8] || '';
  console.log(`\n  G card label: "${gLabel}"`);
  console.log(`  H card label: "${hLabel}"`);
  console.log(`  I card label: "${iLabel}"`);

  // Detect category column (B or C typically)
  const catColIdx = headerRow.findIndex(h => /category|type|kind|service/i.test(h));
  const catLetter = catColIdx >= 0 ? String.fromCharCode(65 + catColIdx) : null;

  // ── Step 8: build replacement formulas ────────────────────────────────────
  // Data starts at the row AFTER the header row
  const dataStart = headerRowNum + 1;
  const dataEnd = 1000;

  // Base YTD = all expenses this year
  const baseDateRange  = `${dateLetter}$${dataStart}:${dateLetter}$${dataEnd}`;
  const baseAmtRange   = `${amtLetter}$${dataStart}:${amtLetter}$${dataEnd}`;
  const baseYTD = `=IFERROR(SUMPRODUCT((YEAR(${baseDateRange})=YEAR(TODAY()))*${baseAmtRange}),0)`;

  // Try to detect if each card filters by a category keyword
  function cardFormula(label) {
    if (!label) return baseYTD;
    // If label mentions "vet" / "medical"
    if (/vet|medical|health|clinic/i.test(label) && catLetter) {
      return `=IFERROR(SUMPRODUCT((YEAR(${baseDateRange})=YEAR(TODAY()))*(ISNUMBER(SEARCH("vet",'${tabTitle}'!${catLetter}$${dataStart}:${catLetter}$${dataEnd})))*${baseAmtRange}),0)`;
    }
    // If label mentions "food" / "nutrition"
    if (/food|nutrition|treat|feed/i.test(label) && catLetter) {
      return `=IFERROR(SUMPRODUCT((YEAR(${baseDateRange})=YEAR(TODAY()))*(ISNUMBER(SEARCH("food",'${tabTitle}'!${catLetter}$${dataStart}:${catLetter}$${dataEnd})))*${baseAmtRange}),0)`;
    }
    // If label mentions "total" or "all" → base YTD
    if (/total|all|overall|grand/i.test(label)) return baseYTD;
    // Default: just show base YTD
    return baseYTD;
  }

  const gFormula = cardFormula(gLabel);
  const hFormula = cardFormula(hLabel);
  const iFormula = cardFormula(iLabel);

  console.log('\nProposed replacement formulas:');
  console.log(`  G${valueRowNum}: ${gFormula}`);
  console.log(`  H${valueRowNum}: ${hFormula}`);
  console.log(`  I${valueRowNum}: ${iFormula}`);

  // ── Step 9: apply the fix ─────────────────────────────────────────────────
  await valuesBatchUpdate(id, [
    { range: `${S}!G${valueRowNum}`, values: [[gFormula]] },
    { range: `${S}!H${valueRowNum}`, values: [[hFormula]] },
    { range: `${S}!I${valueRowNum}`, values: [[iFormula]] },
  ], 'fix-expenses-ytd');

  console.log(`\n✅  Done. G${valueRowNum}:I${valueRowNum} updated with SUMPRODUCT-based YTD formulas.`);
  console.log('    Open the spreadsheet and verify the values look correct.');
  console.log('    If a card should filter by a different category, tell Claude the card labels and desired filter.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
