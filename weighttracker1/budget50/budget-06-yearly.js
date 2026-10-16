'use strict';
/**
 * Step 6: Populate Yearly Insights tab (sheetId 4).
 *
 * Sections:
 *   A. Monthly Rollup      rows 5-19   (Jan-Dec + Annual total)
 *   B. Category Breakdown  rows 23-49  (23 categories × 13 cols)
 *   C. Chart Helper        rows 53-66  (clean data for charts)
 *
 * All cross-tab aggregation uses SUMPRODUCT (never COUNTIFS with YEAR/MONTH).
 * Month numbers stored in column B so formula rows can be copy-pasted down.
 */
const { google } = require('googleapis');
const fs = require('fs'), path = require('path');
const AUTH_DIR = path.join(__dirname, '..');
const CFG = JSON.parse(fs.readFileSync(path.join(__dirname, 'budget-config.json')));
const SID = CFG.spreadsheetId;
const SHEET = CFG.sheets['Yearly Insights']; // sheetId 4

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(AUTH_DIR, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const auth = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  auth.setCredentials(JSON.parse(fs.readFileSync(path.join(AUTH_DIR, 'tokens.json'))));
  return auth;
}

const C = {
  cream:      { red:0.965, green:0.953, blue:0.922 },
  paleSage:   { red:0.906, green:0.929, blue:0.898 },
  softSage:   { red:0.651, green:0.725, blue:0.631 },
  moss:       { red:0.443, green:0.545, blue:0.475 },
  deepForest: { red:0.208, green:0.282, blue:0.243 },
  amber:      { red:0.953, green:0.894, blue:0.784 },
  sand:       { red:0.851, green:0.788, blue:0.682 },
  white:      { red:1, green:1, blue:1 },
};

// Short references to Transactions columns (1-indexed col letters)
const T = "'Transactions'!";
const DR = '$B$10:$B$1009'; // Date range
const ER = '$E$10:$E$1009'; // Category range
const GR = '$G$10:$G$1009'; // Amount range
const PR = '$P$10:$P$1009'; // Needs
const QR = '$Q$10:$Q$1009'; // Wants
const RR = '$R$10:$R$1009'; // Future
const SR = '$S$10:$S$1009'; // Income
const YR = "'Start Here'!$B$10"; // Budget Year

// Generates monthly aggregation formula for a given Transactions column.
// rowRef: e.g. '$B7' — the cell on THIS sheet holding the month number
function mf(txnCol, rowRef, negate = false) {
  const sign = negate ? '-' : '';
  return `${sign}SUMPRODUCT((MONTH(${T}${DR})=${rowRef})*(YEAR(${T}${DR})=${YR})*(${T}${txnCol}))`;
}

// Category-specific monthly formula
function catMF(monthRef, catCell) {
  return `=IFERROR(SUMPRODUCT((MONTH(${T}${DR})=${monthRef})*(YEAR(${T}${DR})=${YR})*(${T}${ER}=${catCell})*ABS(${T}${GR})),0)`;
}

const MONTHS = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];

const CATEGORIES = [
  'Housing','Utilities','Groceries','Transportation','Insurance','Healthcare',
  'Minimum Debt Payments','Childcare',
  'Dining Out','Entertainment','Shopping','Streaming','Hobbies','Travel','Coffee',
  'Emergency Fund','Retirement Contribution','Extra Debt Payment','Vacation Fund',
  'Home Down Payment','Other Savings',
  'Paycheck','Account Transfer',
];

(async () => {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  // ── Section A: Monthly Rollup ──────────────────────────────────────────────
  // Row 5 (0-indexed 4): section header
  // Row 6 (5): column headers — A Month, B #, C Income, D Needs, E Wants, F Future, G Expenses, H Net, I Needs%, J Wants%, K Future%
  // Rows 7-18 (6-17): Jan-Dec data
  // Row 19 (18): Annual Total

  const rollupHeaders = ['Month','#','Income','Needs Spent','Wants Spent','Future Saved','Total Expenses','Net Cash Flow','Needs %','Wants %','Future %'];

  // Build formula row for row 7 (January, month# in B7):
  // We write the formulas in row 7 and copyPaste to rows 8-18
  const janFormulas = [
    'January', 1,
    `=${mf(SR, '$B7')}`,                        // C: Income
    `=${mf(PR, '$B7', true)}`,                  // D: Needs (positive spending)
    `=${mf(QR, '$B7', true)}`,                  // E: Wants (positive spending)
    `=${mf(RR, '$B7', true)}`,                  // F: Future (positive transfers)
    `=D7+E7`,                                   // G: Total Expenses = Needs+Wants
    `=C7-D7-E7-F7`,                             // H: Net
    `=IFERROR(D7/C7,0)`,                        // I: Needs %
    `=IFERROR(E7/C7,0)`,                        // J: Wants %
    `=IFERROR(F7/C7,0)`,                        // K: Future %
  ];

  // Month name and number for rows 8-18
  const monthRows = MONTHS.slice(1).map((m, i) => [m, i + 2]); // Feb=2 ... Dec=12

  // Annual total row (row 19, index 18)
  const totalRow = ['ANNUAL TOTAL', '', '=SUM(C7:C18)', '=SUM(D7:D18)', '=SUM(E7:E18)',
                    '=SUM(F7:F18)', '=SUM(G7:G18)', '=SUM(H7:H18)',
                    '=IFERROR(D19/C19,0)', '=IFERROR(E19/C19,0)', '=IFERROR(F19/C19,0)'];

  // ── Section B: Category Breakdown ─────────────────────────────────────────
  // Row 23 (22): section header
  // Row 24 (23): col headers — A Category, B Bucket, C Annual, D-O = Jan-Dec
  // Rows 25-47 (24-46): 23 categories

  const catColHeaders = ['Category', 'Bucket', 'Annual Total',
    'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  // Build category rows: each row has the category name, VLOOKUP for bucket, annual sum, then 12 monthly formulas
  const catDataRows = CATEGORIES.map((cat, ci) => {
    const row = 25 + ci; // rows 25-47
    const months12 = Array.from({length: 12}, (_, m) => catMF(m + 1, `$A${row}`));
    const annualFormula = `=IFERROR(SUM(D${row}:O${row}),0)`;
    const bucketFormula = `=IFERROR(VLOOKUP($A${row},'Start Here'!$B$32:$C$54,2,0),"")`;
    return [cat, bucketFormula, annualFormula, ...months12];
  });

  // ── Section C: Chart Helper (clean pivot for charts) ──────────────────────
  // Row 53 (52): section header
  // Row 54 (53): headers
  // Rows 55-66 (54-65): month rows with income, expenses, net
  const chartHeaders = ['Month', 'Income', 'Needs', 'Wants', 'Future', 'Total Expenses', 'Net'];
  const chartRows = MONTHS.map((m, i) => {
    const srcRow = 7 + i; // rows 7-18 in monthly rollup
    return [m, `=C${srcRow}`, `=D${srcRow}`, `=E${srcRow}`, `=F${srcRow}`, `=G${srcRow}`, `=H${srcRow}`];
  });

  // ── Write all data ─────────────────────────────────────────────────────────
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [
        { range: `'Yearly Insights'!A1`, values: [['YEARLY INSIGHTS']] },
        { range: `'Yearly Insights'!A2`, values: [[`=TEXT('Start Here'!$B$10,"0")&" Annual Budget Summary — BuildSheetStudio"`]] },
        { range: `'Yearly Insights'!A5`, values: [['📊  MONTHLY ROLLUP']] },
        { range: `'Yearly Insights'!A6:K6`, values: [rollupHeaders] },
        { range: `'Yearly Insights'!A7:K7`, values: [janFormulas] },
        { range: `'Yearly Insights'!A19:K19`, values: [totalRow] },
        { range: `'Yearly Insights'!A23`, values: [['📋  SPENDING BY CATEGORY']] },
        { range: `'Yearly Insights'!A24:O24`, values: [catColHeaders] },
        { range: `'Yearly Insights'!A25:O47`, values: catDataRows },
        { range: `'Yearly Insights'!A53`, values: [['📈  CHART DATA (Monthly Income vs. Expenses)']] },
        { range: `'Yearly Insights'!A54:G54`, values: [chartHeaders] },
        { range: `'Yearly Insights'!A55:G66`, values: chartRows },
      ],
    },
  });

  // ── Fill month names + numbers rows 8-18 ─────────────────────────────────
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [{ range: `'Yearly Insights'!A8:B18`, values: monthRows }],
    },
  });

  // ── Copy formula row 7 (C-K) down to rows 8-18 ────────────────────────────
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      requests: [{
        copyPaste: {
          source: { sheetId: SHEET, startRowIndex: 6, endRowIndex: 7, startColumnIndex: 2, endColumnIndex: 11 },
          destination: { sheetId: SHEET, startRowIndex: 6, endRowIndex: 18, startColumnIndex: 2, endColumnIndex: 11 },
          pasteType: 'PASTE_FORMULA', pasteOrientation: 'NORMAL',
        },
      }],
    },
  });

  // ── Formatting ─────────────────────────────────────────────────────────────
  const requests = [];

  // Full sheet cream bg
  requests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 0, endRowIndex: 200, startColumnIndex: 0, endColumnIndex: 20 },
      cell: { userEnteredFormat: { backgroundColor: C.cream } },
      fields: 'userEnteredFormat.backgroundColor',
    },
  });

  // Title row 1
  requests.push(
    { mergeCells: { range: { sheetId: SHEET, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 15 }, mergeType: 'MERGE_ALL' } },
    { repeatCell: {
        range: { sheetId: SHEET, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 15 },
        cell: { userEnteredFormat: { backgroundColor: C.deepForest, textFormat: { bold: true, foregroundColor: C.white, fontSize: 16 }, padding: { top: 12, bottom: 12, left: 12 } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
      },
    }
  );

  // Subtitle row 2
  requests.push(
    { mergeCells: { range: { sheetId: SHEET, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 15 }, mergeType: 'MERGE_ALL' } },
    { repeatCell: {
        range: { sheetId: SHEET, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 15 },
        cell: { userEnteredFormat: { backgroundColor: C.moss, textFormat: { italic: true, foregroundColor: C.white }, padding: { top: 5, bottom: 5, left: 12 } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
      },
    }
  );

  // Section headers
  const secHdr = (r, cols = 11) => [
    { mergeCells: { range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: cols }, mergeType: 'MERGE_ALL' } },
    { repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: cols },
        cell: { userEnteredFormat: { backgroundColor: C.deepForest, textFormat: { bold: true, foregroundColor: C.white, fontSize: 12 }, padding: { top: 7, bottom: 7, left: 10 } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
      },
    },
  ];

  requests.push(
    ...secHdr(4, 11), // Monthly rollup header
    ...secHdr(22, 15), // Category header
    ...secHdr(52, 7), // Chart helper header
  );

  // Column headers (row 6, 24, 54)
  [5, 23, 53].forEach(r => {
    const cols = r === 23 ? 15 : (r === 53 ? 7 : 11);
    requests.push({
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: cols },
        cell: { userEnteredFormat: { backgroundColor: C.softSage, textFormat: { bold: true, foregroundColor: C.white }, padding: { top: 4, bottom: 4, left: 6 } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
      },
    });
  });

  // Monthly data rows 7-18 alternating
  for (let r = 6; r <= 17; r++) {
    const bg = r % 2 === 0 ? C.cream : C.paleSage;
    requests.push({
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: 11 },
        cell: { userEnteredFormat: { backgroundColor: bg, padding: { top: 3, bottom: 3, left: 6 } } },
        fields: 'userEnteredFormat(backgroundColor,padding)',
      },
    });
  }

  // Annual total row 19 (index 18)
  requests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 18, endRowIndex: 19, startColumnIndex: 0, endColumnIndex: 11 },
      cell: { userEnteredFormat: { backgroundColor: C.moss, textFormat: { bold: true, foregroundColor: C.white }, padding: { top: 5, bottom: 5, left: 6 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
    },
  });

  // Category data rows 25-47 alternating
  for (let r = 24; r <= 46; r++) {
    const bg = r % 2 === 0 ? C.cream : C.paleSage;
    requests.push({
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: 15 },
        cell: { userEnteredFormat: { backgroundColor: bg, padding: { top: 3, bottom: 3, left: 6 } } },
        fields: 'userEnteredFormat(backgroundColor,padding)',
      },
    });
  }

  // Chart helper rows 55-66 alternating
  for (let r = 54; r <= 65; r++) {
    const bg = r % 2 === 0 ? C.cream : C.paleSage;
    requests.push({
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: 7 },
        cell: { userEnteredFormat: { backgroundColor: bg, padding: { top: 3, bottom: 3, left: 6 } } },
        fields: 'userEnteredFormat(backgroundColor,padding)',
      },
    });
  }

  // Currency format: monthly rollup C-H, cat C-O, chart B-F
  requests.push(
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 6, endRowIndex: 19, startColumnIndex: 2, endColumnIndex: 8 }, cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00;[RED]"-$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } },
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 6, endRowIndex: 19, startColumnIndex: 8, endColumnIndex: 11 }, cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } },
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 24, endRowIndex: 47, startColumnIndex: 2, endColumnIndex: 15 }, cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } },
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 54, endRowIndex: 66, startColumnIndex: 1, endColumnIndex: 7 }, cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } },
  );

  // Column widths
  const widths = [130, 40, 110, 90, 90, 90, 110, 110, 80, 80, 80, 80, 80, 80, 80];
  widths.forEach((w, i) => {
    requests.push({
      updateDimensionProperties: {
        range: { sheetId: SHEET, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
        properties: { pixelSize: w }, fields: 'pixelSize',
      },
    });
  });

  // Freeze rows 1-6 (title + section + headers)
  requests.push({
    updateSheetProperties: {
      properties: { sheetId: SHEET, gridProperties: { frozenRowCount: 6 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await sheets.spreadsheets.batchUpdate({ spreadsheetId: SID, requestBody: { requests } });

  // ── Charts ────────────────────────────────────────────────────────────────
  // Chart 1: Monthly Income vs. Expenses (column chart using chart helper data)
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      requests: [{
        addChart: {
          chart: {
            spec: {
              title: 'Monthly Income vs. Total Expenses',
              basicChart: {
                chartType: 'COLUMN',
                legendPosition: 'BOTTOM_LEGEND',
                axis: [
                  { position: 'BOTTOM_AXIS', title: 'Month' },
                  { position: 'LEFT_AXIS', title: 'Amount ($)' },
                ],
                domains: [{
                  domain: { sourceRange: { sources: [{ sheetId: SHEET, startRowIndex: 54, endRowIndex: 66, startColumnIndex: 0, endColumnIndex: 1 }] } },
                }],
                series: [
                  { series: { sourceRange: { sources: [{ sheetId: SHEET, startRowIndex: 54, endRowIndex: 66, startColumnIndex: 1, endColumnIndex: 2 }] } }, targetAxis: 'LEFT_AXIS', color: { red:0.443, green:0.545, blue:0.475 } },
                  { series: { sourceRange: { sources: [{ sheetId: SHEET, startRowIndex: 54, endRowIndex: 66, startColumnIndex: 5, endColumnIndex: 6 }] } }, targetAxis: 'LEFT_AXIS', color: { red:0.847, green:0.647, blue:0.604 } },
                ],
                headerCount: 1,
              },
            },
            position: { overlayPosition: { anchorCell: { sheetId: SHEET, rowIndex: 70, columnIndex: 0 }, widthPixels: 560, heightPixels: 340 } },
          },
        },
      }],
    },
  });

  // Chart 2: Needs/Wants/Future monthly breakdown (COMBO — columns for each)
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      requests: [{
        addChart: {
          chart: {
            spec: {
              title: 'Monthly Spending by Bucket (Needs / Wants / Future)',
              basicChart: {
                chartType: 'COMBO',
                legendPosition: 'BOTTOM_LEGEND',
                axis: [
                  { position: 'BOTTOM_AXIS', title: 'Month' },
                  { position: 'LEFT_AXIS', title: 'Amount ($)' },
                ],
                domains: [{
                  domain: { sourceRange: { sources: [{ sheetId: SHEET, startRowIndex: 54, endRowIndex: 66, startColumnIndex: 0, endColumnIndex: 1 }] } },
                }],
                series: [
                  { series: { sourceRange: { sources: [{ sheetId: SHEET, startRowIndex: 54, endRowIndex: 66, startColumnIndex: 2, endColumnIndex: 3 }] } }, targetAxis: 'LEFT_AXIS', type: 'COLUMN', color: { red:0.651, green:0.725, blue:0.631 } },
                  { series: { sourceRange: { sources: [{ sheetId: SHEET, startRowIndex: 54, endRowIndex: 66, startColumnIndex: 3, endColumnIndex: 4 }] } }, targetAxis: 'LEFT_AXIS', type: 'COLUMN', color: { red:0.953, green:0.894, blue:0.784 } },
                  { series: { sourceRange: { sources: [{ sheetId: SHEET, startRowIndex: 54, endRowIndex: 66, startColumnIndex: 4, endColumnIndex: 5 }] } }, targetAxis: 'LEFT_AXIS', type: 'COLUMN', color: { red:0.847, green:0.647, blue:0.604 } },
                ],
                headerCount: 1,
              },
            },
            position: { overlayPosition: { anchorCell: { sheetId: SHEET, rowIndex: 70, columnIndex: 7 }, widthPixels: 560, heightPixels: 340 } },
          },
        },
      }],
    },
  });

  console.log('✅ Yearly Insights tab populated:');
  console.log('   Monthly rollup: rows 7-18 (Jan-Dec) + Annual total row 19');
  console.log('   Category breakdown: rows 25-47 (23 categories × 13 months)');
  console.log('   Chart helper: rows 55-66 (monthly pivot for charts)');
  console.log('   Charts: 2 charts (column + combo)');
  console.log('   All formulas use SUMPRODUCT — zero COUNTIFS with YEAR/MONTH');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
