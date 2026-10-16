'use strict';
/**
 * Step 7: Populate Budget Dashboard tab (sheetId 5).
 *
 * Key helper cells:
 *   B5 = period reference date (user-editable; default Jan 15, 2026)
 *   D5 = period start (formula)
 *   F5 = period end (formula)
 *
 * Sections:
 *   Row 1-2: Title/subtitle
 *   Row 4-5: Period selector
 *   Row 7-9: 4 period stat cards (Income / Expenses / Net / Savings)
 *   Row 11-15: 50/30/20 Budget Progress
 *   Row 17-22: Savings Goals Snapshot
 *   Row 24-34: Upcoming Bills
 *   Row 36-43: Paycheck Allocation
 *   Row 45+: Chart helper data
 */
const { google } = require('googleapis');
const fs = require('fs'), path = require('path');
const AUTH_DIR = path.join(__dirname, '..');
const CFG = JSON.parse(fs.readFileSync(path.join(__dirname, 'budget-config.json')));
const SID = CFG.spreadsheetId;
const SHEET = CFG.sheets['Budget Dashboard']; // sheetId 5

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
  sand:       { red:0.851, green:0.788, blue:0.682 },
  amber:      { red:0.953, green:0.894, blue:0.784 },
  coral:      { red:0.847, green:0.647, blue:0.604 },
  white:      { red:1, green:1, blue:1 },
  posGreen:   { red:0.15, green:0.55, blue:0.25 },
};

// Period helper references
const pStart = `$D$5`;    // period start cell
const pEnd   = `$F$5`;    // period end cell
const T = "'Transactions'!";
const DR = '$B$10:$B$1009';
const PR = '$P$10:$P$1009'; // Needs
const QR = '$Q$10:$Q$1009'; // Wants
const RR = '$R$10:$R$1009'; // Future
const SR = '$S$10:$S$1009'; // Income
const YR = `'Start Here'!$B$10`; // Budget Year
const CAD = `'Start Here'!$B$11`; // Cadence

// Period SUMPRODUCT filter
const pFilter = `(${T}${DR}>=${pStart})*(${T}${DR}<=${pEnd})`;

// Budget for a bucket based on cadence
function budgetFormula(pctCell) {
  // pctCell e.g. 'Start Here'!$B$14 (Needs %)
  const n  = `'Start Here'!$B$17`; // Monthly take-home
  const bw = `'Start Here'!$B$18`; // Biweekly
  const wk = `'Start Here'!$B$19`; // Weekly
  return `=IF(${CAD}="Monthly",${n}*${pctCell}/100,IF(${CAD}="Biweekly",${bw}*${pctCell}/100,${wk}*${pctCell}/100))`;
}

(async () => {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  // ── Main content ───────────────────────────────────────────────────────────
  const data = [
    // Row 1: Title
    { range: `'Budget Dashboard'!A1`, values: [['BUDGET DASHBOARD']] },

    // Row 2: Subtitle
    { range: `'Budget Dashboard'!A2`, values: [['Your 50/30/20 plan at a glance — powered by BuildSheetStudio']] },

    // Row 4: Period selector labels
    { range: `'Budget Dashboard'!A4:H4`, values: [[
      '📅 Reference Date', '← Change me', '', 'Period Start', '', 'Period End', '', 'Cadence',
    ]]},

    // Row 5: Period selector values
    { range: `'Budget Dashboard'!A5:H5`, values: [[
      'Date:',
      '1/15/2026',                  // B5: user-editable demo date
      '',
      // D5: Period Start
      `=IF($B$5="","",IF(${CAD}="Weekly",$B$5-WEEKDAY($B$5,2)+1,IF(${CAD}="Biweekly",'Start Here'!$B$13+14*INT(($B$5-'Start Here'!$B$13)/14),DATE(YEAR($B$5),MONTH($B$5),1))))`,
      'to',
      // F5: Period End
      `=IF($D$5="","",IF(${CAD}="Weekly",$D$5+6,IF(${CAD}="Biweekly",$D$5+13,EOMONTH($D$5,0))))`,
      '',
      `=${CAD}`,                    // H5: Cadence display
    ]]},

    // Row 6: Period label display
    { range: `'Budget Dashboard'!A6`, values: [[
      `=IFERROR("Showing: "&TEXT($D$5,"mmmm d")&" – "&TEXT($F$5,"mmmm d, yyyy"),"")`
    ]]},

    // Row 7: Section header
    { range: `'Budget Dashboard'!A7`, values: [['📊  PERIOD SUMMARY']] },

    // Row 8: Stat card labels
    { range: `'Budget Dashboard'!A8:H8`, values: [[
      'INCOME', '', 'EXPENSES', '', 'NET CASH FLOW', '', 'SAVINGS TRANSFERRED', '',
    ]]},

    // Row 9: Stat card values
    { range: `'Budget Dashboard'!A9:H9`, values: [[
      `=IFERROR(SUMPRODUCT(${pFilter}*(${T}${SR})),0)`,
      '',
      `=IFERROR(-SUMPRODUCT(${pFilter}*(${T}${PR}))-SUMPRODUCT(${pFilter}*(${T}${QR})),0)`,
      '',
      `=IFERROR(SUMPRODUCT(${pFilter}*(${T}${SR}))-(-SUMPRODUCT(${pFilter}*(${T}${PR}))-SUMPRODUCT(${pFilter}*(${T}${QR})))-(-SUMPRODUCT(${pFilter}*(${T}${RR}))),0)`,
      '',
      `=IFERROR(-SUMPRODUCT(${pFilter}*(${T}${RR})),0)`,
      '',
    ]]},

    // Row 11: Section header — 50/30/20 Progress
    { range: `'Budget Dashboard'!A11`, values: [['🎯  50/30/20 BUDGET PROGRESS']] },

    // Row 12: Budget progress headers
    { range: `'Budget Dashboard'!A12:F12`, values: [['Bucket','Target %','Budget','Spent','Remaining','% Used']] },

    // Row 13: Needs
    { range: `'Budget Dashboard'!A13:F13`, values: [[
      'Needs',
      `='Start Here'!$B$14&"%"`,
      budgetFormula(`'Start Here'!$B$14`),
      `=IFERROR(-SUMPRODUCT(${pFilter}*(${T}${PR})),0)`,
      `=MAX(C13-D13,0)`,
      `=IFERROR(D13/C13,0)`,
    ]]},

    // Row 14: Wants
    { range: `'Budget Dashboard'!A14:F14`, values: [[
      'Wants',
      `='Start Here'!$B$15&"%"`,
      budgetFormula(`'Start Here'!$B$15`),
      `=IFERROR(-SUMPRODUCT(${pFilter}*(${T}${QR})),0)`,
      `=MAX(C14-D14,0)`,
      `=IFERROR(D14/C14,0)`,
    ]]},

    // Row 15: Future
    { range: `'Budget Dashboard'!A15:F15`, values: [[
      'Future',
      `='Start Here'!$B$16&"%"`,
      budgetFormula(`'Start Here'!$B$16`),
      `=IFERROR(-SUMPRODUCT(${pFilter}*(${T}${RR})),0)`,
      `=MAX(C15-D15,0)`,
      `=IFERROR(D15/C15,0)`,
    ]]},

    // Row 16: % validation check
    { range: `'Budget Dashboard'!A16`, values: [[
      `=IF('Start Here'!$B$14+'Start Here'!$B$15+'Start Here'!$B$16=100,"✓ Percentages = 100%","⚠ "&('Start Here'!$B$14+'Start Here'!$B$15+'Start Here'!$B$16)&"% — adjust in Start Here")`
    ]]},

    // Row 18: Savings Goals Snapshot header
    { range: `'Budget Dashboard'!A18`, values: [['🌱  SAVINGS GOALS SNAPSHOT']] },

    // Row 19: Headers
    { range: `'Budget Dashboard'!A19:E19`, values: [['Goal','Target','Saved','Monthly Contrib','% Complete']] },

    // Rows 20-23: Goals (reference Bills & Goals tab)
    { range: `'Budget Dashboard'!A20:E23`, values: [
      [`='Bills & Goals'!B320`, `='Bills & Goals'!D320`, `='Bills & Goals'!E320`, `='Bills & Goals'!F320`, `=IFERROR('Bills & Goals'!H320,0)`],
      [`='Bills & Goals'!B321`, `='Bills & Goals'!D321`, `='Bills & Goals'!E321`, `='Bills & Goals'!F321`, `=IFERROR('Bills & Goals'!H321,0)`],
      [`='Bills & Goals'!B322`, `='Bills & Goals'!D322`, `='Bills & Goals'!E322`, `='Bills & Goals'!F322`, `=IFERROR('Bills & Goals'!H322,0)`],
      [`='Bills & Goals'!B323`, `='Bills & Goals'!D323`, `='Bills & Goals'!E323`, `='Bills & Goals'!F323`, `=IFERROR('Bills & Goals'!H323,0)`],
    ]},

    // Row 25: Upcoming Bills header
    { range: `'Budget Dashboard'!A25`, values: [['📅  UPCOMING BILLS (Current Month)']] },

    // Row 26: Headers
    { range: `'Budget Dashboard'!A26:E26`, values: [['Bill Name','Frequency','Amount','Due Day','Status']] },

    // Rows 27-36: Up to 10 bills from registry (reference Bills & Goals)
    { range: `'Budget Dashboard'!A27:E36`, values: Array.from({length:10}, (_, i) => {
        const r = 10 + i; // Bills & Goals rows 10-19
        return [
          `=IFERROR('Bills & Goals'!B${r},"")`,
          `=IFERROR('Bills & Goals'!D${r},"")`,
          `=IFERROR('Bills & Goals'!E${r},"")`,
          `=IFERROR('Bills & Goals'!F${r},"")`,
          `=IFERROR('Bills & Goals'!I${r},"")`,
        ];
      })
    },

    // Row 38: Paycheck Allocation header
    { range: `'Budget Dashboard'!A38`, values: [['💰  PAYCHECK ALLOCATION (Per-Period Breakdown)']] },

    // Row 39: Labels
    { range: `'Budget Dashboard'!A39:D39`, values: [['Period Income','Needs Budget','Wants Budget','Future Budget']] },

    // Row 40: Formulas
    { range: `'Budget Dashboard'!A40:D40`, values: [[
      `=IFERROR(SUMPRODUCT(${pFilter}*(${T}${SR})),0)`,
      budgetFormula(`'Start Here'!$B$14`),
      budgetFormula(`'Start Here'!$B$15`),
      budgetFormula(`'Start Here'!$B$16`),
    ]]},

    // Row 41: Unbudgeted line
    { range: `'Budget Dashboard'!A41:D41`, values: [[
      'Budget Surplus / Deficit',
      '',
      '',
      `=IFERROR(A40-(B40+C40+D40),0)`,
    ]]},

    // ── Chart helper data (rows 45-50) ───────────────────────────────────────
    { range: `'Budget Dashboard'!A45`, values: [['Chart Helper — Period Bucket Split']] },
    { range: `'Budget Dashboard'!A46:B46`, values: [['Bucket','Amount']] },
    { range: `'Budget Dashboard'!A47:B49`, values: [
      ['Needs',  `=IFERROR(-SUMPRODUCT(${pFilter}*(${T}${PR})),0)`],
      ['Wants',  `=IFERROR(-SUMPRODUCT(${pFilter}*(${T}${QR})),0)`],
      ['Future', `=IFERROR(-SUMPRODUCT(${pFilter}*(${T}${RR})),0)`],
    ]},
  ];

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SID,
    requestBody: { valueInputOption: 'USER_ENTERED', data },
  });

  // ── Formatting ─────────────────────────────────────────────────────────────
  const requests = [];

  // Full sheet cream
  requests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 0, endRowIndex: 120, startColumnIndex: 0, endColumnIndex: 16 },
      cell: { userEnteredFormat: { backgroundColor: C.cream } },
      fields: 'userEnteredFormat.backgroundColor',
    },
  });

  // Title row 1 (index 0)
  requests.push(
    { mergeCells: { range: { sheetId: SHEET, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 8 }, mergeType: 'MERGE_ALL' } },
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 8 }, cell: { userEnteredFormat: { backgroundColor: C.deepForest, textFormat: { bold: true, foregroundColor: C.white, fontSize: 18 }, padding: { top: 14, bottom: 14, left: 16 }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,padding,verticalAlignment)' } }
  );

  // Subtitle row 2 (index 1)
  requests.push(
    { mergeCells: { range: { sheetId: SHEET, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 8 }, mergeType: 'MERGE_ALL' } },
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 8 }, cell: { userEnteredFormat: { backgroundColor: C.moss, textFormat: { italic: true, foregroundColor: C.white, fontSize: 11 }, padding: { top: 6, bottom: 6, left: 16 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat,padding)' } }
  );

  // Period selector rows 4-6 — sage bg
  requests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 3, endRowIndex: 6, startColumnIndex: 0, endColumnIndex: 8 },
      cell: { userEnteredFormat: { backgroundColor: C.paleSage, padding: { top: 4, bottom: 4, left: 8 } } },
      fields: 'userEnteredFormat(backgroundColor,padding)',
    },
  });

  // B5 = editable date cell — highlight with sand bg and bold
  requests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 4, endRowIndex: 5, startColumnIndex: 1, endColumnIndex: 2 },
      cell: { userEnteredFormat: { backgroundColor: C.amber, textFormat: { bold: true }, numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat,horizontalAlignment)',
    },
  });

  // D5 / F5 period start/end: date format
  requests.push(
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 4, endRowIndex: 5, startColumnIndex: 3, endColumnIndex: 4 }, cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' }, textFormat: { bold: true }, backgroundColor: C.sand } }, fields: 'userEnteredFormat(numberFormat,textFormat,backgroundColor)' } },
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 4, endRowIndex: 5, startColumnIndex: 5, endColumnIndex: 6 }, cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' }, textFormat: { bold: true }, backgroundColor: C.sand } }, fields: 'userEnteredFormat(numberFormat,textFormat,backgroundColor)' } },
  );

  // Section header helper
  const secHdr = (r, cols = 8) => [
    { mergeCells: { range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: cols }, mergeType: 'MERGE_ALL' } },
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: cols }, cell: { userEnteredFormat: { backgroundColor: C.deepForest, textFormat: { bold: true, foregroundColor: C.white, fontSize: 12 }, padding: { top: 7, bottom: 7, left: 10 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat,padding)' } },
  ];

  requests.push(
    ...secHdr(6, 8),   // PERIOD SUMMARY
    ...secHdr(10, 8),  // Not used — but row 11
    ...secHdr(17, 6),  // SAVINGS GOALS
    ...secHdr(24, 6),  // UPCOMING BILLS
    ...secHdr(37, 6),  // PAYCHECK ALLOCATION
    ...secHdr(44, 4),  // CHART HELPER
  );

  // 50/30/20 section header (row 11, index 10)
  requests.push(
    { mergeCells: { range: { sheetId: SHEET, startRowIndex: 10, endRowIndex: 11, startColumnIndex: 0, endColumnIndex: 8 }, mergeType: 'MERGE_ALL' } },
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 10, endRowIndex: 11, startColumnIndex: 0, endColumnIndex: 8 }, cell: { userEnteredFormat: { backgroundColor: C.deepForest, textFormat: { bold: true, foregroundColor: C.white, fontSize: 12 }, padding: { top: 7, bottom: 7, left: 10 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat,padding)' } }
  );

  // Stat cards row 8 (index 7): labels
  requests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 7, endRowIndex: 8, startColumnIndex: 0, endColumnIndex: 8 },
      cell: { userEnteredFormat: { backgroundColor: C.softSage, textFormat: { bold: true, foregroundColor: C.white, fontSize: 10 }, horizontalAlignment: 'CENTER', padding: { top: 6, bottom: 4, left: 4 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,padding)',
    },
  });

  // Stat cards row 9 (index 8): values
  requests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 8, endRowIndex: 9, startColumnIndex: 0, endColumnIndex: 8 },
      cell: { userEnteredFormat: { backgroundColor: C.paleSage, textFormat: { bold: true, fontSize: 16 }, numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00;[RED]"-$"#,##0.00' }, horizontalAlignment: 'CENTER', padding: { top: 10, bottom: 10, left: 4 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat,horizontalAlignment,padding)',
    },
  });

  // Row heights
  requests.push(
    { updateDimensionProperties: { range: { sheetId: SHEET, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 54 }, fields: 'pixelSize' } },
    { updateDimensionProperties: { range: { sheetId: SHEET, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } },
    { updateDimensionProperties: { range: { sheetId: SHEET, dimension: 'ROWS', startIndex: 8, endIndex: 9 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } },
  );

  // Column header rows for subsections
  [[11, 6], [18, 5], [25, 5], [38, 4]].forEach(([r, cols]) => {
    requests.push({
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: cols },
        cell: { userEnteredFormat: { backgroundColor: C.softSage, textFormat: { bold: true, foregroundColor: C.white }, padding: { top: 4, bottom: 4, left: 6 } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
      },
    });
  });

  // 50/30/20 progress rows 13-15 (indices 12-14): alternating + currency + %
  [[12, C.paleSage], [13, C.cream], [14, C.amber]].forEach(([r, bg]) => {
    requests.push({
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: 6 },
        cell: { userEnteredFormat: { backgroundColor: bg, padding: { top: 4, bottom: 4, left: 8 }, textFormat: { bold: true } } },
        fields: 'userEnteredFormat(backgroundColor,padding,textFormat)',
      },
    });
  });
  // Currency format cols C-E (indices 2-4) for progress rows
  requests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 12, endRowIndex: 15, startColumnIndex: 2, endColumnIndex: 5 },
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });
  // % format col F (index 5) for progress rows
  requests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 12, endRowIndex: 15, startColumnIndex: 5, endColumnIndex: 6 },
      cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // Goals snapshot rows 20-23 alternating
  for (let r = 19; r <= 22; r++) {
    const bg = r % 2 === 0 ? C.paleSage : C.cream;
    requests.push({
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: 5 },
        cell: { userEnteredFormat: { backgroundColor: bg, padding: { top: 3, bottom: 3, left: 6 } } },
        fields: 'userEnteredFormat(backgroundColor,padding)',
      },
    });
  }
  requests.push(
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 19, endRowIndex: 23, startColumnIndex: 1, endColumnIndex: 4 }, cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } },
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 19, endRowIndex: 23, startColumnIndex: 4, endColumnIndex: 5 }, cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } },
  );

  // Upcoming bills rows 27-36 alternating
  for (let r = 26; r <= 35; r++) {
    const bg = r % 2 === 0 ? C.cream : C.paleSage;
    requests.push({
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: 5 },
        cell: { userEnteredFormat: { backgroundColor: bg, padding: { top: 3, bottom: 3, left: 6 } } },
        fields: 'userEnteredFormat(backgroundColor,padding)',
      },
    });
  }
  requests.push({ repeatCell: { range: { sheetId: SHEET, startRowIndex: 26, endRowIndex: 36, startColumnIndex: 2, endColumnIndex: 3 }, cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Paycheck allocation rows 40-41
  requests.push(
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 38, endRowIndex: 40, startColumnIndex: 0, endColumnIndex: 4 }, cell: { userEnteredFormat: { backgroundColor: C.paleSage, textFormat: { bold: true }, padding: { top: 4, bottom: 4, left: 6 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat,padding)' } },
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 39, endRowIndex: 41, startColumnIndex: 0, endColumnIndex: 4 }, cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00;[RED]"-$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } },
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 40, endRowIndex: 41, startColumnIndex: 0, endColumnIndex: 4 }, cell: { userEnteredFormat: { backgroundColor: C.amber } }, fields: 'userEnteredFormat.backgroundColor' } },
  );

  // Column widths
  const widths = [180, 120, 120, 120, 120, 100, 40, 140];
  widths.forEach((w, i) => {
    requests.push({
      updateDimensionProperties: {
        range: { sheetId: SHEET, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
        properties: { pixelSize: w }, fields: 'pixelSize',
      },
    });
  });

  // Freeze rows 1-6 (title + period selector)
  requests.push({
    updateSheetProperties: {
      properties: { sheetId: SHEET, gridProperties: { frozenRowCount: 6 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await sheets.spreadsheets.batchUpdate({ spreadsheetId: SID, requestBody: { requests } });

  // ── Conditional formatting ─────────────────────────────────────────────────
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      requests: [
        // % Used >= 100%: red background on F col (index 5), progress rows
        {
          addConditionalFormatRule: {
            rule: {
              ranges: [{ sheetId: SHEET, startRowIndex: 12, endRowIndex: 15, startColumnIndex: 5, endColumnIndex: 6 }],
              booleanRule: {
                condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '1' }] },
                format: { backgroundColor: { red: 0.95, green: 0.8, blue: 0.8 }, textFormat: { foregroundColor: { red: 0.7, green: 0.1, blue: 0.1 }, bold: true } },
              },
            }, index: 0,
          },
        },
        // Net Cash Flow positive = green (A9 / col 0, stat card row index 8)
        {
          addConditionalFormatRule: {
            rule: {
              ranges: [{ sheetId: SHEET, startRowIndex: 8, endRowIndex: 9, startColumnIndex: 4, endColumnIndex: 5 }],
              booleanRule: {
                condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0' }] },
                format: { textFormat: { foregroundColor: { red: 0.1, green: 0.5, blue: 0.2 } } },
              },
            }, index: 1,
          },
        },
        // Goal 100% complete: green
        {
          addConditionalFormatRule: {
            rule: {
              ranges: [{ sheetId: SHEET, startRowIndex: 19, endRowIndex: 23, startColumnIndex: 0, endColumnIndex: 5 }],
              booleanRule: {
                condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$E20>=1` }] },
                format: { backgroundColor: { red: 0.88, green: 0.96, blue: 0.88 } },
              },
            }, index: 2,
          },
        },
      ],
    },
  });

  // ── Charts ────────────────────────────────────────────────────────────────
  // Chart 1: Donut — period bucket split (Needs/Wants/Future)
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      requests: [{
        addChart: {
          chart: {
            spec: {
              title: 'Period Spending by Bucket',
              pieChart: {
                legendPosition: 'RIGHT_LEGEND',
                pieHole: 0.5, // donut
                domain: { sourceRange: { sources: [{ sheetId: SHEET, startRowIndex: 46, endRowIndex: 49, startColumnIndex: 0, endColumnIndex: 1 }] } },
                series: { sourceRange: { sources: [{ sheetId: SHEET, startRowIndex: 46, endRowIndex: 49, startColumnIndex: 1, endColumnIndex: 2 }] } },
              },
            },
            position: { overlayPosition: { anchorCell: { sheetId: SHEET, rowIndex: 6, columnIndex: 9 }, widthPixels: 420, heightPixels: 280 } },
          },
        },
      }],
    },
  });

  // Chart 2: Bar chart — budget vs. actual (Needs/Wants/Future)
  // Use a small helper table on the sheet for clean source data
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [{
        range: `'Budget Dashboard'!I46:K49`,
        values: [
          ['Bucket', 'Budget', 'Actual'],
          ['Needs',  `=C13`, `=D13`],
          ['Wants',  `=C14`, `=D14`],
          ['Future', `=C15`, `=D15`],
        ],
      }],
    },
  });

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      requests: [{
        addChart: {
          chart: {
            spec: {
              title: 'Budget vs. Actual Spending',
              basicChart: {
                chartType: 'BAR',
                legendPosition: 'BOTTOM_LEGEND',
                axis: [
                  { position: 'BOTTOM_AXIS', title: 'Amount ($)' },
                  { position: 'LEFT_AXIS', title: 'Bucket' },
                ],
                domains: [{
                  domain: { sourceRange: { sources: [{ sheetId: SHEET, startRowIndex: 47, endRowIndex: 50, startColumnIndex: 8, endColumnIndex: 9 }] } },
                }],
                series: [
                  { series: { sourceRange: { sources: [{ sheetId: SHEET, startRowIndex: 47, endRowIndex: 50, startColumnIndex: 9, endColumnIndex: 10 }] } }, targetAxis: 'BOTTOM_AXIS', color: { red:0.651, green:0.725, blue:0.631 } },
                  { series: { sourceRange: { sources: [{ sheetId: SHEET, startRowIndex: 47, endRowIndex: 50, startColumnIndex: 10, endColumnIndex: 11 }] } }, targetAxis: 'BOTTOM_AXIS', color: { red:0.953, green:0.894, blue:0.784 } },
                ],
                headerCount: 1,
              },
            },
            position: { overlayPosition: { anchorCell: { sheetId: SHEET, rowIndex: 20, columnIndex: 9 }, widthPixels: 420, heightPixels: 280 } },
          },
        },
      }],
    },
  });

  console.log('✅ Budget Dashboard populated:');
  console.log('   Period selector: B5 (user date) → D5/F5 (period start/end)');
  console.log('   Stat cards: Income / Expenses / Net / Savings (SUMPRODUCT period filter)');
  console.log('   50/30/20 progress: Budget vs. Spent vs. Remaining + % Used');
  console.log('   Savings goals snapshot: 4 goals from Bills & Goals tab');
  console.log('   Upcoming bills: 10 bills from Bills & Goals registry');
  console.log('   Paycheck allocation: period income split by bucket');
  console.log('   Charts: 1 donut (bucket split) + 1 bar (budget vs. actual)');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
