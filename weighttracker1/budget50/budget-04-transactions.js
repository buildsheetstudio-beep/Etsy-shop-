'use strict';
/**
 * Step 4: Populate Transactions tab (sheetId 2).
 *
 * Column layout:
 *   A=Txn ID, B=Date, C=Description, D=Type, E=Category
 *   F=Bucket (formula), G=Amount (signed), H=Account, I=Pay Method, J=Notes
 *   K=(spacer), L=Period Start (formula), M=Period Label (formula)
 *   N=Month (formula), O=Year (formula)
 *   P=Needs Amt (formula), Q=Wants Amt (formula)
 *   R=Future Amt (formula), S=Income Amt (formula)
 *   T=Dup Check (formula)
 *
 * Data validation uses 'Start Here'!B32:B54 for category dropdown.
 * F (Bucket) uses VLOOKUP into 'Start Here'!$B$32:$C$54.
 * Period formula uses 'Start Here'!$B$11 (cadence) and $B$13 (anchor).
 */
const { google } = require('googleapis');
const fs = require('fs'), path = require('path');
const AUTH_DIR = path.join(__dirname, '..');
const CFG = JSON.parse(fs.readFileSync(path.join(__dirname, 'budget-config.json')));
const SID = CFG.spreadsheetId;
const SHEET = CFG.sheets['Transactions']; // sheetId 2

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
  negRed:     { red:0.8, green:0.2, blue:0.2 },
  posGreen:   { red:0.2, green:0.6, blue:0.3 },
};

// ── 68 sample transactions ────────────────────────────────────────────────────
// Format: [ID, Date, Description, Type, Category, Amount, Account, PayMethod, Notes]
// Signed amounts: Income=positive, Expense/SavingsTransfer=negative
const TRANSACTIONS = [
  // DEC 2025 (8)
  ['TXN-001','12/1/2025', 'Monthly Rent - Riverside Apts',     'Expense',          'Housing',              -1450.00, 'Checking Account','Bank',         ''],
  ['TXN-002','12/8/2025', 'StreamFlow Monthly',                'Expense',          'Streaming',             -15.99,  'Checking Account','Credit Card',  ''],
  ['TXN-003','12/10/2025','FreshMart Grocery',                 'Expense',          'Groceries',             -87.50,  'Checking Account','Bank',         ''],
  ['TXN-004','12/12/2025','City Electric Co',                  'Expense',          'Utilities',            -118.42,  'Checking Account','Bank',         ''],
  ['TXN-005','12/17/2025','FreshMart Grocery',                 'Expense',          'Groceries',             -94.20,  'Checking Account','Bank',         ''],
  ['TXN-006','12/22/2025','Employer Direct Deposit',           'Income',           'Paycheck',            +2400.00,  'Checking Account','Bank',         'Biweekly paycheck'],
  ['TXN-007','12/24/2025','Holiday Gifts - Amazon',            'Expense',          'Shopping',             -145.00,  'Checking Account','Credit Card',  ''],
  ['TXN-008','12/28/2025','Emergency Fund Transfer',           'Savings Transfer', 'Emergency Fund',       -200.00,  'Savings Account', 'Bank',         'Monthly contribution'],
  // JAN 2026 (22)
  ['TXN-009','1/1/2026',  'Monthly Rent - Riverside Apts',     'Expense',          'Housing',             -1450.00,  'Checking Account','Bank',         ''],
  ['TXN-010','1/5/2026',  'Employer Direct Deposit',           'Income',           'Paycheck',            +2400.00,  'Checking Account','Bank',         'Biweekly paycheck'],
  ['TXN-011','1/5/2026',  'Transit Pass - MetroCard',          'Expense',          'Transportation',        -98.00,  'Checking Account','Bank',         'Monthly transit pass'],
  ['TXN-012','1/6/2026',  'FreshMart Grocery Run',             'Expense',          'Groceries',             -91.35,  'Checking Account','Bank',         ''],
  ['TXN-013','1/8/2026',  'StreamFlow Monthly',                'Expense',          'Streaming',             -15.99,  'Checking Account','Credit Card',  ''],
  ['TXN-014','1/9/2026',  'FitLife Gym Membership',            'Expense',          'Hobbies',               -45.00,  'Checking Account','Credit Card',  ''],
  ['TXN-015','1/10/2026', 'Morning Brew Coffee',               'Expense',          'Coffee',                 -7.20,  'Checking Account','Credit Card',  ''],
  ['TXN-016','1/12/2026', 'City Electric Co',                  'Expense',          'Utilities',            -105.60,  'Checking Account','Bank',         ''],
  ['TXN-017','1/13/2026', 'FreshMart Grocery',                 'Expense',          'Groceries',             -78.45,  'Checking Account','Bank',         ''],
  ['TXN-018','1/14/2026', 'SafeGuard Auto Insurance',          'Expense',          'Insurance',             -95.00,  'Checking Account','Bank',         'Monthly premium'],
  ['TXN-019','1/15/2026', 'SoundWave Music Streaming',         'Expense',          'Streaming',             -10.99,  'Checking Account','Credit Card',  ''],
  ['TXN-020','1/17/2026', 'Thai Garden Restaurant',            'Expense',          'Dining Out',            -42.60,  'Checking Account','Credit Card',  ''],
  ['TXN-021','1/19/2026', 'Employer Direct Deposit',           'Income',           'Paycheck',            +2400.00,  'Checking Account','Bank',         'Biweekly paycheck'],
  ['TXN-022','1/20/2026', 'FreshMart Grocery',                 'Expense',          'Groceries',            -105.80,  'Checking Account','Bank',         ''],
  ['TXN-023','1/21/2026', 'Cineplex Movies x2',                'Expense',          'Entertainment',         -28.00,  'Checking Account','Credit Card',  ''],
  ['TXN-024','1/23/2026', 'Morning Brew Coffee',               'Expense',          'Coffee',                 -6.50,  'Checking Account','Credit Card',  ''],
  ['TXN-025','1/25/2026', 'Emergency Fund Transfer',           'Savings Transfer', 'Emergency Fund',       -200.00,  'Savings Account', 'Bank',         'Monthly contribution'],
  ['TXN-026','1/26/2026', 'Vacation Fund Deposit',             'Savings Transfer', 'Vacation Fund',        -100.00,  'Savings Account', 'Bank',         'Saving for summer trip'],
  ['TXN-027','1/27/2026', 'DigitSpace Cloud Storage',          'Expense',          'Streaming',             -12.99,  'Checking Account','Credit Card',  ''],
  ['TXN-028','1/28/2026', 'Amazon Prime Order',                'Expense',          'Shopping',              -54.99,  'Checking Account','Credit Card',  ''],
  ['TXN-029','1/30/2026', 'Student Loan Extra Payment',        'Expense',          'Extra Debt Payment',   -250.00,  'Checking Account','Bank',         'Above-minimum payment'],
  ['TXN-030','1/31/2026', 'Internet & Phone - ConnectUS',      'Expense',          'Utilities',             -85.00,  'Checking Account','Bank',         ''],
  // FEB 2026 (18)
  ['TXN-031','2/1/2026',  'Monthly Rent - Riverside Apts',     'Expense',          'Housing',             -1450.00,  'Checking Account','Bank',         ''],
  ['TXN-032','2/2/2026',  'Employer Direct Deposit',           'Income',           'Paycheck',            +2400.00,  'Checking Account','Bank',         'Biweekly paycheck'],
  ['TXN-033','2/3/2026',  'FreshMart Grocery',                 'Expense',          'Groceries',             -89.95,  'Checking Account','Bank',         ''],
  ['TXN-034','2/5/2026',  'FitLife Gym Membership',            'Expense',          'Hobbies',               -45.00,  'Checking Account','Credit Card',  ''],
  ['TXN-035','2/7/2026',  'StreamFlow Monthly',                'Expense',          'Streaming',             -15.99,  'Checking Account','Credit Card',  ''],
  ['TXN-036','2/10/2026', 'FreshMart Grocery',                 'Expense',          'Groceries',             -96.20,  'Checking Account','Bank',         ''],
  ['TXN-037','2/11/2026', 'City Electric Co',                  'Expense',          'Utilities',             -99.80,  'Checking Account','Bank',         ''],
  ['TXN-038','2/13/2026', 'SafeGuard Auto Insurance',          'Expense',          'Insurance',             -95.00,  'Checking Account','Bank',         'Monthly premium'],
  ['TXN-039','2/14/2026', "Valentine's Dinner - The Rosette",  'Expense',          'Dining Out',            -89.50,  'Checking Account','Credit Card',  'Special occasion'],
  ['TXN-040','2/16/2026', 'Employer Direct Deposit',           'Income',           'Paycheck',            +2400.00,  'Checking Account','Bank',         'Biweekly paycheck'],
  ['TXN-041','2/17/2026', 'FreshMart Grocery',                 'Expense',          'Groceries',             -82.10,  'Checking Account','Bank',         ''],
  ['TXN-042','2/19/2026', 'SoundWave Music Streaming',         'Expense',          'Streaming',             -10.99,  'Checking Account','Credit Card',  ''],
  ['TXN-043','2/20/2026', 'Uber Ride',                         'Expense',          'Transportation',        -22.00,  'Checking Account','Credit Card',  ''],
  ['TXN-044','2/22/2026', 'Clothing Haul - Zara',              'Expense',          'Shopping',             -120.00,  'Checking Account','Credit Card',  ''],
  ['TXN-045','2/24/2026', 'Emergency Fund Transfer',           'Savings Transfer', 'Emergency Fund',       -200.00,  'Savings Account', 'Bank',         'Monthly contribution'],
  ['TXN-046','2/25/2026', 'DigitSpace Cloud Storage',          'Expense',          'Streaming',             -12.99,  'Checking Account','Credit Card',  ''],
  ['TXN-047','2/26/2026', 'Student Loan Extra Payment',        'Expense',          'Extra Debt Payment',   -250.00,  'Checking Account','Bank',         'Above-minimum payment'],
  ['TXN-048','2/28/2026', 'Morning Brew Coffee',               'Expense',          'Coffee',                 -8.40,  'Checking Account','Credit Card',  ''],
  // MAR 2026 (14)
  ['TXN-049','3/1/2026',  'Monthly Rent - Riverside Apts',     'Expense',          'Housing',             -1450.00,  'Checking Account','Bank',         ''],
  ['TXN-050','3/2/2026',  'Employer Direct Deposit',           'Income',           'Paycheck',            +2400.00,  'Checking Account','Bank',         'Biweekly paycheck'],
  ['TXN-051','3/3/2026',  'FreshMart Grocery',                 'Expense',          'Groceries',             -93.45,  'Checking Account','Bank',         ''],
  ['TXN-052','3/5/2026',  'Transit Pass - MetroCard',          'Expense',          'Transportation',        -98.00,  'Checking Account','Bank',         'Monthly transit pass'],
  ['TXN-053','3/6/2026',  'FitLife Gym Membership',            'Expense',          'Hobbies',               -45.00,  'Checking Account','Credit Card',  ''],
  ['TXN-054','3/8/2026',  'StreamFlow Monthly',                'Expense',          'Streaming',             -15.99,  'Checking Account','Credit Card',  ''],
  ['TXN-055','3/10/2026', 'FreshMart Grocery',                 'Expense',          'Groceries',             -88.75,  'Checking Account','Bank',         ''],
  ['TXN-056','3/12/2026', 'City Electric Co',                  'Expense',          'Utilities',             -94.30,  'Checking Account','Bank',         ''],
  ['TXN-057','3/13/2026', 'SafeGuard Auto Insurance',          'Expense',          'Insurance',             -95.00,  'Checking Account','Bank',         'Monthly premium'],
  ['TXN-058','3/15/2026', 'PetGuard Quarterly Insurance',      'Expense',          'Insurance',            -180.00,  'Checking Account','Bank',         'Quarterly premium due'],
  ['TXN-059','3/16/2026', 'Employer Direct Deposit',           'Income',           'Paycheck',            +2400.00,  'Checking Account','Bank',         'Biweekly paycheck'],
  ['TXN-060','3/22/2026', 'Emergency Fund Transfer',           'Savings Transfer', 'Emergency Fund',       -250.00,  'Savings Account', 'Bank',         'Increased contribution'],
  ['TXN-061','3/24/2026', 'Vacation Fund Deposit',             'Savings Transfer', 'Vacation Fund',        -150.00,  'Savings Account', 'Bank',         'Summer trip saving'],
  ['TXN-062','3/26/2026', 'Student Loan Extra Payment',        'Expense',          'Extra Debt Payment',   -250.00,  'Checking Account','Bank',         'Above-minimum payment'],
  // APR 2026 (6)
  ['TXN-063','4/1/2026',  'Monthly Rent - Riverside Apts',     'Expense',          'Housing',             -1450.00,  'Checking Account','Bank',         ''],
  ['TXN-064','4/3/2026',  'FreshMart Grocery',                 'Expense',          'Groceries',             -96.50,  'Checking Account','Bank',         ''],
  ['TXN-065','4/5/2026',  'FitLife Gym Membership',            'Expense',          'Hobbies',               -45.00,  'Checking Account','Credit Card',  ''],
  ['TXN-066','4/7/2026',  'StreamFlow Monthly',                'Expense',          'Streaming',             -15.99,  'Checking Account','Credit Card',  ''],
  ['TXN-067','4/10/2026', 'FreshMart Grocery',                 'Expense',          'Groceries',             -83.20,  'Checking Account','Bank',         ''],
  ['TXN-068','4/13/2026', 'Employer Direct Deposit',           'Income',           'Paycheck',            +2400.00,  'Checking Account','Bank',         'Biweekly paycheck'],
];

(async () => {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  // ── Formulas for row 10 (to be copied down) ───────────────────────────────
  // F10: Bucket VLOOKUP from category dictionary
  const F10 = `=IFERROR(VLOOKUP(E10,'Start Here'!$B$32:$C$54,2,0),"Uncategorized")`;

  // L10: Period Start (date of period start, based on cadence)
  const L10 = `=IF($B10="","",IF('Start Here'!$B$11="Weekly",$B10-WEEKDAY($B10,2)+1,IF('Start Here'!$B$11="Biweekly",'Start Here'!$B$13+14*INT(($B10-'Start Here'!$B$13)/14),DATE(YEAR($B10),MONTH($B10),1))))`;

  // M10: Period Label (human-readable)
  const M10 = `=IF($L10="","",IF('Start Here'!$B$11="Monthly",TEXT($L10,"mmmm yyyy"),TEXT($L10,"mmm d")&" – "&TEXT($L10+IF('Start Here'!$B$11="Weekly",6,13),"mmm d, yyyy")))`;

  // N10: Month number
  const N10 = `=IF($B10="","",MONTH($B10))`;

  // O10: Year number
  const O10 = `=IF($B10="","",YEAR($B10))`;

  // P10: Needs Amount (negative = spending; positive means refund to needs)
  const P10 = `=IF(IFERROR($F10,"")="Needs",$G10,0)`;

  // Q10: Wants Amount
  const Q10 = `=IF(IFERROR($F10,"")="Wants",$G10,0)`;

  // R10: Future Amount (negative = transfer to savings/debt)
  const R10 = `=IF(IFERROR($F10,"")="Future",$G10,0)`;

  // S10: Income Amount (positive = received)
  const S10 = `=IF(IFERROR($F10,"")="Income",$G10,0)`;

  // T10: Duplicate ID check
  const T10 = `=IFERROR(IF(COUNTIF($A$10:$A10,$A10)>1,"⚠ Dup",""),"")`;

  // ── Summary section formulas (rows 3-7) ───────────────────────────────────
  // Using SUMPRODUCT — NEVER COUNTIFS with YEAR/MONTH (per CLAUDE.md quality gate)
  const yearRef = `'Start Here'!$B$10`; // Budget Year
  const annIncome  = `SUMPRODUCT((O10:O1009=${yearRef})*(S10:S1009))`;
  const annNeeds   = `SUMPRODUCT((O10:O1009=${yearRef})*(P10:P1009))`;
  const annWants   = `SUMPRODUCT((O10:O1009=${yearRef})*(Q10:Q1009))`;
  const annFuture  = `SUMPRODUCT((O10:O1009=${yearRef})*(R10:R1009))`;
  const annExpense = `SUMPRODUCT((O10:O1009=${yearRef})*((P10:P1009)+(Q10:Q1009)))`;
  const annNet     = `${annIncome}+${annExpense}`; // expense is negative so this = income-|expense|

  // ── 1. Write header + summary rows ────────────────────────────────────────
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [
        // Title row 1
        { range: `'Transactions'!A1`, values: [['TRANSACTIONS LEDGER']] },
        // Summary headers row 3
        { range: `'Transactions'!A3:G3`, values: [['Year','Income','Expenses (Needs+Wants)','Needs','Wants','Future (Savings)','Net Cash Flow']] },
        // Summary values row 4
        { range: `'Transactions'!A4:G4`, values: [[
          `=${yearRef}`,
          `=${annIncome}`,
          `=${annExpense}`,
          `=${annNeeds}`,
          `=${annWants}`,
          `=${annFuture}`,
          `=${annNet}`,
        ]]},
        // Info row 6
        { range: `'Transactions'!A6`, values: [['Log every income, expense, and savings transfer below. Category and Bucket are auto-filled.']] },
        // Column headers row 9
        { range: `'Transactions'!A9:T9`, values: [[
          'Txn ID','Date','Description','Type','Category','Bucket','Amount (Signed)',
          'Account','Pay Method','Notes','',
          'Period Start','Period','Month','Year',
          'Needs $','Wants $','Future $','Income $','Check',
        ]]},
      ],
    },
  });

  // ── 2. Write formula row 10 ────────────────────────────────────────────────
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [
        { range: `'Transactions'!F10`, values: [[F10]] },
        { range: `'Transactions'!L10:T10`, values: [[L10, M10, N10, O10, P10, Q10, R10, S10, T10]] },
      ],
    },
  });

  // ── 3. Copy formulas down (row 10 → rows 11-1009) ─────────────────────────
  const copyRequests = [
    // F (Bucket)
    { copyPaste: { source: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 10, startColumnIndex: 5, endColumnIndex: 6 }, destination: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 1009, startColumnIndex: 5, endColumnIndex: 6 }, pasteType: 'PASTE_FORMULA', pasteOrientation: 'NORMAL' } },
    // L-T (Period Start through Check)
    { copyPaste: { source: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 10, startColumnIndex: 11, endColumnIndex: 20 }, destination: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 1009, startColumnIndex: 11, endColumnIndex: 20 }, pasteType: 'PASTE_FORMULA', pasteOrientation: 'NORMAL' } },
  ];

  await sheets.spreadsheets.batchUpdate({ spreadsheetId: SID, requestBody: { requests: copyRequests } });

  // ── 4. Write 68 sample transactions ───────────────────────────────────────
  // Columns: A=ID, B=Date, C=Desc, D=Type, E=Category, G=Amount, H=Account, I=PayMethod, J=Notes
  const txnRows = TRANSACTIONS.map(t => [
    t[0], t[1], t[2], t[3], t[4], '', t[5], t[6], t[7], t[8],
  ]);

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [{ range: `'Transactions'!A10:J${9 + TRANSACTIONS.length}`, values: txnRows }],
    },
  });

  // ── 5. Formatting ──────────────────────────────────────────────────────────
  const fmtRequests = [];

  // Full sheet background = cream
  fmtRequests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 0, endRowIndex: 1050, startColumnIndex: 0, endColumnIndex: 20 },
      cell: { userEnteredFormat: { backgroundColor: C.cream } },
      fields: 'userEnteredFormat.backgroundColor',
    },
  });

  // Title row 1 (index 0)
  fmtRequests.push(
    { mergeCells: { range: { sheetId: SHEET, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 10 }, mergeType: 'MERGE_ALL' } },
    { repeatCell: {
        range: { sheetId: SHEET, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 10 },
        cell: { userEnteredFormat: { backgroundColor: C.deepForest, textFormat: { bold: true, foregroundColor: C.white, fontSize: 14 }, padding: { top: 10, bottom: 10, left: 12 }, verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,padding,verticalAlignment)',
      },
    }
  );

  // Summary section header row 3 (index 2)
  fmtRequests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 7 },
      cell: { userEnteredFormat: { backgroundColor: C.softSage, textFormat: { bold: true, foregroundColor: C.white }, padding: { top: 4, bottom: 4, left: 8 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
    },
  });

  // Summary values row 4 (index 3)
  fmtRequests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 3, endRowIndex: 4, startColumnIndex: 0, endColumnIndex: 7 },
      cell: { userEnteredFormat: { backgroundColor: C.paleSage, textFormat: { bold: true }, numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00;[RED]"-$"#,##0.00' }, padding: { top: 3, bottom: 3, left: 8 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat,padding)',
    },
  });
  // Year column in summary row 4: no currency format
  fmtRequests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 3, endRowIndex: 4, startColumnIndex: 0, endColumnIndex: 1 },
      cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // Column headers row 9 (index 8)
  fmtRequests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 8, endRowIndex: 9, startColumnIndex: 0, endColumnIndex: 20 },
      cell: { userEnteredFormat: { backgroundColor: C.deepForest, textFormat: { bold: true, foregroundColor: C.white, fontSize: 10 }, padding: { top: 4, bottom: 4, left: 6 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
    },
  });

  // Data rows 10-1009 (indices 9-1008): alternate rows
  for (let r = 9; r < 77; r++) { // format first 68 rows + buffer
    const bg = r % 2 === 0 ? C.cream : C.paleSage;
    fmtRequests.push({
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: 20 },
        cell: { userEnteredFormat: { backgroundColor: bg, padding: { top: 3, bottom: 3, left: 6 } } },
        fields: 'userEnteredFormat(backgroundColor,padding)',
      },
    });
  }

  // Amount column G (index 6): currency format for all data rows
  fmtRequests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 1009, startColumnIndex: 6, endColumnIndex: 7 },
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00;[RED]"-$"#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // Date column B: date format
  fmtRequests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 1009, startColumnIndex: 1, endColumnIndex: 2 },
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // Period Start (L) column: date format
  fmtRequests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 1009, startColumnIndex: 11, endColumnIndex: 12 },
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // P/Q/R/S columns (15-18): currency format
  fmtRequests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 1009, startColumnIndex: 15, endColumnIndex: 19 },
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00;[RED]"-$"#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // Freeze rows 1 + 9 (freeze first 9 rows, which includes title + headers)
  fmtRequests.push({
    updateSheetProperties: {
      properties: { sheetId: SHEET, gridProperties: { frozenRowCount: 9 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  // Column widths
  const colWidths = [90, 100, 220, 120, 150, 110, 110, 140, 110, 160, 20, 110, 140, 60, 60, 90, 90, 90, 90, 70];
  colWidths.forEach((w, i) => {
    fmtRequests.push({
      updateDimensionProperties: {
        range: { sheetId: SHEET, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
        properties: { pixelSize: w }, fields: 'pixelSize',
      },
    });
  });

  await sheets.spreadsheets.batchUpdate({ spreadsheetId: SID, requestBody: { requests: fmtRequests } });

  // ── 6. Conditional formatting ─────────────────────────────────────────────
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      requests: [
        // Positive amounts (col G) = green text
        {
          addConditionalFormatRule: {
            rule: {
              ranges: [{ sheetId: SHEET, startRowIndex: 9, endRowIndex: 1009, startColumnIndex: 6, endColumnIndex: 7 }],
              booleanRule: {
                condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0' }] },
                format: { textFormat: { foregroundColor: { red: 0.15, green: 0.55, blue: 0.25 }, bold: true } },
              },
            }, index: 0,
          },
        },
        // Savings Transfer rows: amber background on col A-J
        {
          addConditionalFormatRule: {
            rule: {
              ranges: [{ sheetId: SHEET, startRowIndex: 9, endRowIndex: 1009, startColumnIndex: 0, endColumnIndex: 10 }],
              booleanRule: {
                condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$D10="Savings Transfer"` }] },
                format: { backgroundColor: C.amber },
              },
            }, index: 1,
          },
        },
        // Income rows: pale green on A-J
        {
          addConditionalFormatRule: {
            rule: {
              ranges: [{ sheetId: SHEET, startRowIndex: 9, endRowIndex: 1009, startColumnIndex: 0, endColumnIndex: 10 }],
              booleanRule: {
                condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$D10="Income"` }] },
                format: { backgroundColor: { red: 0.88, green: 0.96, blue: 0.88 } },
              },
            }, index: 2,
          },
        },
        // Dup check T column: orange bg on warning
        {
          addConditionalFormatRule: {
            rule: {
              ranges: [{ sheetId: SHEET, startRowIndex: 9, endRowIndex: 1009, startColumnIndex: 19, endColumnIndex: 20 }],
              booleanRule: {
                condition: { type: 'TEXT_CONTAINS', values: [{ userEnteredValue: 'Dup' }] },
                format: { backgroundColor: { red: 1, green: 0.7, blue: 0.5 } },
              },
            }, index: 3,
          },
        },
      ],
    },
  });

  // ── 7. Data validation ────────────────────────────────────────────────────
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      requests: [
        // D: Transaction Type dropdown
        {
          setDataValidation: {
            range: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 1009, startColumnIndex: 3, endColumnIndex: 4 },
            rule: {
              condition: { type: 'ONE_OF_LIST', values: [
                { userEnteredValue: 'Income' },
                { userEnteredValue: 'Expense' },
                { userEnteredValue: 'Savings Transfer' },
                { userEnteredValue: 'Transfer In' },
                { userEnteredValue: 'Transfer Out' },
                { userEnteredValue: 'Refund' },
              ]},
              strict: true, showCustomUi: true,
            },
          },
        },
        // E: Category dropdown (from Start Here category list)
        {
          setDataValidation: {
            range: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 1009, startColumnIndex: 4, endColumnIndex: 5 },
            rule: {
              condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Start Here'!$B$32:$B$54` }] },
              strict: false, showCustomUi: true,
            },
          },
        },
        // I: Payment Method dropdown
        {
          setDataValidation: {
            range: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 1009, startColumnIndex: 8, endColumnIndex: 9 },
            rule: {
              condition: { type: 'ONE_OF_LIST', values: [
                { userEnteredValue: 'Bank' },
                { userEnteredValue: 'Credit Card' },
                { userEnteredValue: 'Cash' },
                { userEnteredValue: 'Other' },
              ]},
              strict: false, showCustomUi: true,
            },
          },
        },
      ],
    },
  });

  console.log(`✅ Transactions tab populated:`);
  console.log(`   ${TRANSACTIONS.length} sample transactions (Dec 2025 – Apr 2026)`);
  console.log(`   Formula columns: F (Bucket VLOOKUP), L-T (period/bucket/check)`);
  console.log(`   Data validation: Type (D), Category (E), Pay Method (I)`);
  console.log(`   Conditional formatting: income=green, savings=amber, dup=orange`);
  console.log(`   Summary: annual totals rows 3-4 (SUMPRODUCT — no COUNTIFS with YEAR/MONTH)`);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
