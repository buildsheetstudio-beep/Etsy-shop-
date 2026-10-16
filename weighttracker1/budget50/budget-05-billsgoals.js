'use strict';
/**
 * Step 5: Populate Bills & Goals tab (sheetId 3).
 *
 * Three sections:
 *   Section 1 — Bills Registry     rows 8-107  (col A-K)
 *   Section 2 — Payment History    rows 113-312 (col A-F)
 *   Section 3 — Savings Goals      rows 318-417 (col A-J)
 *
 * Bills: 10 sample bills
 * Payment History: 8 sample payments
 * Savings Goals: 4 goals
 */
const { google } = require('googleapis');
const fs = require('fs'), path = require('path');
const AUTH_DIR = path.join(__dirname, '..');
const CFG = JSON.parse(fs.readFileSync(path.join(__dirname, 'budget-config.json')));
const SID = CFG.spreadsheetId;
const SHEET = CFG.sheets['Bills & Goals']; // sheetId 3

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
  white:      { red:1, green:1, blue:1 },
  lightGreen: { red:0.88, green:0.96, blue:0.88 },
};

// ── Bills data ────────────────────────────────────────────────────────────────
// [ID, Name, Category, Frequency, Amount, DueDay, AutoPay, Account, Notes]
const BILLS = [
  ['BILL-01','Monthly Rent',               'Housing',        'Monthly',   1450.00,  1,  true,  'Checking Account','Riverside Apts'],
  ['BILL-02','City Electric Co',           'Utilities',      'Monthly',    120.00, 12,  false, 'Checking Account','Average ~$110-130'],
  ['BILL-03','SafeGuard Auto Insurance',   'Insurance',      'Monthly',     95.00, 13,  true,  'Checking Account','Full coverage policy'],
  ['BILL-04','StreamFlow',                 'Streaming',      'Monthly',     15.99,  8,  true,  'Checking Account','HD streaming plan'],
  ['BILL-05','Transit Pass - MetroCard',   'Transportation', 'Monthly',     98.00,  5,  false, 'Checking Account','Monthly unlimited pass'],
  ['BILL-06','DigitSpace Cloud Storage',   'Streaming',      'Monthly',     12.99, 27,  true,  'Checking Account','2TB cloud plan'],
  ['BILL-07','SoundWave Music',            'Streaming',      'Monthly',     10.99, 19,  true,  'Checking Account','Premium individual plan'],
  ['BILL-08','FitLife Gym',                'Hobbies',        'Monthly',     45.00,  5,  true,  'Checking Account','No-contract membership'],
  ['BILL-09','PetGuard Quarterly',         'Insurance',      'Quarterly',  180.00, 15,  false, 'Checking Account','Due Jan/Apr/Jul/Oct'],
  ['BILL-10','Domain & Hosting Annual',    'Streaming',      'Yearly',      89.00,  1,  true,  'Checking Account','Renews every March'],
];

// ── Payment History ───────────────────────────────────────────────────────────
// [PayID, BillID, Bill Name, Payment Date, Amount, Status, Notes]
const PAYMENTS = [
  ['PAY-001','BILL-01','Monthly Rent',          '1/1/2026',  1450.00,'Paid','January rent on time'],
  ['PAY-002','BILL-05','Transit Pass',          '1/5/2026',    98.00,'Paid','January monthly pass'],
  ['PAY-003','BILL-04','StreamFlow',            '1/8/2026',    15.99,'Paid','Auto-renewed'],
  ['PAY-004','BILL-02','City Electric Co',      '1/12/2026',  105.60,'Paid','Lower in winter'],
  ['PAY-005','BILL-03','SafeGuard Auto',        '1/14/2026',   95.00,'Paid','Auto-deducted'],
  ['PAY-006','BILL-01','Monthly Rent',          '2/1/2026',  1450.00,'Paid','February rent on time'],
  ['PAY-007','BILL-09','PetGuard Quarterly',    '1/15/2026',  180.00,'Paid','Q1 quarterly premium'],
  ['PAY-008','BILL-01','Monthly Rent',          '3/1/2026',  1450.00,'Paid','March rent on time'],
];

// ── Savings Goals ─────────────────────────────────────────────────────────────
// [GoalID, Name, Type, Target, CurrentSaved, MonthlyContrib, TargetDate, Notes]
const GOALS = [
  ['GOAL-01','Emergency Fund',     'Emergency Fund',       5000.00,  850.00, 200.00,'12/31/2027','3-6 month safety net goal'],
  ['GOAL-02','Vacation Fund',      'Sinking Fund',         2000.00,  400.00, 150.00,'6/30/2026', 'Summer road trip'],
  ['GOAL-03','Home Down Payment',  'Savings',             20000.00, 1200.00,   0.00,'12/31/2029','Long-term savings target'],
  ['GOAL-04','Student Loan Extra', 'Extra Debt Payment',   3600.00,  750.00, 250.00,'12/31/2026','Accelerated payoff plan'],
];

(async () => {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  // ── 1. Write section headers and data ─────────────────────────────────────
  // Bills section
  const billsData = BILLS.map(b => [
    b[0], b[1], b[2], b[3],
    b[4], // Amount
    b[5], // Due Day
    `=IFERROR(IF($D${10+BILLS.indexOf(b)}="Monthly",DATE(YEAR(TODAY()),MONTH(TODAY())+(DAY(TODAY())>=$F${10+BILLS.indexOf(b)}),$F${10+BILLS.indexOf(b)}),IF($D${10+BILLS.indexOf(b)}="Quarterly",EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),3),IF($D${10+BILLS.indexOf(b)}="Yearly",DATE(YEAR(TODAY())+1,MONTH(TODAY()),1),TODAY()+7))),"")`,
    b[6], // AutoPay checkbox (TRUE/FALSE)
    'Scheduled', // Status
    b[7], // Account
    b[8], // Notes
  ]);

  // Simpler: compute next due date formula inline
  const billRows = BILLS.map((b, i) => {
    const row = 10 + i; // rows 10-19
    return [
      b[0], b[1], b[2], b[3],
      b[4], // Amount (col E)
      b[5], // Due Day (col F)
      `=IFERROR(IF($D${row}="Monthly",DATE(YEAR(TODAY()),MONTH(TODAY())+(DAY(TODAY())>=$F${row}),$F${row}),IF($D${row}="Quarterly",EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),3),IF($D${row}="Yearly",DATE(YEAR(TODAY())+1,1,1),TODAY()+7))),"")`,
      b[6] ? 'TRUE' : 'FALSE', // AutoPay (col H)
      'Scheduled', // Status (col I)
      b[7], // Account (col J)
      b[8], // Notes (col K)
    ];
  });

  // Payment history rows
  const payRows = PAYMENTS.map(p => [p[0], p[1], p[2], p[3], p[4], p[5], p[6]]);

  // Goal rows
  const goalRows = GOALS.map((g, i) => {
    const row = 320 + i;
    // H = progress %, I = months to complete
    return [
      g[0], g[1], g[2], g[3], g[4], g[5], g[6],
      `=IFERROR($E${row}/$D${row},0)`,
      `=IFERROR(IF($F${row}=0,"—",CEILING(($D${row}-$E${row})/$F${row},1)),"—")`,
      g[7],
    ];
  });

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [
        // ── Bills section ──────────────────────────────────────────────────
        { range: `'Bills & Goals'!A8`, values: [['🧾  BILLS & RECURRING PAYMENTS']] },
        { range: `'Bills & Goals'!A9:K9`, values: [['Bill ID','Name','Category','Frequency','Amount','Due Day','Next Due','Auto Pay','Status','Account','Notes']] },
        { range: `'Bills & Goals'!A10:K${9+BILLS.length}`, values: billRows },

        // Monthly total row
        { range: `'Bills & Goals'!A${10+BILLS.length+1}:K${10+BILLS.length+1}`, values: [[
          'MONTHLY TOTAL','','','',
          `=SUMPRODUCT((D10:D${9+BILLS.length}="Monthly")*E10:E${9+BILLS.length})+SUMPRODUCT((D10:D${9+BILLS.length}="Quarterly")*E10:E${9+BILLS.length})/3+SUMPRODUCT((D10:D${9+BILLS.length}="Yearly")*E10:E${9+BILLS.length})/12`,
          '','','','','','Avg monthly cost',
        ]]},

        // ── Payment History section ────────────────────────────────────────
        { range: `'Bills & Goals'!A113`, values: [['📅  PAYMENT HISTORY']] },
        { range: `'Bills & Goals'!A114:G114`, values: [['Pay ID','Bill ID','Bill Name','Payment Date','Amount Paid','Status','Notes']] },
        { range: `'Bills & Goals'!A115:G${114+PAYMENTS.length}`, values: payRows },

        // ── Savings Goals section ──────────────────────────────────────────
        { range: `'Bills & Goals'!A318`, values: [['🎯  SAVINGS GOALS & SINKING FUNDS']] },
        { range: `'Bills & Goals'!A319:J319`, values: [['Goal ID','Goal Name','Type','Target','Saved So Far','Monthly Contrib','Target Date','Progress','Months to Go','Notes']] },
        { range: `'Bills & Goals'!A320:J${319+GOALS.length}`, values: goalRows },

        // Goals summary row
        { range: `'Bills & Goals'!A${320+GOALS.length}:J${320+GOALS.length}`, values: [[
          'TOTAL','','',
          `=SUM(D320:D${319+GOALS.length})`,
          `=SUM(E320:E${319+GOALS.length})`,
          `=SUM(F320:F${319+GOALS.length})`,
          '','','','',
        ]]},
      ],
    },
  });

  // ── 2. Formatting ──────────────────────────────────────────────────────────
  const requests = [];

  // Full sheet cream background
  requests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 0, endRowIndex: 450, startColumnIndex: 0, endColumnIndex: 14 },
      cell: { userEnteredFormat: { backgroundColor: C.cream } },
      fields: 'userEnteredFormat.backgroundColor',
    },
  });

  // Helper: section header style
  const sectionHeader = (r, cols = 11) => [
    { mergeCells: { range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: cols }, mergeType: 'MERGE_ALL' } },
    { repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: cols },
        cell: { userEnteredFormat: { backgroundColor: C.deepForest, textFormat: { bold: true, foregroundColor: C.white, fontSize: 13 }, padding: { top: 8, bottom: 8, left: 10 } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
      },
    },
  ];

  // Helper: column header style
  const colHeader = (r, cols) => ({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: cols },
      cell: { userEnteredFormat: { backgroundColor: C.softSage, textFormat: { bold: true, foregroundColor: C.white }, padding: { top: 4, bottom: 4, left: 6 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
    },
  });

  requests.push(
    ...sectionHeader(7), colHeader(8, 11),
    ...sectionHeader(112, 7), colHeader(113, 7),
    ...sectionHeader(317, 10), colHeader(318, 10),
  );

  // Bills data rows 10-19 alternating
  for (let r = 9; r <= 9+BILLS.length; r++) {
    const bg = r % 2 === 0 ? C.cream : C.paleSage;
    requests.push({
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: 11 },
        cell: { userEnteredFormat: { backgroundColor: bg, padding: { top: 3, bottom: 3, left: 6 } } },
        fields: 'userEnteredFormat(backgroundColor,padding)',
      },
    });
  }

  // Total row for bills
  requests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 9+BILLS.length+1, endRowIndex: 9+BILLS.length+2, startColumnIndex: 0, endColumnIndex: 11 },
      cell: { userEnteredFormat: { backgroundColor: C.moss, textFormat: { bold: true, foregroundColor: C.white }, padding: { top: 4, bottom: 4, left: 6 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
    },
  });

  // Payment history rows alternating
  for (let r = 114; r <= 114+PAYMENTS.length; r++) {
    const bg = r % 2 === 0 ? C.cream : C.paleSage;
    requests.push({
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: 7 },
        cell: { userEnteredFormat: { backgroundColor: bg, padding: { top: 3, bottom: 3, left: 6 } } },
        fields: 'userEnteredFormat(backgroundColor,padding)',
      },
    });
  }

  // Goals rows alternating
  for (let r = 319; r <= 319+GOALS.length; r++) {
    const bg = r % 2 === 0 ? C.cream : C.paleSage;
    requests.push({
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: 10 },
        cell: { userEnteredFormat: { backgroundColor: bg, padding: { top: 3, bottom: 3, left: 6 } } },
        fields: 'userEnteredFormat(backgroundColor,padding)',
      },
    });
  }

  // Goals total row
  requests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 319+GOALS.length, endRowIndex: 320+GOALS.length, startColumnIndex: 0, endColumnIndex: 10 },
      cell: { userEnteredFormat: { backgroundColor: C.moss, textFormat: { bold: true, foregroundColor: C.white }, padding: { top: 4, bottom: 4, left: 6 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
    },
  });

  // Amount/currency columns
  // Bills: E (amount, index 4), G (next due date, index 6)
  requests.push(
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 107, startColumnIndex: 4, endColumnIndex: 5 }, cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } },
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 107, startColumnIndex: 6, endColumnIndex: 7 }, cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } },
    // Payment history: E (amount, index 4), D (date, index 3)
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 114, endRowIndex: 312, startColumnIndex: 4, endColumnIndex: 5 }, cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } },
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 114, endRowIndex: 312, startColumnIndex: 3, endColumnIndex: 4 }, cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } },
    // Goals: D/E/F (indices 3/4/5), G target date (index 6)
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 319, endRowIndex: 417, startColumnIndex: 3, endColumnIndex: 6 }, cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } },
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 319, endRowIndex: 417, startColumnIndex: 6, endColumnIndex: 7 }, cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } },
    // Goals: H (progress %, index 7)
    { repeatCell: { range: { sheetId: SHEET, startRowIndex: 319, endRowIndex: 417, startColumnIndex: 7, endColumnIndex: 8 }, cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } },
  );

  // Column widths
  const widths = [90, 200, 140, 100, 100, 80, 110, 90, 100, 140, 200];
  widths.forEach((w, i) => {
    requests.push({
      updateDimensionProperties: {
        range: { sheetId: SHEET, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
        properties: { pixelSize: w }, fields: 'pixelSize',
      },
    });
  });

  // Freeze row 9 (col headers) of bills section
  requests.push({
    updateSheetProperties: {
      properties: { sheetId: SHEET, gridProperties: { frozenRowCount: 9 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await sheets.spreadsheets.batchUpdate({ spreadsheetId: SID, requestBody: { requests } });

  // ── 3. Conditional formatting ─────────────────────────────────────────────
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      requests: [
        // Bills: Paid status = green
        {
          addConditionalFormatRule: {
            rule: {
              ranges: [{ sheetId: SHEET, startRowIndex: 9, endRowIndex: 107, startColumnIndex: 0, endColumnIndex: 11 }],
              booleanRule: {
                condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$I10="Paid"` }] },
                format: { backgroundColor: C.lightGreen },
              },
            }, index: 0,
          },
        },
        // Bills: Overdue = light red
        {
          addConditionalFormatRule: {
            rule: {
              ranges: [{ sheetId: SHEET, startRowIndex: 9, endRowIndex: 107, startColumnIndex: 0, endColumnIndex: 11 }],
              booleanRule: {
                condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$I10="Overdue"` }] },
                format: { backgroundColor: { red: 0.95, green: 0.8, blue: 0.8 } },
              },
            }, index: 1,
          },
        },
        // Goals: progress >= 100% = green
        {
          addConditionalFormatRule: {
            rule: {
              ranges: [{ sheetId: SHEET, startRowIndex: 319, endRowIndex: 417, startColumnIndex: 0, endColumnIndex: 10 }],
              booleanRule: {
                condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$H320>=1` }] },
                format: { backgroundColor: C.lightGreen },
              },
            }, index: 2,
          },
        },
      ],
    },
  });

  // ── 4. Data validation ────────────────────────────────────────────────────
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      requests: [
        // Bills: Frequency (col D, index 3)
        {
          setDataValidation: {
            range: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 107, startColumnIndex: 3, endColumnIndex: 4 },
            rule: {
              condition: { type: 'ONE_OF_LIST', values: [
                { userEnteredValue: 'Weekly' }, { userEnteredValue: 'Biweekly' },
                { userEnteredValue: 'Monthly' }, { userEnteredValue: 'Quarterly' },
                { userEnteredValue: 'Yearly' },
              ]},
              strict: true, showCustomUi: true,
            },
          },
        },
        // Bills: Status (col I, index 8)
        {
          setDataValidation: {
            range: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 107, startColumnIndex: 8, endColumnIndex: 9 },
            rule: {
              condition: { type: 'ONE_OF_LIST', values: [
                { userEnteredValue: 'Scheduled' }, { userEnteredValue: 'Paid' },
                { userEnteredValue: 'Late' }, { userEnteredValue: 'Overdue' },
              ]},
              strict: true, showCustomUi: true,
            },
          },
        },
        // Bills: AutoPay checkbox (col H, index 7)
        {
          setDataValidation: {
            range: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 107, startColumnIndex: 7, endColumnIndex: 8 },
            rule: { condition: { type: 'BOOLEAN' }, strict: true, showCustomUi: true },
          },
        },
        // Payment History: Status (col F, index 5)
        {
          setDataValidation: {
            range: { sheetId: SHEET, startRowIndex: 114, endRowIndex: 312, startColumnIndex: 5, endColumnIndex: 6 },
            rule: {
              condition: { type: 'ONE_OF_LIST', values: [
                { userEnteredValue: 'Paid' }, { userEnteredValue: 'Partial' }, { userEnteredValue: 'Late' },
              ]},
              strict: true, showCustomUi: true,
            },
          },
        },
        // Goals: Type (col C, index 2)
        {
          setDataValidation: {
            range: { sheetId: SHEET, startRowIndex: 319, endRowIndex: 417, startColumnIndex: 2, endColumnIndex: 3 },
            rule: {
              condition: { type: 'ONE_OF_LIST', values: [
                { userEnteredValue: 'Emergency Fund' }, { userEnteredValue: 'Sinking Fund' },
                { userEnteredValue: 'Savings' }, { userEnteredValue: 'Extra Debt Payment' },
              ]},
              strict: false, showCustomUi: true,
            },
          },
        },
      ],
    },
  });

  console.log('✅ Bills & Goals tab populated:');
  console.log(`   ${BILLS.length} bills in registry (rows 10-${9+BILLS.length})`);
  console.log(`   ${PAYMENTS.length} payment history entries (rows 115-${114+PAYMENTS.length})`);
  console.log(`   ${GOALS.length} savings goals (rows 320-${319+GOALS.length})`);
  console.log('   Next Due Date: formula using TODAY() for accurate upcoming dates');
  console.log('   Conditional formatting: Paid=green, Overdue=red, Goal complete=green');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
