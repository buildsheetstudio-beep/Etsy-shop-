'use strict';
/**
 * Step 3: Populate "Start Here" tab (sheetId 1).
 *
 * Column layout: A=label, B=value (referenced by other tabs), C=notes
 *
 * Key cell addresses other tabs reference:
 *   B8  = Currency symbol ($)
 *   B9  = Budget Year (2026)
 *   B10 = Cadence (Monthly/Weekly/Biweekly)
 *   B12 = Biweekly Anchor Date (Jan 5 2026)
 *   B13 = Needs % (50)
 *   B14 = Wants % (30)
 *   B15 = Future % (20)
 *   B16 = Monthly take-home pay
 *   B17 = Biweekly take-home pay
 *   B18 = Weekly take-home pay
 *   B19 = Demo Mode (TRUE/FALSE)
 *   B28:B50 = Category names (for Transactions dropdown)
 *   B28:C50 = Category+Bucket (for Transactions VLOOKUP)
 */
const { google } = require('googleapis');
const fs = require('fs'), path = require('path');
const AUTH_DIR = path.join(__dirname, '..');
const CFG = JSON.parse(fs.readFileSync(path.join(__dirname, 'budget-config.json')));
const SID = CFG.spreadsheetId;
const SHEET = CFG.sheets['Start Here'];

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(AUTH_DIR, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const auth = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  auth.setCredentials(JSON.parse(fs.readFileSync(path.join(AUTH_DIR, 'tokens.json'))));
  return auth;
}

// ── Colors ────────────────────────────────────────────────────────────────────
const C = {
  cream:      { red:0.965, green:0.953, blue:0.922 },
  paleSage:   { red:0.906, green:0.929, blue:0.898 },
  softSage:   { red:0.651, green:0.725, blue:0.631 },
  moss:       { red:0.443, green:0.545, blue:0.475 },
  deepForest: { red:0.208, green:0.282, blue:0.243 },
  sand:       { red:0.851, green:0.788, blue:0.682 },
  amber:      { red:0.953, green:0.894, blue:0.784 },
  white:      { red:1,     green:1,     blue:1      },
};

(async () => {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  // ── 1. Write all values ────────────────────────────────────────────────────
  const data = [
    // Row 1: Main title
    { range: `'Start Here'!A1:E1`, values: [['Ultimate 50/30/20 Budget Planner','','','','']] },

    // Row 2: Tagline
    { range: `'Start Here'!A2`, values: [['Track every dollar • Balance every bucket • Hit every goal']] },

    // Row 4-5: Intro
    { range: `'Start Here'!A4`, values: [['Welcome! Configure your settings below, then head to the Budget Dashboard to see your plan come to life.']] },
    { range: `'Start Here'!A5`, values: [['Tip: The only tab you edit directly is Transactions. All other tabs update automatically.']] },

    // Row 7: Settings section header
    { range: `'Start Here'!A7`, values: [['⚙  PLANNER SETTINGS']] },

    // Row 8: Column headers for settings
    { range: `'Start Here'!A8:C8`, values: [['Setting','Value','Notes']] },

    // Rows 9-21: Settings data (label in A, value in B, note in C)
    { range: `'Start Here'!A9:C21`, values: [
      ['Currency Symbol',         '$',                'Change to £, €, ¥, etc.'],
      ['Budget Year',             2026,               '4-digit year for annual rollups'],
      ['Cadence',                 'Monthly',          'Weekly / Biweekly / Monthly'],
      ['Weekly Start Day',        'Monday',           'Week always starts on Monday'],
      ['Biweekly Anchor Date',    '1/5/2026',         'First payday — used to anchor biweekly periods'],
      ['Needs %',                 50,                 'Essentials: housing, food, utilities (default 50)'],
      ['Wants %',                 30,                 'Discretionary: dining, fun, shopping (default 30)'],
      ['Future %',                20,                 'Savings & debt paydown (default 20)'],
      ['Monthly Take-Home Pay',   5200,               'Net income per month (after tax)'],
      ['Biweekly Take-Home Pay',  2400,               'Net pay per paycheck (biweekly cadence)'],
      ['Weekly Take-Home Pay',    1200,               'Net pay per week (weekly cadence)'],
      ['Demo Mode',               'TRUE',             'Set to FALSE to clear sample/demo data'],
      ['Primary Account',         'Checking Account', 'Your main spending account name'],
    ]},

    // Row 23: Section divider
    { range: `'Start Here'!A23`, values: [['📊  50/30/20 ALLOCATION PREVIEW (based on Monthly Take-Home)']] },

    // Row 24: Allocation headers
    { range: `'Start Here'!A24:E24`, values: [['Bucket','Target %','Monthly Budget','Biweekly Budget','Weekly Budget']] },

    // Row 25-27: Allocation rows
    { range: `'Start Here'!A25:E27`, values: [
      ['=IF(B13="","Needs",B13)',
       '=B14&"%"',
       '=TEXT(B17*B14/100,"[$$-en-US]#,##0.00")',
       '=TEXT(B18*B14/100,"[$$-en-US]#,##0.00")',
       '=TEXT(B19*B14/100,"[$$-en-US]#,##0.00")'],
      ['Needs','=B14&"%"','=TEXT(B17*B14/100,"[$$-en-US]#,##0.00")','=TEXT(B18*B14/100,"[$$-en-US]#,##0.00")','=TEXT(B19*B14/100,"[$$-en-US]#,##0.00")'],
      ['Wants','=B15&"%"','=TEXT(B17*B15/100,"[$$-en-US]#,##0.00")','=TEXT(B18*B15/100,"[$$-en-US]#,##0.00")','=TEXT(B19*B15/100,"[$$-en-US]#,##0.00")'],
    ]},

    // Re-do row 25-27 properly (Needs/Wants/Future)
  ];

  // Write initial data in one call
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SID,
    requestBody: { valueInputOption: 'USER_ENTERED', data },
  });

  // Overwrite allocation rows correctly
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [{
        range: `'Start Here'!A25:E27`,
        values: [
          ['Needs', '=B14&"%"',
           '=TEXT(B17*B14/100,"[$$-en-US]#,##0.00")',
           '=TEXT(B18*B14/100,"[$$-en-US]#,##0.00")',
           '=TEXT(B19*B14/100,"[$$-en-US]#,##0.00")'],
          ['Wants', '=B15&"%"',
           '=TEXT(B17*B15/100,"[$$-en-US]#,##0.00")',
           '=TEXT(B18*B15/100,"[$$-en-US]#,##0.00")',
           '=TEXT(B19*B15/100,"[$$-en-US]#,##0.00")'],
          ['Future','=B16&"%"',
           '=TEXT(B17*B16/100,"[$$-en-US]#,##0.00")',
           '=TEXT(B18*B16/100,"[$$-en-US]#,##0.00")',
           '=TEXT(B19*B16/100,"[$$-en-US]#,##0.00")'],
        ],
      }],
    },
  });

  // Row 28: % validation note
  await sheets.spreadsheets.values.update({
    spreadsheetId: SID, range: `'Start Here'!A28`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [['=IF(B14+B15+B16=100,"✓ Percentages sum to 100%","⚠ Needs + Wants + Future must = 100 (currently "&(B14+B15+B16)&"%)")']] },
  });

  // ── Category Dictionary (rows 30-52) ───────────────────────────────────────
  // NOTE: B32:B54 and B32:C54 are referenced by Transactions tab for validation/lookup
  const catHeader = { range: `'Start Here'!A30`, values: [['📋  CATEGORY DICTIONARY']] };
  const catColHdr = { range: `'Start Here'!A31:D31`, values: [['ID','Category','Bucket','Description']] };

  const categories = [
    // Needs (8)
    ['CAT-01','Housing',               'Needs',       'Monthly rent or mortgage payment'],
    ['CAT-02','Utilities',             'Needs',       'Electricity, gas, water, internet, phone'],
    ['CAT-03','Groceries',             'Needs',       'Food shopping and household supplies'],
    ['CAT-04','Transportation',        'Needs',       'Gas, transit pass, car maintenance, parking'],
    ['CAT-05','Insurance',             'Needs',       'Health, auto, renters or homeowner insurance'],
    ['CAT-06','Healthcare',            'Needs',       'Doctor visits, prescriptions, dental, vision'],
    ['CAT-07','Minimum Debt Payments', 'Needs',       'Required minimum loan and credit card payments'],
    ['CAT-08','Childcare',             'Needs',       'Daycare, school fees, after-school programs'],
    // Wants (7)
    ['CAT-09','Dining Out',            'Wants',       'Restaurants, takeout, food delivery'],
    ['CAT-10','Entertainment',         'Wants',       'Movies, concerts, games, events'],
    ['CAT-11','Shopping',              'Wants',       'Clothing, electronics, personal items'],
    ['CAT-12','Streaming',             'Wants',       'Streaming and digital subscriptions'],
    ['CAT-13','Hobbies',               'Wants',       'Gym membership, crafts, sports, interests'],
    ['CAT-14','Travel',                'Wants',       'Flights, hotels, vacation spending'],
    ['CAT-15','Coffee',                'Wants',       'Coffee shops, cafes, tea'],
    // Future (6)
    ['CAT-16','Emergency Fund',        'Future',      'Building a 3-6 month safety net'],
    ['CAT-17','Retirement Contribution','Future',     '401k, IRA, pension contributions'],
    ['CAT-18','Extra Debt Payment',    'Future',      'Above-minimum loan or credit card payments'],
    ['CAT-19','Vacation Fund',         'Future',      'Dedicated saving for planned trips'],
    ['CAT-20','Home Down Payment',     'Future',      'Saving toward buying a home'],
    ['CAT-21','Other Savings',         'Future',      'General or miscellaneous savings goals'],
    // Special (2)
    ['CAT-22','Paycheck',              'Income',      'Mark income deposits with this category'],
    ['CAT-23','Account Transfer',      'Transfer',    'Internal bank transfer — not counted as spending'],
  ];

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [
        catHeader,
        catColHdr,
        { range: `'Start Here'!A32:D54`, values: categories },
      ],
    },
  });

  // ── Onboarding Guide (rows 57-75) ──────────────────────────────────────────
  const guideRows = [
    ['🚀  QUICK START GUIDE'],
    [''],
    ['Step 1 — Set your income', 'Enter your monthly, biweekly, or weekly take-home pay in the Settings above (rows 17-19).'],
    ['Step 2 — Choose your cadence', 'Set Cadence (row 11) to Monthly, Biweekly, or Weekly. If Biweekly, set your anchor pay date in row 13.'],
    ['Step 3 — Adjust percentages', 'Needs + Wants + Future must equal 100. Edit rows 14-16 to match your financial goals.'],
    ['Step 4 — Log your transactions', 'Go to the Transactions tab. Add every income, expense, and savings transfer. Date, description, amount, and category are required.'],
    ['Step 5 — Track your bills', 'Visit Bills & Goals to register recurring bills. Mark each as Paid when done.'],
    ['Step 6 — Set savings goals', 'In Bills & Goals, add savings goals (Emergency Fund, Vacation, etc.) and track progress.'],
    ['Step 7 — Review the Dashboard', 'Budget Dashboard shows your 50/30/20 progress for the current period, plus charts and upcoming bills.'],
    ['Step 8 — Check Yearly Insights', 'At month-end, Yearly Insights shows monthly trends, category breakdowns, and your annual summary.'],
    [''],
    ['📌  TIPS', ''],
    ['• Always use consistent category names', 'Match exactly what is in the Category Dictionary above — the bucket classification is automatic.'],
    ['• Mark transfers correctly', 'Use "Account Transfer" for bank-to-bank moves. These are excluded from spending totals.'],
    ['• Paychecks are Income', 'Category = "Paycheck", amount = positive, type = "Income".'],
    ['• Savings deposits are Future', 'Category = your goal name (e.g. "Emergency Fund"), type = "Savings Transfer".'],
    ['• Turn off Demo Mode', 'Set B20 to FALSE when you\'re ready to start with a clean slate.'],
  ];

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [{ range: `'Start Here'!A57:C73`, values: guideRows }],
    },
  });

  // ── 2. Formatting ──────────────────────────────────────────────────────────
  const requests = [];

  // Background: full sheet = cream
  requests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 0, endRowIndex: 100, startColumnIndex: 0, endColumnIndex: 10 },
      cell: { userEnteredFormat: { backgroundColor: C.cream } },
      fields: 'userEnteredFormat.backgroundColor',
    },
  });

  // Row 1: Big title — deep forest bg, white bold text, 18pt, merged A1:E1
  requests.push(
    { mergeCells: { range: { sheetId: SHEET, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 5 }, mergeType: 'MERGE_ALL' } },
    {
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 5 },
        cell: { userEnteredFormat: {
          backgroundColor: C.deepForest,
          textFormat: { bold: true, foregroundColor: C.white, fontSize: 18 },
          horizontalAlignment: 'CENTER',
          verticalAlignment: 'MIDDLE',
          padding: { top: 12, bottom: 12, left: 16 },
        }},
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,padding)',
      },
    }
  );

  // Row 2: Tagline — moss bg, white italic
  requests.push(
    { mergeCells: { range: { sheetId: SHEET, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 5 }, mergeType: 'MERGE_ALL' } },
    {
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 5 },
        cell: { userEnteredFormat: {
          backgroundColor: C.moss,
          textFormat: { italic: true, foregroundColor: C.white, fontSize: 11 },
          horizontalAlignment: 'CENTER',
          verticalAlignment: 'MIDDLE',
          padding: { top: 6, bottom: 6 },
        }},
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,padding)',
      },
    }
  );

  // Section header row 7 (index 6): ⚙ SETTINGS
  requests.push(
    { mergeCells: { range: { sheetId: SHEET, startRowIndex: 6, endRowIndex: 7, startColumnIndex: 0, endColumnIndex: 5 }, mergeType: 'MERGE_ALL' } },
    {
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: 6, endRowIndex: 7, startColumnIndex: 0, endColumnIndex: 5 },
        cell: { userEnteredFormat: {
          backgroundColor: C.deepForest,
          textFormat: { bold: true, foregroundColor: C.white, fontSize: 12 },
          verticalAlignment: 'MIDDLE',
          padding: { top: 6, bottom: 6, left: 10 },
        }},
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,padding)',
      },
    }
  );

  // Settings column header row 8 (index 7)
  requests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 7, endRowIndex: 8, startColumnIndex: 0, endColumnIndex: 3 },
      cell: { userEnteredFormat: {
        backgroundColor: C.softSage,
        textFormat: { bold: true, foregroundColor: C.white },
        verticalAlignment: 'MIDDLE',
        padding: { top: 4, bottom: 4, left: 8 },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,padding)',
    },
  });

  // Settings data rows 9-21 (indices 8-20): alternating cream / paleSage
  for (let r = 8; r <= 20; r++) {
    const bg = r % 2 === 0 ? C.cream : C.paleSage;
    requests.push({
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: 3 },
        cell: { userEnteredFormat: {
          backgroundColor: bg,
          padding: { top: 3, bottom: 3, left: 8 },
        }},
        fields: 'userEnteredFormat(backgroundColor,padding)',
      },
    });
    // Bold value cells (col B)
    requests.push({
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 1, endColumnIndex: 2 },
        cell: { userEnteredFormat: { textFormat: { bold: true } } },
        fields: 'userEnteredFormat.textFormat.bold',
      },
    });
  }

  // Allocation preview header row 23 (index 22)
  requests.push(
    { mergeCells: { range: { sheetId: SHEET, startRowIndex: 22, endRowIndex: 23, startColumnIndex: 0, endColumnIndex: 5 }, mergeType: 'MERGE_ALL' } },
    {
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: 22, endRowIndex: 23, startColumnIndex: 0, endColumnIndex: 5 },
        cell: { userEnteredFormat: {
          backgroundColor: C.moss,
          textFormat: { bold: true, foregroundColor: C.white, fontSize: 11 },
          padding: { top: 5, bottom: 5, left: 10 },
        }},
        fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
      },
    }
  );

  // Allocation column headers row 24 (index 23)
  requests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 23, endRowIndex: 24, startColumnIndex: 0, endColumnIndex: 5 },
      cell: { userEnteredFormat: {
        backgroundColor: C.softSage,
        textFormat: { bold: true, foregroundColor: C.white },
        padding: { top: 4, bottom: 4, left: 8 },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
    },
  });

  // Allocation data rows 25-27 (indices 24-26)
  const allocColors = [C.paleSage, C.cream, C.amber];
  for (let r = 24; r <= 26; r++) {
    requests.push({
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: 5 },
        cell: { userEnteredFormat: {
          backgroundColor: allocColors[r-24],
          textFormat: { bold: true },
          padding: { top: 4, bottom: 4, left: 8 },
        }},
        fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
      },
    });
  }

  // % check row 28 (index 27): pale amber, italic
  requests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 27, endRowIndex: 28, startColumnIndex: 0, endColumnIndex: 5 },
      cell: { userEnteredFormat: {
        backgroundColor: C.amber,
        textFormat: { italic: true },
        padding: { top: 2, bottom: 2, left: 8 },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
    },
  });

  // Category Dictionary header row 30 (index 29)
  requests.push(
    { mergeCells: { range: { sheetId: SHEET, startRowIndex: 29, endRowIndex: 30, startColumnIndex: 0, endColumnIndex: 5 }, mergeType: 'MERGE_ALL' } },
    {
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: 29, endRowIndex: 30, startColumnIndex: 0, endColumnIndex: 5 },
        cell: { userEnteredFormat: {
          backgroundColor: C.deepForest,
          textFormat: { bold: true, foregroundColor: C.white, fontSize: 12 },
          padding: { top: 6, bottom: 6, left: 10 },
        }},
        fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
      },
    }
  );

  // Cat col headers row 31 (index 30)
  requests.push({
    repeatCell: {
      range: { sheetId: SHEET, startRowIndex: 30, endRowIndex: 31, startColumnIndex: 0, endColumnIndex: 4 },
      cell: { userEnteredFormat: {
        backgroundColor: C.softSage,
        textFormat: { bold: true, foregroundColor: C.white },
        padding: { top: 4, bottom: 4, left: 8 },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
    },
  });

  // Category data rows 32-54 (indices 31-53): alternating, bucket-color col C
  const bucketColors = { Needs: C.paleSage, Wants: C.amber, Future: { red:0.847, green:0.647, blue:0.604 }, Income: { red:0.9, green:0.95, blue:0.9 }, Transfer: C.sand };
  for (let r = 31; r <= 53; r++) {
    const bg = r % 2 === 0 ? C.cream : C.paleSage;
    requests.push({
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: r, endRowIndex: r+1, startColumnIndex: 0, endColumnIndex: 4 },
        cell: { userEnteredFormat: { backgroundColor: bg, padding: { top: 3, bottom: 3, left: 8 } } },
        fields: 'userEnteredFormat(backgroundColor,padding)',
      },
    });
  }

  // Quick Start Guide header row 57 (index 56)
  requests.push(
    { mergeCells: { range: { sheetId: SHEET, startRowIndex: 56, endRowIndex: 57, startColumnIndex: 0, endColumnIndex: 5 }, mergeType: 'MERGE_ALL' } },
    {
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: 56, endRowIndex: 57, startColumnIndex: 0, endColumnIndex: 5 },
        cell: { userEnteredFormat: {
          backgroundColor: C.deepForest,
          textFormat: { bold: true, foregroundColor: C.white, fontSize: 12 },
          padding: { top: 6, bottom: 6, left: 10 },
        }},
        fields: 'userEnteredFormat(backgroundColor,textFormat,padding)',
      },
    }
  );

  // Column widths
  const colWidths = [240, 200, 380, 200, 200]; // A B C D E
  colWidths.forEach((w, i) => {
    requests.push({
      updateDimensionProperties: {
        range: { sheetId: SHEET, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
        properties: { pixelSize: w },
        fields: 'pixelSize',
      },
    });
  });

  // Row heights
  requests.push(
    { updateDimensionProperties: { range: { sheetId: SHEET, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } },
    { updateDimensionProperties: { range: { sheetId: SHEET, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } },
  );

  // Freeze row 1 (title row)
  requests.push({
    updateSheetProperties: {
      properties: { sheetId: SHEET, gridProperties: { frozenRowCount: 1 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  // Row 4-5: Intro text — merge A:E, italic moss text
  requests.push(
    { mergeCells: { range: { sheetId: SHEET, startRowIndex: 3, endRowIndex: 4, startColumnIndex: 0, endColumnIndex: 5 }, mergeType: 'MERGE_ALL' } },
    { mergeCells: { range: { sheetId: SHEET, startRowIndex: 4, endRowIndex: 5, startColumnIndex: 0, endColumnIndex: 5 }, mergeType: 'MERGE_ALL' } },
    {
      repeatCell: {
        range: { sheetId: SHEET, startRowIndex: 3, endRowIndex: 5, startColumnIndex: 0, endColumnIndex: 5 },
        cell: { userEnteredFormat: {
          backgroundColor: C.paleSage,
          textFormat: { italic: true, foregroundColor: C.deepForest },
          padding: { top: 5, bottom: 5, left: 12 },
          wrapStrategy: 'WRAP',
        }},
        fields: 'userEnteredFormat(backgroundColor,textFormat,padding,wrapStrategy)',
      },
    }
  );

  await sheets.spreadsheets.batchUpdate({ spreadsheetId: SID, requestBody: { requests } });

  // ── 3. Data Validation ─────────────────────────────────────────────────────
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      requests: [
        // Cadence dropdown (B11, 0-indexed row 10)
        {
          setDataValidation: {
            range: { sheetId: SHEET, startRowIndex: 10, endRowIndex: 11, startColumnIndex: 1, endColumnIndex: 2 },
            rule: {
              condition: { type: 'ONE_OF_LIST', values: [
                { userEnteredValue: 'Weekly' },
                { userEnteredValue: 'Biweekly' },
                { userEnteredValue: 'Monthly' },
              ]},
              strict: true, showCustomUi: true,
            },
          },
        },
        // Demo Mode dropdown (B20, index 19)
        {
          setDataValidation: {
            range: { sheetId: SHEET, startRowIndex: 19, endRowIndex: 20, startColumnIndex: 1, endColumnIndex: 2 },
            rule: {
              condition: { type: 'ONE_OF_LIST', values: [
                { userEnteredValue: 'TRUE' },
                { userEnteredValue: 'FALSE' },
              ]},
              strict: true, showCustomUi: true,
            },
          },
        },
      ],
    },
  });

  console.log('✅ Start Here tab populated:');
  console.log('   Settings: rows 8-21 (B8-B20 = key reference cells)');
  console.log('   Allocation preview: rows 23-28');
  console.log('   Category dictionary: rows 30-54 (B32:C54 = lookup range)');
  console.log('   Quick Start Guide: rows 57-73');
  console.log('   Data validation: Cadence (B11), Demo Mode (B20)');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
