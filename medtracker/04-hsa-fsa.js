'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const HSA = sheetMap['💳 HSA/FSA Tracker'];

// A=Account Type | B=Tax Year | C=Annual Limit (formula) | D=Your Contributions YTD | E=Employer Contributions
// F=Total Contributed (formula) | G=Withdrawals/Reimbursements (formula) | H=Remaining Balance (formula)
// I=Grace Period/Rollover Note

const HEADERS = ['Account Type','Tax Year','Annual Contrib. Limit','Your Contributions YTD','Employer Contributions YTD','Total Contributed','Withdrawals/Reimbursements','Remaining Balance','Grace Period / Rollover Note'];
const WIDTHS  = [120, 80, 120, 140, 160, 120, 170, 130, 220];

const EXP_TAB = "'🧾 Expense Log'";

// Sample: 2 rows — HSA Self-Only 2024, FSA 2024
// HSA: 1800 personal + 200 employer = 2000 total
//   Withdrawals = SUMPRODUCT where year=2024 AND payment="HSA"
//   2024 HSA expenses: 250+45+180+150+55+85+175 = 940 → Balance = 1060
// FSA: 2000 personal + 0 employer = 2000 total
//   Withdrawals = SUMPRODUCT where year=2024 AND payment="FSA"
//   2024 FSA expenses: 180+320+65+450+150+950 = 2115 → Balance = -115 (RED CF demo!)

function withdrawalFormula(row) {
  // If Account Type starts with "HSA" → match "HSA" payments; else match "FSA" payments
  return `=IFERROR(SUMPRODUCT((YEAR(${EXP_TAB}!$A$3:$A$200)=B${row})*(IF(LEFT(A${row},3)="HSA",${EXP_TAB}!$F$3:$F$200="HSA",${EXP_TAB}!$F$3:$F$200="FSA"))*(${EXP_TAB}!$E$3:$E$200)),0)`;
}

(async () => {
  const reqs = [];

  // Row 0: Title
  reqs.push({ mergeCells: { range: gridRange(HSA, 0, 1, 0, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(HSA, 0, 1, 0, 9),
      cell: {
        userEnteredValue: { stringValue: '💳 HSA / FSA Account Tracker' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepTeal),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: HSA, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Row 1: Headers
  reqs.push({
    repeatCell: {
      range: gridRange(HSA, 1, 2, 0, 9),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.mutedSage),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: HSA, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});

  // Data rows 2-3
  for (let r = 2; r < 4; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(HSA, r, r+1, 0, 9),
        cell: { userEnteredFormat: { backgroundColor: hex(r === 2 ? C.inputBg : C.altRow), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: HSA, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 30 }, fields: 'pixelSize',
    }});
  }

  // Formula columns: C (limit), F (total), G (withdrawals), H (balance) → formulaBg
  for (const col of [2, 5, 6, 7]) {
    reqs.push({ repeatCell: {
      range: gridRange(HSA, 2, 4, col, col+1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg) } },
      fields: 'userEnteredFormat',
    }});
  }

  // Number formats
  for (const col of [2, 3, 4, 5, 6, 7]) {
    reqs.push({ repeatCell: {
      range: gridRange(HSA, 2, 4, col, col+1),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  }
  reqs.push({ repeatCell: {
    range: gridRange(HSA, 2, 4, 1, 2),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(HSA, 2, 4, 0, 1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});

  // Col widths
  for (const [i, px] of WIDTHS.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: HSA, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  // Info box row 5
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: HSA, dimension: 'ROWS', startIndex: 4, endIndex: 5 },
    properties: { pixelSize: 14 }, fields: 'pixelSize',
  }});
  reqs.push({ mergeCells: { range: gridRange(HSA, 5, 6, 0, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(HSA, 5, 6, 0, 9),
      cell: {
        userEnteredValue: { stringValue: 'ℹ️  Add one row per account per tax year. Annual Contribution Limit auto-fills via lookup. Withdrawals pull live from the Expense Log matching Payment Method to account type.' },
        userEnteredFormat: {
          backgroundColor: hex(C.lightSage),
          textFormat: { foregroundColor: hex(C.darkText), italic: true, fontSize: 9 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: HSA, dimension: 'ROWS', startIndex: 5, endIndex: 6 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});

  await batchUpdate(id, reqs, 'hsa-format');

  // Values
  const data = [];
  data.push({ range: `'💳 HSA/FSA Tracker'!A2:I2`, values: [HEADERS] });

  // Row 3 (index 2): HSA Self-Only 2024
  data.push({
    range: `'💳 HSA/FSA Tracker'!A3:I3`,
    values: [[
      'HSA Self-Only',
      2024,
      `=IFERROR(VLOOKUP(B3&" "&A3,'📋 Reference Data'!$F:$G,2,FALSE),"")`,
      1800,
      200,
      `=IFERROR(D3+E3,0)`,
      withdrawalFormula(3),
      `=IFERROR(F3-G3,0)`,
      'N/A for HSA — no grace period or use-it-or-lose-it rule',
    ]],
  });

  // Row 4 (index 3): FSA 2024
  data.push({
    range: `'💳 HSA/FSA Tracker'!A4:I4`,
    values: [[
      'FSA',
      2024,
      `=IFERROR(VLOOKUP(B4&" "&A4,'📋 Reference Data'!$F:$G,2,FALSE),"")`,
      2000,
      0,
      `=IFERROR(D4+E4,0)`,
      withdrawalFormula(4),
      `=IFERROR(F4-G4,0)`,
      'Grace period through Mar 15, 2025 — use-it-or-lose-it rules apply; check plan documents',
    ]],
  });

  await valuesBatchUpdate(id, data, 'hsa-values');
  console.log('HSA/FSA Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
