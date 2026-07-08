'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const PYR = sheetMap['📈 Prior-Year Comparison'];

// A=Tax Year | B=Total Expenses | C=Total Reimbursed | D=Out-of-Pocket
// E=Doctor Visits | F=Prescriptions | G=Dental | H=Vision | I=Emergency
// J=Hospital | K=Therapy | L=Equipment | M=Lab/Diagnostic | N=Preventive | O=Other
// P=YoY Change %

const CATEGORIES = [
  'Doctor Visits', 'Prescription Medications', 'Dental Care', 'Vision Care',
  'Emergency Services', 'Hospital Services', 'Therapy/Mental Health',
  'Medical Equipment', 'Lab/Diagnostic Tests', 'Preventive Care', 'Other',
];
const CAT_COLS = ['E','F','G','H','I','J','K','L','M','N','O'];
const YEARS = [2024, 2025, 2026];

const EXP = "'🧾 Expense Log'";

function totalExpenses(year, row) {
  return `=IFERROR(SUMPRODUCT((YEAR(${EXP}!$A$3:$A$200)=${year})*(${EXP}!$E$3:$E$200)),0)`;
}
function totalReimbursed(year, row) {
  return `=IFERROR(SUMPRODUCT((YEAR(${EXP}!$A$3:$A$200)=${year})*(${EXP}!$H$3:$H$200="Paid")*(${EXP}!$E$3:$E$200)),0)`;
}
function categoryTotal(year, cat) {
  return `=IFERROR(SUMPRODUCT((YEAR(${EXP}!$A$3:$A$200)=${year})*(${EXP}!$C$3:$C$200="${cat}")*(${EXP}!$E$3:$E$200)),0)`;
}

const HEADERS = ['Tax Year','Total Expenses','Total Reimbursed','Out-of-Pocket', ...CATEGORIES, 'YoY Change %'];
const WIDTHS  = [80, 120, 120, 110, 110, 130, 100, 90, 120, 110, 100, 110, 130, 110, 80, 100];

(async () => {
  const reqs = [];

  // Row 0: Title — col A standalone for freeze compatibility
  reqs.push({ mergeCells: { range: gridRange(PYR, 0, 1, 1, 16), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(PYR, 0, 1, 0, 1),
      cell: {
        userEnteredValue: { stringValue: '📈' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepTeal),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(PYR, 0, 1, 1, 16),
      cell: {
        userEnteredValue: { stringValue: 'Prior-Year Medical Expense Comparison' },
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
    range: { sheetId: PYR, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Row 1: Headers
  reqs.push({
    repeatCell: {
      range: gridRange(PYR, 1, 2, 0, 16),
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
    range: { sheetId: PYR, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Data rows 2-4 (3 years)
  for (let r = 2; r < 5; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(PYR, r, r+1, 0, 16),
        cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.inputBg : C.altRow), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: PYR, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
  }

  // All formula cols (B-P): formulaBg
  reqs.push({ repeatCell: {
    range: gridRange(PYR, 2, 5, 1, 16),
    cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg) } },
    fields: 'userEnteredFormat',
  }});

  // YoY Change col P (index 15): different bg
  reqs.push({ repeatCell: {
    range: gridRange(PYR, 2, 5, 15, 16),
    cell: { userEnteredFormat: { backgroundColor: hex(C.lightAmber) } },
    fields: 'userEnteredFormat',
  }});

  // Number formats: B-O = currency, P = percent
  reqs.push({ repeatCell: {
    range: gridRange(PYR, 2, 5, 1, 15),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(PYR, 2, 5, 15, 16),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '+0.0%;-0.0%' }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(PYR, 2, 5, 0, 1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true } } },
    fields: 'userEnteredFormat',
  }});

  // Col widths
  for (const [i, px] of WIDTHS.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: PYR, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  await batchUpdate(id, reqs, 'prioryear-format');

  // Values
  const data = [];
  data.push({ range: `'📈 Prior-Year Comparison'!A2:P2`, values: [HEADERS] });

  for (let i = 0; i < YEARS.length; i++) {
    const year = YEARS[i];
    const row = i + 3;
    const total  = totalExpenses(year, row);
    const reimb  = totalReimbursed(year, row);
    const oop    = `=IFERROR(B${row}-C${row},0)`;
    const catVals = CATEGORIES.map(cat => categoryTotal(year, cat));
    const yoy = row === 3 ? '' : `=IFERROR((B${row}-B${row-1})/B${row-1},"")`;

    data.push({
      range: `'📈 Prior-Year Comparison'!A${row}:P${row}`,
      values: [[year, total, reimb, oop, ...catVals, yoy]],
    });
  }

  await valuesBatchUpdate(id, data, 'prioryear-values');
  console.log('Prior-Year Comparison complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
