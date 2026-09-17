'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const IL = sheetMap['Income Log'];
const EL = sheetMap['Expense Log'];
const MV = sheetMap['Monthly View'];
const AS = sheetMap['Annual Summary'];
const EP = sheetMap['Enterprise Profitability'];
const BA = sheetMap['Budget vs Actual'];
const PG = sheetMap['Profit Goals'];
const TD = sheetMap['Tax Deduction Summary'];
const DB = sheetMap['Dashboard'];

function dropdown(sheetId, r1, r2, c1, c2, sourceRange) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r1, r2, c1, c2),
      rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${sourceRange}` }] },
        showCustomUi: true,
        strict: true,
      },
    },
  };
}

function listDropdown(sheetId, r1, r2, c1, c2, values) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r1, r2, c1, c2),
      rule: {
        condition: { type: 'ONE_OF_LIST', values: values.map(v => ({ userEnteredValue: v })) },
        showCustomUi: true,
        strict: true,
      },
    },
  };
}

function dateValidation(sheetId, r1, r2, c1, c2) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r1, r2, c1, c2),
      rule: {
        condition: { type: 'DATE_IS_VALID' },
        showCustomUi: true,
        strict: false,
      },
    },
  };
}

function numberValidation(sheetId, r1, r2, c1, c2, min) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r1, r2, c1, c2),
      rule: {
        condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: String(min) }] },
        showCustomUi: true,
        strict: false,
      },
    },
  };
}

// Reference Data column refs
const REF = "'Reference Data'";
const ENT_RANGE  = `${REF}!$A$2:$A$16`;   // 15 enterprises
const INC_RANGE  = `${REF}!$C$2:$C$18`;   // 17 income cats
const EXP_RANGE  = `${REF}!$E$2:$E$31`;   // 30 expense cats
const PAY_RANGE  = `${REF}!$G$2:$G$9`;    // 8 payment methods
const ISTATUS    = `${REF}!$H$2:$H$4`;    // Income status
const ESTATUS    = `${REF}!$I$2:$I$4`;    // Expense status
const YESNO      = `${REF}!$J$2:$J$3`;    // Yes/No
const ENT_F_RANGE= `${REF}!$R$2:$R$17`;  // All Enterprises + 15
const PER_RANGE  = `${REF}!$S$2:$S$14`;  // Full Year + 12 months
const QTR_RANGE  = `${REF}!$T$2:$T$6`;   // Full Year + Q1-Q4

(async () => {
  const reqs = [];

  // ── Income Log (rows 6-505 = 0-indexed 5-504) ─────────────────────────────
  // B: date
  reqs.push(dateValidation(IL, 5, 505, 1, 2));
  // F: Enterprise
  reqs.push(dropdown(IL, 5, 505, 5, 6, ENT_RANGE));
  // G: Income Category
  reqs.push(dropdown(IL, 5, 505, 6, 7, INC_RANGE));
  // J: Payment Method
  reqs.push(dropdown(IL, 5, 505, 9, 10, PAY_RANGE));
  // K: Amount >= 0
  reqs.push(numberValidation(IL, 5, 505, 10, 11, 0));
  // M: Status
  reqs.push(dropdown(IL, 5, 505, 12, 13, ISTATUS));

  // ── Expense Log (rows 6-505 = 0-indexed 5-504) ────────────────────────────
  // B: date
  reqs.push(dateValidation(EL, 5, 505, 1, 2));
  // F: Enterprise
  reqs.push(dropdown(EL, 5, 505, 5, 6, ENT_RANGE));
  // G: Expense Category
  reqs.push(dropdown(EL, 5, 505, 6, 7, EXP_RANGE));
  // H: Vendor — free text, no validation
  // J: Payment Method
  reqs.push(dropdown(EL, 5, 505, 9, 10, PAY_RANGE));
  // K,L: >= 0
  reqs.push(numberValidation(EL, 5, 505, 10, 11, 0));
  reqs.push(numberValidation(EL, 5, 505, 11, 12, 0));
  // N: Payment Status
  reqs.push(dropdown(EL, 5, 505, 13, 14, ESTATUS));
  // O: Potentially Deductible?
  reqs.push(dropdown(EL, 5, 505, 14, 15, YESNO));

  // ── Monthly View controls B3:B4 (period + enterprise) ────────────────────
  reqs.push(dropdown(MV, 2, 3, 1, 2, PER_RANGE));
  reqs.push(dropdown(MV, 3, 4, 1, 2, ENT_F_RANGE));
  // Year B2: number
  reqs.push(numberValidation(MV, 1, 2, 1, 2, 2020));

  // ── Annual Summary controls B3 (enterprise) ──────────────────────────────
  reqs.push(dropdown(AS, 2, 3, 1, 2, ENT_F_RANGE));
  reqs.push(numberValidation(AS, 1, 2, 1, 2, 2020));

  // ── Enterprise Profitability controls B3 (period) ────────────────────────
  reqs.push(dropdown(EP, 2, 3, 1, 2, PER_RANGE));
  reqs.push(numberValidation(EP, 1, 2, 1, 2, 2020));

  // ── Budget vs Actual controls B3 (period) ────────────────────────────────
  reqs.push(dropdown(BA, 2, 3, 1, 2, PER_RANGE));
  reqs.push(numberValidation(BA, 1, 2, 1, 2, 2020));

  // ── Profit Goals controls B2 (year) ──────────────────────────────────────
  reqs.push(numberValidation(PG, 1, 2, 1, 2, 2020));

  // ── Tax Deduction Summary controls B3 (quarter) ──────────────────────────
  reqs.push(dropdown(TD, 2, 3, 1, 2, QTR_RANGE));
  reqs.push(numberValidation(TD, 1, 2, 1, 2, 2020));

  // ── Dashboard controls B3/B4 (period/enterprise) ────────────────────────
  reqs.push(dropdown(DB, 2, 3, 1, 2, PER_RANGE));
  reqs.push(dropdown(DB, 3, 4, 1, 2, ENT_F_RANGE));
  reqs.push(numberValidation(DB, 1, 2, 1, 2, 2020));

  await batchUpdate(id, reqs, 'validation');
  console.log('Validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
