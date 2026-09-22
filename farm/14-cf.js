'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
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

function cfText(sheetId, r1, r2, c1, c2, contains, bgHex, textHex) {
  return {
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(sheetId, r1, r2, c1, c2)],
        booleanRule: {
          condition: { type: 'TEXT_CONTAINS', values: [{ userEnteredValue: contains }] },
          format: {
            backgroundColor: hex(bgHex),
            textFormat: { foregroundColor: hex(textHex) },
          },
        },
      },
      index: 0,
    },
  };
}

function cfGT(sheetId, r1, r2, c1, c2, value, bgHex) {
  return {
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(sheetId, r1, r2, c1, c2)],
        booleanRule: {
          condition: { type: 'NUMBER_GREATER', values: [{ userEnteredValue: String(value) }] },
          format: { backgroundColor: hex(bgHex) },
        },
      },
      index: 0,
    },
  };
}

function cfLT(sheetId, r1, r2, c1, c2, value, bgHex) {
  return {
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(sheetId, r1, r2, c1, c2)],
        booleanRule: {
          condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: String(value) }] },
          format: { backgroundColor: hex(bgHex) },
        },
      },
      index: 0,
    },
  };
}

function cfEq(sheetId, r1, r2, c1, c2, value, bgHex, textHex) {
  return {
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(sheetId, r1, r2, c1, c2)],
        booleanRule: {
          condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: value }] },
          format: {
            backgroundColor: hex(bgHex),
            textFormat: { foregroundColor: hex(textHex || C.mainText) },
          },
        },
      },
      index: 0,
    },
  };
}

(async () => {
  const reqs = [];

  // ── Income Log Status column M (col 12, 0-indexed) rows 6-505 ─────────────
  reqs.push(cfEq(IL, 5, 505, 12, 13, 'Paid',    C.success,   C.white));
  reqs.push(cfEq(IL, 5, 505, 12, 13, 'Pending', C.warning,   C.mainText));
  reqs.push(cfEq(IL, 5, 505, 12, 13, 'Void',    C.attention, C.white));

  // ── Expense Log Payment Status column N (col 13) rows 6-505 ───────────────
  reqs.push(cfEq(EL, 5, 505, 13, 14, 'Paid',    C.success,   C.white));
  reqs.push(cfEq(EL, 5, 505, 13, 14, 'Pending', C.warning,   C.mainText));
  reqs.push(cfEq(EL, 5, 505, 13, 14, 'Overdue', C.attention, C.white));

  // Potentially Deductible? col O (14)
  reqs.push(cfEq(EL, 5, 505, 14, 15, 'Yes', '#E8F5E9', C.mainText));
  reqs.push(cfEq(EL, 5, 505, 14, 15, 'No',  '#FFF3E0', C.mainText));

  // ── Annual Summary — profit column D (col 3) rows 9-20 ───────────────────
  reqs.push(cfGT(AS, 8, 21, 3, 4, 0,  '#E8F5E9'));  // positive = sage
  reqs.push(cfLT(AS, 8, 21, 3, 4, 0,  '#FFEBEE'));  // negative = light red

  // ── Enterprise Profitability — Performance Flag col J (9) rows 9-23 ───────
  reqs.push(cfText(EP, 8, 24, 9, 10, 'Strong Profit',       C.success,   C.white));
  reqs.push(cfText(EP, 8, 24, 9, 10, 'Profitable',          '#E8F5E9',   C.mainText));
  reqs.push(cfText(EP, 8, 24, 9, 10, 'Near Break-Even',     C.warning,   C.mainText));
  reqs.push(cfText(EP, 8, 24, 9, 10, 'Loss',                C.attention, C.white));

  // Net profit col D (3) positive/negative
  reqs.push(cfGT(EP, 8, 24, 3, 4, 0,  '#E8F5E9'));
  reqs.push(cfLT(EP, 8, 24, 3, 4, 0,  '#FFEBEE'));

  // ── Budget vs Actual — Status col F (5) income rows 9-25 ─────────────────
  reqs.push(cfText(BA, 8, 26, 5, 6, 'Above Plan',   C.success,   C.white));
  reqs.push(cfText(BA, 8, 26, 5, 6, 'On Track',     '#E8F5E9',   C.mainText));
  reqs.push(cfText(BA, 8, 26, 5, 6, 'Below Plan',   C.attention, C.white));

  // Status col F (5) expense rows 30-59
  reqs.push(cfText(BA, 29, 60, 5, 6, 'Under Budget', C.success,   C.white));
  reqs.push(cfText(BA, 29, 60, 5, 6, 'Near Budget',  '#E8F5E9',   C.mainText));
  reqs.push(cfText(BA, 29, 60, 5, 6, 'Over Budget',  C.attention, C.white));

  // Variance col D (3) — positive green, negative red
  reqs.push(cfGT(BA, 8, 26, 3, 4, 0,  '#E8F5E9'));
  reqs.push(cfLT(BA, 8, 26, 3, 4, 0,  '#FFEBEE'));
  reqs.push(cfGT(BA, 29, 60, 3, 4, 0, '#E8F5E9'));
  reqs.push(cfLT(BA, 29, 60, 3, 4, 0, '#FFEBEE'));

  // ── Profit Goals — Status col G (6) monthly rows 9-20 ────────────────────
  reqs.push(cfText(PG, 8, 21, 6, 7, 'Goal Met',    C.success,   C.white));
  reqs.push(cfText(PG, 8, 21, 6, 7, 'In Progress', C.warning,   C.mainText));
  reqs.push(cfText(PG, 8, 21, 6, 7, 'Below Goal',  C.attention, C.white));

  // Variance col F (5)
  reqs.push(cfGT(PG, 8, 21, 5, 6, 0, '#E8F5E9'));
  reqs.push(cfLT(PG, 8, 21, 5, 6, 0, '#FFEBEE'));

  // Cumulative On Track col E (4)
  reqs.push(cfText(PG, 24, 36, 4, 5, 'On Track', C.success, C.white));
  reqs.push(cfText(PG, 24, 36, 4, 5, 'Behind',   C.attention, C.white));

  // ── Monthly View — net profit col D rows 8-22 (0-indexed 7-21) ───────────
  reqs.push(cfGT(MV, 7, 22, 3, 4, 0, '#E8F5E9'));
  reqs.push(cfLT(MV, 7, 22, 3, 4, 0, '#FFEBEE'));

  // ── Tax Deduction Summary — Deductible? col E rows 28-57 (27-57 0-indexed) ─
  reqs.push(cfText(TD, 27, 58, 4, 5, 'Yes', '#E8F5E9', C.mainText));
  reqs.push(cfText(TD, 27, 58, 4, 5, '—',   '#FFF3E0', C.mainText));

  await batchUpdate(id, reqs, 'cf');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
