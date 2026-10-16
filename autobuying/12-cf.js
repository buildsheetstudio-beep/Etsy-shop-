'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const VQ = sheetMap['Vehicle Details & Quotes'];
const MB = sheetMap['Monthly Budget Impact'];
const VC = sheetMap['Vehicle Comparison'];
const DB = sheetMap['Dashboard'];
const BS = sheetMap['Buyer Setup'];

function cfTextEq(sheetId, r1, r2, c1, c2, text, bgHex, textHex) {
  return {
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(sheetId, r1, r2, c1, c2)],
        booleanRule: {
          condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] },
          format: { backgroundColor: hex(bgHex), textFormat: { foregroundColor: hex(textHex || C.mainText) } }
        }
      },
      index: 0
    }
  };
}

function cfFormula(sheetId, r1, r2, c1, c2, formula, bgHex, textHex) {
  return {
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(sheetId, r1, r2, c1, c2)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: formula }] },
          format: { backgroundColor: hex(bgHex), textFormat: { foregroundColor: hex(textHex || C.mainText) } }
        }
      },
      index: 0
    }
  };
}

(async () => {
  const reqs = [];

  // ── Vehicle Details & Quotes: Quote Status (col AD = index 29) ─────────────
  reqs.push(cfTextEq(VQ, 5, 105, 29, 30, 'Selected',     C.success, '#1A3A1E'));
  reqs.push(cfTextEq(VQ, 5, 105, 29, 30, 'Final Offer',  C.warning, C.mainText));
  reqs.push(cfTextEq(VQ, 5, 105, 29, 30, 'Negotiating',  C.warning, C.mainText));
  reqs.push(cfTextEq(VQ, 5, 105, 29, 30, 'Rejected',     C.border,  C.secText));

  // ── Monthly Budget Impact: Affordability (row 33, index 32, cols D-K = 3-10) ─
  reqs.push(cfTextEq(MB, 32, 33, 3, 11, 'Comfortable',      C.success,   '#1A3A1E'));
  reqs.push(cfTextEq(MB, 32, 33, 3, 11, 'Manageable',       C.warning,   C.mainText));
  reqs.push(cfTextEq(MB, 32, 33, 3, 11, 'Tight',            C.attention, C.white));
  reqs.push(cfTextEq(MB, 32, 33, 3, 11, 'Over Budget',      C.rust,      C.white));
  reqs.push(cfTextEq(MB, 32, 33, 3, 11, 'Insufficient Data',C.slate,     C.white));

  // ── Vehicle Comparison: Affordability (col T = index 19) ─────────────────
  reqs.push(cfTextEq(VC, 16, 26, 19, 20, 'Comfortable',      C.success,   '#1A3A1E'));
  reqs.push(cfTextEq(VC, 16, 26, 19, 20, 'Manageable',       C.warning,   C.mainText));
  reqs.push(cfTextEq(VC, 16, 26, 19, 20, 'Tight',            C.attention, C.white));
  reqs.push(cfTextEq(VC, 16, 26, 19, 20, 'Over Budget',      C.rust,      C.white));
  reqs.push(cfTextEq(VC, 16, 26, 19, 20, 'Insufficient Data',C.slate,     C.white));

  // Vehicle Comparison: Recommendation (col V = index 21)
  reqs.push(cfTextEq(VC, 16, 26, 21, 22, 'Best Overall',        C.success,   '#1A3A1E'));
  reqs.push(cfTextEq(VC, 16, 26, 21, 22, 'Best Monthly Cost',   C.success,   '#1A3A1E'));
  reqs.push(cfTextEq(VC, 16, 26, 21, 22, 'Best Long-Term Value',C.success,   '#1A3A1E'));
  reqs.push(cfTextEq(VC, 16, 26, 21, 22, 'Review Carefully',    C.warning,   C.mainText));

  // Vehicle Comparison: Weight total (B14 = index row 13, col 1)
  // Amber if NOT 100%
  reqs.push(cfFormula(VC, 13, 14, 1, 2, '=B14<>1', C.warning, C.mainText));
  reqs.push(cfFormula(VC, 13, 14, 1, 2, '=B14=1',  C.success, '#1A3A1E'));

  // ── Dashboard: Affordability Status (G12 = index row 11, col 6) ───────────
  reqs.push(cfTextEq(DB, 11, 12, 6, 7, 'Comfortable',      C.success,   '#1A3A1E'));
  reqs.push(cfTextEq(DB, 11, 12, 6, 7, 'Manageable',       C.warning,   C.mainText));
  reqs.push(cfTextEq(DB, 11, 12, 6, 7, 'Tight',            C.attention, C.white));
  reqs.push(cfTextEq(DB, 11, 12, 6, 7, 'Over Budget',      C.rust,      C.white));
  reqs.push(cfTextEq(DB, 11, 12, 6, 7, 'Insufficient Data',C.slate,     C.white));

  // Dashboard: Emergency Fund remaining — flag if below zero (negative = dipped below minimum after down pmt)
  reqs.push(cfFormula(DB, 11, 12, 4, 5, `=E12<0`, C.rust, C.white));
  reqs.push(cfFormula(DB, 11, 12, 4, 5, `=AND(E12>=0,E12<6000)`, C.warning, C.mainText));

  // Buyer Setup: Net Trade-In Equity negative (B35 — index row 34, col 1)
  reqs.push(cfFormula(BS, 34, 35, 1, 2, `=B35<0`, C.rust, C.white));

  await batchUpdate(id, reqs, 'conditional-formatting');
  console.log('✓ Conditional Formatting');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
