'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const PLANTS = sheetMap['Plant Library'];
const PS     = sheetMap['Planting Schedule'];
const WAT    = sheetMap['Watering & Care Schedule'];
const HARV   = sheetMap['Harvest & Yield Tracker'];
const PEST   = sheetMap['Pest & Disease Log'];
const BUD    = sheetMap['Garden Budget'];
const SEED   = sheetMap['Seed Inventory'];

function dropdown(sheetId, row0, row1, col0, col1, refRange) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, row0, row1, col0, col1),
      rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: '=' + refRange }] },
        strict: true,
        showCustomUi: true,
      },
    },
  };
}

function list(sheetId, row0, row1, col0, col1, items) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, row0, row1, col0, col1),
      rule: {
        condition: { type: 'ONE_OF_LIST', values: items.map(v => ({ userEnteredValue: v })) },
        strict: true,
        showCustomUi: true,
      },
    },
  };
}

function numberRange(sheetId, row0, row1, col0, col1, min, max) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, row0, row1, col0, col1),
      rule: {
        condition: { type: 'NUMBER_BETWEEN', values: [{ userEnteredValue: String(min) }, { userEnteredValue: String(max) }] },
        strict: false,
        showCustomUi: false,
        inputMessage: `Enter a value between ${min} and ${max}`,
      },
    },
  };
}

(async () => {
  const reqs = [];

  // ── Plant Library ─────────────────────────────────────────────────────────
  // Type col B (1), rows 3-37 (2-36)
  reqs.push(dropdown(PLANTS, 2, 37, 1, 2, "'Reference Data'!$A$2:$A$9"));
  // Season col C (2)
  reqs.push(dropdown(PLANTS, 2, 37, 2, 3, "'Reference Data'!$E$2:$E$7"));
  // Sun col D (3)
  reqs.push(dropdown(PLANTS, 2, 37, 3, 4, "'Reference Data'!$B$2:$B$5"));
  // Water col E (4)
  reqs.push(dropdown(PLANTS, 2, 37, 4, 5, "'Reference Data'!$C$2:$C$5"));
  // Zone col I (8)
  reqs.push(dropdown(PLANTS, 2, 37, 8, 9, "'Reference Data'!$O$2:$O$10"));
  // Status col L (11)
  reqs.push(dropdown(PLANTS, 2, 37, 11, 12, "'Reference Data'!$D$2:$D$11"));
  // Qty planted col N (13) — number 1-1000
  reqs.push(numberRange(PLANTS, 2, 37, 13, 14, 1, 1000));

  // ── Planting Schedule ─────────────────────────────────────────────────────
  // Season col B (1)
  reqs.push(dropdown(PS, 3, 18, 1, 2, "'Reference Data'!$E$2:$E$7"));
  // Method col C (2)
  reqs.push(list(PS, 3, 18, 2, 3, ['Indoors', 'Direct', 'Transplant']));
  // Zone col H (7)
  reqs.push(dropdown(PS, 3, 18, 7, 8, "'Reference Data'!$O$2:$O$10"));
  // Status col L (11)
  reqs.push(dropdown(PS, 3, 18, 11, 12, "'Reference Data'!$D$2:$D$11"));
  // Succession # col M (12) — number 1-10
  reqs.push(numberRange(PS, 3, 18, 12, 13, 1, 10));

  // ── Watering & Care ───────────────────────────────────────────────────────
  // Zone col B (1)
  reqs.push(dropdown(WAT, 2, 14, 1, 2, "'Reference Data'!$O$2:$O$10"));
  // Frequency col C (2)
  reqs.push(dropdown(WAT, 2, 14, 2, 3, "'Reference Data'!$N$2:$N$8"));
  // Water amount col D (3) — number 0.1-5
  reqs.push(numberRange(WAT, 2, 14, 3, 4, 0.1, 5));
  // Fertilize schedule col H (7)
  reqs.push(list(WAT, 2, 14, 7, 8, ['Daily', 'Weekly', 'Bi-weekly', 'Monthly', 'Once at planting', 'As Needed']));

  // ── Harvest Tracker ───────────────────────────────────────────────────────
  // Unit col D (3)
  reqs.push(dropdown(HARV, 2, 14, 3, 4, "'Reference Data'!$G$2:$G$9"));
  // Amount col C (2) — number > 0
  reqs.push(numberRange(HARV, 2, 14, 2, 3, 0.01, 9999));

  // ── Pest & Disease Log ────────────────────────────────────────────────────
  // Type col C (2)
  reqs.push(dropdown(PEST, 2, 8, 2, 3, "'Reference Data'!$I$2:$I$8"));
  // Severity col E (4)
  reqs.push(dropdown(PEST, 2, 8, 4, 5, "'Reference Data'!$J$2:$J$5"));
  // Status col H (7)
  reqs.push(dropdown(PEST, 2, 8, 7, 8, "'Reference Data'!$K$2:$K$5"));

  // ── Garden Budget ─────────────────────────────────────────────────────────
  // Category col B (1)
  reqs.push(dropdown(BUD, 3, 13, 1, 2, "'Reference Data'!$H$2:$H$9"));
  // Budgeted col E (4) — number > 0
  reqs.push(numberRange(BUD, 3, 13, 4, 5, 0.01, 99999));
  // Actual col F (5) — number ≥ 0
  reqs.push(numberRange(BUD, 3, 13, 5, 6, 0, 99999));

  // ── Seed Inventory ────────────────────────────────────────────────────────
  // Type col B (1)
  reqs.push(dropdown(SEED, 2, 12, 1, 2, "'Reference Data'!$A$2:$A$9"));
  // Season col C (2)
  reqs.push(dropdown(SEED, 2, 12, 2, 3, "'Reference Data'!$E$2:$E$7"));
  // Source col D (3)
  reqs.push(dropdown(SEED, 2, 12, 3, 4, "'Reference Data'!$L$2:$L$5"));
  // # Seeds col F (5) — number 1-10000
  reqs.push(numberRange(SEED, 2, 12, 5, 6, 1, 10000));
  // # Planted col G (6) — number 0-10000
  reqs.push(numberRange(SEED, 2, 12, 6, 7, 0, 10000));
  // Viability col J (9) — number 1-10
  reqs.push(numberRange(SEED, 2, 12, 9, 10, 1, 10));

  await batchUpdate(id, reqs, 'validation');
  console.log(`Data validation complete — ${reqs.length} rules`);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
