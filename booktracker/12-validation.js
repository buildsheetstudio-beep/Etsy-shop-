'use strict';
const { batchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID_REF = sheetMap['Reference Data'];
const SID_LIB = sheetMap['Master Book Library'];
const SID_REV = sheetMap['Book Review & Notes'];
const SID_WL  = sheetMap['Wishlist'];
const SID_GOA = sheetMap['Goals & Challenges'];

// ONE_OF_RANGE dropdown pointing to a column in Reference Data
const rangeVal = (sheetId, r1, r2, c1, c2, refCol) => ({
  setDataValidation: {
    range: gridRange(sheetId, r1, r2, c1, c2),
    rule: {
      condition: {
        type: 'ONE_OF_RANGE',
        values: [{ userEnteredValue: `='Reference Data'!${refCol}2:${refCol}50` }],
      },
      showCustomUi: true,
      strict: false,
    },
  },
});

// BooleanCondition checkbox
const checkVal = (sheetId, r1, r2, c1, c2) => ({
  setDataValidation: {
    range: gridRange(sheetId, r1, r2, c1, c2),
    rule: {
      condition: { type: 'BOOLEAN', values: [] },
      showCustomUi: true,
      strict: true,
    },
  },
});

(async () => {
  const fmt = [];

  // ── Reference Data column mapping ────────────────────────────────────────────
  // A=Genre, B=Status, C=Format, D=Language, E=Priority, F=Source
  // G=Goal Type, H=Goal Status, I=Challenge Type, J=Yes/No, K=Shelf, L=Rating Stars

  // ── Master Book Library (data rows 8-1007) ────────────────────────────────────
  // D: Genre
  fmt.push(rangeVal(SID_LIB, 7, 1007, 3, 4, 'A'));
  // F: Format
  fmt.push(rangeVal(SID_LIB, 7, 1007, 5, 6, 'C'));
  // J: Language
  fmt.push(rangeVal(SID_LIB, 7, 1007, 9, 10, 'D'));
  // L: Shelf
  fmt.push(rangeVal(SID_LIB, 7, 1007, 11, 12, 'K'));
  // M: Status
  fmt.push(rangeVal(SID_LIB, 7, 1007, 12, 13, 'B'));
  // P: Rating (1-5)
  fmt.push(rangeVal(SID_LIB, 7, 1007, 15, 16, 'L'));
  // U: Favorite? checkbox
  fmt.push(checkVal(SID_LIB, 7, 1007, 20, 21));
  // V: Reread? checkbox
  fmt.push(checkVal(SID_LIB, 7, 1007, 21, 22));

  // ── Book Review & Notes (data rows 6-505) ─────────────────────────────────────
  // G: Plot Rating (1-5)
  fmt.push(rangeVal(SID_REV, 5, 505, 6, 7, 'L'));
  // H: Characters Rating
  fmt.push(rangeVal(SID_REV, 5, 505, 7, 8, 'L'));
  // I: Writing Style Rating
  fmt.push(rangeVal(SID_REV, 5, 505, 8, 9, 'L'));
  // J: Pacing Rating
  fmt.push(rangeVal(SID_REV, 5, 505, 9, 10, 'L'));
  // N: Read Again? (Yes/No)
  fmt.push(rangeVal(SID_REV, 5, 505, 13, 14, 'J'));
  // O: Would Recommend? (Yes/No)
  fmt.push(rangeVal(SID_REV, 5, 505, 14, 15, 'J'));

  // ── Wishlist (data rows 6-505) ────────────────────────────────────────────────
  // C: Genre
  fmt.push(rangeVal(SID_WL, 5, 505, 2, 3, 'A'));
  // D: Format
  fmt.push(rangeVal(SID_WL, 5, 505, 3, 4, 'C'));
  // K: Priority
  fmt.push(rangeVal(SID_WL, 5, 505, 10, 11, 'E'));
  // L: Source
  fmt.push(rangeVal(SID_WL, 5, 505, 11, 12, 'F'));
  // O: Acquired? checkbox
  fmt.push(checkVal(SID_WL, 5, 505, 14, 15));

  // ── Goals & Challenges ────────────────────────────────────────────────────────
  // Annual Goals section (rows 6-25): col G = Goal Status
  fmt.push(rangeVal(SID_GOA, 5, 25, 6, 7, 'H'));

  // Reading Challenges section (rows 28-55): col D = Challenge Type, col I = Goal Status
  fmt.push(rangeVal(SID_GOA, 27, 55, 3, 4, 'I'));
  fmt.push(rangeVal(SID_GOA, 27, 55, 8, 9, 'H'));

  // Series Tracking section (rows 58-85): no special dropdowns needed

  await batchUpdate(id, fmt, '12-validation');
  console.log('✅  Data validation done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
