'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID_LIB = sheetMap['Master Book Library'];
const SID_REV = sheetMap['Book Review & Notes'];
const SID_WL  = sheetMap['Wishlist'];
const SID_GOA = sheetMap['Goals & Challenges'];

const cfEq = (sheetId, r1, r2, c1, c2, text, bgColor) => ({
  addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(sheetId, r1, r2, c1, c2)],
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] },
        format: { backgroundColor: hex(bgColor) }
      }
    },
    index: 0
  }
});

const cfGte = (sheetId, r1, r2, c1, c2, val, bgColor) => ({
  addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(sheetId, r1, r2, c1, c2)],
      booleanRule: {
        condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: String(val) }] },
        format: { backgroundColor: hex(bgColor) }
      }
    },
    index: 0
  }
});

const cfLte = (sheetId, r1, r2, c1, c2, val, bgColor) => ({
  addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(sheetId, r1, r2, c1, c2)],
      booleanRule: {
        condition: { type: 'NUMBER_LESS_THAN_EQ', values: [{ userEnteredValue: String(val) }] },
        format: { backgroundColor: hex(bgColor) }
      }
    },
    index: 0
  }
});

(async () => {
  const fmt = [];

  // ── Master Book Library: Status column M (12) ─────────────────────────────
  [
    ['Finished',  C.greenTint ],
    ['Reading',   C.goldTint  ],
    ['To Read',   C.bg        ],
    ['Paused',    C.altRow    ],
    ['DNF',       C.wineTint  ],
    ['Abandoned', C.wineTint  ],
  ].forEach(([text, bg]) => fmt.push(cfEq(SID_LIB, 7, 1008, 12, 13, text, bg)));

  // Master Book Library: Rating column P (15) — color by star level
  fmt.push(cfGte(SID_LIB, 7, 1008, 15, 16, 5, '#F7EDD8'));  // 5★ = gold tint
  fmt.push(cfGte(SID_LIB, 7, 1008, 15, 16, 4, C.greenTint)); // 4★ = green tint
  fmt.push(cfLte(SID_LIB, 7, 1008, 15, 16, 2, C.wineTint));  // 1-2★ = wine tint

  // Master Book Library: Progress % column S (18) — highlight high progress
  fmt.push(cfGte(SID_LIB, 7, 1008, 18, 19, 1, C.greenTint));   // 100% = green
  fmt.push(cfGte(SID_LIB, 7, 1008, 18, 19, 0.5, C.goldTint));  // 50%+ = gold
  fmt.push(cfLte(SID_LIB, 7, 1008, 18, 19, 0.25, C.wineTint)); // ≤25% = wine

  // Master Book Library: Favorite? column U (20) — checkbox true
  fmt.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(SID_LIB, 7, 1008, 20, 21)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$U8=TRUE' }] },
          format: { backgroundColor: hex('#FFFDE0') }
        }
      },
      index: 0
    }
  });

  // ── Book Review & Notes: Overall rating column F (5) ─────────────────────
  fmt.push(cfGte(SID_REV, 5, 506, 5, 6, 5, '#F7EDD8'));   // 5★
  fmt.push(cfGte(SID_REV, 5, 506, 5, 6, 4, C.greenTint));  // 4★
  fmt.push(cfLte(SID_REV, 5, 506, 5, 6, 2, C.wineTint));   // ≤2★

  // Book Review & Notes: Would Recommend? column O (14)
  fmt.push(cfEq(SID_REV, 5, 506, 14, 15, 'Yes', C.greenTint));
  fmt.push(cfEq(SID_REV, 5, 506, 14, 15, 'No',  C.wineTint));

  // ── Wishlist: Priority column K (10) ─────────────────────────────────────
  fmt.push(cfEq(SID_WL, 5, 506, 10, 11, 'High',   '#FFE8E8'));  // red tint
  fmt.push(cfEq(SID_WL, 5, 506, 10, 11, 'Medium', '#FFF8E0'));  // amber tint
  fmt.push(cfEq(SID_WL, 5, 506, 10, 11, 'Low',    C.altRow));

  // Wishlist: Acquired? column O (14) — checkbox true → green
  fmt.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(SID_WL, 5, 506, 14, 15)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$O6=TRUE' }] },
          format: { backgroundColor: hex(C.greenTint) }
        }
      },
      index: 0
    }
  });

  // ── Goals & Challenges: Status columns ────────────────────────────────────
  // Annual goals status col G (6)
  [
    ['Achieved',       C.greenTint ],
    ['In Progress',    C.goldTint  ],
    ['Behind Schedule',C.wineTint  ],
    ['Not Started',    C.altRow    ],
    ['Paused',         C.altRow    ],
    ['Abandoned',      C.wineTint  ],
  ].forEach(([text, bg]) => fmt.push(cfEq(SID_GOA, 5, 100, 6, 7, text, bg)));

  // Goals: % complete col F (5) — color bands
  fmt.push(cfGte(SID_GOA, 5, 100, 5, 6, 1,    C.greenTint));
  fmt.push(cfGte(SID_GOA, 5, 100, 5, 6, 0.75, C.goldTint));
  fmt.push(cfLte(SID_GOA, 5, 100, 5, 6, 0.5,  C.wineTint));

  await batchUpdate(id, fmt, '11-cf');
  console.log('✅  Conditional formatting done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
