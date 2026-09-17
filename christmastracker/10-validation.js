'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const GRT = sheetMap['🎁 Gift Recipient Tracker'];
const WSH = sheetMap['💡 Gift Idea Wishlist'];
const CDL = sheetMap['💌 Holiday Card & Mailing List'];
const RET = sheetMap['🔄 Return & Exchange Tracker'];

function dropdown(range, refRange, strict = true) {
  return {
    setDataValidation: {
      range,
      rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: refRange }] },
        showCustomUi: true,
        strict,
      },
    },
  };
}

function checkbox(range) {
  return {
    setDataValidation: {
      range,
      rule: {
        condition: { type: 'BOOLEAN' },
        strict: true,
        showCustomUi: true,
      },
    },
  };
}

(async () => {
  const reqs = [];

  // ── Gift Recipient Tracker ────────────────────────────────────────────────
  // Col B (idx 1): Relationship — Reference A2:A6 (5 items)
  reqs.push(dropdown(gridRange(GRT, 1, 26, 1, 2), "='📋 Reference Data'!$A$2:$A$6"));
  // Col K (idx 10): Delivery Method — Reference B2:B4 (3 items)
  reqs.push(dropdown(gridRange(GRT, 1, 26, 10, 11), "='📋 Reference Data'!$B$2:$B$4"));
  // Cols H, I, J (idx 7-9): Bought, Wrapped, Sent checkboxes
  reqs.push(checkbox(gridRange(GRT, 1, 26, 7, 10)));

  // ── Gift Idea Wishlist ────────────────────────────────────────────────────
  // Col D (idx 3): Priority — Reference C2:C4 (3 items)
  reqs.push(dropdown(gridRange(WSH, 1, 21, 3, 4), "='📋 Reference Data'!$C$2:$C$4"));
  // Col G (idx 6): Selected checkbox
  reqs.push(checkbox(gridRange(WSH, 1, 21, 6, 7)));

  // ── Holiday Card & Mailing List ───────────────────────────────────────────
  // Col C (idx 2): Card Status — Reference D2:D4 (3 items)
  reqs.push(dropdown(gridRange(CDL, 1, 21, 2, 3), "='📋 Reference Data'!$D$2:$D$4"));

  // ── Return & Exchange Tracker ─────────────────────────────────────────────
  // Col H (idx 7): Return/Exchange status — Reference E2:E5 (4 items)
  reqs.push(dropdown(gridRange(RET, 1, 11, 7, 8), "='📋 Reference Data'!$E$2:$E$5"));

  await batchUpdate(id, reqs, 'validation');
  console.log('Data validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
