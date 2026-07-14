'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const GRT = sheetMap['🎁 Gift Recipient Tracker'];
const WSH = sheetMap['💡 Gift Idea Wishlist'];
const CDL = sheetMap['💌 Holiday Card & Mailing List'];
const RET = sheetMap['🔄 Return & Exchange Tracker'];

function dropdown(range, refRange) {
  return {
    setDataValidation: {
      range,
      rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: refRange }] },
        showCustomUi: true,
        strict: false,
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
  // Col B (idx 1): Relationship
  reqs.push(dropdown(gridRange(GRT, 1, 26, 1, 2), "='📋 Reference Data'!$A$2:$A$10"));
  // Col K (idx 10): Delivery Method
  reqs.push(dropdown(gridRange(GRT, 1, 26, 10, 11), "='📋 Reference Data'!$B$2:$B$6"));
  // Col M (idx 12): Priority
  reqs.push(dropdown(gridRange(GRT, 1, 26, 12, 13), "='📋 Reference Data'!$C$2:$C$4"));
  // Cols H, I, J (idx 7,8,9): Bought, Wrapped, Delivered checkboxes
  reqs.push(checkbox(gridRange(GRT, 1, 26, 7, 10)));

  // ── Gift Idea Wishlist ────────────────────────────────────────────────────
  // Col E (idx 4): Priority
  reqs.push(dropdown(gridRange(WSH, 1, 21, 4, 5), "='📋 Reference Data'!$C$2:$C$4"));
  // Col G (idx 6): Selected checkbox
  reqs.push(checkbox(gridRange(WSH, 1, 21, 6, 7)));

  // ── Holiday Card & Mailing List ───────────────────────────────────────────
  // Col C (idx 2): Card Status
  reqs.push(dropdown(gridRange(CDL, 1, 21, 2, 3), "='📋 Reference Data'!$D$2:$D$4"));
  // Col D (idx 3): Card Received checkbox
  reqs.push(checkbox(gridRange(CDL, 1, 21, 3, 4)));

  // ── Return & Exchange Tracker ─────────────────────────────────────────────
  // Col C (idx 2): Store — from Reference Data col F
  reqs.push(dropdown(gridRange(RET, 1, 11, 2, 3), "='📋 Reference Data'!$F$2:$F$10"));
  // Col H (idx 7): Return Status
  reqs.push(dropdown(gridRange(RET, 1, 11, 7, 8), "='📋 Reference Data'!$E$2:$E$6"));

  await batchUpdate(id, reqs, 'validation');
  console.log('Data validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
