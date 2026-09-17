'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const ITN = sheetMap['🗓️ Multi-City Day-by-Day Itinerary'];
const BGT = sheetMap['💰 Budget & Expense Tracker'];
const VLT = sheetMap['🔐 Travel Document & Confirmation Vault'];
const WIS = sheetMap['✨ Activity & Excursion Wishlist'];
const CHK = sheetMap['✅ Research & Reservation Checklist'];
const PKG = sheetMap['🧳 Packing List'];
const OVW = sheetMap['🌅 Honeymoon Overview'];

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

function dropdownList(range, values, strict = true) {
  return {
    setDataValidation: {
      range,
      rule: {
        condition: { type: 'ONE_OF_LIST', values: values.map(v => ({ userEnteredValue: v })) },
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

  // ── Itinerary ─────────────────────────────────────────────────────────────
  // Col C (idx 2): Time of Day — ONE_OF_LIST
  reqs.push(dropdownList(gridRange(ITN, 1, 31, 2, 3), ['Morning', 'Afternoon', 'Evening', 'All Day']));
  // Col G (idx 6): Booked — checkbox
  reqs.push(checkbox(gridRange(ITN, 1, 31, 6, 7)));

  // ── Budget & Expense ──────────────────────────────────────────────────────
  // Col A (idx 0): Category — Ref!H
  reqs.push(dropdown(gridRange(BGT, 1, 26, 0, 1), "='📋 Reference Data'!$H$2:$H$9"));
  // Col F (idx 5): Currency — Ref!A
  reqs.push(dropdown(gridRange(BGT, 1, 26, 5, 6), "='📋 Reference Data'!$A$2:$A$11"));
  // Col G (idx 6): Payment Method — ONE_OF_LIST
  reqs.push(dropdownList(gridRange(BGT, 1, 26, 6, 7), ['Credit Card', 'Debit Card', 'Cash', 'Cash/Card', 'Bank Transfer', 'Various']));

  // ── Document Vault ────────────────────────────────────────────────────────
  // Col A (idx 0): Document Type — Ref!C
  reqs.push(dropdown(gridRange(VLT, 1, 16, 0, 1), "='📋 Reference Data'!$C$2:$C$10"));
  // Col E (idx 4): Status — Ref!D
  reqs.push(dropdown(gridRange(VLT, 1, 16, 4, 5), "='📋 Reference Data'!$D$2:$D$6"));

  // ── Activity Wishlist ─────────────────────────────────────────────────────
  // Col B (idx 1): Category — Ref!G
  reqs.push(dropdown(gridRange(WIS, 1, 21, 1, 2), "='📋 Reference Data'!$G$2:$G$10"));
  // Col E (idx 4): Priority — ONE_OF_LIST
  reqs.push(dropdownList(gridRange(WIS, 1, 21, 4, 5), ['High', 'Medium', 'Low']));
  // Col F (idx 5): Booking Status — Ref!D
  reqs.push(dropdown(gridRange(WIS, 1, 21, 5, 6), "='📋 Reference Data'!$D$2:$D$6"));

  // ── Checklist ─────────────────────────────────────────────────────────────
  // Col B (idx 1): Category — Ref!B
  reqs.push(dropdown(gridRange(CHK, 1, 26, 1, 2), "='📋 Reference Data'!$B$2:$B$9"));
  // Col D (idx 3): Assigned To — Ref!E
  reqs.push(dropdown(gridRange(CHK, 1, 26, 3, 4), "='📋 Reference Data'!$E$2:$E$4"));
  // Col E (idx 4): Done — checkbox
  reqs.push(checkbox(gridRange(CHK, 1, 26, 4, 5)));

  // ── Packing List ──────────────────────────────────────────────────────────
  // Col B (idx 1): Category — Ref!F
  reqs.push(dropdown(gridRange(PKG, 1, 34, 1, 2), "='📋 Reference Data'!$F$2:$F$9"));
  // Col C (idx 2): Partner — Ref!E
  reqs.push(dropdown(gridRange(PKG, 1, 34, 2, 3), "='📋 Reference Data'!$E$2:$E$4"));
  // Col D (idx 3): Packed — checkbox
  reqs.push(checkbox(gridRange(PKG, 1, 34, 3, 4)));

  // ── Overview: Currency converter From Currency (D15, idx row 14, col 3) ──
  reqs.push(dropdown(gridRange(OVW, 14, 15, 3, 4), "='📋 Reference Data'!$A$2:$A$11"));

  await batchUpdate(id, reqs, 'validation');
  console.log('Data validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
