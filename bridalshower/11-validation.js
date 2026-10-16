'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const REF = sheetMap['📋 Reference Data'];
const GST = sheetMap['👰 Guest List & Seating'];
const BUD = sheetMap['💰 Budget & Expenses'];
const ITN = sheetMap['🗓️ Party Itinerary & Games'];
const VEN = sheetMap['💍 Venue, Vendors & Bridesmaids'];
const GFT = sheetMap['🎁 Gift Tracker & Wishlist'];

function dropdown(sheetId, r1, r2, c1, c2, refCol, refRow1, refRow2) {
  return { setDataValidation: { range: gridRange(sheetId, r1, r2, c1, c2), rule: {
    condition: {
      type: 'ONE_OF_RANGE',
      values: [{ userEnteredValue: `='📋 Reference Data'!$${refCol}$${refRow1}:$${refCol}$${refRow2}` }],
    },
    strict: false, showCustomUi: true,
  }}};
}

function checkbox(sheetId, r1, r2, c1, c2) {
  return { setDataValidation: { range: gridRange(sheetId, r1, r2, c1, c2), rule: {
    condition: { type: 'BOOLEAN' }, strict: true, showCustomUi: true,
  }}};
}

(async () => {
  const reqs = [];

  // ── Guest List ──
  // Col B: Relationship (A2:A10 in ref)
  reqs.push(dropdown(GST, 3, 23, 1, 2,  'A', 2, 10));
  // Col E: RSVP Status (B2:B6 in ref)
  reqs.push(dropdown(GST, 3, 23, 4, 5,  'B', 2, 6));
  // Col G: Dietary (C2:C10 in ref)
  reqs.push(dropdown(GST, 3, 23, 6, 7,  'C', 2, 10));
  // Col H: Plus One? (custom Yes/No)
  reqs.push({ setDataValidation: { range: gridRange(GST, 3, 23, 7, 8), rule: {
    condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'Yes' }, { userEnteredValue: 'No' }] },
    strict: false, showCustomUi: true,
  }}});
  // Col J: Gift Received? (Yes/No)
  reqs.push({ setDataValidation: { range: gridRange(GST, 3, 23, 9, 10), rule: {
    condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'Yes' }, { userEnteredValue: 'No' }] },
    strict: false, showCustomUi: true,
  }}});
  // Col K: Thank You? (Yes/No)
  reqs.push({ setDataValidation: { range: gridRange(GST, 3, 23, 10, 11), rule: {
    condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'Yes' }, { userEnteredValue: 'No' }] },
    strict: false, showCustomUi: true,
  }}});
  // Col L: Table Name (L2:L9 in ref)
  reqs.push(dropdown(GST, 3, 23, 11, 12, 'L', 2, 9));

  // ── Budget ──
  // Col A (expense log): Budget Category (K2:K13 in ref)
  reqs.push(dropdown(BUD, 19, 60, 0, 1, 'K', 2, 13));
  // Col H (expense log): Payment Method (J2:J8 in ref)
  reqs.push(dropdown(BUD, 19, 60, 7, 8, 'J', 2, 8));
  // Col G: Paid? checkbox
  reqs.push(checkbox(BUD, 19, 60, 6, 7));

  // ── Itinerary ──
  // Col E: Activity Type (M2:M12 in ref)
  reqs.push(dropdown(ITN, 4, 18, 4, 5, 'M', 2, 12));
  // Col C (games): Game Type (D2:D8 in ref)
  reqs.push(dropdown(ITN, 21, 31, 1, 2, 'D', 2, 8));

  // ── Vendors / Bridesmaids ──
  // Vendor comparison row 9: Contract Signed (Yes/No)
  reqs.push({ setDataValidation: { range: gridRange(VEN, 12, 13, 1, 8), rule: {
    condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'Yes' }, { userEnteredValue: 'No' }, { userEnteredValue: 'Pending' }] },
    strict: false, showCustomUi: true,
  }}});
  // Vendor comparison row 10: Rating (Q2:Q6 in ref)
  reqs.push(dropdown(VEN, 13, 14, 1, 8, 'Q', 2, 6));
  // BM task col E: Priority (G2:G4 in ref)
  reqs.push(dropdown(VEN, 19, 37, 4, 5, 'G', 2, 4));
  // BM task col F: Task Status (F2:F5 in ref)
  reqs.push(dropdown(VEN, 19, 37, 5, 6, 'F', 2, 5));
  // BM task col B: Role (N2:N5 in ref)
  reqs.push(dropdown(VEN, 19, 37, 1, 2, 'N', 2, 5));
  // BM task col D: Task Category (O2:O8 in ref)
  reqs.push(dropdown(VEN, 19, 37, 3, 4, 'O', 2, 8));
  // BM task col H: Done? checkbox
  reqs.push(checkbox(VEN, 19, 37, 7, 8));
  // Supplier contacts: Category col B (E2:E8 in ref = Vendor Category)
  reqs.push(dropdown(VEN, 40, 48, 1, 2, 'E', 2, 8));

  // ── Gifts ──
  // Wishlist col B: Gift Category (H2:H10 in ref)
  reqs.push(dropdown(GFT, 3, 18, 1, 2, 'H', 2, 10));
  // Wishlist col C: Priority (I2:I4 in ref)
  reqs.push(dropdown(GFT, 3, 18, 2, 3, 'I', 2, 4));
  // Gift log col D: Category (H2:H10 in ref)
  reqs.push(dropdown(GFT, 22, 60, 3, 4, 'H', 2, 10));
  // Gift log col F: Payment Method (J2:J8 in ref)
  reqs.push(dropdown(GFT, 22, 60, 5, 6, 'J', 2, 8));
  // Gift log col G: Received? (Yes/No)
  reqs.push({ setDataValidation: { range: gridRange(GFT, 22, 60, 6, 7), rule: {
    condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'Yes' }, { userEnteredValue: 'No' }] },
    strict: false, showCustomUi: true,
  }}});
  // Gift log col H: Thank You status
  reqs.push({ setDataValidation: { range: gridRange(GFT, 22, 60, 7, 8), rule: {
    condition: { type: 'ONE_OF_LIST', values: [
      { userEnteredValue: 'Thank You Sent' },
      { userEnteredValue: 'Pending' },
      { userEnteredValue: 'Not Started' },
    ]},
    strict: false, showCustomUi: true,
  }}});

  await batchUpdate(id, reqs, 'validation');
  console.log('Data Validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
