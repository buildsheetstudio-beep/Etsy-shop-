'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GST = sheetMap['👥 Guest List & RSVP'];
const BUD = sheetMap['💰 Budget & Vendor Tracker'];
const TML = sheetMap['⏰ Day-Of Timeline'];
const SEA = sheetMap['🪑 Seating Chart'];
const TST = sheetMap['🥂 Toast & Speech Planner'];
const PHO = sheetMap['📸 Photo & Shot List'];
const DEC = sheetMap['🌸 Decor, Food & Drinks'];
const CHK = sheetMap['✅ Party Checklist'];

const REF = "'📋 Reference Data'";

// Column reference ranges (in Reference Data tab, 0-indexed header at row 0, data from row 1):
// A=RSVP Status (rows 1-4 = A2:A5), B=Dietary (B2:B10), C=Meal (C2:C7), D=Gift (D2:D6),
// E=Side (E2:E5), F=Vendor Cat (F2:F11), G=Payment (G2:G5), H=Timeline Type (H2:H9... actually 10 items H2:H11),
// I=Table Shape (I2:I5), J=Toast Status (J2:J6), K=Shot Type (K2:K9), L=Shot Status (L2:L4),
// M=Item Category (M2:M7), N=Task Category (N2:N8), O=Task Status (O2:O6), P=Priority (P2:P4)

function dv(sheetId, r1, r2, c1, c2, refRange, strict) {
  return { setDataValidation: {
    range: gridRange(sheetId, r1, r2, c1, c2),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!${refRange}` }] },
      showCustomUi: true,
      strict: !!strict,
    },
  }};
}

function dvBool(sheetId, r1, r2, c1, c2) {
  return { setDataValidation: {
    range: gridRange(sheetId, r1, r2, c1, c2),
    rule: { condition: { type: 'BOOLEAN' }, strict: true },
  }};
}

(async () => {
  const reqs = [];

  // ── Guest List ──
  reqs.push(dv(GST, 3, 200, 4, 5,  '$E$2:$E$5',  true));  // Side (col E, index 4)
  reqs.push(dv(GST, 3, 200, 5, 6,  '$A$2:$A$5',  true));  // RSVP Status (col F, index 5)
  reqs.push(dv(GST, 3, 200, 7, 8,  '$B$2:$B$10', false)); // Dietary (col H, index 7)
  reqs.push(dv(GST, 3, 200, 8, 9,  '$C$2:$C$7',  false)); // Meal Choice (col I, index 8)
  reqs.push(dvBool(GST, 3, 200, 9, 10));                   // Plus One checkbox (col J, index 9)
  reqs.push(dv(GST, 3, 200, 11, 12,'$D$2:$D$6',  false)); // Gift Status (col L, index 11)
  reqs.push(dvBool(GST, 3, 200, 12, 13));                  // TY Sent checkbox (col M, index 12)

  // ── Budget ──
  reqs.push(dv(BUD, 3, 200, 2, 3, '$F$2:$F$11', false));  // Category (col C, index 2)
  reqs.push(dv(BUD, 3, 200, 6, 7, '$G$2:$G$5',  true));   // Payment Status (col G, index 6, strict)
  reqs.push(dvBool(BUD, 3, 200, 9, 10));                   // Contract Signed (col J, index 9)

  // ── Timeline ──
  reqs.push(dv(TML, 3, 200, 4, 5, '$H$2:$H$11', false));  // Type (col E, index 4)
  reqs.push(dvBool(TML, 3, 200, 9, 10));                   // Done (col J, index 9)

  // ── Seating Chart ──
  reqs.push(dv(SEA, 3, 200, 2, 3, '$I$2:$I$5', false));   // Shape (col C, index 2)

  // ── Toast Planner ──
  reqs.push(dv(TST, 3, 200, 4, 5, '$J$2:$J$6', true));    // Status (col E, index 4, strict)

  // ── Photo List ──
  reqs.push(dv(PHO, 3, 200, 2, 3, '$K$2:$K$9', false));   // Shot Type (col C, index 2)
  reqs.push(dv(PHO, 3, 200, 3, 4, '$L$2:$L$4', true));    // Priority/Shot Status (col D, index 3, strict)
  reqs.push(dv(PHO, 3, 200, 6, 7, '$L$2:$L$4', true));    // Status (col G, index 6, strict)

  // ── Decor ──
  reqs.push(dv(DEC, 3, 200, 3, 4, '$M$2:$M$7', false));   // Category (col D, index 3)
  reqs.push({ setDataValidation: {                                            // Decor Status (col H, index 7)
    range: gridRange(DEC, 3, 200, 7, 8),
    rule: { condition: { type: 'ONE_OF_LIST', values: [
      { userEnteredValue: 'Ordered' }, { userEnteredValue: 'Confirmed' },
      { userEnteredValue: 'Pending' }, { userEnteredValue: 'Cancelled' },
    ]}, showCustomUi: true, strict: false },
  }});

  // ── Checklist ──
  reqs.push(dv(CHK, 3, 200, 2, 3, '$N$2:$N$8', false));   // Category (col C, index 2)
  reqs.push(dv(CHK, 3, 200, 3, 4, '$P$2:$P$4', true));    // Priority (col D, index 3, strict)
  reqs.push(dv(CHK, 3, 200, 4, 5, '$O$2:$O$6', true));    // Status (col E, index 4, strict)
  reqs.push(dvBool(CHK, 3, 200, 7, 8));                    // Done (col H, index 7)

  await batchUpdate(id, reqs, 'validation');
  console.log('Validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
