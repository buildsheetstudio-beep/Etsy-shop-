'use strict';
const { C, hex, batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['🎉 Dashboard'];
const GST = sheetMap['👥 Guest List & RSVP'];
const BUD = sheetMap['💰 Budget & Vendor Tracker'];
const TML = sheetMap['⏰ Day-Of Timeline'];
const SEA = sheetMap['🪑 Seating Chart'];
const TST = sheetMap['🥂 Toast & Speech Planner'];
const PHO = sheetMap['📸 Photo & Shot List'];
const DEC = sheetMap['🌸 Decor, Food & Drinks'];
const CHK = sheetMap['✅ Party Checklist'];

function cfBoolean(sheetId, r1, r2, c1, c2, formula, bgHex, textHex) {
  return { addConditionalFormatRule: { rule: {
    ranges: [gridRange(sheetId, r1, r2, c1, c2)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: formula }] },
      format: { backgroundColor: hex(bgHex), textFormat: { foregroundColor: hex(textHex) } },
    },
  }, index: 0 }};
}

(async () => {
  const reqs = [];

  // ══════ GUEST LIST ══════
  // Full row: Attending → light green
  reqs.push(cfBoolean(GST, 3, 200, 0, 14, '=$F4="Attending"', '#D4EDDA', C.bodyText));
  // Full row: Declined → light red/pink
  reqs.push(cfBoolean(GST, 3, 200, 0, 14, '=$F4="Declined"', '#F8D7DA', C.bodyText));
  // Full row: Maybe → light gold
  reqs.push(cfBoolean(GST, 3, 200, 0, 14, '=$F4="Maybe"', C.lightGold, C.bodyText));
  // RSVP cell (col F, index 5) status colors
  reqs.push(cfBoolean(GST, 3, 200, 5, 6, '=$F4="Attending"',     C.sageGreen,  C.white));
  reqs.push(cfBoolean(GST, 3, 200, 5, 6, '=$F4="Declined"',      C.mutedRust,  C.white));
  reqs.push(cfBoolean(GST, 3, 200, 5, 6, '=$F4="Maybe"',         C.mutedGold,  C.deepPlum));
  reqs.push(cfBoolean(GST, 3, 200, 5, 6, '=$F4="Awaiting Reply"',C.roseGold,   C.white));
  // Thank You Sent (col M=12) checkbox TRUE → green
  reqs.push(cfBoolean(GST, 3, 200, 12, 13, '=$M4=TRUE', C.sageGreen, C.white));

  // ══════ BUDGET ══════
  // Payment Status (col G=6) colors
  reqs.push(cfBoolean(BUD, 3, 200, 0, 13, '=$G4="Fully Paid"',   '#D4EDDA', C.bodyText));
  reqs.push(cfBoolean(BUD, 3, 200, 0, 13, '=$G4="Not Paid"',     '#F8D7DA', C.bodyText));
  reqs.push(cfBoolean(BUD, 3, 200, 6, 7,  '=$G4="Fully Paid"',   C.sageGreen,  C.white));
  reqs.push(cfBoolean(BUD, 3, 200, 6, 7,  '=$G4="Deposit Paid"', C.mutedGold,  C.deepPlum));
  reqs.push(cfBoolean(BUD, 3, 200, 6, 7,  '=$G4="Not Paid"',     C.mutedRust,  C.white));
  // Variance negative (col F=5) → red
  reqs.push(cfBoolean(BUD, 3, 200, 5, 6, '=$F4<0', '#F8D7DA', C.mutedRust));
  // Variance positive → light green
  reqs.push(cfBoolean(BUD, 3, 200, 5, 6, '=$F4>0', '#D4EDDA', C.sageGreen));
  // Contract signed (col J=9) → green
  reqs.push(cfBoolean(BUD, 3, 200, 9, 10, '=$J4=TRUE', C.sageGreen, C.white));

  // ══════ TIMELINE ══════
  // Done (col J=9) TRUE → full row green
  reqs.push(cfBoolean(TML, 3, 200, 0, 10, '=$J4=TRUE', '#D4EDDA', C.bodyText));
  // Type colors (col E=4)
  reqs.push(cfBoolean(TML, 3, 200, 4, 5, '=$E4="Setup"',          C.warmSand,   C.bodyText));
  reqs.push(cfBoolean(TML, 3, 200, 4, 5, '=$E4="Ceremony"',       C.deepPlum,   C.white));
  reqs.push(cfBoolean(TML, 3, 200, 4, 5, '=$E4="Catering"',       C.mutedGold,  C.deepPlum));
  reqs.push(cfBoolean(TML, 3, 200, 4, 5, '=$E4="Entertainment"',  C.roseGold,   C.white));
  reqs.push(cfBoolean(TML, 3, 200, 4, 5, '=$E4="Photography"',    C.mutedRust,  C.white));
  reqs.push(cfBoolean(TML, 3, 200, 4, 5, '=$E4="Cleanup"',        C.lightGray,  C.bodyText));

  // ══════ TOAST PLANNER ══════
  // Status colors (col E=4)
  reqs.push(cfBoolean(TST, 3, 200, 0, 8, '=$E4="Speech Ready"',  '#D4EDDA', C.bodyText));
  reqs.push(cfBoolean(TST, 3, 200, 0, 8, '=$E4="Pending"',       '#F8D7DA', C.bodyText));
  reqs.push(cfBoolean(TST, 3, 200, 4, 5, '=$E4="Speech Ready"',  C.sageGreen,  C.white));
  reqs.push(cfBoolean(TST, 3, 200, 4, 5, '=$E4="Confirmed"',     C.mutedGold,  C.deepPlum));
  reqs.push(cfBoolean(TST, 3, 200, 4, 5, '=$E4="Draft Written"', C.roseGold,   C.white));
  reqs.push(cfBoolean(TST, 3, 200, 4, 5, '=$E4="Pending"',       C.mutedRust,  C.white));
  reqs.push(cfBoolean(TST, 3, 200, 4, 5, '=$E4="Delivered"',     C.deepPlum,   C.white));

  // ══════ PHOTO LIST ══════
  // Priority / status colors
  reqs.push(cfBoolean(PHO, 3, 200, 3, 4, '=$D4="Must Have"',     C.mutedRust,  C.white));
  reqs.push(cfBoolean(PHO, 3, 200, 3, 4, '=$D4="Nice to Have"',  C.mutedGold,  C.deepPlum));
  reqs.push(cfBoolean(PHO, 3, 200, 6, 7, '=$G4="Completed"',     C.sageGreen,  C.white));
  // Full row completed → light green
  reqs.push(cfBoolean(PHO, 3, 200, 0, 8, '=$G4="Completed"',     '#D4EDDA', C.bodyText));

  // ══════ DECOR ══════
  // Status colors (col H=7)
  reqs.push(cfBoolean(DEC, 3, 200, 7, 8, '=$H4="Confirmed"', C.sageGreen,  C.white));
  reqs.push(cfBoolean(DEC, 3, 200, 7, 8, '=$H4="Ordered"',   C.mutedGold,  C.deepPlum));
  reqs.push(cfBoolean(DEC, 3, 200, 7, 8, '=$H4="Pending"',   C.mutedRust,  C.white));
  reqs.push(cfBoolean(DEC, 3, 200, 7, 8, '=$H4="Cancelled"', '#F8D7DA',   C.bodyText));

  // ══════ CHECKLIST ══════
  // Done (col H=7) TRUE → full row green + strikethrough
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(CHK, 3, 200, 0, 9)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$H4=TRUE' }] },
      format: { backgroundColor: hex('#D4EDDA'), textFormat: { strikethrough: true, foregroundColor: hex(C.bodyText) } },
    },
  }, index: 0 }});
  // Priority colors (col D=3)
  reqs.push(cfBoolean(CHK, 3, 200, 3, 4, '=$D4="High"',   C.mutedRust,  C.white));
  reqs.push(cfBoolean(CHK, 3, 200, 3, 4, '=$D4="Medium"', C.mutedGold,  C.deepPlum));
  reqs.push(cfBoolean(CHK, 3, 200, 3, 4, '=$D4="Low"',    C.sageGreen,  C.white));
  // Status colors (col E=4)
  reqs.push(cfBoolean(CHK, 3, 200, 4, 5, '=$E4="Complete"',     C.sageGreen,  C.white));
  reqs.push(cfBoolean(CHK, 3, 200, 4, 5, '=$E4="In Progress"',  C.mutedGold,  C.deepPlum));
  reqs.push(cfBoolean(CHK, 3, 200, 4, 5, '=$E4="Not Started"',  '#E0E0E0',   '#555555'));
  reqs.push(cfBoolean(CHK, 3, 200, 4, 5, '=$E4="Delegated"',    C.roseGold,   C.white));
  reqs.push(cfBoolean(CHK, 3, 200, 4, 5, '=$E4="Cancelled"',    C.mutedRust,  C.white));

  // ══════ DASHBOARD ══════
  // Budget rows: remaining negative → red
  reqs.push(cfBoolean(DSH, 12, 22, 3, 4, '=$D13<0', '#F8D7DA', C.mutedRust));

  await batchUpdate(id, reqs, 'cf');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
