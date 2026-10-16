'use strict';
const { batchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH = sheetMap['🥂 Dashboard'];
const GST  = sheetMap['👰 Guest List & Seating'];
const BUD  = sheetMap['💰 Budget & Expenses'];
const ITN  = sheetMap['🗓️ Party Itinerary & Games'];
const VEN  = sheetMap['💍 Venue, Vendors & Bridesmaids'];
const GFT  = sheetMap['🎁 Gift Tracker & Wishlist'];

function cf(sheetId, r1, r2, c1, c2, formula, bgHex, fgHex, bold) {
  return { addConditionalFormatRule: { rule: {
    ranges: [gridRange(sheetId, r1, r2, c1, c2)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: formula }] },
      format: {
        backgroundColor: hex(bgHex),
        textFormat: { foregroundColor: hex(fgHex || C.text), bold: !!bold },
      },
    },
  }, index: 0 }};
}

(async () => {
  const reqs = [];

  // ── Guest List: RSVP status colours (col E = index 4) ──
  // Confirmed → sage bg
  reqs.push(cf(GST, 3, 23, 0, 13, '=$E4="Confirmed"',    '#E8F5E4', C.sage));
  // Declined → deep rose bg
  reqs.push(cf(GST, 3, 23, 0, 13, '=$E4="Declined"',     '#F9EBEC', C.deepRose));
  // No Response → amber bg
  reqs.push(cf(GST, 3, 23, 0, 13, '=$E4="No Response"',  '#FDF6EC', C.amber));
  // Invited → parchment
  reqs.push(cf(GST, 3, 23, 0, 13, '=$E4="Invited"',      C.parchment, C.muted));

  // Thank you NOT sent (col K = index 10) where gift received (col J = index 9)
  reqs.push(cf(GST, 3, 23, 10, 11, '=AND($J4="Yes",$K4<>"Yes")', '#FFF3CD', C.amber, true));

  // ── Budget: overspent rows (Remaining < 0, col D = index 3) ──
  reqs.push(cf(BUD, 3, 15, 0, 6, '=$D4<0', '#F9EBEC', C.deepRose));
  // On budget (0..10% over): highlight remaining col green if positive
  reqs.push(cf(BUD, 3, 15, 3, 4, '=$D4>0', '#E8F5E4', C.sage));
  // Paid checkbox = TRUE: sage row in expense log (col G = index 6)
  reqs.push(cf(BUD, 19, 34, 0, 9, '=$G20=TRUE', '#F2FBF2', C.sage));
  // Negative variance in expense log (col F = index 5): highlight red
  reqs.push(cf(BUD, 19, 34, 5, 6, '=$F20<0', '#F9EBEC', C.deepRose, false));

  // ── Itinerary: TIME CHECK warning (col H = index 7) ──
  reqs.push(cf(ITN, 4, 18, 7, 8, '=$H5="⚠ OVERLAP"', '#F9EBEC', C.deepRose, true));

  // ── Vendors / Bridesmaids: task status colours (col F = index 5, rows 19-36) ──
  reqs.push(cf(VEN, 19, 37, 0, 9, '=$F20="Done"',        '#E8F5E4', C.sage));
  reqs.push(cf(VEN, 19, 37, 0, 9, '=$F20="In Progress"', '#FDF6EC', C.amber));
  reqs.push(cf(VEN, 19, 37, 0, 9, '=$F20="Not Started"', C.ivory,   C.muted));
  reqs.push(cf(VEN, 19, 37, 0, 9, '=$F20="Delegated"',   '#F0EFF9', '#5E5A8B'));

  // Contract signed? → col J (index 9) in vendor comparison (rows 3-14)
  reqs.push(cf(VEN, 3, 15, 0, 9, '=$J4="Yes"', '#E8F5E4', C.sage));

  // ── Gifts: Thank You status (col H = index 7, rows 22-50) ──
  reqs.push(cf(GFT, 22, 50, 0, 9, '=$H23="Thank You Sent"', '#E8F5E4', C.sage));
  reqs.push(cf(GFT, 22, 50, 0, 9, '=$H23="Pending"',        '#FDF6EC', C.amber));
  reqs.push(cf(GFT, 22, 50, 0, 9, '=$H23="Not Started"',    '#F9EBEC', C.deepRose));

  // Wishlist priority highlights (col C = index 2, rows 3-17)
  reqs.push(cf(GFT, 3, 18, 0, 6, '=$C4="Must Have"',    '#FFF0F3', C.deepRose, true));
  reqs.push(cf(GFT, 3, 18, 0, 6, '=$C4="Would Love"',   '#FDF6EC', C.amber));
  reqs.push(cf(GFT, 3, 18, 0, 6, '=$C4="Nice to Have"', C.parchment, C.muted));

  await batchUpdate(id, reqs, 'cf-rules');
  console.log('Conditional Formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
