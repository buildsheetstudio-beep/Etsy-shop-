'use strict';
const { batchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const CK = sheetMap['✅ Party Checklist'];
const GL = sheetMap['👥 Guest List & RSVPs'];
const BU = sheetMap['💰 Budget & Expenses'];
const GA = sheetMap['🎮 Games & Activities'];
const GF = sheetMap['🎁 Gift Tracker'];
const DC = sheetMap['🎀 Décor, Food & Drinks'];

function cfRule(ranges, condition, format) {
  return {
    addConditionalFormatRule: {
      rule: {
        ranges,
        booleanRule: { condition, format },
      },
      index: 0,
    },
  };
}

(async () => {
  const reqs = [];

  // ── CHECKLIST ──
  // Done rows (A=TRUE): strike-through + sage text on cols B-H
  reqs.push(cfRule(
    [gridRange(CK, 3, 48, 1, 8)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$A4=TRUE' }] },
    { textFormat: { foregroundColor: hex(C.sage), strikethrough: true } }
  ));
  // High priority col D — bold amber
  reqs.push(cfRule(
    [gridRange(CK, 3, 48, 3, 4)],
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'High' }] },
    { textFormat: { foregroundColor: hex(C.amber), bold: true } }
  ));
  // Status "Done" — sage
  reqs.push(cfRule(
    [gridRange(CK, 3, 48, 6, 7)],
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Done' }] },
    { textFormat: { foregroundColor: hex(C.sage), bold: true } }
  ));
  // Status "In Progress" — amber
  reqs.push(cfRule(
    [gridRange(CK, 3, 48, 6, 7)],
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'In Progress' }] },
    { textFormat: { foregroundColor: hex(C.amber), bold: true } }
  ));

  // ── GUEST LIST ──
  // RSVP Confirmed — sage
  reqs.push(cfRule(
    [gridRange(GL, 3, 23, 4, 5)],
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Confirmed' }] },
    { textFormat: { foregroundColor: hex(C.sage), bold: true } }
  ));
  // RSVP Declined — deepRose
  reqs.push(cfRule(
    [gridRange(GL, 3, 23, 4, 5)],
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Declined' }] },
    { textFormat: { foregroundColor: hex(C.deepRose), bold: true } }
  ));
  // RSVP No Response — amber
  reqs.push(cfRule(
    [gridRange(GL, 3, 23, 4, 5)],
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'No Response' }] },
    { textFormat: { foregroundColor: hex(C.amber), bold: true } }
  ));
  // Gift received = TRUE — sage on gift col K (index 10)
  reqs.push(cfRule(
    [gridRange(GL, 3, 23, 10, 11)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$J4=TRUE' }] },
    { textFormat: { foregroundColor: hex(C.sage) } }
  ));

  // ── BUDGET ──
  // Remaining < 0 (over budget) — deepRose bg on col E rows 4-15
  reqs.push(cfRule(
    [gridRange(BU, 3, 15, 4, 5)],
    { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
    { backgroundColor: hex(C.deepRose), textFormat: { foregroundColor: hex(C.white), bold: true } }
  ));
  // Paid = TRUE — sage text on expense row cols A-I
  reqs.push(cfRule(
    [gridRange(BU, 17, 32, 0, 9)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G18=TRUE' }] },
    { textFormat: { foregroundColor: hex(C.sage) } }
  ));
  // Variance negative (over budget) on col F expense rows
  reqs.push(cfRule(
    [gridRange(BU, 17, 32, 5, 6)],
    { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
    { textFormat: { foregroundColor: hex(C.deepRose), bold: true } }
  ));

  // ── GAMES ──
  // Status "Done" — sage
  reqs.push(cfRule(
    [gridRange(GA, 2, 14, 8, 9)],
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Done' }] },
    { textFormat: { foregroundColor: hex(C.sage), bold: true } }
  ));
  // Status "Planned" — amber
  reqs.push(cfRule(
    [gridRange(GA, 2, 14, 8, 9)],
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Planned' }] },
    { textFormat: { foregroundColor: hex(C.amber) } }
  ));

  // ── GIFT TRACKER ──
  // Received = TRUE — full row sage
  reqs.push(cfRule(
    [gridRange(GF, 3, 23, 0, 9)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F4=TRUE' }] },
    { textFormat: { foregroundColor: hex(C.sage) } }
  ));
  // TY NOT sent but gift received — amber on col G
  reqs.push(cfRule(
    [gridRange(GF, 3, 23, 6, 7)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($F4=TRUE,$G4=FALSE)' }] },
    { backgroundColor: hex('#FFF3CD'), textFormat: { foregroundColor: hex(C.amber), bold: true } }
  ));

  // ── DÉCOR ──
  // Status "Arrived"/"Done" — sage on décor status col G (index 6) rows 4-17
  reqs.push(cfRule(
    [gridRange(DC, 3, 17, 6, 7)],
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Arrived' }] },
    { textFormat: { foregroundColor: hex(C.sage), bold: true } }
  ));
  reqs.push(cfRule(
    [gridRange(DC, 3, 17, 6, 7)],
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Done' }] },
    { textFormat: { foregroundColor: hex(C.sage), bold: true } }
  ));
  // Food status "Confirmed" — sage on food status col G (index 6) rows 21-35
  reqs.push(cfRule(
    [gridRange(DC, 20, 35, 6, 7)],
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Confirmed' }] },
    { textFormat: { foregroundColor: hex(C.sage), bold: true } }
  ));

  await batchUpdate(id, reqs, 'cf-rules');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
