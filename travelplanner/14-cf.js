const { hex, COLOR, gridRange, batchUpdate } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SIDs = {
  tripOverview: 732094577,
  itinerary: 353766069,
  budget: 1601326951,
  costPerPerson: 1289386811,
  currency: 1614915889,
  reservations: 128079104,
  sightseeing: 2144247023,
  foodDiary: 1291844518,
  visa: 1488564375,
  packing: 1128916394,
  reference: 2028483192,
};

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

  // ── Travel Budget tab ──
  // Over budget rows (Status col I = "Over Budget") — row highlight red
  reqs.push(cfRule(
    [gridRange(SIDs.budget, 12, 212, 0, 12)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$I13="Over Budget"' }] },
    { backgroundColor: hex(COLOR.lightRed), textFormat: { foregroundColor: hex(COLOR.deepCrimson) } }
  ));
  // Under budget rows — green
  reqs.push(cfRule(
    [gridRange(SIDs.budget, 12, 212, 0, 12)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$I13="Under Budget"' }] },
    { backgroundColor: hex(COLOR.lightGreen), textFormat: { foregroundColor: hex(COLOR.forestGreen) } }
  ));
  // Difference positive (over budget) — red text
  reqs.push(cfRule(
    [gridRange(SIDs.budget, 12, 212, 6, 7)],
    { type: 'NUMBER_GREATER', values: [{ userEnteredValue: '0' }] },
    { textFormat: { foregroundColor: hex(COLOR.deepCrimson), bold: true } }
  ));
  // Difference negative (under budget) — green text
  reqs.push(cfRule(
    [gridRange(SIDs.budget, 12, 212, 6, 7)],
    { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
    { textFormat: { foregroundColor: hex(COLOR.forestGreen), bold: true } }
  ));

  // ── Research & Reservations tab ──
  // Confirmed — green
  reqs.push(cfRule(
    [gridRange(SIDs.reservations, 2, 102, 0, 11)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$H3="Confirmed"' }] },
    { backgroundColor: hex(COLOR.lightGreen) }
  ));
  // Pending — amber
  reqs.push(cfRule(
    [gridRange(SIDs.reservations, 2, 102, 0, 11)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$H3="Pending"' }] },
    { backgroundColor: hex(COLOR.lightAmber) }
  ));
  // Cancelled — red
  reqs.push(cfRule(
    [gridRange(SIDs.reservations, 2, 102, 0, 11)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$H3="Cancelled"' }] },
    { backgroundColor: hex(COLOR.lightRed) }
  ));
  // Status cell text formatting
  reqs.push(cfRule(
    [gridRange(SIDs.reservations, 2, 102, 7, 8)],
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Confirmed' }] },
    { textFormat: { foregroundColor: hex(COLOR.forestGreen), bold: true } }
  ));
  reqs.push(cfRule(
    [gridRange(SIDs.reservations, 2, 102, 7, 8)],
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Pending' }] },
    { textFormat: { foregroundColor: hex(COLOR.amber), bold: true } }
  ));
  reqs.push(cfRule(
    [gridRange(SIDs.reservations, 2, 102, 7, 8)],
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Cancelled' }] },
    { textFormat: { foregroundColor: hex(COLOR.deepCrimson), bold: true } }
  ));

  // ── Sightseeing tab ──
  // Booked (col I = 8) = TRUE — green tick
  reqs.push(cfRule(
    [gridRange(SIDs.sightseeing, 2, 102, 8, 9)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$I3=TRUE' }] },
    { textFormat: { foregroundColor: hex(COLOR.forestGreen), bold: true } }
  ));
  // Must See (col J = 9) = TRUE — gold row highlight
  reqs.push(cfRule(
    [gridRange(SIDs.sightseeing, 2, 102, 0, 11)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$J3=TRUE' }] },
    { backgroundColor: hex(COLOR.gold20) }
  ));
  // Rating 5 — gold
  reqs.push(cfRule(
    [gridRange(SIDs.sightseeing, 2, 102, 7, 8)],
    { type: 'NUMBER_EQ', values: [{ userEnteredValue: '5' }] },
    { backgroundColor: hex(COLOR.lightGold), textFormat: { bold: true, foregroundColor: hex(COLOR.amber) } }
  ));
  // Rating ≤ 2 — light red
  reqs.push(cfRule(
    [gridRange(SIDs.sightseeing, 2, 102, 7, 8)],
    { type: 'NUMBER_LESS_THAN_EQ', values: [{ userEnteredValue: '2' }] },
    { backgroundColor: hex(COLOR.lightRed), textFormat: { foregroundColor: hex(COLOR.deepCrimson) } }
  ));

  // ── Food Diary tab ──
  // Must Return = TRUE — gold highlight
  reqs.push(cfRule(
    [gridRange(SIDs.foodDiary, 2, 102, 0, 11)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$I3=TRUE' }] },
    { backgroundColor: hex(COLOR.gold20) }
  ));
  // Rating 5 — gold
  reqs.push(cfRule(
    [gridRange(SIDs.foodDiary, 2, 102, 7, 8)],
    { type: 'NUMBER_EQ', values: [{ userEnteredValue: '5' }] },
    { backgroundColor: hex(COLOR.lightGold), textFormat: { bold: true, foregroundColor: hex(COLOR.amber) } }
  ));
  // High cost (> £50) — light red highlight on cost cell
  reqs.push(cfRule(
    [gridRange(SIDs.foodDiary, 2, 102, 6, 7)],
    { type: 'NUMBER_GREATER', values: [{ userEnteredValue: '50' }] },
    { textFormat: { foregroundColor: hex(COLOR.deepCrimson), bold: true } }
  ));

  // ── Visa tab ──
  // No Visa Required — green
  reqs.push(cfRule(
    [gridRange(SIDs.visa, 3, 33, 0, 11)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$B4="No Visa Required"' }] },
    { backgroundColor: hex(COLOR.lightGreen) }
  ));
  // Visa Required — amber
  reqs.push(cfRule(
    [gridRange(SIDs.visa, 3, 33, 0, 11)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$B4="Embassy Visa Required"' }] },
    { backgroundColor: hex(COLOR.lightAmber) }
  ));
  // E-Visa — light blue
  reqs.push(cfRule(
    [gridRange(SIDs.visa, 3, 33, 0, 11)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$B4="E-Visa / eTA"' }] },
    { backgroundColor: hex(COLOR.lightBlue) }
  ));

  // ── Packing Checklist ──
  // Packed items (D col = TRUE) — strikethrough + muted color
  reqs.push(cfRule(
    [gridRange(SIDs.packing, 3, 44, 0, 8)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$D4=TRUE' }] },
    { textFormat: { strikethrough: true, foregroundColor: hex(COLOR.midGrey) } }
  ));
  reqs.push(cfRule(
    [gridRange(SIDs.packing, 3, 44, 3, 4)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$D4=TRUE' }] },
    { backgroundColor: hex(COLOR.lightGreen), textFormat: { foregroundColor: hex(COLOR.forestGreen), bold: true } }
  ));

  // ── Itinerary tab ──
  // Daily budget over actual — green difference
  reqs.push(cfRule(
    [gridRange(SIDs.itinerary, 2, 24, 11, 12)],
    { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
    { textFormat: { foregroundColor: hex(COLOR.deepCrimson), bold: true } }
  ));
  reqs.push(cfRule(
    [gridRange(SIDs.itinerary, 2, 24, 11, 12)],
    { type: 'NUMBER_GREATER', values: [{ userEnteredValue: '0' }] },
    { textFormat: { foregroundColor: hex(COLOR.forestGreen), bold: true } }
  ));

  // ── Trip Overview ──
  // Destination rows — alternate color enhancement (already done with bg, so just highlight the nights cell)
  reqs.push(cfRule(
    [gridRange(SIDs.tripOverview, 19, 24, 0, 9)],
    { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$B20<>""' }] },
    { textFormat: { foregroundColor: hex(COLOR.deepNavy) } }
  ));

  await batchUpdate(id, reqs, 'conditional-formatting');
  console.log('Conditional formatting applied:', reqs.length, 'rules');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
