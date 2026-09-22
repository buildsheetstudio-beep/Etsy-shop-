const { gridRange, batchUpdate } = require('./lib');
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

function dropdownRange(sheetId, r0, r1, c0, c1, refRange, strict = true) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r0, r1, c0, c1),
      rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${refRange}` }] },
        strict,
        showCustomUi: true,
      },
    },
  };
}

function checkbox(sheetId, r0, r1, c0, c1) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r0, r1, c0, c1),
      rule: { condition: { type: 'BOOLEAN' }, strict: true, showCustomUi: true },
    },
  };
}

(async () => {
  const reqs = [];

  // ── Trip Overview ──
  // Traveller Type (E9 = row 8, col 4)
  reqs.push(dropdownRange(SIDs.tripOverview, 8, 9, 4, 5, `'Reference Data'!$N$3:$N$7`));
  // Currency (E8 = row 7, col 4)
  reqs.push(dropdownRange(SIDs.tripOverview, 7, 8, 4, 5, `'Reference Data'!$L$3:$L$11`));
  // Currency Converter quick converter: From and To
  reqs.push(dropdownRange(SIDs.currency, 16, 17, 3, 4, `'Reference Data'!$L$3:$L$11`)); // D17
  reqs.push(dropdownRange(SIDs.currency, 16, 17, 6, 7, `'Reference Data'!$L$3:$L$11`)); // G17

  // ── Travel Budget ──
  // Category (col B = 1, rows 12-211)
  reqs.push(dropdownRange(SIDs.budget, 12, 212, 1, 2, `'Reference Data'!$A$3:$A$11`));
  // Destination (col D = 3)
  reqs.push(dropdownRange(SIDs.budget, 12, 212, 3, 4, `'Reference Data'!$B$3:$B$18`));
  // Payment Method — inline list
  reqs.push({
    setDataValidation: {
      range: gridRange(SIDs.budget, 12, 212, 10, 11),
      rule: {
        condition: { type: 'ONE_OF_LIST', values: [
          { userEnteredValue: 'Credit Card' },
          { userEnteredValue: 'Debit Card' },
          { userEnteredValue: 'Cash' },
          { userEnteredValue: 'Bank Transfer' },
          { userEnteredValue: 'Mixed' },
        ]},
        strict: false, showCustomUi: true,
      },
    },
  });

  // ── Day-by-Day Itinerary ──
  // Destination (col D = 3, rows 2-23)
  reqs.push(dropdownRange(SIDs.itinerary, 2, 24, 3, 4, `'Reference Data'!$B$3:$B$18`));

  // ── Research & Reservations ──
  // Type (col B = 1, rows 2-101)
  reqs.push({
    setDataValidation: {
      range: gridRange(SIDs.reservations, 2, 102, 1, 2),
      rule: {
        condition: { type: 'ONE_OF_LIST', values: [
          { userEnteredValue: 'Accommodation' },
          { userEnteredValue: 'Transport' },
          { userEnteredValue: 'Activities' },
          { userEnteredValue: 'Tours' },
          { userEnteredValue: 'Restaurant' },
          { userEnteredValue: 'Other' },
        ]},
        strict: true, showCustomUi: true,
      },
    },
  });
  // Destination (col D = 3)
  reqs.push(dropdownRange(SIDs.reservations, 2, 102, 3, 4, `'Reference Data'!$B$3:$B$18`));
  // Status (col H = 7)
  reqs.push(dropdownRange(SIDs.reservations, 2, 102, 7, 8, `'Reference Data'!$D$3:$D$7`));

  // ── Sightseeing & Activities ──
  // Type (col C = 2)
  reqs.push(dropdownRange(SIDs.sightseeing, 2, 102, 2, 3, `'Reference Data'!$E$3:$E$14`));
  // Destination (col D = 3)
  reqs.push(dropdownRange(SIDs.sightseeing, 2, 102, 3, 4, `'Reference Data'!$B$3:$B$18`));
  // Booked? (col I = 8)
  reqs.push(checkbox(SIDs.sightseeing, 2, 102, 8, 9));
  // Must See? (col J = 9)
  reqs.push(checkbox(SIDs.sightseeing, 2, 102, 9, 10));
  // Rating 1-5 number validation
  reqs.push({
    setDataValidation: {
      range: gridRange(SIDs.sightseeing, 2, 102, 7, 8),
      rule: {
        condition: { type: 'NUMBER_BETWEEN', values: [{ userEnteredValue: '1' }, { userEnteredValue: '5' }] },
        strict: false, showCustomUi: false,
      },
    },
  });

  // ── Restaurant & Food Diary ──
  // Destination (col C = 2)
  reqs.push(dropdownRange(SIDs.foodDiary, 2, 102, 2, 3, `'Reference Data'!$B$3:$B$18`));
  // Meal Type (col E = 4)
  reqs.push(dropdownRange(SIDs.foodDiary, 2, 102, 4, 5, `'Reference Data'!$F$3:$F$11`));
  // Cuisine (col F = 5)
  reqs.push(dropdownRange(SIDs.foodDiary, 2, 102, 5, 6, `'Reference Data'!$G$3:$G$16`));
  // Must Return? (col I = 8)
  reqs.push(checkbox(SIDs.foodDiary, 2, 102, 8, 9));
  // Rating 1-5
  reqs.push({
    setDataValidation: {
      range: gridRange(SIDs.foodDiary, 2, 102, 7, 8),
      rule: {
        condition: { type: 'NUMBER_BETWEEN', values: [{ userEnteredValue: '1' }, { userEnteredValue: '5' }] },
        strict: false, showCustomUi: false,
      },
    },
  });

  // ── Visa tab ──
  // Visa Type (col C = 2, rows 3-32)
  reqs.push(dropdownRange(SIDs.visa, 3, 33, 2, 3, `'Reference Data'!$H$3:$H$7`));
  // Status (col J = 9) — inline
  reqs.push({
    setDataValidation: {
      range: gridRange(SIDs.visa, 3, 33, 9, 10),
      rule: {
        condition: { type: 'ONE_OF_LIST', values: [
          { userEnteredValue: 'Not Required' },
          { userEnteredValue: 'Applied' },
          { userEnteredValue: 'Approved' },
          { userEnteredValue: 'Rejected' },
          { userEnteredValue: 'Expired' },
        ]},
        strict: true, showCustomUi: true,
      },
    },
  });

  // ── Packing Checklist ──
  // Category (col B = 1, rows 3-53)
  reqs.push(dropdownRange(SIDs.packing, 3, 53, 1, 2, `'Reference Data'!$I$3:$I$10`));

  await batchUpdate(id, reqs, 'data-validation');
  console.log('Data validation applied:', reqs.length, 'rules');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
