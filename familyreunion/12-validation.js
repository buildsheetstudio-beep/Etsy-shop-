'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CKL = sheetMap['✅ Planning Checklist'];
const GST = sheetMap['👨‍👩‍👧‍👦 Guest List & RSVPs'];
const BDG = sheetMap['💰 Budget & Expenses'];
const ITN = sheetMap['🗓️ Event Itinerary & Activities'];
const FD  = sheetMap['🍖 Food, Drinks & Potluck'];
const ACC = sheetMap['🏠 Accommodation & Travel'];

const REF = "'📋 Reference Data'";

// ONE_OF_RANGE dropdown helper
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

(async () => {
  const reqs = [];

  // ── PLANNING CHECKLIST (data rows 0-indexed 3–47) ──
  reqs.push(dv(CKL, 3, 48, 1, 2,  '$A$3:$A$11'));  // B: Timeframe
  reqs.push(dv(CKL, 3, 48, 3, 4,  '$C$3:$C$5'));   // D: Priority
  reqs.push(dv(CKL, 3, 48, 6, 7,  '$B$3:$B$6'));   // G: Task Status

  // ── GUEST LIST (data rows 0-indexed 3–27) ──
  reqs.push(dv(GST, 3, 28, 1, 2,  '$U$3:$U$8'));   // B: Family Branch
  reqs.push(dv(GST, 3, 28, 3, 4,  '$D$3:$D$7'));   // D: Age Group
  reqs.push(dv(GST, 3, 28, 7, 8,  '$E$3:$E$7', true)); // H: RSVP Status (strict)
  reqs.push(dv(GST, 3, 28, 9, 10, '$F$3:$F$11'));  // J: Dietary

  // ── BUDGET EXPENSES (data rows 0-indexed 18–32) ──
  reqs.push(dv(BDG, 18, 33, 0, 1, '$T$3:$T$14'));  // A: Budget Category

  // ── EVENT ITINERARY (data rows 0-indexed 3–22) ──
  reqs.push(dv(ITN, 3, 23, 0, 1,  '$P$3:$P$7'));   // A: Day
  reqs.push(dv(ITN, 3, 23, 5, 6,  '$O$3:$O$13'));  // F: Itinerary Type

  // ── ACTIVITIES (rows 0-indexed 25–36) ──
  reqs.push(dv(ITN, 25, 37, 1, 2, '$N$3:$N$11'));  // B: Activity Type
  reqs.push(dv(ITN, 25, 37, 3, 4, '$M$3:$M$7'));   // D: Activity Ages
  reqs.push(dv(ITN, 25, 37, 6, 7, '$L$3:$L$6'));   // G: Activity Status

  // ── FOOD — MENU (rows 0-indexed 3–24) ──
  reqs.push(dv(FD, 3, 25, 0, 1,  '$Q$3:$Q$12'));  // A: Meal Slot
  reqs.push(dv(FD, 3, 25, 2, 3,  '$G$3:$G$11'));  // C: Food Type
  reqs.push(dv(FD, 3, 25, 3, 4,  '$H$3:$H$7'));   // D: Who's Bringing
  reqs.push(dv(FD, 3, 25, 5, 6,  '$S$3:$S$9'));   // F: Dietary Tags
  reqs.push(dv(FD, 3, 25, 6, 7,  '$R$3:$R$7'));   // G: Food Status

  // ── FOOD — POTLUCK (rows 0-indexed 28–37) ──
  reqs.push(dv(FD, 28, 38, 2, 3, '$Q$3:$Q$12'));  // C: Meal Slot
  reqs.push(dv(FD, 28, 38, 3, 4, '$S$3:$S$9'));   // D: Dietary Tags

  // ── ACCOMMODATION (rows 0-indexed 3–10) ──
  reqs.push(dv(ACC, 3, 11, 1, 2,  '$I$3:$I$9'));   // B: Accommodation Type
  reqs.push(dv(ACC, 3, 11, 7, 8,  '$J$3:$J$6'));   // H: Booking Status

  // ── TRAVEL (rows 0-indexed 14–21) ──
  reqs.push(dv(ACC, 14, 22, 6, 7, '$K$3:$K$7'));   // G: Transport

  await batchUpdate(id, reqs, 'validation');
  console.log('Validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
