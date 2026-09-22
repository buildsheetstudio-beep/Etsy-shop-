'use strict';
// Fix script for Ultimate Kids Birthday Party Planner Dashboard
// Bug 1: KPI card data placement — labels and formulas written to same row in adjacent
//         columns, but the stacked card layout needs label row above and value row below.
//         Result: all 16 KPI cards (8 primary + 8 secondary) showed blank values.
//         Also: A6 (section header row) was overwritten with "Days Until Party" text.
// Bug 2: remainingBudget and costPerGuest referenced other dashboard cells (B10, D10, B9)
//         that were always empty — fixed by using direct source-data formulas.
// Bug 3: guestCapRemaining referenced B9 (always empty) — fixed with direct formula.
// Bug 4: cakeStatus used FD!B124 (Flavor, always filled) instead of FD!B128 (Order Date)
//         as the "ordered" indicator — fixed to use B128.
// Bug 5: Number formats (currency/percent) applied to wrong cells — fixed to match
//         the corrected value cell positions.
const { batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DB = sheetMap['Dashboard'];
const S = "'Dashboard'";

const GL = "'Guest List & RSVP'";
const BE = "'Budget & Expenses'";
const FD = "'Food, Drinks & Cake'";
const SL = "'Shopping List'";
const PC = "'Prep, Setup & Cleanup'";
const IT = "'Invitation Tracker'";

// All formulas use direct source references (no cross-dashboard cell refs)
const daysUntilParty    = `=IFERROR(MAX(0,DATEVALUE("August 15, 2026")-TODAY()),"—")`;
const confirmedGuests   = `=IFERROR(SUMIF(${GL}!$I$6:$I$205,"Confirmed",${GL}!$J$6:$J$205),0)`;
const pendingRsvps      = `=IFERROR(COUNTIF(${GL}!$I$6:$I$205,"No Response")+COUNTIF(${GL}!$I$6:$I$205,"Maybe"),0)`;
const totalBudget       = `=IFERROR('Party Setup'!B16,0)`;
const amountSpent       = `=IFERROR(SUM(${BE}!$I$31:$I$229),0)`;
const remainingBudget   = `=IFERROR('Party Setup'!B16-SUM(${BE}!$I$31:$I$229),0)`;
const costPerGuest      = `=IFERROR(IF(SUMIF(${GL}!$I$6:$I$205,"Confirmed",${GL}!$J$6:$J$205)=0,0,SUM(${BE}!$I$31:$I$229)/MAX(1,SUMIF(${GL}!$I$6:$I$205,"Confirmed",${GL}!$J$6:$J$205))),0)`;
const planningPct       = `=IFERROR(0.20*IFERROR(COUNTIF(${GL}!$H$6:$H$205,"Invitation Sent")/MAX(1,COUNTA(${GL}!$C$6:$C$205)),0)+0.20*IFERROR((COUNTIF(${GL}!$I$6:$I$205,"Confirmed")+COUNTIF(${GL}!$I$6:$I$205,"Declined"))/MAX(1,COUNTA(${GL}!$C$6:$C$205)),0)+0.15*IFERROR(IF(SUM(${BE}!$B$11:$B$25)>0,1,0),0)+0.20*IFERROR(COUNTIF(${SL}!$K$10:$K$209,TRUE)/MAX(1,COUNTA(${SL}!$C$10:$C$209)),0)+0.25*IFERROR(COUNTIF(${PC}!$J$10:$J$209,TRUE)/MAX(1,COUNTA(${PC}!$C$10:$C$209)),0),0)`;

const guestCapRemaining = `=IFERROR('Party Setup'!B17-SUMIF(${GL}!$I$6:$I$205,"Confirmed",${GL}!$J$6:$J$205),0)`;
const invNotSent        = `=IFERROR(COUNTIF(${GL}!$H$6:$H$205,"Invitation Ready")+COUNTIF(${GL}!$H$6:$H$205,"Not Invited"),0)`;
const followUpsNeeded   = `=IFERROR(COUNTIF(${IT}!$I$7:$I$206,"Yes"),0)`;
const dietaryRestrictions = `=IFERROR(SUMPRODUCT((${GL}!$I$6:$I$205="Confirmed")*(${GL}!$K$6:$K$205<>"None")*(${GL}!$K$6:$K$205<>"")),0)`;
const shoppingRemaining = `=IFERROR(COUNTA(${SL}!$C$10:$C$209)-COUNTIF(${SL}!$K$10:$K$209,TRUE),0)`;
const prepTasksRemaining= `=IFERROR(COUNTA(${PC}!$C$10:$C$209)-COUNTIF(${PC}!$J$10:$J$209,TRUE),0)`;
const overdueTasks      = `=IFERROR(COUNTIF(${PC}!$M$10:$M$209,"Yes"),0)`;
// Bug 4 fix: was FD!B124 (Flavor — always non-empty), now FD!B128 (Order Date)
const cakeStatus        = `=IFERROR(IF(${FD}!B136="Yes","Paid",IF(${FD}!B128<>"","Ordered","Not Ordered")),"—")`;

(async () => {
  const data = [];

  // ── FIX 1 & 2: Primary KPI cards ─────────────────────────────────────────
  // Correct layout: label in first col of each 2-col card span, formula one row below.
  // First card row: label row = A1 row 7 (index 6), value row = A1 row 8 (index 7)
  // Second card row: label row = A1 row 10 (index 9), value row = A1 row 11 (index 10)
  // Columns: card 0→A, card 1→C, card 2→E, card 3→G

  const primaryKPIs = [
    ['Days Until Party',  daysUntilParty ],   // i=0 → A col
    ['Confirmed Guests',  confirmedGuests],    // i=1 → C col
    ['Pending RSVPs',     pendingRsvps   ],    // i=2 → E col
    ['Total Budget',      totalBudget    ],    // i=3 → G col
    ['Amount Spent',      amountSpent    ],    // i=4 → A col (second row)
    ['Remaining Budget',  remainingBudget],    // i=5 → C col
    ['Cost per Guest',    costPerGuest   ],    // i=6 → E col
    ['Planning Progress', planningPct    ],    // i=7 → G col
  ];

  primaryKPIs.forEach(([label, formula], i) => {
    const offset = i < 4 ? i : i - 4;
    const col = String.fromCharCode(65 + offset * 2); // A, C, E, G
    const labelRow = i < 4 ? 7 : 10;   // A1 rows (index 6 and 9)
    const valRow   = i < 4 ? 8 : 11;   // A1 rows (index 7 and 10)
    data.push({ range: `${S}!${col}${labelRow}`, values: [[label]]   });
    data.push({ range: `${S}!${col}${valRow}`,   values: [[formula]] });
  });

  // Clear A6 which was incorrectly overwritten with the first KPI label
  data.push({ range: `${S}!A6`, values: [['']] });

  // ── FIX 3: Secondary KPI cards ───────────────────────────────────────────
  // Labels were placed correctly (rows 13 and 16) but formulas went into adjacent
  // columns on the same row instead of the value row below (rows 14 and 17).

  const secondaryKPIs = [
    ['Guest Capacity Left',   guestCapRemaining  ], // i=0 → A col
    ['Invitations Not Sent',  invNotSent         ], // i=1 → C col
    ['Follow-Ups Needed',     followUpsNeeded    ], // i=2 → E col
    ['Dietary Restrictions',  dietaryRestrictions], // i=3 → G col
    ['Shopping Remaining',    shoppingRemaining  ], // i=4 → A col (second row)
    ['Prep Tasks Remaining',  prepTasksRemaining ], // i=5 → C col
    ['Overdue Tasks',         overdueTasks       ], // i=6 → E col
    ['Cake Status',           cakeStatus         ], // i=7 → G col (Bug 4 fixed above)
  ];

  secondaryKPIs.forEach(([label, formula], i) => {
    const offset = i < 4 ? i : i - 4;
    const col = String.fromCharCode(65 + offset * 2);
    const labelRow = i < 4 ? 13 : 16;  // A1 rows (index 12 and 15)
    const valRow   = i < 4 ? 14 : 17;  // A1 rows (index 13 and 16)
    data.push({ range: `${S}!${col}${labelRow}`, values: [[label]]   });
    data.push({ range: `${S}!${col}${valRow}`,   values: [[formula]] });
  });

  await valuesBatchUpdate(id, data, 'fix-birthday-dashboard-values');

  // ── FIX 5: Number formats at correct value cell positions ─────────────────
  // Original applied formats to wrong cells (D8, F8, H8, B11, H11 — all within
  // merges or off-position). Apply to the correct value cells.
  const currF = { type: 'CURRENCY', pattern: '$#,##0.00' };
  const pctF  = { type: 'PERCENT',  pattern: '0%' };
  const reqs  = [];

  // Primary row 1 value cells (A1 row 8 = index 7):
  // A8=Days, C8=Confirmed, E8=Pending, G8=Total Budget (currency)
  reqs.push({ repeatCell: { range: gridRange(DB,7,8,6,7), cell: { userEnteredFormat: { numberFormat: currF } }, fields: 'userEnteredFormat.numberFormat' } }); // G8: Total Budget

  // Primary row 2 value cells (A1 row 11 = index 10):
  // A11=Amount Spent (currency), C11=Remaining Budget (currency), E11=Cost/Guest (currency), G11=Planning % (percent)
  reqs.push({ repeatCell: { range: gridRange(DB,10,11,0,1), cell: { userEnteredFormat: { numberFormat: currF } }, fields: 'userEnteredFormat.numberFormat' } }); // A11: Amount Spent
  reqs.push({ repeatCell: { range: gridRange(DB,10,11,2,3), cell: { userEnteredFormat: { numberFormat: currF } }, fields: 'userEnteredFormat.numberFormat' } }); // C11: Remaining Budget
  reqs.push({ repeatCell: { range: gridRange(DB,10,11,4,5), cell: { userEnteredFormat: { numberFormat: currF } }, fields: 'userEnteredFormat.numberFormat' } }); // E11: Cost per Guest
  reqs.push({ repeatCell: { range: gridRange(DB,10,11,6,7), cell: { userEnteredFormat: { numberFormat: pctF  } }, fields: 'userEnteredFormat.numberFormat' } }); // G11: Planning %

  // Clear old wrong number formats (D8, F8, H8, B11, H11) — set to no specific format
  const noFmt = { type: 'TEXT' };
  for (const [r1, r2, c1, c2] of [[7,8,3,4],[7,8,5,6],[7,8,7,8],[10,11,1,2],[10,11,7,8]]) {
    reqs.push({ repeatCell: { range: gridRange(DB,r1,r2,c1,c2), cell: { userEnteredFormat: { numberFormat: noFmt } }, fields: 'userEnteredFormat.numberFormat' } });
  }

  await batchUpdate(id, reqs, 'fix-birthday-dashboard-format');

  console.log('Birthday Planner Dashboard fixes applied:');
  console.log('  ✓ Primary KPIs: all 8 labels + formulas placed at correct stacked positions');
  console.log('  ✓ A6 section header cleared (was incorrectly showing "Days Until Party")');
  console.log('  ✓ Secondary KPIs: all 8 formulas moved to value rows (was in adjacent cols)');
  console.log('  ✓ remainingBudget: now uses direct Party Setup B16 minus SUM(BE col I)');
  console.log('  ✓ costPerGuest: now uses direct source-data formula');
  console.log('  ✓ guestCapRemaining: now uses direct source-data formula');
  console.log('  ✓ cakeStatus: fixed to check FD!B128 (Order Date) not FD!B124 (Flavor)');
  console.log('  ✓ Number formats: currency/percent now applied to correct value cells');
  console.log('Spreadsheet:', `https://docs.google.com/spreadsheets/d/${id}/edit`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
