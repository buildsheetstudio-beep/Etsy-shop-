'use strict';
const { batchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CKL = sheetMap['✅ Planning Checklist'];
const GST = sheetMap['👨‍👩‍👧‍👦 Guest List & RSVPs'];
const BDG = sheetMap['💰 Budget & Expenses'];
const ITN = sheetMap['🗓️ Event Itinerary & Activities'];
const ACC = sheetMap['🏠 Accommodation & Travel'];

function cfBool(sheetId, r1, r2, c1, c2, formula, fg, bg, idx) {
  return { addConditionalFormatRule: {
    rule: {
      ranges: [{ sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 }],
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: formula }] },
        format: { textFormat: { foregroundColor: hex(fg) }, backgroundColor: hex(bg) },
      },
    },
    index: idx || 0,
  }};
}

function cfText(sheetId, r1, r2, c1, c2, text, fg, bg, idx) {
  return { addConditionalFormatRule: {
    rule: {
      ranges: [{ sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 }],
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] },
        format: { textFormat: { foregroundColor: hex(fg) }, backgroundColor: hex(bg) },
      },
    },
    index: idx || 0,
  }};
}

function cfNum(sheetId, r1, r2, c1, c2, type, val, fg, bg, idx) { // type: 'NUMBER_LESS' | 'NUMBER_GREATER'
  return { addConditionalFormatRule: {
    rule: {
      ranges: [{ sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 }],
      booleanRule: {
        condition: { type, values: [{ userEnteredValue: String(val) }] },
        format: { textFormat: { foregroundColor: hex(fg) }, backgroundColor: hex(bg) },
      },
    },
    index: idx || 0,
  }};
}

(async () => {
  const reqs = [];

  // ── PLANNING CHECKLIST ──
  // Done rows dimmed (A4:H200, $A=TRUE)
  reqs.push(cfBool(CKL, 3, 200, 0, 8, '=$A4=TRUE', '#AAAAAA', '#F0F5EE', 0));
  // Overdue (date past, not done): F col red
  reqs.push(cfBool(CKL, 3, 200, 5, 6, '=AND($A4=FALSE,$F4<TODAY(),$F4<>"")', C.burnSienna, '#FDE8D8', 1));
  // Status = Done: green
  reqs.push(cfText(CKL, 3, 200, 6, 7, 'Done', C.forestGreen, '#D8EED8', 2));
  // Priority = High: amber
  reqs.push(cfText(CKL, 3, 200, 3, 4, 'High', '#8B5A00', '#FEF5E8', 3));

  // ── GUEST LIST ──
  // Family branch colours on col B (index 1)
  const branchCF = [
    ['Henderson Side', '#FDE8D0', '#8B3A1A'],
    ['Thompson Side',  '#D0E8FD', '#1A3A8B'],
    ['Wilson Side',    '#D0FDE8', '#1A8B4A'],
    ['Smith Side',     '#FDD0E8', '#8B1A5A'],
    ['Out-of-Town',    '#E8D0FD', '#5A1A8B'],
    ['Family Friends', '#FDFFE0', '#6A6A00'],
  ];
  branchCF.forEach(([branch, bg, fg], i) => {
    reqs.push(cfText(GST, 3, 200, 1, 2, branch, fg, bg, i));
  });
  // RSVP Confirmed: green (col H = index 7)
  reqs.push(cfText(GST, 3, 200, 7, 8, 'Confirmed', C.forestGreen, '#D8EED8', 6));
  // RSVP Declined: red
  reqs.push(cfText(GST, 3, 200, 7, 8, 'Declined', C.burnSienna, '#FDE8D8', 7));
  // RSVP No Response: amber
  reqs.push(cfText(GST, 3, 200, 7, 8, 'No Response', '#8B5A00', '#FEF5E8', 8));

  // ── BUDGET ──
  // Variance col G (index 6) on expense rows (0-indexed 18+): negative = red, positive = green
  reqs.push(cfBool(BDG, 18, 200, 6, 7, '=$G19<0', C.burnSienna, '#FDE8D8', 0));
  reqs.push(cfBool(BDG, 18, 200, 6, 7, '=$G19>0', C.forestGreen, '#D8EED8', 1));
  // Paid = FALSE: amber (col H = index 7)
  reqs.push(cfBool(BDG, 18, 200, 7, 8, '=$H19=FALSE', '#8B5A00', '#FEF5E8', 2));

  // ── ITINERARY ──
  // Time overlap col J (index 9): red bold
  reqs.push({ addConditionalFormatRule: {
    rule: {
      ranges: [{ sheetId: ITN, startRowIndex: 3, endRowIndex: 23, startColumnIndex: 9, endColumnIndex: 10 }],
      booleanRule: {
        condition: { type: 'TEXT_CONTAINS', values: [{ userEnteredValue: 'OVERLAP' }] },
        format: { textFormat: { foregroundColor: hex(C.burnSienna), bold: true }, backgroundColor: hex('#FDE8D8') },
      },
    },
    index: 0,
  }});

  // ── ACCOMMODATION ──
  // Booking Status col H (index 7)
  reqs.push(cfText(ACC, 3, 11, 7, 8, 'Fully Booked',  C.forestGreen, '#D8EED8', 0));
  reqs.push(cfText(ACC, 3, 11, 7, 8, 'Not Booked',    C.burnSienna,  '#FDE8D8', 1));
  reqs.push(cfText(ACC, 3, 11, 7, 8, 'Deposit Paid',  '#8B5A00',     '#FEF5E8', 2));
  reqs.push(cfText(ACC, 3, 11, 7, 8, 'Cancelled',     C.burnSienna,  '#FDE8D8', 3));

  await batchUpdate(id, reqs, 'cf-rules');
  console.log('Conditional Formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
