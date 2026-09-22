'use strict';
const { batchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['🎓 Dashboard'];
const TRK = sheetMap['📋 Application Tracker'];
const DL  = sheetMap['⏰ Deadlines & To-Do'];
const RES = sheetMap['🏆 Results & Awards Log'];

function cfText(sheetId, r1, r2, c1, c2, text, fg, bg, idx, bold) {
  return { addConditionalFormatRule: {
    rule: {
      ranges: [{ sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 }],
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] },
        format: { textFormat: { foregroundColor: hex(fg), bold: !!bold }, backgroundColor: hex(bg) },
      },
    },
    index: idx || 0,
  }};
}

function cfFormula(sheetId, r1, r2, c1, c2, formula, fg, bg, idx, bold) {
  return { addConditionalFormatRule: {
    rule: {
      ranges: [{ sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 }],
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: formula }] },
        format: { textFormat: { foregroundColor: hex(fg), bold: !!bold }, backgroundColor: hex(bg) },
      },
    },
    index: idx || 0,
  }};
}

function cfNum(sheetId, r1, r2, c1, c2, type, val, fg, bg, idx, bold) {
  return { addConditionalFormatRule: {
    rule: {
      ranges: [{ sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 }],
      booleanRule: {
        condition: { type, values: [{ userEnteredValue: String(val) }] },
        format: { textFormat: { foregroundColor: hex(fg), bold: !!bold }, backgroundColor: hex(bg) },
      },
    },
    index: idx || 0,
  }};
}

(async () => {
  const reqs = [];

  // ── APPLICATION TRACKER CF (TRK, data rows 3-27, 0-indexed) ──
  // Full row: Awarded → very pale green
  reqs.push(cfFormula(TRK, 3, 28, 0, 15, '=$G4="Awarded"',  '#1A6B3A', '#F0FBF4', 0));
  // Full row: Rejected → very pale red
  reqs.push(cfFormula(TRK, 3, 28, 0, 15, '=$G4="Rejected"', '#8B1A1A', '#FBF0F0', 1));
  // Status cell rules (priority over full-row rules — higher index = higher priority in Sheets)
  reqs.push(cfText(TRK,  3, 28, 6, 7, 'Awarded',    '#1A6B3A', '#D4EDDA', 2, true));
  reqs.push(cfText(TRK,  3, 28, 6, 7, 'Submitted',  '#1A3A8B', '#D6E8FF', 3));
  reqs.push(cfText(TRK,  3, 28, 6, 7, 'In Progress','#7A6000', '#FFF8D4', 4));
  reqs.push(cfText(TRK,  3, 28, 6, 7, 'Rejected',   '#8B1A1A', '#F8D7DA', 5));
  reqs.push(cfText(TRK,  3, 28, 6, 7, 'Waitlisted', '#6A5A8B', '#EDE8F8', 6));
  reqs.push(cfText(TRK,  3, 28, 6, 7, 'Withdrawn',  '#888888', '#F0F0F0', 7));
  // Deadline urgency (col F = index 5): ≤7 days → red bold, ≤30 days → gold
  reqs.push(cfFormula(TRK, 3, 28, 5, 6, '=AND(F4<TODAY()+7,F4>=TODAY(),G4<>"Submitted",G4<>"Awarded",G4<>"Rejected",G4<>"Withdrawn")',  '#8B1A1A', '#F8D7DA', 8, true));
  reqs.push(cfFormula(TRK, 3, 28, 5, 6, '=AND(F4<TODAY()+30,F4>=TODAY(),G4<>"Submitted",G4<>"Awarded",G4<>"Rejected",G4<>"Withdrawn")', '#7A6000', '#FFF8D4', 9));
  // Priority High → gold bg on col H (index 7)
  reqs.push(cfText(TRK, 3, 28, 7, 8, 'High', '#7A6000', '#FFF8D4', 10));

  // ── DEADLINES TAB CF ──
  // Days Left (col E, index 4) — FILTER output at rows 4-19 (0-indexed)
  reqs.push(cfFormula(DL, 4, 20, 4, 5, '=E5<7',  '#8B1A1A', '#F8D7DA', 0, true));
  reqs.push(cfFormula(DL, 4, 20, 4, 5, '=AND(E5>=7,E5<=30)', '#7A6000', '#FFF8D4', 1));
  // Status In Progress col F (index 5)
  reqs.push(cfText(DL, 4, 20, 5, 6, 'In Progress', '#1A3A8B', '#D6E8FF', 2));
  // Checklist checkboxes: TRUE → green bg
  reqs.push(cfText(DL, 24, 32, 1, 7, 'TRUE', '#1A6B3A', '#D4EDDA', 3));

  // ── RESULTS TAB CF ──
  // Status col D (index 3) at rows 4-20 (0-indexed)
  reqs.push(cfText(RES, 4, 21, 3, 4, 'Awarded',    '#1A6B3A', '#D4EDDA', 0, true));
  reqs.push(cfText(RES, 4, 21, 3, 4, 'Rejected',   '#8B1A1A', '#F8D7DA', 1));
  reqs.push(cfText(RES, 4, 21, 3, 4, 'Waitlisted', '#6A5A8B', '#EDE8F8', 2));
  reqs.push(cfText(RES, 4, 21, 3, 4, 'Submitted',  '#1A3A8B', '#D6E8FF', 3));

  // ── DASHBOARD CF ──
  // Days to next deadline (B19, 0-indexed row 18, col B = index 1)
  reqs.push(cfFormula(DSH, 18, 19, 2, 6, '=C19<7',           '#8B1A1A', '#F8D7DA', 0, true));
  reqs.push(cfFormula(DSH, 18, 19, 2, 6, '=AND(C19>=7,C19<=30)', '#7A6000', '#FFF8D4', 1));

  await batchUpdate(id, reqs, 'cf-rules');
  console.log('Conditional Formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
