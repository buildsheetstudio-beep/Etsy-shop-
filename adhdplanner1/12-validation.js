'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

(async () => {
  const fmt = [];

  const REF = sheetMap['Reference Data'];
  const BD  = sheetMap['Brain Dump'];
  const MT  = sheetMap['Master Tasks'];
  const TB  = sheetMap['Task Breakdown'];
  const DF  = sheetMap['Daily Focus'];
  const WP  = sheetMap['Weekly Planner'];

  // Helper: ONE_OF_RANGE validation — value must be prefixed with =
  function rangeVal(sheetId, startRow, endRow, startCol, endCol, refRange) {
    return { setDataValidation: { range: gridRange(sheetId, startRow, endRow, startCol, endCol),
      rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${refRange}` }] }, strict: false, showCustomUi: true },
    }};
  }

  // Helper: ONE_OF_LIST validation
  function listVal(sheetId, startRow, endRow, startCol, endCol, items) {
    return { setDataValidation: { range: gridRange(sheetId, startRow, endRow, startCol, endCol),
      rule: { condition: { type: 'ONE_OF_LIST', values: items.map(v => ({ userEnteredValue: v })) }, strict: false, showCustomUi: true },
    }};
  }

  // Helper: BooleanCondition checkbox
  function checkVal(sheetId, startRow, endRow, startCol, endCol) {
    return { setDataValidation: { range: gridRange(sheetId, startRow, endRow, startCol, endCol),
      rule: { condition: { type: 'BOOLEAN' }, strict: true },
    }};
  }

  // Helper: Number range
  function numVal(sheetId, startRow, endRow, startCol, endCol, min, max) {
    return { setDataValidation: { range: gridRange(sheetId, startRow, endRow, startCol, endCol),
      rule: { condition: { type: 'NUMBER_BETWEEN', values: [{ userEnteredValue: String(min) }, { userEnteredValue: String(max) }] }, strict: false },
    }};
  }

  // Reference ranges (0-indexed row/col)
  // Col A: Task Status (A3:A10 = rows 2-9)  → "Reference Data"!$A$3:$A$10
  // Col B: Focus Stage (B3:B7)
  // Col C: Today Capacity (C3:C5)
  // Col D: Priority (D3:D6)
  // Col E: Effort Level (E3:E5)
  // Col F: Time Estimate (F3:F10)
  // Col G: Task Type (G3:G14)
  // Col H: Rescue Action (H3:H6)
  // Col I: Day Type (I3:I7)
  // Col J: Energy Match (J3:J6)
  // Col K: Yes/No (K3:K4)

  const R = "'Reference Data'";

  // ===== BRAIN DUMP =====
  // Col D: Task Type (G in Reference = col 6, rows 3-14 → $G$3:$G$14)
  fmt.push(rangeVal(BD, 7, 2007, 3, 4, `${R}!$G$3:$G$14`));
  // Col F: Effort Level (col E in Reference = $E$3:$E$5)
  fmt.push(rangeVal(BD, 7, 2007, 5, 6, `${R}!$E$3:$E$5`));
  // Col G: Ready? checkbox
  fmt.push(checkVal(BD, 7, 2007, 6, 7));
  // Col H: Moved? checkbox
  fmt.push(checkVal(BD, 7, 2007, 7, 8));

  // ===== MASTER TASKS =====
  // Col B (index 1): Focus Stage
  fmt.push(rangeVal(MT, 7, 5007, 1, 2, `${R}!$B$3:$B$7`));
  // Col D (index 3): Priority
  fmt.push(rangeVal(MT, 7, 5007, 3, 4, `${R}!$D$3:$D$6`));
  // Col E (index 4): Status
  fmt.push(rangeVal(MT, 7, 5007, 4, 5, `${R}!$A$3:$A$10`));
  // Col G (index 6): Effort Level
  fmt.push(rangeVal(MT, 7, 5007, 6, 7, `${R}!$E$3:$E$5`));
  // Col H (index 7): Time Estimate
  fmt.push(rangeVal(MT, 7, 5007, 7, 8, `${R}!$F$3:$F$10`));
  // Col I (index 8): Task Type
  fmt.push(rangeVal(MT, 7, 5007, 8, 9, `${R}!$G$3:$G$14`));
  // Col J (index 9): Energy Match
  fmt.push(rangeVal(MT, 7, 5007, 9, 10, `${R}!$J$3:$J$6`));
  // Col K (index 10): Breakdown? checkbox
  fmt.push(checkVal(MT, 7, 5007, 10, 11));
  // Col L (index 11): Rescue Action
  fmt.push(rangeVal(MT, 7, 5007, 11, 12, `${R}!$H$3:$H$6`));

  // ===== TASK BREAKDOWN =====
  // Col F (index 5): Effort
  fmt.push(rangeVal(TB, 25, 600, 5, 6, `${R}!$E$3:$E$5`));
  // Col H (index 7): Status
  fmt.push(listVal(TB, 25, 600, 7, 8, ['Not Started','In Progress','Completed','Blocked','Waiting']));
  // Col I (index 8): Done? checkbox
  fmt.push(checkVal(TB, 25, 600, 8, 9));
  // Also checkbox for Task Breakdown parent reference rows (8-15)
  fmt.push(checkVal(TB, 7, 16, 8, 9));

  // ===== DAILY FOCUS =====
  // Row 2, col D (index 3): Day Type
  fmt.push(rangeVal(DF, 1, 2, 3, 4, `${R}!$I$3:$I$7`));
  // Row 5 (index 4): Energy Level 1-5 col B
  fmt.push(numVal(DF, 4, 5, 1, 2, 1, 5));
  // Row 5: Sleep Quality 1-5 col F (index 5)
  fmt.push(numVal(DF, 4, 5, 5, 6, 1, 5));
  // Row 6 (index 5): Today Capacity col D
  fmt.push(rangeVal(DF, 5, 6, 3, 4, `${R}!$C$3:$C$5`));
  // Priority 3 rows done? checkboxes col E (index 4)
  fmt.push(checkVal(DF, 12, 15, 4, 5));
  // Task queue cols B (Focus Stage), C (Priority), D (Effort), F (Status), H (Rescue), I (Done?)
  fmt.push(rangeVal(DF, 18, 49, 1, 2, `${R}!$B$3:$B$7`));
  fmt.push(rangeVal(DF, 18, 49, 2, 3, `${R}!$D$3:$D$6`));
  fmt.push(rangeVal(DF, 18, 49, 3, 4, `${R}!$E$3:$E$5`));
  fmt.push(rangeVal(DF, 18, 49, 6, 7, `${R}!$A$3:$A$10`));
  fmt.push(rangeVal(DF, 18, 49, 7, 8, `${R}!$H$3:$H$6`));
  fmt.push(checkVal(DF, 18, 49, 8, 9));

  // ===== WEEKLY PLANNER =====
  // Top 5 priorities done? col E (index 4)
  fmt.push(checkVal(WP, 5, 10, 4, 5));
  // Task grid col B: Focus Stage
  fmt.push(rangeVal(WP, 13, 200, 1, 2, `${R}!$B$3:$B$7`));
  // Col C: Priority
  fmt.push(rangeVal(WP, 13, 200, 2, 3, `${R}!$D$3:$D$6`));
  // Col D: Effort
  fmt.push(rangeVal(WP, 13, 200, 3, 4, `${R}!$E$3:$E$5`));
  // Col F: Day
  fmt.push(rangeVal(WP, 13, 200, 5, 6, `${R}!$N$3:$N$9`));
  // Col H: Done? checkbox
  fmt.push(checkVal(WP, 13, 200, 7, 8));

  await batchUpdate(id, fmt, '12-validation');
  console.log('✅  Data validation done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
