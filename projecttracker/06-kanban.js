'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Kanban Board'];
const S = "'Kanban Board'";
const MTL = "'Master Task Log'";
const PROJ = "'Project Setup'!$A$4";

const STATUSES = [
  { label: 'BACKLOG',      status: 'Backlog',      bg: C.gray,      fg: C.white },
  { label: 'NOT STARTED',  status: 'Not Started',  bg: C.secText,   fg: C.white },
  { label: 'IN PROGRESS',  status: 'In Progress',  bg: C.primary,   fg: C.white },
  { label: 'WAITING',      status: 'Waiting',       bg: C.warning,   fg: C.text  },
  { label: 'BLOCKED',      status: 'Blocked',       bg: C.attention, fg: C.white },
  { label: 'REVIEW',       status: 'Review',        bg: C.info,      fg: C.text  },
  { label: 'COMPLETE',     status: 'Complete',      bg: C.success,   fg: C.white },
];
const NC = STATUSES.length; // 7 columns
const NROWS = 50; // task rows per column

// FILTER formula for card at position cardIdx (1-based) in given column
const cardFml = (status, cardIdx) => {
  const cond = `(${MTL}!$N$8:$N$3007="${status}")*(IF(${PROJ}="",1,${MTL}!$B$8:$B$3007=${PROJ}))`;
  // Show: task name — assignee
  const arr = `${MTL}!$F$8:$F$3007&" · "&${MTL}!$E$8:$E$3007`;
  return `=IFERROR(INDEX(FILTER(${arr},${cond}),${cardIdx}),"")`;
};

(async () => {
  const fmt = [];
  const vals = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['KANBAN BOARD']] });

  // ── Subtitle ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.teal), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [['Live view from Master Task Log. Set "Selected Project ID" in Project Setup (A4) to filter by project, or leave blank to see all tasks.']] });

  // ── Instructions row ──────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.terra), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A3`, values: [['READ-ONLY VIEW — Edit tasks in the Master Task Log. Card format: Task Name · Assignee. Filtered by "Selected Project ID" in Project Setup!A4 (blank = show all).']] });

  // ── Count row (row 3, 0-indexed) ──────────────────────────────────────────
  STATUSES.forEach(({ status }, ci) => {
    const col = String.fromCharCode(65+ci);
    fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.bg), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
    const countFml = `=COUNTIFS(${MTL}!$N$8:$N$3007,"${status}",IF(${PROJ}="",${MTL}!$N$8:$N$3007,${MTL}!$B$8:$B$3007),IF(${PROJ}="","${status}",${PROJ}))`;
    vals.push({ range: `${S}!${col}4`, values: [[countFml]] });
  });

  // ── Status column headers (row 4, 0-indexed) ──────────────────────────────
  STATUSES.forEach(({ label, bg, fg }, ci) => {
    fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(fg) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!${String.fromCharCode(65+ci)}5`, values: [[label]] });
  });

  // ── Task card rows (rows 5-54, 0-indexed = rows 6-55) ────────────────────
  STATUSES.forEach(({ status, bg }, ci) => {
    const col = String.fromCharCode(65+ci);
    // Slightly lighter bg for card area
    fmt.push({ repeatCell: { range: gridRange(SID, 5, 5+NROWS, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });
    // Every other card slightly tinted
    for (let r = 0; r < NROWS; r++) {
      if (r % 2 !== 0) {
        fmt.push({ repeatCell: { range: gridRange(SID, 5+r, 6+r, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
      }
    }
    // Write FILTER formulas for each card
    const cardFormulas = Array.from({ length: NROWS }, (_, r) => [cardFml(status, r+1)]);
    vals.push({ range: `${S}!${col}6:${col}${5+NROWS}`, values: cardFormulas });
  });

  // ── Freeze row 1-5 ────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Row heights ───────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  // Card rows: taller for readability
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 5+NROWS }, properties: { pixelSize: 38 }, fields: 'pixelSize' } });

  // ── Column widths (equal) ─────────────────────────────────────────────────
  for (let i = 0; i < NC; i++) {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: 200 }, fields: 'pixelSize' } });
  }

  // ── Column borders (thin separators) ─────────────────────────────────────
  STATUSES.forEach((_, ci) => {
    if (ci < NC-1) {
      fmt.push({
        updateBorders: {
          range: gridRange(SID, 4, 5+NROWS, ci+1, ci+2),
          left: { style: 'SOLID_MEDIUM', color: hex(C.border) },
        },
      });
    }
  });

  await batchUpdate(id, fmt, 'kb-fmt');
  await valuesBatchUpdate(id, vals, 'kb-vals');
  console.log('✓ Kanban Board complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
