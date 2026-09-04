'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Gantt Chart'];
const S = "'Gantt Chart'";
const MTL = "'Master Task Log'";
const PS  = "'Project Setup'";

// Task info columns A-J (0-indexed 0-9), then weekly date columns K(10) onwards
// Total: 10 info cols + 48 weekly cols = 58 cols (A-BF)
const INFO_COLS = 10;
const N_WEEKS   = 48; // Jan 5 - Nov 16, 2026
const GANTT_START = 'DATE(2026,1,5)'; // first Monday of 2026
const NC = INFO_COLS + N_WEEKS;
const N_TASK_ROWS = 150;

// FILTER condition: project filter from Project Setup!A4
const COND = `(IF(${PS}!$A$4="",1,${MTL}!$B$8:$B$3007=${PS}!$A$4))*(${MTL}!$F$8:$F$3007<>"")`;

// MTL source columns (0-indexed): A=0(taskId), B=1(projId), E=4(assignee), F=5(name),
//   G=6(priority), H=7(pct), K=10(start), M=12(due), N=13(status)
const SRC = {
  taskId:   `${MTL}!$A$8:$A$3007`,
  projId:   `${MTL}!$B$8:$B$3007`,
  taskName: `${MTL}!$F$8:$F$3007`,
  assignee: `${MTL}!$E$8:$E$3007`,
  priority: `${MTL}!$G$8:$G$3007`,
  pct:      `${MTL}!$H$8:$H$3007`,
  startDt:  `${MTL}!$K$8:$K$3007`,
  dueDt:    `${MTL}!$L$8:$L$3007`,
  status:   `${MTL}!$N$8:$N$3007`,
};

const filterFml = (srcRange, rowIdx) =>
  `=IFERROR(INDEX(FILTER(${srcRange},${COND}),${rowIdx}),"")`;

(async () => {
  const fmt = [];
  const vals = [];

  // Expand sheet to fit info cols + week cols (default is 26)
  const neededCols = NC;
  if (neededCols > 26) {
    await batchUpdate(id, [{
      appendDimension: { sheetId: SID, dimension: 'COLUMNS', length: neededCols - 26 },
    }], 'gantt-expand-cols');
  }

  // ── Title ─────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['GANTT CHART — 2026']] });

  // ── Subtitle ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.teal), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [['Task timeline Jan–Nov 2026. Bars highlighted by conditional formatting. Filter by project: set Selected Project ID in Project Setup (A4).']] });

  // ── Instructions ──────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.terra), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A3`, values: [['READ-ONLY VIEW — Edit tasks in Master Task Log. Gantt bars filled by conditional formatting (green=on track, orange=at risk, red=overdue).']] });

  // ── Month label row (row 3, 0-indexed) ───────────────────────────────────
  // Group weeks into months visually with merged labels
  const MONTHS = [
    { label: 'January 2026',   startWk: 0,  endWk: 3  },
    { label: 'February 2026',  startWk: 4,  endWk: 7  },
    { label: 'March 2026',     startWk: 8,  endWk: 11 },
    { label: 'April 2026',     startWk: 12, endWk: 15 },
    { label: 'May 2026',       startWk: 16, endWk: 19 },
    { label: 'June 2026',      startWk: 20, endWk: 23 },
    { label: 'July 2026',      startWk: 24, endWk: 27 },
    { label: 'August 2026',    startWk: 28, endWk: 31 },
    { label: 'September 2026', startWk: 32, endWk: 35 },
    { label: 'October 2026',   startWk: 36, endWk: 39 },
    { label: 'November 2026',  startWk: 40, endWk: 43 },
    { label: 'Q4 Wrap',        startWk: 44, endWk: 47 },
  ];
  // Task info col header (merged, cols 0-9)
  fmt.push({ mergeCells: { range: gridRange(SID, 3, 4, 0, INFO_COLS), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 0, INFO_COLS), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A4`, values: [['TASK INFORMATION']] });

  MONTHS.forEach(({ label, startWk, endWk }) => {
    const c1 = INFO_COLS + startWk;
    const c2 = INFO_COLS + endWk + 1;
    if (c2 > NC) return;
    fmt.push({ mergeCells: { range: gridRange(SID, 3, 4, c1, c2), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, c1, c2), cell: { userEnteredFormat: { backgroundColor: hex(C.teal), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!${colLetter(c1)}4`, values: [[label]] });
  });

  // ── Column headers (row 4, 0-indexed = row 5) ────────────────────────────
  const INFO_HEADERS = ['Task ID','Project','Task Name','Assignee','Priority','Start Date','Due Date','Status','% Done','Days'];
  fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, 0, INFO_COLS), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)' } });
  vals.push({ range: `${S}!A5`, values: [INFO_HEADERS] });

  // Week date headers (cols K onwards, row 5, 0-indexed)
  fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, INFO_COLS, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 7, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', numberFormat: { type: 'DATE', pattern: 'M/d' } } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,numberFormat)' } });
  const weekDateFormulas = Array.from({ length: N_WEEKS }, (_, w) => [`=${GANTT_START}+${w*7}`]);
  vals.push({ range: `${S}!${colLetter(INFO_COLS)}5:${colLetter(NC-1)}5`, values: [weekDateFormulas.map(f => f[0])] });

  // ── Freeze rows 1-5 and cols A-J ─────────────────────────────────────────
  // Title row merges full width — so cannot freeze columns (freeze+merge conflict rule)
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Task rows (rows 5-154, 0-indexed = rows 6-155) ───────────────────────
  const taskInfoSrcs = [
    SRC.taskId, SRC.projId, SRC.taskName, SRC.assignee, SRC.priority,
    SRC.startDt, SRC.dueDt, SRC.status, SRC.pct, null, // null = duration formula
  ];

  taskInfoSrcs.forEach((src, ci) => {
    const col = colLetter(ci);
    const formulas = Array.from({ length: N_TASK_ROWS }, (_, r) => {
      if (src === null) {
        // Duration = due - start (col J = due col G - start col F)
        return [`=IF(F${6+r}="","",G${6+r}-F${6+r})`];
      }
      return [filterFml(src, r+1)];
    });
    vals.push({ range: `${S}!${col}6:${col}${5+N_TASK_ROWS}`, values: formulas });
  });

  // ── Gantt region styling ──────────────────────────────────────────────────
  // White/light background for all Gantt cells; CF will color bars
  fmt.push({ repeatCell: {
    range: gridRange(SID, 5, 5+N_TASK_ROWS, INFO_COLS, NC),
    cell: { userEnteredFormat: { backgroundColor: hex(C.panel) } },
    fields: 'userEnteredFormat.backgroundColor',
  }});

  // Alternating light tint on every other task row (info + gantt)
  for (let r = 0; r < N_TASK_ROWS; r++) {
    if (r % 2 !== 0) {
      fmt.push({ repeatCell: { range: gridRange(SID, 5+r, 6+r, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
    }
  }

  // Vertical lines at month boundaries in the gantt
  MONTHS.forEach(({ startWk }) => {
    const c = INFO_COLS + startWk;
    if (c < NC) {
      fmt.push({ updateBorders: { range: gridRange(SID, 4, 5+N_TASK_ROWS, c, c+1), left: { style: 'SOLID', color: hex(C.border) } } });
    }
  });

  // ── Number formats for task info ──────────────────────────────────────────
  // Col F (5) = Start Date, Col G (6) = Due Date
  [5, 6].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID, 5, 5+N_TASK_ROWS, ci, ci+1), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'M/d/yy' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });
  // Col I (8) = % Complete
  fmt.push({ repeatCell: { range: gridRange(SID, 5, 5+N_TASK_ROWS, 8, 9), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });
  // Formula col backgrounds (task info formula cols)
  [0,1,2,3,4,5,6,7,8,9].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID, 5, 5+N_TASK_ROWS, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // ── Column widths ─────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  const infoWidths = [65, 65, 180, 80, 65, 72, 72, 80, 50, 45];
  infoWidths.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  // Weekly columns: narrow
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: INFO_COLS, endIndex: NC }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  await batchUpdate(id, fmt, 'gantt-fmt');
  await valuesBatchUpdate(id, vals, 'gantt-vals');
  console.log('✓ Gantt Chart complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });

function colLetter(idx) {
  if (idx < 26) return String.fromCharCode(65+idx);
  return String.fromCharCode(64+Math.floor(idx/26)) + String.fromCharCode(65+(idx%26));
}
