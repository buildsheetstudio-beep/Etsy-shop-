'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Eisenhower Matrix'];
const S = "'Eisenhower Matrix'";
const MTL = "'Master Task Log'";
const PS  = "'Project Setup'";
const PROJ = `${PS}!$A$4`;

// Quadrant definitions: [Important, Urgent]
// Q1: I=TRUE, U=TRUE  → Do Now   (top-left)
// Q2: I=TRUE, U=FALSE → Schedule (top-right)
// Q3: I=FALSE,U=TRUE  → Delegate (bottom-left)
// Q4: I=FALSE,U=FALSE → Eliminate(bottom-right)
const QUADS = [
  { label: 'Q1 — DO NOW',      imp: true,  urg: true,  bg: C.attention, fg: C.white,   desc: 'Important & Urgent — act immediately',        col: 0 },
  { label: 'Q2 — SCHEDULE',    imp: true,  urg: false, bg: C.primary,   fg: C.white,   desc: 'Important, Not Urgent — plan & protect time',  col: 1 },
  { label: 'Q3 — DELEGATE',    imp: false, urg: true,  bg: C.warning,   fg: C.text,    desc: 'Not Important, Urgent — delegate if possible',  col: 0 },
  { label: 'Q4 — ELIMINATE',   imp: false, urg: false, bg: C.gray,      fg: C.white,   desc: 'Not Important, Not Urgent — defer or remove',   col: 1 },
];

const NCOLS = 2;
const NROWS = 60; // task rows per quadrant

// Card formula for quadrant (1=important, 0=not; 1=urgent, 0=not)
const cardFml = (imp, urg, cardIdx) => {
  const impVal = imp ? 'TRUE' : 'FALSE';
  const urgVal = urg ? 'TRUE' : 'FALSE';
  const projFilter = `IF(${PROJ}="",1,${MTL}!$B$8:$B$3007=${PROJ})`;
  const cond = `(${MTL}!$I$8:$I$3007=${impVal})*(${MTL}!$J$8:$J$3007=${urgVal})*(${MTL}!$N$8:$N$3007<>"Complete")*(${MTL}!$N$8:$N$3007<>"Cancelled")*(${projFilter})`;
  const arr  = `${MTL}!$F$8:$F$3007&" · "&${MTL}!$E$8:$E$3007&" ["&${MTL}!$G$8:$G$3007&"]"`;
  return `=IFERROR(INDEX(FILTER(${arr},${cond}),${cardIdx}),"")`;
};

const countFml = (imp, urg) => {
  const impVal = imp ? 'TRUE' : 'FALSE';
  const urgVal = urg ? 'TRUE' : 'FALSE';
  const projFilter = `IF(${PROJ}="",1,${MTL}!$B$8:$B$3007=${PROJ})`;
  return `=IFERROR(SUMPRODUCT((${MTL}!$I$8:$I$3007=${impVal})*(${MTL}!$J$8:$J$3007=${urgVal})*(${MTL}!$N$8:$N$3007<>"Complete")*(${MTL}!$N$8:$N$3007<>"Cancelled")*(${projFilter})),"0")`;
};

(async () => {
  const fmt = [];
  const vals = [];

  // ── Title ──────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NCOLS), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NCOLS), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['EISENHOWER MATRIX']] });

  // ── Subtitle ───────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NCOLS), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NCOLS), cell: { userEnteredFormat: { backgroundColor: hex(C.teal), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [['Prioritize tasks by importance and urgency. Mark Important (col I) and Urgent (col J) checkboxes in Master Task Log. Filter by project: set Project ID in Project Setup A4.']] });

  // ── Instructions ───────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, NCOLS), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, NCOLS), cell: { userEnteredFormat: { backgroundColor: hex(C.terra), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A3`, values: [['READ-ONLY VIEW — Cards show: Task Name · Assignee [Priority]. Check Important and Urgent boxes in Master Task Log to place tasks in the correct quadrant.']] });

  // ── Axis labels row ────────────────────────────────────────────────────────
  // Col A header = "← URGENT" axis label, Col B = "NOT URGENT →" axis label
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(C.attention), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A4`, values: [['↑ URGENT']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!B4`, values: [['NOT URGENT ↑']] });

  // ── Quadrant headers + body (Q1/Q2 = row group 1, Q3/Q4 = row group 2)
  // Layout: header row, count row, then NROWS card rows
  // Q1 top-left (col A, rows 4-...), Q2 top-right (col B, rows 4-...)
  // Q3 bottom-left (col A, rows after Q1), Q4 bottom-right (col B, rows after Q2)
  const Q_HEADER_START = 4; // 0-indexed
  const Q_COUNT_ROW    = Q_HEADER_START + 1;
  const Q_CARD_START   = Q_HEADER_START + 2;
  const Q3_HEADER      = Q_CARD_START + NROWS;
  const Q3_COUNT_ROW   = Q3_HEADER + 1;
  const Q3_CARD_START  = Q3_HEADER + 2;

  // Q1 header (col A, row 4)
  fmt.push({ repeatCell: { range: gridRange(SID, Q_HEADER_START, Q_HEADER_START+1, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(C.attention), textFormat: { bold: true, fontSize: 12, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A5`, values: [['Q1 — DO NOW  |  Important & Urgent']] });

  // Q2 header (col B, row 4)
  fmt.push({ repeatCell: { range: gridRange(SID, Q_HEADER_START, Q_HEADER_START+1, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 12, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!B5`, values: [['Q2 — SCHEDULE  |  Important, Not Urgent']] });

  // Q1 count row
  fmt.push({ repeatCell: { range: gridRange(SID, Q_COUNT_ROW, Q_COUNT_ROW+1, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(C.bg), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.attention) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A6`, values: [[`=${countFml(true,true).slice(1)} & " tasks"`]] });
  // Q2 count row
  fmt.push({ repeatCell: { range: gridRange(SID, Q_COUNT_ROW, Q_COUNT_ROW+1, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.bg), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!B6`, values: [[`=${countFml(true,false).slice(1)} & " tasks"`]] });

  // Q1 cards (col A, rows Q_CARD_START to Q_CARD_START+NROWS-1)
  fmt.push({ repeatCell: { range: gridRange(SID, Q_CARD_START, Q_CARD_START+NROWS, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex('#FFF5F3'), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });
  for (let r = 0; r < NROWS; r++) {
    if (r % 2 !== 0) fmt.push({ repeatCell: { range: gridRange(SID, Q_CARD_START+r, Q_CARD_START+r+1, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }
  const q1Formulas = Array.from({ length: NROWS }, (_, r) => [cardFml(true, true, r+1)]);
  vals.push({ range: `${S}!A7:A${6+NROWS}`, values: q1Formulas });

  // Q2 cards (col B)
  fmt.push({ repeatCell: { range: gridRange(SID, Q_CARD_START, Q_CARD_START+NROWS, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });
  for (let r = 0; r < NROWS; r++) {
    if (r % 2 !== 0) fmt.push({ repeatCell: { range: gridRange(SID, Q_CARD_START+r, Q_CARD_START+r+1, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }
  const q2Formulas = Array.from({ length: NROWS }, (_, r) => [cardFml(true, false, r+1)]);
  vals.push({ range: `${S}!B7:B${6+NROWS}`, values: q2Formulas });

  // Divider row between Q1/Q2 and Q3/Q4
  fmt.push({ mergeCells: { range: gridRange(SID, Q3_HEADER-1, Q3_HEADER, 0, NCOLS), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, Q3_HEADER-1, Q3_HEADER, 0, NCOLS), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A${Q3_HEADER}`, values: [['↓ NOT URGENT']] });

  // Q3 header (col A)
  fmt.push({ repeatCell: { range: gridRange(SID, Q3_HEADER, Q3_HEADER+1, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(C.warning), textFormat: { bold: true, fontSize: 12, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A${Q3_HEADER+1}`, values: [['Q3 — DELEGATE  |  Not Important, Urgent']] });

  // Q4 header (col B)
  fmt.push({ repeatCell: { range: gridRange(SID, Q3_HEADER, Q3_HEADER+1, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.gray), textFormat: { bold: true, fontSize: 12, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!B${Q3_HEADER+1}`, values: [['Q4 — ELIMINATE  |  Not Important, Not Urgent']] });

  // Q3 count row
  fmt.push({ repeatCell: { range: gridRange(SID, Q3_COUNT_ROW, Q3_COUNT_ROW+1, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(C.bg), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.warning) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A${Q3_COUNT_ROW+1}`, values: [[`=${countFml(false,true).slice(1)} & " tasks"`]] });
  // Q4 count row
  fmt.push({ repeatCell: { range: gridRange(SID, Q3_COUNT_ROW, Q3_COUNT_ROW+1, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.bg), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.gray) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!B${Q3_COUNT_ROW+1}`, values: [[`=${countFml(false,false).slice(1)} & " tasks"`]] });

  // Q3 cards (col A)
  fmt.push({ repeatCell: { range: gridRange(SID, Q3_CARD_START, Q3_CARD_START+NROWS, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex('#FFFBF0'), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });
  for (let r = 0; r < NROWS; r++) {
    if (r % 2 !== 0) fmt.push({ repeatCell: { range: gridRange(SID, Q3_CARD_START+r, Q3_CARD_START+r+1, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }
  const q3Formulas = Array.from({ length: NROWS }, (_, r) => [cardFml(false, true, r+1)]);
  vals.push({ range: `${S}!A${Q3_CARD_START+1}:A${Q3_CARD_START+NROWS}`, values: q3Formulas });

  // Q4 cards (col B)
  fmt.push({ repeatCell: { range: gridRange(SID, Q3_CARD_START, Q3_CARD_START+NROWS, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });
  for (let r = 0; r < NROWS; r++) {
    if (r % 2 !== 0) fmt.push({ repeatCell: { range: gridRange(SID, Q3_CARD_START+r, Q3_CARD_START+r+1, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }
  const q4Formulas = Array.from({ length: NROWS }, (_, r) => [cardFml(false, false, r+1)]);
  vals.push({ range: `${S}!B${Q3_CARD_START+1}:B${Q3_CARD_START+NROWS}`, values: q4Formulas });

  // ── Freeze 3 header rows ───────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Row heights ────────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });
  // Card rows
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 6+NROWS }, properties: { pixelSize: 34 }, fields: 'pixelSize' } });
  // Divider
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: Q3_HEADER-1, endIndex: Q3_HEADER }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: Q3_HEADER, endIndex: Q3_HEADER+1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: Q3_COUNT_ROW, endIndex: Q3_COUNT_ROW+1 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: Q3_CARD_START, endIndex: Q3_CARD_START+NROWS }, properties: { pixelSize: 34 }, fields: 'pixelSize' } });

  // ── Column widths ──────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 520 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 520 }, fields: 'pixelSize' } });

  // Column separator
  fmt.push({ updateBorders: { range: gridRange(SID, 3, Q3_CARD_START+NROWS, 1, 2), left: { style: 'SOLID_MEDIUM', color: hex(C.border) } } });

  await batchUpdate(id, fmt, 'em-fmt');
  await valuesBatchUpdate(id, vals, 'em-vals');
  console.log('✓ Eisenhower Matrix complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
