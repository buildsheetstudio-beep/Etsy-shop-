'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Weekly Schedule'];
const S = "'Weekly Schedule'";
const MTL = "'Master Task Log'";
const PS  = "'Project Setup'";
// Week Start Date is at Project Setup row 5 (value row for control i=4)
// 0-indexed row 5 = 1-indexed row 6, col A → PS!$A$6
const WK  = `${PS}!$G$6`;
const PROJ = `${PS}!$A$4`;

const DAYS = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
const NC = 7;
const NROWS = 40; // task rows per day column

// Project filter array (scalar 1 when no project set, else project match)
const PROJ_ARR = `IF(${PROJ}="",1,${MTL}!$B$8:$B$3007=${PROJ})`;

// Filter formula: tasks active on a given day (start <= day <= due), not complete/cancelled
const dayFml = (dayOffset, cardIdx) => {
  const DAY = `${WK}+${dayOffset}`;
  const cond = `(${MTL}!$K$8:$K$3007<=${DAY})*(${MTL}!$L$8:$L$3007>=${DAY})*(${MTL}!$N$8:$N$3007<>"Complete")*(${MTL}!$N$8:$N$3007<>"Cancelled")*(${PROJ_ARR})`;
  const arr  = `${MTL}!$F$8:$F$3007&" · "&${MTL}!$E$8:$E$3007`;
  return `=IFERROR(INDEX(FILTER(${arr},${cond}),${cardIdx}),"")`;
};

// Count formula for that day
const countFml = (dayOffset) => {
  const DAY = `${WK}+${dayOffset}`;
  return `=COUNTIFS(${MTL}!$K$8:$K$3007,"<="&${DAY},${MTL}!$L$8:$L$3007,">="&${DAY},${MTL}!$N$8:$N$3007,"<>Complete",${MTL}!$N$8:$N$3007,"<>Cancelled"${PROJ==='all'?'':`,${MTL}!$B$8:$B$3007,IF(${PROJ}="","<>",${PROJ})`})`;
};

(async () => {
  const fmt = [];
  const vals = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['WEEKLY SCHEDULE']] });

  // ── Subtitle ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.teal), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [[`=IFERROR("Week of "&TEXT(${WK},"mmm d")&" – "&TEXT(${WK}+6,"mmm d, yyyy"),"Weekly Schedule")`]] });

  // ── Instructions ──────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.terra), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A3`, values: [['Shows all active tasks spanning each day (start date ≤ day ≤ due date). Change week: update Week Start Date in Project Setup (A6). Filter by project: set Project ID in A4.']] });

  // ── Task count row (row 3, 0-indexed) ────────────────────────────────────
  DAYS.forEach((_, di) => {
    fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, di, di+1), cell: { userEnteredFormat: { backgroundColor: hex(C.bg), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
    // Simple count formula
    const DAY = `${WK}+${di}`;
    vals.push({ range: `${S}!${String.fromCharCode(65+di)}4`, values: [[
      `=IFERROR(COUNTIFS(${MTL}!$K$8:$K$3007,"<="&${DAY},${MTL}!$L$8:$L$3007,">="&${DAY},${MTL}!$N$8:$N$3007,"<>Complete",${MTL}!$N$8:$N$3007,"<>Cancelled")&" tasks","")`
    ]] });
  });

  // ── Day headers (row 4, 0-indexed = row 5) ───────────────────────────────
  DAYS.forEach((dayName, di) => {
    fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, di, di+1), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!${String.fromCharCode(65+di)}5`, values: [[
      `=${dayName}!A5` === '' ? dayName : `=IFERROR(UPPER(TEXT(${WK}+${di},"ddd mmm d")),"${dayName.slice(0,3).toUpperCase()}")`
    ]] });
  });

  // Highlight weekend columns (Saturday=5, Sunday=6) with lighter background
  [5, 6].forEach(di => {
    fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, di, di+1), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary) } }, fields: 'userEnteredFormat.backgroundColor' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 5, 5+NROWS, di, di+1), cell: { userEnteredFormat: { backgroundColor: hex(C.terra) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // ── Task card rows (rows 5-44, 0-indexed = rows 6-45) ────────────────────
  DAYS.forEach((_, di) => {
    const col = String.fromCharCode(65+di);
    fmt.push({ repeatCell: { range: gridRange(SID, 5, 5+NROWS, di, di+1), cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, wrapStrategy: 'WRAP', verticalAlignment: 'TOP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });
    for (let r = 0; r < NROWS; r++) {
      if (r % 2 !== 0) {
        fmt.push({ repeatCell: { range: gridRange(SID, 5+r, 6+r, di, di+1), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
      }
    }
    const cardFormulas = Array.from({ length: NROWS }, (_, r) => [dayFml(di, r+1)]);
    vals.push({ range: `${S}!${col}6:${col}${5+NROWS}`, values: cardFormulas });
  });

  // ── Freeze 5 rows ─────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Row + col dimensions ──────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 5+NROWS }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  for (let i = 0; i < NC; i++) {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: 200 }, fields: 'pixelSize' } });
  }

  // Column separators
  for (let i = 1; i < NC; i++) {
    fmt.push({ updateBorders: { range: gridRange(SID, 4, 5+NROWS, i, i+1), left: { style: 'SOLID', color: hex(C.border) } } });
  }

  await batchUpdate(id, fmt, 'ws-fmt');
  await valuesBatchUpdate(id, vals, 'ws-vals');
  console.log('✓ Weekly Schedule complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
