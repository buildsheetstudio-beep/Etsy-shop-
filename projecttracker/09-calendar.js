'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Monthly Calendar'];
const S = "'Monthly Calendar'";
const MTL = "'Master Task Log'";
const PS  = "'Project Setup'";

// Calendar Month = PS!$M$6 (control i=6), Calendar Year = PS!$S$6 (control i=7)
const MONTH = `${PS}!$M$6`;
const YEAR  = `${PS}!$S$6`;
const PROJ  = `${PS}!$A$4`;
const CAL   = `DATE(${YEAR},${MONTH},1)-WEEKDAY(DATE(${YEAR},${MONTH},1),2)+1`;
const PROJ_ARR = `IF(${PROJ}="",1,${MTL}!$B$8:$B$3007=${PROJ})`;

const NC = 7; // Mon-Sun
const WEEKS = 6;
const ROWS_PER_WEEK = 4; // day# row + 2 task rows + separator row
// Header: rows 0-6, calendar body: rows 7-30
const CAL_START_ROW = 7;

(async () => {
  const fmt = [];
  const vals = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['MONTHLY CALENDAR']] });

  // ── Subtitle (dynamic month/year) ─────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.teal), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [[`=IFERROR(TEXT(DATE(${YEAR},${MONTH},1),"MMMM yyyy")&" — "&COUNTIFS(${MTL}!$L$8:$L$3007,">="&DATE(${YEAR},${MONTH},1),${MTL}!$L$8:$L$3007,"<"&DATE(${YEAR},${MONTH}+1,1),${MTL}!$N$8:$N$3007,"<>Cancelled")&" tasks due this month","Monthly Calendar")`]] });

  // ── Instructions ──────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.terra), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A3`, values: [['Shows tasks due each day. Change month/year: update Calendar Month (M6) and Calendar Year (S6) in Project Setup. Filter by project: set Project ID in A4.']] });

  // ── Stats row ─────────────────────────────────────────────────────────────
  const statsW = Math.floor(NC / 4); // ~1-2 cols each
  const stats = [
    [`=COUNTIFS(${MTL}!$L$8:$L$3007,">="&DATE(${YEAR},${MONTH},1),${MTL}!$L$8:$L$3007,"<"&DATE(${YEAR},${MONTH}+1,1),${MTL}!$N$8:$N$3007,"<>Complete",${MTL}!$N$8:$N$3007,"<>Cancelled")&" Open Tasks"`],
    [`=COUNTIFS(${MTL}!$L$8:$L$3007,">="&DATE(${YEAR},${MONTH},1),${MTL}!$L$8:$L$3007,"<"&TODAY(),${MTL}!$N$8:$N$3007,"<>Complete",${MTL}!$N$8:$N$3007,"<>Cancelled")&" Overdue"`],
    [`=COUNTIFS(${MTL}!$L$8:$L$3007,">="&TODAY(),${MTL}!$L$8:$L$3007,"<="&(TODAY()+7),${MTL}!$N$8:$N$3007,"<>Complete",${MTL}!$N$8:$N$3007,"<>Cancelled")&" Due This Week"`],
    [`=COUNTIFS(${MTL}!$L$8:$L$3007,">="&DATE(${YEAR},${MONTH},1),${MTL}!$L$8:$L$3007,"<"&DATE(${YEAR},${MONTH}+1,1),${MTL}!$N$8:$N$3007,"Complete")&" Completed"`],
  ];
  stats.forEach(([fml], i) => {
    const c = i * Math.floor(NC / 4);
    const c2 = i < 3 ? c + Math.floor(NC/4) : NC;
    fmt.push({ mergeCells: { range: gridRange(SID, 3, 4, c, c2), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, c, c2), cell: { userEnteredFormat: { backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}4`, values: [[fml]] });
  });

  // ── Spacer row 4 (0-indexed) ──────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // ── Month header row 5 (0-indexed) ───────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 5, 6, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 5, 6, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 14, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A6`, values: [[`=IFERROR(TEXT(DATE(${YEAR},${MONTH},1),"MMMM YYYY"),"")`]] });

  // ── Day-of-week headers (row 6, 0-indexed = row 7) ───────────────────────
  const DOW = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
  const DOW_COLORS = [C.primary, C.primary, C.primary, C.primary, C.primary, C.secondary, C.secondary];
  DOW.forEach((day, d) => {
    fmt.push({ repeatCell: { range: gridRange(SID, 6, 7, d, d+1), cell: { userEnteredFormat: { backgroundColor: hex(DOW_COLORS[d]), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!${String.fromCharCode(65+d)}7`, values: [[day]] });
  });

  // ── Calendar grid (weeks 0-5, each 4 rows) ───────────────────────────────
  for (let w = 0; w < WEEKS; w++) {
    const baseRow = CAL_START_ROW + w * ROWS_PER_WEEK;

    for (let d = 0; d < NC; d++) {
      const col = String.fromCharCode(65+d);
      const cellDate = `${CAL}+${w*7+d}`;
      const isWeekend = d >= 5;

      // Day number row
      const dayNumBg = isWeekend ? C.terra : C.panel;
      fmt.push({ repeatCell: { range: gridRange(SID, baseRow, baseRow+1, d, d+1), cell: { userEnteredFormat: { backgroundColor: hex(dayNumBg), textFormat: { bold: true, fontSize: 13, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
      vals.push({ range: `${S}!${col}${baseRow+1}`, values: [[
        `=IF(MONTH(${cellDate})=${MONTH},DAY(${cellDate}),"")`
      ]] });

      // Task row 1 (in-month only)
      const taskBg1 = isWeekend ? '#F5EAEA' : C.panel;
      fmt.push({ repeatCell: { range: gridRange(SID, baseRow+1, baseRow+2, d, d+1), cell: { userEnteredFormat: { backgroundColor: hex(taskBg1), textFormat: { fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) }, wrapStrategy: 'WRAP', verticalAlignment: 'TOP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });
      vals.push({ range: `${S}!${col}${baseRow+2}`, values: [[
        `=IFERROR(IF(MONTH(${cellDate})<>${MONTH},"",INDEX(FILTER(${MTL}!$F$8:$F$3007,(${MTL}!$L$8:$L$3007=${cellDate})*(${MTL}!$N$8:$N$3007<>"Cancelled")*(${PROJ_ARR})),1)),"")`
      ]] });

      // Task row 2 + overflow count
      const taskBg2 = isWeekend ? '#F5EAEA' : C.altRow;
      fmt.push({ repeatCell: { range: gridRange(SID, baseRow+2, baseRow+3, d, d+1), cell: { userEnteredFormat: { backgroundColor: hex(taskBg2), textFormat: { fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, wrapStrategy: 'WRAP', verticalAlignment: 'TOP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });
      vals.push({ range: `${S}!${col}${baseRow+3}`, values: [[
        `=IFERROR(IF(MONTH(${cellDate})<>${MONTH},"",LET(cnt,COUNTIFS(${MTL}!$L$8:$L$3007,${cellDate},${MTL}!$N$8:$N$3007,"<>Cancelled"),IF(cnt=0,"",IF(cnt=1,INDEX(FILTER(${MTL}!$F$8:$F$3007,(${MTL}!$L$8:$L$3007=${cellDate})*(${MTL}!$N$8:$N$3007<>"Cancelled")*(${PROJ_ARR})),1),IF(cnt=2,INDEX(FILTER(${MTL}!$F$8:$F$3007,(${MTL}!$L$8:$L$3007=${cellDate})*(${MTL}!$N$8:$N$3007<>"Cancelled")*(${PROJ_ARR})),2),"+"&(cnt-1)&" more"))))),"")`
      ]] });

      // Separator row
      fmt.push({ repeatCell: { range: gridRange(SID, baseRow+3, baseRow+4, d, d+1), cell: { userEnteredFormat: { backgroundColor: hex(C.border) } }, fields: 'userEnteredFormat.backgroundColor' } });
    }
  }

  // ── Freeze header rows 1-7 ────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Row heights ───────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  for (let w = 0; w < WEEKS; w++) {
    const base = CAL_START_ROW + w * ROWS_PER_WEEK;
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: base, endIndex: base+1 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } }); // day#
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: base+1, endIndex: base+3 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } }); // tasks
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: base+3, endIndex: base+4 }, properties: { pixelSize: 4 }, fields: 'pixelSize' } }); // separator
  }

  // ── Column widths ─────────────────────────────────────────────────────────
  for (let i = 0; i < NC; i++) {
    const w = i >= 5 ? 130 : 175; // weekends slightly narrower
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  }

  // Cell borders for calendar grid
  for (let w = 0; w < WEEKS; w++) {
    const r = CAL_START_ROW + w * ROWS_PER_WEEK;
    for (let d = 0; d < NC; d++) {
      const borderStyle = { style: 'SOLID', color: hex(C.border) };
      fmt.push({ updateBorders: { range: gridRange(SID, r, r+3, d, d+1), top: borderStyle, bottom: borderStyle, left: borderStyle, right: borderStyle } });
    }
  }

  await batchUpdate(id, fmt, 'cal-fmt');
  await valuesBatchUpdate(id, vals, 'cal-vals');
  console.log('✓ Monthly Calendar complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
