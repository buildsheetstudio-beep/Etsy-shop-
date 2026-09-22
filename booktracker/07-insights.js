'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Reading Insights'];
const S = "'Reading Insights'";
const LIB = "'Master Book Library'";

// All SUMPRODUCT-based formulas (never COUNTIFS with YEAR/MONTH — causes #VALUE)
const sp = (cond) => `=SUMPRODUCT(${cond})`;
const spYr = (yr, extra = '') => `(YEAR(IFERROR(DATEVALUE(${LIB}!$O$8:$O$1008),0))=${yr})*(${LIB}!$M$8:$M$1008="Finished")${extra}`;
const GENRES = ['Fantasy','Mystery','Science Fiction','Literary Fiction','Romance','Historical Fiction','Thriller','Horror','Non-Fiction','Biography','Memoir','Self-Help','Graphic Novel','Young Adult','Other'];
const YEARS  = [2023,2024,2025,2026];

(async () => {
  const fmt  = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,200,0,18), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 10, fontFamily: 'Georgia', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Title ──────────────────────────────────────────────────────────────────
  vals.push({ range: `${S}!A1`, values: [['📊 Reading Insights']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,18), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,18), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 42 }, fields: 'pixelSize' }});

  // Subtitle
  vals.push({ range: `${S}!A2`, values: [['Analytics and statistics automatically calculated from your Master Book Library. All formulas update as you add books.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,18), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,18), cell: { userEnteredFormat: {
    backgroundColor: hex(C.goldTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  const sectionHeader = (row, text, colSpan, bg) => {
    vals.push({ range: `${S}!A${row}`, values: [[text]] });
    fmt.push({ mergeCells: { range: gridRange(SID,row-1,row,0,colSpan), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,row-1,row,0,colSpan), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: row-1, endIndex: row }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  };

  const tableHeader = (row, headers, bg) => {
    vals.push({ range: `${S}!A${row}`, values: [headers] });
    fmt.push({ repeatCell: { range: gridRange(SID,row-1,row,0,headers.length), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: row-1, endIndex: row }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  };

  const dataRow = (row, nCols, bg) => {
    fmt.push({ repeatCell: { range: gridRange(SID,row-1,row,0,nCols), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: row-1, endIndex: row }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  };

  // ── Section 1: Annual Reading Summary (rows 4-10) ─────────────────────────
  sectionHeader(4, 'SECTION 1 — ANNUAL READING SUMMARY', 10, C.primary);
  tableHeader(5, ['Year','Books Finished','Pages Read','Avg Rating','Avg Pages/Book','Best Rating (Most Common)','Non-Fiction Count','Fiction Count','Favorites','% Favorites'], C.primary);

  YEARS.forEach((yr, yi) => {
    const r = 6 + yi;
    dataRow(r, 10, yi % 2 === 0 ? C.white : C.altRow);
    vals.push({ range: `${S}!A${r}`, values: [[yr]] });
    // Books Finished
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,1,2), cell: { userEnteredValue: { formulaValue: sp(spYr(yr)) }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }}, fields: 'userEnteredValue,userEnteredFormat' }});
    // Pages Read
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,2,3), cell: { userEnteredValue: { formulaValue: sp(spYr(yr,`*${LIB}!$Q$8:$Q$1008`)) }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '#,##0' } }}, fields: 'userEnteredValue,userEnteredFormat' }});
    // Avg Rating
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,3,4), cell: { userEnteredValue: { formulaValue: `=IFERROR(ROUND(SUMPRODUCT(${spYr(yr)}*${LIB}!$P$8:$P$1008)/B${r},1),"—")` }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '0.0' } }}, fields: 'userEnteredValue,userEnteredFormat' }});
    // Avg Pages/Book
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,4,5), cell: { userEnteredValue: { formulaValue: `=IFERROR(ROUND(C${r}/B${r},0),"—")` }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }}, fields: 'userEnteredValue,userEnteredFormat' }});
    // Best Rating
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,5,6), cell: { userEnteredValue: { formulaValue: `=IFERROR(MODE(IF((YEAR(IFERROR(DATEVALUE(${LIB}!$O$8:$O$1008),0))=${yr})*(${LIB}!$M$8:$M$1008="Finished"),${LIB}!$P$8:$P$1008)),"—")` }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }}, fields: 'userEnteredValue,userEnteredFormat' }});
    // Non-Fiction Count
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,6,7), cell: { userEnteredValue: { formulaValue: sp(`${spYr(yr)}*(ISNUMBER(SEARCH("Non-Fiction",${LIB}!$D$8:$D$1008))+ISNUMBER(SEARCH("Biography",${LIB}!$D$8:$D$1008))+ISNUMBER(SEARCH("Memoir",${LIB}!$D$8:$D$1008))+ISNUMBER(SEARCH("Self-Help",${LIB}!$D$8:$D$1008))>0)`) }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }}, fields: 'userEnteredValue,userEnteredFormat' }});
    // Fiction Count
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,7,8), cell: { userEnteredValue: { formulaValue: `=IFERROR(B${r}-G${r},0)` }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }}, fields: 'userEnteredValue,userEnteredFormat' }});
    // Favorites
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,8,9), cell: { userEnteredValue: { formulaValue: sp(`${spYr(yr)}*(${LIB}!$U$8:$U$1008=TRUE)`) }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }}, fields: 'userEnteredValue,userEnteredFormat' }});
    // % Favorites
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,9,10), cell: { userEnteredValue: { formulaValue: `=IFERROR(I${r}/B${r},0)` }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER', numberFormat: { type: 'PERCENT', pattern: '0%' } }}, fields: 'userEnteredValue,userEnteredFormat' }});
  });

  // ── Section 2: Genre Breakdown (rows 12-30) ────────────────────────────────
  const GS = 12;
  sectionHeader(GS, 'SECTION 2 — BOOKS BY GENRE', 8, C.secondary);
  tableHeader(GS+1, ['Genre','Total Books','Finished','Avg Rating','Total Pages','% of Library','Favorites','DNF Count'], C.secondary);

  GENRES.forEach((genre, gi) => {
    const r = GS + 2 + gi;
    dataRow(r, 8, gi % 2 === 0 ? C.white : C.altRow);
    vals.push({ range: `${S}!A${r}`, values: [[genre]] });
    // Total
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,1,2), cell: { userEnteredValue: { formulaValue: sp(`(${LIB}!$D$8:$D$1008=A${r})*(${LIB}!$B$8:$B$1008<>"")`) }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }}, fields: 'userEnteredValue,userEnteredFormat' }});
    // Finished
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,2,3), cell: { userEnteredValue: { formulaValue: sp(`(${LIB}!$D$8:$D$1008=A${r})*(${LIB}!$M$8:$M$1008="Finished")`) }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }}, fields: 'userEnteredValue,userEnteredFormat' }});
    // Avg Rating
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,3,4), cell: { userEnteredValue: { formulaValue: `=IFERROR(ROUND(SUMPRODUCT((${LIB}!$D$8:$D$1008=A${r})*(${LIB}!$M$8:$M$1008="Finished")*${LIB}!$P$8:$P$1008)/C${r},1),"—")` }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '0.0' } }}, fields: 'userEnteredValue,userEnteredFormat' }});
    // Total Pages
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,4,5), cell: { userEnteredValue: { formulaValue: sp(`(${LIB}!$D$8:$D$1008=A${r})*(${LIB}!$M$8:$M$1008="Finished")*${LIB}!$Q$8:$Q$1008`) }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '#,##0' } }}, fields: 'userEnteredValue,userEnteredFormat' }});
    // % of Library
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,5,6), cell: { userEnteredValue: { formulaValue: `=IFERROR(B${r}/COUNTA(${LIB}!$B$8:$B$1008),0)` }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER', numberFormat: { type: 'PERCENT', pattern: '0%' } }}, fields: 'userEnteredValue,userEnteredFormat' }});
    // Favorites
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,6,7), cell: { userEnteredValue: { formulaValue: sp(`(${LIB}!$D$8:$D$1008=A${r})*(${LIB}!$U$8:$U$1008=TRUE)`) }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }}, fields: 'userEnteredValue,userEnteredFormat' }});
    // DNF Count
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,7,8), cell: { userEnteredValue: { formulaValue: sp(`(${LIB}!$D$8:$D$1008=A${r})*(${LIB}!$M$8:$M$1008="DNF")`) }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }}, fields: 'userEnteredValue,userEnteredFormat' }});
  });

  // ── Section 3: Rating Distribution (after genre table) ────────────────────
  const RS = GS + 2 + GENRES.length + 2;
  sectionHeader(RS, 'SECTION 3 — RATING DISTRIBUTION', 8, C.accent);
  tableHeader(RS+1, ['Star Rating','Total Books','% of Rated','Avg Length (pages)','Most Common Genre','2023 Count','2024 Count','2025 Count'], C.accent);

  [5,4,3,2,1].forEach((rating, ri) => {
    const r = RS + 2 + ri;
    dataRow(r, 8, ri % 2 === 0 ? C.white : C.altRow);
    vals.push({ range: `${S}!A${r}`, values: [[`${rating} ★ ${'★'.repeat(rating)}`]] });
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,1,2), cell: { userEnteredValue: { formulaValue: sp(`(${LIB}!$P$8:$P$1008=${rating})*(${LIB}!$M$8:$M$1008="Finished")`) }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }}, fields: 'userEnteredValue,userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,2,3), cell: { userEnteredValue: { formulaValue: `=IFERROR(B${r}/SUMPRODUCT((${LIB}!$M$8:$M$1008="Finished")*(${LIB}!$P$8:$P$1008>0)),0)` }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER', numberFormat: { type: 'PERCENT', pattern: '0%' } }}, fields: 'userEnteredValue,userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,3,4), cell: { userEnteredValue: { formulaValue: `=IFERROR(ROUND(SUMPRODUCT((${LIB}!$P$8:$P$1008=${rating})*(${LIB}!$M$8:$M$1008="Finished")*${LIB}!$Q$8:$Q$1008)/B${r},0),"—")` }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }}, fields: 'userEnteredValue,userEnteredFormat' }});
    // Most common genre for this rating
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,4,5), cell: { userEnteredValue: { formulaValue: `=IFERROR(INDEX(${LIB}!$D$8:$D$1008,MATCH(MAX(COUNTIFS(${LIB}!$P$8:$P$1008,${rating},${LIB}!$M$8:$M$1008,"Finished",${LIB}!$D$8:$D$1008,${LIB}!$D$8:$D$1008)),COUNTIFS(${LIB}!$P$8:$P$1008,${rating},${LIB}!$M$8:$M$1008,"Finished",${LIB}!$D$8:$D$1008,${LIB}!$D$8:$D$1008),0)),"—")` }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }}, fields: 'userEnteredValue,userEnteredFormat' }});
    YEARS.slice(0,3).forEach((yr, yi) => {
      fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,5+yi,6+yi), cell: { userEnteredValue: { formulaValue: sp(`(${LIB}!$P$8:$P$1008=${rating})*(${LIB}!$M$8:$M$1008="Finished")*(YEAR(IFERROR(DATEVALUE(${LIB}!$O$8:$O$1008),0))=${yr})`) }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }}, fields: 'userEnteredValue,userEnteredFormat' }});
    });
  });

  // ── Section 4: Format Breakdown (after rating table) ──────────────────────
  const FS = RS + 2 + 5 + 2;
  sectionHeader(FS, 'SECTION 4 — FORMAT BREAKDOWN', 6, C.primary);
  tableHeader(FS+1, ['Format','Total Books','Finished','Avg Rating','Avg Pages','% of Library'], C.primary);
  const FORMATS = ['Hardcover','Paperback','eBook','Audiobook','Library Book','Digital Library'];
  FORMATS.forEach((fmt_name, fi) => {
    const r = FS + 2 + fi;
    dataRow(r, 6, fi % 2 === 0 ? C.white : C.altRow);
    vals.push({ range: `${S}!A${r}`, values: [[fmt_name]] });
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,1,2), cell: { userEnteredValue: { formulaValue: sp(`(${LIB}!$F$8:$F$1008=A${r})*(${LIB}!$B$8:$B$1008<>"")`) }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }}, fields: 'userEnteredValue,userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,2,3), cell: { userEnteredValue: { formulaValue: sp(`(${LIB}!$F$8:$F$1008=A${r})*(${LIB}!$M$8:$M$1008="Finished")`) }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }}, fields: 'userEnteredValue,userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,3,4), cell: { userEnteredValue: { formulaValue: `=IFERROR(ROUND(SUMPRODUCT((${LIB}!$F$8:$F$1008=A${r})*(${LIB}!$M$8:$M$1008="Finished")*${LIB}!$P$8:$P$1008)/C${r},1),"—")` }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '0.0' } }}, fields: 'userEnteredValue,userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,4,5), cell: { userEnteredValue: { formulaValue: `=IFERROR(ROUND(SUMPRODUCT((${LIB}!$F$8:$F$1008=A${r})*(${LIB}!$M$8:$M$1008="Finished")*${LIB}!$Q$8:$Q$1008)/C${r},0),"—")` }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }}, fields: 'userEnteredValue,userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,5,6), cell: { userEnteredValue: { formulaValue: `=IFERROR(B${r}/COUNTA(${LIB}!$B$8:$B$1008),0)` }, userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER', numberFormat: { type: 'PERCENT', pattern: '0%' } }}, fields: 'userEnteredValue,userEnteredFormat' }});
  });

  // Column widths
  [160,80,80,80,80,80,80,80,80,80,80,80,80,80,80,80,80,80].forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Freeze
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' }});

  await batchUpdate(id, fmt, '07-insights format');
  await valuesBatchUpdate(id, vals, '07-insights values');
  console.log('✅  Reading Insights done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
