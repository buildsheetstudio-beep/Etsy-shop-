'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Search & Filter'];
const S = "'Search & Filter'";
const LIB = "'Master Book Library'";

(async () => {
  const fmt  = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,150,0,18), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 10, fontFamily: 'Georgia', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Title ──────────────────────────────────────────────────────────────────
  vals.push({ range: `${S}!A1`, values: [['🔍 Search & Filter']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,18), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,18), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 42 }, fields: 'pixelSize' }});

  // Subtitle
  vals.push({ range: `${S}!A2`, values: [['Use the filter controls below to search your library. Select "All" in any dropdown to show all values for that field.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,18), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,18), cell: { userEnteredFormat: {
    backgroundColor: hex(C.goldTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // ── Filter Control Panel (rows 3-8) ───────────────────────────────────────
  vals.push({ range: `${S}!A3`, values: [['FILTER CONTROLS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,2,3,0,10), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Filter labels and input cells
  const FILTERS = [
    ['Genre Filter:', 'D4', 'All'],
    ['Status Filter:', 'D5', 'All'],
    ['Min Rating:', 'D6', 'All'],
    ['Year Finished:', 'D7', 'All'],
    ['Format Filter:', 'D8', 'All'],
  ];
  const filterData = [
    [null,'Genre Filter:',null,'All'],
    [null,'Status Filter:',null,'All'],
    [null,'Min Rating (1-5):',null,'All'],
    [null,'Year Finished:',null,'All'],
    [null,'Format:',null,'All'],
    [null,'Favorite Only?',null,false],
  ];
  filterData.forEach((row, ri) => {
    const r = 4 + ri;
    vals.push({ range: `${S}!A${r}:D${r}`, values: [row] });
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,1,2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE', padding: { right: 8 },
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,2,4), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,2,4), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      borders: { bottom: { style: 'SOLID', width: 2, color: hex(C.secondary) } },
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r-1, endIndex: r }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  });

  // Results count
  vals.push({ range: `${S}!A10`, values: [['RESULTS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,9,10,0,2), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,9,10,0,2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});

  // Match count formula
  fmt.push({ repeatCell: {
    range: gridRange(SID,9,10,2,8),
    cell: {
      userEnteredValue: { formulaValue: `=IFERROR("Showing "&SUMPRODUCT((${LIB}!$B$8:$B$1008<>"")*IF(D4="All",1,${LIB}!$D$8:$D$1008=D4)*IF(D5="All",1,${LIB}!$M$8:$M$1008=D5)*IF(D6="All",1,${LIB}!$P$8:$P$1008>=VALUE(D6))*IF(D7="All",1,YEAR(IFERROR(DATEVALUE(${LIB}!$O$8:$O$1008),0))=VALUE(D7))*IF(D8="All",1,${LIB}!$F$8:$F$1008=D8)*IF(D9=FALSE,1,${LIB}!$U$8:$U$1008=TRUE))&" matching books","No results")` },
      userEnteredFormat: { backgroundColor: hex(C.wineTint), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial' }, horizontalAlignment: 'CENTER' },
    },
    fields: 'userEnteredValue,userEnteredFormat',
  }});
  fmt.push({ mergeCells: { range: gridRange(SID,9,10,2,8), mergeType: 'MERGE_ALL' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 9, endIndex: 10 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // ── Results Header Row 11 ─────────────────────────────────────────────────
  const RESULT_HDR = ['Book ID','Title','Author','Genre','Status','Date Finished','Rating','Total Pages','Progress %','Format','Favorite?','Notes'];
  vals.push({ range: `${S}!A11`, values: [RESULT_HDR] });
  fmt.push({ repeatCell: { range: gridRange(SID,10,11,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // ── FILTER formulas in rows 12-111 (50 rows constrained) ─────────────────
  // Condition: row in library matches all active filters
  const cond = `(${LIB}!$B$8:$B$1008<>"")*IF(D4="All",1,${LIB}!$D$8:$D$1008=D4)*IF(D5="All",1,${LIB}!$M$8:$M$1008=D5)*IF(D6="All",1,${LIB}!$P$8:$P$1008>=VALUE(D6))*IF(D7="All",1,YEAR(IFERROR(DATEVALUE(${LIB}!$O$8:$O$1008),0))=VALUE(D7))*IF(D8="All",1,${LIB}!$F$8:$F$1008=D8)*IF(D9=FALSE,1,${LIB}!$U$8:$U$1008=TRUE)`;
  const cf = (src, fallback = '""') =>
    `=IFERROR(ARRAY_CONSTRAIN(FILTER(${LIB}!${src},${cond}),50,1),${fallback})`;

  const RESULT_COLS = [
    ['A12', cf('$A$8:$A$1008','""')],         // Book ID
    ['B12', cf('$B$8:$B$1008','"No results — try adjusting your filters"')],  // Title
    ['C12', cf('$C$8:$C$1008','""')],          // Author
    ['D12', cf('$D$8:$D$1008','""')],          // Genre
    ['E12', cf('$M$8:$M$1008','""')],          // Status
    ['F12', cf('$O$8:$O$1008','""')],          // Date Finished
    ['G12', cf('$P$8:$P$1008','""')],          // Rating
    ['H12', cf('$Q$8:$Q$1008','""')],          // Total Pages
    ['I12', cf('$S$8:$S$1008','""')],          // Progress %
    ['J12', cf('$F$8:$F$1008','""')],          // Format
    ['K12', cf('$U$8:$U$1008','""')],          // Favorite?
    ['L12', cf('$X$8:$X$1008','""')],          // Notes
  ];

  RESULT_COLS.forEach(([cell, formula]) => {
    vals.push({ range: `${S}!${cell}`, values: [[formula]] });
  });

  // Format result area
  for (let i = 0; i < 50; i++) {
    const bg = i % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,11+i,12+i,0,12), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }
  // Progress % column I (8): percent format
  fmt.push({ repeatCell: { range: gridRange(SID,11,61,8,9), cell: { userEnteredFormat: {
    numberFormat: { type: 'PERCENT', pattern: '0%' },
  }}, fields: 'userEnteredFormat.numberFormat' }});

  // Row heights for results
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 11, endIndex: 111 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // Column widths
  [90,220,150,110,80,90,60,80,70,90,65,200].forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Freeze top 11 rows (title + filters + header)
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 11 } }, fields: 'gridProperties.frozenRowCount' }});

  await batchUpdate(id, fmt, '08-search format');
  await valuesBatchUpdate(id, vals, '08-search values');
  console.log('✅  Search & Filter done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
