'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Library Gallery'];
const S = "'Library Gallery'";
const LIB = "'Master Book Library'";

// Genre color map for visual variety in the gallery cards
const GENRE_COLORS = {
  'Fantasy':          C.primary,
  'Mystery':          '#4A6E8C',  // steel blue
  'Science Fiction':  '#2D5F4E',  // teal
  'Literary Fiction': '#7A5C2E',  // brown
  'Romance':          '#8B3A6E',  // magenta-wine
  'Historical Fiction':'#5A4A2E', // aged brown
  'Thriller':         '#3A3A5C',  // dark purple
  'Horror':           '#5C2020',  // dark red
  'Non-Fiction':      '#2E4A5C',  // dark steel
  'Biography':        '#3E5C3A',  // forest
  'Memoir':           '#3E5C3A',
  'Self-Help':        '#5C4A2E',
  'Graphic Novel':    '#4A2E5C',
  'Young Adult':      '#2E5C4A',
  'Other':            C.secText,
};

// Top Favorites to display in gallery — using stable Book IDs
// We'll build a gallery showing favorites and high-rated books
// Gallery layout: 3 books per row, 4 rows per "panel", each card = 4 rows × 4 cols
// Cols A-D = book 1, E-H = book 2, I-L = book 3, M-N = spare

(async () => {
  const fmt  = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,500,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 10, fontFamily: 'Georgia', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Title ──────────────────────────────────────────────────────────────────
  vals.push({ range: `${S}!A1`, values: [['🖼️ Library Gallery']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 42 }, fields: 'pixelSize' }});

  // Subtitle
  vals.push({ range: `${S}!A2`, values: [['A visual showcase of your favourite reads. Stars are drawn from your ratings in the Master Book Library — automatically updated.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.goldTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // ── Gallery stats bar ──────────────────────────────────────────────────────
  vals.push({ range: `${S}!A3`, values: [[
    '💝 Total Favorites',
    `=SUMPRODUCT((${LIB}!$U$8:$U$1008=TRUE)*1)`,
    '⭐ 5-Star Books',
    `=SUMPRODUCT((${LIB}!$P$8:$P$1008=5)*(${LIB}!$M$8:$M$1008="Finished"))`,
    '📚 Currently Showing',
    '=MIN(24,SUMPRODUCT((\'Master Book Library\'!$U$8:$U$1008=TRUE)*1))',
  ]] });
  for (let pair = 0; pair < 3; pair++) {
    const c1 = pair * 4;
    fmt.push({ mergeCells: { range: gridRange(SID,2,3,c1,c1+2), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,2,3,c1,c1+2), cell: { userEnteredFormat: {
      backgroundColor: hex(pair % 2 === 0 ? C.primary : C.secondary),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ mergeCells: { range: gridRange(SID,2,3,c1+2,c1+4), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,2,3,c1+2,c1+4), cell: { userEnteredFormat: {
      backgroundColor: hex(pair % 2 === 0 ? C.wineTint : C.goldTint),
      textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 38 }, fields: 'pixelSize' }});

  // ── Section header: Hall of Favorites ─────────────────────────────────────
  vals.push({ range: `${S}!A4`, values: [['⭐ HALL OF FAVORITES — YOUR HIGHEST-RATED READS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,3,4,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,3,4,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.white), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // ── Book Cards — 3 per row, using VLOOKUP from Library ────────────────────
  // Each card = 4 rows × 4 cols (A-D for col set, E-H for col set 2, I-L for col set 3)
  // Row 1 of card: Genre color banner with book number (VLOOKUP genre → color auto-set by CF)
  // Row 2 of card: Title + Author
  // Row 3 of card: Stars + Pages
  // Row 4 of card: Date Finished + Shelf
  // Cards start at row 5

  // We'll show 24 favorite books (8 rows of 3 cards)
  // Using VLOOKUP with SMALL(IF(Fav=TRUE, ROW()-7)) approach via helper formulas
  // Simpler: use FILTER to get favorite Book IDs, then show them in cards

  // Col positions: Book 1 = A-D (0-3), Book 2 = E-H (4-7), Book 3 = I-L (8-11)
  const COL_SETS = [[0,4], [4,8], [8,12]];
  const CARD_HEIGHT = 5; // rows per card
  const CARDS_PER_ROW = 3;
  const NUM_ROWS = 8; // 8 rows of 3 = 24 books

  // Helper: get the nth favorite book's field
  // We filter favorites and get the nth one using INDEX(FILTER(...), n)
  const favField = (n, libCol) =>
    `=IFERROR(INDEX(FILTER(${LIB}!${libCol}$8:${libCol}$1008,${LIB}!$U$8:$U$1008=TRUE,${LIB}!$M$8:$M$1008="Finished"),${n}),"")`;

  const starRating = (n) =>
    `=IFERROR(REPT("★",INDEX(FILTER(${LIB}!$P$8:$P$1008,${LIB}!$U$8:$U$1008=TRUE,${LIB}!$M$8:$M$1008="Finished"),${n}))&REPT("☆",5-INDEX(FILTER(${LIB}!$P$8:$P$1008,${LIB}!$U$8:$U$1008=TRUE,${LIB}!$M$8:$M$1008="Finished"),${n})),"")`;

  let bookNum = 0;
  for (let rowIdx = 0; rowIdx < NUM_ROWS; rowIdx++) {
    const baseRow = 5 + rowIdx * CARD_HEIGHT; // 1-indexed row

    for (let colIdx = 0; colIdx < CARDS_PER_ROW; colIdx++) {
      bookNum++;
      const c1 = colIdx * 4; // 0-indexed start column for this card (A=0,E=4,I=8)
      const br = baseRow - 1; // 0-indexed base row

      // Card background (all 4 rows × 4 cols)
      fmt.push({ repeatCell: { range: gridRange(SID, br, br+CARD_HEIGHT, c1, c1+4), cell: { userEnteredFormat: {
        backgroundColor: hex(C.white),
        borders: {
          top:    { style: 'SOLID', width: 2, color: hex(C.border) },
          bottom: { style: 'SOLID', width: 2, color: hex(C.border) },
          left:   { style: 'SOLID', width: 2, color: hex(C.border) },
          right:  { style: 'SOLID', width: 2, color: hex(C.border) },
        },
      }}, fields: 'userEnteredFormat' }});

      // Row 1: Genre color banner — Genre text on left, Book number on right
      const genreCol = String.fromCharCode(65 + c1);
      const numCol   = String.fromCharCode(65 + c1 + 3);
      vals.push({ range: `${S}!${genreCol}${baseRow}`, values: [[favField(bookNum, '$D')]] });
      vals.push({ range: `${S}!${numCol}${baseRow}`,   values: [[`#${bookNum}`]] });
      fmt.push({ mergeCells: { range: gridRange(SID, br, br+1, c1, c1+3), mergeType: 'MERGE_ALL' }});
      fmt.push({ repeatCell: { range: gridRange(SID, br, br+1, c1, c1+3), cell: { userEnteredFormat: {
        backgroundColor: hex(C.primary),
        textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 6 },
      }}, fields: 'userEnteredFormat' }});
      fmt.push({ repeatCell: { range: gridRange(SID, br, br+1, c1+3, c1+4), cell: { userEnteredFormat: {
        backgroundColor: hex(C.primary),
        textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.goldTint), fontFamily: 'Arial' },
        horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE', padding: { right: 4 },
      }}, fields: 'userEnteredFormat' }});

      // Row 2: Title (merged 3 cols) + Author (1 col)
      vals.push({ range: `${S}!${genreCol}${baseRow+1}`, values: [[favField(bookNum, '$B')]] });
      vals.push({ range: `${S}!${numCol}${baseRow+1}`, values: [[favField(bookNum, '$C')]] });
      fmt.push({ mergeCells: { range: gridRange(SID, br+1, br+2, c1, c1+3), mergeType: 'MERGE_ALL' }});
      fmt.push({ repeatCell: { range: gridRange(SID, br+1, br+2, c1, c1+3), cell: { userEnteredFormat: {
        backgroundColor: hex(C.wineTint),
        textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.text), fontFamily: 'Georgia' },
        horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 6 }, wrapStrategy: 'WRAP',
      }}, fields: 'userEnteredFormat' }});
      fmt.push({ repeatCell: { range: gridRange(SID, br+1, br+2, c1+3, c1+4), cell: { userEnteredFormat: {
        backgroundColor: hex(C.goldTint),
        textFormat: { italic: true, fontSize: 8, foregroundColor: hex(C.secText), fontFamily: 'Georgia' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
      }}, fields: 'userEnteredFormat' }});

      // Row 3: Star rating (merged 2 cols) + Pages (2 cols)
      const starCol  = String.fromCharCode(65 + c1);
      const pageCol  = String.fromCharCode(65 + c1 + 2);
      vals.push({ range: `${S}!${starCol}${baseRow+2}`, values: [[starRating(bookNum)]] });
      vals.push({ range: `${S}!${pageCol}${baseRow+2}`, values: [[`=IFERROR(INDEX(FILTER(${LIB}!$Q$8:$Q$1008,${LIB}!$U$8:$U$1008=TRUE,${LIB}!$M$8:$M$1008="Finished"),${bookNum})&" pages","")`]] });
      fmt.push({ mergeCells: { range: gridRange(SID, br+2, br+3, c1, c1+2), mergeType: 'MERGE_ALL' }});
      fmt.push({ repeatCell: { range: gridRange(SID, br+2, br+3, c1, c1+2), cell: { userEnteredFormat: {
        backgroundColor: hex(C.white),
        textFormat: { fontSize: 12, foregroundColor: hex(C.warning), fontFamily: 'Arial' },
        horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 6 },
      }}, fields: 'userEnteredFormat' }});
      fmt.push({ mergeCells: { range: gridRange(SID, br+2, br+3, c1+2, c1+4), mergeType: 'MERGE_ALL' }});
      fmt.push({ repeatCell: { range: gridRange(SID, br+2, br+3, c1+2, c1+4), cell: { userEnteredFormat: {
        backgroundColor: hex(C.white),
        textFormat: { fontSize: 9, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
        horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE', padding: { right: 6 },
      }}, fields: 'userEnteredFormat' }});

      // Row 4: Date Finished + Shelf
      const dateCol  = String.fromCharCode(65 + c1);
      const shelfCol = String.fromCharCode(65 + c1 + 2);
      vals.push({ range: `${S}!${dateCol}${baseRow+3}`, values: [[`=IFERROR("Finished: "&TEXT(DATEVALUE(INDEX(FILTER(${LIB}!$O$8:$O$1008,${LIB}!$U$8:$U$1008=TRUE,${LIB}!$M$8:$M$1008="Finished"),${bookNum})),"mmm yyyy"),"")`]] });
      vals.push({ range: `${S}!${shelfCol}${baseRow+3}`, values: [[favField(bookNum, '$L')]] });
      fmt.push({ mergeCells: { range: gridRange(SID, br+3, br+4, c1, c1+2), mergeType: 'MERGE_ALL' }});
      fmt.push({ repeatCell: { range: gridRange(SID, br+3, br+4, c1, c1+2), cell: { userEnteredFormat: {
        backgroundColor: hex(C.altRow),
        textFormat: { fontSize: 8, foregroundColor: hex(C.secText), fontFamily: 'Arial', italic: true },
        horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 6 },
      }}, fields: 'userEnteredFormat' }});
      fmt.push({ mergeCells: { range: gridRange(SID, br+3, br+4, c1+2, c1+4), mergeType: 'MERGE_ALL' }});
      fmt.push({ repeatCell: { range: gridRange(SID, br+3, br+4, c1+2, c1+4), cell: { userEnteredFormat: {
        backgroundColor: hex(C.altRow),
        textFormat: { fontSize: 8, foregroundColor: hex(C.secText), fontFamily: 'Arial', italic: true },
        horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE', padding: { right: 6 },
      }}, fields: 'userEnteredFormat' }});
    }

    // Row heights for this card set (5 rows)
    [32, 36, 30, 22, 8].forEach((h, hi) => {
      fmt.push({ updateDimensionProperties: {
        range: { sheetId: SID, dimension: 'ROWS', startIndex: baseRow - 1 + hi, endIndex: baseRow + hi },
        properties: { pixelSize: h }, fields: 'pixelSize',
      }});
    });
  }

  // ── All-Time Stats Section at bottom ─────────────────────────────────────
  const STAT_ROW = 5 + NUM_ROWS * CARD_HEIGHT + 2;
  vals.push({ range: `${S}!A${STAT_ROW}`, values: [['📖 ALL-TIME READING MILESTONES']] });
  fmt.push({ mergeCells: { range: gridRange(SID,STAT_ROW-1,STAT_ROW,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,STAT_ROW-1,STAT_ROW,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.white), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: STAT_ROW-1, endIndex: STAT_ROW }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  const MILESTONES = [
    ['📚 Total Books in Library', `=COUNTA(${LIB}!$B$8:$B$1008)`],
    ['✅ Books Finished',          `=SUMPRODUCT((${LIB}!$M$8:$M$1008="Finished")*1)`],
    ['📄 Total Pages Read',        `=SUMPRODUCT((${LIB}!$M$8:$M$1008="Finished")*${LIB}!$Q$8:$Q$1008)`],
    ['⭐ Average Rating',          `=IFERROR(ROUND(AVERAGEIF(${LIB}!$M$8:$M$1008,"Finished",${LIB}!$P$8:$P$1008),2),"—")`],
    ['💝 Books Marked Favorite',   `=SUMPRODUCT((${LIB}!$U$8:$U$1008=TRUE)*1)`],
    ['🔁 Books Reread',            `=SUMPRODUCT((${LIB}!$V$8:$V$1008=TRUE)*1)`],
    ['🌟 Wishlist Size',           `=COUNTA('Wishlist'!$B$6:$B$506)`],
    ['📝 Total Reviews Written',   `=COUNTA('Book Review & Notes'!$A$6:$A$506)`],
  ];

  MILESTONES.forEach((ms, mi) => {
    const r = STAT_ROW + 1 + Math.floor(mi/4);
    const c = (mi % 4) * 3;
    if (mi % 4 === 0) {
      fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r-1, endIndex: r }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});
    }
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}${r}`, values: [[ms[0]]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+c+1)}${r}`, values: [[ms[1]]] });
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,c,c+1), cell: { userEnteredFormat: {
      backgroundColor: hex(mi % 2 === 0 ? C.primary : C.secondary),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', padding: { left: 4 },
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,c+1,c+2), cell: { userEnteredFormat: {
      backgroundColor: hex(mi % 2 === 0 ? C.wineTint : C.goldTint),
      textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });

  // Column widths (4 cols per card × 3 cards = 12 cols)
  [60,60,60,60, 60,60,60,60, 60,60,60,60, 60,60].forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Freeze title and stats bar
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 4 } }, fields: 'gridProperties.frozenRowCount' }});

  await batchUpdate(id, fmt, '09-gallery format');
  await valuesBatchUpdate(id, vals, '09-gallery values');
  console.log('✅  Library Gallery done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
