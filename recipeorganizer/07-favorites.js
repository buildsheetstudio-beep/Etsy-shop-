'use strict';
const { hex, batchUpdate, valuesBatchUpdate, gridRange, colL, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Favorites & Ratings'];
const S = "'Favorites & Ratings'";

const MRI = "'Master Recipe Index'";
const RECIPE_ROWS = 60; // number of sample recipes in Master Recipe Index

// Master Recipe Index column letters (for range references):
// A=ID B=Name C=MealType D=Category E=Cuisine F=Dietary G=Allergen H=Servings
// I=PrepTime J=CookTime K=TotalTime L=Method M=Difficulty N=Status O=Favorite
// P=Seasonal Q=Freezer R=Kid S=MealPrep T=Rating U=DateCreated V=DateModified
// W=Source X=SourceURL Y=CostBatch Z=CostServing AA=TimesCooked AB=LastCooked

(async () => {
  const reqs = [];
  const vals = [];

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 0, 1, 0, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.blush),
      textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A1`, values: [['⭐ FAVORITES & RATINGS']] });

  // Row 2: subtitle
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 1, 2, 0, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.bg),
      textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A2`, values: [['Track your favorites, ratings, and cooking history across all recipes']] });

  // Row 3: spacer
  reqs.push({ repeatCell: {
    range: gridRange(SID, 2, 3, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.blush) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Row 4: KPI Cards Header ───────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 3, 4, 0, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A4`, values: [['📊 AT A GLANCE']] });

  // ── Rows 5-6: KPI Cards (4 cards across cols A-S) ────────────────────────
  const kpiCards = [
    { label: 'AVG RATING',       formula: `=IFERROR(ROUND(AVERAGEIF(${MRI}!$M$8:$M$2007,">=1"),1),"—")`,    cols: [0,4],  bg: C.butter },
    { label: 'TOTAL FAVORITES',  formula: `=IFERROR(COUNTIF(${MRI}!$O$8:$O$2007,TRUE),"—")`,                 cols: [5,10], bg: C.blush  },
    { label: '5★ RECIPES',       formula: `=IFERROR(COUNTIF(${MRI}!$M$8:$M$2007,5),"—")`,                    cols: [10,15],bg: C.lavender},
    { label: 'IN TESTING',       formula: `=IFERROR(COUNTIF(${MRI}!$AF$8:$AF$2007,"Testing"),"—")`,          cols: [15,20],bg: C.sage   },
  ];
  kpiCards.forEach(({ label, formula, cols: [c1, c2], bg }) => {
    // Label row (row 5, idx 4)
    reqs.push({ mergeCells: { range: gridRange(SID, 4, 5, c1, c2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, 4, 5, c1, c2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!${colL(c1)}5`, values: [[label]] });
    // Value row (row 6, idx 5)
    reqs.push({ mergeCells: { range: gridRange(SID, 5, 6, c1, c2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, 5, 6, c1, c2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.panel),
        textFormat: { bold: true, fontSize: 20, foregroundColor: hex(C.primaryDeep), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!${colL(c1)}6`, values: [[formula]] });
  });

  // ── Row 7: Spacer ─────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID, 6, 7, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.blush) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Rows 8-20: Side-by-side count tables ──────────────────────────────────
  // Left (A-C): STAR RATING DISTRIBUTION  |  Right (E-G): STATUS BREAKDOWN

  // Section header: Star Ratings (A8:C8)
  reqs.push({ mergeCells: { range: gridRange(SID, 7, 8, 0, 4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 7, 8, 0, 4),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.warning),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A8`, values: [['⭐ RATING DISTRIBUTION']] });

  // Star rating rows (rows 9-13, idx 8-12)
  const starRatings = [5, 4, 3, 2, 1];
  const starLabels = ['★★★★★ (5 Stars)', '★★★★☆ (4 Stars)', '★★★☆☆ (3 Stars)', '★★☆☆☆ (2 Stars)', '★☆☆☆☆ (1 Star)'];
  starRatings.forEach((star, i) => {
    const r = 8 + i;
    const bg = i % 2 === 0 ? C.altRow : C.panel;
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 0, 4),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } });
    vals.push({ range: `${S}!A${r+1}:D${r+1}`, values: [[
      starLabels[i],
      `=COUNTIF(${MRI}!$M$8:$M$2007,${star})`,
      `=IFERROR(COUNTIF(${MRI}!$M$8:$M$2007,${star})/COUNTA(${MRI}!$B$8:$B$2007),"")`,
      '',
    ]] });
  });
  // Column headers for star table (row 9, idx 8) - actually put headers above data
  // We skipped headers, values in A9:C13; let's add % col header in a note
  vals.push({ range: `${S}!C9`, values: [['% of total']] }); // overwrite with header text only for row 9

  // Section header: Status Breakdown (E8)
  reqs.push({ mergeCells: { range: gridRange(SID, 7, 8, 5, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 7, 8, 5, 10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.info),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!F8`, values: [['📋 STATUS BREAKDOWN']] });

  const statuses = ['Draft', 'Testing', 'Approved', 'Favorite', 'Seasonal', 'Archived'];
  statuses.forEach((st, i) => {
    const r = 8 + i;
    const bg = i % 2 === 0 ? C.altRow : C.panel;
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 5, 10),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } });
    vals.push({ range: `${S}!F${r+1}:J${r+1}`, values: [[
      st,
      `=COUNTIF(${MRI}!$AF$8:$AF$2007,"${st}")`,
      '', '', '',
    ]] });
  });

  // Favorite flag counts (col L onwards, same area)
  reqs.push({ mergeCells: { range: gridRange(SID, 7, 8, 11, 16), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 7, 8, 11, 16),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.blush),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!L8`, values: [['🏷️ SPECIAL FLAGS']] });

  const flags = [
    ['❤️ Favorites',    `=COUNTIF(${MRI}!$O$8:$O$2007,TRUE)`],
    ['👧 Kid-Friendly', `=COUNTIF(${MRI}!$P$8:$P$2007,TRUE)`],
    ['❄️ Freezer',      `=COUNTIF(${MRI}!$Q$8:$Q$2007,TRUE)`],
    ['🥡 Meal Prep',    `=COUNTIF(${MRI}!$R$8:$R$2007,TRUE)`],
  ];
  flags.forEach(([label, formula], i) => {
    const r = 8 + i;
    const bg = i % 2 === 0 ? C.altRow : C.panel;
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 11, 16),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } });
    vals.push({ range: `${S}!L${r+1}:P${r+1}`, values: [[label, formula, '', '', '']] });
  });

  // ── Row 14: Spacer ────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID, 13, 14, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.blush) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Rows 15-25: TESTING QUEUE ─────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 14, 15, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 14, 15, 0, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.warning),
      textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A15`, values: [['🧪 TESTING QUEUE — Recipes Currently in Testing Status']] });

  // Column headers row 16 (idx 15)
  const testHdrs = ['Recipe ID', 'Recipe Name', 'Category', 'Difficulty', 'Rating', 'Status', 'Testing Notes'];
  testHdrs.forEach((h, i) => {
    reqs.push({ repeatCell: {
      range: gridRange(SID, 15, 16, i, i+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.sage),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
  });
  vals.push({ range: `${S}!A16:G16`, values: [testHdrs] });

  // Testing recipe rows: FILTER by Status="Testing" (rows 17-26, idx 16-25, 10 rows)
  const TEST_ROWS = 10;
  const testFilter = `${MRI}!$AF$8:$AF$2007="Testing"`;
  for (let i = 0; i < TEST_ROWS; i++) {
    const r = 16 + i;
    const bg = i % 2 === 0 ? C.altRow : C.panel;
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 0, 7),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } });
  }
  const testCols = ['A', 'B', 'C', 'L', 'M', 'AF'];
  testCols.forEach((mriCol, ci) => {
    const formulas = Array.from({length: TEST_ROWS}, (_, i) =>
      [`=IFERROR(INDEX(FILTER(${MRI}!$${mriCol}$8:$${mriCol}$2007,${testFilter}),${i+1}),"")`]);
    vals.push({ range: `${S}!${colL(ci)}17:${colL(ci)}${16+TEST_ROWS}`, values: formulas });
  });
  // G col: leave blank for manual testing notes
  reqs.push({ repeatCell: {
    range: gridRange(SID, 16, 16 + TEST_ROWS, 6, 7),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.input),
      textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
      verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  } });

  // ── Row 27: Spacer ────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID, 26, 27, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.blush) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Rows 28+: Full Ratings & Favorites Table ──────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 27, 28, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 27, 28, 0, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A28`, values: [['📋 ALL RECIPES — RATINGS & FAVORITES']] });

  // Column headers row 29 (idx 28)
  const allHdrs = ['Recipe ID', 'Recipe Name', 'Category', 'Difficulty', 'Fav?', 'Rating', 'Status', 'Times Cooked', 'Last Cooked', 'Source'];
  allHdrs.forEach((h, i) => {
    reqs.push({ repeatCell: {
      range: gridRange(SID, 28, 29, i, i+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.sage),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
  });
  vals.push({ range: `${S}!A29:J29`, values: [allHdrs] });

  // All recipe rows (rows 30-89, idx 29-88, 60 rows using INDEX from Master Recipe Index)
  // INDEX('Master Recipe Index'!$X$8:$X$2007, ROW()-29) at row 30 → INDEX(range, 1) = row 8 of MRI
  const allCols = ['A', 'B', 'C', 'L', 'O', 'M', 'AF', 'AE', 'AD', 'AA'];
  for (let i = 0; i < RECIPE_ROWS; i++) {
    const r = 29 + i;
    const bg = i % 2 === 0 ? C.altRow : C.panel;
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 0, 10),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } });
  }
  allCols.forEach((mriCol, ci) => {
    const formulas = Array.from({length: RECIPE_ROWS}, (_, i) =>
      [`=IFERROR(INDEX(${MRI}!$${mriCol}$8:$${mriCol}$2007,${i+1}),"")`]);
    vals.push({ range: `${S}!${colL(ci)}30:${colL(ci)}${29+RECIPE_ROWS}`, values: formulas });
  });

  // ── Column Widths ─────────────────────────────────────────────────────────
  const widths = [85, 215, 90, 80, 55, 65, 80, 85, 95, 130, 40, 100, 40, 40, 40, 40, 40, 40, 40, 40];
  widths.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    } });
  });

  // ── Row Heights ───────────────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 3 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 26 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 90 },
    properties: { pixelSize: 22 }, fields: 'pixelSize',
  } });

  // ── Freeze ────────────────────────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } },
    fields: 'gridProperties.frozenRowCount',
  } });

  // ── Tab Color ─────────────────────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, tabColorStyle: { rgbColor: hex(C.blush) } },
    fields: 'tabColorStyle',
  } });

  await batchUpdate(id, reqs, 'fav-fmt');
  await valuesBatchUpdate(id, vals, 'fav-data');
  console.log('✓ Favorites & Ratings tab complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
