'use strict';
const { hex, batchUpdate, valuesBatchUpdate, gridRange, colL, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Recipe Dashboard'];
const S = "'Recipe Dashboard'";

const MRI = "'Master Recipe Index'";
// MRI column letters: A=ID B=Name C=MealType D=Category E=Cuisine F=Dietary G=Allergen
// H=Servings I=PrepTime J=CookTime K=TotalTime L=Method M=Difficulty N=Status
// O=Favorite P=Seasonal Q=Freezer R=Kid S=MealPrep T=Rating U=DateCreated V=DateModified
// W=Source X=URL Y=CostBatch Z=CostServing AA=TimesCooked AB=LastCooked

// Featured recipe index formula (cycles daily through 60 recipes)
const FEAT_IDX = `MOD(TODAY()-DATE(2026,1,1),COUNTA(${MRI}!$B$8:$B$2007))+1`;
const featField = (mriColLetter) =>
  `=IFERROR(INDEX(${MRI}!$${mriColLetter}$8:$${mriColLetter}$2007,${FEAT_IDX}),"")`;

const MEAL_TYPES   = ['Breakfast','Brunch','Lunch','Dinner','Snack','Appetizer','Side Dish','Dessert','Beverage','Sauce / Dressing','Other'];
const DIFFICULTIES = ['Very Easy','Easy','Moderate','Advanced'];
const CUISINES     = ['American','Italian','Mexican','Mediterranean','Greek','French','Spanish','Indian','Chinese','Japanese','Korean','Thai','Vietnamese','Middle Eastern','Caribbean','Latin American','African','Fusion','Other'];

(async () => {
  const reqs = [];
  const vals = [];

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 0, 1, 0, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold: true, fontSize: 18, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A1`, values: [['🍴 RECIPE DASHBOARD']] });

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
  vals.push({ range: `${S}!A2`, values: [['Your complete recipe library overview — all stats update automatically as you add recipes']] });

  // Row 3: spacer
  reqs.push({ repeatCell: {
    range: gridRange(SID, 2, 3, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.primary) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Row 4: KPI Banner ─────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 3, 4, 0, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A4`, values: [['📊 LIBRARY AT A GLANCE']] });

  // ── Rows 5-8: KPI Cards (2 rows × 4 cards) ───────────────────────────────
  const kpiRows = [
    [
      { label: 'TOTAL RECIPES',   formula: `=IFERROR(COUNTA(${MRI}!$B$8:$B$2007),0)`,                              bg: C.butter   },
      { label: 'TOTAL FAVORITES', formula: `=IFERROR(COUNTIF(${MRI}!$O$8:$O$2007,TRUE),0)`,                        bg: C.blush    },
      { label: 'AVG RATING ★',   formula: `=IFERROR(ROUND(AVERAGEIF(${MRI}!$T$8:$T$2007,">=1"),1),"—")`,          bg: C.lavender },
      { label: 'UNIQUE CUISINES', formula: `=IFERROR(SUMPRODUCT(1/COUNTIF(${MRI}!$E$8:$E$67,${MRI}!$E$8:$E$67)),0)`, bg: C.sage  },
    ],
    [
      { label: 'IN TESTING',     formula: `=IFERROR(COUNTIF(${MRI}!$N$8:$N$2007,"Testing"),0)`,                   bg: C.warning  },
      { label: '5★ RECIPES',     formula: `=IFERROR(COUNTIF(${MRI}!$T$8:$T$2007,5),0)`,                           bg: C.butter   },
      { label: 'AVG TOTAL TIME', formula: `=IFERROR(ROUND(AVERAGEIF(${MRI}!$B$8:$B$2007,"<>",${MRI}!$K$8:$K$2007),0)&" min","—")`, bg: C.info },
      { label: 'FREEZER MEALS',  formula: `=IFERROR(COUNTIF(${MRI}!$Q$8:$Q$2007,TRUE),0)`,                        bg: C.sage     },
    ],
  ];

  kpiRows.forEach((cards, rowSet) => {
    const labelRow = 4 + rowSet * 2; // idx 4 and 6
    const valRow   = 5 + rowSet * 2; // idx 5 and 7
    cards.forEach(({ label, formula, bg }, ci) => {
      const c1 = ci * 5, c2 = (ci + 1) * 5;
      reqs.push({ mergeCells: { range: gridRange(SID, labelRow, labelRow+1, c1, c2), mergeType: 'MERGE_ALL' } });
      reqs.push({ repeatCell: {
        range: gridRange(SID, labelRow, labelRow+1, c1, c2),
        cell: { userEnteredFormat: {
          backgroundColor: hex(bg),
          textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      } });
      vals.push({ range: `${S}!${colL(c1)}${labelRow+1}`, values: [[label]] });

      reqs.push({ mergeCells: { range: gridRange(SID, valRow, valRow+1, c1, c2), mergeType: 'MERGE_ALL' } });
      reqs.push({ repeatCell: {
        range: gridRange(SID, valRow, valRow+1, c1, c2),
        cell: { userEnteredFormat: {
          backgroundColor: hex(C.panel),
          textFormat: { bold: true, fontSize: 22, foregroundColor: hex(C.primaryDeep), fontFamily: 'Arial' },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      } });
      vals.push({ range: `${S}!${colL(c1)}${valRow+1}`, values: [[formula]] });
    });
  });

  // ── Row 9: Spacer ─────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID, 8, 9, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.primary) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Row 10: Three Panel Headers ───────────────────────────────────────────
  const panelHdrs = [
    { label: '🌟 FEATURED TODAY', c1: 0, c2: 5, bg: C.lavender, fg: C.text },
    { label: '🕐 RECENTLY ADDED', c1: 6, c2: 12, bg: C.sage, fg: C.white },
    { label: '❤️ TOP-RATED FAVORITES', c1: 13, c2: 20, bg: C.blush, fg: C.text },
  ];
  // Spacer col 5, col 12
  [5, 12].forEach(ci => {
    reqs.push({ repeatCell: {
      range: gridRange(SID, 9, 25, ci, ci+1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
      fields: 'userEnteredFormat(backgroundColor)',
    } });
  });

  panelHdrs.forEach(({ label, c1, c2, bg, fg }) => {
    reqs.push({ mergeCells: { range: gridRange(SID, 9, 10, c1, c2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, 9, 10, c1, c2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { bold: true, fontSize: 10, foregroundColor: hex(fg), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!${colL(c1)}10`, values: [[label]] });
  });

  // ── Rows 11-18: Three-Panel Content ──────────────────────────────────────
  // LEFT PANEL: Featured Recipe (cols 0-4, A-E)
  const featFields = [
    { label: '📖', formula: featField('B'), fontSize: 14, bold: true, fg: C.primaryDeep },
    { label: 'Category', formula: featField('D'), fontSize: 9,  bold: false, fg: C.text },
    { label: 'Meal Type', formula: featField('C'), fontSize: 9,  bold: false, fg: C.text },
    { label: 'Prep/Cook', formula: `=IFERROR("⏱ "&INDEX(${MRI}!$I$8:$I$2007,${FEAT_IDX})&" + "&INDEX(${MRI}!$J$8:$J$2007,${FEAT_IDX})&" min","")`, fontSize: 9, bold: false, fg: C.text },
    { label: 'Difficulty', formula: featField('M'), fontSize: 9, bold: false, fg: C.text },
    { label: 'Rating', formula: `=IFERROR("★ "&INDEX(${MRI}!$T$8:$T$2007,${FEAT_IDX})&" / 5","")`, fontSize: 9, bold: false, fg: C.warning },
    { label: 'Dietary', formula: featField('F'), fontSize: 9, bold: false, fg: C.text },
    { label: 'Description', formula: featField('AE'), fontSize: 8, bold: false, fg: C.secText },
  ];
  featFields.forEach(({ label, formula, fontSize, bold, fg }, i) => {
    const r = 10 + i;
    const bg = i === 0 ? C.lavender : (i % 2 === 0 ? C.altRow : C.panel);
    // Label col (A)
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 0, 1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.secondary),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!A${r+1}`, values: [[label]] });
    // Value (B-E merged)
    reqs.push({ mergeCells: { range: gridRange(SID, r, r+1, 1, 5), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 1, 5),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { bold, fontSize, foregroundColor: hex(fg), fontFamily: 'Arial' },
        horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        wrapStrategy: i === 7 ? 'WRAP' : 'OVERFLOW_CELL',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    } });
    vals.push({ range: `${S}!B${r+1}`, values: [[formula]] });
  });

  // MIDDLE PANEL: Recently Added (cols 6-11, G-L, 5 recipes)
  const recAddHdrs = ['Recipe Name', 'Date Added', 'Rating'];
  [7, 8, 9].forEach((ci, i) => {
    reqs.push({ repeatCell: {
      range: gridRange(SID, 10, 11, ci, ci+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.sage),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
  });
  vals.push({ range: `${S}!H11:J11`, values: [recAddHdrs] });

  const totalRecipes = 60;
  for (let i = 0; i < 5; i++) {
    const r = 11 + i;
    const mriRow = totalRecipes - i; // last 5 in MRI order
    const bg = i % 2 === 0 ? C.altRow : C.panel;
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 7, 11),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } });
    vals.push({ range: `${S}!H${r+1}:J${r+1}`, values: [[
      `=IFERROR(INDEX(${MRI}!$B$8:$B$2007,${mriRow}),"")`,
      `=IFERROR(TEXT(INDEX(${MRI}!$U$8:$U$2007,${mriRow}),"mmm d, yyyy"),"")`,
      `=IFERROR(INDEX(${MRI}!$T$8:$T$2007,${mriRow}),"")`,
    ]] });
  }

  // RIGHT PANEL: Top-Rated Favorites (cols 13-19, N-T, 5 recipes)
  const topFavHdrs = ['Recipe Name', 'Rating', 'Category'];
  [14, 15, 16].forEach((ci, i) => {
    reqs.push({ repeatCell: {
      range: gridRange(SID, 10, 11, ci, ci+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.blush),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.text), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
  });
  vals.push({ range: `${S}!O11:Q11`, values: [topFavHdrs] });

  for (let i = 0; i < 5; i++) {
    const r = 11 + i;
    const N = i + 1;
    const bg = i % 2 === 0 ? C.altRow : C.panel;
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 14, 19),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } });
    const filt = `${MRI}!$T$8:$T$2007=5,${MRI}!$O$8:$O$2007=TRUE`;
    vals.push({ range: `${S}!O${r+1}:Q${r+1}`, values: [[
      `=IFERROR(INDEX(FILTER(${MRI}!$B$8:$B$2007,${filt}),${N}),"")`,
      `=IFERROR(INDEX(FILTER(${MRI}!$T$8:$T$2007,${filt}),${N}),"")`,
      `=IFERROR(INDEX(FILTER(${MRI}!$D$8:$D$2007,${filt}),${N}),"")`,
    ]] });
  }

  // ── Row 19: Spacer ────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID, 18, 19, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.primary) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Row 20: Recently Cooked Header ────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 19, 20, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 19, 20, 0, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.sage),
      textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A20`, values: [['🍳 RECENTLY COOKED']] });

  // Row 21: headers
  const cookedHdrs = ['Recipe Name', 'Category', 'Difficulty', 'Last Cooked', 'Times Cooked', 'Rating'];
  cookedHdrs.forEach((h, ci) => {
    reqs.push({ repeatCell: {
      range: gridRange(SID, 20, 21, ci, ci+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.sageDeep),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
  });
  vals.push({ range: `${S}!A21:F21`, values: [cookedHdrs] });

  // Rows 22-26: 5 recently cooked (sorted by last cooked date desc using LARGE/MATCH)
  for (let i = 0; i < 5; i++) {
    const r = 21 + i;
    const N = i + 1;
    const bg = i % 2 === 0 ? C.altRow : C.panel;
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 0, 6),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } });
    // Use LARGE on AB (LastCooked) to get Nth most recent date, then INDEX/MATCH to find recipe
    const dateFormula = `IFERROR(LARGE(${MRI}!$AB$8:$AB$2007,${N}),0)`;
    const matchFormula = `IFERROR(MATCH(${dateFormula},${MRI}!$AB$8:$AB$2007,0),1)`;
    vals.push({ range: `${S}!A${r+1}:F${r+1}`, values: [[
      `=IFERROR(INDEX(${MRI}!$B$8:$B$2007,${matchFormula}),"")`,
      `=IFERROR(INDEX(${MRI}!$D$8:$D$2007,${matchFormula}),"")`,
      `=IFERROR(INDEX(${MRI}!$M$8:$M$2007,${matchFormula}),"")`,
      `=IFERROR(TEXT(${dateFormula},"mmm d, yyyy"),"")`,
      `=IFERROR(INDEX(${MRI}!$AA$8:$AA$2007,${matchFormula}),"")`,
      `=IFERROR(INDEX(${MRI}!$T$8:$T$2007,${matchFormula}),"")`,
    ]] });
  }

  // ── Row 27: Spacer ────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID, 26, 27, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.primary) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Row 28: Chart Source Data Header ─────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 27, 28, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 27, 28, 0, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A28`, values: [['📊 CHART SOURCE DATA (auto-calculated — do not edit)']] });

  // Helper: write a count table
  const writeCountTable = (labelCol, countCol, rowStart, header, items, mriFilterCol) => {
    const rIdx = rowStart;
    reqs.push({ mergeCells: { range: gridRange(SID, rIdx, rIdx+1, labelCol, labelCol+2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, rIdx, rIdx+1, labelCol, labelCol+2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.secondary),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!${colL(labelCol)}${rIdx+1}`, values: [[header]] });

    items.forEach((item, i) => {
      const r = rIdx + 1 + i;
      const bg = i % 2 === 0 ? C.altRow : C.panel;
      reqs.push({ repeatCell: {
        range: gridRange(SID, r, r+1, labelCol, labelCol+2),
        cell: { userEnteredFormat: {
          backgroundColor: hex(bg),
          textFormat: { fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) },
          verticalAlignment: 'MIDDLE',
        } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      } });
      vals.push({ range: `${S}!${colL(labelCol)}${r+1}`, values: [[item]] });
      vals.push({ range: `${S}!${colL(countCol)}${r+1}`, values: [[
        `=COUNTIF(${MRI}!$${mriFilterCol}$8:$${mriFilterCol}$2007,"${item.replace(/"/g,'""')}")`
      ]] });
    });
  };

  writeCountTable(0, 1, 28, 'MEAL TYPE', MEAL_TYPES, 'C');
  writeCountTable(3, 4, 28, 'DIFFICULTY', DIFFICULTIES, 'M');
  writeCountTable(6, 7, 28, 'CUISINE', CUISINES, 'E');

  // ── Column Widths ─────────────────────────────────────────────────────────
  const widths = [80, 160, 100, 90, 90, 20, 75, 185, 90, 60, 20, 20, 20, 20, 185, 70, 110, 40, 40, 40];
  widths.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    } });
  });

  // ── Row Heights ───────────────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 44 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 4 },
    properties: { pixelSize: 22 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 },
    properties: { pixelSize: 22 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 },
    properties: { pixelSize: 38 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 },
    properties: { pixelSize: 22 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 8 },
    properties: { pixelSize: 38 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 8, endIndex: 10 },
    properties: { pixelSize: 10 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 10, endIndex: 50 },
    properties: { pixelSize: 22 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 10, endIndex: 11 },
    properties: { pixelSize: 26 }, fields: 'pixelSize',
  } });
  // Featured recipe description row taller
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 17, endIndex: 18 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  } });

  // ── Freeze ────────────────────────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } },
    fields: 'gridProperties.frozenRowCount',
  } });

  // ── Tab Color ─────────────────────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, tabColorStyle: { rgbColor: hex(C.primary) } },
    fields: 'tabColorStyle',
  } });

  await batchUpdate(id, reqs, 'dash-fmt');
  await valuesBatchUpdate(id, vals, 'dash-data');
  console.log('✓ Recipe Dashboard tab complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
