'use strict';
const { hex, batchUpdate, valuesBatchUpdate, gridRange, colL, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Search & Filter'];
const S = "'Search & Filter'";

const MRI = "'Master Recipe Index'";

// Filter selector cell references (GS col C = index 2)
// C5=MealType  C6=Category  C7=Cuisine  C8=Difficulty
// C9=MaxTime   C10=Dietary  C11=FavOnly  C12=MinRating
const CONDITIONS = [
  `((LEN($C$5)=0)+(${MRI}!$C$8:$C$2007=$C$5))`,   // Meal Type (col C)
  `((LEN($C$6)=0)+(${MRI}!$D$8:$D$2007=$C$6))`,   // Category (col D)
  `((LEN($C$7)=0)+(${MRI}!$E$8:$E$2007=$C$7))`,   // Cuisine (col E)
  `((LEN($C$8)=0)+(${MRI}!$M$8:$M$2007=$C$8))`,   // Difficulty (col M)
  `(($C$9="")+(($C$9<>"")*1)*(${MRI}!$K$8:$K$2007<=$C$9))`, // Max Total Time (col K)
  `((LEN($C$10)=0)+(${MRI}!$F$8:$F$2007=$C$10))`, // Dietary Tag (col F)
  `((NOT($C$11))+(($C$11)*(${MRI}!$O$8:$O$2007=TRUE)))`, // Favorites Only (col O)
  `(($C$12="")+(($C$12<>"")*1)*(${MRI}!$T$8:$T$2007>=$C$12))`, // Min Rating (col T)
  `(${MRI}!$B$8:$B$2007<>"")`, // non-blank recipe
];
const COND_STR = CONDITIONS.join(' * \n  ');

const buildFilter = (mriCol, N) =>
  `=IFERROR(INDEX(FILTER(${MRI}!$${mriCol}$8:$${mriCol}$2007,\n  ${COND_STR}\n), ${N}),"")`;

const RESULT_ROWS = 100;
const RESULTS_START = 19; // 0-indexed (GS row 20)

// Count formula using SUMPRODUCT
const countConds = CONDITIONS.map(c => `(${c}>=1)`).join(' * ');
const countFormula = `=IFERROR(SUMPRODUCT(${countConds}),0)`;
const avgRatingFormula = `=IFERROR(AVERAGEIFS(${MRI}!$T$8:$T$2007,${MRI}!$B$8:$B$2007,"<>",${MRI}!$T$8:$T$2007,">=1"),"")`; // simplified
const avgTimeFormula   = `=IFERROR(AVERAGEIF(${MRI}!$B$8:$B$2007,"<>",${MRI}!$K$8:$K$2007),"")`;

// Quick Search Presets (shown as reference - user reads and enters values manually)
const PRESETS = [
  { name: '⚡ Weeknight Dinners',     mealType:'Dinner',    category:'',       maxTime:45,  minRating:3,  dietary:'',            favOnly:false },
  { name: '🌅 Quick Breakfasts',      mealType:'Breakfast', category:'',       maxTime:20,  minRating:'', dietary:'',            favOnly:false },
  { name: '🌱 Vegan Dinners',         mealType:'Dinner',    category:'',       maxTime:'',  minRating:'', dietary:'Vegan',       favOnly:false },
  { name: '❤️ My Favorites',          mealType:'',          category:'',       maxTime:'',  minRating:'', dietary:'',            favOnly:true  },
  { name: '⭐ Top-Rated (4★+)',       mealType:'',          category:'',       maxTime:'',  minRating:4,  dietary:'',            favOnly:false },
  { name: '❄️ Freezer-Friendly',      mealType:'',          category:'Freezer Meals', maxTime:'', minRating:'', dietary:'',     favOnly:false },
  { name: '👧 Kid-Friendly Snacks',   mealType:'Snack',     category:'',       maxTime:'',  minRating:'', dietary:'',            favOnly:false },
  { name: '🥗 Gluten-Free Options',   mealType:'',          category:'',       maxTime:'',  minRating:'', dietary:'Gluten-Free', favOnly:false },
];

(async () => {
  const reqs = [];
  const vals = [];

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 0, 1, 0, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.info),
      textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A1`, values: [['🔍 SEARCH & FILTER']] });

  // Row 2
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
  vals.push({ range: `${S}!A2`, values: [['Choose filter values below — leave blank to ignore that filter — results update automatically']] });

  // Row 3: spacer
  reqs.push({ repeatCell: {
    range: gridRange(SID, 2, 3, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.info) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Row 4: Filter Controls Header ─────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, 0, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 3, 4, 0, 5),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A4`, values: [['🎛️ FILTER CONTROLS']] });

  // Quick Presets header (right panel G4)
  reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, 6, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 3, 4, 6, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.sage),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!G4`, values: [['⚡ QUICK SEARCH PRESETS — Reference (enter values in filter cells manually)']] });

  // ── Filter Input Rows (5-12) ──────────────────────────────────────────────
  const filterDefs = [
    { label: 'MEAL TYPE',        dropSrc: "='Reference Data'!$A$4:$A$14",  type: 'dropdown' },
    { label: 'CATEGORY',         dropSrc: "='Reference Data'!$B$4:$B$32",  type: 'dropdown' },
    { label: 'CUISINE',          dropSrc: "='Reference Data'!$C$4:$C$22",  type: 'dropdown' },
    { label: 'DIFFICULTY',       dropSrc: "='Reference Data'!$G$4:$G$7",   type: 'dropdown' },
    { label: 'MAX TIME (min)',    dropSrc: null,                             type: 'number'   },
    { label: 'DIETARY TAG',      dropSrc: "='Reference Data'!$D$4:$D$18",  type: 'dropdown' },
    { label: 'FAVORITES ONLY?',  dropSrc: null,                             type: 'checkbox' },
    { label: 'MIN RATING (1-5)', dropSrc: "='Reference Data'!$I$4:$I$8",   type: 'dropdown' },
  ];

  filterDefs.forEach(({ label, dropSrc, type }, i) => {
    const r = 4 + i; // 0-indexed (GS rows 5-12)

    // Label (A:B merged)
    reqs.push({ mergeCells: { range: gridRange(SID, r, r+1, 0, 2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 0, 2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.secondary),
        textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!A${r+1}`, values: [[label]] });

    // Input cell (C:E merged)
    if (type === 'checkbox') {
      reqs.push({ repeatCell: {
        range: gridRange(SID, r, r+1, 2, 3),
        cell: { userEnteredFormat: {
          backgroundColor: hex(C.input),
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        } },
        fields: 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment)',
      } });
      reqs.push({ setDataValidation: {
        range: gridRange(SID, r, r+1, 2, 3),
        rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true, strict: true },
      } });
      vals.push({ range: `${S}!C${r+1}`, values: [[false]] });
    } else {
      reqs.push({ mergeCells: { range: gridRange(SID, r, r+1, 2, 5), mergeType: 'MERGE_ALL' } });
      reqs.push({ repeatCell: {
        range: gridRange(SID, r, r+1, 2, 5),
        cell: { userEnteredFormat: {
          backgroundColor: hex(C.input),
          textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.text), fontFamily: 'Arial' },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          borders: {
            top:    { style: 'SOLID', color: hex(C.primary) },
            bottom: { style: 'SOLID', color: hex(C.primary) },
            left:   { style: 'SOLID', color: hex(C.primary) },
            right:  { style: 'SOLID', color: hex(C.primary) },
          },
        } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,borders)',
      } });
      if (dropSrc) {
        reqs.push({ setDataValidation: {
          range: gridRange(SID, r, r+1, 2, 5),
          rule: {
            condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: dropSrc }] },
            showCustomUi: true, strict: false,
          },
        } });
      }
    }
  });

  // ── Quick Presets Table (right side, rows 5-12, cols G-T) ─────────────────
  const presetHdrs = ['Preset', 'Meal Type', 'Category', 'Max Time', 'Dietary', 'Fav Only?', 'Min Rating'];
  presetHdrs.forEach((h, ci) => {
    reqs.push({ repeatCell: {
      range: gridRange(SID, 4, 5, 6+ci, 7+ci),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.sage),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
  });
  vals.push({ range: `${S}!G5:M5`, values: [presetHdrs] });

  PRESETS.forEach(({ name, mealType, category, maxTime, minRating, dietary, favOnly }, i) => {
    const r = 5 + i;
    const bg = i % 2 === 0 ? C.altRow : C.panel;
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 6, 13),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } });
    vals.push({ range: `${S}!G${r+1}:M${r+1}`, values: [[
      name,
      mealType || '(any)',
      category || '(any)',
      maxTime  || '(any)',
      dietary  || '(any)',
      favOnly ? 'YES' : '(any)',
      minRating || '(any)',
    ]] });
  });

  // ── Rows 13-16: Summary Cards ─────────────────────────────────────────────
  // Row 13: spacer within filter area
  reqs.push({ repeatCell: {
    range: gridRange(SID, 12, 13, 0, 5),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // Row 14: RESULTS SUMMARY header
  reqs.push({ mergeCells: { range: gridRange(SID, 13, 14, 0, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 13, 14, 0, 5),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.info),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A14`, values: [['📊 RESULTS SUMMARY']] });

  // Rows 15-16: Count, Avg Rating, Avg Time cards
  const summaryCards = [
    { label: 'MATCHING RECIPES', formula: countFormula,    cols: [0, 2] },
    { label: 'AVG COOK TIME (min)', formula: `=IFERROR(ROUND(AVERAGEIF(${MRI}!$B$8:$B$2007,"<>",${MRI}!$K$8:$K$2007),0),"")`, cols: [2, 5] },
  ];
  summaryCards.forEach(({ label, formula, cols: [c1, c2] }) => {
    reqs.push({ mergeCells: { range: gridRange(SID, 14, 15, c1, c2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, 14, 15, c1, c2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.butter),
        textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!${colL(c1)}15`, values: [[label]] });

    reqs.push({ mergeCells: { range: gridRange(SID, 15, 16, c1, c2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, 15, 16, c1, c2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.panel),
        textFormat: { bold: true, fontSize: 20, foregroundColor: hex(C.primaryDeep), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!${colL(c1)}16`, values: [[formula]] });
  });

  // ── Row 17: Spacer ────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID, 16, 17, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.info) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Row 18: Results Section Header ────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 17, 18, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 17, 18, 0, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A18`, values: [['📋 SEARCH RESULTS']] });

  // ── Row 19: Column Headers ────────────────────────────────────────────────
  const resultHdrs = ['Recipe ID', 'Recipe Name', 'Meal Type', 'Category', 'Cuisine', 'Difficulty', 'Total Time', 'Dietary Tags', 'Fav?', 'Rating', 'Status'];
  resultHdrs.forEach((h, ci) => {
    reqs.push({ repeatCell: {
      range: gridRange(SID, 18, 19, ci, ci+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.sage),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
  });
  vals.push({ range: `${S}!A19:K19`, values: [resultHdrs] });

  // ── Result Data Rows (20 to 119, idx 19 to 118, RESULT_ROWS rows) ─────────
  for (let i = 0; i < RESULT_ROWS; i++) {
    const r = RESULTS_START + i;
    const bg = i % 2 === 0 ? C.altRow : C.panel;
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 0, 11),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } });
  }

  // Filtered result formulas column by column
  const resultCols = ['A', 'B', 'C', 'D', 'E', 'M', 'K', 'F', 'O', 'T', 'N'];
  resultCols.forEach((mriCol, ci) => {
    const formulas = Array.from({length: RESULT_ROWS}, (_, i) => {
      const N = i + 1;
      return [`=IFERROR(INDEX(FILTER(${MRI}!$${mriCol}$8:$${mriCol}$2007,${COND_STR}),${N}),"")`];
    });
    vals.push({ range: `${S}!${colL(ci)}20:${colL(ci)}${19+RESULT_ROWS}`, values: formulas });
  });

  // ── Column Widths ─────────────────────────────────────────────────────────
  const widths = [95, 220, 95, 100, 95, 75, 75, 120, 50, 60, 80, 40, 110, 90, 40, 60, 40, 40, 40, 40];
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
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 17 },
    properties: { pixelSize: 26 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 17, endIndex: 120 },
    properties: { pixelSize: 22 }, fields: 'pixelSize',
  } });

  // ── Freeze ────────────────────────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, gridProperties: { frozenRowCount: 19 } },
    fields: 'gridProperties.frozenRowCount',
  } });

  // ── Tab Color ─────────────────────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, tabColorStyle: { rgbColor: hex(C.info) } },
    fields: 'tabColorStyle',
  } });

  await batchUpdate(id, reqs, 'search-fmt');
  await valuesBatchUpdate(id, vals, 'search-data');
  console.log('✓ Search & Filter tab complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
