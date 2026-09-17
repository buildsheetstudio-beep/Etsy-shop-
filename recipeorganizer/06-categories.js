'use strict';
const { hex, batchUpdate, valuesBatchUpdate, gridRange, colL, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Categories & Tags'];
const S = "'Categories & Tags'";

// Master Recipe Index columns used in FILTER/VLOOKUP:
// A(1)=ID, B(2)=Name, C(3)=MealType, D(4)=Category, E(5)=Cuisine
// F(6)=Dietary, G(7)=Allergen, M(13)=Difficulty, N(14)=Status, T(20)=Rating

const MRI = "'Master Recipe Index'";

// Browser display columns: A=RecipeID B=Name C=MealType D=Cuisine E=Difficulty F=Rating G=Status
const BROWSER_COLS = [
  { header: 'Recipe ID',   w: 85,  filterCol: 'A', label: 'Recipe ID'   },
  { header: 'Recipe Name', w: 220, filterCol: 'B', label: 'Recipe Name' },
  { header: 'Meal Type',   w: 90,  filterCol: 'E', label: 'Meal Type'   },
  { header: 'Cuisine',     w: 95,  filterCol: 'F', label: 'Cuisine'     },
  { header: 'Difficulty',  w: 80,  filterCol: 'L', label: 'Difficulty'  },
  { header: 'Rating',      w: 60,  filterCol: 'M', label: 'Rating'      },
  { header: 'Status',      w: 75,  filterCol: 'AF', label: 'Status'     },
];

const MEAL_TYPES   = ['Breakfast','Brunch','Lunch','Dinner','Snack','Appetizer','Side Dish','Dessert','Beverage','Sauce / Dressing','Other'];
const CUISINES     = ['American','Italian','Mexican','Mediterranean','Greek','French','Spanish','Indian','Chinese','Japanese','Korean','Thai','Vietnamese','Middle Eastern','Caribbean','Latin American','African','Fusion','Other'];
const DIFFICULTIES = ['Very Easy','Easy','Moderate','Advanced'];
const DIETARY_TAGS = ['No Restriction','Vegetarian','Vegan','Pescatarian','Gluten-Free','Dairy-Free','Nut-Free','Egg-Free','Soy-Free','Halal','Kosher','Low Sodium','Low Sugar','High Protein','Other'];

// Summary table positions (row 0-indexed, col 0-indexed)
// Meal Types: col 0-1, rows 4-14
// Difficulty: col 3-4, rows 4-7
// Cuisine: col 6-7, rows 4-22
// Dietary: col 9-10, rows 4-18

(async () => {
  const reqs = [];
  const vals = [];

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 0, 1, 0, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A1`, values: [['🏷️ CATEGORIES & TAGS']] });

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
  vals.push({ range: `${S}!A2`, values: [['Browse your recipe collection by category, cuisine, dietary tag, or difficulty']] });

  // Row 3: spacer
  reqs.push({ repeatCell: {
    range: gridRange(SID, 2, 3, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.secondary) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Summary Stats Header (Row 4) ──────────────────────────────────────────
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
  vals.push({ range: `${S}!A4`, values: [['📊 RECIPE COUNT SUMMARY  (chart source data)']] });

  // Helper: write a summary table
  const writeSummaryTable = (labelCol, countCol, rowStart, header, items, mriFilterCol) => {
    const rIdx = rowStart;
    // Header row
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

    // Data rows
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
      vals.push({ range: `${S}!${colL(countCol)}${r+1}`, values: [
        [`=COUNTIF(${MRI}!$${mriFilterCol}$8:$${mriFilterCol}$2007,"${item.replace(/"/g, '""')}")`]
      ] });
    });
  };

  // Summary tables (all start at row index 4 = row 5 in sheets)
  writeSummaryTable(0, 1, 4, 'MEAL TYPES', MEAL_TYPES, 'E');
  writeSummaryTable(3, 4, 4, 'DIFFICULTY', DIFFICULTIES, 'L');
  writeSummaryTable(6, 7, 4, 'CUISINES', CUISINES, 'F');
  writeSummaryTable(9, 10, 4, 'DIETARY TAGS', DIETARY_TAGS, 'S');

  // Row 25: spacer before browsers
  reqs.push({ repeatCell: {
    range: gridRange(SID, 25, 26, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.secondary) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Browser sections ──────────────────────────────────────────────────────
  // Each browser has: section header, selector row, col header row, N data rows

  const writeBrowser = (startRowIdx, sectionTitle, selectorLabel, dropdownSource, filterMriCol, rowCount, selectorCellCol) => {
    const r = startRowIdx;

    // Section header
    reqs.push({ mergeCells: { range: gridRange(SID, r, r+1, 0, 20), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 0, 20),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.primary),
        textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!A${r+1}`, values: [[sectionTitle]] });

    // Selector row (r+1)
    reqs.push({ repeatCell: {
      range: gridRange(SID, r+1, r+2, 0, 3),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.primary),
        textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!A${r+2}`, values: [[selectorLabel]] });

    reqs.push({ mergeCells: { range: gridRange(SID, r+1, r+2, selectorCellCol, selectorCellCol+3), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, r+1, r+2, selectorCellCol, selectorCellCol+3),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.input),
        textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.text), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        borders: { top: { style: 'SOLID', color: hex(C.primary) }, bottom: { style: 'SOLID', color: hex(C.primary) },
                   left: { style: 'SOLID', color: hex(C.primary) }, right: { style: 'SOLID', color: hex(C.primary) } },
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,borders)',
    } });
    reqs.push({ setDataValidation: {
      range: gridRange(SID, r+1, r+2, selectorCellCol, selectorCellCol+3),
      rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: dropdownSource }] },
        showCustomUi: true, strict: false,
      },
    } });

    // Result count (after dropdown)
    reqs.push({ repeatCell: {
      range: gridRange(SID, r+1, r+2, selectorCellCol+3, selectorCellCol+7),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.butter),
        textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    const selRef = `${colL(selectorCellCol)}${r+2}`;
    vals.push({ range: `${S}!${colL(selectorCellCol+3)}${r+2}`,
      values: [[`=IFERROR(COUNTIF(${MRI}!$${filterMriCol}$8:$${filterMriCol}$2007,${selRef})&" recipes found","—")`]] });

    // Column headers row (r+2)
    const colHdrRow = r + 2;
    BROWSER_COLS.forEach((col, ci) => {
      reqs.push({ repeatCell: {
        range: gridRange(SID, colHdrRow, colHdrRow+1, ci, ci+1),
        cell: { userEnteredFormat: {
          backgroundColor: hex(C.sage),
          textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      } });
    });
    vals.push({ range: `${S}!A${colHdrRow+1}:G${colHdrRow+1}`, values: [BROWSER_COLS.map(c => c.header)] });

    // Data rows (r+3 to r+3+rowCount-1)
    const dataStart = r + 3;
    const filterPredicate = `${MRI}!$${filterMriCol}$8:$${filterMriCol}$2007=${colL(selectorCellCol)}$${r+2}`;

    for (let i = 0; i < rowCount; i++) {
      const dr = dataStart + i;
      const bg = i % 2 === 0 ? C.altRow : C.panel;
      reqs.push({ repeatCell: {
        range: gridRange(SID, dr, dr+1, 0, 7),
        cell: { userEnteredFormat: {
          backgroundColor: hex(bg),
          textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
          verticalAlignment: 'MIDDLE',
        } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      } });
    }

    // Write formulas column by column
    BROWSER_COLS.forEach(({ filterCol }, ci) => {
      const formulas = Array.from({ length: rowCount }, (_, i) => {
        const N = i + 1;
        return [`=IFERROR(IF(${colL(selectorCellCol)}$${r+2}="","",INDEX(FILTER(${MRI}!$${filterCol}$8:$${filterCol}$2007,${filterPredicate}),${N})),"")`];
      });
      vals.push({ range: `${S}!${colL(ci)}${dataStart+1}:${colL(ci)}${dataStart+rowCount}`, values: formulas });
    });

    return dataStart + rowCount; // next available row index
  };

  // Write all four browser sections
  let nextRow = 26; // row index after spacer at 25

  nextRow = writeBrowser(nextRow,
    '🔍 BROWSE BY RECIPE CATEGORY',
    'SELECT CATEGORY →', "='Reference Data'!$B$4:$B$32", 'C', 50, 3);

  // Spacer
  reqs.push({ repeatCell: {
    range: gridRange(SID, nextRow, nextRow+1, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.secondary) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });
  nextRow += 1;

  nextRow = writeBrowser(nextRow,
    '🌍 BROWSE BY CUISINE',
    'SELECT CUISINE →', "='Reference Data'!$C$4:$C$22", 'F', 40, 3);

  reqs.push({ repeatCell: {
    range: gridRange(SID, nextRow, nextRow+1, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.secondary) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });
  nextRow += 1;

  nextRow = writeBrowser(nextRow,
    '🥗 BROWSE BY DIETARY TAG',
    'SELECT DIETARY TAG →', "='Reference Data'!$D$4:$D$18", 'S', 40, 3);

  reqs.push({ repeatCell: {
    range: gridRange(SID, nextRow, nextRow+1, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.secondary) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });
  nextRow += 1;

  nextRow = writeBrowser(nextRow,
    '⭐ BROWSE BY DIFFICULTY',
    'SELECT DIFFICULTY →', "='Reference Data'!$G$4:$G$7", 'L', 20, 3);

  // ── Column Widths ─────────────────────────────────────────────────────────
  const widths = [85, 220, 90, 95, 80, 60, 75, 40, 40, 120, 90, 40, 40, 40, 40, 40, 40, 40, 40, 40];
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
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 4 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: nextRow },
    properties: { pixelSize: 22 }, fields: 'pixelSize',
  } });

  // ── Freeze ────────────────────────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } },
    fields: 'gridProperties.frozenRowCount',
  } });

  // ── Tab Color ─────────────────────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, tabColorStyle: { rgbColor: hex(C.secondary) } },
    fields: 'tabColorStyle',
  } });

  await batchUpdate(id, reqs, 'cat-fmt');
  await valuesBatchUpdate(id, vals, 'cat-data');
  console.log(`✓ Categories & Tags tab complete (${nextRow} rows used)`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
