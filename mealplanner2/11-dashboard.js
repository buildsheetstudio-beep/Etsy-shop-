'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Meal Planning Dashboard'];
const S   = "'Meal Planning Dashboard'";
const MP  = "'Weekly Meal Planner'";
const RB  = "'Recipe Book'";
const GL  = "'Automated Grocery List'";
const PI  = "'Pantry Inventory'";
const MC  = "'Meal Cost Tracker'";
const HS  = "'Household Setup'";

const NC = 16;

(async () => {
  const fmt = [];
  const vals = [];

  // ─── Title banner ─────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.butter), textFormat: { bold: true, fontSize: 24, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  vals.push({ range: `${S}!A1`, values: [['WEEKLY MEAL PLANNING DASHBOARD']] });

  // Subtitle
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.peach), textFormat: { italic: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  vals.push({ range: `${S}!A2`, values: [['Meals  •  Recipes  •  Pantry  •  Grocery List  •  Costs']] });

  // ─── Controls (row 3-4, 0-indexed 2-3) ────────────────────────────────────
  const controls = [
    ['Week Start Date',          `=DATE(2026,8,18)`],
    ['Household Member / All',   'All'],
    ['Meal Type / All',          'All'],
    ['Recipe Category / All',    'All'],
    ['Grocery Category / All',   'All'],
  ];

  // 5 controls in one row across full width
  const CTRL_COLS = Math.floor(NC / controls.length); // 3 each approx
  controls.forEach(([label, value], i) => {
    const col = i * 3;
    fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, col, col+3), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, 3, 4, col, col+3), mergeType: 'MERGE_ALL' } });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, 2, 3, col, col+3),
        cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, 3, 4, col, col+3),
        cell: { userEnteredFormat: { backgroundColor: hex(C.input), textFormat: { fontSize: 10, fontFamily: 'Arial' }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}3`, values: [[label]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}4`, values: [[value]] });
  });

  const WK = `$A$4`; // Week start date control cell

  // Date format for week start control
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 0, 3), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Freeze rows 1:5
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  // Separator row 4 (0-indexed)
  fmt.push({ mergeCells: { range: gridRange(SID, 4, 5, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.butter) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // ─── Primary KPI Cards (rows 5-7, 0-indexed) ─────────────────────────────
  const primaryKPIs = [
    ['Meals Planned',        `=IFERROR(COUNTIFS(${MP}!$B$12:$B$511,${WK}),0)`],
    ['Unique Recipes',       `=IFERROR(SUMPRODUCT((${MP}!$B$12:$B$511=${WK})*(${MP}!$F$12:$F$511<>"")/COUNTIFS(${MP}!$F$12:$F$511,IF(${MP}!$F$12:$F$511<>"",${MP}!$F$12:$F$511,"~"),${MP}!$B$12:$B$511,${WK})),0)`],
    ['Planned Servings',     `=IFERROR(SUMIFS(${MP}!$L$12:$L$511,${MP}!$B$12:$B$511,${WK}),0)`],
    ['Pantry Coverage %',    `=IFERROR(${GL}!D10,0)`],
    ['Grocery Items Needed', `=IFERROR(${GL}!C10,0)`],
    ['Est. Grocery Cost',    `=IFERROR(${GL}!E10,0)`],
    ['Weekly Budget',        `=IFERROR(${HS}!F8,200)`],
    ['Budget Remaining',     `=IFERROR(${GL}!G10,0)`],
  ];

  primaryKPIs.forEach(([label, formula], i) => {
    const col = (i % 4) * 4;
    const row = 5 + Math.floor(i / 4) * 3;
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, col, col+4), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, row+1, row+3, col, col+4), mergeType: 'MERGE_ALL' } });
    const cardBg = i < 4 ? C.butter : C.mint;
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, col, col+4),
        cell: { userEnteredFormat: { backgroundColor: hex(cardBg), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row+1, row+3, col, col+4),
        cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+1}`, values: [[label]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+2}`, values: [[formula]] });
  });

  // Format card values
  // Row 6 and 8 (1-indexed) have values
  // Pantry % (col 12-16, row 7)
  fmt.push({ repeatCell: { range: gridRange(SID, 6, 8, 12, 16), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });
  // Grocery cost (col 4-8, row 10), Budget (col 8-12), Budget Remaining (col 12-16)
  fmt.push({ repeatCell: { range: gridRange(SID, 9, 11, 4, 16), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Separator row 11 (0-indexed)
  fmt.push({ mergeCells: { range: gridRange(SID, 11, 12, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 11, 12, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.butter) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // ─── Secondary KPI Cards (rows 12-14, 0-indexed) ─────────────────────────
  const secondaryKPIs = [
    ['Favorite Recipes Used',   `=IFERROR(SUMPRODUCT((${MP}!$B$12:$B$511=${WK})*(IFERROR(VLOOKUP(${MP}!$F$12:$F$511,${RB}!$A$6:$M$1005,13,FALSE),FALSE))),0)`],
    ['Dietary Reviews Needed',  `=IFERROR(COUNTIFS(${MP}!$B$12:$B$511,${WK},${MP}!$P$12:$P$511,"Review*"),0)`],
    ['Pantry Items Low',        `=IFERROR(${PI}!B4,0)`],
    ['Expiring Soon',           `=IFERROR(${PI}!D4,0)`],
    ['Weekly Meal Cost',        `=IFERROR(${MC}!A8,0)`],
    ['Cost per Serving',        `=IFERROR(${MC}!D8,0)`],
    ['Grocery Items Purchased', `=IFERROR(COUNTIF(${GL}!$M$12:$M$1505,TRUE),0)`],
    ['Planning Completion %',   `=IFERROR(COUNTIFS(${MP}!$B$12:$B$511,${WK},${MP}!$F$12:$F$511,"REC-*")/COUNTIFS(${MP}!$B$12:$B$511,${WK}),0)`],
  ];

  secondaryKPIs.forEach(([label, formula], i) => {
    const col = (i % 4) * 4;
    const row = 12 + Math.floor(i / 4) * 2;
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, col, col+4), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, row+1, row+2, col, col+4), mergeType: 'MERGE_ALL' } });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, col, col+4),
        cell: { userEnteredFormat: { backgroundColor: hex(C.blush), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row+1, row+2, col, col+4),
        cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { bold: true, fontSize: 16, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+1}`, values: [[label]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+2}`, values: [[formula]] });
  });

  // Currency for meal cost cards
  fmt.push({ repeatCell: { range: gridRange(SID, 13, 14, 0, 8), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 15, 16, 0, 8), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 15, 16, 12, 16), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Separator row 16
  fmt.push({ mergeCells: { range: gridRange(SID, 16, 17, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 16, 17, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.butter) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // ─── Weekly Meal Snapshot (rows 17-24, 0-indexed) ────────────────────────
  const SNAP_ROW = 17;
  fmt.push({ mergeCells: { range: gridRange(SID, SNAP_ROW, SNAP_ROW+1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, SNAP_ROW, SNAP_ROW+1, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.text), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'LEFT' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  vals.push({ range: `${S}!A${SNAP_ROW+1}`, values: [['  📅 WEEKLY MEAL SNAPSHOT']] });

  // Day columns (Mon-Sun) in grid
  const DAYS = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
  const SNAP_HDR = SNAP_ROW + 1;
  const SNAP_D0  = SNAP_HDR + 1;

  // Header row: Label + 7 days
  const dayColWidth = Math.floor((NC - 2) / 7); // ~2 cols each
  fmt.push({ mergeCells: { range: gridRange(SID, SNAP_HDR, SNAP_HDR+1, 0, 2), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, SNAP_HDR, SNAP_HDR+1, 0, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.text), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A${SNAP_HDR+1}`, values: [['Meal Type']] });

  DAYS.forEach((day, di) => {
    const col = 2 + di * 2;
    fmt.push({ mergeCells: { range: gridRange(SID, SNAP_HDR, SNAP_HDR+1, col, col+2), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, SNAP_HDR, SNAP_HDR+1, col, col+2), cell: { userEnteredFormat: { backgroundColor: hex(C.text), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${SNAP_HDR+1}`, values: [[day]] });
  });

  // Snapshot rows: Breakfast, Lunch, Dinner, Snack/Dessert, Servings
  const snapshotMeals = [
    ['Breakfast', C.butter, 'Breakfast'],
    ['Lunch',     C.mint,   'Lunch'],
    ['Dinner',    C.powder, 'Dinner'],
    ['Snack / Dessert', C.peach, 'Afternoon Snack'],
  ];
  snapshotMeals.forEach(([label, color, mealType], ri) => {
    const row = SNAP_D0 + ri;
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, 0, 2), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, 0, 2), cell: { userEnteredFormat: { backgroundColor: hex(color), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!A${row+1}`, values: [[label]] });

    DAYS.forEach((day, di) => {
      const col = 2 + di * 2;
      fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, col, col+2), mergeType: 'MERGE_ALL' } });
      fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, col, col+2), cell: { userEnteredFormat: { backgroundColor: hex(color), textFormat: { fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' } });
      const dateExpr = `(${WK}+${di})`;
      const formula = `=IFERROR(INDEX(${MP}!$G$12:$G$511,MATCH(1,(IFERROR(${MP}!$C$12:$C$511,0)=${dateExpr})*(IFERROR(${MP}!$E$12:$E$511,"")="${mealType}"),0)),"—")`;
      vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+1}`, values: [[formula]] });
    });
  });

  // ─── Grocery Snapshot (rows 25-35) ────────────────────────────────────────
  const GR_SNAP_ROW = SNAP_D0 + snapshotMeals.length + 1;
  fmt.push({ mergeCells: { range: gridRange(SID, GR_SNAP_ROW, GR_SNAP_ROW+1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, GR_SNAP_ROW, GR_SNAP_ROW+1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.text), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'LEFT' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A${GR_SNAP_ROW+1}`, values: [['  🛒 GROCERY SNAPSHOT (Top Items Needed)']] });

  // Grocery snapshot header
  const GR_HDR_ROW = GR_SNAP_ROW + 1;
  const GR_D0      = GR_HDR_ROW + 1;
  const grCols     = ['Ingredient','Needed Qty','Unit','Pantry Qty','Qty to Buy','Category','Est. Cost','Purchased?'];
  fmt.push({ repeatCell: { range: gridRange(SID, GR_HDR_ROW, GR_HDR_ROW+1, 0, grCols.length), cell: { userEnteredFormat: { backgroundColor: hex(C.blush), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A${GR_HDR_ROW+1}`, values: [grCols] });

  // Pull top 10 grocery items from Automated Grocery List
  for (let i = 0; i < 10; i++) {
    const r = GR_D0 + i;
    const srcRow = 12 + i;
    const bgColor = i % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 0, grCols.length), cell: { userEnteredFormat: { backgroundColor: hex(bgColor), textFormat: { fontSize: 9, fontFamily: 'Arial' } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
    vals.push({ range: `${S}!A${r+1}`, values: [[
      `=IFERROR(${GL}!B${srcRow},"")`,
      `=IFERROR(${GL}!D${srcRow},"")`,
      `=IFERROR(${GL}!E${srcRow},"")`,
      `=IFERROR(${GL}!F${srcRow},"")`,
      `=IFERROR(${GL}!H${srcRow},"")`,
      `=IFERROR(${GL}!C${srcRow},"")`,
      `=IFERROR(${GL}!J${srcRow},"")`,
      `=IFERROR(${GL}!M${srcRow},FALSE)`,
    ]] });
    // Checkbox for purchased col
    fmt.push({ setDataValidation: { range: gridRange(SID, r, r+1, 7, 8), rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true } } });
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 6, 7), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  }

  // ─── Pantry Snapshot ──────────────────────────────────────────────────────
  const PAN_SNAP_ROW = GR_D0 + 11;
  fmt.push({ mergeCells: { range: gridRange(SID, PAN_SNAP_ROW, PAN_SNAP_ROW+1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, PAN_SNAP_ROW, PAN_SNAP_ROW+1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.text), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'LEFT' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A${PAN_SNAP_ROW+1}`, values: [['  🫙 PANTRY SNAPSHOT']] });

  const panCards = [
    ['Low Stock',        `=IFERROR(COUNTIF(${PI}!$G$6:$G$1505,"Low Stock"),0)`],
    ['Out of Stock',     `=IFERROR(COUNTIF(${PI}!$G$6:$G$1505,"Out of Stock"),0)`],
    ['Expiring Soon',    `=IFERROR(COUNTIF(${PI}!$G$6:$G$1505,"Expiring Soon"),0)`],
    ['Pantry Value',     `=IFERROR(SUM(${PI}!$L$6:$L$1505),0)`],
  ];
  panCards.forEach(([label, formula], i) => {
    const col = i * 4;
    const row = PAN_SNAP_ROW + 1;
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, col, col+4), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, row+1, row+2, col, col+4), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, col, col+4), cell: { userEnteredFormat: { backgroundColor: hex(C.mint), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ repeatCell: { range: gridRange(SID, row+1, row+2, col, col+4), cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { bold: true, fontSize: 16, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+1}`, values: [[label]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+2}`, values: [[formula]] });
  });

  // Currency for pantry value
  fmt.push({ repeatCell: { range: gridRange(SID, PAN_SNAP_ROW+2, PAN_SNAP_ROW+3, 12, 16), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // ─── Dietary Compatibility Snapshot ───────────────────────────────────────
  const DIET_SNAP_ROW = PAN_SNAP_ROW + 3;
  fmt.push({ mergeCells: { range: gridRange(SID, DIET_SNAP_ROW, DIET_SNAP_ROW+1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, DIET_SNAP_ROW, DIET_SNAP_ROW+1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.text), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'LEFT' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A${DIET_SNAP_ROW+1}`, values: [['  🏷️ DIETARY COMPATIBILITY SNAPSHOT']] });

  const dietCards = [
    ['Recipes Requiring Review',       `=IFERROR(COUNTIFS(${MP}!$B$12:$B$511,${WK},${MP}!$P$12:$P$511,"Review*"),0)`],
    ['Allergen Review Count',          `=IFERROR(COUNTIFS(${MP}!$B$12:$B$511,${WK},${MP}!$P$12:$P$511,"Review Allergen"),0)`],
    ['Dietary Restriction Reviews',    `=IFERROR(COUNTIFS(${MP}!$B$12:$B$511,${WK},${MP}!$P$12:$P$511,"Review Restriction"),0)`],
    ['Recipes Missing Tags',           `=IFERROR(COUNTIFS(${MP}!$B$12:$B$511,${WK},${MP}!$P$12:$P$511,"Missing Tags"),0)`],
  ];

  const dietNote = '⚠ Compatibility counts are for organizational reference only. Always verify ingredient labels, cross-contamination risks, and individual dietary needs before serving.';
  dietCards.forEach(([label, formula], i) => {
    const col = i * 4;
    const row = DIET_SNAP_ROW + 1;
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, col, col+4), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, row+1, row+2, col, col+4), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, col, col+4), cell: { userEnteredFormat: { backgroundColor: hex(C.lavender), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ repeatCell: { range: gridRange(SID, row+1, row+2, col, col+4), cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { bold: true, fontSize: 16, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+1}`, values: [[label]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+2}`, values: [[formula]] });
  });

  // ─── Dashboard Disclaimer ──────────────────────────────────────────────────
  const DISC_ROW = DIET_SNAP_ROW + 4;
  fmt.push({ mergeCells: { range: gridRange(SID, DISC_ROW, DISC_ROW+3, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, DISC_ROW, DISC_ROW+3, 0, NC),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.attention),
          textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white), italic: true },
          horizontalAlignment: 'LEFT', verticalAlignment: 'TOP', wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });
  vals.push({ range: `${S}!A${DISC_ROW+1}`, values: [['⚠ DASHBOARD DISCLAIMER: For organizational meal planning only. Dietary and allergen tags are not a substitute for ingredient-label review, medical advice, allergy management, or professional dietary guidance. Grocery quantities and costs are estimates based on the information entered by the user.']] });

  // ─── Column widths ────────────────────────────────────────────────────────
  const colWidths = [100,80,80,80,80,80,80,80,80,80,80,80,80,80,80,80];
  colWidths.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 60 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: DISC_ROW, endIndex: DISC_ROW+3 }, properties: { pixelSize: 80 }, fields: 'pixelSize' } });

  await batchUpdate(id, fmt, 'dashboard-fmt');
  await valuesBatchUpdate(id, vals, 'dashboard-vals');
  console.log('✓ Meal Planning Dashboard complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
