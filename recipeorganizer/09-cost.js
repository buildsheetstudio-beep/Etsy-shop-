'use strict';
const { hex, batchUpdate, valuesBatchUpdate, gridRange, colL, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Recipe Cost Calculator'];
const S = "'Recipe Cost Calculator'";

const MRI = "'Master Recipe Index'";
const RECIPE_ROWS = 60;

// Budget goal cell: H8 (user enters max cost per batch)
// Manual override column: G (col 6, GS col G)
// Effective batch cost: H (col 7) = IF(G11="", E11, G11)
// Effective cost/serving: I (col 8) = H11/D11
// vs Budget: J (col 9) = H11 - $H$8

(async () => {
  const reqs = [];
  const vals = [];

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 0, 1, 0, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.success),
      textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A1`, values: [['💰 RECIPE COST CALCULATOR']] });

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
  vals.push({ range: `${S}!A2`, values: [['Costs pulled from Recipe Ingredients tab — override in column G if needed — set budget goal in H8']] });

  // Row 3: spacer
  reqs.push({ repeatCell: {
    range: gridRange(SID, 2, 3, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.success) } },
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
  vals.push({ range: `${S}!A4`, values: [['📊 COST OVERVIEW']] });

  // ── Rows 5-6: KPI Cards ───────────────────────────────────────────────────
  const kpiCards = [
    { label: 'AVG COST / BATCH',    formula: `=IFERROR("$"&ROUND(AVERAGEIF(${MRI}!$B$8:$B$2007,"<>",${MRI}!$Y$8:$Y$2007),2),"—")`,  cols: [0, 5]  },
    { label: 'AVG COST / SERVING',  formula: `=IFERROR("$"&ROUND(AVERAGEIF(${MRI}!$B$8:$B$2007,"<>",${MRI}!$Z$8:$Z$2007),2),"—")`,  cols: [5, 10] },
    { label: 'MOST EXPENSIVE',      formula: `=IFERROR("$"&ROUND(MAX(${MRI}!$Y$8:$Y$2007),2),"—")`,                                    cols: [10, 15]},
    { label: 'LEAST EXPENSIVE',     formula: `=IFERROR("$"&ROUND(MINIFS(${MRI}!$Y$8:$Y$2007,${MRI}!$Y$8:$Y$2007,">0"),2),"—")`,      cols: [15, 20]},
  ];
  kpiCards.forEach(({ label, formula, cols: [c1, c2] }) => {
    reqs.push({ mergeCells: { range: gridRange(SID, 4, 5, c1, c2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, 4, 5, c1, c2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.sage),
        textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!${colL(c1)}5`, values: [[label]] });

    reqs.push({ mergeCells: { range: gridRange(SID, 5, 6, c1, c2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, 5, 6, c1, c2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.panel),
        textFormat: { bold: true, fontSize: 20, foregroundColor: hex(C.success), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!${colL(c1)}6`, values: [[formula]] });
  });

  // ── Row 7: Spacer ─────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID, 6, 7, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.success) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Row 8: Budget Goal Row ────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 7, 8, 0, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 7, 8, 0, 5),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.butter),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A8`, values: [['🎯 BUDGET GOAL — max cost per batch ($)  →']] });

  // H8: budget goal input cell
  reqs.push({ repeatCell: {
    range: gridRange(SID, 7, 8, 5, 7),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.input),
      textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      borders: {
        top:    { style: 'SOLID_MEDIUM', color: hex(C.success) },
        bottom: { style: 'SOLID_MEDIUM', color: hex(C.success) },
        left:   { style: 'SOLID_MEDIUM', color: hex(C.success) },
        right:  { style: 'SOLID_MEDIUM', color: hex(C.success) },
      },
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,borders)',
  } });
  vals.push({ range: `${S}!F8`, values: [[15]] }); // default budget: $15

  // Budget note
  reqs.push({ mergeCells: { range: gridRange(SID, 7, 8, 7, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 7, 8, 7, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.butter),
      textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!H8`, values: [['← Enter your max spend per batch; "vs Budget" column turns green (under) or red (over)']] });

  // ── Row 9: Spacer ─────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID, 8, 9, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.success) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Row 10: Cost Table Header ─────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 9, 10, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 9, 10, 0, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A10`, values: [['📋 RECIPE COST TABLE']] });

  // ── Row 11: Column Headers ────────────────────────────────────────────────
  const hdrs = ['Recipe ID', 'Recipe Name', 'Category', 'Servings', 'Calc Cost/Batch ($)', 'Calc Cost/Serving ($)', 'Manual Override ($)', 'Effective Cost/Batch', 'Effective/Serving', 'vs Budget', 'Notes'];
  hdrs.forEach((h, i) => {
    reqs.push({ repeatCell: {
      range: gridRange(SID, 10, 11, i, i+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.sage),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        wrapStrategy: 'WRAP',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    } });
  });
  vals.push({ range: `${S}!A11:K11`, values: [hdrs] });

  // ── Data Rows 12-71 (idx 11-70, 60 recipes) ───────────────────────────────
  for (let i = 0; i < RECIPE_ROWS; i++) {
    const r = 11 + i;
    const bg = i % 2 === 0 ? C.altRow : C.panel;
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 0, 9),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } });
    // Manual override (col G = col 6) gets input styling
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 6, 7),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.input),
        textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } });
    // Notes col (col 10)
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 10, 11),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.input),
        textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } });
  }

  // Write pull formulas column-by-column (cols A-F from MRI)
  const pullCols = [
    { mriCol: 'A', targetCol: 0 }, // Recipe ID
    { mriCol: 'B', targetCol: 1 }, // Recipe Name
    { mriCol: 'C', targetCol: 2 }, // Category
    { mriCol: 'H', targetCol: 3 }, // Servings
    { mriCol: 'Y', targetCol: 4 }, // Est. Cost/Batch
    { mriCol: 'Z', targetCol: 5 }, // Est. Cost/Serving
  ];
  pullCols.forEach(({ mriCol, targetCol }) => {
    const formulas = Array.from({length: RECIPE_ROWS}, (_, i) =>
      [`=IFERROR(INDEX(${MRI}!$${mriCol}$8:$${mriCol}$2007,${i+1}),"")`]);
    vals.push({ range: `${S}!${colL(targetCol)}12:${colL(targetCol)}${11+RECIPE_ROWS}`, values: formulas });
  });

  // Effective Cost/Batch (col H = col 7): =IF(G12="", E12, G12)
  const effBatchFormulas = Array.from({length: RECIPE_ROWS}, (_, i) => {
    const gsRow = 12 + i;
    return [`=IF(G${gsRow}="",E${gsRow},G${gsRow})`];
  });
  vals.push({ range: `${S}!H12:H${11+RECIPE_ROWS}`, values: effBatchFormulas });

  // Effective Cost/Serving (col I = col 8): =IFERROR(H12/D12, "")
  const effServFormulas = Array.from({length: RECIPE_ROWS}, (_, i) => {
    const gsRow = 12 + i;
    return [`=IFERROR(H${gsRow}/D${gsRow},"")`];
  });
  vals.push({ range: `${S}!I12:I${11+RECIPE_ROWS}`, values: effServFormulas });

  // vs Budget (col J = col 9): =IF(H12="","",H12-$F$8)
  const vsBudgetFormulas = Array.from({length: RECIPE_ROWS}, (_, i) => {
    const gsRow = 12 + i;
    return [`=IF(H${gsRow}="","",H${gsRow}-$F$8)`];
  });
  vals.push({ range: `${S}!J12:J${11+RECIPE_ROWS}`, values: vsBudgetFormulas });

  // ── Conditional Formatting: vs Budget column ──────────────────────────────
  // Green if <= 0 (under budget), red if > 0 (over budget)
  reqs.push({ addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(SID, 11, 11 + RECIPE_ROWS, 9, 10)],
      booleanRule: {
        condition: { type: 'NUMBER_LESS_THAN_EQ', values: [{ userEnteredValue: '0' }] },
        format: { backgroundColor: hex(C.success) },
      },
    },
    index: 0,
  } });
  reqs.push({ addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(SID, 11, 11 + RECIPE_ROWS, 9, 10)],
      booleanRule: {
        condition: { type: 'NUMBER_GREATER', values: [{ userEnteredValue: '0' }] },
        format: { backgroundColor: hex(C.attention) },
      },
    },
    index: 1,
  } });

  // Format currency cols (E, F, G, H, I, J) with $ number format
  [4, 5, 6, 7, 8, 9].forEach(ci => {
    reqs.push({ repeatCell: {
      range: gridRange(SID, 11, 11 + RECIPE_ROWS, ci, ci+1),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } },
      fields: 'userEnteredFormat(numberFormat)',
    } });
  });

  // ── Column Widths ─────────────────────────────────────────────────────────
  const widths = [85, 210, 100, 65, 100, 100, 90, 100, 90, 80, 130, 40, 40, 40, 40, 40, 40, 40, 40, 40];
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
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 7 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 },
    properties: { pixelSize: 38 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 8 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 10, endIndex: 11 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 11, endIndex: 72 },
    properties: { pixelSize: 22 }, fields: 'pixelSize',
  } });

  // ── Freeze ────────────────────────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, gridProperties: { frozenRowCount: 11 } },
    fields: 'gridProperties.frozenRowCount',
  } });

  // ── Tab Color ─────────────────────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, tabColorStyle: { rgbColor: hex(C.success) } },
    fields: 'tabColorStyle',
  } });

  await batchUpdate(id, reqs, 'cost-fmt');
  await valuesBatchUpdate(id, vals, 'cost-data');
  console.log('✓ Recipe Cost Calculator tab complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
