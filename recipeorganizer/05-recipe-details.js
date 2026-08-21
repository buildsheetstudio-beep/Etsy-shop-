'use strict';
const { hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Recipe Details'];
const S = "'Recipe Details'";

// Column layout (0-indexed):
// A(0)=#  B(1)=Qty  C(2)=Unit  D(3)=Ingredient  E(4)=Category  F(5)=Opt?
// G(6)=spacer
// H(7)=Step#  I(8)=Type  J(9)=Instruction  K(10)=Min  L(11)=Equip  M(12)=MakeAhead
// T(19)=scale factor helper
//
// Row layout:
// 0=title  1=subtitle  2=spacer  3=selector  4=metadata1  5=metadata2
// 6=dietary+serving note  7=description  8=spacer
// 9=section headers  10=col sub-headers
// 11-51=ingredient rows (41)  11-68=instruction rows (58)

const ING_ROWS  = 41;
const INST_ROWS = 58;
const DATA_START = 11; // 0-indexed row index for first data row

// VLOOKUP helper
const vl = (colIdx) =>
  `=IFERROR(VLOOKUP($B$4,'Master Recipe Index'!$A$8:$AH$2007,${colIdx},0),"")`;

(async () => {
  const reqs = [];
  const vals = [];

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 0, 1, 0, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lavender),
      textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A1`, values: [['🍴 RECIPE CARD VIEWER']] });

  // ── Row 2: Subtitle ───────────────────────────────────────────────────────
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
  vals.push({ range: `${S}!A2`, values: [['Select a recipe below — ingredients auto-scale when Scale Mode is ON']] });

  // ── Row 3: Spacer ─────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID, 2, 3, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.lavender) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Row 4: Recipe Selector Controls ───────────────────────────────────────
  const labelFmt = {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  };
  const inputFmt = {
    backgroundColor: hex(C.input),
    textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    borders: {
      top:    { style: 'SOLID', color: hex(C.primary) },
      bottom: { style: 'SOLID', color: hex(C.primary) },
      left:   { style: 'SOLID', color: hex(C.primary) },
      right:  { style: 'SOLID', color: hex(C.primary) },
    },
  };
  const valFmt = {
    backgroundColor: hex(C.formula),
    textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryDeep), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  };

  // A4: "RECIPE ID" label
  reqs.push({ repeatCell: { range: gridRange(SID, 3, 4, 0, 1), cell: { userEnteredFormat: labelFmt },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A4`, values: [['RECIPE ID']] });

  // B4:C4 merged — recipe ID input cell (default REC-0001)
  reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, 1, 3), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 3, 4, 1, 3),
    cell: { userEnteredFormat: inputFmt },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,borders)' } });
  vals.push({ range: `${S}!B4`, values: [['REC-0001']] });

  // Data validation for Recipe ID
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 3, 4, 1, 3),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: "='Master Recipe Index'!$A$8:$A$2007" }] },
      showCustomUi: true, strict: false,
    },
  } });

  // D4: "RECIPE" label
  reqs.push({ repeatCell: { range: gridRange(SID, 3, 4, 3, 4), cell: { userEnteredFormat: labelFmt },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!D4`, values: [['RECIPE NAME']] });

  // E4:I4 merged — recipe name formula
  reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, 4, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 3, 4, 4, 9),
    cell: { userEnteredFormat: valFmt },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!E4`, values: [[vl(2)]] });

  // J4: "CUSTOM SERVINGS" label
  reqs.push({ repeatCell: { range: gridRange(SID, 3, 4, 9, 10), cell: { userEnteredFormat: labelFmt },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!J4`, values: [['CUSTOM SERVINGS']] });

  // K4: custom servings input
  reqs.push({ repeatCell: { range: gridRange(SID, 3, 4, 10, 11),
    cell: { userEnteredFormat: { ...inputFmt, textFormat: { ...inputFmt.textFormat, fontSize: 14 } } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,borders)' } });

  // L4: "/ ORIGINAL" label
  reqs.push({ repeatCell: { range: gridRange(SID, 3, 4, 11, 12), cell: { userEnteredFormat: labelFmt },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!L4`, values: [['/ ORIGINAL']] });

  // M4: original servings formula
  reqs.push({ repeatCell: { range: gridRange(SID, 3, 4, 12, 13),
    cell: { userEnteredFormat: valFmt },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!M4`, values: [[vl(8)]] }); // H = Servings

  // N4: "SCALE MODE" label
  reqs.push({ repeatCell: { range: gridRange(SID, 3, 4, 13, 14), cell: { userEnteredFormat: labelFmt },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!N4`, values: [['SCALE MODE']] });

  // O4: scale checkbox
  reqs.push({ repeatCell: { range: gridRange(SID, 3, 4, 14, 15),
    cell: { userEnteredFormat: { backgroundColor: hex(C.input), horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 3, 4, 14, 15),
    rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true, strict: true },
  } });
  vals.push({ range: `${S}!O4`, values: [[false]] });

  // T4: scale factor helper (hidden helper formula)
  vals.push({ range: `${S}!T4`, values: [['=IFERROR(IF(AND($O$4=TRUE,$M$4>0,$K$4>0),$K$4/$M$4,1),1)']] });

  // ── Rows 5-6: Metadata ────────────────────────────────────────────────────
  const mLabel = (r, c, txt) => {
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, c, c+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.secondary),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}${r+1}`, values: [[txt]] });
  };
  const mVal = (r, c1, c2, formula) => {
    reqs.push({ mergeCells: { range: gridRange(SID, r, r+1, c1, c2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, c1, c2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.altRow),
        textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!${String.fromCharCode(65+c1)}${r+1}`, values: [[formula]] });
  };

  // Row 5 (idx 4): Meal Type | Category | Cuisine | Difficulty | Method
  mLabel(4, 0, 'MEAL TYPE');  mVal(4, 1, 3, vl(3));
  mLabel(4, 3, 'CATEGORY');   mVal(4, 4, 6, vl(4));
  mLabel(4, 6, 'CUISINE');    mVal(4, 7, 9, vl(5));
  mLabel(4, 9, 'DIFFICULTY'); mVal(4, 10, 12, vl(13));
  mLabel(4, 12, 'METHOD');    mVal(4, 13, 15, vl(12));

  // Row 6 (idx 5): Prep | Cook | Total Time | Status | Rating
  mLabel(5, 0, 'PREP (min)');    mVal(5, 1, 3, vl(9));
  mLabel(5, 3, 'COOK (min)');    mVal(5, 4, 6, vl(10));
  mLabel(5, 6, 'TOTAL (min)');   mVal(5, 7, 9, vl(11));
  mLabel(5, 9, 'STATUS');        mVal(5, 10, 12, vl(14));
  mLabel(5, 12, 'RATING ★');     mVal(5, 13, 15, vl(20));

  // ── Row 7: Dietary, Allergens, Serving Note ───────────────────────────────
  mLabel(6, 0, 'DIETARY');
  mVal(6, 1, 5, vl(6));
  mLabel(6, 5, 'ALLERGENS');
  mVal(6, 6, 10, vl(7));

  // Serving note K7:T7
  reqs.push({ mergeCells: { range: gridRange(SID, 6, 7, 10, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 6, 7, 10, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.butter),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!K7`, values: [['=IF($O$4=TRUE,"📐 Scaled: "&$M$4&" → "&$K$4&" servings (×"&ROUND(IFERROR($K$4/$M$4,1),2)&")","📋 Original — "&$M$4&" servings")']] });

  // ── Row 8: Description ────────────────────────────────────────────────────
  mLabel(7, 0, 'DESCRIPTION');
  reqs.push({ mergeCells: { range: gridRange(SID, 7, 8, 1, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 7, 8, 1, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      wrapStrategy: 'WRAP',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
  } });
  vals.push({ range: `${S}!B8`, values: [[vl(31)]] }); // Description = col AE = index 31

  // ── Row 9: Spacer/Divider ─────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID, 8, 9, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.lavender) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Row 10: Section Headers ───────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 9, 10, 0, 6), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 9, 10, 0, 6),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.sage),
      textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A10`, values: [['🥗 INGREDIENTS']] });

  reqs.push({ repeatCell: {
    range: gridRange(SID, 9, 10, 6, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  reqs.push({ mergeCells: { range: gridRange(SID, 9, 10, 7, 20), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 9, 10, 7, 20),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.sageDeep),
      textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!H10`, values: [['📝 STEP-BY-STEP INSTRUCTIONS']] });

  // ── Row 11: Column Sub-Headers ─────────────────────────────────────────────
  const ingHdrs  = ['#', 'Qty', 'Unit', 'Ingredient', 'Category', 'Opt?'];
  const instHdrs = ['Step', 'Type', 'Instruction', 'Min', 'Equipment', 'Make Ahead?'];

  ingHdrs.forEach((_, i) => {
    reqs.push({ repeatCell: {
      range: gridRange(SID, 10, 11, i, i+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.sage),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
  });
  vals.push({ range: `${S}!A11:F11`, values: [ingHdrs] });

  reqs.push({ repeatCell: {
    range: gridRange(SID, 10, 11, 6, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  instHdrs.forEach((_, i) => {
    reqs.push({ repeatCell: {
      range: gridRange(SID, 10, 11, 7+i, 8+i),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.sageDeep),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
  });
  vals.push({ range: `${S}!H11:M11`, values: [instHdrs] });

  reqs.push({ repeatCell: {
    range: gridRange(SID, 10, 11, 13, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Data Row Backgrounds ──────────────────────────────────────────────────
  const totalDataRows = Math.max(ING_ROWS, INST_ROWS);
  for (let r = DATA_START; r < DATA_START + totalDataRows; r++) {
    const bg = r % 2 === 0 ? C.altRow : C.panel;
    // Ingredient cols A-F
    if (r < DATA_START + ING_ROWS) {
      reqs.push({ repeatCell: {
        range: gridRange(SID, r, r+1, 0, 6),
        cell: { userEnteredFormat: {
          backgroundColor: hex(bg),
          textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
          verticalAlignment: 'MIDDLE',
        } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      } });
    }
    // Instruction cols H-M
    if (r < DATA_START + INST_ROWS) {
      reqs.push({ repeatCell: {
        range: gridRange(SID, r, r+1, 7, 13),
        cell: { userEnteredFormat: {
          backgroundColor: hex(bg),
          textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
          verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
        } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,wrapStrategy)',
      } });
    }
    // Spacer col G
    reqs.push({ repeatCell: {
      range: gridRange(SID, r, r+1, 6, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
      fields: 'userEnteredFormat(backgroundColor)',
    } });
  }

  // ── Ingredient Formulas (columns A-F, rows 12 to 12+ING_ROWS-1) ──────────
  // FILTER returns ingredients for selected recipe in order; INDEX picks Nth
  const ingFilter = (col) =>
    `'Recipe Ingredients'!$${col}$8:$${col}$12007,'Recipe Ingredients'!$B$8:$B$12007=$B$4`;

  vals.push({ range: `${S}!A12:A${11+ING_ROWS}`, values:
    Array.from({length: ING_ROWS}, (_, i) => [`=IF(B${12+i}="","",${i+1})`]) });
  vals.push({ range: `${S}!B12:B${11+ING_ROWS}`, values:
    Array.from({length: ING_ROWS}, (_, i) =>
      [`=IFERROR(INDEX(FILTER(${ingFilter('F')}),${i+1})*$T$4,"")`]) });
  vals.push({ range: `${S}!C12:C${11+ING_ROWS}`, values:
    Array.from({length: ING_ROWS}, (_, i) =>
      [`=IFERROR(INDEX(FILTER(${ingFilter('G')}),${i+1}),"")`]) });
  vals.push({ range: `${S}!D12:D${11+ING_ROWS}`, values:
    Array.from({length: ING_ROWS}, (_, i) =>
      [`=IFERROR(INDEX(FILTER(${ingFilter('E')}),${i+1}),"")`]) });
  vals.push({ range: `${S}!E12:E${11+ING_ROWS}`, values:
    Array.from({length: ING_ROWS}, (_, i) =>
      [`=IFERROR(INDEX(FILTER(${ingFilter('H')}),${i+1}),"")`]) });
  vals.push({ range: `${S}!F12:F${11+ING_ROWS}`, values:
    Array.from({length: ING_ROWS}, (_, i) =>
      [`=IFERROR(INDEX(FILTER(${ingFilter('L')}),${i+1}),"")`]) });

  // ── Instruction Formulas (columns H-M, rows 12 to 12+INST_ROWS-1) ────────
  // Filter by Recipe ID AND exact Step# (N) for sorted display
  const instFilter = (col, n) =>
    `'Recipe Instructions'!$${col}$8:$${col}$10007,'Recipe Instructions'!$B$8:$B$10007=$B$4,'Recipe Instructions'!$D$8:$D$10007=${n}`;

  vals.push({ range: `${S}!H12:H${11+INST_ROWS}`, values:
    Array.from({length: INST_ROWS}, (_, i) => [`=IF(I${12+i}="","",${i+1})`]) });
  vals.push({ range: `${S}!I12:I${11+INST_ROWS}`, values:
    Array.from({length: INST_ROWS}, (_, i) =>
      [`=IFERROR(INDEX(FILTER(${instFilter('E', i+1)}),1),"")`]) });
  vals.push({ range: `${S}!J12:J${11+INST_ROWS}`, values:
    Array.from({length: INST_ROWS}, (_, i) =>
      [`=IFERROR(INDEX(FILTER(${instFilter('F', i+1)}),1),"")`]) });
  vals.push({ range: `${S}!K12:K${11+INST_ROWS}`, values:
    Array.from({length: INST_ROWS}, (_, i) =>
      [`=IFERROR(INDEX(FILTER(${instFilter('G', i+1)}),1),"")`]) });
  vals.push({ range: `${S}!L12:L${11+INST_ROWS}`, values:
    Array.from({length: INST_ROWS}, (_, i) =>
      [`=IFERROR(INDEX(FILTER(${instFilter('H', i+1)}),1),"")`]) });
  vals.push({ range: `${S}!M12:M${11+INST_ROWS}`, values:
    Array.from({length: INST_ROWS}, (_, i) =>
      [`=IFERROR(INDEX(FILTER(${instFilter('I', i+1)}),1),"")`]) });

  // ── Column Widths ─────────────────────────────────────────────────────────
  const colWidths = [35, 70, 65, 200, 110, 60,  // A-F ingredients
                     18,                          // G spacer
                     45, 85, 290, 60, 105, 75,   // H-M instructions
                     20, 20, 20, 20, 20, 20, 55]; // N-T (helper)
  colWidths.forEach((w, i) => {
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
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 8 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 8, endIndex: 9 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 9, endIndex: 10 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 10, endIndex: 11 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 11, endIndex: 11 + INST_ROWS },
    properties: { pixelSize: 22 }, fields: 'pixelSize',
  } });

  // ── Freeze Rows ───────────────────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, gridProperties: { frozenRowCount: 11 } },
    fields: 'gridProperties.frozenRowCount',
  } });

  // ── Tab Color ─────────────────────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, tabColorStyle: { rgbColor: hex(C.lavender) } },
    fields: 'tabColorStyle',
  } });

  await batchUpdate(id, reqs, 'details-fmt');
  await valuesBatchUpdate(id, vals, 'details-data');
  console.log('✓ Recipe Details tab complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
