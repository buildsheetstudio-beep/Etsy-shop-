'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Personal & Household Setup'];
const S = "'Personal & Household Setup'";
const REF = "'Reference Data'";

// Sample couple data: James & Patricia Whitmore
// P1 = James Whitmore  DOB 1970-03-15  age 56  retire 67
// P2 = Patricia Whitmore  DOB 1972-09-28  age 53  retire 65

(async () => {
  const vals = [];
  const fmt  = [];

  // ----- Column widths -----
  const colWidths = [200, 165, 14, 200, 165, 14, 185, 150];
  colWidths.forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  // Background fill
  fmt.push({ repeatCell: { range: gridRange(SID,0,200,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // ==================== ROW 0 (row 1): TITLE ====================
  vals.push({ range: `${S}!A1`, values: [['PERSONAL & HOUSEHOLD SETUP']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,8), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // ROW 1 (row 2): subtitle
  vals.push({ range: `${S}!A2`, values: [['Enter your household profile below. All projection tabs reference these inputs. Yellow cells are editable.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,8), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrB),
    textFormat: { fontSize: 9, foregroundColor: hex(C.primaryText), italic: true, fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // ==================== ROW 3: HOUSEHOLD CONFIGURATION ====================
  // Section header row 3 (0-indexed row 2)
  vals.push({ range: `${S}!A3`, values: [['HOUSEHOLD CONFIGURATION']] });
  fmt.push({ mergeCells: { range: gridRange(SID,2,3,0,8), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Row 4 (0-indexed 3): Household Type
  vals.push({ range: `${S}!A4`, values: [['Household Type', 'Couple']] });
  fmt.push({ repeatCell: { range: gridRange(SID,3,4,0,1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow),
    textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    verticalAlignment: 'MIDDLE', padding: { left: 6 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,3,4,1,2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input),
    textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  // Data validation — Household Type dropdown
  fmt.push({ setDataValidation: { range: gridRange(SID,3,4,1,2), rule: {
    condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$4:$A$5` }] },
    showCustomUi: true, strict: true,
  }}});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 32 }, fields: 'pixelSize' }});

  // Row 5 (0-indexed 4): spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 },
    properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // ==================== PARTICIPANT HEADERS (row 6, 0-indexed 5) ====================
  vals.push({ range: `${S}!A6`, values: [['PARTICIPANT 1', '', '', 'PARTICIPANT 2']] });
  // P1 header: A6:B6
  fmt.push({ mergeCells: { range: gridRange(SID,5,6,0,2), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,5,6,0,2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  // Separator col C
  fmt.push({ repeatCell: { range: gridRange(SID,5,6,2,3), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});
  // P2 header: D6:E6
  fmt.push({ mergeCells: { range: gridRange(SID,5,6,3,5), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,5,6,3,5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrC),
    textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 },
    properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // ==================== PARTICIPANT DATA ROWS (rows 7-20, 0-indexed 6-19) ====================
  const FIELDS = [
    { label: 'Full Name',                   p1: 'James Whitmore',           p2: 'Patricia Whitmore',      type: 'text' },
    { label: 'Date of Birth',               p1: '1970-03-15',               p2: '1972-09-28',             type: 'date' },
    { label: 'Current Age',                 p1: `=IFERROR(DATEDIF(B8,TODAY(),"Y"),"")`, p2: `=IFERROR(DATEDIF(E8,TODAY(),"Y"),"")`, type: 'formula' },
    { label: 'Target Retirement Age',       p1: 67,                         p2: 65,                       type: 'number' },
    { label: 'Projected Retirement Year',   p1: `=IFERROR(YEAR(B8)+B10,"")`,  p2: `=IFERROR(YEAR(E8)+E10,"")`,  type: 'formula' },
    { label: 'Years to Retirement',         p1: `=IFERROR(B11-YEAR(TODAY()),"")`, p2: `=IFERROR(E11-YEAR(TODAY()),"")`, type: 'formula' },
    { label: 'Employment Status',           p1: 'Full-time',                p2: 'Full-time',              type: 'text' },
    { label: 'Annual Gross Income',         p1: 128500,                     p2: 95000,                    type: 'currency' },
    { label: 'Annual Savings Rate',         p1: 0.18,                       p2: 0.14,                     type: 'percent' },
    { label: 'Annual Savings Contribution', p1: `=IFERROR(B14*B15,"")`,   p2: `=IFERROR(E14*E15,"")`,   type: 'currency_formula' },
    { label: 'Risk Profile',                p1: 'Moderate',                 p2: 'Moderate-Conservative',  type: 'text' },
    { label: 'Social Security Claim Age',   p1: 67,                         p2: 67,                       type: 'number' },
    { label: 'Est. SS Benefit (Monthly)',   p1: 3400,                       p2: 2750,                     type: 'currency' },
    { label: 'Est. SS Benefit (Annual)',    p1: `=IFERROR(B19*12,"")`,    p2: `=IFERROR(E19*12,"")`,    type: 'currency_formula' },
  ];

  FIELDS.forEach((f, fi) => {
    const rIdx = 6 + fi; // 0-indexed row
    const rowNum = rIdx + 1; // 1-indexed

    // Label col A
    vals.push({ range: `${S}!A${rowNum}`, values: [[f.label]] });
    // P1 value col B
    vals.push({ range: `${S}!B${rowNum}`, values: [[f.p1]] });
    // P2 label col D (same label)
    vals.push({ range: `${S}!D${rowNum}`, values: [[f.label]] });
    // P2 value col E
    vals.push({ range: `${S}!E${rowNum}`, values: [[f.p2]] });

    const isEven = fi % 2 === 0;
    const rowBg = isEven ? C.panel : C.altRow;

    // Format label cols A and D
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,0,1), cell: { userEnteredFormat: {
      backgroundColor: hex(rowBg),
      textFormat: { fontSize: 9, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE', padding: { left: 6 },
      borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,3,4), cell: { userEnteredFormat: {
      backgroundColor: hex(rowBg),
      textFormat: { fontSize: 9, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE', padding: { left: 6 },
      borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
    }}, fields: 'userEnteredFormat' }});

    // Input vs formula background
    const isFormula = f.type === 'formula' || f.type === 'currency_formula';
    const valBg = isFormula ? C.formula : C.input;

    // P1 value col B
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,1,2), cell: { userEnteredFormat: {
      backgroundColor: hex(valBg),
      textFormat: { bold: isFormula, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE',
      horizontalAlignment: ['currency','number','currency_formula','percent'].includes(f.type) ? 'RIGHT' : 'LEFT',
      borders: { bottom: { style: 'SOLID', color: hex(C.border) }, right: { style: 'SOLID', color: hex(C.border) } },
    }}, fields: 'userEnteredFormat' }});

    // P2 value col E
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,4,5), cell: { userEnteredFormat: {
      backgroundColor: hex(valBg),
      textFormat: { bold: isFormula, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE',
      horizontalAlignment: ['currency','number','currency_formula','percent'].includes(f.type) ? 'RIGHT' : 'LEFT',
      borders: { bottom: { style: 'SOLID', color: hex(C.border) }, right: { style: 'SOLID', color: hex(C.border) } },
    }}, fields: 'userEnteredFormat' }});

    // Number formats
    if (f.type === 'currency' || f.type === 'currency_formula') {
      fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,1,2), cell: { userEnteredFormat: {
        numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' },
      }}, fields: 'userEnteredFormat.numberFormat' }});
      fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,4,5), cell: { userEnteredFormat: {
        numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' },
      }}, fields: 'userEnteredFormat.numberFormat' }});
    } else if (f.type === 'percent') {
      fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,1,2), cell: { userEnteredFormat: {
        numberFormat: { type: 'PERCENT', pattern: '0%' },
      }}, fields: 'userEnteredFormat.numberFormat' }});
      fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,4,5), cell: { userEnteredFormat: {
        numberFormat: { type: 'PERCENT', pattern: '0%' },
      }}, fields: 'userEnteredFormat.numberFormat' }});
    }

    // Dropdowns for Risk Profile and Employment Status
    if (f.label === 'Risk Profile') {
      fmt.push({ setDataValidation: { range: gridRange(SID,rIdx,rIdx+1,1,2), rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$95:$A$98` }] },
        showCustomUi: true, strict: true,
      }}});
      fmt.push({ setDataValidation: { range: gridRange(SID,rIdx,rIdx+1,4,5), rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$95:$A$98` }] },
        showCustomUi: true, strict: true,
      }}});
    }

    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: rIdx, endIndex: rIdx+1 },
      properties: { pixelSize: 24 }, fields: 'pixelSize' }});
  });

  // Row 20 (0-indexed 20): spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 20, endIndex: 21 },
    properties: { pixelSize: 10 }, fields: 'pixelSize' }});

  // ==================== ASSUMPTIONS SECTION (rows 21-31, 0-indexed 20-30) ====================
  const ASS_START = 21; // 0-indexed row of assumption header

  vals.push({ range: `${S}!A${ASS_START+1}`, values: [['HOUSEHOLD ASSUMPTIONS & GOALS']] });
  fmt.push({ mergeCells: { range: gridRange(SID, ASS_START, ASS_START+1, 0, 8), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID, ASS_START, ASS_START+1, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: ASS_START, endIndex: ASS_START+1 },
    properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  const ASSUMPTIONS_DATA = [
    { label: 'Inflation Rate (Annual)',           val: 0.03,   fmt: 'PERCENT', pat: '0.0%', editable: true  },
    { label: 'Healthcare Inflation (Annual)',      val: 0.055,  fmt: 'PERCENT', pat: '0.0%', editable: true  },
    { label: 'Expected Portfolio Return',          val: 0.07,   fmt: 'PERCENT', pat: '0.0%', editable: true  },
    { label: 'Life Expectancy (Years)',            val: 90,     fmt: 'NUMBER',  pat: '0',    editable: true  },
    { label: 'Social Security COLA',              val: 0.025,  fmt: 'PERCENT', pat: '0.0%', editable: true  },
    { label: 'Safe Withdrawal Rate',              val: 0.04,   fmt: 'PERCENT', pat: '0.0%', editable: true  },
    { label: 'Monthly Retirement Income Goal',    val: 9500,   fmt: 'CURRENCY',pat: '"$"#,##0', editable: true },
    { label: 'Annual Retirement Income Goal',     val: `=IFERROR(B29*12,"")`, fmt: 'CURRENCY', pat: '"$"#,##0', editable: false },
    { label: 'Target Nest Egg (4% Rule)',         val: `=IFERROR(B30/B28,"")`, fmt: 'CURRENCY', pat: '"$"#,##0', editable: false },
    { label: 'Combined Household Income (P1+P2)', val: `=IFERROR(B14+E14,"")`, fmt: 'CURRENCY', pat: '"$"#,##0', editable: false },
    { label: 'Combined Annual Savings (P1+P2)',   val: `=IFERROR(B16+E16,"")`, fmt: 'CURRENCY', pat: '"$"#,##0', editable: false },
  ];

  ASSUMPTIONS_DATA.forEach((a, ai) => {
    const rIdx = ASS_START + 1 + ai;
    const rowNum = rIdx + 1;
    const isEven = ai % 2 === 0;
    const rowBg = isEven ? C.panel : C.altRow;
    const valBg = a.editable ? C.input : C.formula;

    vals.push({ range: `${S}!A${rowNum}`, values: [[a.label]] });
    vals.push({ range: `${S}!B${rowNum}`, values: [[a.val]] });

    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,0,1), cell: { userEnteredFormat: {
      backgroundColor: hex(rowBg),
      textFormat: { fontSize: 9, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE', padding: { left: 6 },
      borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,1,2), cell: { userEnteredFormat: {
      backgroundColor: hex(valBg),
      textFormat: { bold: !a.editable, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE', horizontalAlignment: 'RIGHT',
      numberFormat: { type: a.fmt, pattern: a.pat },
      borders: { bottom: { style: 'SOLID', color: hex(C.border) }, right: { style: 'SOLID', color: hex(C.border) } },
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: rIdx, endIndex: rIdx+1 },
      properties: { pixelSize: 24 }, fields: 'pixelSize' }});
  });

  // ==================== SUMMARY CARDS (rows 33-40, 0-indexed 32-39) ====================
  const CARD_START = 33; // 0-indexed

  // Spacer before cards
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: CARD_START-1, endIndex: CARD_START },
    properties: { pixelSize: 10 }, fields: 'pixelSize' }});

  vals.push({ range: `${S}!A${CARD_START+1}`, values: [['HOUSEHOLD SNAPSHOT']] });
  fmt.push({ mergeCells: { range: gridRange(SID, CARD_START, CARD_START+1, 0, 8), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID, CARD_START, CARD_START+1, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: CARD_START, endIndex: CARD_START+1 },
    properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // 8 KPI cards in 2 rows of 4, each card = 2 merged cols
  // Card row 1: rows CARD_START+1, CARD_START+2 (label/value)
  // Card row 2: rows CARD_START+4, CARD_START+5 (with spacer at CARD_START+3)
  const CARDS = [
    [
      { label: 'Household Type',      formula: '=IFERROR(B4,"—")',         fmt: null },
      { label: 'Years to Retirement', formula: '=IFERROR(MIN(B13,E13),"—")', fmt: '0 "yrs"' },
      { label: 'Combined Income',     formula: '=IFERROR(B15+E15,"—")',    fmt: '"$"#,##0' },
      { label: 'Combined Savings',    formula: '=IFERROR(B17+E17,"—")',    fmt: '"$"#,##0' },
    ],
    [
      { label: 'Monthly Income Goal', formula: '=IFERROR(B29,"—")',       fmt: '"$"#,##0' },
      { label: 'Target Nest Egg',     formula: '=IFERROR(B31,"—")',        fmt: '"$"#,##0' },
      { label: 'Est. Combined SS',    formula: '=IFERROR((B20+E20)*12,"—")', fmt: '"$"#,##0 "/ yr"' },
      { label: 'Life Expectancy',     formula: '=IFERROR(B26,"—")',        fmt: '0 "yrs"' },
    ],
  ];

  CARDS.forEach((cardRow, ri) => {
    const labelRow = CARD_START + 1 + ri * 3; // 0-indexed
    const valueRow = labelRow + 1;
    const spacerRow = labelRow + 2;

    cardRow.forEach((card, ci) => {
      const c1 = ci * 2;
      const c2 = c1 + 2;
      const colLetter = String.fromCharCode(65 + c1); // A, C, E, G

      // Label row (merged 2 cols)
      fmt.push({ mergeCells: { range: gridRange(SID, labelRow, labelRow+1, c1, c2), mergeType: 'MERGE_ALL' }});
      fmt.push({ repeatCell: { range: gridRange(SID, labelRow, labelRow+1, c1, c2), cell: { userEnteredFormat: {
        backgroundColor: hex(C.hdrB),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex('#D0D4E8'), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      }}, fields: 'userEnteredFormat' }});
      vals.push({ range: `${S}!${colLetter}${labelRow+1}`, values: [[card.label]] });

      // Value row (merged 2 cols)
      fmt.push({ mergeCells: { range: gridRange(SID, valueRow, valueRow+1, c1, c2), mergeType: 'MERGE_ALL' }});
      fmt.push({ repeatCell: { range: gridRange(SID, valueRow, valueRow+1, c1, c2), cell: { userEnteredFormat: {
        backgroundColor: hex(C.panel),
        textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
        ...(card.fmt ? { numberFormat: { type: 'NUMBER', pattern: card.fmt } } : {}),
      }}, fields: 'userEnteredFormat' }});
      vals.push({ range: `${S}!${colLetter}${valueRow+1}`, values: [[card.formula]] });
    });

    // Row heights
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: labelRow, endIndex: labelRow+1 },
      properties: { pixelSize: 22 }, fields: 'pixelSize' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: valueRow, endIndex: valueRow+1 },
      properties: { pixelSize: 40 }, fields: 'pixelSize' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: spacerRow, endIndex: spacerRow+1 },
      properties: { pixelSize: 8 }, fields: 'pixelSize' }});
  });

  // ==================== DISCLAIMER (row 43, 0-indexed 42) ====================
  const DISC_ROW = 42;
  vals.push({ range: `${S}!A${DISC_ROW+1}`, values: [[
    'DISCLAIMER: This spreadsheet is for educational planning purposes only. ' +
    'It does not constitute financial, investment, tax, or legal advice. ' +
    'Consult a licensed financial professional before making retirement decisions.',
  ]] });
  fmt.push({ mergeCells: { range: gridRange(SID, DISC_ROW, DISC_ROW+1, 0, 8), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID, DISC_ROW, DISC_ROW+1, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow),
    textFormat: { fontSize: 8, foregroundColor: hex(C.secText), italic: true, fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: DISC_ROW, endIndex: DISC_ROW+1 },
    properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Freeze top 2 rows only
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '03-setup values');
  await batchUpdate(id, fmt, '03-setup format');

  console.log('✅ Personal & Household Setup done.');
  console.log('  Key cells:');
  console.log('  B4  = Household Type');
  console.log('  B7  = P1 Name  |  E7 = P2 Name');
  console.log('  B8  = P1 DOB   |  E8 = P2 DOB');
  console.log('  B10 = P1 Retire Age | E10 = P2 Retire Age');
  console.log('  B12 = P1 Years to Retire | E12 = P2 Years to Retire');
  console.log('  B14 = P1 Income    | E14 = P2 Income');
  console.log('  B22 = Inflation Rate');
  console.log('  B25 = Expected Return');
  console.log('  B26 = Life Expectancy');
  console.log('  B28 = Safe Withdrawal Rate');
  console.log('  B29 = Monthly Income Goal');
  console.log('  B31 = Target Nest Egg');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
