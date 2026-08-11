'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID  = sheetMap['Growth Forecast'];
const S    = "'Growth Forecast'";
const SETUP = "'Portfolio Setup'";
const HOLD  = "'Holdings'";

// ── Row layout (0-indexed) ────────────────────────────────────────────
const R_TITLE   = 0;   // rows 0-1 merged
const R_NOTE    = 2;
const R_CARDHDR = 3;
const R_CARD    = 4;
const R_BLANK1  = 5;
const R_SHDR    = 6;   // Scenario Settings section header
const R_SCOLS   = 7;   // scenario col headers
const R_SCON    = 8;   // Conservative (row 9)
const R_SEXP    = 9;   // Expected (row 10)
const R_SOPT    = 10;  // Optimistic (row 11)
const R_SCU1    = 11;  // Custom 1 (row 12) — rate input at B12
const R_SCU2    = 12;  // Custom 2 (row 13) — rate input at B13
const R_BLANK2  = 13;
const R_FHDR    = 14;  // 30-Year Forecast section header
const R_FCOLS   = 15;  // forecast col headers
const R_FY0     = 16;  // first forecast year: 0-indexed 16 → 1-indexed 17 (Year 2026)
const R_FY_END  = 45;  // last forecast year: 0-indexed 45 → 1-indexed 46 (Year 2055)
const R_BLANK3  = 46;

// 1-indexed convenience
const FY0   = R_FY0 + 1;            // 17  (Year 2026 row, 1-indexed)
const FY_END= R_FY_END + 1;         // 46  (Year 2055 row, 1-indexed)
const CU1R  = '$B$12';              // Custom 1 rate cell
const CU2R  = '$B$13';              // Custom 2 rate cell
const CON_R = `${SETUP}!$B$22`;    // Conservative return %
const EXP_R = `${SETUP}!$B$23`;    // Expected return %
const OPT_R = `${SETUP}!$B$24`;    // Optimistic return %
const CON_G = `${SETUP}!$B$25`;    // Contribution growth %
const INF   = `${SETUP}!$B$26`;    // Inflation %
const FEE   = `${SETUP}!$B$27`;    // Fee %
const TARGET= `${SETUP}!$B$15`;    // Target portfolio value
const ANNC  = `${SETUP}!$B$13`;    // Annual contribution

function colLetter(c) { return String.fromCharCode(65 + c); }

(async () => {
  const vals = [];
  const fmt  = [];

  // ── Background ────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID,0,500,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // ── Column widths ─────────────────────────────────────────────────
  [
    [0,65],[1,125],[2,130],[3,130],[4,130],[5,130],[6,130],[7,130],[8,130],[9,110],
    [10,16],[11,120],[12,120],[13,120],[14,120],[15,120],
  ].forEach(([ci,px]) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // ── Title banner ──────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_TITLE, R_TITLE+2, 0, 10), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${S}!A1`, values: [['GROWTH FORECAST\n30-Year Portfolio Projection  ·  5 Scenarios  ·  Inflation-Adjusted View']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_TITLE, R_TITLE+2, 0, 10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_TITLE, endIndex: R_TITLE+2 },
    properties: { pixelSize: 46 }, fields: 'pixelSize' }});
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' }});

  // ── Disclaimer / note ────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_NOTE, R_NOTE+1, 0, 10), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${S}!A3`, values: [['DISCLAIMER: Projections are illustrative only and not guaranteed. Past performance does not predict future results. Rates, contributions, and inflation assumptions may change. Consult a financial advisor before making investment decisions.']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_NOTE, R_NOTE+1, 0, 10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.attention),
    textFormat: { fontSize: 8, italic: true, bold: false, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_NOTE, endIndex: R_NOTE+1 },
    properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // ── Summary cards (rows 3-4) ──────────────────────────────────────
  const CARDS = [
    { label: 'Current Portfolio Value', val: `=IFERROR(SUM(${HOLD}!$O$6:$O$1005),0)`, type: 'CURRENCY', pat: '$#,##0' },
    { label: 'Target Portfolio Value',  val: `=${TARGET}`,                              type: 'CURRENCY', pat: '$#,##0' },
    { label: 'Expected at Forecast End',val: `=D${FY_END}`,                            type: 'CURRENCY', pat: '$#,##0' },
    { label: '% of Target (Expected)',  val: `=IFERROR(D${FY_END}/${TARGET},0)`,       type: 'PERCENT',  pat: '0.0%' },
  ];
  const cardCols = [[0,2],[2,5],[5,7],[7,10]];
  CARDS.forEach((card, i) => {
    const [c0, c1] = cardCols[i];
    fmt.push({ mergeCells: { range: gridRange(SID, R_CARDHDR, R_CARDHDR+1, c0, c1), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID, R_CARD, R_CARD+1, c0, c1), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!${colLetter(c0)}${R_CARDHDR+1}`, values: [[card.label]] });
    vals.push({ range: `${S}!${colLetter(c0)}${R_CARD+1}`, values: [[card.val]] });
    fmt.push({ repeatCell: { range: gridRange(SID, R_CARDHDR, R_CARDHDR+1, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrB),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID, R_CARD, R_CARD+1, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.highlight),
      textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      numberFormat: { type: card.type, pattern: card.pat },
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_CARDHDR, endIndex: R_CARD+1 },
    properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // ── Blank ─────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_BLANK1, endIndex: R_BLANK1+1 },
    properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // ── Scenario Settings ─────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_SHDR, R_SHDR+1, 0, 10), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${S}!A7`, values: [['  Scenario Settings  —  Conservative / Expected / Optimistic pull from Portfolio Setup. Adjust Custom rates below.']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_SHDR, R_SHDR+1, 0, 10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrA),
    textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_SHDR, endIndex: R_SHDR+1 },
    properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Scenario col headers
  vals.push({ range: `${S}!A${R_SCOLS+1}`, values: [['Scenario','Annual Return %','Projected Value at 2055','× Current Value','Reaches $2M Target?','']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_SCOLS, R_SCOLS+1, 0, 6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrC),
    textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_SCOLS, endIndex: R_SCOLS+1 },
    properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  const SCENARIOS = [
    { name: 'Conservative', rateFormula: `=${CON_R}`, valCol: 'C', rowI: R_SCON },
    { name: 'Expected',     rateFormula: `=${EXP_R}`, valCol: 'D', rowI: R_SEXP },
    { name: 'Optimistic',   rateFormula: `=${OPT_R}`, valCol: 'E', rowI: R_SOPT },
    { name: 'Custom 1',     rateFormula: null,         valCol: 'F', rowI: R_SCU1, inputRate: 0.06 },
    { name: 'Custom 2',     rateFormula: null,         valCol: 'G', rowI: R_SCU2, inputRate: 0.085 },
  ];

  SCENARIOS.forEach((sc, i) => {
    const r0 = sc.rowI;
    const r  = r0 + 1;
    const bg = i % 2 === 0 ? C.panel : C.altRow;

    vals.push({ range: `${S}!A${r}`, values: [[sc.name]] });
    // B: rate
    if (sc.rateFormula) {
      vals.push({ range: `${S}!B${r}`, values: [[sc.rateFormula]] });
    } else {
      vals.push({ range: `${S}!B${r}`, values: [[sc.inputRate]] });
    }
    // C: projected end value (referencing last forecast row)
    vals.push({ range: `${S}!C${r}`, values: [[`=${sc.valCol}${FY_END}`]] });
    // D: multiple of current value
    vals.push({ range: `${S}!D${r}`, values: [[`=IFERROR(C${r}/SUM(${HOLD}!$O$6:$O$1005),"")`]] });
    // E: meets target?
    vals.push({ range: `${S}!E${r}`, values: [[`=IF(C${r}>=${TARGET},"Yes ✓","Not yet")`]] });

    fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 0, 6), cell: { userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { fontSize: 9, fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    // Input tint on rate cell for Custom scenarios
    if (!sc.rateFormula) {
      fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 1, 2), cell: { userEnteredFormat: {
        backgroundColor: hex(C.input),
        textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
        borders: { bottom: { style: 'SOLID', color: hex(C.secondary) } },
      }}, fields: 'userEnteredFormat' }});
    } else {
      fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 1, 2), cell: { userEnteredFormat: {
        backgroundColor: hex(C.formula),
      }}, fields: 'userEnteredFormat.backgroundColor' }});
    }
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r0, endIndex: r0+1 },
      properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  });

  // Scenario number formats
  fmt.push({ repeatCell: { range: gridRange(SID, R_SCON, R_SCU2+1, 1, 2), cell: { userEnteredFormat: {
    numberFormat: { type: 'PERCENT', pattern: '0.00%' },
  }}, fields: 'userEnteredFormat.numberFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID, R_SCON, R_SCU2+1, 2, 3), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0' },
  }}, fields: 'userEnteredFormat.numberFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID, R_SCON, R_SCU2+1, 3, 4), cell: { userEnteredFormat: {
    numberFormat: { type: 'NUMBER', pattern: '0.0"x"' },
  }}, fields: 'userEnteredFormat.numberFormat' }});
  // Formula tint on computed cols (C, D, E) for non-custom rows
  [0,2,3,4].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID, R_SCON, R_SCU2+1, c, c+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula),
    }}, fields: 'userEnteredFormat.backgroundColor' }});
  });
  // But override A col with panel bg for the name
  fmt.push({ repeatCell: { range: gridRange(SID, R_SCON, R_SCU2+1, 0, 1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.panel),
    textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat' }});
  // Conditional: "Yes ✓" → success, "Not yet" → warning
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_SCON, R_SCU2+1, 4, 5)],
    booleanRule: {
      condition: { type: 'TEXT_CONTAINS', values: [{ userEnteredValue: 'Yes' }] },
      format: { backgroundColor: hex(C.success), textFormat: { foregroundColor: hex(C.primaryText) } },
    },
  }, index: 0 }});
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_SCON, R_SCU2+1, 4, 5)],
    booleanRule: {
      condition: { type: 'TEXT_CONTAINS', values: [{ userEnteredValue: 'Not yet' }] },
      format: { backgroundColor: hex(C.warning) },
    },
  }, index: 0 }});

  // Blank
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_BLANK2, endIndex: R_BLANK2+1 },
    properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // ── 30-Year Forecast Table ────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_FHDR, R_FHDR+1, 0, 10), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${S}!A15`, values: [['  30-Year Growth Projection  —  Rates, fees, and contribution growth pull from Portfolio Setup. Adjust inputs there to update all scenarios.']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_FHDR, R_FHDR+1, 0, 10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrA),
    textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_FHDR, endIndex: R_FHDR+1 },
    properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  const FCOLS_HEADERS = [
    'Year','Annual Contribution','Conservative','Expected','Optimistic',
    'Custom 1','Custom 2','Inflation-Adj Expected','Cumulative Contributions','Target',
  ];
  vals.push({ range: `${S}!A${R_FCOLS+1}`, values: [FCOLS_HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_FCOLS, R_FCOLS+1, 0, 10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrC),
    textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_FCOLS, endIndex: R_FCOLS+1 },
    properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Forecast data rows
  const numYears = 30;  // 2026-2055
  for (let i = 0; i < numYears; i++) {
    const r0 = R_FY0 + i;
    const r  = r0 + 1;  // 1-indexed
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    const prev = r - 1; // previous row (1-indexed)

    if (i === 0) {
      // Starting row: current portfolio value, no growth yet
      vals.push({ range: `${S}!A${r}`, values: [[`=${SETUP}!$B$28`]] });     // Year 2026
      vals.push({ range: `${S}!B${r}`, values: [[`=${ANNC}`]] });             // Annual contribution
      vals.push({ range: `${S}!C${r}`, values: [[`=IFERROR(SUM(${HOLD}!$O$6:$O$1005),0)`]] }); // Conservative start
      vals.push({ range: `${S}!D${r}`, values: [[`=C${r}`]] });               // Expected start
      vals.push({ range: `${S}!E${r}`, values: [[`=C${r}`]] });               // Optimistic start
      vals.push({ range: `${S}!F${r}`, values: [[`=C${r}`]] });               // Custom 1 start
      vals.push({ range: `${S}!G${r}`, values: [[`=C${r}`]] });               // Custom 2 start
      vals.push({ range: `${S}!H${r}`, values: [[`=C${r}`]] });               // Inflation-adj start
      vals.push({ range: `${S}!I${r}`, values: [[0]] });                      // Cumulative contributions
      vals.push({ range: `${S}!J${r}`, values: [[`=${TARGET}`]] });           // Target (constant)
    } else {
      // Growth rows: chain from previous year
      vals.push({ range: `${S}!A${r}`, values: [[`=A${prev}+1`]] });
      vals.push({ range: `${S}!B${r}`, values: [[`=B${prev}*(1+${CON_G})`]] });
      vals.push({ range: `${S}!C${r}`, values: [[`=IFERROR(C${prev}*(1+${CON_R}-${FEE})+B${r},0)`]] });
      vals.push({ range: `${S}!D${r}`, values: [[`=IFERROR(D${prev}*(1+${EXP_R}-${FEE})+B${r},0)`]] });
      vals.push({ range: `${S}!E${r}`, values: [[`=IFERROR(E${prev}*(1+${OPT_R}-${FEE})+B${r},0)`]] });
      vals.push({ range: `${S}!F${r}`, values: [[`=IFERROR(F${prev}*(1+${CU1R}-${FEE})+B${r},0)`]] });
      vals.push({ range: `${S}!G${r}`, values: [[`=IFERROR(G${prev}*(1+${CU2R}-${FEE})+B${r},0)`]] });
      vals.push({ range: `${S}!H${r}`, values: [[`=IFERROR(D${r}/(1+${INF})^(A${r}-$A$${FY0}),0)`]] });
      vals.push({ range: `${S}!I${r}`, values: [[`=I${prev}+B${r}`]] });
      vals.push({ range: `${S}!J${r}`, values: [[`=${TARGET}`]] });
    }

    fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 0, 10), cell: { userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { fontSize: 9, fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r0, endIndex: r0+1 },
      properties: { pixelSize: 18 }, fields: 'pixelSize' }});
  }

  // Forecast number formats
  const D0 = R_FY0; const DN = R_FY_END + 1;
  // Year column: integer
  fmt.push({ repeatCell: { range: gridRange(SID, D0, DN, 0, 1), cell: { userEnteredFormat: {
    numberFormat: { type: 'NUMBER', pattern: '0' }, horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});
  // Currency columns (B-J)
  fmt.push({ repeatCell: { range: gridRange(SID, D0, DN, 1, 10), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0' },
  }}, fields: 'userEnteredFormat.numberFormat' }});
  // Formula tint on all computed columns
  [1,2,3,4,5,6,7,8,9].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID, D0, DN, c, c+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula),
    }}, fields: 'userEnteredFormat.backgroundColor' }});
  });
  // Year col bg (alternating already set by row loop)
  // Highlight rows where Expected value crosses target (conditional)
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, D0, DN, 3, 4)],
    booleanRule: {
      condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: `=${TARGET}` }] },
      format: { backgroundColor: hex(C.success), textFormat: { bold: true, foregroundColor: hex(C.primaryText) } },
    },
  }, index: 0 }});
  // Target column: soft styling
  fmt.push({ repeatCell: { range: gridRange(SID, D0, DN, 9, 10), cell: { userEnteredFormat: {
    textFormat: { italic: true, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
  }}, fields: 'userEnteredFormat.textFormat' }});

  // Blank after forecast
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_BLANK3, endIndex: R_BLANK3+1 },
    properties: { pixelSize: 10 }, fields: 'pixelSize' }});

  // ── Charts ────────────────────────────────────────────────────────
  // Chart 1: Multi-line — 5 scenarios + target over 30 years
  fmt.push({ addChart: { chart: {
    spec: {
      title: '30-Year Portfolio Projection — All Scenarios',
      titleTextFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      basicChart: {
        chartType: 'LINE',
        legendPosition: 'BOTTOM_LEGEND',
        domains: [{ domain: {
          sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_FCOLS, endRowIndex: R_FY_END+1, startColumnIndex: 0, endColumnIndex: 1 }] },
        }}],
        series: [
          { series: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_FCOLS, endRowIndex: R_FY_END+1, startColumnIndex: 2, endColumnIndex: 3 }] }}, targetAxis: 'LEFT_AXIS' },
          { series: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_FCOLS, endRowIndex: R_FY_END+1, startColumnIndex: 3, endColumnIndex: 4 }] }}, targetAxis: 'LEFT_AXIS' },
          { series: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_FCOLS, endRowIndex: R_FY_END+1, startColumnIndex: 4, endColumnIndex: 5 }] }}, targetAxis: 'LEFT_AXIS' },
          { series: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_FCOLS, endRowIndex: R_FY_END+1, startColumnIndex: 9, endColumnIndex: 10 }] }}, targetAxis: 'LEFT_AXIS' },
        ],
        headerCount: 1,
      },
      backgroundColor: hex(C.bg),
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_SHDR, columnIndex: 11 },
      widthPixels: 420, heightPixels: 300,
    }},
  }}});

  // Chart 2: Line — Expected vs Inflation-Adjusted + Cumulative Contributions
  fmt.push({ addChart: { chart: {
    spec: {
      title: 'Expected Value vs Inflation-Adjusted vs Contributions',
      titleTextFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      basicChart: {
        chartType: 'LINE',
        legendPosition: 'BOTTOM_LEGEND',
        domains: [{ domain: {
          sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_FCOLS, endRowIndex: R_FY_END+1, startColumnIndex: 0, endColumnIndex: 1 }] },
        }}],
        series: [
          { series: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_FCOLS, endRowIndex: R_FY_END+1, startColumnIndex: 3, endColumnIndex: 4 }] }}, targetAxis: 'LEFT_AXIS' },
          { series: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_FCOLS, endRowIndex: R_FY_END+1, startColumnIndex: 7, endColumnIndex: 8 }] }}, targetAxis: 'LEFT_AXIS' },
          { series: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_FCOLS, endRowIndex: R_FY_END+1, startColumnIndex: 8, endColumnIndex: 9 }] }}, targetAxis: 'LEFT_AXIS' },
        ],
        headerCount: 1,
      },
      backgroundColor: hex(C.bg),
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_FHDR, columnIndex: 11 },
      widthPixels: 420, heightPixels: 300,
    }},
  }}});

  // ── Flush ─────────────────────────────────────────────────────────
  await batchUpdate(id, fmt, '09-forecast');
  await valuesBatchUpdate(id, vals, '09-forecast');
  console.log('✓ Growth Forecast tab complete');
})();
