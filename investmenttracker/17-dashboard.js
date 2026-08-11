'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Investment Dashboard'];
const S   = "'Investment Dashboard'";

const CLR = {
  hdrA: hex('2F5860'), hdrB: hex('3D6E78'), hdrC: hex('547F87'),
  bg:   hex('F4F5F2'), highlight: hex('DDE8E7'), panel: hex('FFFFFF'),
  altRow: hex('E8ECEA'), input: hex('FDFCF8'), formula: hex('EDF1F3'),
  success: hex('8EAA92'), warning: hex('D2B16A'), attention: hex('B97972'),
  info: hex('AFC4D4'), border: hex('C8CBC6'), secondary: hex('A56F82'),
  text: hex('2D3436'), secText: hex('6A7375'), white: hex('FFFFFF'),
};

// Row constants (0-indexed)
const R_TITLE   = 0;
const R_NOTE    = 1;
const R_DATE    = 2;
// Section A: Portfolio Snapshot cards
const R_PAHDR   = 3;
const R_PA      = 4;
// row 5 = separator
// Section B: Allocation & Top Holdings
const R_PBHDR   = 6;
const R_PBCOLS  = 7;
const R_PBD0    = 8;
const NPB       = 8;
const R_PBD_END = R_PBD0 + NPB - 1; // 15
// row 16 = separator
// Section C: Dividend Income
const R_PCHDR   = 17;
const R_PCCOLS  = 18;
const R_PCD0    = 19;
const NPC       = 5;
const R_PCD_END = R_PCD0 + NPC - 1; // 23
// row 24 = separator
// Section D: Goals at a Glance
const R_PDHDR   = 25;
const R_PDCOLS  = 26;
const R_PDD0    = 27;
const NPD       = 6;
const R_PDD_END = R_PDD0 + NPD - 1; // 32
// row 33 = separator
// Section E: Accounts Overview
const R_PEHDR   = 34;
const R_PECOLS  = 35;
const R_PED0    = 36;
const NPE       = 12;
const R_PED_END = R_PED0 + NPE - 1; // 47
// row 48 = separator
// Section F: Watchlist Highlights
const R_PFHDR   = 49;
const R_PFCOLS  = 50;
const R_PFD0    = 51;
const NPF       = 5;
const R_PFD_END = R_PFD0 + NPF - 1; // 55
// row 56 = separator
// Disclaimer
const R_DISC    = 57;

const TOTAL_ROWS = 120;
const TOTAL_COLS = 14; // A-N

// 1-indexed row refs
const PA1   = R_PA + 1;    // 5
const PBD1  = R_PBD0 + 1;  // 9
const PBDN  = R_PBD_END + 1; // 16
const PCD1  = R_PCD0 + 1;  // 20
const PCDN  = R_PCD_END + 1; // 24
const PDD1  = R_PDD0 + 1;  // 28
const PDDN  = R_PDD_END + 1; // 33
const PED1  = R_PED0 + 1;  // 37
const PEDN  = R_PED_END + 1; // 48
const PFD1  = R_PFD0 + 1;  // 52
const PFDN  = R_PFD_END + 1; // 56
const DISC1 = R_DISC + 1;  // 58

const fmt    = [];
const vals   = [];
const charts = [];

// ── SHEET DIMENSIONS ──────────────────────────────────────────────────────────
fmt.push({
  updateSheetProperties: {
    properties: {
      sheetId: SID,
      gridProperties: { rowCount: TOTAL_ROWS, columnCount: TOTAL_COLS, frozenRowCount: 2 }
    },
    fields: 'gridProperties.rowCount,gridProperties.columnCount,gridProperties.frozenRowCount'
  }
});

// Column widths (A-N) — balanced for a 14-col dashboard
[120, 120, 100, 115, 115, 100, 95, 115, 95, 100, 115, 95, 115, 80].forEach((w, i) => {
  fmt.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize'
    }
  });
});

// Fixed row heights
[
  [R_TITLE,   52], [R_NOTE,    20], [R_DATE,    22],
  [R_PAHDR,   22], [R_PA,      72], [5,          8],
  [R_PBHDR,   30], [R_PBCOLS,  24],
  [16,         8],
  [R_PCHDR,   30], [R_PCCOLS,  24],
  [24,         8],
  [R_PDHDR,   30], [R_PDCOLS,  24],
  [33,         8],
  [R_PEHDR,   30], [R_PECOLS,  24],
  [48,         8],
  [R_PFHDR,   30], [R_PFCOLS,  24],
  [56,         8],
  [R_DISC,    44],
].forEach(([r, h]) => {
  fmt.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r + 1 },
      properties: { pixelSize: h }, fields: 'pixelSize'
    }
  });
});

// Data rows: 26px
[
  [R_PBD0, R_PBD_END], [R_PCD0, R_PCD_END],
  [R_PDD0, R_PDD_END], [R_PED0, R_PED_END], [R_PFD0, R_PFD_END],
].forEach(([start, end]) => {
  for (let r = start; r <= end; r++) {
    fmt.push({
      updateDimensionProperties: {
        range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r + 1 },
        properties: { pixelSize: 26 }, fields: 'pixelSize'
      }
    });
  }
});

// ── MERGES ────────────────────────────────────────────────────────────────────
const NC = TOTAL_COLS;
[R_TITLE, R_NOTE, R_DATE, R_PAHDR, R_PBHDR, R_PCHDR, R_PDHDR, R_PEHDR, R_PFHDR, R_DISC].forEach(r => {
  fmt.push({ mergeCells: { range: gridRange(SID, r, r+1, 0, NC), mergeType: 'MERGE_ALL' } });
});

// 4 big KPI cards (A-C, D-G, H-J, K-N)
[[0,3],[3,7],[7,10],[10,14]].forEach(([c1,c2]) => {
  fmt.push({ mergeCells: { range: gridRange(SID, R_PA, R_PA+1, c1, c2), mergeType: 'MERGE_ALL' } });
});

// ── CELL FORMATTING ───────────────────────────────────────────────────────────
// Title
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_TITLE, R_TITLE+1, 0, NC),
    cell: { userEnteredFormat: {
      backgroundColor: CLR.hdrA,
      textFormat: { foregroundColor: CLR.white, bold: true, fontSize: 22 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE'
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }
});

// Note
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_NOTE, R_NOTE+1, 0, NC),
    cell: { userEnteredFormat: {
      backgroundColor: CLR.hdrB,
      textFormat: { foregroundColor: CLR.white, italic: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE'
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }
});

// Date row
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_DATE, R_DATE+1, 0, NC),
    cell: { userEnteredFormat: {
      backgroundColor: CLR.bg,
      textFormat: { foregroundColor: CLR.secText, italic: true, fontSize: 9 },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE'
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }
});

// Section A header
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_PAHDR, R_PAHDR+1, 0, NC),
    cell: { userEnteredFormat: {
      backgroundColor: CLR.hdrC,
      textFormat: { foregroundColor: CLR.white, bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE'
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }
});

// Cards row
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_PA, R_PA+1, 0, NC),
    cell: { userEnteredFormat: {
      backgroundColor: CLR.panel,
      textFormat: { foregroundColor: CLR.hdrA, bold: true, fontSize: 20 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE'
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }
});

// Section headers B-F
[R_PBHDR, R_PCHDR, R_PDHDR, R_PEHDR, R_PFHDR].forEach(r => {
  fmt.push({
    repeatCell: {
      range: gridRange(SID, r, r+1, 0, NC),
      cell: { userEnteredFormat: {
        backgroundColor: CLR.hdrA,
        textFormat: { foregroundColor: CLR.white, bold: true, fontSize: 12 },
        horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE'
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }
  });
});

// Column header rows
[[R_PBCOLS, 14],[R_PCCOLS, 8],[R_PDCOLS, 10],[R_PECOLS, 10],[R_PFCOLS, 10]].forEach(([r, cols]) => {
  fmt.push({
    repeatCell: {
      range: gridRange(SID, r, r+1, 0, cols),
      cell: { userEnteredFormat: {
        backgroundColor: CLR.hdrB,
        textFormat: { foregroundColor: CLR.white, bold: true, fontSize: 9 },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP'
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)'
    }
  });
});

// Data rows (alternating)
[
  [R_PBD0, R_PBD_END, 14],
  [R_PCD0, R_PCD_END, 8],
  [R_PDD0, R_PDD_END, 10],
  [R_PED0, R_PED_END, 10],
  [R_PFD0, R_PFD_END, 10],
].forEach(([start, end, cols]) => {
  for (let r = start; r <= end; r++) {
    const even = (r - start) % 2 === 1;
    fmt.push({
      repeatCell: {
        range: gridRange(SID, r, r+1, 0, cols),
        cell: { userEnteredFormat: {
          backgroundColor: even ? CLR.altRow : CLR.panel,
          textFormat: { foregroundColor: CLR.text, fontSize: 10 },
          verticalAlignment: 'MIDDLE'
        }},
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)'
      }
    });
  }
});

// Separators
[5, 16, 24, 33, 48, 56].forEach(r => {
  fmt.push({
    repeatCell: {
      range: gridRange(SID, r, r+1, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: CLR.bg } },
      fields: 'userEnteredFormat(backgroundColor)'
    }
  });
});

// Disclaimer
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_DISC, R_DISC+1, 0, NC),
    cell: { userEnteredFormat: {
      backgroundColor: CLR.highlight,
      textFormat: { foregroundColor: CLR.secText, italic: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP'
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)'
  }
});

// ── NUMBER FORMATS ────────────────────────────────────────────────────────────
// Card 1 (A-C): total portfolio — currency
fmt.push({ repeatCell: { range: gridRange(SID, R_PA, R_PA+1, 0, 3),   cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
// Card 2 (D-G): total net worth — currency
fmt.push({ repeatCell: { range: gridRange(SID, R_PA, R_PA+1, 3, 7),   cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
// Card 3 (H-J): unrealized G/L — currency with sign
fmt.push({ repeatCell: { range: gridRange(SID, R_PA, R_PA+1, 7, 10),  cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
// Card 4 (K-N): annual dividends — currency
fmt.push({ repeatCell: { range: gridRange(SID, R_PA, R_PA+1, 10, 14), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });

// Section B: Alloc% (C=2), GL$ (D=3), GL% (E=4), Cost (F=5)
fmt.push({ repeatCell: { range: gridRange(SID, R_PBD0, R_PBD_END+1, 2, 3), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } });
fmt.push({ repeatCell: { range: gridRange(SID, R_PBD0, R_PBD_END+1, 3, 5), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
fmt.push({ repeatCell: { range: gridRange(SID, R_PBD0, R_PBD_END+1, 5, 6), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } });
// Right side (cols H-M = 7-12): alloc%, value, cost, GL$, GL%
fmt.push({ repeatCell: { range: gridRange(SID, R_PBD0, R_PBD_END+1, 8, 9),   cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } });
fmt.push({ repeatCell: { range: gridRange(SID, R_PBD0, R_PBD_END+1, 9, 11),  cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
fmt.push({ repeatCell: { range: gridRange(SID, R_PBD0, R_PBD_END+1, 11, 13), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
fmt.push({ repeatCell: { range: gridRange(SID, R_PBD0, R_PBD_END+1, 13, 14), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

// Section C: dividend amounts (B=1 to E=4), yield% (F=5)
fmt.push({ repeatCell: { range: gridRange(SID, R_PCD0, R_PCD_END+1, 1, 5), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
fmt.push({ repeatCell: { range: gridRange(SID, R_PCD0, R_PCD_END+1, 5, 6), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.00%' } } }, fields: 'userEnteredFormat.numberFormat' } });

// Section D: target/current currency (F=5, G=6), progress% (H=7)
fmt.push({ repeatCell: { range: gridRange(SID, R_PDD0, R_PDD_END+1, 5, 7), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
fmt.push({ repeatCell: { range: gridRange(SID, R_PDD0, R_PDD_END+1, 7, 8), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

// Section E: Total Value (D=3), Holdings (E=4), Cash (F=5)
fmt.push({ repeatCell: { range: gridRange(SID, R_PED0, R_PED_END+1, 3, 6), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });

// Section F: Current Price (D=3), Target Price (E=4), Upside% (F=5)
fmt.push({ repeatCell: { range: gridRange(SID, R_PFD0, R_PFD_END+1, 3, 5), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
fmt.push({ repeatCell: { range: gridRange(SID, R_PFD0, R_PFD_END+1, 5, 6), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

// ── CONDITIONAL FORMATTING ─────────────────────────────────────────────────────
// Section B allocation gradient
fmt.push({
  addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(SID, R_PBD0, R_PBD_END+1, 2, 3)],
      gradientRule: { minpoint: { color: CLR.panel, type: 'MIN' }, maxpoint: { color: CLR.success, type: 'MAX' } }
    }, index: 0
  }
});

// Section B: G/L $ positive/negative text color
fmt.push({
  addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(SID, R_PBD0, R_PBD_END+1, 3, 4)],
      booleanRule: {
        condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0' }] },
        format: { textFormat: { foregroundColor: CLR.success } }
      }
    }, index: 0
  }
});
fmt.push({
  addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(SID, R_PBD0, R_PBD_END+1, 3, 4)],
      booleanRule: {
        condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
        format: { textFormat: { foregroundColor: CLR.attention } }
      }
    }, index: 0
  }
});

// Section D: Goals Status colors
[
  { val: 'Achieved',    bg: CLR.success   },
  { val: 'In Progress', bg: CLR.info      },
  { val: 'Delayed',     bg: CLR.warning   },
  { val: 'Not Started', bg: CLR.altRow    },
  { val: 'Reassess',    bg: CLR.attention },
].forEach(({ val, bg }) => {
  fmt.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(SID, R_PDD0, R_PDD_END+1, 0, 10)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$C${PDD1}="${val}"` }] },
          format: { backgroundColor: bg }
        }
      }, index: 0
    }
  });
});

// Section F: Upside % text color
fmt.push({
  addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(SID, R_PFD0, R_PFD_END+1, 5, 6)],
      booleanRule: {
        condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0' }] },
        format: { textFormat: { foregroundColor: CLR.success } }
      }
    }, index: 0
  }
});
fmt.push({
  addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(SID, R_PFD0, R_PFD_END+1, 5, 6)],
      booleanRule: {
        condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
        format: { textFormat: { foregroundColor: CLR.attention } }
      }
    }, index: 0
  }
});

// ── BORDERS ───────────────────────────────────────────────────────────────────
const solidBorder = { style: 'SOLID', color: CLR.border };
[
  [R_PBCOLS, R_PBD_END+1, 14],
  [R_PCCOLS, R_PCD_END+1, 8],
  [R_PDCOLS, R_PDD_END+1, 10],
  [R_PECOLS, R_PED_END+1, 10],
  [R_PFCOLS, R_PFD_END+1, 10],
].forEach(([r1, r2, cols]) => {
  fmt.push({
    updateBorders: {
      range: gridRange(SID, r1, r2, 0, cols),
      top: solidBorder, bottom: solidBorder, left: solidBorder,
      right: solidBorder, innerHorizontal: solidBorder, innerVertical: solidBorder
    }
  });
});

// ── CHARTS ────────────────────────────────────────────────────────────────────
const arSID = sheetMap['Annual Review'];
const R_CH0 = R_DISC + 2; // start charts below disclaimer

// Chart 1: Asset Allocation DONUT — from Section B (cols A=AssetClass, C=Alloc%)
charts.push({
  addChart: {
    chart: {
      spec: {
        title: 'Portfolio Allocation by Asset Class',
        pieChart: {
          legendPosition: 'RIGHT_LEGEND', pieHole: 0.45,
          domain: { sourceRange: { sources: [gridRange(SID, R_PBD0, R_PBD_END+1, 0, 1)] } },
          series: { sourceRange: { sources: [gridRange(SID, R_PBD0, R_PBD_END+1, 2, 3)] } }
        }
      },
      position: { overlayPosition: { anchorCell: { sheetId: SID, rowIndex: R_CH0, columnIndex: 0 }, widthPixels: 420, heightPixels: 300 } }
    }
  }
});

// Chart 2: Holdings Value by Asset Class — COLUMN
charts.push({
  addChart: {
    chart: {
      spec: {
        title: 'Market Value by Asset Class',
        basicChart: {
          chartType: 'COLUMN', legendPosition: 'NONE',
          axis: [{ position: 'BOTTOM_AXIS', title: 'Asset Class' }, { position: 'LEFT_AXIS', title: 'Value ($)' }],
          domains: [{ domain: { sourceRange: { sources: [gridRange(SID, R_PBD0, R_PBD_END+1, 0, 1)] } } }],
          series: [{ series: { sourceRange: { sources: [gridRange(SID, R_PBD0, R_PBD_END+1, 3, 4)] } }, targetAxis: 'LEFT_AXIS', color: CLR.hdrA }],
          headerCount: 0
        }
      },
      position: { overlayPosition: { anchorCell: { sheetId: SID, rowIndex: R_CH0, columnIndex: 5 }, widthPixels: 420, heightPixels: 300 } }
    }
  }
});

// Chart 3: Annual Dividend Trend — LINE (from Annual Review KPI table)
charts.push({
  addChart: {
    chart: {
      spec: {
        title: 'Annual Dividend Income Trend',
        basicChart: {
          chartType: 'LINE', legendPosition: 'NONE',
          axis: [{ position: 'BOTTOM_AXIS', title: 'Year' }, { position: 'LEFT_AXIS', title: 'Dividends ($)' }],
          domains: [{ domain: { sourceRange: { sources: [{ sheetId: arSID, startRowIndex: 22, endRowIndex: 28, startColumnIndex: 0, endColumnIndex: 1 }] } } }],
          series: [{ series: { sourceRange: { sources: [{ sheetId: arSID, startRowIndex: 22, endRowIndex: 28, startColumnIndex: 6, endColumnIndex: 7 }] } }, targetAxis: 'LEFT_AXIS', color: CLR.success }],
          headerCount: 0
        }
      },
      position: { overlayPosition: { anchorCell: { sheetId: SID, rowIndex: R_CH0 + 18, columnIndex: 0 }, widthPixels: 420, heightPixels: 300 } }
    }
  }
});

// Chart 4: Net Worth Over Time — LINE (from Annual Review KPI table)
charts.push({
  addChart: {
    chart: {
      spec: {
        title: 'Net Worth Over Time',
        basicChart: {
          chartType: 'LINE', legendPosition: 'NONE',
          axis: [{ position: 'BOTTOM_AXIS', title: 'Year' }, { position: 'LEFT_AXIS', title: 'Net Worth ($)' }],
          domains: [{ domain: { sourceRange: { sources: [{ sheetId: arSID, startRowIndex: 22, endRowIndex: 28, startColumnIndex: 0, endColumnIndex: 1 }] } } }],
          series: [{ series: { sourceRange: { sources: [{ sheetId: arSID, startRowIndex: 22, endRowIndex: 28, startColumnIndex: 1, endColumnIndex: 2 }] } }, targetAxis: 'LEFT_AXIS', color: CLR.secondary }],
          headerCount: 0
        }
      },
      position: { overlayPosition: { anchorCell: { sheetId: SID, rowIndex: R_CH0 + 18, columnIndex: 5 }, widthPixels: 420, heightPixels: 300 } }
    }
  }
});

// ── VALUES ────────────────────────────────────────────────────────────────────
vals.push({ range: `${S}!A1`, values: [['  Walsh Family Investment Dashboard']] });
vals.push({ range: `${S}!A2`, values: [['All values are live — formulas pull automatically from Portfolio, Holdings, Dividends, Goals, and other tabs']] });
vals.push({ range: `${S}!A3`, values: [[`=IFERROR("Last updated: "&TEXT(NOW(),"mmmm d, yyyy"),"")`]] });

// Section A header + cards
vals.push({ range: `${S}!A4`, values: [['  PORTFOLIO SNAPSHOT']] });

const pfVal   = `IFERROR(SUM('Holdings'!$O$6:$O$1005),0)`;
const costBas = `IFERROR(SUM('Holdings'!$L$6:$L$1005),0)`;
const totNW   = `IFERROR('Net Worth Tracker'!$E$28,0)`;  // NW1=28, col E = Total
const totalGL = `IFERROR(${pfVal}-${costBas},0)`;
const annDiv  = `IFERROR(SUMPRODUCT((YEAR('Dividend Income'!$B$6:$B$2005)=2025)*('Dividend Income'!$S$6:$S$2005="Received")*('Dividend Income'!$O$6:$O$2005)),0)`;

vals.push({
  range: `${S}!A${PA1}:N${PA1}`,
  values: [[
    `=${pfVal}`, '', '',
    `=${totNW}`, '', '', '',
    `=${totalGL}`, '', '',
    `=${annDiv}`, '', '', ''
  ]]
});

// Section B: Allocation & Top Holdings
vals.push({ range: `${S}!A7`, values: [['  ALLOCATION  &  TOP HOLDINGS']] });
vals.push({
  range: `${S}!A8:N8`,
  values: [['Asset Class', 'Count', 'Alloc %', 'Market Value', 'Unrlzd G/L', 'G/L %', '', 'Rank', 'Alloc %', 'Market Value', 'Cost Basis', 'Unrlzd G/L', 'G/L $', 'G/L %']]
});

// Asset class breakdown (A-F, rows PBD1-PBDN)
const assetClasses = [
  'Domestic Equity', 'International Equity', 'Fixed Income',
  'Real Estate', 'Commodities', 'Crypto / Digital Assets',
  'Cash & Equivalents', 'Other'
];
const allocationRows = assetClasses.map(ac => {
  const valFml  = `IFERROR(SUMPRODUCT(('Holdings'!$I$6:$I$1005="${ac}")*('Holdings'!$O$6:$O$1005)),0)`;
  const cstFml  = `IFERROR(SUMPRODUCT(('Holdings'!$I$6:$I$1005="${ac}")*('Holdings'!$L$6:$L$1005)),0)`;
  const glFml   = `IFERROR(${valFml}-${cstFml},0)`;
  const glPFml  = `IFERROR((${valFml}-${cstFml})/${cstFml},0)`;
  const alcFml  = `IFERROR(${valFml}/${pfVal},0)`;
  const cntFml  = `IFERROR(COUNTIF('Holdings'!$I$6:$I$1005,"${ac}"),0)`;
  return [
    ac,
    `=IFERROR(${cntFml},0)`,
    `=IFERROR(${alcFml},0)`,
    `=IFERROR(${valFml},0)`,
    `=IFERROR(${glFml},0)`,
    `=IFERROR(${glPFml},0)`
  ];
});
vals.push({ range: `${S}!A${PBD1}:F${PBDN}`, values: allocationRows });

// Top 8 holdings (ranked by current value using LARGE/INDEX/MATCH)
const holdVal = `'Holdings'!$O$6:$O$1005`;
const holdAlc = `'Holdings'!$T$6:$T$1005`;
const holdCst = `'Holdings'!$L$6:$L$1005`;
const holdTkr = `'Holdings'!$F$6:$F$1005`;

const topHoldRows = Array.from({ length: NPB }, (_, i) => {
  const rank = i + 1;
  const lv   = `LARGE(${holdVal},${rank})`;
  const mr   = `IFERROR(MATCH(${lv},${holdVal},0),0)`;
  const tkr  = `IFERROR(INDEX(${holdTkr},${mr}),"")`;
  const val  = `IFERROR(${lv},0)`;
  const alc  = `IFERROR(INDEX(${holdAlc},${mr}),0)`;
  const cst  = `IFERROR(INDEX(${holdCst},${mr}),0)`;
  const glD  = `IFERROR(${val}-${cst},0)`;
  const glP  = `IFERROR((${val}-${cst})/${cst},0)`;
  return [rank, `=IFERROR(${alc},0)`, `=IFERROR(${val},0)`, `=IFERROR(${cst},0)`, `=IFERROR(${glD},0)`, `=IFERROR(${glP},0)`];
});
vals.push({ range: `${S}!H${PBD1}:M${PBDN}`, values: topHoldRows });

// Section C: Dividend Income
vals.push({ range: `${S}!A18`, values: [['  DIVIDEND INCOME BY OWNER']] });
vals.push({ range: `${S}!A19:H19`, values: [['Owner / Group', '2023 Dividends', '2024 Dividends', '2025 YTD', 'YoY Change', 'YoY % Change', 'Avg Yield', 'Notes']] });

const divByOwner = (owner, yr) => {
  if (owner === 'All') {
    return `IFERROR(SUMPRODUCT((YEAR('Dividend Income'!$B$6:$B$2005)=${yr})*('Dividend Income'!$S$6:$S$2005="Received")*('Dividend Income'!$O$6:$O$2005)),0)`;
  }
  return `IFERROR(SUMPRODUCT((YEAR('Dividend Income'!$B$6:$B$2005)=${yr})*('Dividend Income'!$H$6:$H$2005="${owner}")*('Dividend Income'!$S$6:$S$2005="Received")*('Dividend Income'!$O$6:$O$2005)),0)`;
};

const divRowDefs = [
  ['Daniel Walsh',    'Daniel Walsh'],
  ['Emily Walsh',     'Emily Walsh'],
  ['Joint Household', 'Joint Household'],
  ['All — 2024',      'All'],
  ['All — 2025',      'All'],
];
const divDataRows = divRowDefs.map(([label, owner], i) => {
  const r = PCD1 + i;
  return [
    label,
    `=${divByOwner(owner, 2023)}`,
    `=${divByOwner(owner, 2024)}`,
    `=${divByOwner(owner, 2025)}`,
    `=IFERROR(D${r}-C${r},0)`,
    `=IFERROR((D${r}-C${r})/C${r},0)`,
    '', ''
  ];
});
vals.push({ range: `${S}!A${PCD1}:H${PCDN}`, values: divDataRows });

// Section D: Goals at a Glance
vals.push({ range: `${S}!A26`, values: [['  GOALS AT A GLANCE']] });
vals.push({ range: `${S}!A27:J27`, values: [['#', 'Goal Name', 'Status', 'Type', 'Owner', 'Target', 'Current', 'Progress', 'Priority', 'Target Date']] });

const goalRows = Array.from({ length: NPD }, (_, i) => {
  const gr = 9 + i; // rows 9-14 in Goals & Milestones
  return [
    `=IFERROR('Goals & Milestones'!A${gr},"")`,
    `=IFERROR('Goals & Milestones'!B${gr},"")`,
    `=IFERROR('Goals & Milestones'!E${gr},"")`,
    `=IFERROR('Goals & Milestones'!C${gr},"")`,
    `=IFERROR('Goals & Milestones'!D${gr},"")`,
    `=IFERROR('Goals & Milestones'!G${gr},0)`,
    `=IFERROR('Goals & Milestones'!H${gr},0)`,
    `=IFERROR('Goals & Milestones'!I${gr},0)`,
    `=IFERROR('Goals & Milestones'!F${gr},"")`,
    `=IFERROR(TEXT('Goals & Milestones'!K${gr},"mm/dd/yyyy"),"")`
  ];
});
vals.push({ range: `${S}!A${PDD1}:J${PDDN}`, values: goalRows });

// Section E: Accounts Overview
vals.push({ range: `${S}!A35`, values: [['  ACCOUNTS OVERVIEW']] });
vals.push({ range: `${S}!A36:J36`, values: [['Account Name', 'Owner', 'Type', 'Total Value', 'Holdings Value', 'Cash Balance', 'Tax Treatment', 'Active?', 'Last Updated', 'Notes']] });

const acctRows = Array.from({ length: NPE }, (_, i) => {
  const ar = 6 + i; // Account Tracker rows 6-17
  return [
    `=IFERROR('Account Tracker'!B${ar},"")`,
    `=IFERROR('Account Tracker'!C${ar},"")`,
    `=IFERROR('Account Tracker'!E${ar},"")`,
    `=IFERROR('Account Tracker'!K${ar},0)`,
    `=IFERROR('Account Tracker'!J${ar},0)`,
    `=IFERROR('Account Tracker'!I${ar},0)`,
    `=IFERROR('Account Tracker'!F${ar},"")`,
    `=IFERROR('Account Tracker'!O${ar},"")`,
    `=IFERROR(TEXT('Account Tracker'!Q${ar},"mm/dd/yyyy"),"")`,
    `=IFERROR('Account Tracker'!R${ar},"")`
  ];
});
vals.push({ range: `${S}!A${PED1}:J${PEDN}`, values: acctRows });

// Section F: Watchlist Highlights
vals.push({ range: `${S}!A50`, values: [['  WATCHLIST HIGHLIGHTS']] });
vals.push({ range: `${S}!A51:J51`, values: [['Ticker', 'Security Name', 'Status', 'Current Price', 'Target Price', 'Upside %', 'Risk Level', 'Priority', 'Asset Class', 'Notes']] });

const wlRows = Array.from({ length: NPF }, (_, i) => {
  const wr = 9 + i; // Watchlist rows 9-13
  return [
    `=IFERROR(Watchlist!B${wr},"")`,
    `=IFERROR(Watchlist!C${wr},"")`,
    `=IFERROR(Watchlist!F${wr},"")`,
    `=IFERROR(Watchlist!I${wr},0)`,
    `=IFERROR(Watchlist!H${wr},0)`,
    `=IFERROR(Watchlist!J${wr},"")`,
    `=IFERROR(Watchlist!G${wr},"")`,
    `=IFERROR(Watchlist!M${wr},"")`,
    `=IFERROR(Watchlist!E${wr},"")`,
    `=IFERROR(Watchlist!P${wr},"")`
  ];
});
vals.push({ range: `${S}!A${PFD1}:J${PFDN}`, values: wlRows });

// Disclaimer
vals.push({
  range: `${S}!A${DISC1}`,
  values: [['DISCLAIMER: This spreadsheet is for personal financial tracking and educational purposes only. It does not constitute financial, investment, tax, or legal advice. Past performance does not guarantee future results. Consult a licensed financial advisor before making investment decisions. All values are self-reported and not verified by any financial institution.']]
});

async function main() {
  await batchUpdate(id, fmt, 'dashboard-fmt');
  await valuesBatchUpdate(id, vals, 'dashboard-vals');
  await batchUpdate(id, charts, 'dashboard-charts');
  console.log('17-dashboard.js complete');
}

main().catch(e => { console.error(e); process.exit(1); });
