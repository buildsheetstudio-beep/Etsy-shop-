'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Annual Review'];
const S   = "'Annual Review'";

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
const R_CTRLHDR = 2;
const R_CTRL    = 3;   // year dropdown at B4 (1-indexed)
const R_CARDHDR = 4;
const R_CARD    = 5;
// row 6 = separator
const R_CHDR   = 7;   // checklist section header
const R_CCOLS  = 8;   // checklist column headers
const R_CD0    = 9;   // checklist data start (0-indexed) = row 10 (1-indexed)
const NC       = 10;
const R_CD_END = R_CD0 + NC - 1; // = 18
// row 19 = separator
const R_KHDR   = 20;  // KPI section header
const R_KCOLS  = 21;  // KPI column headers
const R_KD0    = 22;  // KPI data start (0-indexed) = row 23 (1-indexed)
const NK       = 6;
const R_KD_END = R_KD0 + NK - 1; // = 27

const TOTAL_ROWS = 100;
const TOTAL_COLS = 14; // A-N

// 1-indexed refs
const CD1  = R_CD0 + 1;    // 10
const CDN  = R_CD_END + 1; // 19
const KD1  = R_KD0 + 1;    // 23
const KDN  = R_KD_END + 1; // 28

// Pie chart data in cols L-M (11-12) at the first 4 KPI rows
const R_PIE0 = R_KD0;
const PIE1   = KD1;

const fmt   = [];
const vals  = [];
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

// Column widths (A-N)
[45, 180, 120, 120, 100, 100, 80, 170, 85, 145, 80, 110, 55, 80].forEach((w, i) => {
  fmt.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize'
    }
  });
});

// Row heights
[
  [R_TITLE,   44], [R_NOTE,    20], [R_CTRLHDR, 18], [R_CTRL,    32],
  [R_CARDHDR, 20], [R_CARD,    60], [6,          8],
  [R_CHDR,    30], [R_CCOLS,   24],
  [19,         8],
  [R_KHDR,    30], [R_KCOLS,   26],
].forEach(([r, h]) => {
  fmt.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r + 1 },
      properties: { pixelSize: h }, fields: 'pixelSize'
    }
  });
});
for (let r = R_CD0; r <= R_CD_END; r++) {
  fmt.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r + 1 },
      properties: { pixelSize: 26 }, fields: 'pixelSize'
    }
  });
}
for (let r = R_KD0; r <= R_KD_END; r++) {
  fmt.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r + 1 },
      properties: { pixelSize: 26 }, fields: 'pixelSize'
    }
  });
}

// ── MERGES ────────────────────────────────────────────────────────────────────
fmt.push({ mergeCells: { range: gridRange(SID, R_TITLE,   R_TITLE+1,   0, TOTAL_COLS), mergeType: 'MERGE_ALL' } });
fmt.push({ mergeCells: { range: gridRange(SID, R_NOTE,    R_NOTE+1,    0, TOTAL_COLS), mergeType: 'MERGE_ALL' } });
fmt.push({ mergeCells: { range: gridRange(SID, R_CTRLHDR, R_CTRLHDR+1, 0, 2),         mergeType: 'MERGE_ALL' } });
fmt.push({ mergeCells: { range: gridRange(SID, R_CARDHDR, R_CARDHDR+1, 0, TOTAL_COLS), mergeType: 'MERGE_ALL' } });
[[0,2],[2,4],[4,6],[6,8]].forEach(([c1, c2]) => {
  fmt.push({ mergeCells: { range: gridRange(SID, R_CARD, R_CARD+1, c1, c2), mergeType: 'MERGE_ALL' } });
});
fmt.push({ mergeCells: { range: gridRange(SID, R_CHDR, R_CHDR+1, 0, TOTAL_COLS), mergeType: 'MERGE_ALL' } });
fmt.push({ mergeCells: { range: gridRange(SID, R_KHDR, R_KHDR+1, 0, TOTAL_COLS), mergeType: 'MERGE_ALL' } });

// ── CELL FORMATTING ───────────────────────────────────────────────────────────
// Title
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_TITLE, R_TITLE+1, 0, TOTAL_COLS),
    cell: { userEnteredFormat: {
      backgroundColor: CLR.hdrA,
      textFormat: { foregroundColor: CLR.white, bold: true, fontSize: 18 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE'
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }
});

// Note
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_NOTE, R_NOTE+1, 0, TOTAL_COLS),
    cell: { userEnteredFormat: {
      backgroundColor: CLR.info,
      textFormat: { foregroundColor: CLR.text, italic: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE'
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }
});

// Year control
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_CTRLHDR, R_CTRLHDR+1, 0, 4),
    cell: { userEnteredFormat: {
      backgroundColor: CLR.bg,
      textFormat: { foregroundColor: CLR.secText, bold: true, fontSize: 9 },
      verticalAlignment: 'MIDDLE'
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)'
  }
});
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_CTRL, R_CTRL+1, 0, 4),
    cell: { userEnteredFormat: {
      backgroundColor: CLR.input,
      textFormat: { foregroundColor: CLR.hdrA, fontSize: 12, bold: true },
      verticalAlignment: 'MIDDLE'
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)'
  }
});

// Cards header
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_CARDHDR, R_CARDHDR+1, 0, TOTAL_COLS),
    cell: { userEnteredFormat: {
      backgroundColor: CLR.hdrC,
      textFormat: { foregroundColor: CLR.white, bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE'
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }
});

// Cards
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_CARD, R_CARD+1, 0, 8),
    cell: { userEnteredFormat: {
      backgroundColor: CLR.panel,
      textFormat: { foregroundColor: CLR.hdrA, bold: true, fontSize: 18 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE'
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }
});

// Separator row 6
fmt.push({
  repeatCell: {
    range: gridRange(SID, 6, 7, 0, TOTAL_COLS),
    cell: { userEnteredFormat: { backgroundColor: CLR.bg } },
    fields: 'userEnteredFormat(backgroundColor)'
  }
});

// Checklist section header
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_CHDR, R_CHDR+1, 0, TOTAL_COLS),
    cell: { userEnteredFormat: {
      backgroundColor: CLR.hdrA,
      textFormat: { foregroundColor: CLR.white, bold: true, fontSize: 12 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE'
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }
});

// Checklist col headers
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_CCOLS, R_CCOLS+1, 0, 8),
    cell: { userEnteredFormat: {
      backgroundColor: CLR.hdrB,
      textFormat: { foregroundColor: CLR.white, bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE'
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }
});

// Checklist data rows (alternating)
for (let r = R_CD0; r <= R_CD_END; r++) {
  const even = (r - R_CD0) % 2 === 1;
  fmt.push({
    repeatCell: {
      range: gridRange(SID, r, r+1, 0, 8),
      cell: { userEnteredFormat: {
        backgroundColor: even ? CLR.altRow : CLR.panel,
        textFormat: { foregroundColor: CLR.text, fontSize: 10 },
        verticalAlignment: 'MIDDLE'
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)'
    }
  });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, r, r+1, 0, 1),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true, fontSize: 10 } } },
      fields: 'userEnteredFormat(horizontalAlignment,textFormat)'
    }
  });
}

// Separator row 19
fmt.push({
  repeatCell: {
    range: gridRange(SID, 19, 20, 0, TOTAL_COLS),
    cell: { userEnteredFormat: { backgroundColor: CLR.bg } },
    fields: 'userEnteredFormat(backgroundColor)'
  }
});

// KPI section header
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_KHDR, R_KHDR+1, 0, TOTAL_COLS),
    cell: { userEnteredFormat: {
      backgroundColor: CLR.hdrA,
      textFormat: { foregroundColor: CLR.white, bold: true, fontSize: 12 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE'
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }
});

// KPI col headers
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_KCOLS, R_KCOLS+1, 0, 10),
    cell: { userEnteredFormat: {
      backgroundColor: CLR.hdrB,
      textFormat: { foregroundColor: CLR.white, bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP'
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)'
  }
});

// KPI data rows
for (let r = R_KD0; r <= R_KD_END; r++) {
  const even = (r - R_KD0) % 2 === 1;
  fmt.push({
    repeatCell: {
      range: gridRange(SID, r, r+1, 0, 10),
      cell: { userEnteredFormat: {
        backgroundColor: even ? CLR.altRow : CLR.panel,
        textFormat: { foregroundColor: CLR.text, fontSize: 10 },
        verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER'
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)'
    }
  });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, r, r+1, 0, 1),
      cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 11 }, horizontalAlignment: 'LEFT' } },
      fields: 'userEnteredFormat(textFormat,horizontalAlignment)'
    }
  });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, r, r+1, 9, 10),
      cell: { userEnteredFormat: { textFormat: { fontSize: 9 }, horizontalAlignment: 'LEFT' } },
      fields: 'userEnteredFormat(textFormat,horizontalAlignment)'
    }
  });
}

// ── NUMBER FORMATS ────────────────────────────────────────────────────────────
// Checklist dates (E=4, F=5)
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_CD0, R_CD_END+1, 4, 6),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mm/dd/yyyy' } } },
    fields: 'userEnteredFormat.numberFormat'
  }
});

// Cards
[
  [0, 2, { type: 'CURRENCY', pattern: '"$"#,##0' }],
  [2, 4, { type: 'PERCENT',  pattern: '0%' }],
  [4, 6, { type: 'NUMBER',   pattern: '#,##0' }],
  [6, 8, { type: 'CURRENCY', pattern: '"$"#,##0' }],
].forEach(([c1, c2, nf]) => {
  fmt.push({
    repeatCell: {
      range: gridRange(SID, R_CARD, R_CARD+1, c1, c2),
      cell: { userEnteredFormat: { numberFormat: nf } },
      fields: 'userEnteredFormat.numberFormat'
    }
  });
});

// KPI: Net Worth (B=1), NW Change $ (C=2) — currency
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_KD0, R_KD_END+1, 1, 3),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } },
    fields: 'userEnteredFormat.numberFormat'
  }
});
// NW Change % (D=3)
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_KD0, R_KD_END+1, 3, 4),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } },
    fields: 'userEnteredFormat.numberFormat'
  }
});
// Portfolio (E=4), Contributions (F=5), Dividends (G=6)
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_KD0, R_KD_END+1, 4, 7),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } },
    fields: 'userEnteredFormat.numberFormat'
  }
});
// Div Yield % (H=7), Return % (I=8)
fmt.push({
  repeatCell: {
    range: gridRange(SID, R_KD0, R_KD_END+1, 7, 9),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } },
    fields: 'userEnteredFormat.numberFormat'
  }
});

// ── DATA VALIDATION ───────────────────────────────────────────────────────────
// Year dropdown (B4)
fmt.push({
  setDataValidation: {
    range: gridRange(SID, R_CTRL, R_CTRL+1, 1, 2),
    rule: {
      condition: {
        type: 'ONE_OF_LIST',
        values: ['2020','2021','2022','2023','2024','2025'].map(y => ({ userEnteredValue: y }))
      },
      strict: false, showCustomUi: true
    }
  }
});

// Checklist Category (B=1)
fmt.push({
  setDataValidation: {
    range: gridRange(SID, R_CD0, R_CD_END+1, 1, 2),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: "='Reference Data'!$A$120:$A$129" }] },
      strict: true, showCustomUi: true
    }
  }
});

// Checklist Status (C=2)
fmt.push({
  setDataValidation: {
    range: gridRange(SID, R_CD0, R_CD_END+1, 2, 3),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: "='Reference Data'!$A$132:$A$135" }] },
      strict: true, showCustomUi: true
    }
  }
});

// ── CONDITIONAL FORMATTING ─────────────────────────────────────────────────────
[
  { val: 'Complete',         bg: CLR.success  },
  { val: 'In Progress',      bg: CLR.info     },
  { val: 'Follow-Up Needed', bg: CLR.warning  },
  { val: 'Not Started',      bg: CLR.altRow   },
].forEach(({ val, bg }) => {
  fmt.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(SID, R_CD0, R_CD_END+1, 0, 8)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$C${CD1}="${val}"` }] },
          format: { backgroundColor: bg }
        }
      }, index: 0
    }
  });
});

// KPI NW Change % positive / negative
fmt.push({
  addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(SID, R_KD0, R_KD_END+1, 3, 4)],
      booleanRule: {
        condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0' }] },
        format: { backgroundColor: CLR.success }
      }
    }, index: 0
  }
});
fmt.push({
  addConditionalFormatRule: {
    rule: {
      ranges: [gridRange(SID, R_KD0, R_KD_END+1, 3, 4)],
      booleanRule: {
        condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
        format: { backgroundColor: CLR.attention }
      }
    }, index: 0
  }
});

// ── BORDERS ───────────────────────────────────────────────────────────────────
const solidBorder = { style: 'SOLID', color: CLR.border };
[[R_CCOLS, R_CD_END+1, 8], [R_KCOLS, R_KD_END+1, 10]].forEach(([r1, r2, cols]) => {
  fmt.push({
    updateBorders: {
      range: gridRange(SID, r1, r2, 0, cols),
      top: solidBorder, bottom: solidBorder, left: solidBorder,
      right: solidBorder, innerHorizontal: solidBorder, innerVertical: solidBorder
    }
  });
});

// ── CHARTS ────────────────────────────────────────────────────────────────────
const arSid = sheetMap['Annual Review'];
const CHART_ROW = R_KD_END + 2;

// Chart 1: Annual Dividends — COLUMN
charts.push({
  addChart: {
    chart: {
      spec: {
        title: 'Annual Dividends Received',
        basicChart: {
          chartType: 'COLUMN', legendPosition: 'NO_LEGEND',
          axis: [{ position: 'BOTTOM_AXIS', title: 'Year' }, { position: 'LEFT_AXIS', title: 'Dividends ($)' }],
          domains: [{ domain: { sourceRange: { sources: [gridRange(SID, R_KD0, R_KD_END+1, 0, 1)] } } }],
          series: [{ series: { sourceRange: { sources: [gridRange(SID, R_KD0, R_KD_END+1, 6, 7)] } }, targetAxis: 'LEFT_AXIS', color: CLR.success }],
          headerCount: 0
        }
      },
      position: { overlayPosition: { anchorCell: { sheetId: SID, rowIndex: CHART_ROW, columnIndex: 0 }, widthPixels: 400, heightPixels: 270 } }
    }
  }
});

// Chart 2: Net Worth Trend — LINE
charts.push({
  addChart: {
    chart: {
      spec: {
        title: 'Year-End Net Worth Trend',
        basicChart: {
          chartType: 'LINE', legendPosition: 'NO_LEGEND',
          axis: [{ position: 'BOTTOM_AXIS', title: 'Year' }, { position: 'LEFT_AXIS', title: 'Net Worth ($)' }],
          domains: [{ domain: { sourceRange: { sources: [gridRange(SID, R_KD0, R_KD_END+1, 0, 1)] } } }],
          series: [{ series: { sourceRange: { sources: [gridRange(SID, R_KD0, R_KD_END+1, 1, 2)] } }, targetAxis: 'LEFT_AXIS', color: CLR.hdrA }],
          headerCount: 0
        }
      },
      position: { overlayPosition: { anchorCell: { sheetId: SID, rowIndex: CHART_ROW, columnIndex: 5 }, widthPixels: 400, heightPixels: 270 } }
    }
  }
});

// Chart 3: Contributions — COLUMN
charts.push({
  addChart: {
    chart: {
      spec: {
        title: 'Annual Contributions',
        basicChart: {
          chartType: 'COLUMN', legendPosition: 'NO_LEGEND',
          axis: [{ position: 'BOTTOM_AXIS', title: 'Year' }, { position: 'LEFT_AXIS', title: 'Contributions ($)' }],
          domains: [{ domain: { sourceRange: { sources: [gridRange(SID, R_KD0, R_KD_END+1, 0, 1)] } } }],
          series: [{ series: { sourceRange: { sources: [gridRange(SID, R_KD0, R_KD_END+1, 5, 6)] } }, targetAxis: 'LEFT_AXIS', color: CLR.secondary }],
          headerCount: 0
        }
      },
      position: { overlayPosition: { anchorCell: { sheetId: SID, rowIndex: CHART_ROW + 16, columnIndex: 0 }, widthPixels: 400, heightPixels: 270 } }
    }
  }
});

// Chart 4: Portfolio Return % — COLUMN
charts.push({
  addChart: {
    chart: {
      spec: {
        title: 'Annual Portfolio Return %',
        basicChart: {
          chartType: 'COLUMN', legendPosition: 'NO_LEGEND',
          axis: [{ position: 'BOTTOM_AXIS', title: 'Year' }, { position: 'LEFT_AXIS', title: 'Return (%)' }],
          domains: [{ domain: { sourceRange: { sources: [gridRange(SID, R_KD0, R_KD_END+1, 0, 1)] } } }],
          series: [{ series: { sourceRange: { sources: [gridRange(SID, R_KD0, R_KD_END+1, 8, 9)] } }, targetAxis: 'LEFT_AXIS', color: CLR.warning }],
          headerCount: 0
        }
      },
      position: { overlayPosition: { anchorCell: { sheetId: SID, rowIndex: CHART_ROW + 16, columnIndex: 5 }, widthPixels: 400, heightPixels: 270 } }
    }
  }
});

// Chart 5: Checklist Status DONUT (data in cols L-M at R_PIE0 rows)
charts.push({
  addChart: {
    chart: {
      spec: {
        title: 'Review Checklist Status',
        pieChart: {
          legendPosition: 'RIGHT_LEGEND', pieHole: 0.4,
          domain: { sourceRange: { sources: [gridRange(SID, R_PIE0, R_PIE0+4, 11, 12)] } },
          series: { sourceRange: { sources: [gridRange(SID, R_PIE0, R_PIE0+4, 12, 13)] } }
        }
      },
      position: { overlayPosition: { anchorCell: { sheetId: SID, rowIndex: CHART_ROW, columnIndex: 10 }, widthPixels: 360, heightPixels: 270 } }
    }
  }
});

// ── VALUES ────────────────────────────────────────────────────────────────────
vals.push({ range: `${S}!A1`, values: [['  Annual Review']] });
vals.push({ range: `${S}!A2`, values: [['Track yearly performance, review completion, and multi-year trends  •  Use the Year dropdown to filter card metrics']] });
vals.push({ range: `${S}!A3:B4`, values: [['REVIEW YEAR', ''], ['', '2025']] });
vals.push({ range: `${S}!A5`, values: [['  ANNUAL SNAPSHOT  (filtered by selected year)']] });

// Card formulas (row 6)
const divFml  = `SUMPRODUCT((YEAR('Dividend Income'!$B$6:$B$2005)=VALUE($B$4))*('Dividend Income'!$S$6:$S$2005="Received")*('Dividend Income'!$O$6:$O$2005))`;
const chkFml  = `COUNTIF($C$${CD1}:$C$${CDN},"Complete")/${NC}`;
const goalFml = `COUNTIF('Goals & Milestones'!$E$9:$E$28,"Achieved")`;
const pfFml   = `SUM('Holdings'!$O$6:$O$1005)`;

vals.push({
  range: `${S}!A6:H6`,
  values: [[
    `=IFERROR(${divFml},0)`, '',
    `=IFERROR(${chkFml},0)`, '',
    `=IFERROR(${goalFml},0)`, '',
    `=IFERROR(${pfFml},0)`, ''
  ]]
});

// Checklist section header + col headers
vals.push({ range: `${S}!A8`, values: [['  ANNUAL REVIEW CHECKLIST']] });
vals.push({ range: `${S}!A9:H9`, values: [['#', 'Review Category', 'Status', 'Owner / Lead', 'Due Date', 'Completed Date', 'Priority', 'Notes & Action Items']] });

vals.push({
  range: `${S}!A${CD1}:H${CDN}`,
  values: [
    [1,  'Contributions Review',  'Not Started', 'Daniel Walsh',    '01/31/2026', '', 'High',   'Verify 401k/Roth IRA/HSA contributions hit IRS limits for the year'],
    [2,  'Performance Review',    'Not Started', 'Daniel Walsh',    '02/15/2026', '', 'High',   'Compare portfolio return vs. benchmark (60/40 VTI+BND blend)'],
    [3,  'Allocation Check',      'Not Started', 'Daniel Walsh',    '02/15/2026', '', 'High',   'Check asset class drift vs. target; rebalance if any class is >5% off'],
    [4,  'Dividend Review',       'Not Started', 'Emily Walsh',     '02/28/2026', '', 'Medium', 'Review YoY dividend growth; confirm DRIP elections are correct'],
    [5,  'Fee Analysis',          'Not Started', 'Daniel Walsh',    '02/28/2026', '', 'Medium', 'Audit expense ratios across all holdings; flag anything above 0.20%'],
    [6,  'Risk Assessment',       'Not Started', 'Daniel Walsh',    '03/15/2026', '', 'Medium', 'Reassess risk tolerance; review correlation between holdings'],
    [7,  'Goals Review',          'Not Started', 'Emily Walsh',     '01/31/2026', '', 'High',   'Update all goal current values; set new milestones for the year ahead'],
    [8,  'Tax Organization',      'Not Started', 'Daniel Walsh',    '03/15/2026', '', 'High',   'Collect 1099-DIV, 1099-B; review tax-loss harvesting opportunities'],
    [9,  'Account Maintenance',   'Not Started', 'Emily Walsh',     '02/28/2026', '', 'Low',    'Update beneficiaries; review linked accounts and access credentials'],
    [10, 'Estate & Insurance',    'Not Started', 'Joint Household', '03/31/2026', '', 'Low',    'Review life/disability coverage; confirm trust and will documents are current'],
  ]
});

// KPI section header + col headers
vals.push({ range: `${S}!A21`, values: [['  YEAR-OVER-YEAR PERFORMANCE']] });
vals.push({ range: `${S}!A22:J22`, values: [['Year', 'Year-End Net Worth', 'NW Change ($)', 'NW Change %', 'Portfolio Value', 'Contributions', "Dividends Rec'd", 'Div Yield %', 'Est. Return %', 'Notes']] });

// KPI historical data 2020-2024
vals.push({
  range: `${S}!A${KD1}:J${KD1 + 4}`,
  values: [
    [2020, 418250,       0,       0,       312500,  22000,     0, 0,       0.162, 'Pre-tracker baseline'],
    [2021, 523640,  105390,  0.2519,  412600,  28400,  3847, 0.0093, 0.211, 'Strong recovery; Roth IRA & HSA opened'],
    [2022, 482150,  -41490, -0.0792,  368200,  31200,  5124, 0.0124, -0.148,'Bear market — held positions, increased contributions'],
    [2023, 614720,  132570,  0.2749,  492300,  34600,  6382, 0.013,  0.262, 'Recovery year; Emily started 403(b) max contributions'],
    [2024, 742380,  127660,  0.2077,  608400,  38100,  7841, 0.0129, 0.188, 'Continued growth; 529 plan opened'],
  ]
});

// KPI 2025 row — formulas for live data
const nw25  = `IFERROR(SUM('Holdings'!$O$6:$O$1005),0)`;
const div25 = `IFERROR(SUMPRODUCT((YEAR('Dividend Income'!$B$6:$B$2005)=2025)*('Dividend Income'!$S$6:$S$2005="Received")*('Dividend Income'!$O$6:$O$2005)),0)`;

vals.push({
  range: `${S}!A${KDN}:J${KDN}`,
  values: [[
    2025,
    `=${nw25}`,
    `=IFERROR(B${KDN}-B${KDN-1},0)`,
    `=IFERROR((B${KDN}-B${KDN-1})/B${KDN-1},0)`,
    `=${nw25}`,
    42500,
    `=${div25}`,
    `=IFERROR(G${KDN}/E${KDN},0)`,
    0.24,
    '2025 — portfolio value is live from Holdings; dividends from Dividend Income log'
  ]]
});

// Pie chart mini-table (cols L-M, rows PIE1 to PIE1+3)
vals.push({
  range: `${S}!L${PIE1}:M${PIE1 + 3}`,
  values: [
    ['Not Started',      `=COUNTIF($C$${CD1}:$C$${CDN},"Not Started")`],
    ['In Progress',      `=COUNTIF($C$${CD1}:$C$${CDN},"In Progress")`],
    ['Complete',         `=COUNTIF($C$${CD1}:$C$${CDN},"Complete")`],
    ['Follow-Up Needed', `=COUNTIF($C$${CD1}:$C$${CDN},"Follow-Up Needed")`],
  ]
});

async function main() {
  await batchUpdate(id, fmt, 'annual-review-fmt');
  await valuesBatchUpdate(id, vals, 'annual-review-vals');
  await batchUpdate(id, charts, 'annual-review-charts');
  console.log('16-annualreview.js complete');
}

main().catch(e => { console.error(e); process.exit(1); });
