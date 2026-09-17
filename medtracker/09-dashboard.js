'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard'];

// Layout:
// Row 0: Title banner
// Row 1: Subtitle
// Row 2: Input row (Plan Year Start=B3, End=D3, OOP Max=F3, Participant=H3)
// Rows 3-4: KPI card labels + values (5 cards, 2 cols each = A-J)
// Row 5: OOP progress bar (G5:H5 merged)
// Row 6: Disclaimer
// Row 7: spacer
// Row 8: "Expenses by Category" section header
// Rows 9-21: category chart data (13 rows)
// Row 22: spacer
// Row 23: "Reimbursement Status" section header
// Rows 24-29: status chart data (6 statuses)
// Row 30: spacer
// Row 31: "Monthly Totals" section header
// Rows 32-43: monthly data (12 months)
// Row 44: spacer
// Row 45: "By Participant" section header
// Rows 46-49: participant data (4 participants)

const EXP = "'🧾 Expense Log'";
const YEAR = 2024; // default plan year

const CATEGORIES = [
  'Doctor Visits','Prescription Medications','Dental Care','Vision Care',
  'Emergency Services','Hospital Services','Therapy/Mental Health',
  'Medical Equipment','Lab/Diagnostic Tests','Preventive Care','Other',
];
const STATUSES = ['Not Submitted','Submitted','Processing','Approved','Denied','Paid'];
const PARTICIPANTS = ['Self','Spouse','Child 1','Child 2'];
const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];

// All SUMIFS use B3/D3 (date range) and H3 (participant filter)
function spentFormula() {
  return `=IFERROR(IF(H3="All",SUMIFS(${EXP}!E:E,${EXP}!A:A,">="&B3,${EXP}!A:A,"<="&D3),SUMIFS(${EXP}!E:E,${EXP}!A:A,">="&B3,${EXP}!A:A,"<="&D3,${EXP}!B:B,H3)),"")`;
}
function reimbFormula() {
  return `=IFERROR(IF(H3="All",SUMIFS(${EXP}!E:E,${EXP}!A:A,">="&B3,${EXP}!A:A,"<="&D3,${EXP}!H:H,"Paid"),SUMIFS(${EXP}!E:E,${EXP}!A:A,">="&B3,${EXP}!A:A,"<="&D3,${EXP}!H:H,"Paid",${EXP}!B:B,H3)),"")`;
}

(async () => {
  const reqs = [];

  // Row 0: Title banner
  reqs.push({ mergeCells: { range: gridRange(DSH, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 0, 1, 0, 10),
      cell: {
        userEnteredValue: { stringValue: '🏥 ULTIMATE MEDICAL EXPENSE TRACKER' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepTeal),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});

  // Row 1: Subtitle
  reqs.push({ mergeCells: { range: gridRange(DSH, 1, 2, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 1, 2, 0, 10),
      cell: {
        userEnteredValue: { stringValue: 'Track every medical expense · Maximize HSA/FSA reimbursements · Stay ahead of your out-of-pocket maximum' },
        userEnteredFormat: {
          backgroundColor: hex(C.darkText),
          textFormat: { foregroundColor: hex(C.warmCoral), italic: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Row 2: Input labels + fields
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 2, 3, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.lightSage),
          textFormat: { fontSize: 10 }, verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  // KPI cards: 5 cards × 2 cols = cols 0-9, rows 3-4 (labels + values)
  const kpis = [
    { icon: '💰', label: 'Total Amount Spent',        format: 'CURRENCY', pattern: '$#,##0.00', color: C.deepTeal },
    { icon: '✅', label: 'Total Reimbursed',          format: 'CURRENCY', pattern: '$#,##0.00', color: C.mutedSage },
    { icon: '🏥', label: 'Out-of-Pocket Cost',        format: 'CURRENCY', pattern: '$#,##0.00', color: C.warmCoral },
    { icon: '📊', label: 'OOP Max Progress',          format: 'PERCENT',  pattern: '0.0%',       color: C.amber },
    { icon: '📅', label: 'Avg Monthly Spending',      format: 'CURRENCY', pattern: '$#,##0.00', color: C.darkText },
  ];

  for (let k = 0; k < kpis.length; k++) {
    const c1 = k * 2;
    const c2 = c1 + 2;
    const kpi = kpis[k];

    reqs.push({ mergeCells: { range: gridRange(DSH, 3, 4, c1, c2), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, 3, 4, c1, c2),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(kpi.color),
            textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat',
      },
    });

    reqs.push({ mergeCells: { range: gridRange(DSH, 4, 5, c1, c2), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, 4, 5, c1, c2),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.softIvory),
            textFormat: { foregroundColor: hex(kpi.color), bold: true, fontSize: 16 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
            numberFormat: { type: kpi.format, pattern: kpi.pattern },
          },
        },
        fields: 'userEnteredFormat',
      },
    });
  }

  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});
  for (let r = 4; r < 6; r++) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
  }

  // Row 5: OOP progress bar (G5:H5)
  reqs.push({ mergeCells: { range: gridRange(DSH, 5, 6, 6, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 5, 6, 6, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.lightAmber),
          textFormat: { foregroundColor: hex(C.amber), bold: false, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 5, endIndex: 6 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Row 6: Disclaimer
  reqs.push({ mergeCells: { range: gridRange(DSH, 6, 7, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 6, 7, 0, 10),
      cell: {
        userEnteredValue: { stringValue: '⚠️ This tracker is for organizational purposes only and does not constitute tax, legal, or insurance advice. Confirm HSA/FSA eligibility, contribution limits, and reimbursement rules with your plan administrator, insurer, or a tax professional.' },
        userEnteredFormat: {
          backgroundColor: hex(C.lightGray),
          textFormat: { foregroundColor: hex(C.gray), italic: true, fontSize: 9 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 6, endIndex: 7 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});

  // Spacer row 7
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 7, endIndex: 8 },
    properties: { pixelSize: 14 }, fields: 'pixelSize',
  }});

  // Section headers helper
  const sectionHdr = (r, text) => [
    { mergeCells: { range: gridRange(DSH, r, r+1, 0, 10), mergeType: 'MERGE_ALL' } },
    { repeatCell: {
      range: gridRange(DSH, r, r+1, 0, 10),
      cell: {
        userEnteredValue: { stringValue: text },
        userEnteredFormat: {
          backgroundColor: hex(C.darkText),
          textFormat: { foregroundColor: hex(C.warmCoral), bold: true, fontSize: 11 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    }},
    { updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 30 }, fields: 'pixelSize',
    }},
  ];

  // Sections
  reqs.push(...sectionHdr(8, '🗂 Expenses by Category'));
  reqs.push(...sectionHdr(23, '📋 Reimbursement Status'));
  reqs.push(...sectionHdr(31, '📅 Monthly Totals (2024)'));
  reqs.push(...sectionHdr(45, '👥 By Participant'));

  // Chart data rows backgrounds
  const dataRowBg = (r1, r2) => {
    const arr = [];
    for (let r = r1; r < r2; r++) {
      arr.push({ repeatCell: {
        range: gridRange(DSH, r, r+1, 0, 3),
        cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.inputBg : C.altRow), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      }});
      arr.push({ updateDimensionProperties: {
        range: { sheetId: DSH, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
        properties: { pixelSize: 22 }, fields: 'pixelSize',
      }});
    }
    return arr;
  };

  reqs.push(...dataRowBg(9, 22));   // Category
  reqs.push(...dataRowBg(24, 30));  // Status
  reqs.push(...dataRowBg(32, 44));  // Monthly
  reqs.push(...dataRowBg(46, 50));  // Participant

  // Formula bg on value col B for chart data
  for (const [r1, r2] of [[9,22],[24,30],[32,44],[46,50]]) {
    reqs.push({ repeatCell: {
      range: gridRange(DSH, r1, r2, 1, 3),
      cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
      fields: 'userEnteredFormat',
    }});
  }

  // Status count col uses NUMBER
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 24, 30, 2, 3),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});

  // Spacers
  for (const r of [22, 30, 44]) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 12 }, fields: 'pixelSize',
    }});
  }

  // Col widths
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 160 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 140 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 90 }, fields: 'pixelSize',
  }});
  for (let c = 3; c < 10; c++) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: c, endIndex: c+1 },
      properties: { pixelSize: 110 }, fields: 'pixelSize',
    }});
  }

  await batchUpdate(id, reqs, 'dashboard-format');

  // Values
  const data = [];

  // Row 3: Input labels
  data.push({ range: `'📊 Dashboard'!A3`, values: [['Plan Year Start:']] });
  data.push({ range: `'📊 Dashboard'!B3`, values: [['2024-01-01']] });
  data.push({ range: `'📊 Dashboard'!C3`, values: [['Plan Year End:']] });
  data.push({ range: `'📊 Dashboard'!D3`, values: [['2024-12-31']] });
  data.push({ range: `'📊 Dashboard'!E3`, values: [['OOP Maximum:']] });
  data.push({ range: `'📊 Dashboard'!F3`, values: [[9000]] });
  data.push({ range: `'📊 Dashboard'!G3`, values: [['Participant:']] });
  data.push({ range: `'📊 Dashboard'!H3`, values: [['All']] });

  // Row 4: KPI labels
  const kpiLabels = kpis.map((k, i) => {
    const c = String.fromCharCode(65 + i*2); // A, C, E, G, I
    return { range: `'📊 Dashboard'!${c}4`, values: [[`${k.icon} ${k.label}`]] };
  });
  data.push(...kpiLabels);

  // Row 5: KPI values
  data.push({ range: `'📊 Dashboard'!A5`, values: [[spentFormula()]] });
  data.push({ range: `'📊 Dashboard'!C5`, values: [[reimbFormula()]] });
  data.push({ range: `'📊 Dashboard'!E5`, values: [[`=IFERROR(A5-C5,"")`]] });
  data.push({ range: `'📊 Dashboard'!G5`, values: [[`=IFERROR(IF(F3=0,"",E5/F3),"")`]] });
  data.push({ range: `'📊 Dashboard'!I5`, values: [[`=IFERROR(A5/12,"")`]] });

  // Row 6: OOP progress bar
  data.push({
    range: `'📊 Dashboard'!G6`,
    values: [[`=IFERROR("OOP Progress: "&REPT("█",MIN(ROUND(E5/F3*20,0),20))&REPT("░",MAX(20-ROUND(E5/F3*20,0),0))&"  "&TEXT(G5,"0%"),"░░░░░░░░░░░░░░░░░░░░")`]],
  });

  // Category chart data rows 10-22
  for (let i = 0; i < CATEGORIES.length; i++) {
    const row = i + 10;
    const cat = CATEGORIES[i];
    const formula = `=IFERROR(IF(H3="All",SUMIFS(${EXP}!E:E,${EXP}!A:A,">="&B3,${EXP}!A:A,"<="&D3,${EXP}!C:C,"${cat}"),SUMIFS(${EXP}!E:E,${EXP}!A:A,">="&B3,${EXP}!A:A,"<="&D3,${EXP}!C:C,"${cat}",${EXP}!B:B,H3)),"")`;
    data.push({ range: `'📊 Dashboard'!A${row}:B${row}`, values: [[cat, formula]] });
  }

  // Reimbursement status chart data rows 25-30
  for (let i = 0; i < STATUSES.length; i++) {
    const row = i + 25;
    const status = STATUSES[i];
    const amtFormula = `=IFERROR(IF(H3="All",SUMIFS(${EXP}!E:E,${EXP}!A:A,">="&B3,${EXP}!A:A,"<="&D3,${EXP}!H:H,"${status}"),SUMIFS(${EXP}!E:E,${EXP}!A:A,">="&B3,${EXP}!A:A,"<="&D3,${EXP}!H:H,"${status}",${EXP}!B:B,H3)),"")`;
    const cntFormula = `=IFERROR(IF(H3="All",COUNTIFS(${EXP}!A:A,">="&B3,${EXP}!A:A,"<="&D3,${EXP}!H:H,"${status}"),COUNTIFS(${EXP}!A:A,">="&B3,${EXP}!A:A,"<="&D3,${EXP}!H:H,"${status}",${EXP}!B:B,H3)),"")`;
    data.push({ range: `'📊 Dashboard'!A${row}:C${row}`, values: [[status, amtFormula, cntFormula]] });
  }

  // Monthly totals rows 33-44
  for (let m = 0; m < 12; m++) {
    const row = m + 33;
    const monthNum = m + 1;
    const formula = `=IFERROR(IF(H3="All",SUMPRODUCT((YEAR(${EXP}!$A$3:$A$200)=YEAR(B3))*(MONTH(${EXP}!$A$3:$A$200)=${monthNum})*(${EXP}!$E$3:$E$200)),SUMPRODUCT((YEAR(${EXP}!$A$3:$A$200)=YEAR(B3))*(MONTH(${EXP}!$A$3:$A$200)=${monthNum})*(${EXP}!$B$3:$B$200=H3)*(${EXP}!$E$3:$E$200))),"")`;
    data.push({ range: `'📊 Dashboard'!A${row}:B${row}`, values: [[MONTHS[m], formula]] });
  }

  // Participant data rows 47-50
  for (let p = 0; p < PARTICIPANTS.length; p++) {
    const row = p + 47;
    const part = PARTICIPANTS[p];
    const formula = `=IFERROR(SUMIFS(${EXP}!E:E,${EXP}!A:A,">="&B3,${EXP}!A:A,"<="&D3,${EXP}!B:B,"${part}"),"")`;
    data.push({ range: `'📊 Dashboard'!A${row}:B${row}`, values: [[part, formula]] });
  }

  await valuesBatchUpdate(id, data, 'dashboard-values');

  // Format input date cells
  const fmtReqs = [];
  fmtReqs.push({ repeatCell: {
    range: gridRange(DSH, 2, 3, 1, 2),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  fmtReqs.push({ repeatCell: {
    range: gridRange(DSH, 2, 3, 3, 4),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  fmtReqs.push({ repeatCell: {
    range: gridRange(DSH, 2, 3, 5, 6),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  // Input cells: coral border
  const cb = { style: 'SOLID', color: hex(C.border) };
  for (const col of [1,3,5,7]) {
    fmtReqs.push({ repeatCell: {
      range: gridRange(DSH, 2, 3, col, col+1),
      cell: { userEnteredFormat: { borders: { bottom: cb, top: cb, left: cb, right: cb }, backgroundColor: hex(C.inputBg) } },
      fields: 'userEnteredFormat.borders,userEnteredFormat.backgroundColor',
    }});
  }
  await batchUpdate(id, fmtReqs, 'dashboard-input-fmt');

  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
