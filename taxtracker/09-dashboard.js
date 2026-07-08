'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard'];

// Dashboard layout:
// Rows 0-1: Title + subtitle
// Row 2: Section header "2024 Tax Year At a Glance"
// Rows 3-7: KPI cards (4 wide)
// Rows 8-9: Spacer + section header "Deductions by Category"
// Rows 10-22: Chart data (category + amount) — 13 categories
// Row 23: Spacer
// Row 24: Section "Documentation Status"
// Rows 25-26: Doc data (Documented / Not Documented)
// Row 27: Spacer
// Row 28: Section "Monthly Deduction Totals"
// Rows 29-40: Monthly data (12 months)
// Row 41: Spacer
// Row 42: Section "Multi-Year Comparison"
// Rows 43-45: Year data (2024/2025/2026)

const LOG = "'📝 Deduction Log'";
const MIL = "'🚗 Mileage Log'";
const HOM = "'🏠 Home Office'";
const QTR = "'📅 Quarterly Tax'";
const CHK = "'✅ Doc Checklist'";
const MYR = "'📈 Multi-Year'";

const YEAR = 2024;

(async () => {
  const reqs = [];

  // Row 0: Title
  reqs.push({ mergeCells: { range: gridRange(DSH, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 0, 1, 0, 10),
      cell: {
        userEnteredValue: { stringValue: '💼 Ultimate Tax Deduction Tracker' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepSapphire),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 48 }, fields: 'pixelSize',
  }});

  // Row 1: Subtitle
  reqs.push({ mergeCells: { range: gridRange(DSH, 1, 2, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 1, 2, 0, 10),
      cell: {
        userEnteredValue: { stringValue: 'Track every deduction. Maximize every dollar. Never miss a write-off again.' },
        userEnteredFormat: {
          backgroundColor: hex(C.charcoal),
          textFormat: { foregroundColor: hex(C.warmGold), italic: true, fontSize: 10 },
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

  // Row 2: Section header
  reqs.push({ mergeCells: { range: gridRange(DSH, 2, 3, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 2, 3, 0, 10),
      cell: {
        userEnteredValue: { stringValue: `📊 ${YEAR} Tax Year At a Glance` },
        userEnteredFormat: {
          backgroundColor: hex(C.warmGold),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});

  // KPI Cards — 4 cards: rows 3-7 (merged 2-tall each, cols 0-2 each)
  // Card 1: Total Deductions (0-2)
  // Card 2: Mileage Deductible (3-5)  [actually cols 3-5 wait, 10 cols total, 4 cards = col 0-2, 2-4, 5-7, 7-9]
  // Let's do cols 0-2, 2-4, 5-7, 7-9 → nah, let's do 4 cards at cols 0-2, 2-4... better: cols 0,2,5,7 with width 2
  // Actually let's do uniform: 4 cards at (0-2), (2-4), (5-7), (7-9) — wait those overlap. Let me do:
  // Cols 0-2, 3-5, 5-7, 7-9? No. Let's just do 5 KPIs in cols 0-1, 2-3, 4-5, 6-7, 8-9

  const kpiCards = [
    { label: 'Total Deductions',     formula: `=IFERROR(SUMPRODUCT((YEAR(${LOG}!$A$3:$A$200)=${YEAR})*(${LOG}!$G$3:$G$200)),0)`,                     format: '$#,##0.00', icon: '💰', color: C.deepSapphire },
    { label: 'Mileage Deductible',   formula: `=IFERROR(SUMPRODUCT((YEAR(${MIL}!$A$3:$A$200)=${YEAR})*(${MIL}!$F$3:$F$200)),0)`,                     format: '$#,##0.00', icon: '🚗', color: C.mutedSage },
    { label: 'Home Office Deduction',formula: `=IFERROR(MAX(${HOM}!B25,0),0)`,                                                                       format: '$#,##0.00', icon: '🏠', color: C.warmGold },
    { label: 'Quarterly Tax Due',    formula: `=IFERROR(SUM(${QTR}!H4:H7),0)`,                                                                       format: '$#,##0.00', icon: '📅', color: C.amber },
    { label: 'Documentation Rate',   formula: `=IFERROR(SUMPRODUCT((YEAR(${LOG}!$A$3:$A$200)=${YEAR})*(${LOG}!$F$3:$F$200=TRUE))/SUMPRODUCT((YEAR(${LOG}!$A$3:$A$200)=${YEAR})*(${LOG}!$A$3:$A$200<>"")),0)`, format: '0.0%', icon: '✅', color: C.mutedSage },
  ];

  for (let k = 0; k < kpiCards.length; k++) {
    const c1 = k * 2;
    const c2 = c1 + 2;
    const card = kpiCards[k];

    // Label row (row 3)
    reqs.push({ mergeCells: { range: gridRange(DSH, 3, 4, c1, c2), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, 3, 4, c1, c2),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(card.color),
            textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat',
      },
    });

    // Value row (rows 4-6)
    reqs.push({ mergeCells: { range: gridRange(DSH, 4, 7, c1, c2), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, 4, 7, c1, c2),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.warmIvory),
            textFormat: { foregroundColor: hex(card.color), bold: true, fontSize: 16 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
            numberFormat: { type: card.format === '0.0%' ? 'PERCENT' : 'CURRENCY', pattern: card.format },
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
  for (let r = 4; r < 7; r++) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 22 }, fields: 'pixelSize',
    }});
  }

  // Spacer row 7
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 7, endIndex: 8 },
    properties: { pixelSize: 14 }, fields: 'pixelSize',
  }});

  // Section: Deductions by Category (row 8)
  const sectionHeader = (r, text) => [
    { mergeCells: { range: gridRange(DSH, r, r+1, 0, 10), mergeType: 'MERGE_ALL' } },
    { repeatCell: {
      range: gridRange(DSH, r, r+1, 0, 10),
      cell: {
        userEnteredValue: { stringValue: text },
        userEnteredFormat: {
          backgroundColor: hex(C.charcoal),
          textFormat: { foregroundColor: hex(C.warmGold), bold: true, fontSize: 11 },
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

  reqs.push(...sectionHeader(8, '🗂 Deductions by Category'));

  // Category data rows 9-21 (13 categories from Reference)
  const categories = [
    'Office Expenses','Meals & Entertainment','Travel','Vehicle/Mileage','Home Office',
    'Professional Services','Marketing & Advertising','Insurance','Education & Training',
    'Equipment & Technology','Subscriptions & Software','Rent & Utilities','Other',
  ];

  for (let r = 0; r < 13; r++) {
    const row = r + 9;
    const bg = r % 2 === 0 ? C.inputBg : C.altRow;
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, row, row+1, 0, 3),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'ROWS', startIndex: row, endIndex: row+1 },
      properties: { pixelSize: 22 }, fields: 'pixelSize',
    }});
  }

  // Col B formula bg
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 9, 22, 1, 3),
    cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
    fields: 'userEnteredFormat',
  }});

  // Spacer row 22
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 22, endIndex: 23 },
    properties: { pixelSize: 14 }, fields: 'pixelSize',
  }});

  // Section: Documentation Status (row 23)
  reqs.push(...sectionHeader(23, '📋 Documentation Status'));

  // Doc data rows 24-25
  for (let r = 24; r < 26; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, r, r+1, 0, 3),
        cell: { userEnteredFormat: { backgroundColor: hex(r === 24 ? C.lightSage : C.lightRed), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 26 }, fields: 'pixelSize',
    }});
  }
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 24, 26, 1, 3),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' }, backgroundColor: hex(C.formulaBg) } },
    fields: 'userEnteredFormat',
  }});

  // Spacer 26
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 26, endIndex: 27 },
    properties: { pixelSize: 14 }, fields: 'pixelSize',
  }});

  // Section: Monthly Deduction Totals (row 27)
  reqs.push(...sectionHeader(27, '📅 Monthly Deduction Totals (2024)'));

  // Monthly data rows 28-39
  for (let r = 28; r < 40; r++) {
    const bg = r % 2 === 0 ? C.inputBg : C.altRow;
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, r, r+1, 0, 3),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 22 }, fields: 'pixelSize',
    }});
  }
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 28, 40, 1, 3),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' }, backgroundColor: hex(C.formulaBg) } },
    fields: 'userEnteredFormat',
  }});

  // Spacer 40
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 40, endIndex: 41 },
    properties: { pixelSize: 14 }, fields: 'pixelSize',
  }});

  // Section: Multi-Year (row 41)
  reqs.push(...sectionHeader(41, '📈 Multi-Year Deduction Comparison'));

  // Multi-year rows 42-44
  for (let r = 42; r < 45; r++) {
    const bg = r % 2 === 0 ? C.inputBg : C.altRow;
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, r, r+1, 0, 3),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 26 }, fields: 'pixelSize',
    }});
  }
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 42, 45, 1, 3),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' }, backgroundColor: hex(C.formulaBg) } },
    fields: 'userEnteredFormat',
  }});

  // Col widths: A=160, B=140, C=100, D-J=90 each
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 160 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 140 }, fields: 'pixelSize',
  }});
  for (let c = 2; c < 10; c++) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: c, endIndex: c+1 },
      properties: { pixelSize: 100 }, fields: 'pixelSize',
    }});
  }

  await batchUpdate(id, reqs, 'dashboard-format');

  // Values
  const data = [];

  // KPI labels + values
  for (let k = 0; k < kpiCards.length; k++) {
    const c1 = k * 2;
    const colA = String.fromCharCode(65 + c1);
    data.push({ range: `'📊 Dashboard'!${colA}4`, values: [[`${kpiCards[k].icon} ${kpiCards[k].label}`]] });
    data.push({ range: `'📊 Dashboard'!${colA}5`, values: [[kpiCards[k].formula]] });
  }

  // Category data (rows 10-22, cols A-B)
  const MONTHS_NUM = ['1','2','3','4','5','6','7','8','9','10','11','12'];
  for (let r = 0; r < categories.length; r++) {
    const row = r + 10;
    const cat = categories[r];
    const catFormula = `=IFERROR(SUMPRODUCT((YEAR(${LOG}!$A$3:$A$200)=${YEAR})*(${LOG}!$B$3:$B$200="${cat}")*(${LOG}!$G$3:$G$200)),0)`;
    data.push({ range: `'📊 Dashboard'!A${row}:B${row}`, values: [[cat, catFormula]] });
  }

  // Documentation status (rows 25-26)
  const docDocFormula = `=IFERROR(SUMPRODUCT((YEAR(${LOG}!$A$3:$A$200)=${YEAR})*(${LOG}!$F$3:$F$200=TRUE)),0)`;
  const docUndocFormula = `=IFERROR(SUMPRODUCT((YEAR(${LOG}!$A$3:$A$200)=${YEAR})*(${LOG}!$F$3:$F$200=FALSE)*(${LOG}!$A$3:$A$200<>"")),0)`;
  data.push({ range: `'📊 Dashboard'!A25:B25`, values: [['Documented', docDocFormula]] });
  data.push({ range: `'📊 Dashboard'!A26:B26`, values: [['Not Documented', docUndocFormula]] });

  // Monthly data (rows 29-40)
  const MONTH_NAMES = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  for (let m = 0; m < 12; m++) {
    const row = m + 29;
    const monthNum = m + 1;
    const monthFormula = `=IFERROR(SUMPRODUCT((YEAR(${LOG}!$A$3:$A$200)=${YEAR})*(MONTH(${LOG}!$A$3:$A$200)=${monthNum})*(${LOG}!$G$3:$G$200)),0)`;
    data.push({ range: `'📊 Dashboard'!A${row}:B${row}`, values: [[MONTH_NAMES[m], monthFormula]] });
  }

  // Multi-year data (rows 43-45)
  for (let y = 0; y < 3; y++) {
    const year = 2024 + y;
    const row = y + 43;
    data.push({ range: `'📊 Dashboard'!A${row}:B${row}`, values: [[year, `=${MYR}!B${y+3}`]] });
  }

  await valuesBatchUpdate(id, data, 'dashboard-values');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
