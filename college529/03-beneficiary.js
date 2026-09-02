'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Beneficiary Setup'];
const S   = "'Beneficiary Setup'";
const REF = "'Reference Data'";
const NC  = 18; // cols A-R

// ── Layout ───────────────────────────────────────────────────────────────────
// Row 1:   Title (merged A1:R1)
// Row 2:   Subtitle
// Rows 3-4: Summary cards (8 cards × 2 cols)
// Row 5:   Control labels
// Row 6:   Control values  ← controls referenced by other tabs
// Row 7:   Column headers  ← FROZEN ROW
// Rows 8-507: Beneficiary data
// Row 510: Contributor table header
// Rows 511-520: Contributor data

// Controls cell references used by other tabs:
//   Growth Scenario → B6!J6  wait, controls in row 6
//   let me put them more carefully in row 6:
//   A6: "Beneficiary" B6: [control]
//   C6: "Account" D6: [control]
//   E6: "College" F6: [control]
//   G6: "Report Year" H6: [value=2026]
//   I6: "Scenario" J6: [dropdown=Base]  ← 'Beneficiary Setup'!$J$6
//   K6: "Planning Date" L6: =TODAY()
//   M6: "Paycheck Freq" N6: [dropdown=Monthly]
//   O6: "Contribs/Yr" P6: [value=12]

const COL_WIDTHS = [85,180,110,90,70,90,110,90,80,150,200,110,120,110,110,110,120,200];

const BENEFICIARIES = [
  // [name, role, dob, startYear, attendYrs, collegeType, school, costGoal, aid, status, notes]
  ['Emma Hartley',     'Parent',      '2016-03-15', 2034, 4, 'In-State Public',     'Maple Ridge State University', 95000,  10000, 'Saving',          'Enrolled in automatic monthly savings plan'],
  ['Lucas Hartley',    'Parent',      '2019-07-22', 2037, 4, 'Private Nonprofit',   'Crestwood University',         220000, 25000, 'Planning',         'Early stage — reviewing school options'],
  ['Sofia Delgado',    'Parent',      '2008-11-08', 2026, 4, 'Out-of-State Public', 'Fairview University',          140000, 20000, 'Near Enrollment',  'Starting fall 2026; final contributions underway'],
  ['Marcus Washington','Parent',      '2006-05-14', 2024, 4, 'In-State Public',     'Maple Ridge State University', 85000,  15000, 'Enrolled',         'Enrolled fall 2024; withdrawals started'],
  ['Claire Beaumont',  'Grandparent', '2012-09-20', 2030, 4, 'Private Nonprofit',   'Heritage University',          180000, 30000, 'Saving',           'Grandparent-managed account; strong growth trajectory'],
];

const CONTRIBUTORS = [
  // [name, role, primaryBen, frequency, typicalContrib, active, notes]
  ['Rachel Hartley',   'Parent',      'BEN-001', 'Monthly',             300,  true,  'Automated payroll deduction'],
  ['David Hartley',    'Parent',      'BEN-001', 'Monthly',             300,  true,  'Automated payroll deduction'],
  ['Margaret Chen',    'Grandparent', 'BEN-001', 'Quarterly',           500,  true,  'Gift contributions to Emma'],
  ['Carlos Delgado',   'Parent',      'BEN-003', 'Monthly',             400,  true,  ''],
  ['Maria Delgado',    'Parent',      'BEN-003', 'Biweekly',            200,  true,  ''],
  ['James Washington', 'Parent',      'BEN-004', 'Monthly',             350,  false, 'Paused 2024 — enrolled'],
  ['Patricia Washington','Parent',    'BEN-004', 'Monthly',             250,  true,  ''],
  ['Rose Beaumont',    'Grandparent', 'BEN-005', 'Annual',             5000,  true,  'Holiday + birthday gifts combined'],
  ['Thomas Beaumont',  'Parent',      'BEN-005', 'Monthly',             350,  true,  ''],
];

(async () => {
  const vals = [];
  const fmt  = [];

  // Column widths
  COL_WIDTHS.forEach((px,ci) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Background wash
  fmt.push({ repeatCell: { range: gridRange(SID,0,600,0,NC), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  vals.push({ range: `${S}!A1`, values: [['BENEFICIARY SETUP & FAMILY SAVINGS OVERVIEW']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,NC), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,NC), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 42 }, fields: 'pixelSize' }});

  // ── Row 2: Subtitle ───────────────────────────────────────────────────────
  vals.push({ range: `${S}!A2`, values: [['Track beneficiaries, college cost goals, savings progress, and funding gaps. Yellow cells = editable. Blue-gray cells = formulas.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,NC), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,NC), cell: { userEnteredFormat: {
    backgroundColor: hex(C.aubergTint), textFormat: { fontSize: 9, foregroundColor: hex(C.text), italic: true, fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // ── Rows 3-4: Summary Cards ───────────────────────────────────────────────
  const CARDS = [
    { label: 'Active Beneficiaries',       formula: `=IFERROR(SUMPRODUCT(($B$8:$B$507<>"")*($Q$8:$Q$507<>"Archived")*($Q$8:$Q$507<>"Graduated")),0)`, fmt: '0' },
    { label: 'Total College Cost Goal',    formula: `=IFERROR(SUM($L$8:$L$507),"—")`, fmt: '"$"#,##0' },
    { label: 'Current Saved',              formula: `=IFERROR(SUM($O$8:$O$507),"—")`, fmt: '"$"#,##0' },
    { label: 'Total Funding Gap',          formula: `=IFERROR(SUM($P$8:$P$507),"—")`, fmt: '"$"#,##0' },
    { label: 'Avg Yrs Until Enrollment',   formula: `=IFERROR(AVERAGEIF($H$8:$H$507,">"&0),"—")`, fmt: '0.0' },
    { label: 'Accounts Tracked',           formula: `=IFERROR(COUNTA('529 Accounts'!$A$8:$A$1007),"—")`, fmt: '0' },
    { label: 'Contributions YTD',          formula: `=IFERROR(SUMPRODUCT((YEAR('Contribution Log'!$B$8:$B$5007)=YEAR(TODAY()))*('Contribution Log'!$I$8:$I$5007="Contribution")*'Contribution Log'!$L$8:$L$5007),"—")`, fmt: '"$"#,##0' },
    { label: 'Goals Reached',              formula: `=IFERROR(COUNTIF('Goals & Milestones'!$L$8:$L$1507,"Reached"),0)`, fmt: '0' },
  ];
  // 2 cols per card = 16 cols, plus empty Q3:R3
  CARDS.forEach((card, ci) => {
    const c1 = ci * 2, c2 = c1 + 2;
    const col = String.fromCharCode(65 + c1);
    fmt.push({ mergeCells: { range: gridRange(SID,2,3,c1,c2), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,2,3,c1,c2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.aubergTint), textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    vals.push({ range: `${S}!${col}3`, values: [[card.label]] });
    fmt.push({ mergeCells: { range: gridRange(SID,3,4,c1,c2), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,3,4,c1,c2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.white), textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      numberFormat: { type: ci === 0 || ci === 4 || ci === 5 || ci === 7 ? 'NUMBER' : 'CURRENCY', pattern: card.fmt },
    }}, fields: 'userEnteredFormat' }});
    vals.push({ range: `${S}!${col}4`, values: [[card.formula]] });
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 20 }, fields: 'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // ── Row 5: Control Labels ─────────────────────────────────────────────────
  const ctrlLabels = [
    [0,'Beneficiary / All'],[2,'Account / All'],[4,'College / All'],
    [6,'Reporting Year'],[8,'Growth Scenario'],[10,'Planning Date'],
    [12,'Paycheck Frequency'],[14,'Contributions / Year'],[16,''],
  ];
  ctrlLabels.forEach(([ci, lbl]) => {
    if (lbl) {
      vals.push({ range: `${S}!${String.fromCharCode(65+ci)}5`, values: [[lbl]] });
      fmt.push({ repeatCell: { range: gridRange(SID,4,5,ci,ci+1), cell: { userEnteredFormat: {
        backgroundColor: hex(C.aubergTint), textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.text), fontFamily: 'Arial' },
        horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
      }}, fields: 'userEnteredFormat' }});
    }
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 18 }, fields: 'pixelSize' }});

  // ── Row 6: Control Values (referenced by other tabs) ─────────────────────
  // B6=Beneficiary, D6=Account, F6=College, H6=Year, J6=Scenario, L6=Today, N6=PayFreq, P6=ContribsPerYr
  vals.push({ range: `${S}!B6`, values: [['All']] });
  vals.push({ range: `${S}!D6`, values: [['All']] });
  vals.push({ range: `${S}!F6`, values: [['All']] });
  vals.push({ range: `${S}!H6`, values: [[2026]] });
  vals.push({ range: `${S}!J6`, values: [['Base']] });
  vals.push({ range: `${S}!L6`, values: [['=TODAY()']] });
  vals.push({ range: `${S}!N6`, values: [['Monthly']] });
  vals.push({ range: `${S}!P6`, values: [[12]] });

  // Format control value cells as inputs
  [1,3,5,7,9,11,13,15].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID,5,6,ci,ci+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });
  // L6 formula cell
  fmt.push({ repeatCell: { range: gridRange(SID,5,6,11,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula), textFormat: { fontSize: 9, fontFamily: 'Arial' },
    numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // ── Row 7: Column Headers ─────────────────────────────────────────────────
  const HEADERS = [
    'Beneficiary ID','Beneficiary Name','Family Role','Birth Date','Current Age',
    'College Start Year','College Start Date','Years Until Enrollment','Planned Years of Attendance',
    'Primary College Type','Primary School Choice','College Cost Goal',
    'Aid / Scholarship Assumption','Family Savings Goal','Current Total Saved',
    'Funding Gap','Status','Notes',
  ];
  vals.push({ range: `${S}!A7`, values: [HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID,6,7,0,NC), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
    borders: { bottom: { style: 'SOLID_MEDIUM', color: hex(C.accent) } },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // ── Rows 8+: Beneficiary Data ─────────────────────────────────────────────
  BENEFICIARIES.forEach((ben, bi) => {
    const rIdx = 7 + bi;
    const rowNum = rIdx + 1;
    const isEven = bi % 2 === 0;
    const rowBg = isEven ? C.white : C.altRow;

    // A=ID (formula), B=Name, C=Role, D=DOB, E=Age(f), F=StartYear, G=StartDate(f),
    // H=YearsUntilEnroll(f), I=AttendYrs, J=CollegeType, K=School, L=CostGoal,
    // M=Aid, N=FamSavingsGoal(f), O=CurrTotalSaved(f), P=FundingGap(f), Q=Status, R=Notes

    vals.push({ range: `${S}!A${rowNum}`, values: [[
      `=IF(B${rowNum}="","","BEN-"&TEXT(ROW()-7,"000"))`,  // A
      ben[0],  // B name
      ben[1],  // C role
      ben[2],  // D dob
      `=IFERROR(DATEDIF(D${rowNum},TODAY(),"Y"),"")`,  // E age
      ben[3],  // F start year
      `=IFERROR(DATE(F${rowNum},9,1),"")`,  // G start date (Sep 1)
      `=IFERROR(MAX(0,YEARFRAC(TODAY(),G${rowNum})),"")`,  // H years until
      ben[4],  // I attend years
      ben[5],  // J college type
      ben[6],  // K school
      ben[7],  // L cost goal
      ben[8],  // M aid
      `=IFERROR(MAX(0,L${rowNum}-M${rowNum}),"")`,  // N family savings goal
      `=IFERROR(SUMIF('529 Accounts'!$B$8:$B$1007,$A${rowNum},'529 Accounts'!$N$8:$N$1007),0)`,  // O current saved
      `=IFERROR(MAX(0,L${rowNum}-M${rowNum}-O${rowNum}),0)`,  // P funding gap
      ben[9],  // Q status
      ben[10], // R notes
    ]] });

    // Row background
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,0,NC), cell: { userEnteredFormat: {
      backgroundColor: hex(rowBg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});

    // ID (A) — formula, small gray
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,0,1), cell: { userEnteredFormat: {
      textFormat: { fontSize: 8, foregroundColor: hex(C.secText), italic: true },
    }}, fields: 'userEnteredFormat.textFormat' }});

    // Name (B) — bold
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,1,2), cell: { userEnteredFormat: {
      textFormat: { bold: true, fontSize: 9 }, padding: { left: 4 },
    }}, fields: 'userEnteredFormat.textFormat,userEnteredFormat.padding' }});

    // DOB (D) — date format
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,3,4), cell: { userEnteredFormat: {
      numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' },
    }}, fields: 'userEnteredFormat.numberFormat' }});

    // Age (E) — formula bg
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,4,5), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula), horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat.backgroundColor,userEnteredFormat.horizontalAlignment' }});

    // Start Year (F) — input
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,5,6), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), horizontalAlignment: 'CENTER',
      numberFormat: { type: 'NUMBER', pattern: '0' },
    }}, fields: 'userEnteredFormat' }});

    // Start Date (G) — formula, date
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,6,7), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula), numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' }, horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat' }});

    // Years Until Enrollment (H) — formula
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,7,8), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula), numberFormat: { type: 'NUMBER', pattern: '0.0' }, horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat' }});

    // Attend Years (I) — input, center
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,8,9), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat' }});

    // Currency cols: L(11), M(12), N(13), O(14), P(15)
    [11,12].forEach(ci => {
      fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,ci,ci+1), cell: { userEnteredFormat: {
        backgroundColor: hex(C.input), numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' }, horizontalAlignment: 'RIGHT',
      }}, fields: 'userEnteredFormat' }});
    });
    [13,14,15].forEach(ci => {
      fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,ci,ci+1), cell: { userEnteredFormat: {
        backgroundColor: hex(C.formula), numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' }, horizontalAlignment: 'RIGHT',
      }}, fields: 'userEnteredFormat' }});
    });

    // Row border
    fmt.push({ updateBorders: { range: gridRange(SID,rIdx,rIdx+1,0,NC), bottom: { style: 'SOLID', color: hex(C.border) }}});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: rIdx, endIndex: rIdx+1 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  });

  // ── Contributor Table (starting row 512) ──────────────────────────────────
  const contribStart = 511; // 0-indexed
  vals.push({ range: `${S}!A512`, values: [['CONTRIBUTOR DIRECTORY']] });
  fmt.push({ mergeCells: { range: gridRange(SID,511,512,0,8), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,511,512,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 511, endIndex: 512 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  const contribHdr = ['Contributor ID','Contributor Name','Family Role','Primary Beneficiary','Contribution Frequency','Typical Contribution','Active?','Notes'];
  vals.push({ range: `${S}!A513`, values: [contribHdr] });
  fmt.push({ repeatCell: { range: gridRange(SID,512,513,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 512, endIndex: 513 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  CONTRIBUTORS.forEach((con, ci) => {
    const rIdx = 513 + ci;
    const rowNum = rIdx + 1;
    const isEven = ci % 2 === 0;

    vals.push({ range: `${S}!A${rowNum}`, values: [[
      `=IF(B${rowNum}="","","CON-"&TEXT(ROW()-513,"000"))`,
      con[0], con[1], con[2], con[3], con[4], con[5], con[6],
    ]] });

    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,0,8), cell: { userEnteredFormat: {
      backgroundColor: hex(isEven ? C.white : C.altRow), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});

    // Typical Contribution — currency
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,5,6), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' }, horizontalAlignment: 'RIGHT',
    }}, fields: 'userEnteredFormat' }});

    // Active? col (index 6) — will be set to checkbox by validation script
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,6,7), cell: { userEnteredFormat: {
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat.horizontalAlignment' }});

    fmt.push({ updateBorders: { range: gridRange(SID,rIdx,rIdx+1,0,8), bottom: { style: 'SOLID', color: hex(C.border) }}});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: rIdx, endIndex: rIdx+1 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  });

  await valuesBatchUpdate(id, vals, '03-beneficiary values');
  await batchUpdate(id, fmt, '03-beneficiary format');
  console.log('✅  Beneficiary Setup done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
