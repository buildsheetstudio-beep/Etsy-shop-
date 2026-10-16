'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DB  = sheetMap['🛡️ Security Dashboard'];
const AV  = sheetMap['🔐 Account Vault'];

// 16 cols: A-P
// Dashboard layout:
// Row 1: Title
// Row 2: Subtitle/tagline
// Row 3-4: Spacer
// Row 5-12: KPI Block (2x4 grid)
// Row 13: Spacer
// Row 14-15: Section header "Security Score"
// Row 16: Score bar
// Row 17: Spacer
// Row 18-19: Section header "Accounts Needing Attention"
// Row 20-29: FILTER table (accounts with Weak/Critical strength or no MFA)
// Row 30: Spacer
// Row 31-32: Section header "Password Strength Summary"
// Row 33-37: Strength breakdown
// Row 38: Spacer
// Row 39-40: Section header "Account Status Overview"
// Row 41-45: Status breakdown
// Row 46: Spacer
// Row 47-48: Section header "Category Breakdown"
// Row 49-62: Category counts
// Row 63+: Spacer

const NCOLS = 16; // A-P

(async () => {
  const reqs = [];

  // ── Column widths ─────────────────────────────────────────────────────
  const COL_WIDTHS = [180,140,120,120,140,140,120,120,120,120,120,120,120,120,120,120];
  COL_WIDTHS.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DB, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w },
      fields: 'pixelSize',
    }});
  });

  // ── Row heights ───────────────────────────────────────────────────────
  const rowHeights = {
    0: 48,  // title
    1: 26,  // subtitle
    2: 12,  // spacer
    3: 12,  // spacer
    4: 28,  // KPI section header
    5: 60,  // KPI row 1
    6: 60,  // KPI row 2
    7: 12,  // spacer
    8: 26,  // score header
    9: 44,  // score bar row
    10: 12, // spacer
    11: 26, // attention header
    12: 26, // attention col headers
    13: 22,14: 22,15: 22,16: 22,17: 22,18: 22,19: 22,20: 22,21: 22,22: 22, // 10 attention rows
    23: 12, // spacer
    24: 26, // strength header
    25: 22,26: 22,27: 22,28: 22, // 4 strength rows
    29: 12, // spacer
    30: 26, // status header
    31: 22,32: 22,33: 22,34: 22,35: 22, // 5 status rows
    36: 12, // spacer
    37: 26, // category header
    38: 22,39: 22,40: 22,41: 22,42: 22,43: 22,44: 22,45: 22,46: 22,47: 22,48: 22,49: 22,50: 22,51: 22, // 14 cats
  };
  Object.entries(rowHeights).forEach(([r, h]) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DB, dimension: 'ROWS', startIndex: +r, endIndex: +r + 1 },
      properties: { pixelSize: h },
      fields: 'pixelSize',
    }});
  });

  function merge(r1, r2, c1, c2) {
    reqs.push({ mergeCells: { range: gridRange(DB, r1, r2, c1, c2), mergeType: 'MERGE_ALL' }});
  }
  function bg(r1, r2, c1, c2, color) {
    reqs.push({ repeatCell: {
      range: gridRange(DB, r1, r2, c1, c2),
      cell: { userEnteredFormat: { backgroundColor: hex(color) }},
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }
  function fmt(r1, r2, c1, c2, opts) {
    reqs.push({ repeatCell: {
      range: gridRange(DB, r1, r2, c1, c2),
      cell: { userEnteredFormat: opts },
      fields: 'userEnteredFormat(' + Object.keys(opts).join(',') + ')',
    }});
  }

  // ── TITLE (row 1) ─────────────────────────────────────────────────────
  bg(0, 1, 0, NCOLS, C.deepIndigo);
  merge(0, 1, 0, NCOLS);
  fmt(0, 1, 0, NCOLS, {
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18, fontFamily: 'Montserrat' },
    verticalAlignment: 'MIDDLE',
    horizontalAlignment: 'LEFT',
  });

  // ── SUBTITLE (row 2) ──────────────────────────────────────────────────
  bg(1, 2, 0, NCOLS, C.midnightIndigo);
  merge(1, 2, 0, NCOLS);
  fmt(1, 2, 0, NCOLS, {
    textFormat: { foregroundColor: hex(C.clearBlue), fontSize: 10, italic: true },
    verticalAlignment: 'MIDDLE',
    horizontalAlignment: 'LEFT',
  });

  // ── SPACERS ───────────────────────────────────────────────────────────
  [2, 3].forEach(r => bg(r, r + 1, 0, NCOLS, C.paleIndigo));

  // ── KPI SECTION HEADER (row 5) ────────────────────────────────────────
  bg(4, 5, 0, NCOLS, C.midnightIndigo);
  merge(4, 5, 0, NCOLS);
  fmt(4, 5, 0, NCOLS, {
    textFormat: { foregroundColor: hex(C.clearBlue), bold: true, fontSize: 11 },
    verticalAlignment: 'MIDDLE',
    horizontalAlignment: 'LEFT',
  });

  // ── KPI CARDS (rows 6-7, 2 rows × 4 cols) ────────────────────────────
  // Cards: [label, value-formula, col-start, color-scheme]
  const KPI_CARDS = [
    // Row 6: cols A-D, E-H, I-L, M-P
    { label: 'Total Accounts',         formula: '=COUNTA(\'🔐 Account Vault\'!B10:B109)',      cs: 0,  ce: 4,  bg: C.deepIndigo,     fg: C.white,      vfg: C.clearBlue },
    { label: 'Active Accounts',        formula: '=COUNTIF(\'🔐 Account Vault\'!M10:M109,"Active")', cs: 4, ce: 8, bg: C.mediumIndigo, fg: C.white,      vfg: C.white },
    { label: 'Accounts with MFA',      formula: '=COUNTIF(\'🔐 Account Vault\'!G10:G109,"Yes")', cs: 8, ce: 12, bg: C.veryPaleIndigo, fg: C.bodyText, vfg: C.deepIndigo },
    { label: 'Weak / Critical Passwords', formula: '=COUNTIF(\'🔐 Account Vault\'!L10:L109,"Weak")+COUNTIF(\'🔐 Account Vault\'!L10:L109,"Critical")', cs: 12, ce: 16, bg: C.errorBg, fg: C.errorFg, vfg: C.errorFg },
    // Row 7: cols A-D, E-H, I-L, M-P
    { label: 'MFA Coverage %',         formula: '=IFERROR(COUNTIF(\'🔐 Account Vault\'!G10:G109,"Yes")/COUNTA(\'🔐 Account Vault\'!B10:B109),0)', cs: 0, ce: 4, row: 1, bg: C.paleIndigo, fg: C.bodyText, vfg: C.deepIndigo, pct: true },
    { label: 'Strong Passwords',       formula: '=COUNTIF(\'🔐 Account Vault\'!L10:L109,"Strong")', cs: 4, ce: 8, row: 1, bg: C.doneBg, fg: C.doneFg, vfg: C.deepGreen },
    { label: 'Total Monthly Spend',    formula: '=SUMPRODUCT((\'💳 Subscriptions & Billing\'!I10:I24="Active")*(\'💳 Subscriptions & Billing\'!E10:E24))', cs: 8, ce: 12, row: 1, bg: C.veryPaleIndigo, fg: C.bodyText, vfg: C.deepIndigo, currency: true },
    { label: 'Subscriptions Active',   formula: '=COUNTIF(\'💳 Subscriptions & Billing\'!I10:I24,"Active")', cs: 12, ce: 16, row: 1, bg: C.paleIndigo, fg: C.bodyText, vfg: C.deepIndigo },
  ];

  KPI_CARDS.forEach(k => {
    const row = 5 + (k.row || 0);
    bg(row, row + 1, k.cs, k.ce, k.bg);
    merge(row, row + 1, k.cs, k.ce);
    // We'll set values later; just format the box
    reqs.push({ updateBorders: {
      range: gridRange(DB, row, row + 1, k.cs, k.ce),
      top: { style: 'SOLID', color: hex(C.border), width: 1 },
      bottom: { style: 'SOLID', color: hex(C.border), width: 1 },
      left: { style: 'SOLID', color: hex(C.border), width: 1 },
      right: { style: 'SOLID', color: hex(C.border), width: 1 },
    }});
  });

  // ── SECURITY SCORE (rows 9-10) ────────────────────────────────────────
  bg(7, 8, 0, NCOLS, C.paleIndigo); // spacer
  bg(8, 9, 0, NCOLS, C.midnightIndigo);
  merge(8, 9, 0, NCOLS);
  fmt(8, 9, 0, NCOLS, {
    textFormat: { foregroundColor: hex(C.clearBlue), bold: true, fontSize: 11 },
    verticalAlignment: 'MIDDLE',
    horizontalAlignment: 'LEFT',
  });

  bg(9, 10, 0, NCOLS, C.veryPaleIndigo);
  // Score label: cols A-B
  merge(9, 10, 0, 2);
  fmt(9, 10, 0, 2, {
    textFormat: { foregroundColor: hex(C.bodyText), bold: true, fontSize: 12 },
    verticalAlignment: 'MIDDLE',
    horizontalAlignment: 'RIGHT',
  });
  // Score value: cols C-D
  merge(9, 10, 2, 4);
  fmt(9, 10, 2, 4, {
    textFormat: { foregroundColor: hex(C.deepIndigo), bold: true, fontSize: 22 },
    verticalAlignment: 'MIDDLE',
    horizontalAlignment: 'LEFT',
  });
  // Score bar: cols E-P
  merge(9, 10, 4, NCOLS);
  fmt(9, 10, 4, NCOLS, {
    textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
    verticalAlignment: 'MIDDLE',
    horizontalAlignment: 'LEFT',
  });

  // ── ACCOUNTS NEEDING ATTENTION (rows 12-23) ───────────────────────────
  bg(10, 11, 0, NCOLS, C.paleIndigo); // spacer
  bg(11, 12, 0, NCOLS, C.midnightIndigo);
  merge(11, 12, 0, NCOLS);
  fmt(11, 12, 0, NCOLS, {
    textFormat: { foregroundColor: hex(C.clearBlue), bold: true, fontSize: 11 },
    verticalAlignment: 'MIDDLE',
    horizontalAlignment: 'LEFT',
  });

  // Table header row 13
  bg(12, 13, 0, 8, C.mediumIndigo);
  fmt(12, 13, 0, 8, {
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
    horizontalAlignment: 'CENTER',
    verticalAlignment: 'MIDDLE',
  });

  // Attention table data rows 14-23
  for (let r = 13; r < 23; r++) {
    const bgColor = (r - 13) % 2 === 0 ? C.paleIndigo : C.silkyWhite;
    bg(r, r + 1, 0, 8, bgColor);
    fmt(r, r + 1, 0, 8, {
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    });
  }

  // ── PASSWORD STRENGTH SUMMARY (rows 25-29) ────────────────────────────
  bg(23, 24, 0, NCOLS, C.paleIndigo); // spacer
  bg(24, 25, 0, NCOLS, C.midnightIndigo);
  merge(24, 25, 0, NCOLS);
  fmt(24, 25, 0, NCOLS, {
    textFormat: { foregroundColor: hex(C.clearBlue), bold: true, fontSize: 11 },
    verticalAlignment: 'MIDDLE',
    horizontalAlignment: 'LEFT',
  });

  const STRENGTHS = [
    { label: 'Strong',   fgColor: C.doneFg,   bgColor: C.doneBg   },
    { label: 'Medium',   fgColor: C.warnFg,   bgColor: C.warnBg   },
    { label: 'Weak',     fgColor: C.errorFg,  bgColor: C.errorBg  },
    { label: 'Critical', fgColor: C.deepRed,  bgColor: '#FFD0D0'  },
  ];
  STRENGTHS.forEach((s, i) => {
    const r = 25 + i;
    bg(r, r + 1, 0, 2, s.bgColor);
    bg(r, r + 1, 2, 5, C.veryPaleIndigo);
    bg(r, r + 1, 5, 8, C.paleIndigo);
    fmt(r, r + 1, 0, 2, {
      textFormat: { foregroundColor: hex(s.fgColor), bold: true, fontSize: 10 },
      verticalAlignment: 'MIDDLE',
      horizontalAlignment: 'CENTER',
    });
    fmt(r, r + 1, 2, 5, {
      textFormat: { foregroundColor: hex(C.deepIndigo), bold: true, fontSize: 12 },
      verticalAlignment: 'MIDDLE',
      horizontalAlignment: 'CENTER',
    });
    fmt(r, r + 1, 5, 8, {
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    });
    reqs.push({ updateBorders: {
      range: gridRange(DB, r, r + 1, 0, 8),
      bottom: { style: 'SOLID', color: hex(C.border), width: 1 },
    }});
  });

  // ── ACCOUNT STATUS OVERVIEW (rows 31-35) ─────────────────────────────
  bg(29, 30, 0, NCOLS, C.paleIndigo); // spacer
  bg(30, 31, 0, NCOLS, C.midnightIndigo);
  merge(30, 31, 0, NCOLS);
  fmt(30, 31, 0, NCOLS, {
    textFormat: { foregroundColor: hex(C.clearBlue), bold: true, fontSize: 11 },
    verticalAlignment: 'MIDDLE',
    horizontalAlignment: 'LEFT',
  });

  const STATUSES = ['Active','Inactive','Suspended','Closed','Under Review'];
  STATUSES.forEach((s, i) => {
    const r = 31 + i;
    bg(r, r + 1, 0, 2, C.veryPaleIndigo);
    bg(r, r + 1, 2, 4, C.paleIndigo);
    fmt(r, r + 1, 0, 2, {
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 10 },
      verticalAlignment: 'MIDDLE',
      horizontalAlignment: 'LEFT',
    });
    fmt(r, r + 1, 2, 4, {
      textFormat: { foregroundColor: hex(C.deepIndigo), bold: true, fontSize: 11 },
      verticalAlignment: 'MIDDLE',
      horizontalAlignment: 'CENTER',
    });
    reqs.push({ updateBorders: {
      range: gridRange(DB, r, r + 1, 0, 4),
      bottom: { style: 'SOLID', color: hex(C.border), width: 1 },
    }});
  });

  // ── CATEGORY BREAKDOWN (rows 38-51) ──────────────────────────────────
  bg(36, 37, 0, NCOLS, C.paleIndigo); // spacer
  bg(37, 38, 0, NCOLS, C.midnightIndigo);
  merge(37, 38, 0, NCOLS);
  fmt(37, 38, 0, NCOLS, {
    textFormat: { foregroundColor: hex(C.clearBlue), bold: true, fontSize: 11 },
    verticalAlignment: 'MIDDLE',
    horizontalAlignment: 'LEFT',
  });

  const CATS = [
    'Email & Communication','Social Media','Banking & Finance','Shopping & Retail',
    'Streaming & Entertainment','Work & Productivity','Health & Wellness',
    'Government & Legal','Travel & Transport','Gaming','Cloud Storage',
    'Security & Password','Utilities & Home','Other',
  ];
  CATS.forEach((cat, i) => {
    const r = 38 + i;
    bg(r, r + 1, 0, 2, i % 2 === 0 ? C.paleIndigo : C.veryPaleIndigo);
    bg(r, r + 1, 2, 4, i % 2 === 0 ? C.veryPaleIndigo : C.paleIndigo);
    fmt(r, r + 1, 0, 2, {
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    });
    fmt(r, r + 1, 2, 4, {
      textFormat: { foregroundColor: hex(C.deepIndigo), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER',
      verticalAlignment: 'MIDDLE',
    });
    reqs.push({ updateBorders: {
      range: gridRange(DB, r, r + 1, 0, 4),
      bottom: { style: 'SOLID', color: hex(C.border), width: 1 },
    }});
  });

  // Overall borders for main sections
  [[4,7,0,NCOLS],[8,10,0,NCOLS],[11,23,0,8],[24,29,0,8],[30,36,0,4],[37,52,0,4]].forEach(([r1,r2,c1,c2]) => {
    reqs.push({ updateBorders: {
      range: gridRange(DB, r1, r2, c1, c2),
      top:    { style: 'SOLID_MEDIUM', color: hex(C.deepIndigo), width: 2 },
      bottom: { style: 'SOLID_MEDIUM', color: hex(C.deepIndigo), width: 2 },
      left:   { style: 'SOLID_MEDIUM', color: hex(C.deepIndigo), width: 2 },
      right:  { style: 'SOLID_MEDIUM', color: hex(C.deepIndigo), width: 2 },
    }});
  });

  await batchUpdate(id, reqs, 'dashboard-format');

  // ── VALUES ─────────────────────────────────────────────────────────────
  const vals = [];

  vals.push({ range: "'🛡️ Security Dashboard'!A1", values: [['  🛡️  Security Dashboard — Account & Password Health']] });
  vals.push({ range: "'🛡️ Security Dashboard'!A2", values: [['   Auto-calculated from Account Vault & Subscriptions data  ·  Last refreshed on open']] });

  vals.push({ range: "'🛡️ Security Dashboard'!A5", values: [['  📊  Key Metrics at a Glance']] });

  // KPI card labels and values (each card is a merged cell, so we write to the top-left cell)
  const KPI_VALS = [
    // row 6 (index 5)
    { range: "'🛡️ Security Dashboard'!A6", label: 'TOTAL ACCOUNTS', formula: `=COUNTA('🔐 Account Vault'!B10:B109)` },
    { range: "'🛡️ Security Dashboard'!E6", label: 'ACTIVE ACCOUNTS', formula: `=COUNTIF('🔐 Account Vault'!M10:M109,"Active")` },
    { range: "'🛡️ Security Dashboard'!I6", label: 'WITH MFA ENABLED', formula: `=COUNTIF('🔐 Account Vault'!G10:G109,"Yes")` },
    { range: "'🛡️ Security Dashboard'!M6", label: 'WEAK / CRITICAL PASSWORDS', formula: `=COUNTIF('🔐 Account Vault'!L10:L109,"Weak")+COUNTIF('🔐 Account Vault'!L10:L109,"Critical")` },
    // row 7 (index 6)
    { range: "'🛡️ Security Dashboard'!A7", label: 'MFA COVERAGE', formula: `=IFERROR(TEXT(COUNTIF('🔐 Account Vault'!G10:G109,"Yes")/COUNTA('🔐 Account Vault'!B10:B109),"0%"),"—")` },
    { range: "'🛡️ Security Dashboard'!E7", label: 'STRONG PASSWORDS', formula: `=COUNTIF('🔐 Account Vault'!L10:L109,"Strong")` },
    { range: "'🛡️ Security Dashboard'!I7", label: 'MONTHLY SUBSCRIPTION SPEND', formula: `=IFERROR(TEXT(SUMPRODUCT(('💳 Subscriptions & Billing'!I10:I24="Active")*('💳 Subscriptions & Billing'!E10:E24)),"$#,##0.00"),"—")` },
    { range: "'🛡️ Security Dashboard'!M7", label: 'ACTIVE SUBSCRIPTIONS', formula: `=COUNTIF('💳 Subscriptions & Billing'!I10:I24,"Active")` },
  ];
  // Write combined label + value as two lines using CHAR(10) join won't work; use separate rows conceptually
  // Since each KPI is a single merged cell, write a text formula that shows label on top
  // Actually since we can't easily do multi-line in one cell with different sizes, write label as the cell value
  // and handle display via the large font for numbers approach
  KPI_VALS.forEach(k => {
    // Write just the formula — label will be the cell note or handled via formatting
    vals.push({ range: k.range, values: [[k.formula]] });
  });

  // Override: write rich KPI display
  // For row 6 we'll actually write label text before the merge, using a helper text row approach.
  // Better: just write the formula, and the number will be big. The user will understand context from section label.

  vals.push({ range: "'🛡️ Security Dashboard'!A9", values: [['  🔐  Security Score']] });
  vals.push({ range: "'🛡️ Security Dashboard'!A10", values: [['Security Score:']] });
  // Score: weighted formula: 50% MFA coverage + 30% strong passwords ratio + 20% no weak/critical
  vals.push({ range: "'🛡️ Security Dashboard'!C10", values: [['=IFERROR(TEXT(ROUND((COUNTIF(\'🔐 Account Vault\'!G10:G109,"Yes")/COUNTA(\'🔐 Account Vault\'!B10:B109))*50 + (COUNTIF(\'🔐 Account Vault\'!L10:L109,"Strong")/COUNTA(\'🔐 Account Vault\'!B10:B109))*30 + MAX(0,20-(COUNTIF(\'🔐 Account Vault\'!L10:L109,"Weak")+COUNTIF(\'🔐 Account Vault\'!L10:L109,"Critical"))*4),0)&"/100","—")']] });
  vals.push({ range: "'🛡️ Security Dashboard'!E10", values: [['=IFERROR(SPARKLINE({ROUND((COUNTIF(\'🔐 Account Vault\'!G10:G109,"Yes")/COUNTA(\'🔐 Account Vault\'!B10:B109))*50+(COUNTIF(\'🔐 Account Vault\'!L10:L109,"Strong")/COUNTA(\'🔐 Account Vault\'!B10:B109))*30+MAX(0,20-(COUNTIF(\'🔐 Account Vault\'!L10:L109,"Weak")+COUNTIF(\'🔐 Account Vault\'!L10:L109,"Critical"))*4),0),100-ROUND((COUNTIF(\'🔐 Account Vault\'!G10:G109,"Yes")/COUNTA(\'🔐 Account Vault\'!B10:B109))*50+(COUNTIF(\'🔐 Account Vault\'!L10:L109,"Strong")/COUNTA(\'🔐 Account Vault\'!B10:B109))*30+MAX(0,20-(COUNTIF(\'🔐 Account Vault\'!L10:L109,"Weak")+COUNTIF(\'🔐 Account Vault\'!L10:L109,"Critical"))*4),0)},{"charttype","bar";"color1","#2D2B8F";"color2","#E8ECFF"}),"—")']] });

  vals.push({ range: "'🛡️ Security Dashboard'!A12", values: [['  ⚠️  Accounts Needing Attention']] });
  vals.push({ range: "'🛡️ Security Dashboard'!A13", values: [['Service','Category','Username / Email','Strength','Status','MFA?','MFA Type','Action Needed']] });

  // FILTER: accounts where strength = Weak or Critical, or no MFA
  // =IFERROR(FILTER(range, condition), "No accounts needing attention")
  vals.push({ range: "'🛡️ Security Dashboard'!A14", values: [
    [`=IFERROR(FILTER('🔐 Account Vault'!B10:B109,('🔐 Account Vault'!L10:L109="Weak")+('🔐 Account Vault'!L10:L109="Critical")+('🔐 Account Vault'!G10:G109="No")),"✅ All accounts are in good standing")`],
    [`=IFERROR(FILTER('🔐 Account Vault'!E10:E109,('🔐 Account Vault'!L10:L109="Weak")+('🔐 Account Vault'!L10:L109="Critical")+('🔐 Account Vault'!G10:G109="No")),"")`,],
    [`=IFERROR(FILTER('🔐 Account Vault'!D10:D109,('🔐 Account Vault'!L10:L109="Weak")+('🔐 Account Vault'!L10:L109="Critical")+('🔐 Account Vault'!G10:G109="No")),"")`,],
    [`=IFERROR(FILTER('🔐 Account Vault'!L10:L109,('🔐 Account Vault'!L10:L109="Weak")+('🔐 Account Vault'!L10:L109="Critical")+('🔐 Account Vault'!G10:G109="No")),"")`,],
    [`=IFERROR(FILTER('🔐 Account Vault'!M10:M109,('🔐 Account Vault'!L10:L109="Weak")+('🔐 Account Vault'!L10:L109="Critical")+('🔐 Account Vault'!G10:G109="No")),"")`,],
    [`=IFERROR(FILTER('🔐 Account Vault'!G10:G109,('🔐 Account Vault'!L10:L109="Weak")+('🔐 Account Vault'!L10:L109="Critical")+('🔐 Account Vault'!G10:G109="No")),"")`,],
    [`=IFERROR(FILTER('🔐 Account Vault'!H10:H109,('🔐 Account Vault'!L10:L109="Weak")+('🔐 Account Vault'!L10:L109="Critical")+('🔐 Account Vault'!G10:G109="No")),"")`,],
  ].map((v, i) => {
    // These are COLUMN formulas, so each fills down. But we want row-by-row.
    // We'll place them as single-cell formulas that spill (modern Sheets supports this).
    return null; // handled below
  }).filter(Boolean)});

  // Actually let's write column-by-column FILTER formulas (they spill down)
  const filterBase = `('🔐 Account Vault'!L10:L109="Weak")+('🔐 Account Vault'!L10:L109="Critical")+('🔐 Account Vault'!G10:G109="No")`;
  const filterCols = [
    { col: 'A', src: `'🔐 Account Vault'!B10:B109` },
    { col: 'B', src: `'🔐 Account Vault'!E10:E109` },
    { col: 'C', src: `'🔐 Account Vault'!D10:D109` },
    { col: 'D', src: `'🔐 Account Vault'!L10:L109` },
    { col: 'E', src: `'🔐 Account Vault'!M10:M109` },
    { col: 'F', src: `'🔐 Account Vault'!G10:G109` },
    { col: 'G', src: `'🔐 Account Vault'!H10:H109` },
  ];
  filterCols.forEach(fc => {
    vals.push({ range: `'🛡️ Security Dashboard'!${fc.col}14`,
      values: [[`=IFERROR(FILTER(${fc.src},${filterBase}),${fc.col === 'A' ? '"✅ All accounts in good standing"' : '""'})`]]
    });
  });
  // Action Needed col H
  vals.push({ range: `'🛡️ Security Dashboard'!H14`,
    values: [[`=IFERROR(IF(FILTER('🔐 Account Vault'!L10:L109,${filterBase})="Weak","🔑 Update password",IF(FILTER('🔐 Account Vault'!L10:L109,${filterBase})="Critical","🚨 Update password immediately","📱 Enable MFA")),"—")`]]
  });

  vals.push({ range: "'🛡️ Security Dashboard'!A25", values: [['  🔑  Password Strength Breakdown']] });

  STRENGTHS.forEach((s, i) => {
    vals.push({ range: `'🛡️ Security Dashboard'!A${26 + i}`, values: [[s.label]] });
    vals.push({ range: `'🛡️ Security Dashboard'!C${26 + i}`, values: [[`=COUNTIF('🔐 Account Vault'!L10:L109,"${s.label}")`]] });
    vals.push({ range: `'🛡️ Security Dashboard'!F${26 + i}`, values: [[`=IFERROR(SPARKLINE({COUNTIF('🔐 Account Vault'!L10:L109,"${s.label}"),COUNTA('🔐 Account Vault'!B10:B109)-COUNTIF('🔐 Account Vault'!L10:L109,"${s.label}")},{"charttype","bar";"color1","#2D2B8F";"color2","#E8ECFF"}),"—")`]] });
  });

  vals.push({ range: "'🛡️ Security Dashboard'!A31", values: [['  📋  Account Status Overview']] });
  STATUSES.forEach((s, i) => {
    vals.push({ range: `'🛡️ Security Dashboard'!A${32 + i}`, values: [[s]] });
    vals.push({ range: `'🛡️ Security Dashboard'!C${32 + i}`, values: [[`=COUNTIF('🔐 Account Vault'!M10:M109,"${s}")`]] });
  });

  vals.push({ range: "'🛡️ Security Dashboard'!A38", values: [['  🗂️  Accounts by Category']] });
  CATS.forEach((cat, i) => {
    vals.push({ range: `'🛡️ Security Dashboard'!A${39 + i}`, values: [[cat]] });
    vals.push({ range: `'🛡️ Security Dashboard'!C${39 + i}`, values: [[`=COUNTIF('🔐 Account Vault'!E10:E109,"${cat}")`]] });
  });

  await valuesBatchUpdate(id, vals, 'dashboard-values');

  console.log('Security Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
