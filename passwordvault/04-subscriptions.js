'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SUB = sheetMap['💳 Subscriptions & Billing'];

// 15 cols: A-O
// A: # | B: Service | C: Category | D: Billing Cycle | E: Monthly Cost
// F: Annual Cost | G: Next Billing | H: Payment Method | I: Status
// J: Free Trial End | K: Auto-Renew | L: Cancellation Notes | M: Account Link
// N: Date Started | O: Notes
// Summary block: Q-S cols (index 16-18)

const NCOLS = 15;  // A-O data
const NCOLS_FULL = 19; // up to col S for summary

const COL_WIDTHS = [
  30,   // A: #
  160,  // B: Service
  140,  // C: Category
  100,  // D: Billing Cycle
  100,  // E: Monthly Cost
  100,  // F: Annual Cost
  100,  // G: Next Billing
  160,  // H: Payment Method
  90,   // I: Status
  90,   // J: Trial End
  80,   // K: Auto-Renew
  180,  // L: Cancel Notes
  180,  // M: Account Link
  90,   // N: Date Started
  180,  // O: Notes
  30,   // P: gap
  120,  // Q: Summary label
  130,  // R: Summary value
  30,   // S: spare
];

// Row map
const ROW_TITLE    = 0;
const ROW_NOTICE_1 = 1;
const ROW_NOTICE_2 = 2;
const ROW_NOTICE_3 = 3;
const ROW_KPI_HDR  = 4;
const ROW_KPI_1    = 5; // Total Monthly Cost
const ROW_KPI_2    = 6; // Total Annual Cost
const ROW_KPI_3    = 7; // Active Subscriptions
const ROW_HEADER   = 8;
const ROW_DATA_START = 9;
const ROW_DATA_END   = 24; // 15 rows (rows 10-24)

// 15 sample subscriptions
const SUBS = [
  [1, 'Netflix',          'Streaming Video',   'Monthly',  15.49, '=E10*12',  '15 Jul 2025','Visa •••4521','Active',  '—',         'Yes','—',                     'netflix.com',         '10 Oct 2019','Family Standard plan – 4 screens'],
  [2, 'Spotify Premium',  'Streaming Music',   'Monthly',  10.99, '=E11*12',  '01 Jul 2025','Visa •••4521','Active',  '—',         'Yes','—',                     'spotify.com',         '01 Apr 2018','Individual premium'],
  [3, 'Disney+',          'Streaming Video',   'Monthly',   7.99, '=E12*12',  '22 Jul 2025','Visa •••4521','Active',  '—',         'Yes','—',                     'disneyplus.com',      '22 Nov 2019','Bundled with Hulu'],
  [4, 'Hulu',             'Streaming Video',   'Monthly',   7.99, '=E13*12',  '22 Jul 2025','Visa •••4521','Active',  '—',         'Yes','—',                     'hulu.com',            '22 Nov 2019','Bundle with Disney+'],
  [5, 'Amazon Prime',     'Shopping Membership','Annual',   14.99, 139.00,    '15 Mar 2026','Mastercard •••8833','Active','—',      'Yes','—',                     'amazon.com/prime',    '15 Mar 2022','Free shipping + Prime Video'],
  [6, 'YouTube Premium',  'Streaming Video',   'Monthly',  13.99, '=E15*12',  '14 Jul 2025','Google Pay','Active',   '—',         'Yes','—',                     'youtube.com/premium', '14 Mar 2021','Ad-free + YouTube Music'],
  [7, 'Microsoft 365',    'Software / SaaS',   'Annual',   12.50, 149.99,    '01 Jan 2026','Mastercard •••8833','Active','—',      'Yes','—',                     'microsoft.com/365',   '01 Jan 2021','Work subscription (personal use share)'],
  [8, 'Dropbox Plus',     'Cloud Storage',     'Annual',   11.99, 119.99,    '05 May 2026','Visa •••4521','Active',  '—',         'Yes','—',                     'dropbox.com',         '05 May 2017','2 TB storage'],
  [9, 'Adobe CC',         'Software / SaaS',   'Monthly',  54.99, '=E18*12',  '10 Jul 2025','Mastercard •••8833','Active','—',      'Yes','Consider downgrading',  'adobe.com/creativecloud','10 Jun 2022','Photography plan + Lightroom'],
  [10,'1Password',        'Security',          'Annual',    2.99, 35.88,     '01 Jun 2026','Visa •••4521','Active',  '—',         'Yes','—',                     '1password.com',       '01 Jun 2023','Password manager – 5 family members'],
  [11,'Headspace',        'Health & Fitness',  'Annual',    5.83, 69.99,     '14 Feb 2026','Visa •••4521','Active',  '—',         'Yes','—',                     'headspace.com',       '14 Feb 2022','Meditation & mindfulness'],
  [12,'The New York Times','News & Media',      'Monthly',   4.00, '=E21*12',  '08 Jul 2025','Visa •••4521','Active',  '—',         'Yes','Intro rate ends Sep 25', 'nytimes.com',         '08 Jan 2025','Digital only – basic plan'],
  [13,'Duolingo Plus',    'Health & Fitness',  'Annual',    6.67, 79.99,     '20 Nov 2025','Google Pay','Active',   '—',         'Yes','—',                     'duolingo.com',        '20 Nov 2023','Language learning – Spanish'],
  [14,'LinkedIn Premium', 'Software / SaaS',   'Monthly',  39.99, '=E23*12',  '30 Jul 2025','Mastercard •••8833','Paused','—',     'No', 'Paused – evaluating ROI','linkedin.com/premium', '30 Apr 2024','Career plan – paused pending review'],
  [15,'Calm',             'Health & Fitness',  'Annual',    3.75, 44.99,     '10 Jan 2026','Visa •••4521','Trial',  '10 Jul 2025','No', 'Cancel before 10 Jul',  'calm.com',            '10 Jan 2025','Sleep & mindfulness – free trial'],
];

(async () => {
  const reqs = [];

  // ── Column widths ─────────────────────────────────────────────────────
  COL_WIDTHS.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SUB, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w },
      fields: 'pixelSize',
    }});
  });

  // ── Row heights ───────────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SUB, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 44 }, fields: 'pixelSize',
  }});
  [1,2,3].forEach(r => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SUB, dimension: 'ROWS', startIndex: r, endIndex: r + 1 },
      properties: { pixelSize: 22 }, fields: 'pixelSize',
    }});
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SUB, dimension: 'ROWS', startIndex: ROW_KPI_HDR, endIndex: ROW_KPI_HDR + 1 },
    properties: { pixelSize: 26 }, fields: 'pixelSize',
  }});
  [ROW_KPI_1, ROW_KPI_2, ROW_KPI_3].forEach(r => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SUB, dimension: 'ROWS', startIndex: r, endIndex: r + 1 },
      properties: { pixelSize: 24 }, fields: 'pixelSize',
    }});
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SUB, dimension: 'ROWS', startIndex: ROW_HEADER, endIndex: ROW_HEADER + 1 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SUB, dimension: 'ROWS', startIndex: ROW_DATA_START, endIndex: ROW_DATA_END },
    properties: { pixelSize: 22 }, fields: 'pixelSize',
  }});

  // ── TITLE ROW ─────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SUB, ROW_TITLE, ROW_TITLE + 1, 0, NCOLS_FULL),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepIndigo) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ mergeCells: { range: gridRange(SUB, ROW_TITLE, ROW_TITLE + 1, 2, NCOLS_FULL), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(SUB, ROW_TITLE, ROW_TITLE + 1, 2, NCOLS_FULL),
    cell: { userEnteredFormat: {
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16, fontFamily: 'Montserrat' },
      verticalAlignment: 'MIDDLE',
      horizontalAlignment: 'LEFT',
    }},
    fields: 'userEnteredFormat(textFormat,verticalAlignment,horizontalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(SUB, ROW_TITLE, ROW_TITLE + 1, 0, 2),
    cell: { userEnteredFormat: {
      textFormat: { foregroundColor: hex(C.clearBlue), bold: true, fontSize: 14 },
      verticalAlignment: 'MIDDLE',
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(textFormat,verticalAlignment,horizontalAlignment)',
  }});

  // ── SECURITY NOTICE ROWS 2-4 ──────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SUB, 1, 4, 0, NCOLS_FULL),
    cell: { userEnteredFormat: { backgroundColor: hex(C.noticeBg) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  [1, 2, 3].forEach(r => {
    reqs.push({ mergeCells: { range: gridRange(SUB, r, r + 1, 2, NCOLS_FULL), mergeType: 'MERGE_ALL' }});
    reqs.push({ repeatCell: {
      range: gridRange(SUB, r, r + 1, 2, NCOLS_FULL),
      cell: { userEnteredFormat: {
        textFormat: { foregroundColor: hex(C.noticeFg), bold: r === 1, fontSize: r === 1 ? 10 : 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(textFormat,verticalAlignment)',
    }});
  });
  reqs.push({ updateBorders: {
    range: gridRange(SUB, 3, 4, 0, NCOLS_FULL),
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.noticeBorder), width: 2 },
  }});

  // ── KPI HEADER ROW (row 5) ─────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SUB, ROW_KPI_HDR, ROW_KPI_HDR + 1, 0, NCOLS_FULL),
    cell: { userEnteredFormat: { backgroundColor: hex(C.midnightIndigo) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ mergeCells: { range: gridRange(SUB, ROW_KPI_HDR, ROW_KPI_HDR + 1, 2, 8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(SUB, ROW_KPI_HDR, ROW_KPI_HDR + 1, 2, 8),
    cell: { userEnteredFormat: {
      textFormat: { foregroundColor: hex(C.clearBlue), bold: true, fontSize: 10 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(textFormat,verticalAlignment)',
  }});
  // KPI value boxes: cols I(8), J(9), K(10) → merge pairs
  // Monthly label + value → cols I-J; Annual cols K-L; Count cols M-N
  [[8,10],[10,12],[12,14]].forEach(([s,e]) => {
    reqs.push({ mergeCells: { range: gridRange(SUB, ROW_KPI_HDR, ROW_KPI_HDR + 1, s, e), mergeType: 'MERGE_ALL' }});
    reqs.push({ repeatCell: {
      range: gridRange(SUB, ROW_KPI_HDR, ROW_KPI_HDR + 1, s, e),
      cell: { userEnteredFormat: {
        textFormat: { foregroundColor: hex(C.mutedIndigo), bold: true, fontSize: 8 },
        horizontalAlignment: 'CENTER',
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)',
    }});
  });

  // KPI data rows (rows 6-8)
  [ROW_KPI_1, ROW_KPI_2, ROW_KPI_3].forEach(r => {
    reqs.push({ repeatCell: {
      range: gridRange(SUB, r, r + 1, 0, NCOLS_FULL),
      cell: { userEnteredFormat: { backgroundColor: hex(C.veryPaleIndigo) }},
      fields: 'userEnteredFormat.backgroundColor',
    }});
  });

  // ── COLUMN HEADERS (row 9) ─────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SUB, ROW_HEADER, ROW_HEADER + 1, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.mediumIndigo),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
      verticalAlignment: 'MIDDLE',
      wrapStrategy: 'WRAP',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
  }});

  // ── DATA ROWS ──────────────────────────────────────────────────────────
  for (let r = ROW_DATA_START; r < ROW_DATA_END; r++) {
    const bg = (r - ROW_DATA_START) % 2 === 0 ? C.paleIndigo : C.silkyWhite;
    reqs.push({ repeatCell: {
      range: gridRange(SUB, r, r + 1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // # col — center
  reqs.push({ repeatCell: {
    range: gridRange(SUB, ROW_DATA_START, ROW_DATA_END, 0, 1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { foregroundColor: hex(C.mutedIndigo), fontSize: 9 } }},
    fields: 'userEnteredFormat(horizontalAlignment,textFormat)',
  }});

  // Cost columns (E,F) — currency format
  [4, 5].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(SUB, ROW_DATA_START, ROW_DATA_END, c, c + 1),
      cell: { userEnteredFormat: {
        numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
        horizontalAlignment: 'RIGHT',
        backgroundColor: hex(C.veryPaleIndigo),
      }},
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment,backgroundColor)',
    }});
  });

  // Status + AutoRenew — center
  [8, 10].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(SUB, ROW_DATA_START, ROW_DATA_END, c, c + 1),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
  });

  // Date cols (G, J, N) — center
  [6, 9, 13].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(SUB, ROW_DATA_START, ROW_DATA_END, c, c + 1),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
  });

  // Borders
  reqs.push({ updateBorders: {
    range: gridRange(SUB, ROW_HEADER, ROW_DATA_END, 0, NCOLS),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.deepIndigo), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.deepIndigo), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.deepIndigo), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.deepIndigo), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Category Summary block (Q-S, rows 9-26)
  reqs.push({ repeatCell: {
    range: gridRange(SUB, ROW_HEADER, ROW_DATA_END + 3, 16, 19),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.veryPaleIndigo),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(SUB, ROW_HEADER, ROW_HEADER + 1, 16, 19),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.mediumIndigo),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateBorders: {
    range: gridRange(SUB, ROW_HEADER, ROW_DATA_END + 3, 16, 19),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.deepIndigo), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.deepIndigo), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.deepIndigo), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.deepIndigo), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  await batchUpdate(id, reqs, 'sub-format');

  // ── VALUES ─────────────────────────────────────────────────────────────
  const vals = [];

  vals.push({ range: "'💳 Subscriptions & Billing'!C1", values: [['  💳  Subscriptions & Billing Tracker']] });

  vals.push({ range: "'💳 Subscriptions & Billing'!C2", values: [['⚠️  BILLING NOTICE — Review your subscriptions regularly. Cancel unused services before trial periods end.']] });
  vals.push({ range: "'💳 Subscriptions & Billing'!C3", values: [['💡  Tip: Annual billing usually saves 20-40% vs monthly. Check the "Annual Cost" column for savings opportunities.']] });
  vals.push({ range: "'💳 Subscriptions & Billing'!C4", values: [['📌  Highlighted rows = Paused or Trial subscriptions requiring action. Red = cancellation needed soon.']] });

  vals.push({ range: "'💳 Subscriptions & Billing'!C5", values: [['  📊  Spending Summary']] });
  vals.push({ range: "'💳 Subscriptions & Billing'!I5", values: [['Total Monthly']] });
  vals.push({ range: "'💳 Subscriptions & Billing'!K5", values: [['Total Annual']] });
  vals.push({ range: "'💳 Subscriptions & Billing'!M5", values: [['Active Count']] });

  vals.push({ range: "'💳 Subscriptions & Billing'!C6", values: [['Monthly spend (active only):']] });
  vals.push({ range: "'💳 Subscriptions & Billing'!I6", values: [['=SUMPRODUCT((I10:I24="Active")*(E10:E24))']] });
  vals.push({ range: "'💳 Subscriptions & Billing'!C7", values: [['Annual spend (billed annually):']] });
  vals.push({ range: "'💳 Subscriptions & Billing'!I7", values: [['=SUMPRODUCT((I10:I24="Active")*(F10:F24))']] });
  vals.push({ range: "'💳 Subscriptions & Billing'!C8", values: [['Active subscriptions:']] });
  vals.push({ range: "'💳 Subscriptions & Billing'!I8", values: [['=COUNTIF(I10:I24,"Active")']] });

  // KPI value formats
  const kpiCurrencyReqs = [];
  [[5,9],[6,9]].forEach(([r0idx, c]) => {
    kpiCurrencyReqs.push({ repeatCell: {
      range: gridRange(SUB, ROW_KPI_1 + (r0idx - 5), ROW_KPI_1 + (r0idx - 5) + 1, 8, 10),
      cell: { userEnteredFormat: {
        numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
        textFormat: { foregroundColor: hex(C.deepIndigo), bold: true, fontSize: 12 },
        horizontalAlignment: 'CENTER',
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(numberFormat,textFormat,horizontalAlignment,verticalAlignment)',
    }});
  });
  kpiCurrencyReqs.push({ repeatCell: {
    range: gridRange(SUB, ROW_KPI_3, ROW_KPI_3 + 1, 8, 10),
    cell: { userEnteredFormat: {
      textFormat: { foregroundColor: hex(C.deepIndigo), bold: true, fontSize: 12 },
      horizontalAlignment: 'CENTER',
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)',
  }});
  await batchUpdate(id, kpiCurrencyReqs, 'sub-kpi-format');

  // Column headers
  vals.push({ range: "'💳 Subscriptions & Billing'!A9", values: [['#','Service','Category','Billing Cycle','Monthly Cost','Annual Cost','Next Billing','Payment Method','Status','Trial End','Auto-Renew','Cancellation Notes','Account Link','Date Started','Notes']] });

  // Summary block headers
  vals.push({ range: "'💳 Subscriptions & Billing'!Q9", values: [['Category','Monthly $','Count']] });

  // Data rows
  vals.push({ range: "'💳 Subscriptions & Billing'!A10", values: SUBS });

  // Category summary (using SUMIF on data)
  const cats = ['Streaming Video','Streaming Music','Cloud Storage','Software / SaaS','News & Media','Gaming','Health & Fitness','Shopping Membership','Financial Tool','Communication','Security','Other'];
  const summaryVals = cats.map(cat => [
    cat,
    { formulaValue: `=IFERROR(SUMIF(C10:C24,"${cat}",E10:E24),0)` },
    { formulaValue: `=IFERROR(COUNTIF(C10:C24,"${cat}"),0)` },
  ]);
  // Manually build the range values
  const summaryRows = cats.map(cat => [
    cat,
    `=IFERROR(SUMIF(C$10:C$24,"${cat}",E$10:E$24),0)`,
    `=IFERROR(COUNTIF(C$10:C$24,"${cat}"),0)`,
  ]);
  vals.push({ range: "'💳 Subscriptions & Billing'!Q10", values: summaryRows });

  await valuesBatchUpdate(id, vals, 'sub-values');

  // Format summary value column as currency
  const summaryFmtReqs = [];
  summaryFmtReqs.push({ repeatCell: {
    range: gridRange(SUB, ROW_DATA_START, ROW_DATA_START + cats.length, 17, 18),
    cell: { userEnteredFormat: {
      numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
      horizontalAlignment: 'RIGHT',
    }},
    fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
  }});
  summaryFmtReqs.push({ repeatCell: {
    range: gridRange(SUB, ROW_DATA_START, ROW_DATA_START + cats.length, 18, 19),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
    fields: 'userEnteredFormat.horizontalAlignment',
  }});
  await batchUpdate(id, summaryFmtReqs, 'sub-summary-format');

  console.log('Subscriptions & Billing complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
