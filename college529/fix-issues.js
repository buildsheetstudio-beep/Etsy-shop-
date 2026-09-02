'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID_BS = sheetMap['Beneficiary Setup'];
const SID_CE = sheetMap['College Cost Estimator'];
const SID_GM = sheetMap['Goals & Milestones'];

// ── Benchmark lookup by college type (matches 02-reference.js BENCHMARKS) ──
// [Tuition, Fees, Housing, Food, Books, Transport, Personal, TuitionInfl, OtherInfl]
const BENCH = {
  'In-State Public':        [11500, 1100, 10200, 5100, 1150, 1800, 2000, 0.04,  0.03 ],
  'Out-of-State Public':    [27000, 1400, 12500, 5500, 1200, 1500, 2200, 0.04,  0.03 ],
  'Private Nonprofit':      [39500, 1700, 13800, 6300, 1350, 1000, 2500, 0.045, 0.03 ],
  'Private For-Profit':     [21500, 1100, 11000, 5000, 1200, 1500, 2000, 0.035, 0.025],
  'Community College':      [ 4200,  580,  8500, 4200,  950, 2200, 1600, 0.03,  0.025],
  'Trade / Technical School':[ 7800,  750,  9200, 4600, 1450, 1600, 1800, 0.035, 0.025],
  'Custom School':          [    0,    0,     0,    0,    0,    0,    0, 0.04,  0.03 ],
};

// College type for each of the 9 CE rows (rows 10-18)
const CE_TYPES = [
  'In-State Public',
  'Out-of-State Public',
  'Out-of-State Public',
  'Private Nonprofit',
  'Private Nonprofit',
  'Community College',
  'Trade / Technical School',
  'Private For-Profit',
  'Custom School',
];

(async () => {
  const vals = [];
  const fmt  = [];

  // ── FIX 1: Beneficiary Setup M4 — wrong Contribution Log column refs ──────
  // Was: $I (Frequency) and $L (Account Name), rows starting at 8
  // Fix: $F (Transaction Type) and $H (Amount), rows starting at 6
  vals.push({
    range: "'Beneficiary Setup'!M4",
    values: [["=IFERROR(SUMPRODUCT((YEAR('Contribution Log'!$B$6:$B$5005)=YEAR(TODAY()))*('Contribution Log'!$F$6:$F$5005=\"Contribution\")*'Contribution Log'!$H$6:$H$5005),\"—\")"]]
  });

  // ── FIX 2: College Cost Estimator E10:M18 — seed with benchmark values ────
  // E-K = cost components, L-M = inflation rates
  CE_TYPES.forEach((type, i) => {
    const rowNum = 10 + i;
    const b = BENCH[type] || BENCH['Custom School'];
    vals.push({
      range: `'College Cost Estimator'!E${rowNum}:M${rowNum}`,
      values: [[b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7], b[8]]]
    });
  });

  // Apply input-cell formatting to the newly populated E-M cells
  fmt.push({ repeatCell: {
    range: gridRange(SID_CE, 9, 18, 4, 13),
    cell: { userEnteredFormat: { backgroundColor: hex(C.input) } },
    fields: 'userEnteredFormat.backgroundColor'
  }});
  // Currency format E-K
  fmt.push({ repeatCell: {
    range: gridRange(SID_CE, 9, 18, 4, 11),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } },
    fields: 'userEnteredFormat.numberFormat'
  }});
  // Percent format L-M
  fmt.push({ repeatCell: {
    range: gridRange(SID_CE, 9, 18, 11, 13),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } },
    fields: 'userEnteredFormat.numberFormat'
  }});

  // ── FIX 3: Goals & Milestones col C — replace 'Savings Milestone' with ────
  // the specific Reference Data milestone type, and give col D a proper title.
  // Row 8 = GOALS[0], sheet row = array index + 8
  const GM_FIXES = [
    // BEN-001 Emma (rows 8-12)
    { row:8,  c:'10% Funded',  d:'Emma — 10% Savings Goal'    },
    { row:9,  c:'25% Funded',  d:'Emma — 25% Savings Goal'    },
    { row:10, c:'50% Funded',  d:'Emma — 50% Savings Goal'    },
    { row:11, c:'75% Funded',  d:'Emma — 75% Savings Goal'    },
    { row:12, c:'100% Funded', d:'Emma — 100% Savings Goal'   },
    // BEN-002 Lucas (rows 15-19)
    { row:15, c:'10% Funded',  d:'Lucas — 10% Savings Goal'   },
    { row:16, c:'25% Funded',  d:'Lucas — 25% Savings Goal'   },
    { row:17, c:'50% Funded',  d:'Lucas — 50% Savings Goal'   },
    { row:18, c:'75% Funded',  d:'Lucas — 75% Savings Goal'   },
    { row:19, c:'100% Funded', d:'Lucas — 100% Savings Goal'  },
    // BEN-003 Sofia (rows 22-23)
    { row:22, c:'10% Funded',  d:'Sofia — 10% Savings Goal'   },
    { row:23, c:'25% Funded',  d:'Sofia — 25% Savings Goal'   },
    // BEN-004 Marcus (row 28)
    { row:28, c:'100% Funded', d:'Enrollment Balance — $60,000' },
    // BEN-005 Claire (rows 34-38)
    { row:34, c:'10% Funded',  d:'Claire — 10% Savings Goal'  },
    { row:35, c:'25% Funded',  d:'Claire — 25% Savings Goal'  },
    { row:36, c:'50% Funded',  d:'Claire — 50% Savings Goal'  },
    { row:37, c:'75% Funded',  d:'Claire — 75% Savings Goal'  },
    { row:38, c:'100% Funded', d:'Claire — 100% Savings Goal' },
  ];
  GM_FIXES.forEach(({ row, c, d }) => {
    vals.push({ range: `'Goals & Milestones'!C${row}:D${row}`, values: [[c, d]] });
  });

  await valuesBatchUpdate(id, vals, 'fix-issues values');
  await batchUpdate(id, fmt, 'fix-issues format');
  console.log('fix-issues done ✓');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
