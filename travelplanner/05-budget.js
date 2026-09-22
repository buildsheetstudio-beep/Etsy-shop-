const { hex, COLOR, gridRange, batchUpdate, valuesBatchUpdate, vr } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = 1601326951; // Travel Budget

// 20 sample budget entries
const entries = [
  // Category, Description, Destination, Budgeted, Actual, DatePaid, PaymentMethod, Notes
  ['Accommodation','Hotel Duquesne Eiffel (7 nights)','Paris',980,975,'2025-06-01','Credit Card','Pre-paid, non-refundable'],
  ['Accommodation','Airbnb Eixample (8 nights)','Barcelona',720,735,'2025-06-15','Credit Card','Includes cleaning fee'],
  ['Accommodation','Hotel Artemide (7 nights)','Rome',840,850,'2025-06-01','Credit Card','Breakfast included'],
  ['Transport','Eurostar London → Paris','Paris',180,178,'2025-05-20','Credit Card','1st Class, 2 seats'],
  ['Transport','Flight Paris → Barcelona (Vueling)','Barcelona',120,118,'2025-05-20','Debit Card','Hand luggage only'],
  ['Transport','Flight Barcelona → Rome (Ryanair)','Rome',90,95,'2025-05-20','Debit Card','Extra bag fee'],
  ['Transport','Flight Rome → London (EasyJet)','Rome',140,145,'2025-05-20','Credit Card','Priority boarding'],
  ['Transport','Paris Metro 10-trip carnet','Paris',35,38,'2025-07-15','Cash','Navigo top-up'],
  ['Transport','Barcelona Metro/Bus passes','Barcelona',25,27,'2025-07-21','Cash','T-Casual card x2'],
  ['Transport','Rome Metro & buses','Rome',20,20,'2025-07-29','Cash','BIT tickets'],
  ['Food & Drink','Paris dining (7 days)','Paris',280,310,'2025-07-21','Mixed','Bistros, boulangeries, cafes'],
  ['Food & Drink','Barcelona dining (8 days)','Barcelona',240,265,'2025-07-29','Mixed','Tapas, paella, sangria'],
  ['Food & Drink','Rome dining (7 days)','Rome',210,228,'2025-08-05','Mixed','Pasta, pizza, gelato'],
  ['Activities','Louvre Museum tickets x2','Paris',25,25,'2025-07-16','Credit Card','Skip-the-line booked online'],
  ['Activities','Eiffel Tower summit x2','Paris',30,30,'2025-07-17','Credit Card','Evening visit'],
  ['Activities','Sagrada Família + towers x2','Barcelona',44,44,'2025-07-23','Credit Card','Morning slot'],
  ['Activities','Vatican Museums & Sistine Chapel x2','Rome',56,56,'2025-07-30','Credit Card','Early access tour'],
  ['Activities','Colosseum & Roman Forum x2','Rome',36,36,'2025-07-31','Credit Card','Arena floor access'],
  ['Shopping','Souvenirs, fashion, gifts','Mixed',250,220,'2025-08-04','Mixed','Under budget!'],
  ['Insurance','Annual travel insurance x2','Mixed',85,85,'2025-04-01','Credit Card','Covers entire trip'],
];

(async () => {
  const reqs = [];

  // ── Row 1: Title ──
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 12),
      cell: {
        userEnteredValue: { stringValue: '💰 TRAVEL BUDGET TRACKER' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.deepNavy),
          textFormat: { foregroundColor: hex(COLOR.antiqueGold), bold: true, fontSize: 18 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // ── Row 2: Subtitle ──
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, 12),
      cell: {
        userEnteredValue: { stringValue: 'European Summer Escape  ·  £4,500 budget  ·  Jul 15 – Aug 5, 2025' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.midnightBlue),
          textFormat: { foregroundColor: hex(COLOR.champagne), italic: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // ── Rows 3-8: Budget Summary KPIs ──
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 2, 3, 0, 12),
      cell: {
        userEnteredValue: { stringValue: '📊  BUDGET SUMMARY' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.forestGreen),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 12 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // KPI layout: 3 KPIs per row × 2 rows, each KPI = 4 cols
  const kpis = [
    { label: 'TOTAL BUDGET', formula: "=IFERROR('Trip Overview'!B9,4500)", fmt: '£#,##0.00', color: COLOR.deepNavy },
    { label: 'TOTAL SPENT', formula: '=SUM(F12:F211)', fmt: '£#,##0.00', color: COLOR.forestGreen },
    { label: 'REMAINING', formula: "=IFERROR('Trip Overview'!B9,4500)-SUM(F12:F211)", fmt: '£#,##0.00', color: COLOR.amber },
    { label: 'BUDGETED (PLANNED)', formula: '=SUM(E12:E211)', fmt: '£#,##0.00', color: COLOR.midnightBlue },
    { label: '% OF BUDGET USED', formula: '=IFERROR(SUM(F12:F211)/IFERROR(\'Trip Overview\'!B9,4500)*100,0)', fmt: '0.00"%"', color: COLOR.deepCrimson },
    { label: 'DAYS LEFT', formula: "=IFERROR(MAX(0,'Trip Overview'!C7-TODAY()),\"—\")", fmt: '0', color: COLOR.midnightBlue },
  ];

  kpis.forEach((k, i) => {
    const row = 3 + Math.floor(i / 3);
    const col = (i % 3) * 4;

    reqs.push({ mergeCells: { range: gridRange(SID, row, row+1, col, col+2), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, col, col+2),
        cell: {
          userEnteredValue: { stringValue: k.label },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.champagne),
            textFormat: { foregroundColor: hex(COLOR.deepNavy), bold: true, fontSize: 9 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
    reqs.push({ mergeCells: { range: gridRange(SID, row, row+1, col+2, col+4), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, col+2, col+4),
        cell: {
          userEnteredValue: { formulaValue: k.formula },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.formulaBg),
            textFormat: { foregroundColor: hex(k.color), bold: true, fontSize: 14 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
            numberFormat: { type: 'NUMBER', pattern: k.fmt },
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });

  // ── Category Summary section (rows 5-10) — right side ──
  // Actually, let me put category totals in columns N-P (helper area) to feed the donut chart
  // We'll add that separately. For now just the main entries table.

  // ── Row 11: Section header ──
  reqs.push({ mergeCells: { range: gridRange(SID, 10, 11, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 10, 11, 0, 12),
      cell: {
        userEnteredValue: { stringValue: '📋  BUDGET ENTRIES' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.midnightBlue),
          textFormat: { foregroundColor: hex(COLOR.antiqueGold), bold: true, fontSize: 12 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // ── Row 12: Column headers ──
  const cols = ['#','Category','Description','Destination','Budgeted (£)','Actual (£)','Difference','% Used','Status','Date Paid','Payment Method','Notes'];
  cols.forEach((h, c) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 11, 12, c, c+1),
        cell: {
          userEnteredValue: { stringValue: h },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.deepNavy),
            textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });

  // Data rows 12-31 (entries.length = 20, rows index 12-31)
  entries.forEach((e, i) => {
    const row = 12 + i;
    const bg = i % 2 === 0 ? hex(COLOR.white) : hex(COLOR.altRow);

    // # (auto number)
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 0, 1),
        cell: {
          userEnteredValue: { formulaValue: `=IF(B${row+1}="","",ROW()-11)` },
          userEnteredFormat: { backgroundColor: bg, textFormat: { bold: true }, horizontalAlignment: 'CENTER' },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });

    // Difference (col G = 6)
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 6, 7),
        cell: {
          userEnteredValue: { formulaValue: `=IF(OR(E${row+1}="",F${row+1}=""),"",F${row+1}-E${row+1})` },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.formulaBg),
            numberFormat: { type: 'NUMBER', pattern: '£#,##0.00' },
            horizontalAlignment: 'CENTER',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });

    // % Used (col H = 7)
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 7, 8),
        cell: {
          userEnteredValue: { formulaValue: `=IF(OR(E${row+1}=0,E${row+1}=""),"",F${row+1}/E${row+1})` },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.formulaBg),
            numberFormat: { type: 'NUMBER', pattern: '0.0%' },
            horizontalAlignment: 'CENTER',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });

    // Status (col I = 8)
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 8, 9),
        cell: {
          userEnteredValue: { formulaValue: `=IF(F${row+1}="","",IF(F${row+1}>E${row+1},"Over Budget",IF(F${row+1}=E${row+1},"On Track","Under Budget")))` },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.formulaBg),
            horizontalAlignment: 'CENTER', textFormat: { bold: true },
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });

    // Input cells bg
    [1,2,3,4,5,9,10,11].forEach(c => {
      reqs.push({
        repeatCell: {
          range: gridRange(SID, row, row+1, c, c+1),
          cell: {
            userEnteredFormat: {
              backgroundColor: bg,
              numberFormat: (c === 4 || c === 5) ? { type: 'NUMBER', pattern: '£#,##0.00' } :
                            c === 9 ? { type: 'DATE', pattern: 'dd mmm yyyy' } : undefined,
            },
          },
          fields: 'userEnteredFormat',
        },
      });
    });
  });

  // Blank rows 32-211 for future entries (#/Diff/Status formulas)
  for (let i = 20; i < 200; i++) {
    const row = 12 + i;
    const bg = i % 2 === 0 ? hex(COLOR.white) : hex(COLOR.altRow);
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 0, 12),
        cell: { userEnteredFormat: { backgroundColor: bg } },
        fields: 'userEnteredFormat.backgroundColor',
      },
    });
    reqs.push({
      updateCells: {
        range: gridRange(SID, row, row+1, 0, 1),
        rows: [{ values: [{ userEnteredValue: { formulaValue: `=IF(B${row+1}="","",ROW()-11)` } }] }],
        fields: 'userEnteredValue',
      },
    });
    reqs.push({
      updateCells: {
        range: gridRange(SID, row, row+1, 6, 7),
        rows: [{ values: [{ userEnteredValue: { formulaValue: `=IF(OR(E${row+1}="",F${row+1}=""),"",F${row+1}-E${row+1})` } }] }],
        fields: 'userEnteredValue',
      },
    });
    reqs.push({
      updateCells: {
        range: gridRange(SID, row, row+1, 7, 8),
        rows: [{ values: [{ userEnteredValue: { formulaValue: `=IF(OR(E${row+1}=0,E${row+1}=""),"",F${row+1}/E${row+1})` } }] }],
        fields: 'userEnteredValue',
      },
    });
    reqs.push({
      updateCells: {
        range: gridRange(SID, row, row+1, 8, 9),
        rows: [{ values: [{ userEnteredValue: { formulaValue: `=IF(F${row+1}="","",IF(F${row+1}>E${row+1},"Over Budget",IF(F${row+1}=E${row+1},"On Track","Under Budget")))` } }] }],
        fields: 'userEnteredValue',
      },
    });
  }

  // ── Category totals helper (cols N-P, rows 3-13) for donut chart ──
  const categories = ['Accommodation','Transport','Food & Drink','Activities','Shopping','Insurance','Tours','Visa & Fees','Miscellaneous'];
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 2, 3, 13, 16),
      cell: {
        userEnteredValue: { stringValue: '' },
        userEnteredFormat: { backgroundColor: hex(COLOR.champagne), textFormat: { bold: true, fontSize: 9 } },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  ['Category','Budgeted','Actual'].forEach((h, ci) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 2, 3, 13+ci, 14+ci),
        cell: {
          userEnteredValue: { stringValue: h },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.midnightBlue),
            textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 9 },
            horizontalAlignment: 'CENTER',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });

  // Column widths
  const widths = [40,140,220,100,110,110,100,80,120,100,120,180];
  widths.forEach((w, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: c, endIndex: c+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  // Helper cols N,O,P
  [120,100,100].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 13+i, endIndex: 14+i }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 12 }, properties: { pixelSize: 38 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 11, endIndex: 12 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 12, endIndex: 212 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'budget-structure');

  // ── Values ──
  const data = [];
  const entryRows = entries.map(e => [e[0],e[1],e[2],e[3],e[4],e[5],e[6],e[7]]);
  data.push(vr(`'Travel Budget'!B13`, entryRows));

  // Category totals helper
  const catRows = categories.map(cat => [
    cat,
    `=SUMIF(B$13:B$211,"${cat}",E$13:E$211)`,
    `=SUMIF(B$13:B$211,"${cat}",F$13:F$211)`,
  ]);
  data.push(vr(`'Travel Budget'!N4`, catRows));

  await valuesBatchUpdate(id, data, 'budget-values');

  // Format helper column cells
  const fmtReqs = [];
  categories.forEach((_, i) => {
    const row = 3 + i;
    fmtReqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 14, 16),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(COLOR.formulaBg),
            numberFormat: { type: 'NUMBER', pattern: '£#,##0.00' },
          },
        },
        fields: 'userEnteredFormat',
      },
    });
    fmtReqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 13, 14),
        cell: {
          userEnteredFormat: { backgroundColor: hex(COLOR.altRow) },
        },
        fields: 'userEnteredFormat.backgroundColor',
      },
    });
  });
  await batchUpdate(id, fmtReqs, 'budget-helper-fmt');

  console.log('Travel Budget built with', entries.length, 'entries');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
