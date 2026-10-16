'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BE = sheetMap['Budget & Expenses'];
const RD = sheetMap['Reference Data'];
const S = "'Budget & Expenses'";
const R = "'Reference Data'";
const G = "'Guest List & RSVP'";

// 16 expense categories + planned budgets
const CATEGORIES = [
  ['Venue', 900],
  ['Invitations', 80],
  ['Decorations', 350],
  ['Food', 800],
  ['Drinks', 400],
  ['Cake & Desserts', 200],
  ['Entertainment', 300],
  ['Music', 150],
  ['Audio / Visual', 200],
  ['Gifts', 250],
  ['Awards', 100],
  ['Photography', 300],
  ['Supplies', 150],
  ['Transportation', 0],
  ['Setup & Cleanup', 100],
  ['Miscellaneous', 220],
];

// 32 sample expenses
const EXPENSES = [
  // [date, category, vendor, item, qty, unitCost, actualCost, payStatus, dueDate, paidDate, receipt, notes]
  ['May 1, 2026','Venue','Lakeside Event Hall','Venue rental fee',1,900,900,'Deposit Paid','Jun 15, 2026','May 1, 2026',true,'50% deposit paid; balance due Jun 15'],
  ['May 3, 2026','Invitations','PrintSmart Co.','Premium card invitations - 120 pack',120,0.55,66,'Paid','May 3, 2026','May 3, 2026',true,'Includes envelopes'],
  ['May 3, 2026','Invitations','PrintSmart Co.','Postage stamps',50,0.68,34,'Paid','May 3, 2026','May 3, 2026',true,''],
  ['May 5, 2026','Decorations','Party Palace','Balloon arch kit',1,120,120,'Paid','May 5, 2026','May 5, 2026',true,'Deep indigo and champagne colors'],
  ['May 5, 2026','Decorations','Party Palace','Table centerpieces x12',12,18,216,'Paid','May 5, 2026','May 5, 2026',true,''],
  ['May 6, 2026','Decorations','Banner Express','Custom retirement banner - 6ft',1,75,75,'Paid','May 6, 2026','May 6, 2026',true,''],
  ['May 10, 2026','Food','Lakeside Catering','Plated dinner - 80 guests',80,18,1440,'Deposit Paid','Jun 15, 2026','May 10, 2026',true,'30% deposit; balance at delivery'],
  ['May 10, 2026','Food','Lakeside Catering','Vegetarian option uplift - 8 guests',8,5,40,'Planned','Jun 15, 2026','',false,''],
  ['May 12, 2026','Drinks','Wine & Spirits Co.','Champagne - 6 bottles',6,28,168,'Paid','May 12, 2026','May 12, 2026',true,'For toasts'],
  ['May 12, 2026','Drinks','Wine & Spirits Co.','Red & white wine - 12 bottles',12,15,180,'Paid','May 12, 2026','May 12, 2026',true,''],
  ['May 12, 2026','Drinks','Beverage Depot','Sparkling water and sodas - assorted',1,85,85,'Paid','May 12, 2026','May 12, 2026',true,''],
  ['May 14, 2026','Cake & Desserts','Sweet Endings Bakery','3-tier retirement cake',1,250,250,'Deposit Paid','Jun 20, 2026','May 14, 2026',true,'Balance due on delivery'],
  ['May 14, 2026','Cake & Desserts','Sweet Endings Bakery','Dessert table items',1,95,95,'Planned','Jun 20, 2026','',false,''],
  ['May 15, 2026','Entertainment','Riverside AV','Sound system rental',1,350,350,'Deposit Paid','Jun 25, 2026','May 15, 2026',true,''],
  ['May 15, 2026','Audio / Visual','Riverside AV','Projector and screen rental',1,200,200,'Deposit Paid','Jun 25, 2026','May 15, 2026',true,'Package deal with sound'],
  ['May 18, 2026','Music','DJ Smooth','DJ services - 3.5 hours',1,450,450,'Deposit Paid','Jun 25, 2026','May 18, 2026',true,'Includes playlist consultation'],
  ['May 20, 2026','Gifts','Trophy & Award Co.','Personalized crystal award',1,120,120,'Paid','May 20, 2026','May 20, 2026',true,'Engraved with name and years'],
  ['May 20, 2026','Gifts','Memory Lane Prints','Custom photo book',1,85,85,'Paid','May 20, 2026','May 20, 2026',true,'32 years photo collection'],
  ['May 20, 2026','Awards','Trophy & Award Co.','Retirement certificate - custom framed',1,75,75,'Paid','May 20, 2026','May 20, 2026',true,''],
  ['May 22, 2026','Photography','Riverside Photos','Event photography - 4 hours',1,400,400,'Deposit Paid','Jun 25, 2026','May 22, 2026',true,'Includes edited digital photos'],
  ['May 25, 2026','Supplies','Party Palace','Paper goods - plates, napkins, cups',1,95,95,'Paid','May 25, 2026','May 25, 2026',true,''],
  ['May 25, 2026','Supplies','Office Depot','Name tags, signage',1,35,35,'Paid','May 25, 2026','May 25, 2026',true,''],
  ['May 28, 2026','Decorations','Party Palace','Photo booth backdrop and props',1,150,150,'Paid','May 28, 2026','May 28, 2026',true,''],
  ['Jun 1, 2026','Setup & Cleanup','Event Services Inc.','Setup crew - 2 hours',2,45,90,'Planned','Jun 27, 2026','',false,'Day-of setup'],
  ['Jun 1, 2026','Setup & Cleanup','Event Services Inc.','Cleanup crew - 2 hours',2,35,70,'Planned','Jun 27, 2026','',false,'Post-party cleanup'],
  ['Jun 1, 2026','Miscellaneous','Michaels','Guest book and memory jar',1,45,45,'Paid','Jun 1, 2026','Jun 1, 2026',true,''],
  ['Jun 1, 2026','Miscellaneous','Amazon','Table numbers and place cards',1,28,28,'Paid','Jun 1, 2026','Jun 1, 2026',true,''],
  ['Jun 5, 2026','Food','Lakeside Catering','Appetizer add-on - 80 guests',80,4,320,'Planned','Jun 15, 2026','',false,'Added after headcount confirmed'],
  ['Jun 5, 2026','Entertainment','Event Supplies','Photo booth rental (4-hour)',1,250,250,'Deposit Paid','Jun 20, 2026','Jun 5, 2026',true,'Includes prints'],
  ['Jun 5, 2026','Miscellaneous','Canva (online)','Menu cards and program printing',1,32,32,'Paid','Jun 5, 2026','Jun 5, 2026',true,''],
  ['Jun 8, 2026','Gifts','Memory Lane Prints','Memory box for messages',1,55,55,'Paid','Jun 8, 2026','Jun 8, 2026',true,'Guests write retirement wishes'],
  ['Jun 8, 2026','Supplies','Party Palace','Table linens - 12 tables',12,12,144,'Planned','Jun 20, 2026','',false,''],
];

(async () => {
  const reqs = [];
  const data = [];

  // Background
  reqs.push({ repeatCell: {
    range: gridRange(BE,0,400,0,14),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) }},
    fields:'userEnteredFormat(backgroundColor)'
  }});

  // Freeze first
  reqs.push({ updateSheetProperties: {
    properties: { sheetId:BE, gridProperties:{ frozenRowCount:5 } },
    fields:'gridProperties.frozenRowCount'
  }});

  // Title
  data.push({ range:`${S}!A1`, values:[['BUDGET & EXPENSES']] });
  reqs.push({ mergeCells: { range: gridRange(BE,0,1,0,14), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(BE,0,1,0,14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:16, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:BE, dimension:'ROWS', startIndex:0, endIndex:1 },
    properties: { pixelSize:44 }, fields:'pixelSize'
  }});

  // Subtitle
  data.push({ range:`${S}!A2`, values:[['Track planned and actual costs by category — budget status updates automatically']] });
  reqs.push({ mergeCells: { range: gridRange(BE,1,2,0,14), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(BE,1,2,0,14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { italic:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});

  // Info rows 2-3
  data.push({ range:`${S}!A3`, values:[['Actual Spent by category calculates automatically from the expense log below (row 31+). Estimated Cost = Qty × Unit Cost.']] });
  data.push({ range:`${S}!A4`, values:[['Log expenses in rows 31+. Enter Category, Item, Qty, Unit Cost, and Actual Cost. Payment Status dropdown controls budget tracking.']] });
  [2,3].forEach(ri => {
    reqs.push({ mergeCells: { range: gridRange(BE,ri,ri+1,0,14), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(BE,ri,ri+1,0,14),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.bg),
        textFormat: { italic:true, fontSize:9, foregroundColor:hex(C.secText), fontFamily:'Arial' },
        horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:BE, dimension:'ROWS', startIndex:ri, endIndex:ri+1 },
      properties: { pixelSize:22 }, fields:'pixelSize'
    }});
  });

  // ── 8 Summary cards rows 4-9 (indices 4-9) ────────────────────────────────
  // 4 cards per row, 2 rows
  const summaryCards = [
    { label:'Total Budget',           formula:`=IFERROR('Party Setup'!B18,0)` },
    { label:'Planned Spending',       formula:`=IFERROR(SUM(B12:B27),0)` },
    { label:'Actual Spending',        formula:`=IFERROR(SUM(I31:I329),0)` },
    { label:'Remaining Budget',       formula:`=IFERROR('Party Setup'!B18-SUM(I31:I329),0)` },
    { label:'Amount Over / Under',    formula:`=IFERROR(SUM(B12:B27)-SUM(I31:I329),0)` },
    { label:'Cost per Confirmed Guest',formula:`=IFERROR(IF(COUNTIF(${G}!$I$6:$I$305,"Confirmed")=0,"—",SUM(I31:I329)/COUNTIF(${G}!$I$6:$I$305,"Confirmed")),"—")` },
    { label:'Paid Expenses',          formula:`=IFERROR(SUMIF(J31:J329,"Paid",I31:I329),0)` },
    { label:'Unpaid / Planned',       formula:`=IFERROR(SUMIF(J31:J329,"Planned",I31:I329)+SUMIF(J31:J329,"Deposit Paid",I31:I329),0)` },
  ];

  // Row indices for cards: 4,5 with 7 cols wide each... use 4 cards across 7 cols, 2 rows
  // Layout: cols A(0),B(1) label; C(2),D(3) value; E(4) spacer; F(5),G(6) label; H(7),I(8) value
  // But we only have 14 cols (A-N). Let me use a simpler 4-wide layout:
  // Card: cols A-C (label) + D-F (value) ... 6 cols per card pair, 2 pairs per row?
  // Actually: 2 cols per label + 2 cols per value = 4 cols per card, 3 cols separator
  // 4 cards × 3.5 cols = 14. Let me do: each card occupies 3 cols (label=1, value=2)
  // 4 cards = 12 cols. Remaining 2 cols can be spacers.

  // Simpler: 2 rows of 4 cards each
  // Card positions (0-indexed col start): 0, 4, 7, 11 (uneven but workable)
  // Let me use: label cols [0-1], value cols [2-3]; then label [4-5], value [6]; label [7-8], value [9]; label [10-12], value [13]
  // Actually simplest: label in col i*3+0..1, value in col i*3+2 for i=0,1,2,3 = cols 0-11, then cols 12-13
  // 4 cards fit in 12 cols: label(2)+value(1) x4 = exactly 12. We have 14 total.

  // Let me use 7 col groups: label(3) + value(4) each pair = 7 cols. 2 pairs = 14 cols.
  // Row 4: card 0 (A-C label, D-G value), card 1 (H-J label, K-N value)
  // Row 5: separator
  // Row 6: card 2, card 3
  // Row 7: separator
  // ...
  // Actually let me just do 2 cards per row across cols 0-6 and 7-13

  const cardPositions = [
    [4, 0, 3, 3, 7],   // row4, labelC1=0, labelC2=3, valC1=3, valC2=7
    [4, 7, 10, 10, 14], // row4, labelC1=7, labelC2=10, valC1=10, valC2=14
    [5, 0, 3, 3, 7],
    [5, 7, 10, 10, 14],
    [7, 0, 3, 3, 7],
    [7, 7, 10, 10, 14],
    [8, 0, 3, 3, 7],
    [8, 7, 10, 10, 14],
  ];

  summaryCards.forEach((card, i) => {
    const [ri, lc1, lc2, vc1, vc2] = cardPositions[i];
    const sheetRow = ri + 1;

    const labelCell = String.fromCharCode(65 + lc1) + sheetRow;
    const valueCell = String.fromCharCode(65 + vc1) + sheetRow;
    data.push({ range:`${S}!${labelCell}`, values:[[card.label]] });
    data.push({ range:`${S}!${valueCell}`, values:[[card.formula]] });

    reqs.push({ mergeCells: { range: gridRange(BE,ri,ri+1,lc1,lc2), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(BE,ri,ri+1,lc1,lc2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.primary),
        textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});

    reqs.push({ mergeCells: { range: gridRange(BE,ri,ri+1,vc1,vc2), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(BE,ri,ri+1,vc1,vc2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.panel),
        textFormat: { bold:true, fontSize:13, foregroundColor:hex(C.secondary), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});

    reqs.push({ updateDimensionProperties: {
      range: { sheetId:BE, dimension:'ROWS', startIndex:ri, endIndex:ri+1 },
      properties: { pixelSize:38 }, fields:'pixelSize'
    }});
  });

  // Separator between card groups (row index 6)
  reqs.push({ repeatCell: {
    range: gridRange(BE,6,7,0,14),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) }},
    fields:'userEnteredFormat(backgroundColor)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:BE, dimension:'ROWS', startIndex:6, endIndex:7 },
    properties: { pixelSize:8 }, fields:'pixelSize'
  }});

  // ── Category Budget Table (rows 10-26, index 9-26) ─────────────────────────
  // Section header row 9
  data.push({ range:`${S}!A10`, values:[['CATEGORY BUDGET — Planned vs. Actual']] });
  reqs.push({ mergeCells: { range: gridRange(BE,9,10,0,14), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(BE,9,10,0,14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { bold:true, fontSize:10, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:BE, dimension:'ROWS', startIndex:9, endIndex:10 },
    properties: { pixelSize:28 }, fields:'pixelSize'
  }});

  // Category table headers row 10 (index 10)
  data.push({ range:`${S}!A11`, values:[['Expense Category','Planned Budget','Actual Spent','Remaining','Variance %','Status']] });
  reqs.push({ repeatCell: {
    range: gridRange(BE,10,11,0,6),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});

  // Category data rows 11-26 (index 11-26, sheet rows 12-27)
  CATEGORIES.forEach(([cat, planned], i) => {
    const sheetRow = 12 + i;
    const ri = 11 + i;
    data.push({ range:`${S}!A${sheetRow}`, values:[[cat]] });
    data.push({ range:`${S}!B${sheetRow}`, values:[[planned]] });
    // C: Actual Spent SUMIF
    data.push({ range:`${S}!C${sheetRow}`, values:[[`=IFERROR(SUMIF($C$31:$C$329,"${cat}",$I$31:$I$329),0)`]] });
    // D: Remaining
    data.push({ range:`${S}!D${sheetRow}`, values:[[`=IFERROR(B${sheetRow}-C${sheetRow},0)`]] });
    // E: Variance %
    data.push({ range:`${S}!E${sheetRow}`, values:[[`=IFERROR(IF(B${sheetRow}=0,"—",(C${sheetRow}-B${sheetRow})/B${sheetRow}),"—")`]] });
    // F: Status
    data.push({ range:`${S}!F${sheetRow}`, values:[[`=IFERROR(IF(B${sheetRow}=0,"No Budget Set",IF(C${sheetRow}>B${sheetRow},"Over Budget",IF(C${sheetRow}/B${sheetRow}>0.85,"Near Budget","Under Budget"))),"—")`]] });

    // Row styling
    reqs.push({ repeatCell: {
      range: gridRange(BE,ri,ri+1,0,1),
      cell: { userEnteredFormat: {
        backgroundColor: ri%2===1 ? hex(C.altRow) : hex(C.panel),
        textFormat: { bold:true, fontSize:9, fontFamily:'Arial' },
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat)'
    }});
    reqs.push({ repeatCell: {
      range: gridRange(BE,ri,ri+1,1,2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.input),
        textFormat: { fontSize:9, fontFamily:'Arial' },
        numberFormat: { type:'CURRENCY', pattern:'$#,##0.00' }
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,numberFormat)'
    }});
    [2,3].forEach(ci => {
      reqs.push({ repeatCell: {
        range: gridRange(BE,ri,ri+1,ci,ci+1),
        cell: { userEnteredFormat: {
          backgroundColor: hex(C.formula),
          textFormat: { fontSize:9, fontFamily:'Arial' },
          numberFormat: { type:'CURRENCY', pattern:'$#,##0.00' }
        }},
        fields:'userEnteredFormat(backgroundColor,textFormat,numberFormat)'
      }});
    });
  });

  // Status CF
  const statusCF = [
    { val:'Under Budget', bg:C.success },
    { val:'Near Budget',  bg:C.warning },
    { val:'Over Budget',  bg:C.attention },
    { val:'No Budget Set',bg:C.border },
  ];
  statusCF.forEach((r, idx) => {
    reqs.push({ addConditionalFormatRule: { index:idx, rule: {
      ranges:[gridRange(BE,11,27,5,6)],
      booleanRule: {
        condition:{ type:'TEXT_EQ', values:[{userEnteredValue:r.val}] },
        format:{ backgroundColor:hex(r.bg) }
      }
    }}});
  });

  // ── Expense Log section (rows 28-329) ──────────────────────────────────────
  // Section header row 28 (index 28, but after cat table ends at index 27 = row 28)
  // Actually: category table: header at index 10 (row 11), data at indices 11-26 (rows 12-27)
  // So expense section starts at row 28 (index 27) → let's put section header at row 28 (index 27)
  data.push({ range:`${S}!A28`, values:[['EXPENSE LOG']] });
  reqs.push({ mergeCells: { range: gridRange(BE,27,28,0,14), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(BE,27,28,0,14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { bold:true, fontSize:10, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:BE, dimension:'ROWS', startIndex:27, endIndex:28 },
    properties: { pixelSize:28 }, fields:'pixelSize'
  }});

  // Expense log headers row 29 (index 29, sheet row 29... wait the spec says rows 30:329 and expense ID = ROW()-30)
  // Let me put headers at row 29 (index 28), data at rows 31-329 (index 30-328)
  // Wait: EXPENSE LOG header at row 28 (index 27). Then sub-header at row 29 (index 28). Then skip row 30 for nothing?
  // Actually spec says log rows 30:329, expense ID = IF(E31="","","EXP-"&TEXT(ROW()-30,"000"))
  // So data starts at row 31 (index 30). Header at row 30 (index 29). Section header at row 28 (index 27).
  // Row 29 can be empty/spacer or part of section header.

  // Expense header row 30 (index 29)
  const expHeaders = ['Expense ID','Date','Category','Vendor / Store','Item / Service','Quantity','Unit Cost','Estimated Cost','Actual Cost','Payment Status','Due Date','Paid Date','Receipt Saved?','Notes'];
  data.push({ range:`${S}!A30`, values:[expHeaders] });
  reqs.push({ repeatCell: {
    range: gridRange(BE,29,30,0,14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:BE, dimension:'ROWS', startIndex:29, endIndex:30 },
    properties: { pixelSize:40 }, fields:'pixelSize'
  }});

  // Expense data rows 31-62 (index 30-61)
  EXPENSES.forEach(([date, cat, vendor, item, qty, unitCost, actualCost, payStatus, dueDate, paidDate, receipt, notes], i) => {
    const row = 31 + i;
    data.push({ range:`${S}!A${row}`, values:[[`=IF(E${row}="","","EXP-"&TEXT(ROW()-30,"000"))`]] });
    data.push({ range:`${S}!B${row}`, values:[[date]] });
    data.push({ range:`${S}!C${row}`, values:[[cat]] });
    data.push({ range:`${S}!D${row}`, values:[[vendor]] });
    data.push({ range:`${S}!E${row}`, values:[[item]] });
    data.push({ range:`${S}!F${row}`, values:[[qty]] });
    data.push({ range:`${S}!G${row}`, values:[[unitCost]] });
    data.push({ range:`${S}!H${row}`, values:[[`=IFERROR(IF(F${row}="","",F${row}*G${row}),"")`]] });
    data.push({ range:`${S}!I${row}`, values:[[actualCost]] });
    data.push({ range:`${S}!J${row}`, values:[[payStatus]] });
    if (dueDate) data.push({ range:`${S}!K${row}`, values:[[dueDate]] });
    if (paidDate) data.push({ range:`${S}!L${row}`, values:[[paidDate]] });
    data.push({ range:`${S}!M${row}`, values:[[receipt]] });
    data.push({ range:`${S}!N${row}`, values:[[notes]] });
  });

  // Formula-only rows 63-329 (index 62-328)
  for (let row = 63; row <= 329; row++) {
    data.push({ range:`${S}!A${row}`, values:[[`=IF(E${row}="","","EXP-"&TEXT(ROW()-30,"000"))`]] });
    data.push({ range:`${S}!H${row}`, values:[[`=IFERROR(IF(F${row}="","",F${row}*G${row}),"")`]] });
  }

  // Alternating rows in expense log (index 30-328)
  for (let ri = 30; ri < 329; ri += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(BE,ri,ri+1,0,14),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) }},
      fields:'userEnteredFormat(backgroundColor)'
    }});
  }

  // Formula cols A,H styling in expense log
  [0,7].forEach(ci => {
    reqs.push({ repeatCell: {
      range: gridRange(BE,30,329,ci,ci+1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.formula), textFormat:{fontSize:9,fontFamily:'Arial'} }},
      fields:'userEnteredFormat(backgroundColor,textFormat)'
    }});
  });

  // Receipt checkbox col M (index 12)
  const dvReqs = [];
  dvReqs.push({ setDataValidation: {
    range: gridRange(BE,30,329,12,13),
    rule: { condition:{ type:'BOOLEAN' }, showCustomUi:true }
  }});

  // Payment Status dropdown col J (index 9)
  dvReqs.push({ setDataValidation: {
    range: gridRange(BE,30,329,9,10),
    rule: { condition:{ type:'ONE_OF_RANGE', values:[{ userEnteredValue:`=${R}!$I$2:$I$7` }] }, showCustomUi:true, strict:false }
  }});

  // Category dropdown col C (index 2)
  dvReqs.push({ setDataValidation: {
    range: gridRange(BE,30,329,2,3),
    rule: { condition:{ type:'ONE_OF_RANGE', values:[{ userEnteredValue:`=${R}!$G$2:$G$17` }] }, showCustomUi:true, strict:false }
  }});

  // Payment Status CF col J
  const paymentCF = [
    { val:'Paid', bg:C.success },
    { val:'Deposit Paid', bg:C.mutedBlue },
    { val:'Partially Paid', bg:C.warning },
    { val:'Planned', bg:C.bg },
    { val:'Cancelled', bg:C.border },
    { val:'Refunded', bg:C.mutedMauve },
  ];
  paymentCF.forEach((r, idx) => {
    reqs.push({ addConditionalFormatRule: { index:4+idx, rule: {
      ranges:[gridRange(BE,30,329,9,10)],
      booleanRule: {
        condition:{ type:'TEXT_EQ', values:[{userEnteredValue:r.val}] },
        format:{ backgroundColor:hex(r.bg) }
      }
    }}});
  });

  // ── Pivot data for charts (cols I-N, rows 1-20 area outside data) ──────────
  // Actually let's use a helper area at col H+ of the category table
  // Spending by category pivot at H11:J27
  CATEGORIES.forEach(([cat], i) => {
    const row = 12 + i;
    data.push({ range:`${S}!H${row}`, values:[[cat]] });
    data.push({ range:`${S}!I${row}`, values:[[`=IFERROR(SUMIF($C$31:$C$329,"${cat}",$I$31:$I$329),0)`]] });
  });
  data.push({ range:`${S}!H11`, values:[['Category']] });
  data.push({ range:`${S}!I11`, values:[['Actual Spent']] });

  // Payment status pivot at K11:L14
  data.push({ range:`${S}!K11`, values:[['Payment Status']] });
  data.push({ range:`${S}!L11`, values:[['Amount']] });
  [['Paid',`=IFERROR(SUMIF($J$31:$J$329,"Paid",$I$31:$I$329),0)`],
   ['Deposit Paid',`=IFERROR(SUMIF($J$31:$J$329,"Deposit Paid",$I$31:$I$329),0)`],
   ['Planned',`=IFERROR(SUMIF($J$31:$J$329,"Planned",$I$31:$I$329),0)`],
   ['Partially Paid',`=IFERROR(SUMIF($J$31:$J$329,"Partially Paid",$I$31:$I$329),0)`],
  ].forEach(([label, formula], i) => {
    data.push({ range:`${S}!K${12+i}`, values:[[label]] });
    data.push({ range:`${S}!L${12+i}`, values:[[formula]] });
  });

  // ── 3 Charts ──────────────────────────────────────────────────────────────
  const chartReqs = [];

  // Chart 1: Donut — Actual Spending by Category (H11:I27 = cols 7,8)
  chartReqs.push({ addChart: { chart: {
    spec: {
      title: 'Actual Spending by Category',
      pieChart: {
        legendPosition:'RIGHT_LEGEND',
        pieHole:0.4,
        domain: { sourceRange:{ sources:[{ sheetId:BE, startRowIndex:10, endRowIndex:27, startColumnIndex:7, endColumnIndex:8 }]}},
        series: { sourceRange:{ sources:[{ sheetId:BE, startRowIndex:10, endRowIndex:27, startColumnIndex:8, endColumnIndex:9 }]}}
      }
    },
    position:{ overlayPosition:{ anchorCell:{ sheetId:BE, rowIndex:4, columnIndex:9 }, widthPixels:420, heightPixels:250 } }
  }}});

  // Chart 2: Grouped column — Planned vs Actual (A11:C27)
  chartReqs.push({ addChart: { chart: {
    spec: {
      title: 'Planned vs. Actual by Category',
      basicChart: {
        chartType:'COLUMN',
        legendPosition:'RIGHT_LEGEND',
        domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:BE, startRowIndex:10, endRowIndex:27, startColumnIndex:0, endColumnIndex:1 }]}}}],
        series:[
          { series:{ sourceRange:{ sources:[{ sheetId:BE, startRowIndex:10, endRowIndex:27, startColumnIndex:1, endColumnIndex:2 }]}}, targetAxis:'LEFT_AXIS' },
          { series:{ sourceRange:{ sources:[{ sheetId:BE, startRowIndex:10, endRowIndex:27, startColumnIndex:2, endColumnIndex:3 }]}}, targetAxis:'LEFT_AXIS' },
        ],
        headerCount:1
      }
    },
    position:{ overlayPosition:{ anchorCell:{ sheetId:BE, rowIndex:4, columnIndex:9 }, widthPixels:450, heightPixels:250, offsetXPixels:430 } }
  }}});

  // Chart 3: Bar — Payment Status (K11:L15)
  chartReqs.push({ addChart: { chart: {
    spec: {
      title: 'Expenses by Payment Status',
      basicChart: {
        chartType:'BAR',
        legendPosition:'BOTTOM_LEGEND',
        domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:BE, startRowIndex:11, endRowIndex:15, startColumnIndex:10, endColumnIndex:11 }]}}}],
        series:[{ series:{ sourceRange:{ sources:[{ sheetId:BE, startRowIndex:11, endRowIndex:15, startColumnIndex:11, endColumnIndex:12 }]}}, targetAxis:'BOTTOM_AXIS' }],
        headerCount:1
      }
    },
    position:{ overlayPosition:{ anchorCell:{ sheetId:BE, rowIndex:9, columnIndex:9 }, widthPixels:400, heightPixels:220 } }
  }}});

  // Column widths
  const colW = [80,90,120,140,200,60,80,100,100,110,90,90,70,160];
  colW.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:BE, dimension:'COLUMNS', startIndex:i, endIndex:i+1 },
      properties: { pixelSize:w }, fields:'pixelSize'
    }});
  });

  // Tab color
  reqs.push({ updateSheetProperties: {
    properties: { sheetId:BE, tabColorStyle:{ rgbColor:hex(C.primary) } },
    fields:'tabColorStyle'
  }});

  await batchUpdate(id, reqs);
  await batchUpdate(id, dvReqs);
  await batchUpdate(id, chartReqs);
  await valuesBatchUpdate(id, data);
  console.log('Budget & Expenses complete');
})();
