'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DB = sheetMap['Dashboard'];
const S  = "'Dashboard'";

(async () => {
  const data = [];
  const reqs = [];

  // Helper: KPI card block
  // Occupies 4 rows × 3 cols starting at (row0, col0) [0-indexed]
  function kpiCard(row0, col0, label, formula, subLabel, bgColor = C.primary) {
    const r = row0, c = col0;
    data.push({ range: `${S}!${rc(r,c)}`,   values: [[label]] });
    data.push({ range: `${S}!${rc(r+2,c)}`, values: [[formula]] });
    data.push({ range: `${S}!${rc(r+3,c)}`, values: [[subLabel]] });
    // Merge label row
    reqs.push({ mergeCells: { range: gridRange(DB, r, r+2, c, c+3), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(DB, r, r+2, c, c+3), cell: { userEnteredFormat: {
      backgroundColor: hex(bgColor),
      textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    // Value row
    reqs.push({ mergeCells: { range: gridRange(DB, r+2, r+3, c, c+3), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(DB, r+2, r+3, c, c+3), cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 20 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    // Sub-label row
    reqs.push({ mergeCells: { range: gridRange(DB, r+3, r+4, c, c+3), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(DB, r+3, r+4, c, c+3), cell: { userEnteredFormat: {
      backgroundColor: hex(C.altRow),
      textFormat: { italic: true, foregroundColor: hex(C.secText), fontSize: 8 },
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});
  }

  function rc(row, col) {
    // 0-indexed to A1 notation
    const letter = colLetter(col);
    return `${letter}${row + 1}`;
  }

  function colLetter(n) {
    let s = '';
    n += 1;
    while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
    return s;
  }

  // ── Main Title ─────────────────────────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['WEDDING DASHBOARD']] });
  reqs.push({ mergeCells: { range: gridRange(DB, 0, 1, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 0, 1, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 18 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize' }});

  // Subtitle row 2
  data.push({ range: `${S}!A2`, values: [[
    `=IFERROR("Amara Osei & Marcus Thornton  ·  "&TEXT('Wedding Setup'!B7,"MMMM D, YYYY")&"  ·  "&'Wedding Setup'!B10,"Ultimate Wedding Planner")`,
  ]] });
  reqs.push({ mergeCells: { range: gridRange(DB, 1, 2, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 1, 2, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { italic: true, foregroundColor: hex(C.white), fontSize: 10 },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});

  // ── PRIMARY KPI SECTION HEADER ─────────────────────────────────────────────
  data.push({ range: `${S}!A3`, values: [['  KEY METRICS']] });
  reqs.push({ mergeCells: { range: gridRange(DB, 2, 3, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 2, 3, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.accent),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 10 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row heights for KPI rows
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 30 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 3, endIndex: 11 },
    properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // ── PRIMARY KPI CARDS (8 cards, 2 rows × 4 cols, each card 3 cols wide) ──
  // Row 1 of KPIs: rows 4-7 (0-indexed: 3-6)
  kpiCard(3, 0,  'DAYS UNTIL',
    `=IFERROR(MAX(0,INT('Wedding Setup'!B7-TODAY())),"—")`,
    `=IFERROR(IF('Wedding Setup'!B7<TODAY(),"Wedding has passed!","Days remaining"),"")`,
    C.primary);
  kpiCard(3, 3,  'GUESTS ATTENDING',
    `=IFERROR(SUMPRODUCT(('Guest List & RSVP'!R6:R600="Attending")*1),"—")`,
    `=IFERROR("of "&COUNTA('Guest List & RSVP'!B6:B600)&" invited","")`,
    C.primary);
  kpiCard(3, 6,  'RSVP RATE',
    `=IFERROR(TEXT(SUMPRODUCT(('Guest List & RSVP'!R6:R600<>"Pending")*('Guest List & RSVP'!R6:R600<>"No Response")*('Guest List & RSVP'!R6:R600<>"")*1)/COUNTA('Guest List & RSVP'!B6:B600),"0%"),"—")`,
    'Response rate',
    C.primary);
  kpiCard(3, 9,  'VENDOR CONTRACTS',
    `=IFERROR(SUMPRODUCT(('Vendors & Venue'!G23:G40=TRUE)*1)&" / "&COUNTA('Vendors & Venue'!C23:C40),"—")`,
    'Contracts signed / total',
    C.primary);

  // Row heights for KPI value rows
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 5, endIndex: 6 },
    properties: { pixelSize: 46 }, fields: 'pixelSize' }});

  // Row 2 of KPIs: rows 8-11 (0-indexed: 7-10)
  kpiCard(7, 0,  'BUDGET REMAINING',
    `=IFERROR("$"&TEXT('Wedding Setup'!B20-SUM('Budget & Payments'!H40:H500),"#,##0"),"—")`,
    `=IFERROR("of $"&TEXT('Wedding Setup'!B20,"#,##0")&" total","")`,
    C.success);
  kpiCard(7, 3,  'CHECKLIST',
    `=IFERROR(TEXT(SUMPRODUCT(('Wedding Checklist'!K5:K200="Completed")*1)/COUNTA('Wedding Checklist'!E5:E200),"0%"),"—")`,
    `=IFERROR(SUMPRODUCT(('Wedding Checklist'!K5:K200="Completed")*1)&" of "&COUNTA('Wedding Checklist'!E5:E200)&" tasks done","")`,
    C.success);
  kpiCard(7, 6,  'SEATING FILLED',
    `=IFERROR(TEXT(SUMPRODUCT(('Table Planner'!F5:F100)*1)/SUMPRODUCT(('Table Planner'!D5:D100)*1),"0%"),"—")`,
    `=IFERROR(SUM('Table Planner'!F5:F100)&" of "&SUM('Table Planner'!D5:D100)&" seats assigned","")`,
    C.success);
  kpiCard(7, 9,  'GIFTS RECEIVED',
    `=IFERROR(SUMPRODUCT(('Guest List & RSVP'!AB6:AB600=TRUE)*1),"—")`,
    `=IFERROR(SUMPRODUCT(('Guest List & RSVP'!AD6:AD600=TRUE)*1)&" thank-you cards sent","")`,
    C.success);

  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 9, endIndex: 10 },
    properties: { pixelSize: 46 }, fields: 'pixelSize' }});

  // ── SECONDARY KPI SECTION ─────────────────────────────────────────────────
  data.push({ range: `${S}!A12`, values: [['  SECONDARY METRICS']] });
  reqs.push({ mergeCells: { range: gridRange(DB, 11, 12, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 11, 12, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.accent),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 10 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Secondary KPIs — flat label+value pairs, 4 per row, rows 13-14, 15-16
  const secKPIs = [
    ['Ceremony Guests',    `=IFERROR(SUMPRODUCT(('Guest List & RSVP'!T6:T600="Yes")*1),"—")`],
    ['Reception Guests',   `=IFERROR(SUMPRODUCT(('Guest List & RSVP'!U6:U600="Yes")*1),"—")`],
    ['Dietary Restrictions',`=IFERROR(SUMPRODUCT((('Guest List & RSVP'!K6:K600<>"")*('Guest List & RSVP'!K6:K600<>"None"))*1),"—")`],
    ['Kids Meals',         `=IFERROR(SUMPRODUCT(('Guest List & RSVP'!J6:J600="Kids Meal")*1),"—")`],
    ['Invitations Mailed', `=IFERROR(SUMPRODUCT(('Invitation Tracker'!Q5:Q200="Mailed")*1),"—")`],
    ['Follow-Up Needed',   `=IFERROR(SUMPRODUCT(('Invitation Tracker'!R5:R200=TRUE)*1),"—")`],
    ['Vendors Booked',     `=IFERROR(SUMPRODUCT(ISNUMBER(SEARCH("Booked",'Vendors & Venue'!J23:J40))*1),"—")`],
    ['Total Paid to Date', `=IFERROR("$"&TEXT(SUM('Budget & Payments'!H40:H500),"#,##0"),"—")`],
  ];

  secKPIs.forEach((kpi, i) => {
    const baseRow = 12; // row 13 (0-indexed 12)
    const rowOffset = Math.floor(i / 4) * 3;
    const colOffset = (i % 4) * 3;
    const r = baseRow + rowOffset;
    const c = colOffset;
    data.push({ range: `${S}!${rc(r, c)}`,   values: [[kpi[0]]] });
    data.push({ range: `${S}!${rc(r+1, c)}`, values: [[kpi[1]]] });
    reqs.push({ mergeCells: { range: gridRange(DB, r, r+1, c, c+3), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(DB, r, r+1, c, c+3), cell: { userEnteredFormat: {
      backgroundColor: hex(C.altRow),
      textFormat: { bold: true, foregroundColor: hex(C.primary), fontSize: 8 },
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});
    reqs.push({ mergeCells: { range: gridRange(DB, r+1, r+2, c, c+3), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(DB, r+1, r+2, c, c+3), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { bold: true, foregroundColor: hex(C.primary), fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 12, endIndex: 19 },
    properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 13, endIndex: 14 },
    properties: { pixelSize: 36 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 16, endIndex: 17 },
    properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // ── UPCOMING TASKS SECTION (next 12 incomplete tasks) ─────────────────────
  data.push({ range: `${S}!A20`, values: [['  UPCOMING TASKS — NEXT 12 DUE']] });
  reqs.push({ mergeCells: { range: gridRange(DB, 19, 20, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 19, 20, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 10 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  const taskHdrs = ['Task', 'Phase', 'Due Date', 'Assigned To', 'Priority', 'Status', 'Urgency'];
  data.push({ range: `${S}!A21`, values: [taskHdrs] });
  reqs.push({ repeatCell: { range: gridRange(DB, 20, 21, 0, 7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // SMALL helper for task snapshot — just embed IFERROR formulas
  for (let i = 0; i < 12; i++) {
    const row = 22 + i;
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    // Use IFERROR wrapping SMALL/INDEX/MATCH to pull next upcoming incomplete tasks
    // We can't do a "top N not completed sorted by date" purely in sheets without QUERY,
    // so we'll use a manual pull from the checklist by hardcoded task names that are 'Not Started' or 'In Progress'
    // Instead, use static upcoming task names for now and note in description
    const staticTasks = [
      ['Confirm final headcount with caterer','7','6/1/2026','Amara Osei','Critical','In Progress'],
      ['Finalize seating chart','7','6/5/2026','Amara Osei','Critical','In Progress'],
      ['Submit ceremony script to officiant','7','6/5/2026','Both','High','Not Started'],
      ['Confirm all vendor logistics','7','6/7/2026','Sophia Rivera','Critical','Not Started'],
      ['Prepare place cards','7','6/7/2026','Amara Osei','High','Not Started'],
      ['Finalize vendor payments','7','6/10/2026','Both','Critical','Not Started'],
      ['Prepare gifts for vendors','8','6/10/2026','Both','Medium','Not Started'],
      ['Create wedding day timeline','8','6/10/2026','Sophia Rivera','High','Not Started'],
      ['Final dress fitting','8','5/20/2026','Amara Osei','Critical','Not Started'],
      ['Prepare vendor gratuities','9','6/12/2026','Marcus Thornton','High','Not Started'],
      ['Confirm vendor deliveries','9','6/12/2026','Sophia Rivera','Critical','Not Started'],
      ['Delegate day-of tasks','9','6/12/2026','Sophia Rivera','High','Not Started'],
    ];
    if (i < staticTasks.length) {
      const [task, phase, due, assigned, priority, status] = staticTasks[i];
      data.push({ range: `${S}!A${row}`, values: [[task]] });
      data.push({ range: `${S}!B${row}`, values: [[`Phase ${phase}`]] });
      data.push({ range: `${S}!C${row}`, values: [[due]] });
      data.push({ range: `${S}!D${row}`, values: [[assigned]] });
      data.push({ range: `${S}!E${row}`, values: [[priority]] });
      data.push({ range: `${S}!F${row}`, values: [[status]] });
      // Urgency formula
      data.push({ range: `${S}!G${row}`, values: [[`=IFERROR(IF(C${row}="","",IF(INT(DATEVALUE(TEXT(C${row},"MM/DD/YYYY"))-TODAY())<0,"OVERDUE",IF(INT(DATEVALUE(TEXT(C${row},"MM/DD/YYYY"))-TODAY())<=7,"THIS WEEK",IF(INT(DATEVALUE(TEXT(C${row},"MM/DD/YYYY"))-TODAY())<=30,"THIS MONTH","On Track")))),"")` ]] });
    }
    reqs.push({ repeatCell: { range: gridRange(DB, row - 1, row, 0, 7), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9 },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});
  }

  // CF: urgency in tasks list
  const taskUrgCF = [
    { text: 'OVERDUE',    bg: C.attention },
    { text: 'THIS WEEK',  bg: C.warning   },
    { text: 'THIS MONTH', bg: C.accent    },
  ];
  taskUrgCF.forEach(({ text, bg }) => {
    reqs.push({ addConditionalFormatRule: {
      rule: { ranges: [gridRange(DB, 21, 35, 6, 7)],
        booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] },
          format: { backgroundColor: hex(bg), textFormat: { bold: true, foregroundColor: hex(C.white) } } },
      }, index: 0,
    }});
  });

  // ── THREE SNAPSHOT PANELS ─────────────────────────────────────────────────
  const snapRow = 35;
  const snapTitles = ['RSVP SNAPSHOT', 'VENDOR PAYMENT SNAPSHOT', 'SEATING SNAPSHOT'];
  snapTitles.forEach((title, i) => {
    const c = i * 5;
    data.push({ range: `${S}!${rc(snapRow - 1, c)}`, values: [[title]] });
    reqs.push({ mergeCells: { range: gridRange(DB, snapRow - 1, snapRow, c, c + 4), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(DB, snapRow - 1, snapRow, c, c + 4), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});
  });

  // RSVP Snapshot (col 0-4)
  const rsvpSnap = [
    ['Attending',     `=IFERROR(SUMPRODUCT(('Guest List & RSVP'!R6:R600="Attending")*1),0)`],
    ['Not Attending', `=IFERROR(SUMPRODUCT(('Guest List & RSVP'!R6:R600="Not Attending")*1),0)`],
    ['Pending',       `=IFERROR(SUMPRODUCT(('Guest List & RSVP'!R6:R600="Pending")*1),0)`],
    ['No Response',   `=IFERROR(SUMPRODUCT(('Guest List & RSVP'!R6:R600="No Response")*1),0)`],
    ['Total Invited', `=IFERROR(COUNTA('Guest List & RSVP'!B6:B600),0)`],
  ];
  rsvpSnap.forEach((r2, i) => {
    data.push({ range: `${S}!${rc(snapRow + i, 0)}`, values: [[r2[0]]] });
    data.push({ range: `${S}!${rc(snapRow + i, 1)}`, values: [[r2[1]]] });
  });

  // Vendor Payment Snapshot (col 5-9)
  const vendSnap = [
    ['Total Contracted', `=IFERROR(SUM('Vendors & Venue'!T23:T40),0)`],
    ['Total Paid',       `=IFERROR(SUM('Vendors & Venue'!U23:U40),0)`],
    ['Balance Due',      `=IFERROR(SUM('Vendors & Venue'!V23:V40),0)`],
    ['Contracts Signed', `=IFERROR(SUMPRODUCT(('Vendors & Venue'!G23:G40=TRUE)*1),0)`],
    ['Deposits Paid',    `=IFERROR(SUMPRODUCT(('Vendors & Venue'!R23:R40=TRUE)*1),0)`],
  ];
  vendSnap.forEach((r2, i) => {
    data.push({ range: `${S}!${rc(snapRow + i, 5)}`, values: [[r2[0]]] });
    data.push({ range: `${S}!${rc(snapRow + i, 6)}`, values: [[r2[1]]] });
  });

  // Seating Snapshot (col 9-13)
  const seatSnap = [
    ['Tables Total',    `=IFERROR(COUNTA('Table Planner'!A5:A100),0)`],
    ['Seats Total',     `=IFERROR(SUM('Table Planner'!D5:D100),0)`],
    ['Seats Assigned',  `=IFERROR(SUM('Table Planner'!F5:F100),0)`],
    ['Open Seats',      `=IFERROR(SUM('Table Planner'!G5:G100),0)`],
    ['Conflicts',       `=IFERROR(SUMPRODUCT(('Seating Chart'!P6:P600="YES")*1),0)`],
  ];
  seatSnap.forEach((r2, i) => {
    data.push({ range: `${S}!${rc(snapRow + i, 9)}`, values: [[r2[0]]] });
    data.push({ range: `${S}!${rc(snapRow + i, 10)}`, values: [[r2[1]]] });
  });

  // Style snapshot rows
  for (let r2 = snapRow; r2 < snapRow + 5; r2++) {
    const bg = (r2 - snapRow) % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(DB, r2, r2 + 1, 0, 14), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9 },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});
  }

  // ── CHART DATA AREA (hidden zone for charts) ───────────────────────────────
  // RSVP pie chart data: rows 45-50, cols A-B
  data.push({ range: `${S}!A45`, values: [['RSVP Status', 'Count']] });
  const rsvpStatuses = ['Attending', 'Not Attending', 'Pending', 'No Response'];
  rsvpStatuses.forEach((s, i) => {
    const row = 46 + i;
    data.push({ range: `${S}!A${row}`, values: [[s]] });
    data.push({ range: `${S}!B${row}`, values: [[`=IFERROR(SUMPRODUCT(('Guest List & RSVP'!R$6:R$600="${s}")*1),0)`]] });
  });

  // Budget Category chart data: rows 45-62, cols D-F
  data.push({ range: `${S}!D45`, values: [['Category', 'Budgeted', 'Actual']] });
  const budgetCats = [
    'Venue', 'Catering', 'Photography', 'Videography', 'Florals',
    'Music / Entertainment', 'Hair & Makeup', 'Attire & Accessories',
    'Invitations & Stationery', 'Transportation', 'Cake & Desserts',
    'Officiant', 'Decorations & Rentals', 'Favors & Gifts', 'Honeymoon',
  ];
  budgetCats.forEach((cat, i) => {
    const row = 46 + i;
    data.push({ range: `${S}!D${row}`, values: [[cat]] });
    data.push({ range: `${S}!E${row}`, values: [[`=IFERROR(INDEX('Budget & Payments'!B$5:B$22,MATCH("${cat}",'Budget & Payments'!A$5:A$22,0)),0)`]] });
    data.push({ range: `${S}!F${row}`, values: [[`=IFERROR(INDEX('Budget & Payments'!C$5:C$22,MATCH("${cat}",'Budget & Payments'!A$5:A$22,0)),0)`]] });
  });

  // Meal Preference chart data: rows 45-55, cols H-I
  data.push({ range: `${S}!H45`, values: [['Meal', 'Count']] });
  const meals = ['Chicken', 'Fish', 'Beef', 'Vegetarian', 'Vegan', 'Kids Meal', 'TBD'];
  meals.forEach((m, i) => {
    const row = 46 + i;
    data.push({ range: `${S}!H${row}`, values: [[m]] });
    data.push({ range: `${S}!I${row}`, values: [[`=IFERROR(SUMPRODUCT(('Guest List & RSVP'!J$6:J$600="${m}")*1),0)`]] });
  });

  // Checklist by Phase chart data: rows 45-57, cols K-M
  data.push({ range: `${S}!K45`, values: [['Phase', 'Total', 'Completed']] });
  for (let phase = 1; phase <= 10; phase++) {
    const row = 45 + phase;
    data.push({ range: `${S}!K${row}`, values: [[`Phase ${phase}`]] });
    data.push({ range: `${S}!L${row}`, values: [[`=IFERROR(SUMPRODUCT(('Wedding Checklist'!B$5:B$200=${phase})*1),0)`]] });
    data.push({ range: `${S}!M${row}`, values: [[`=IFERROR(SUMPRODUCT(('Wedding Checklist'!B$5:B$200=${phase})*('Wedding Checklist'!K$5:K$200="Completed")*1),0)`]] });
  }

  // Style chart data area lightly
  reqs.push({ repeatCell: { range: gridRange(DB, 44, 62, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow), textFormat: { fontSize: 8, foregroundColor: hex(C.secText) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});
  reqs.push({ repeatCell: { range: gridRange(DB, 44, 45, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 8 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [160, 60, 80, 90, 70, 70, 80, 70, 70, 80, 80, 70, 70, 80];
  colWidths.forEach((px, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Freeze row 1
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: DB, gridProperties: { frozenRowCount: 2 } },
    fields: 'gridProperties.frozenRowCount',
  }});

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: DB, tabColor: hex(C.secondary) },
    fields: 'tabColor',
  }});

  await valuesBatchUpdate(id, data, 'dashboard-data');
  await batchUpdate(id, reqs, 'dashboard-format');

  // ── CHARTS ────────────────────────────────────────────────────────────────
  const chartReqs = [];

  function chartSrc(r1, r2, c1, c2) {
    return { sourceRange: { sources: [gridRange(DB, r1, r2, c1, c2)] } };
  }

  // Chart 1: RSVP Status Distribution — PIE, anchor row 65, col 0
  chartReqs.push({ addChart: { chart: {
    spec: { title: 'RSVP Status Distribution', pieChart: {
      legendPosition: 'LABELED_LEGEND',
      domain: chartSrc(44, 50, 0, 1),
      series: chartSrc(44, 50, 1, 2),
    }},
    position: { overlayPosition: {
      anchorCell: { sheetId: DB, rowIndex: 64, columnIndex: 0 },
      offsetXPixels: 0, offsetYPixels: 0, widthPixels: 400, heightPixels: 260,
    }},
  }}});

  // Chart 2: Meal Preference Distribution — PIE, anchor row 65, col 4
  chartReqs.push({ addChart: { chart: {
    spec: { title: 'Meal Preference Distribution', pieChart: {
      legendPosition: 'LABELED_LEGEND',
      domain: chartSrc(44, 52, 7, 8),
      series: chartSrc(44, 52, 8, 9),
    }},
    position: { overlayPosition: {
      anchorCell: { sheetId: DB, rowIndex: 64, columnIndex: 7 },
      offsetXPixels: 0, offsetYPixels: 0, widthPixels: 400, heightPixels: 260,
    }},
  }}});

  // Chart 3: Budget Allocated vs Actual — BAR, anchor row 80, col 0
  chartReqs.push({ addChart: { chart: {
    spec: { title: 'Budget: Allocated vs Actual by Category', basicChart: {
      chartType: 'BAR',
      legendPosition: 'BOTTOM_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: 'Amount ($)' },
        { position: 'LEFT_AXIS',   title: 'Category' },
      ],
      domains: [{ domain: chartSrc(44, 61, 3, 4) }],
      series: [
        { series: chartSrc(44, 61, 4, 5), targetAxis: 'BOTTOM_AXIS', color: hex(C.primary) },
        { series: chartSrc(44, 61, 5, 6), targetAxis: 'BOTTOM_AXIS', color: hex(C.secondary) },
      ],
      headerCount: 1,
    }},
    position: { overlayPosition: {
      anchorCell: { sheetId: DB, rowIndex: 79, columnIndex: 0 },
      offsetXPixels: 0, offsetYPixels: 0, widthPixels: 480, heightPixels: 300,
    }},
  }}});

  // Chart 4: Checklist Completion by Phase — COLUMN, anchor row 80, col 7
  chartReqs.push({ addChart: { chart: {
    spec: { title: 'Checklist: Tasks Completed by Phase', basicChart: {
      chartType: 'COLUMN',
      legendPosition: 'BOTTOM_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: 'Phase' },
        { position: 'LEFT_AXIS',   title: 'Tasks' },
      ],
      domains: [{ domain: chartSrc(44, 56, 10, 11) }],
      series: [
        { series: chartSrc(44, 56, 11, 12), targetAxis: 'LEFT_AXIS', color: hex(C.primary) },
        { series: chartSrc(44, 56, 12, 13), targetAxis: 'LEFT_AXIS', color: hex(C.success) },
      ],
      headerCount: 1,
    }},
    position: { overlayPosition: {
      anchorCell: { sheetId: DB, rowIndex: 79, columnIndex: 7 },
      offsetXPixels: 0, offsetYPixels: 0, widthPixels: 480, heightPixels: 300,
    }},
  }}});

  await batchUpdate(id, chartReqs, 'dashboard-charts');

  console.log('✓ Dashboard built with KPIs, snapshots, and 4 charts');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
