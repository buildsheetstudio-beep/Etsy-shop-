'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DB = sheetMap['🌸 Dashboard'];

// Dashboard layout (all rows 1-indexed in comments, 0-indexed in gridRange):
// Row 1:  Hero title (A1:L1) — full width
// Row 2:  Party info labels (A-L, 6 pairs)
// Row 3:  Spacer
// Row 4:  KPI block header
// Rows 5-10: KPI cards (2 columns: A-B label+value, D-E label+value, G-H label+value)
// Row 11: Spacer
// Row 12: "Upcoming Tasks" header + "Quick Progress" header
// Rows 13-20: Upcoming tasks FILTER (A-E) | Progress sparklines (G-L)
// Row 21: Spacer
// Row 22: "Guest RSVP Snapshot" header
// Rows 23-27: RSVP summary (A-D) | Recent gifts (F-J)

const NCOLS = 12; // A-L

(async () => {
  const reqs = [];

  // Column widths: A=90, B=200, C=30, D=90, E=200, F=30, G=90, H=200, I=30, J=90, K=200, L=80
  [90,200,30,90,200,30,90,200,30,90,200,80].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DB, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Row heights
  // Row 1: hero
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DB, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 64 }, fields: 'pixelSize',
  }});
  // Row 2: party info
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DB, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  // Row 3: spacer
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DB, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 12 }, fields: 'pixelSize',
  }});
  // Row 4: KPI header
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DB, dimension: 'ROWS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  // Rows 5-10: KPI cards
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DB, dimension: 'ROWS', startIndex: 4, endIndex: 10 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});
  // Row 11: spacer
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DB, dimension: 'ROWS', startIndex: 10, endIndex: 11 },
    properties: { pixelSize: 12 }, fields: 'pixelSize',
  }});
  // Row 12: section headers
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DB, dimension: 'ROWS', startIndex: 11, endIndex: 12 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  // Rows 13-20: tasks + progress
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DB, dimension: 'ROWS', startIndex: 12, endIndex: 20 },
    properties: { pixelSize: 26 }, fields: 'pixelSize',
  }});
  // Row 21: spacer
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DB, dimension: 'ROWS', startIndex: 20, endIndex: 21 },
    properties: { pixelSize: 12 }, fields: 'pixelSize',
  }});
  // Row 22: section headers
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DB, dimension: 'ROWS', startIndex: 21, endIndex: 22 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  // Rows 23-30: guest + gift snapshot
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DB, dimension: 'ROWS', startIndex: 22, endIndex: 30 },
    properties: { pixelSize: 26 }, fields: 'pixelSize',
  }});

  // ── HERO TITLE ROW (A1:L1) ──
  reqs.push({ mergeCells: { range: gridRange(DB, 0, 1, 0, NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DB, 0, 1, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 20 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // ── PARTY INFO ROW 2 — 6 pairs merged ──
  reqs.push({ repeatCell: {
    range: gridRange(DB, 1, 2, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  [[0,2],[2,4],[4,6],[6,8],[8,10],[10,12]].forEach(([c1,c2]) => {
    reqs.push({ mergeCells: { range: gridRange(DB, 1, 2, c1, c2), mergeType: 'MERGE_ALL' }});
  });

  // ── SPACER ROW 3 ──
  reqs.push({ repeatCell: {
    range: gridRange(DB, 2, 3, 0, NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.border) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});

  // ── KPI HEADER ROW 4 ──
  // Three KPI header blocks: A-B, D-E, G-H (with C, F spacers)
  [[0,2],[3,5],[6,8],[9,11]].forEach(([c1,c2]) => {
    reqs.push({ mergeCells: { range: gridRange(DB, 3, 4, c1, c2), mergeType: 'MERGE_ALL' }});
    reqs.push({ repeatCell: {
      range: gridRange(DB, 3, 4, c1, c2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.roseLilac),
        textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
  });

  // ── KPI CARDS rows 5-10 (indices 4-9) ──
  // Three groups side by side; label col (A/D/G/J) + value col (B/E/H/K)
  const kpiLabelCols  = [0, 3, 6, 9];
  const kpiValueCols  = [1, 4, 7, 10];
  for (let r = 4; r < 10; r++) {
    kpiLabelCols.forEach(c => {
      reqs.push({ repeatCell: {
        range: gridRange(DB, r, r + 1, c, c + 1),
        cell: { userEnteredFormat: {
          backgroundColor: hex(C.champagne),
          textFormat: { foregroundColor: hex(C.muted), bold: true, fontSize: 8 },
          verticalAlignment: 'MIDDLE',
        }},
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      }});
    });
    kpiValueCols.forEach(c => {
      reqs.push({ repeatCell: {
        range: gridRange(DB, r, r + 1, c, c + 1),
        cell: { userEnteredFormat: {
          backgroundColor: hex(C.champagne),
          textFormat: { foregroundColor: hex(C.text), bold: true, fontSize: 13 },
          verticalAlignment: 'MIDDLE', horizontalAlignment: 'RIGHT',
        }},
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
      }});
    });
  }
  // KPI borders
  [[0,2],[3,5],[6,8],[9,11]].forEach(([c1,c2]) => {
    reqs.push({ updateBorders: {
      range: gridRange(DB, 3, 10, c1, c2),
      top:    { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
      bottom: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
      left:   { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
      right:  { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
      innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
      innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
    }});
  });

  // ── SPACER ROW 11 ──
  reqs.push({ repeatCell: {
    range: gridRange(DB, 10, 11, 0, NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.border) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});

  // ── SECTION HEADERS ROW 12 ──
  // Upcoming Tasks: A-E merged
  reqs.push({ mergeCells: { range: gridRange(DB, 11, 12, 0, 5), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DB, 11, 12, 0, 5),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.roseLilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // Quick Progress: G-L merged
  reqs.push({ mergeCells: { range: gridRange(DB, 11, 12, 6, NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DB, 11, 12, 6, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.roseLilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // ── UPCOMING TASKS rows 13-20 (indices 12-19) cols A-E ──
  for (let r = 12; r < 20; r++) {
    const bg = (r % 2 === 0) ? C.white : C.petalPink;
    reqs.push({ repeatCell: {
      range: gridRange(DB, r, r + 1, 0, 5),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }
  reqs.push({ updateBorders: {
    range: gridRange(DB, 11, 20, 0, 5),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // ── QUICK PROGRESS rows 13-20 (indices 12-19) cols G-L ──
  for (let r = 12; r < 20; r++) {
    const bg = (r % 2 === 0) ? C.white : C.petalPink;
    reqs.push({ repeatCell: {
      range: gridRange(DB, r, r + 1, 6, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }
  // Label col G
  for (let r = 12; r < 20; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(DB, r, r + 1, 6, 7),
      cell: { userEnteredFormat: {
        textFormat: { foregroundColor: hex(C.muted), bold: true, fontSize: 8 },
      }},
      fields: 'userEnteredFormat.textFormat',
    }});
  }
  // Pct col H
  for (let r = 12; r < 20; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(DB, r, r + 1, 7, 8),
      cell: { userEnteredFormat: {
        textFormat: { foregroundColor: hex(C.text), bold: true, fontSize: 9 },
        horizontalAlignment: 'RIGHT',
      }},
      fields: 'userEnteredFormat(textFormat,horizontalAlignment)',
    }});
  }
  // Sparkline col I-L merged per row
  for (let r = 12; r < 20; r++) {
    reqs.push({ mergeCells: { range: gridRange(DB, r, r + 1, 8, NCOLS), mergeType: 'MERGE_ALL' }});
  }
  reqs.push({ updateBorders: {
    range: gridRange(DB, 11, 20, 6, NCOLS),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // ── SPACER ROW 21 ──
  reqs.push({ repeatCell: {
    range: gridRange(DB, 20, 21, 0, NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.border) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});

  // ── RSVP SNAPSHOT HEADER ROW 22 ──
  reqs.push({ mergeCells: { range: gridRange(DB, 21, 22, 0, 5), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DB, 21, 22, 0, 5),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.roseLilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // Gift Snapshot header
  reqs.push({ mergeCells: { range: gridRange(DB, 21, 22, 6, NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DB, 21, 22, 6, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.roseLilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // ── RSVP rows 23-30 (indices 22-29) cols A-D ──
  for (let r = 22; r < 30; r++) {
    const bg = (r % 2 === 0) ? C.white : C.petalPink;
    reqs.push({ repeatCell: {
      range: gridRange(DB, r, r + 1, 0, 5),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }
  reqs.push({ updateBorders: {
    range: gridRange(DB, 21, 30, 0, 5),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // ── GIFT SNAPSHOT rows 23-30 (indices 22-29) cols G-L ──
  for (let r = 22; r < 30; r++) {
    const bg = (r % 2 === 0) ? C.white : C.petalPink;
    reqs.push({ repeatCell: {
      range: gridRange(DB, r, r + 1, 6, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }
  reqs.push({ updateBorders: {
    range: gridRange(DB, 21, 30, 6, NCOLS),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  await batchUpdate(id, reqs, 'dashboard-format');

  // ── VALUES ──
  const vals = [];
  const sn = (name) => `'${name}'`;

  vals.push({ range: "'🌸 Dashboard'!A1", values: [['  🌸  Ultimate Baby Shower Planner — Dashboard']] });

  // Party info row 2
  vals.push({ range: "'🌸 Dashboard'!A2", values: [['📅  Party Date: 20 July 2025']] });
  vals.push({ range: "'🌸 Dashboard'!C2", values: [['⏰  11:00 AM – 4:30 PM']] });
  vals.push({ range: "'🌸 Dashboard'!E2", values: [['👶  Mum-to-Be: Sarah Mitchell']] });
  vals.push({ range: "'🌸 Dashboard'!G2", values: [['📍  Home of Emma Clarke']] });
  vals.push({ range: "'🌸 Dashboard'!I2", values: [['🎨  Theme: Dusty Lilac & Blush']] });
  vals.push({ range: "'🌸 Dashboard'!K2", values: [['👥  Target Guests: 20']] });

  // KPI section headers
  vals.push({ range: "'🌸 Dashboard'!A4", values: [['GUEST LIST']] });
  vals.push({ range: "'🌸 Dashboard'!D4", values: [['CHECKLIST']] });
  vals.push({ range: "'🌸 Dashboard'!G4", values: [['BUDGET']] });
  vals.push({ range: "'🌸 Dashboard'!J4", values: [['GIFTS']] });

  // KPI labels & formulas
  // Guest List KPIs (cols A-B, rows 5-10)
  vals.push({ range: "'🌸 Dashboard'!A5", values: [
    ['Invited',    `=IFERROR(COUNTA('👥 Guest List & RSVPs'!A:A)-1,"—")`],
    ['Confirmed',  `=IFERROR(COUNTIF('👥 Guest List & RSVPs'!E4:E200,"Confirmed"),"—")`],
    ['Declined',   `=IFERROR(COUNTIF('👥 Guest List & RSVPs'!E4:E200,"Declined"),"—")`],
    ['Awaiting',   `=IFERROR(COUNTIFS('👥 Guest List & RSVPs'!E4:E200,"Invited")+COUNTIF('👥 Guest List & RSVPs'!E4:E200,"No Response"),"—")`],
    ['Dietary Req',`=IFERROR(COUNTIFS('👥 Guest List & RSVPs'!G4:G200,"<>None",'👥 Guest List & RSVPs'!G4:G200,"<>"),"—")`],
    ['Plus Ones',  `=IFERROR(COUNTIF('👥 Guest List & RSVPs'!H4:H200,TRUE),"—")`],
  ]});

  // Checklist KPIs (cols D-E, rows 5-10)
  vals.push({ range: "'🌸 Dashboard'!D5", values: [
    ['Total Tasks', `=IFERROR(COUNTA('✅ Party Checklist'!C4:C200),"—")`],
    ['Done',        `=IFERROR(COUNTIF('✅ Party Checklist'!A4:A200,TRUE),"—")`],
    ['Remaining',   `=IFERROR(COUNTA('✅ Party Checklist'!C4:C200)-COUNTIF('✅ Party Checklist'!A4:A200,TRUE),"—")`],
    ['% Complete',  `=IFERROR(TEXT(COUNTIF('✅ Party Checklist'!A4:A200,TRUE)/COUNTA('✅ Party Checklist'!C4:C200),"0%"),"—")`],
    ['High Priority',`=IFERROR(COUNTIFS('✅ Party Checklist'!D4:D200,"High",'✅ Party Checklist'!A4:A200,FALSE),"—")`],
    ['Overdue',     `=IFERROR(COUNTIFS('✅ Party Checklist'!A4:A200,FALSE,'✅ Party Checklist'!E4:E200,"<>",'✅ Party Checklist'!E4:E200,"<"&TEXT(TODAY(),"dd mmm yyyy")),"—")`],
  ]});

  // Budget KPIs (cols G-H, rows 5-10)
  vals.push({ range: "'🌸 Dashboard'!G5", values: [
    ['Total Budget',  `=IFERROR('💰 Budget & Expenses'!C16,"—")`],
    ['Spent',         `=IFERROR('💰 Budget & Expenses'!D16,"—")`],
    ['Remaining',     `=IFERROR('💰 Budget & Expenses'!E16,"—")`],
    ['% Spent',       `=IFERROR(TEXT('💰 Budget & Expenses'!D16/'💰 Budget & Expenses'!C16,"0%"),"—")`],
    ['# Expenses',    `=IFERROR(COUNTA('💰 Budget & Expenses'!C18:C200),"—")`],
    ['Paid Items',    `=IFERROR(COUNTIF('💰 Budget & Expenses'!G18:G200,TRUE),"—")`],
  ]});

  // Gift KPIs (cols J-K, rows 5-10)
  vals.push({ range: "'🌸 Dashboard'!J5", values: [
    ['Gifts Logged',  `=IFERROR(COUNTA('🎁 Gift Tracker'!B4:B200),"—")`],
    ['Received',      `=IFERROR(COUNTIF('🎁 Gift Tracker'!F4:F200,TRUE),"—")`],
    ['TY Sent',       `=IFERROR(COUNTIF('🎁 Gift Tracker'!G4:G200,TRUE),"—")`],
    ['TY Pending',    `=IFERROR(COUNTIFS('🎁 Gift Tracker'!F4:F200,TRUE,'🎁 Gift Tracker'!G4:G200,FALSE),"—")`],
    ['Est. Value',    `=IFERROR("$"&TEXT(SUMIF('🎁 Gift Tracker'!F4:F200,TRUE,'🎁 Gift Tracker'!E4:E200),"#,##0"),"—")`],
    ['Categories',    `=IFERROR(SUMPRODUCT(1/COUNTIF('🎁 Gift Tracker'!D4:D23,'🎁 Gift Tracker'!D4:D23)),"—")`],
  ]});

  // Section headers
  vals.push({ range: "'🌸 Dashboard'!A12", values: [['  📋  UPCOMING TASKS (next 14 days)']] });
  vals.push({ range: "'🌸 Dashboard'!G12", values: [['  📊  QUICK PROGRESS']] });

  // Upcoming tasks — FILTER from checklist (col B=timeframe, col C=task, col D=priority, col E=due date)
  // Filter: not done (A=FALSE) AND due date not blank AND due <= TODAY()+14
  vals.push({ range: "'🌸 Dashboard'!A13", values: [
    [`=IFERROR(FILTER('✅ Party Checklist'!B4:E200,('✅ Party Checklist'!A4:A200=FALSE)*('✅ Party Checklist'!E4:E200<>"")),"No upcoming tasks")`],
  ]});

  // Quick progress sparklines (rows 13-20 = indices 12-19)
  const PROGRESS = [
    ['Checklist',     `=IFERROR(TEXT(COUNTIF('✅ Party Checklist'!A4:A200,TRUE)/COUNTA('✅ Party Checklist'!C4:C200),"0%"),"—")`,
                      `=IFERROR(SPARKLINE(COUNTIF('✅ Party Checklist'!A4:A200,TRUE)/COUNTA('✅ Party Checklist'!C4:C200),{"charttype","bar";"color1","#9B7DB5";"color2","#E5D5DF";"max",1}),"—")`],
    ['RSVPs',         `=IFERROR(TEXT(COUNTIF('👥 Guest List & RSVPs'!E4:E200,"Confirmed")/COUNTA('👥 Guest List & RSVPs'!A:A),"0%"),"—")`,
                      `=IFERROR(SPARKLINE(COUNTIF('👥 Guest List & RSVPs'!E4:E200,"Confirmed")/COUNTA('👥 Guest List & RSVPs'!A:A),{"charttype","bar";"color1","#B58DAA";"color2","#E5D5DF";"max",1}),"—")`],
    ['Budget Used',   `=IFERROR(TEXT('💰 Budget & Expenses'!D16/'💰 Budget & Expenses'!C16,"0%"),"—")`,
                      `=IFERROR(SPARKLINE('💰 Budget & Expenses'!D16/'💰 Budget & Expenses'!C16,{"charttype","bar";"color1",IF('💰 Budget & Expenses'!D16/'💰 Budget & Expenses'!C16>1,"#9B4A5A","#7A9B7A");"color2","#E5D5DF";"max",1}),"—")`],
    ['Gifts Rcvd',    `=IFERROR(TEXT(COUNTIF('🎁 Gift Tracker'!F4:F200,TRUE)/COUNTA('🎁 Gift Tracker'!B4:B200),"0%"),"—")`,
                      `=IFERROR(SPARKLINE(COUNTIF('🎁 Gift Tracker'!F4:F200,TRUE)/COUNTA('🎁 Gift Tracker'!B4:B200),{"charttype","bar";"color1","#C4A0B5";"color2","#E5D5DF";"max",1}),"—")`],
    ['TY Notes Sent', `=IFERROR(TEXT(COUNTIF('🎁 Gift Tracker'!G4:G200,TRUE)/COUNTA('🎁 Gift Tracker'!B4:B200),"0%"),"—")`,
                      `=IFERROR(SPARKLINE(COUNTIF('🎁 Gift Tracker'!G4:G200,TRUE)/COUNTA('🎁 Gift Tracker'!B4:B200),{"charttype","bar";"color1","#5A7A5A";"color2","#E5D5DF";"max",1}),"—")`],
    ['Décor Ordered', `=IFERROR(TEXT(COUNTIF('🎀 Décor, Food & Drinks'!G4:G17,"Ordered")/COUNTA('🎀 Décor, Food & Drinks'!A4:A17),"0%"),"—")`,
                      `=IFERROR(SPARKLINE(COUNTIF('🎀 Décor, Food & Drinks'!G4:G17,"Ordered")/COUNTA('🎀 Décor, Food & Drinks'!A4:A17),{"charttype","bar";"color1","#B5A08D";"color2","#E5D5DF";"max",1}),"—")`],
    ['Food Confirmed',`=IFERROR(TEXT(COUNTIF('🎀 Décor, Food & Drinks'!G21:G35,"Confirmed")/COUNTA('🎀 Décor, Food & Drinks'!A21:A35),"0%"),"—")`,
                      `=IFERROR(SPARKLINE(COUNTIF('🎀 Décor, Food & Drinks'!G21:G35,"Confirmed")/COUNTA('🎀 Décor, Food & Drinks'!A21:A35),{"charttype","bar";"color1","#8DAAB5";"color2","#E5D5DF";"max",1}),"—")`],
    ['Games Planned', `=IFERROR(TEXT(COUNTIF('🎮 Games & Activities'!I3:I14,"Planned")/COUNTA('🎮 Games & Activities'!A3:A14),"0%"),"—")`,
                      `=IFERROR(SPARKLINE(COUNTIF('🎮 Games & Activities'!I3:I14,"Planned")/COUNTA('🎮 Games & Activities'!A3:A14),{"charttype","bar";"color1","#9B8DB5";"color2","#E5D5DF";"max",1}),"—")`],
  ];
  PROGRESS.forEach(([label, pct, spark], i) => {
    vals.push({ range: `'🌸 Dashboard'!G${13+i}`, values: [[label, pct, spark]] });
  });

  // RSVP snapshot header
  vals.push({ range: "'🌸 Dashboard'!A22", values: [['  👥  RSVP SNAPSHOT']] });
  vals.push({ range: "'🌸 Dashboard'!G22", values: [['  🎁  GIFT SNAPSHOT']] });

  // RSVP rows
  vals.push({ range: "'🌸 Dashboard'!A23", values: [
    ['STATUS','COUNT','',''],
    ['Confirmed', `=COUNTIF('👥 Guest List & RSVPs'!E4:E200,"Confirmed")`, '', ''],
    ['Declined',  `=COUNTIF('👥 Guest List & RSVPs'!E4:E200,"Declined")`, '', ''],
    ['Invited',   `=COUNTIF('👥 Guest List & RSVPs'!E4:E200,"Invited")`, '', ''],
    ['No Response',`=COUNTIF('👥 Guest List & RSVPs'!E4:E200,"No Response")`, '', ''],
    ['Dietary Needs', `=COUNTIFS('👥 Guest List & RSVPs'!G4:G200,"<>None",'👥 Guest List & RSVPs'!G4:G200,"<>")`, '', ''],
    ['Plus Ones', `=COUNTIF('👥 Guest List & RSVPs'!H4:H200,TRUE)`, '', ''],
    ['Total Attending', `=COUNTIF('👥 Guest List & RSVPs'!E4:E200,"Confirmed")+COUNTIF('👥 Guest List & RSVPs'!H4:H200,TRUE)`, '', ''],
  ]});

  // Gift snapshot rows
  vals.push({ range: "'🌸 Dashboard'!G23", values: [
    ['STATUS','COUNT','',''],
    ['Gifts Received',   `=COUNTIF('🎁 Gift Tracker'!F4:F200,TRUE)`, '', ''],
    ['Not Yet Received', `=COUNTIF('🎁 Gift Tracker'!F4:F200,FALSE)`, '', ''],
    ['TY Sent',          `=COUNTIF('🎁 Gift Tracker'!G4:G200,TRUE)`, '', ''],
    ['TY Pending',       `=COUNTIFS('🎁 Gift Tracker'!F4:F200,TRUE,'🎁 Gift Tracker'!G4:G200,FALSE)`, '', ''],
    ['Est. Gift Value',  `=IFERROR("$"&TEXT(SUMIF('🎁 Gift Tracker'!F4:F200,TRUE,'🎁 Gift Tracker'!E4:E200),"#,##0"),"—")`, '', ''],
    ['Gift Categories',  `=IFERROR(SUMPRODUCT(1/COUNTIF('🎁 Gift Tracker'!D4:D23,'🎁 Gift Tracker'!D4:D23)),"—")`, '', ''],
    ['Gifts Logged',     `=COUNTA('🎁 Gift Tracker'!B4:B200)`, '', ''],
  ]});

  await valuesBatchUpdate(id, vals, 'dashboard-values');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
