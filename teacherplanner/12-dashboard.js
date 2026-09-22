'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, sheets } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Teacher Dashboard'];
const S = "'Teacher Dashboard'";
const ROSTER = "'Student Roster'";
const ASSIGN = "'Assignments & Assessments'";
const GRADE  = "'Gradebook'";
const ATT    = "'Attendance'";
const SETUP  = "'School Year Setup'";
const TODO   = "'Teacher To-Do'";

// Dashboard layout (18 cols × 120 rows):
// Row 1-2:   Title band
// Row 3:     Subtitle / last updated
// Row 4:     Spacer
// Row 5-8:   KPI cards row (4 KPI cards across)
// Row 9:     Spacer
// Row 10-11: Section header "Class Snapshot"
// Row 12-16: Class snapshot table (5 classes + header)
// Row 17:    Spacer
// Row 18:    Section header "Gradebook Summary"
// Row 19-25: Grade summary table (by grade category)
// Row 26:    Spacer
// Row 27:    Section header "Attendance Summary"
// Row 28-34: Attendance table (by class)
// Row 35:    Spacer
// Row 36:    Section header "To-Do Summary"
// Row 37-43: To-Do summary by category
// Row 44:    Spacer
// Row 45:    Section header "Assignment Status"
// Row 46-52: Assignment pipeline (by status)
// Row 53-60: Spacer / chart anchor rows
// Row 61:    Disclaimer

// Helper lookup tables for charts (non-adjacent data → helper table)
// Placed cols J-N (indices 9-13), rows 5-50

// IMPORTANT: All cross-sheet formulas must use IFERROR
// SUMPRODUCT instead of COUNTIFS with YEAR/MONTH functions

(async () => {
  const vals = [];
  const fmt  = [];

  // ── BG ──────────────────────────────────────────────────────────────
  fmt.push({ repeatCell:{ range: gridRange(SID,0,120,0,18), cell:{userEnteredFormat:{backgroundColor:hex(C.bg)}}, fields:'userEnteredFormat.backgroundColor' }});

  // ── TITLE BAND ──────────────────────────────────────────────────────
  vals.push({ range:`${S}!A1`, values:[['TEACHER DASHBOARD']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,9), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,9), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3437'), textFormat:{bold:true,fontSize:18,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:12},
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1}, properties:{pixelSize:50}, fields:'pixelSize' }});

  vals.push({ range:`${S}!A2`, values:[['Sarah Mitchell | Maplewood Elementary | Grade 4 | 2025–26']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,1,2,0,9), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,9), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:10,foregroundColor:hex('#D5D8DB'),fontFamily:'Arial'},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:12},
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:1,endIndex:2}, properties:{pixelSize:28}, fields:'pixelSize' }});

  vals.push({ range:`${S}!A3`, values:[[`=IFERROR("Last updated: "&TEXT(NOW(),"mmm d, yyyy h:mm AM/PM"),"")`]] });
  fmt.push({ mergeCells:{ range: gridRange(SID,2,3,0,9), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,2,3,0,9), cell:{userEnteredFormat:{
    backgroundColor:hex(C.bg), textFormat:{fontSize:8,foregroundColor:hex(C.secText),italic:true,fontFamily:'Arial'},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:12},
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:2,endIndex:3}, properties:{pixelSize:20}, fields:'pixelSize' }});

  // Spacer row 4
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:3,endIndex:4}, properties:{pixelSize:8}, fields:'pixelSize' }});

  // ── KPI CARDS (rows 5-8) ────────────────────────────────────────────
  // 4 KPI cards, each 2 cols wide: A-B, C-D, E-F, G-H (cols 0-7)
  const KPI_CARDS = [
    {
      label: 'Total Students',
      formula: `=IFERROR(COUNTA(${ROSTER}!B6:B1005),"–")`,
      sub: 'Active roster',
      color: C.ELA,
    },
    {
      label: 'Assignments Graded',
      formula: `=IFERROR(COUNTIF(${ASSIGN}!O6:O1505,"Returned"),"–")`,
      sub: 'Status: Returned',
      color: C.Math,
    },
    {
      label: 'Class Avg Score',
      formula: `=IFERROR(TEXT(AVERAGEIF(${GRADE}!Q6:Q5005,TRUE,${GRADE}!M6:M5005),"0%"),"–")`,
      sub: 'All graded records',
      color: C.Science,
    },
    {
      label: 'Tasks Outstanding',
      formula: `=IFERROR(COUNTIFS(${TODO}!E6:E1505,"<>Complete",${TODO}!E6:E1505,"<>"),"–")`,
      sub: 'Not yet complete',
      color: C.SocialStudies,
    },
  ];

  const KPI_COLS = [[0,2],[2,4],[4,6],[6,8]];
  KPI_CARDS.forEach((kpi, ki) => {
    const [c1, c2] = KPI_COLS[ki];
    // Label row (row 5, index 4)
    fmt.push({ mergeCells:{ range: gridRange(SID,4,5,c1,c2), mergeType:'MERGE_ALL' }});
    fmt.push({ repeatCell:{ range: gridRange(SID,4,5,c1,c2), cell:{userEnteredFormat:{
      backgroundColor:hex(kpi.color), textFormat:{bold:false,fontSize:9,foregroundColor:hex(C.text),fontFamily:'Arial'},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
    vals.push({ range:`${S}!${String.fromCharCode(65+c1)}5`, values:[[kpi.label]] });

    // Value row (rows 6-7, index 5-6)
    fmt.push({ mergeCells:{ range: gridRange(SID,5,7,c1,c2), mergeType:'MERGE_ALL' }});
    fmt.push({ repeatCell:{ range: gridRange(SID,5,7,c1,c2), cell:{userEnteredFormat:{
      backgroundColor:hex(kpi.color), textFormat:{bold:true,fontSize:26,foregroundColor:hex(C.text),fontFamily:'Arial'},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
    vals.push({ range:`${S}!${String.fromCharCode(65+c1)}6`, values:[[kpi.formula]] });

    // Sub-label row (row 8, index 7)
    fmt.push({ mergeCells:{ range: gridRange(SID,7,8,c1,c2), mergeType:'MERGE_ALL' }});
    fmt.push({ repeatCell:{ range: gridRange(SID,7,8,c1,c2), cell:{userEnteredFormat:{
      backgroundColor:hex(kpi.color), textFormat:{fontSize:8,foregroundColor:hex(C.text),fontFamily:'Arial'},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
    vals.push({ range:`${S}!${String.fromCharCode(65+c1)}8`, values:[[kpi.sub]] });
  });

  [4,5,6,7].forEach((r, ri) => {
    const px = [22,38,38,22][ri];
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:r,endIndex:r+1}, properties:{pixelSize:px}, fields:'pixelSize' }});
  });

  // Spacer row 9
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:8,endIndex:9}, properties:{pixelSize:10}, fields:'pixelSize' }});

  // ── SECTION HELPER ───────────────────────────────────────────────────
  function sectionHeader(rowIdx, label, color) {
    fmt.push({ mergeCells:{ range: gridRange(SID,rowIdx,rowIdx+1,0,9), mergeType:'MERGE_ALL' }});
    fmt.push({ repeatCell:{ range: gridRange(SID,rowIdx,rowIdx+1,0,9), cell:{userEnteredFormat:{
      backgroundColor:hex(color), textFormat:{bold:true,fontSize:10,foregroundColor:hex(C.text),fontFamily:'Arial'},
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:8},
      borders:{ bottom:{style:'SOLID',color:hex(C.border)} },
    }}, fields:'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:rowIdx,endIndex:rowIdx+1}, properties:{pixelSize:26}, fields:'pixelSize' }});
    vals.push({ range:`${S}!A${rowIdx+1}`, values:[[label]] });
  }

  function tableHeader(rowIdx, headers, endCol) {
    vals.push({ range:`${S}!A${rowIdx+1}`, values:[headers] });
    fmt.push({ repeatCell:{ range: gridRange(SID,rowIdx,rowIdx+1,0,endCol), cell:{userEnteredFormat:{
      backgroundColor:hex(C.altRow), textFormat:{bold:true,fontSize:8,foregroundColor:hex(C.secText),fontFamily:'Arial'},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
      borders:{ bottom:{style:'SOLID',color:hex(C.border)} },
    }}, fields:'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:rowIdx,endIndex:rowIdx+1}, properties:{pixelSize:22}, fields:'pixelSize' }});
  }

  function tableRow(rowIdx, rowData, endCol, bgColor) {
    vals.push({ range:`${S}!A${rowIdx+1}`, values:[rowData] });
    fmt.push({ repeatCell:{ range: gridRange(SID,rowIdx,rowIdx+1,0,endCol), cell:{userEnteredFormat:{
      backgroundColor:hex(bgColor || C.panel), textFormat:{fontSize:9,foregroundColor:hex(C.text),fontFamily:'Arial'},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
      borders:{ bottom:{style:'DOTTED',color:hex(C.border)} },
    }}, fields:'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:rowIdx,endIndex:rowIdx+1}, properties:{pixelSize:22}, fields:'pixelSize' }});
  }

  function spacer(rowIdx) {
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:rowIdx,endIndex:rowIdx+1}, properties:{pixelSize:10}, fields:'pixelSize' }});
  }

  // ── CLASS SNAPSHOT (rows 10-16) ──────────────────────────────────────
  sectionHeader(9, '📋  Class Snapshot', C.ELA);
  tableHeader(10, ['Class ID','Class Name','Enrolled','Assignments','Avg Class Score','Missing Work','Absent (last 22 days)','Avg Attendance %'], 8);

  const CLASS_SNAP = [
    ['CLS-001','ELA'],
    ['CLS-002','Reading'],
    ['CLS-003','Mathematics'],
    ['CLS-004','Science'],
    ['CLS-005','Social Studies'],
  ];
  CLASS_SNAP.forEach(([clsId, name], ci) => {
    const r = 11 + ci;
    const bg = ci % 2 === 0 ? C.panel : C.altRow;
    tableRow(r, [
      clsId,
      name,
      `=IFERROR(COUNTIF(${ROSTER}!E6:E1005,"${clsId}"),"–")`,
      `=IFERROR(COUNTIF(${ASSIGN}!B6:B1505,"${clsId}"),"–")`,
      `=IFERROR(TEXT(AVERAGEIFS(${GRADE}!M6:M5005,${GRADE}!D6:D5005,"${clsId}",${GRADE}!Q6:Q5005,TRUE),"0%"),"–")`,
      `=IFERROR(COUNTIFS(${GRADE}!D6:D5005,"${clsId}",${GRADE}!N6:N5005,TRUE),"–")`,
      `=IFERROR(COUNTIFS(${ATT}!D6:D5005,"${clsId}",${ATT}!H6:H5005,"Absent"),"–")`,
      `=IFERROR(TEXT(SUMPRODUCT((${ATT}!D6:D5005="${clsId}")*(${ATT}!H6:H5005="Present"))/MAX(1,COUNTIF(${ATT}!D6:D5005,"${clsId}")),"0%"),"–")`,
    ], 8, bg);
  });
  spacer(16);

  // ── GRADEBOOK SUMMARY (rows 17-25) ───────────────────────────────────
  sectionHeader(17, '📊  Gradebook Summary — All Classes', C.Math);
  tableHeader(18, ['Grade Category','Total Entries','Graded','Missing','Avg Score','Avg %'], 6);

  const GRADE_CATS = ['Homework','Classwork','Quizzes','Tests','Projects','Participation','Labs','Essays'];
  GRADE_CATS.forEach((cat, ci) => {
    const r = 19 + ci;
    const bg = ci % 2 === 0 ? C.panel : C.altRow;
    tableRow(r, [
      cat,
      `=IFERROR(COUNTIF(${GRADE}!H6:H5005,"${cat}"),"–")`,
      `=IFERROR(COUNTIFS(${GRADE}!H6:H5005,"${cat}",${GRADE}!Q6:Q5005,TRUE),"–")`,
      `=IFERROR(COUNTIFS(${GRADE}!H6:H5005,"${cat}",${GRADE}!N6:N5005,TRUE),"–")`,
      `=IFERROR(TEXT(AVERAGEIFS(${GRADE}!L6:L5005,${GRADE}!H6:H5005,"${cat}",${GRADE}!Q6:Q5005,TRUE),"0.0"),"–")`,
      `=IFERROR(TEXT(AVERAGEIFS(${GRADE}!M6:M5005,${GRADE}!H6:H5005,"${cat}",${GRADE}!Q6:Q5005,TRUE),"0%"),"–")`,
    ], 6, bg);
  });
  spacer(27);

  // ── ATTENDANCE SUMMARY (rows 28-34) ──────────────────────────────────
  sectionHeader(28, '📅  Attendance Summary', C.Science);
  tableHeader(29, ['Class','Present','Absent','Tardy','Excused','Total Days','Attendance Rate'], 7);

  CLASS_SNAP.forEach(([clsId, name], ci) => {
    const r = 30 + ci;
    const bg = ci % 2 === 0 ? C.panel : C.altRow;
    tableRow(r, [
      `${clsId} — ${name}`,
      `=IFERROR(COUNTIFS(${ATT}!D6:D5005,"${clsId}",${ATT}!H6:H5005,"Present"),"–")`,
      `=IFERROR(COUNTIFS(${ATT}!D6:D5005,"${clsId}",${ATT}!H6:H5005,"Absent"),"–")`,
      `=IFERROR(COUNTIFS(${ATT}!D6:D5005,"${clsId}",${ATT}!H6:H5005,"Tardy"),"–")`,
      `=IFERROR(COUNTIFS(${ATT}!D6:D5005,"${clsId}",${ATT}!H6:H5005,"Excused"),"–")`,
      `=IFERROR(COUNTIF(${ATT}!D6:D5005,"${clsId}"),"–")`,
      `=IFERROR(TEXT(COUNTIFS(${ATT}!D6:D5005,"${clsId}",${ATT}!H6:H5005,"Present")/MAX(1,COUNTIF(${ATT}!D6:D5005,"${clsId}")),"0%"),"–")`,
    ], 7, bg);
  });
  spacer(35);

  // ── TO-DO SUMMARY (rows 36-42) ───────────────────────────────────────
  sectionHeader(36, '✅  To-Do Summary', C.SocialStudies);
  tableHeader(37, ['Category','Total','Not Started','In Progress','Waiting','Complete','% Done'], 7);

  const TODO_CATS = ['Lesson Planning','Grading','Parent Communication','Meetings','Classroom Management','Materials','Professional Development','Administrative','Student Support','Other'];
  const TODO_SHOW = TODO_CATS.slice(0, 5); // show top 5 in dashboard
  TODO_SHOW.forEach((cat, ci) => {
    const r = 38 + ci;
    const bg = ci % 2 === 0 ? C.panel : C.altRow;
    tableRow(r, [
      cat,
      `=IFERROR(COUNTIF(${TODO}!C6:C1505,"${cat}"),"–")`,
      `=IFERROR(COUNTIFS(${TODO}!C6:C1505,"${cat}",${TODO}!E6:E1505,"Not Started"),"–")`,
      `=IFERROR(COUNTIFS(${TODO}!C6:C1505,"${cat}",${TODO}!E6:E1505,"In Progress"),"–")`,
      `=IFERROR(COUNTIFS(${TODO}!C6:C1505,"${cat}",${TODO}!E6:E1505,"Waiting"),"–")`,
      `=IFERROR(COUNTIFS(${TODO}!C6:C1505,"${cat}",${TODO}!E6:E1505,"Complete"),"–")`,
      `=IFERROR(TEXT(COUNTIFS(${TODO}!C6:C1505,"${cat}",${TODO}!E6:E1505,"Complete")/MAX(1,COUNTIF(${TODO}!C6:C1505,"${cat}")),"0%"),"–")`,
    ], 7, bg);
  });
  spacer(43);

  // ── ASSIGNMENT STATUS (rows 44-51) ───────────────────────────────────
  sectionHeader(44, '📝  Assignment Pipeline — All Classes', C.Writing);
  tableHeader(45, ['Status','Count','% of Total','Action Needed'], 4);

  const ASSIGN_STATUSES = ['Planned','Assigned','In Progress','Collected','Grading','Returned','Cancelled'];
  const ASSIGN_ACTIONS  = ['No action','Prepare materials','Check student progress','Enter scores','Grade and return','','Remove from active list'];
  ASSIGN_STATUSES.forEach((st, si) => {
    const r = 46 + si;
    const bg = si % 2 === 0 ? C.panel : C.altRow;
    tableRow(r, [
      st,
      `=IFERROR(COUNTIF(${ASSIGN}!O6:O1505,"${st}"),"–")`,
      `=IFERROR(TEXT(COUNTIF(${ASSIGN}!O6:O1505,"${st}")/MAX(1,COUNTA(${ASSIGN}!O6:O1505)),"0%"),"–")`,
      ASSIGN_ACTIONS[si],
    ], 4, bg);
  });
  spacer(53);

  // ── DISCLAIMER (row 54) ──────────────────────────────────────────────
  vals.push({ range:`${S}!A54`, values:[['This dashboard displays live data from all tabs. Scores update automatically when grades are entered. No macros, scripts, or external connections — formula-powered only. | Ultimate Teacher Planner']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,53,54,0,9), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,53,54,0,9), cell:{userEnteredFormat:{
    backgroundColor:hex(C.bg), textFormat:{fontSize:7,foregroundColor:hex(C.border),italic:true,fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:53,endIndex:54}, properties:{pixelSize:18}, fields:'pixelSize' }});

  // Column widths
  const colWidths = [110,120,90,90,110,90,110,110,80,80,80,80,80,80,80,80,80,80];
  colWidths.forEach((px, c) => {
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:c,endIndex:c+1}, properties:{pixelSize:px}, fields:'pixelSize' }});
  });

  // Freeze rows 1-3
  fmt.push({ updateSheetProperties:{ properties:{ sheetId:SID, gridProperties:{ frozenRowCount:3 }}, fields:'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '12-dashboard values');
  await batchUpdate(id, fmt, '12-dashboard format');

  // ── CHARTS ───────────────────────────────────────────────────────────
  // Build helper lookup tables for chart data (contiguous ranges)
  // Helper table A: Class Avg Score by Class (rows 56-62)
  // Helper table B: Assignment Status counts (rows 64-72)
  // Helper table C: Attendance by class (rows 74-80)
  // Helper table D: To-Do by status (rows 82-88)
  // All at cols K-L (indices 10-11, right side, not visible in main view)

  const helperVals = [];
  const K = 'K', L = 'L', M = 'M';

  // Table A: Class avg scores (for bar chart)
  helperVals.push({ range:`${S}!K56:L56`, values:[['Class','Avg Score %']] });
  CLASS_SNAP.forEach(([clsId, name], ci) => {
    helperVals.push({ range:`${S}!K${57+ci}:L${57+ci}`, values:[[
      name,
      `=IFERROR(AVERAGEIFS(${GRADE}!M6:M5005,${GRADE}!D6:D5005,"${clsId}",${GRADE}!Q6:Q5005,TRUE),0)`,
    ]]});
  });

  // Table B: Assignment status counts (for pie/column chart)
  helperVals.push({ range:`${S}!K64:L64`, values:[['Status','Count']] });
  ASSIGN_STATUSES.forEach((st, si) => {
    helperVals.push({ range:`${S}!K${65+si}:L${65+si}`, values:[[
      st,
      `=IFERROR(COUNTIF(${ASSIGN}!O6:O1505,"${st}"),0)`,
    ]]});
  });

  // Table C: Attendance present % by class (for bar chart)
  helperVals.push({ range:`${S}!K74:L74`, values:[['Class','Attendance %']] });
  CLASS_SNAP.forEach(([clsId, name], ci) => {
    helperVals.push({ range:`${S}!K${75+ci}:L${75+ci}`, values:[[
      name,
      `=IFERROR(COUNTIFS(${ATT}!D6:D5005,"${clsId}",${ATT}!H6:H5005,"Present")/MAX(1,COUNTIF(${ATT}!D6:D5005,"${clsId}")),0)`,
    ]]});
  });

  // Table D: To-Do by status (all tasks)
  helperVals.push({ range:`${S}!K82:L82`, values:[['Status','Tasks']] });
  ['Not Started','In Progress','Waiting','Complete','Cancelled'].forEach((st, si) => {
    helperVals.push({ range:`${S}!K${83+si}:L${83+si}`, values:[[
      st,
      `=IFERROR(COUNTIF(${TODO}!E6:E1505,"${st}"),0)`,
    ]]});
  });

  await valuesBatchUpdate(id, helperVals, '12-dashboard helpers');

  // Style helper tables (small, muted)
  const helperFmt = [];
  [[56,62],[64,72],[74,80],[82,88]].forEach(([r1, r2]) => {
    helperFmt.push({ repeatCell:{ range: gridRange(SID,r1-1,r2,10,12), cell:{userEnteredFormat:{
      backgroundColor:hex(C.altRow), textFormat:{fontSize:7,foregroundColor:hex(C.secText),fontFamily:'Arial'},
    }}, fields:'userEnteredFormat' }});
  });
  await batchUpdate(id, helperFmt, '12-dashboard helper fmt');

  // ── ADD CHARTS ───────────────────────────────────────────────────────
  // Sheet row indices for helper tables (0-based):
  // Table A: rows 55-60 (header at 55, data 56-60)
  // Table B: rows 63-70 (header at 63, data 64-70)
  // Table C: rows 73-78 (header at 73, data 74-78)
  // Table D: rows 81-86 (header at 81, data 82-86)

  const chartRequests = [
    // Chart 1: Class Avg Score (BAR) — anchored rows 5-20, cols 9-18
    { addChart:{ chart:{
      spec:{
        title:'Average Class Score (%)',
        titleTextFormat:{fontSize:10,bold:true},
        basicChart:{
          chartType:'BAR',
          legendPosition:'NO_LEGEND',
          headerCount:1,
          domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:SID, startRowIndex:55, endRowIndex:61, startColumnIndex:10, endColumnIndex:11 }] }}}],
          series:[{ series:{ sourceRange:{ sources:[{ sheetId:SID, startRowIndex:55, endRowIndex:61, startColumnIndex:11, endColumnIndex:12 }] }}, color:hex(C.Math) }],
        },
      },
      position:{ overlayPosition:{ anchorCell:{ sheetId:SID, rowIndex:4, columnIndex:9 }, offsetXPixels:0, offsetYPixels:0, widthPixels:360, heightPixels:200 }},
    }}},
    // Chart 2: Assignment Status (COLUMN)
    { addChart:{ chart:{
      spec:{
        title:'Assignment Pipeline by Status',
        titleTextFormat:{fontSize:10,bold:true},
        basicChart:{
          chartType:'COLUMN',
          legendPosition:'NO_LEGEND',
          headerCount:1,
          domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:SID, startRowIndex:63, endRowIndex:71, startColumnIndex:10, endColumnIndex:11 }] }}}],
          series:[{ series:{ sourceRange:{ sources:[{ sheetId:SID, startRowIndex:63, endRowIndex:71, startColumnIndex:11, endColumnIndex:12 }] }}, color:hex(C.ELA) }],
        },
      },
      position:{ overlayPosition:{ anchorCell:{ sheetId:SID, rowIndex:4, columnIndex:14 }, offsetXPixels:0, offsetYPixels:0, widthPixels:360, heightPixels:200 }},
    }}},
    // Chart 3: Attendance % (BAR)
    { addChart:{ chart:{
      spec:{
        title:'Attendance Rate by Class (%)',
        titleTextFormat:{fontSize:10,bold:true},
        basicChart:{
          chartType:'BAR',
          legendPosition:'NO_LEGEND',
          headerCount:1,
          domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:SID, startRowIndex:73, endRowIndex:79, startColumnIndex:10, endColumnIndex:11 }] }}}],
          series:[{ series:{ sourceRange:{ sources:[{ sheetId:SID, startRowIndex:73, endRowIndex:79, startColumnIndex:11, endColumnIndex:12 }] }}, color:hex(C.Science) }],
        },
      },
      position:{ overlayPosition:{ anchorCell:{ sheetId:SID, rowIndex:20, columnIndex:9 }, offsetXPixels:0, offsetYPixels:0, widthPixels:360, heightPixels:200 }},
    }}},
    // Chart 4: To-Do status (COLUMN)
    { addChart:{ chart:{
      spec:{
        title:'To-Do Tasks by Status',
        titleTextFormat:{fontSize:10,bold:true},
        basicChart:{
          chartType:'COLUMN',
          legendPosition:'NO_LEGEND',
          headerCount:1,
          domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:SID, startRowIndex:81, endRowIndex:87, startColumnIndex:10, endColumnIndex:11 }] }}}],
          series:[{ series:{ sourceRange:{ sources:[{ sheetId:SID, startRowIndex:81, endRowIndex:87, startColumnIndex:11, endColumnIndex:12 }] }}, color:hex(C.SocialStudies) }],
        },
      },
      position:{ overlayPosition:{ anchorCell:{ sheetId:SID, rowIndex:20, columnIndex:14 }, offsetXPixels:0, offsetYPixels:0, widthPixels:360, heightPixels:200 }},
    }}},
  ];

  await batchUpdate(id, chartRequests, '12-dashboard charts');
  console.log('✅ Teacher Dashboard done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
