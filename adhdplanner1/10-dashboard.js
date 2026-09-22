'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['ADHD Focus Dashboard'];
const S = "'ADHD Focus Dashboard'";

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,80,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // ===== Row 1: Title =====
  vals.push({ range: `${S}!A1`, values: [['ADHD FOCUS FLOW PLANNER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' }});

  // ===== Row 2: Subtitle / tagline =====
  vals.push({ range: `${S}!A2`, values: [['Your complete system for ADHD focus — Capture, Organize, Prioritize, Execute, Review']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.headerMid), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // ===== Row 3: Spacer =====
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,16), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // ===== Row 4: Today's date + Day type =====
  vals.push({ range: `${S}!A4`, values: [['TODAY:','=TODAY()','','Day Type:','','Brain State:','','','','','','','','','','']] });
  fmt.push({ repeatCell: { range: gridRange(SID,3,4,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coralTint), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,3,4,1,2), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'dddd, mmmm d yyyy' },
    backgroundColor: hex(C.coralTint), textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.coral) },
    horizontalAlignment: 'LEFT',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // ===== Section 1: FOCUS STAGE PIPELINE (row 5-6) =====
  // Row 5: section header
  vals.push({ range: `${S}!A5`, values: [['FOCUS STAGE PIPELINE']] });
  fmt.push({ mergeCells: { range: gridRange(SID,4,5,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,4,5,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coral), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // Row 6: stage labels (5 stages)
  const STAGES = ['NOW','NEXT','LATER','EVENT','REVISIT'];
  const STAGE_COLORS = [C.coral, C.blue, C.lavender, C.teal, C.sand];
  const STAGE_FORMULAS = STAGES.map(s => `=COUNTIF('Master Tasks'!B8:B5007,"${s.charAt(0)+s.slice(1).toLowerCase()}")`);

  // Stage labels row 6 — 5 cols of 3 each, labels
  const stageLabelRow = STAGES.map(s => s);
  const paddedStageLabels = [];
  STAGES.forEach(s => { paddedStageLabels.push(s, '', ''); });
  paddedStageLabels.push('');
  vals.push({ range: `${S}!A6`, values: [paddedStageLabels.slice(0,16)] });

  STAGES.forEach((s, si) => {
    const startCol = si * 3;
    fmt.push({ mergeCells: { range: gridRange(SID,5,6,startCol,startCol+3), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,5,6,startCol,startCol+3), cell: { userEnteredFormat: {
      backgroundColor: hex(STAGE_COLORS[si]),
      textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // Row 7: stage count values
  const stageCountRow = [];
  STAGE_FORMULAS.forEach(f => { stageCountRow.push(f, '', ''); });
  stageCountRow.push('');
  vals.push({ range: `${S}!A7`, values: [stageCountRow.slice(0,16)] });

  STAGES.forEach((s, si) => {
    const startCol = si * 3;
    fmt.push({ mergeCells: { range: gridRange(SID,6,7,startCol,startCol+3), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,6,7,startCol,startCol+3), cell: { userEnteredFormat: {
      backgroundColor: hex(STAGE_COLORS[si]+'33'.padStart(2,'0')),
      textFormat: { bold: true, fontSize: 20, foregroundColor: hex(STAGE_COLORS[si]), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });
  // Tint backgrounds for count rows
  const STAGE_TINTS = [C.coralTint, C.blueTint, C.lavTint, C.tealTint, C.sandTint];
  STAGES.forEach((s, si) => {
    const startCol = si * 3;
    fmt.push({ repeatCell: { range: gridRange(SID,6,7,startCol,startCol+3), cell: { userEnteredFormat: {
      backgroundColor: hex(STAGE_TINTS[si]),
      textFormat: { bold: true, fontSize: 20, foregroundColor: hex(STAGE_COLORS[si]), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // ===== Spacer =====
  fmt.push({ repeatCell: { range: gridRange(SID,7,8,0,16), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 8 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // ===== Section 2: KEY METRICS ROW (rows 9-11) =====
  vals.push({ range: `${S}!A9`, values: [['KEY METRICS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,8,9,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,8,9,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.blue), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 8, endIndex: 9 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // 8 metric cards: 2 cols each across 16 cols
  const METRICS = [
    { label: 'Total Tasks', formula: `=COUNTA('Master Tasks'!C8:C5007)`, color: C.blue },
    { label: 'Must Do Today', formula: `=COUNTIF('Master Tasks'!D8:D5007,"Must Do")`, color: C.coral },
    { label: 'Completed', formula: `=COUNTIF('Master Tasks'!E8:E5007,"Completed")`, color: C.teal },
    { label: 'Overdue', formula: `=COUNTIFS('Master Tasks'!F8:F5007,"<"&TODAY(),'Master Tasks'!E8:E5007,"<>Completed",'Master Tasks'!E8:E5007,"<>Archived",'Master Tasks'!C8:C5007,"<>")`, color: C.coral },
    { label: 'Due This Week', formula: `=COUNTIFS('Master Tasks'!F8:F5007,">="&TODAY(),'Master Tasks'!F8:F5007,"<="&TODAY()+7,'Master Tasks'!E8:E5007,"<>Completed",'Master Tasks'!E8:E5007,"<>Archived")`, color: C.sand },
    { label: 'Brain Dump Items', formula: `=COUNTA('Brain Dump'!C8:C2007)`, color: C.lavender },
    { label: 'Unprocessed', formula: `=COUNTA('Brain Dump'!C8:C2007)-COUNTIF('Brain Dump'!G8:G2007,TRUE)`, color: C.sand },
    { label: 'Completion %', formula: `=IFERROR(TEXT(COUNTIF('Master Tasks'!E8:E5007,"Completed")/COUNTA('Master Tasks'!C8:C5007),"0%"),"0%")`, color: C.teal },
  ];

  // Labels row 10
  const metricLabelRow = [];
  METRICS.forEach(m => { metricLabelRow.push(m.label, ''); });
  vals.push({ range: `${S}!A10`, values: [metricLabelRow] });
  METRICS.forEach((m, mi) => {
    const startCol = mi * 2;
    fmt.push({ mergeCells: { range: gridRange(SID,9,10,startCol,startCol+2), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,9,10,startCol,startCol+2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.headerMid),
      textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 9, endIndex: 10 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // Values row 11
  const metricValRow = [];
  METRICS.forEach(m => { metricValRow.push(m.formula, ''); });
  vals.push({ range: `${S}!A11`, values: [metricValRow] });
  const METRIC_TINTS = [C.blueTint, C.coralTint, C.tealTint, C.coralTint, C.sandTint, C.lavTint, C.sandTint, C.tealTint];
  METRICS.forEach((m, mi) => {
    const startCol = mi * 2;
    fmt.push({ mergeCells: { range: gridRange(SID,10,11,startCol,startCol+2), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,10,11,startCol,startCol+2), cell: { userEnteredFormat: {
      backgroundColor: hex(METRIC_TINTS[mi]),
      textFormat: { bold: true, fontSize: 18, foregroundColor: hex(m.color), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 38 }, fields: 'pixelSize' }});

  // Spacer
  fmt.push({ repeatCell: { range: gridRange(SID,11,12,0,16), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 11, endIndex: 12 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // ===== Section 3: TODAY'S TOP 5 TASKS (rows 13-19) =====
  vals.push({ range: `${S}!A13`, values: [['TODAY\'S FOCUS — TOP 5 TASKS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,12,13,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,12,13,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.teal), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 12, endIndex: 13 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  const T5_HDRS = ['#','Task Name','Focus Stage','Priority','Effort','Due Date','Status','','','','','','','','',''];
  vals.push({ range: `${S}!A14`, values: [T5_HDRS] });
  fmt.push({ repeatCell: { range: gridRange(SID,13,14,0,7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 13, endIndex: 14 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  const TOP5 = [
    ['1','Finish Q3 report executive summary','Now','Must Do','Deep Focus','09/15/2026','In Progress'],
    ['2','Send weekly status update email','Now','Must Do','Quick','09/14/2026','Ready'],
    ['3','Pay credit card bill','Now','Must Do','Quick','09/14/2026','Ready'],
    ['4','Prepare presentation slides','Now','Must Do','Deep Focus','09/17/2026','In Progress'],
    ['5','Return Amazon package','Next','Important','Quick','09/19/2026','Ready'],
  ];
  const top5Padded = TOP5.map(r => r.concat(Array(16-r.length).fill('')));
  vals.push({ range: `${S}!A15`, values: top5Padded });
  TOP5.forEach((r, ri) => {
    const tints = [C.coralTint, C.blueTint, C.tealTint, C.sandTint, C.lavTint];
    fmt.push({ repeatCell: { range: gridRange(SID,14+ri,15+ri,0,16), cell: { userEnteredFormat: {
      backgroundColor: hex(tints[ri]),
      textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 14, endIndex: 19 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Spacer
  fmt.push({ repeatCell: { range: gridRange(SID,19,20,0,16), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 19, endIndex: 20 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // ===== Section 4: STATUS BREAKDOWN (rows 21-28) =====
  vals.push({ range: `${S}!A21`, values: [['TASK STATUS OVERVIEW']] });
  fmt.push({ mergeCells: { range: gridRange(SID,20,21,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,20,21,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.lavender), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 20, endIndex: 21 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  const STATUSES = ['Inbox','Ready','In Progress','Waiting','Scheduled','Completed','Rescheduled','Archived'];
  const STATUS_COLORS = [C.medGray, C.blue, C.coral, C.sand, C.lavender, C.teal, C.sand, C.archivedBg];

  const statusHdrRow = ['Status','Count','% of Total','','','Status','Count','% of Total','','','','','','','',''];
  vals.push({ range: `${S}!A22`, values: [statusHdrRow] });
  fmt.push({ repeatCell: { range: gridRange(SID,21,22,0,3), cell: { userEnteredFormat: {
    backgroundColor: hex(C.headerMid), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white) },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,21,22,5,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.headerMid), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white) },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 21, endIndex: 22 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // 4 statuses in left col, 4 in right col
  const leftStats = STATUSES.slice(0,4);
  const rightStats = STATUSES.slice(4);
  const totalFormula = `COUNTA('Master Tasks'!C8:C5007)`;

  const statusDataLeft = leftStats.map((s, si) => {
    const count = `=COUNTIF('Master Tasks'!E8:E5007,"${s}")`;
    const pct = `=IFERROR(TEXT(COUNTIF('Master Tasks'!E8:E5007,"${s}")/${totalFormula},"0%"),"0%")`;
    return [s, count, pct, '', ''];
  });
  const statusDataRight = rightStats.map((s, si) => {
    const count = `=COUNTIF('Master Tasks'!E8:E5007,"${s}")`;
    const pct = `=IFERROR(TEXT(COUNTIF('Master Tasks'!E8:E5007,"${s}")/${totalFormula},"0%"),"0%")`;
    return [s, count, pct, '', '', '', '', '', '', '', ''];
  });

  statusDataLeft.forEach((row, ri) => {
    const fullRow = row.concat(statusDataRight[ri] || ['','','','','','','','','','','']);
    vals.push({ range: `${S}!A${23+ri}`, values: [fullRow.slice(0,16)] });
    const bg = ri % 2 === 0 ? C.lavTint : C.white;
    fmt.push({ repeatCell: { range: gridRange(SID,22+ri,23+ri,0,16), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 22, endIndex: 26 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // Spacer
  fmt.push({ repeatCell: { range: gridRange(SID,26,27,0,16), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 26, endIndex: 27 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // ===== Section 5: RESCUE PLANNER (rows 28-33) =====
  vals.push({ range: `${S}!A28`, values: [['RESCUE PLANNER — Stuck? Start Here']] });
  fmt.push({ mergeCells: { range: gridRange(SID,27,28,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,27,28,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.sand), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 27, endIndex: 28 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  const RESCUE_DATA = [
    ['Feeling Overwhelmed?','→ Go to Brain Dump. Write everything. Sort nothing. Just empty your head.','','','','','','','','','','','','','',''],
    ['Can\'t start a task?','→ Pick the SMALLEST possible first step. Just open the file. Just find the phone number.','','','','','','','','','','','','','',''],
    ['Distracted or hyper-focused?','→ Check Daily Focus tab. Use a 25-min Pomodoro. Set one visible timer.','','','','','','','','','','','','','',''],
    ['Running out of time?','→ Go to Master Tasks. Filter by "Must Do" + "Quick". Do those first.','','','','','','','','','','','','','',''],
    ['Bad day? Low energy?','→ Use a Low-Capacity Day type. Permission to do the bare minimum today.','','','','','','','','','','','','','',''],
    ['Nothing feels important?','→ Open Weekly Planner. What was the ONE thing you said mattered this week?','','','','','','','','','','','','','',''],
  ];
  vals.push({ range: `${S}!A29`, values: RESCUE_DATA });
  RESCUE_DATA.forEach((r, ri) => {
    const bg = ri % 2 === 0 ? C.sandTint : C.white;
    fmt.push({ repeatCell: { range: gridRange(SID,28+ri,29+ri,0,16), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    // Bold the label column A
    fmt.push({ repeatCell: { range: gridRange(SID,28+ri,29+ri,0,1), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.sand) },
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 28, endIndex: 28+RESCUE_DATA.length }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Spacer
  fmt.push({ repeatCell: { range: gridRange(SID,34,35,0,16), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 34, endIndex: 35 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // ===== Section 6: QUICK NAVIGATION (rows 36-38) =====
  vals.push({ range: `${S}!A36`, values: [['QUICK NAVIGATION — ADHD FOCUS FLOW PLANNER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,35,36,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,35,36,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.headerDark), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 35, endIndex: 36 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  const NAV_LABELS = ['Brain Dump','Master Tasks','Task Breakdown','Daily Focus','Weekly Planner','Monthly Calendar','Annual Calendar'];
  const NAV_DESCS = [
    'Capture everything on your mind',
    'Organize and prioritize all tasks',
    'Break big tasks into small steps',
    'Today\'s focus, priorities, and review',
    'Plan the week ahead',
    'Month-at-a-glance calendar view',
    'Annual goals and key milestones',
  ];
  const NAV_COLORS = [C.coral, C.blue, C.lavender, C.teal, C.sand, C.coral, C.blue];

  vals.push({ range: `${S}!A37`, values: [NAV_LABELS.concat(Array(16-NAV_LABELS.length).fill(''))] });
  vals.push({ range: `${S}!A38`, values: [NAV_DESCS.concat(Array(16-NAV_DESCS.length).fill(''))] });

  NAV_LABELS.forEach((label, ci) => {
    if (ci >= 16) return;
    fmt.push({ repeatCell: { range: gridRange(SID,36,37,ci,ci+1), cell: { userEnteredFormat: {
      backgroundColor: hex(NAV_COLORS[ci]),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,37,38,ci,ci+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.white),
      textFormat: { italic: true, fontSize: 8, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'TOP',
      wrapStrategy: 'WRAP',
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 36, endIndex: 37 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 37, endIndex: 38 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // Spacer
  fmt.push({ repeatCell: { range: gridRange(SID,38,39,0,16), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 38, endIndex: 39 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // ===== Section 7: PRIORITY BREAKDOWN (rows 40-44) =====
  vals.push({ range: `${S}!A40`, values: [['PRIORITY OVERVIEW']] });
  fmt.push({ mergeCells: { range: gridRange(SID,39,40,0,8), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,39,40,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coral), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 39, endIndex: 40 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  const PRIORITIES = ['Must Do','Important','Useful','Optional'];
  const PRI_COLORS = [C.coral, C.blue, C.teal, C.lavender];
  const PRI_TINTS = [C.coralTint, C.blueTint, C.tealTint, C.lavTint];

  const priHdr = PRIORITIES.map(p => p).concat(Array(12).fill(''));
  vals.push({ range: `${S}!A41`, values: [priHdr] });
  PRIORITIES.forEach((p, pi) => {
    fmt.push({ repeatCell: { range: gridRange(SID,40,41,pi*2,pi*2+2), cell: { userEnteredFormat: {
      backgroundColor: hex(PRI_COLORS[pi]),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ mergeCells: { range: gridRange(SID,40,41,pi*2,pi*2+2), mergeType: 'MERGE_ALL' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 40, endIndex: 41 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  const priCounts = PRIORITIES.map(p => `=COUNTIF('Master Tasks'!D8:D5007,"${p}")`).concat(Array(8).fill(''));
  // Interleave with blanks for merged cells
  const priCountInterleaved = [];
  PRIORITIES.forEach(p => { priCountInterleaved.push(`=COUNTIF('Master Tasks'!D8:D5007,"${p}")`, ''); });
  while (priCountInterleaved.length < 16) priCountInterleaved.push('');
  vals.push({ range: `${S}!A42`, values: [priCountInterleaved] });
  PRIORITIES.forEach((p, pi) => {
    fmt.push({ mergeCells: { range: gridRange(SID,41,42,pi*2,pi*2+2), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,41,42,pi*2,pi*2+2), cell: { userEnteredFormat: {
      backgroundColor: hex(PRI_TINTS[pi]),
      textFormat: { bold: true, fontSize: 18, foregroundColor: hex(PRI_COLORS[pi]), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 41, endIndex: 42 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Column widths
  const WIDTHS = [140,120,120,100,100,100,100,100,100,100,100,100,100,100,100,100];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze row 1
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '10-dashboard values');
  await batchUpdate(id, fmt, '10-dashboard format');
  console.log('✅  Dashboard done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
