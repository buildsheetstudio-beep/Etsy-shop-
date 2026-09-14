'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Annual Calendar'];
const S = "'Annual Calendar'";

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,80,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Title row
  vals.push({ range: `${S}!A1`, values: [['ANNUAL CALENDAR — ADHD FOCUS FLOW PLANNER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Year selector row 2
  vals.push({ range: `${S}!A2`, values: [['Year:','2026','','Tasks Completed:','','Tasks Total:','','Completion Rate:','','','','','','','','']] });
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coralTint), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,1,2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.white), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.coral) },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});

  // Year summary stats
  const yearSummary = [
    `=COUNTIF('Master Tasks'!E8:E5007,"Completed")`,
    `=COUNTA('Master Tasks'!C8:C5007)`,
    `=IFERROR(TEXT(E2/F2,"0%"),"0%")`,
  ];
  vals.push({ range: `${S}!E2`, values: [[yearSummary[0]]] });
  vals.push({ range: `${S}!G2`, values: [[yearSummary[1]]] });
  vals.push({ range: `${S}!I2`, values: [[yearSummary[2]]] });
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,4,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coralTint), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.coral) },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // Spacer row 3
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,16), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // 3×4 mini-calendar grid layout
  // Each mini-calendar is 10 columns wide (cols A-J, K-T ... but we have 16 cols available)
  // Layout: 3 calendars per row, 4 rows = 12 months
  // Each mini-calendar: 8 rows tall (1 month header + 1 day header + 6 week rows)
  // 3 mini-cals per row, 1 spacer col between each: cols 0-4, 6-10, 12-15 (0-indexed)
  // Actually simpler: use 5 cols per mini-calendar with cols 0-4, 5-9, 10-14 and no spacer col

  // Mini-calendar column width = 5 cols per month
  // Days: S M T W T F S in 7 slots but we only have 5 cols — use 7 cols total per month
  // Layout: Month 1 (cols A-G), Month 2 (cols H-N), Month 3 (cols O-U) — but we only have 16 cols
  // Use compact layout: 7 day cols, 1 spacer, repeat: 7+1+7+1+7 = 23 cols — too wide
  // Use 5-col compact: Mon-Fri merged, Sat+Sun merged — that loses info
  // Better: 3 months in a row with 4 cols each + 3 spacer cols = 15 cols = A-O
  // Each month: Mon, Tue, Wed, Thu, Fri, Sa, Su across 7 cols...
  // Final decision: use 8 cols per month (7 day cols + 1 spacer) across 3 months = 24 — still too wide
  // Use 7 cols for month 1 (A-G), 7 for month 2 (I-O), 7 for month 3 (Q-W) ... 21+2 spacer = 23 too wide
  // Compromise: use cols 0-4 (5), 5-9 (5), 10-14 (5) — 5 cols per month with abbreviated day headers (M T W T F S S)
  // Actually the grid is 16 cols. Let's use 3 months × 5 cols each = 15 + 1 spacer = 16 exactly.
  // 5 cols: M+T merged (1 col), W (1 col), T+F merged (1 col), S (1), S (1) — loses readability
  // BEST: just 7 cols per mini-calendar in a 2×6 arrangement (2 months wide, 6 month-blocks tall)
  // But that's very tall. Let's do 12 months as a list instead — one row per month with summary stats.
  // That's the most readable for ADHD users and avoids the 24-col problem.

  // Monthly summary list: row per month with stats
  const MONTH_HEADER = ['Month','Tasks Planned','Tasks Completed','Completion %','High Priority Done','Events','Notes / Theme','','','','','','','','',''];
  vals.push({ range: `${S}!A4`, values: [MONTH_HEADER] });
  fmt.push({ repeatCell: { range: gridRange(SID,3,4,0,7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const MONTH_COLORS = [C.blueTint,C.lavTint,C.coralTint,C.sandTint,C.tealTint,C.blueTint,C.coralTint,C.sandTint,C.lavTint,C.tealTint,C.blueTint,C.lavTint];

  const monthRows = MONTHS.map((m, mi) => {
    const mn = mi + 1;
    // bare expressions (no leading =) for embedding
    const tasksPlannedExpr = `SUMPRODUCT((YEAR('Master Tasks'!F8:F5007)=2026)*(MONTH('Master Tasks'!F8:F5007)=${mn})*('Master Tasks'!C8:C5007<>""))`;
    const tasksDoneExpr = `SUMPRODUCT((YEAR('Master Tasks'!F8:F5007)=2026)*(MONTH('Master Tasks'!F8:F5007)=${mn})*('Master Tasks'!E8:E5007="Completed"))`;
    const tasksPlanned = `=${tasksPlannedExpr}`;
    const tasksDone = `=${tasksDoneExpr}`;
    const compRate = `=IFERROR(TEXT(${tasksDoneExpr}/${tasksPlannedExpr},"0%"),"—")`;
    const hpDone = `=SUMPRODUCT((YEAR('Master Tasks'!F8:F5007)=2026)*(MONTH('Master Tasks'!F8:F5007)=${mn})*('Master Tasks'!D8:D5007="Must Do")*('Master Tasks'!E8:E5007="Completed"))`;
    const events = `=SUMPRODUCT((YEAR('Master Tasks'!F8:F5007)=2026)*(MONTH('Master Tasks'!F8:F5007)=${mn})*('Master Tasks'!B8:B5007="Event"))`;

    return [m, tasksPlanned, tasksDone, compRate, hpDone, events, '', '', '', '', '', '', '', '', '', ''];
  });

  vals.push({ range: `${S}!A5`, values: monthRows });

  MONTHS.forEach((m, mi) => {
    fmt.push({ repeatCell: { range: gridRange(SID,4+mi,5+mi,0,16), cell: { userEnteredFormat: {
      backgroundColor: hex(MONTH_COLORS[mi]),
      textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });

  // Highlight current month (September = index 8)
  fmt.push({ repeatCell: { range: gridRange(SID,12,13,0,7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coral), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});

  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 4+12 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Spacer row 17
  fmt.push({ repeatCell: { range: gridRange(SID,16,17,0,16), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 16, endIndex: 17 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // Annual goals section
  vals.push({ range: `${S}!A18`, values: [['ANNUAL GOALS & KEY MILESTONES']] });
  fmt.push({ mergeCells: { range: gridRange(SID,17,18,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,17,18,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.blue), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 17, endIndex: 18 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  const GOAL_HDRS = ['Goal / Milestone','Category','Target Month','Status','Notes / Progress','','','','','','','','','','',''];
  vals.push({ range: `${S}!A19`, values: [GOAL_HDRS] });
  fmt.push({ repeatCell: { range: gridRange(SID,18,19,0,5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.headerMid), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 18, endIndex: 19 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  const GOALS = [
    ['Get promoted to Senior level','Work','December','In Progress','Performance review in September','','','','','','','','','','',''],
    ['Build a consistent morning routine','Personal','October','In Progress','Currently at 3/5 days per week','','','','','','','','','','',''],
    ['Save $5,000 for emergency fund','Personal','December','In Progress','At $2,400 so far','','','','','','','','','','',''],
    ['Run a 5K','Personal','November','Not Started','Need to start training plan','','','','','','','','','','',''],
    ['Read 12 books (1/month)','Personal','December','In Progress','7 books completed so far','','','','','','','','','','',''],
    ['Implement weekly review system','Personal','September','Not Started','Start this Sunday','','','','','','','','','','',''],
    ['Clear email inbox to zero','Work','October','Not Started','1,200 unread currently','','','','','','','','','','',''],
    ['Plan and book winter family vacation','Personal','October','Not Started','Destination TBD','','','','','','','','','','',''],
    ['Develop consistent ADHD management routine','Personal','December','In Progress','Using this planner as the system','','','','','','','','','',''],
    ['Complete Q3 and Q4 at work above expectations','Work','December','In Progress','Q3 report this week','','','','','','','','','','',''],
  ];
  vals.push({ range: `${S}!A20`, values: GOALS });
  GOALS.forEach((g, gi) => {
    const bg = gi % 2 === 0 ? C.blueTint : C.white;
    fmt.push({ repeatCell: { range: gridRange(SID,19+gi,20+gi,0,16), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 19, endIndex: 19+GOALS.length }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // Spacer
  fmt.push({ repeatCell: { range: gridRange(SID,19+GOALS.length,20+GOALS.length,0,16), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 19+GOALS.length, endIndex: 20+GOALS.length }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // Key dates section
  const keyRow = 20 + GOALS.length;
  vals.push({ range: `${S}!A${keyRow+1}`, values: [['KEY DATES & RECURRING EVENTS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,keyRow,keyRow+1,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,keyRow,keyRow+1,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.teal), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: keyRow, endIndex: keyRow+1 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  vals.push({ range: `${S}!A${keyRow+2}`, values: [['Date / Month','Event / Deadline','Category','Recurring?','Notes','','','','','','','','','','','']] });
  fmt.push({ repeatCell: { range: gridRange(SID,keyRow+1,keyRow+2,0,5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: keyRow+1, endIndex: keyRow+2 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  const KEY_DATES = [
    ['September 14','Today — Weekly Review','Personal','Weekly',''],
    ['September 15','Q3 Report Due','Work','Quarterly',''],
    ['September 17','Team Presentation','Work','Once',''],
    ['September 21','Dentist Appointment','Health','Annual',''],
    ['September 25','Mom\'s Birthday','Personal','Annual','Book restaurant by Sep 22'],
    ['September 28','Performance Review','Work','Annual',''],
    ['October 31','Halloween','Personal','Annual',''],
    ['November (last Thu)','Thanksgiving','Personal','Annual',''],
    ['December 25','Christmas / Holidays','Personal','Annual',''],
    ['January 1','New Year — Annual Goal Review','Personal','Annual',''],
    ['April 15','Tax Day (US)','Financial','Annual','File or extend by this date'],
    ['Monthly (last Mon)','Monthly Budget Review','Financial','Monthly',''],
    ['Weekly (Monday)','Weekly Review & Planning','Personal','Weekly',''],
    ['Daily','Daily Focus Check-in','Personal','Daily',''],
  ];
  vals.push({ range: `${S}!A${keyRow+3}`, values: KEY_DATES.map(r => r.concat(['','','','','','','','','','','']))});
  KEY_DATES.forEach((r, ri) => {
    const bg = ri % 2 === 0 ? C.tealTint : C.white;
    fmt.push({ repeatCell: { range: gridRange(SID,keyRow+2+ri,keyRow+3+ri,0,16), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: keyRow+2, endIndex: keyRow+2+KEY_DATES.length }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // Column widths
  const WIDTHS = [130,100,100,100,120,100,100,100,100,100,100,100,100,100,100,100];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1-4
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 4 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '09-annualcal values');
  await batchUpdate(id, fmt, '09-annualcal format');
  console.log('✅  Annual Calendar done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
