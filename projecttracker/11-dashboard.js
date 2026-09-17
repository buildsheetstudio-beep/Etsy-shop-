'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Project Dashboard'];
const S = "'Project Dashboard'";
const MTL = "'Master Task Log'";
const MST = "'Milestones & Progress'";
const PS  = "'Project Setup'";
const REC = "'Recurring Tasks'";

const PROJ  = `${PS}!$A$4`;
const YEAR  = `${PS}!$S$4`;

const NC = 8; // 8 columns A-H

// Helper: SUMPRODUCT-based count (avoids COUNTIFS with computed criteria)
const spCount = (range, cond) => `SUMPRODUCT((${cond})*1)`;

(async () => {
  const fmt = [];
  const vals = [];

  // ── Title ──────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 24, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['PROJECT DASHBOARD']] });

  // ── Subtitle ───────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.teal), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [[`=IFERROR("Portfolio snapshot — "&COUNTA(${PS}!$A$8:$A$507)&" projects | As of "&TEXT(TODAY(),"mmm d, yyyy"),"Project Dashboard")`]] });

  // ── Instructions ───────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.terra), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A3`, values: [['READ-ONLY — aggregates from all sheets. Filter by project in Project Setup A4. Reporting Year in S4 controls YTD metrics. Refresh by pressing Ctrl+Shift+F9.']] });

  // ── Section: Portfolio KPIs (row 3-4, 0-indexed) ─────────────────────────
  // 8 KPI tiles across cols A-H (1 col each)
  const KPIs = [
    { label: 'Total Projects',     fml: `=COUNTA(${PS}!$A$8:$A$507)` },
    { label: 'Active Projects',    fml: `=COUNTIF(${PS}!$D$8:$D$507,"Active")` },
    { label: 'Total Tasks',        fml: `=COUNTA(${MTL}!$F$8:$F$3007)` },
    { label: 'Tasks Complete',     fml: `=COUNTIF(${MTL}!$N$8:$N$3007,"Complete")` },
    { label: 'Overdue Tasks',      fml: `=COUNTIF(${MTL}!$V$8:$V$3007,TRUE)` },
    { label: 'Milestones',         fml: `=COUNTA(${MST}!$D$8:$D$507)` },
    { label: 'Milestones Done',    fml: `=COUNTIF(${MST}!$H$8:$H$507,"Complete")` },
    { label: 'Recurring Templates',fml: `=COUNTA(${REC}!$B$8:$B$507)` },
  ];
  const KPI_COLORS = [C.primary, C.primary, C.teal, C.success, C.attention, C.secondary, C.success, C.info];

  // KPI label row (row 3)
  KPIs.forEach(({ label }, i) => {
    fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, i, i+1), cell: { userEnteredFormat: { backgroundColor: hex(KPI_COLORS[i]), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!${String.fromCharCode(65+i)}4`, values: [[label]] });
  });
  // KPI value row (row 4)
  KPIs.forEach(({ fml }, i) => {
    fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, i, i+1), cell: { userEnteredFormat: { backgroundColor: hex(KPI_COLORS[i]), textFormat: { bold: true, fontSize: 26, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'TOP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!${String.fromCharCode(65+i)}5`, values: [[fml]] });
  });

  // ── Spacer row 5 (0-indexed) ───────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 5, 6, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // ── Section header: Project Health Overview ────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 6, 7, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 6, 7, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 12, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A7`, values: [['  PROJECT HEALTH OVERVIEW']] });

  // ── Project table headers (row 7, 0-indexed) ─────────────────────────────
  const PRJ_HDRS = ['Project ID','Project Name','Status','Owner','Tasks','Done','Overdue','Health'];
  PRJ_HDRS.forEach((h, i) => {
    fmt.push({ repeatCell: { range: gridRange(SID, 7, 8, i, i+1), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  });
  vals.push({ range: `${S}!A8`, values: [PRJ_HDRS] });

  // ── Project rows (rows 8-27, 0-indexed = rows 9-28, 20 projects) ─────────
  for (let r = 0; r < 20; r++) {
    const row1 = 9 + r; // 1-indexed sheet row
    const psRow = 8 + r; // PS data starts at row 8 (1-indexed)
    const bgColor = r % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, 8+r, 9+r, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(bgColor), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });

    const projId   = `${PS}!$A$${psRow}`;
    const projName = `${PS}!$B$${psRow}`;
    const status   = `${PS}!$D$${psRow}`;
    const owner    = `${PS}!$C$${psRow}`;

    const rowVals = [
      `=IFERROR(${projId},"")`,
      `=IFERROR(${projName},"")`,
      `=IFERROR(${status},"")`,
      `=IFERROR(${owner},"")`,
      `=IFERROR(COUNTIF(${MTL}!$B$8:$B$3007,${projId}),"")`,
      `=IFERROR(COUNTIFS(${MTL}!$B$8:$B$3007,${projId},${MTL}!$N$8:$N$3007,"Complete"),"")`,
      `=IFERROR(COUNTIFS(${MTL}!$B$8:$B$3007,${projId},${MTL}!$V$8:$V$3007,TRUE),"")`,
      `=IFERROR(IF(${projId}="","",IF(COUNTIF(${MTL}!$B$8:$B$3007,${projId})=0,"No Tasks",IF(COUNTIFS(${MTL}!$B$8:$B$3007,${projId},${MTL}!$V$8:$V$3007,TRUE)>0,"At Risk","On Track"))),"")`,
    ];
    vals.push({ range: `${S}!A${row1}:H${row1}`, values: [rowVals] });
  }

  // ── Spacer ─────────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 28, 29, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // ── Section header: Task Status Breakdown ────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 29, 30, 0, 4), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 29, 30, 0, 4), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 12, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A30`, values: [['  TASK STATUS BREAKDOWN']] });

  // Section header: Upcoming Milestones
  fmt.push({ mergeCells: { range: gridRange(SID, 29, 30, 4, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 29, 30, 4, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 12, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!E30`, values: [['  UPCOMING MILESTONES (Next 30 Days)']] });

  // ── Task status rows (rows 30-36, 0-indexed) ──────────────────────────────
  const STATUSES = [
    { label: 'Backlog',      color: C.gray },
    { label: 'Not Started',  color: C.secText },
    { label: 'In Progress',  color: C.primary },
    { label: 'Waiting',      color: C.warning },
    { label: 'Blocked',      color: C.attention },
    { label: 'Review',       color: C.info },
    { label: 'Complete',     color: C.success },
  ];

  // Status table headers
  ['Status','Count','% of Total','',''].forEach((h, i) => {
    if (i < 3) {
      fmt.push({ repeatCell: { range: gridRange(SID, 30, 31, i, i+1), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
    }
  });
  vals.push({ range: `${S}!A31`, values: [['Status','Count','% of Total']] });

  STATUSES.forEach(({ label, color }, si) => {
    const r0 = 31 + si; // 0-indexed
    const row1 = 32 + si; // 1-indexed
    const bgColor = si % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 0, 3), cell: { userEnteredFormat: { backgroundColor: hex(bgColor), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
    // Color dot in status col
    fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(color), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
    const countF = `=COUNTIF(${MTL}!$N$8:$N$3007,"${label}")`;
    const pctF   = `=IFERROR(COUNTIF(${MTL}!$N$8:$N$3007,"${label}")/COUNTA(${MTL}!$F$8:$F$3007),0)`;
    vals.push({ range: `${S}!A${row1}:C${row1}`, values: [[label, countF, pctF]] });
    // Format pct as %
    fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 2, 3), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });

  // ── Upcoming milestones (cols E-H, rows 30-38) ───────────────────────────
  const MST_HDRS = ['Milestone','Project','Due Date','Status'];
  MST_HDRS.forEach((h, i) => {
    fmt.push({ repeatCell: { range: gridRange(SID, 30, 31, 4+i, 5+i), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  });
  vals.push({ range: `${S}!E31`, values: [MST_HDRS] });

  for (let m = 0; m < 10; m++) {
    const r0 = 31 + m;
    const row1 = 32 + m;
    const bgColor = m % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 4, NC), cell: { userEnteredFormat: { backgroundColor: hex(bgColor), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
    // upcoming milestone FILTER: due >= today, due <= today+30, not complete
    const mstFilter = `(${MST}!$F$8:$F$507>=TODAY())*(${MST}!$F$8:$F$507<=TODAY()+30)*(${MST}!$H$8:$H$507<>"Complete")`;
    vals.push({ range: `${S}!E${row1}:H${row1}`, values: [[
      `=IFERROR(INDEX(FILTER(${MST}!$D$8:$D$507,${mstFilter}),${m+1}),"")`,
      `=IFERROR(INDEX(FILTER(${MST}!$B$8:$B$507,${mstFilter}),${m+1}),"")`,
      `=IFERROR(INDEX(FILTER(${MST}!$F$8:$F$507,${mstFilter}),${m+1}),"")`,
      `=IFERROR(INDEX(FILTER(${MST}!$H$8:$H$507,${mstFilter}),${m+1}),"")`,
    ]] });
    // Date format
    fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 6, 7), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'M/d/yy' } } }, fields: 'userEnteredFormat.numberFormat' } });
  }

  // ── Spacer ─────────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 42, 43, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // ── Section header: Workload by Assignee ─────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 43, 44, 0, 4), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 43, 44, 0, 4), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 12, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A44`, values: [['  WORKLOAD BY ASSIGNEE']] });

  // Section header: Monthly Completion (YTD)
  fmt.push({ mergeCells: { range: gridRange(SID, 43, 44, 4, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 43, 44, 4, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 12, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!E44`, values: [['  MONTHLY COMPLETION (YTD)']] });

  // Workload table headers
  ['Assignee','Open','Complete','Total'].forEach((h, i) => {
    fmt.push({ repeatCell: { range: gridRange(SID, 44, 45, i, i+1), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  });
  vals.push({ range: `${S}!A45`, values: [['Assignee','Open','Complete','Total']] });

  // Assignees from Reference Data
  const ASSIGNEES = ['Alex Johnson','Maria Rodriguez','Sam Chen','Jordan Lee','Taylor Kim','Casey Williams','Morgan Davis'];
  ASSIGNEES.forEach((name, ai) => {
    const r0 = 45 + ai;
    const row1 = 46 + ai;
    const bgColor = ai % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 0, 4), cell: { userEnteredFormat: { backgroundColor: hex(bgColor), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
    vals.push({ range: `${S}!A${row1}:D${row1}`, values: [[
      name,
      `=COUNTIFS(${MTL}!$E$8:$E$3007,"${name}",${MTL}!$N$8:$N$3007,"<>Complete",${MTL}!$N$8:$N$3007,"<>Cancelled")`,
      `=COUNTIFS(${MTL}!$E$8:$E$3007,"${name}",${MTL}!$N$8:$N$3007,"Complete")`,
      `=COUNTIF(${MTL}!$E$8:$E$3007,"${name}")`,
    ]] });
  });

  // Monthly completion table
  const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  ['Month','Due','Completed','On-Time Rate'].forEach((h, i) => {
    fmt.push({ repeatCell: { range: gridRange(SID, 44, 45, 4+i, 5+i), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  });
  vals.push({ range: `${S}!E45`, values: [['Month','Due','Completed','On-Time Rate']] });

  MONTHS.forEach((mon, mi) => {
    const r0 = 45 + mi;
    const row1 = 46 + mi;
    const bgColor = mi % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 4, NC), cell: { userEnteredFormat: { backgroundColor: hex(bgColor), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
    const mIdx = mi + 1;
    // SUMPRODUCT to avoid COUNTIFS with computed date criteria
    const dueFml   = `=SUMPRODUCT((MONTH(IF(${MTL}!$L$8:$L$3007="",DATE(1900,1,1),${MTL}!$L$8:$L$3007))=${mIdx})*(YEAR(IF(${MTL}!$L$8:$L$3007="",DATE(1900,1,1),${MTL}!$L$8:$L$3007))=${YEAR})*(${MTL}!$F$8:$F$3007<>""))`;
    const doneFml  = `=SUMPRODUCT((MONTH(IF(${MTL}!$L$8:$L$3007="",DATE(1900,1,1),${MTL}!$L$8:$L$3007))=${mIdx})*(YEAR(IF(${MTL}!$L$8:$L$3007="",DATE(1900,1,1),${MTL}!$L$8:$L$3007))=${YEAR})*(${MTL}!$N$8:$N$3007="Complete"))`;
    const rateFml  = `=IFERROR(F${row1}/E${row1},0)`;
    vals.push({ range: `${S}!E${row1}:H${row1}`, values: [[mon, dueFml, doneFml, rateFml]] });
    fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 7, 8), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });

  // ── Freeze 3 header rows ───────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Row heights ────────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 60 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 8 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 8 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 8, endIndex: 28 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 28, endIndex: 29 }, properties: { pixelSize: 8 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 30, endIndex: 31 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 31, endIndex: 42 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 42, endIndex: 43 }, properties: { pixelSize: 8 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 43, endIndex: 44 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 44, endIndex: 58 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // ── Column widths ──────────────────────────────────────────────────────────
  const COL_WIDTHS = [80, 200, 90, 120, 90, 70, 80, 80];
  COL_WIDTHS.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Column separator between left/right panels in lower sections
  fmt.push({ updateBorders: { range: gridRange(SID, 29, 58, 4, 5), left: { style: 'SOLID_MEDIUM', color: hex(C.border) } } });

  await batchUpdate(id, fmt, 'db-fmt');
  await valuesBatchUpdate(id, vals, 'db-vals');
  console.log('✓ Project Dashboard complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
