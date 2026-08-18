'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Home Maintenance Dashboard'];
const S = "'Home Maintenance Dashboard'";
const SETUP = "'Home & Property Setup'";
const TASK  = "'Maintenance Task Log'";
const COST  = "'Maintenance Costs'";
const REP   = "'Home Repair Log'";
const REC   = "'Recurring Maintenance'";
const AST   = "'Appliances & Systems'";
const NC = 14; // A-N

(async () => {
  const fmt = [];
  const vals = [];

  // ── Main title ────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['Home Maintenance Dashboard']] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  // ── Subtitle / property name ───────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, 8), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.white), italic: true }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [[`=IFERROR(${SETUP}!C6&" · "&${SETUP}!D6&" · Built "&TEXT(${SETUP}!E6,"0"),"Home Maintenance Dashboard")`]] });

  // ── Date updated ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 8, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 8, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!I2`, values: [['="As of " & TEXT(TODAY(),"mmmm d, yyyy")']] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // ── Disclaimer ────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.clay), textFormat: { fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text), italic: true }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A3`, values: [['DISCLAIMER: This tracker is for personal planning only. It is not a substitute for professional inspection, licensed contractor assessment, or official building code compliance. All cost estimates are illustrative examples — verify for your actual property.']] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // ── Divider ───────────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.wheat) } }, fields: 'userEnteredFormat.backgroundColor' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 4 }, fields: 'pixelSize' } });

  // ──────────────────────────────────────────────────────────────────────────────
  // ROW 4 (0-indexed) — KPI section label
  // ──────────────────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 4, 5, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.olive), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A5`, values: [['  TASK OVERVIEW']] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 20 }, fields: 'pixelSize' } });

  // ── KPI label row (row 5, 0-indexed) ─────────────────────────────────────────
  const kpis = [
    { label: 'Total Tasks',      col: 0, span: 2, color: C.primary },
    { label: 'Complete',         col: 2, span: 2, color: C.success },
    { label: 'In Progress',      col: 4, span: 2, color: C.info },
    { label: 'Scheduled',        col: 6, span: 2, color: C.wheat },
    { label: 'Not Started',      col: 8, span: 2, color: C.gray },
    { label: 'Overdue',          col: 10, span: 2, color: C.attention },
    { label: '% Complete',       col: 12, span: 2, color: C.secondary },
  ];
  for (const k of kpis) {
    fmt.push({ mergeCells: { range: gridRange(SID, 5, 6, k.col, k.col + k.span), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 5, 6, k.col, k.col + k.span), cell: { userEnteredFormat: { backgroundColor: hex(k.color), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ mergeCells: { range: gridRange(SID, 6, 7, k.col, k.col + k.span), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 6, 7, k.col, k.col + k.span), cell: { userEnteredFormat: { backgroundColor: hex(C.bg), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  }

  const kpiFormulas = [
    { col: 'A', formula: `=COUNTA(${TASK}!A7:A1000)`, fmt: '#,##0' },
    { col: 'C', formula: `=SUMPRODUCT((${TASK}!H7:H1000="Complete")*1)`, fmt: '#,##0' },
    { col: 'E', formula: `=SUMPRODUCT((${TASK}!H7:H1000="In Progress")*1)`, fmt: '#,##0' },
    { col: 'G', formula: `=SUMPRODUCT((${TASK}!H7:H1000="Scheduled")*1)`, fmt: '#,##0' },
    { col: 'I', formula: `=SUMPRODUCT((${TASK}!H7:H1000="Not Started")*1)`, fmt: '#,##0' },
    { col: 'K', formula: `=SUMPRODUCT((${TASK}!T7:T1000=TRUE)*1)`, fmt: '#,##0' },
    { col: 'M', formula: `=IFERROR(SUMPRODUCT((${TASK}!H7:H1000="Complete")*1)/COUNTA(${TASK}!A7:A1000),0)`, fmt: '0%' },
  ];

  for (const { col, formula, fmt: nfmt } of kpiFormulas) {
    vals.push({ range: `${S}!${col}7`, values: [[formula]] });
    const ci = col.charCodeAt(0) - 65;
    const kpiK = kpis.find(k => k.col === ci);
    const span = kpiK ? kpiK.span : 2;
    fmt.push({ repeatCell: { range: gridRange(SID, 6, 7, ci, ci + span), cell: { userEnteredFormat: { numberFormat: { type: nfmt === '#,##0' ? 'NUMBER' : 'PERCENT', pattern: nfmt } } }, fields: 'userEnteredFormat.numberFormat' } });
  }

  for (const [ri, px] of [[5,20],[6,44]]) {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  }

  // ── Divider ───────────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 8, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.olive) } }, fields: 'userEnteredFormat.backgroundColor' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 8 }, properties: { pixelSize: 4 }, fields: 'pixelSize' } });

  // ──────────────────────────────────────────────────────────────────────────────
  // ROW 8 — COST OVERVIEW label
  // ──────────────────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 8, 9, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 8, 9, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.olive), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A9`, values: [['  COST OVERVIEW']] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 8, endIndex: 9 }, properties: { pixelSize: 20 }, fields: 'pixelSize' } });

  // ── Cost KPI row ─────────────────────────────────────────────────────────────
  const costKpis = [
    { label: 'Total Paid (All Time)', col: 0, span: 3, color: C.primary },
    { label: `Total ${new Date().getFullYear()} Costs`, col: 3, span: 3, color: C.secondary },
    { label: 'Total Planned', col: 6, span: 3, color: C.warning },
    { label: 'DIY Savings', col: 9, span: 3, color: C.success },
    { label: 'Reimbursable', col: 12, span: 2, color: C.info },
  ];
  for (const k of costKpis) {
    fmt.push({ mergeCells: { range: gridRange(SID, 9, 10, k.col, k.col + k.span), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 9, 10, k.col, k.col + k.span), cell: { userEnteredFormat: { backgroundColor: hex(k.color), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ mergeCells: { range: gridRange(SID, 10, 11, k.col, k.col + k.span), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 10, 11, k.col, k.col + k.span), cell: { userEnteredFormat: { backgroundColor: hex(C.bg), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' }, userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' }, backgroundColor: hex(C.bg), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(numberFormat,backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  }

  const costKpiFormulas = [
    { col: 'A', formula: `=SUMPRODUCT((${COST}!O7:O1000="Paid")*(ISNUMBER(${COST}!M7:M1000))*${COST}!M7:M1000)` },
    { col: 'D', formula: `=SUMPRODUCT((${COST}!O7:O1000="Paid")*(${COST}!P7:P1000=${new Date().getFullYear()})*(ISNUMBER(${COST}!M7:M1000))*${COST}!M7:M1000)` },
    { col: 'G', formula: `=SUMPRODUCT((${COST}!O7:O1000="Planned")*(ISNUMBER(${COST}!M7:M1000))*${COST}!M7:M1000)` },
    { col: 'J', formula: `=SUMPRODUCT((${COST}!N7:N1000="DIY")*(${COST}!O7:O1000="Paid")*(ISNUMBER(${COST}!M7:M1000))*${COST}!M7:M1000)` },
    { col: 'M', formula: `=SUMPRODUCT((${COST}!S7:S1000=TRUE)*(ISNUMBER(${COST}!M7:M1000))*${COST}!M7:M1000)` },
  ];
  for (const { col, formula } of costKpiFormulas) {
    vals.push({ range: `${S}!${col}11`, values: [[formula]] });
  }

  for (const [ri, px] of [[9,20],[10,44]]) {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  }

  // ── Divider ───────────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 11, 12, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.olive) } }, fields: 'userEnteredFormat.backgroundColor' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 11, endIndex: 12 }, properties: { pixelSize: 4 }, fields: 'pixelSize' } });

  // ──────────────────────────────────────────────────────────────────────────────
  // ROWS 12-27: Upcoming Tasks (left) + Cost by Category (right)
  // ──────────────────────────────────────────────────────────────────────────────
  // Section label row (row 12)
  fmt.push({ mergeCells: { range: gridRange(SID, 12, 13, 0, 7), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 12, 13, 0, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A13`, values: [['UPCOMING TASKS (Next 90 Days)']] });

  fmt.push({ mergeCells: { range: gridRange(SID, 12, 13, 7, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 12, 13, 7, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!H13`, values: [[`COST BY CATEGORY — ${new Date().getFullYear()}`]] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 12, endIndex: 13 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // Sub-header row (row 13)
  const upcomingHdr = ['#','Task Name','Category','Priority','Due Date','Status','DIY/Pro'];
  const catHdr = ['Category','Tasks Done','Cost Paid','Planned'];
  fmt.push({ repeatCell: { range: gridRange(SID, 13, 14, 0, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A14:G14`, values: [upcomingHdr] });
  fmt.push({ repeatCell: { range: gridRange(SID, 13, 14, 7, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!H14:K14`, values: [catHdr] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 13, endIndex: 14 }, properties: { pixelSize: 18 }, fields: 'pixelSize' } });

  // Upcoming task rows (rows 14-27, 0-indexed = row 15-28 in spreadsheet)
  // Filter tasks where due date <= TODAY()+90 AND status not Complete/Cancelled/Skipped
  const upcomingFilter = `(${TASK}!K7:K1000<>"")*` +
    `(${TASK}!K7:K1000<=TODAY()+90)*` +
    `(${TASK}!H7:H1000<>"Complete")*` +
    `(${TASK}!H7:H1000<>"Cancelled")*` +
    `(${TASK}!H7:H1000<>"Skipped")`;

  const upcomingCols = [
    `${TASK}!C7:C1000`, // name
    `${TASK}!D7:D1000`, // category
    `${TASK}!G7:G1000`, // priority
    `${TASK}!K7:K1000`, // due date
    `${TASK}!H7:H1000`, // status
    `${TASK}!M7:M1000`, // diy/pro
  ];

  for (let ri = 0; ri < 14; ri++) {
    const n = ri + 1;
    const rowIdx = 14 + ri; // 0-indexed
    const bg = ri % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, rowIdx, rowIdx+1, 0, 7), cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

    // Row number
    vals.push({ range: `${S}!A${rowIdx+1}`, values: [[`=IFERROR(IF(IFERROR(INDEX(FILTER(${TASK}!C7:C1000,${upcomingFilter}),${n}),"")="","",${n}),"")`]] });
    // Name col B
    vals.push({ range: `${S}!B${rowIdx+1}`, values: [[`=IFERROR(INDEX(FILTER(${upcomingCols[0]},${upcomingFilter}),${n}),"")`]] });
    // Category col C
    vals.push({ range: `${S}!C${rowIdx+1}`, values: [[`=IFERROR(INDEX(FILTER(${upcomingCols[1]},${upcomingFilter}),${n}),"")`]] });
    // Priority col D
    vals.push({ range: `${S}!D${rowIdx+1}`, values: [[`=IFERROR(INDEX(FILTER(${upcomingCols[2]},${upcomingFilter}),${n}),"")`]] });
    // Due Date col E
    vals.push({ range: `${S}!E${rowIdx+1}`, values: [[`=IFERROR(INDEX(FILTER(${upcomingCols[3]},${upcomingFilter}),${n}),"")`]] });
    // Status col F
    vals.push({ range: `${S}!F${rowIdx+1}`, values: [[`=IFERROR(INDEX(FILTER(${upcomingCols[4]},${upcomingFilter}),${n}),"")`]] });
    // DIY/Pro col G
    vals.push({ range: `${S}!G${rowIdx+1}`, values: [[`=IFERROR(INDEX(FILTER(${upcomingCols[5]},${upcomingFilter}),${n}),"")`]] });
  }

  // Date format for due date col E in upcoming rows
  fmt.push({ repeatCell: { range: gridRange(SID, 14, 28, 4, 5), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'M/d/yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });
  // Formula bg
  fmt.push({ repeatCell: { range: gridRange(SID, 14, 28, 0, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Cost by Category (right panel, rows 14-27, cols H-K = indices 7-10)
  const CATS = [
    'HVAC','Plumbing','Electrical','Roof & Gutters','Exterior','Interior',
    'Appliances','Safety','Landscaping','Cleaning','Pest Prevention','Water Systems','Other'
  ];
  const CY = new Date().getFullYear();
  for (let ci2 = 0; ci2 < Math.min(CATS.length, 14); ci2++) {
    const rowIdx = 14 + ci2;
    const cat = CATS[ci2];
    const bg = ci2 % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, rowIdx, rowIdx+1, 7, NC), cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

    vals.push({ range: `${S}!H${rowIdx+1}`, values: [[cat]] });
    // Tasks done this year in category
    vals.push({ range: `${S}!I${rowIdx+1}`, values: [[`=SUMPRODUCT((${TASK}!D7:D1000="${cat}")*(${TASK}!H7:H1000="Complete")*(${TASK}!W7:W1000=${CY})*1)`]] });
    // Cost paid this year in category
    vals.push({ range: `${S}!J${rowIdx+1}`, values: [[`=SUMPRODUCT((${COST}!D7:D1000="${cat}")*(${COST}!O7:O1000="Paid")*(${COST}!P7:P1000=${CY})*(ISNUMBER(${COST}!M7:M1000))*${COST}!M7:M1000)`]] });
    // Planned cost in category
    vals.push({ range: `${S}!K${rowIdx+1}`, values: [[`=SUMPRODUCT((${COST}!D7:D1000="${cat}")*(${COST}!O7:O1000="Planned")*(ISNUMBER(${COST}!M7:M1000))*${COST}!M7:M1000)`]] });
  }

  // Currency formats for cost by category
  fmt.push({ repeatCell: { range: gridRange(SID, 14, 28, 9, 11), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });

  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 14, endIndex: 28 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // ── Divider ───────────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 28, 29, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.wheat) } }, fields: 'userEnteredFormat.backgroundColor' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 28, endIndex: 29 }, properties: { pixelSize: 4 }, fields: 'pixelSize' } });

  // ──────────────────────────────────────────────────────────────────────────────
  // ROWS 29-44: Annual cost trend (left) + Asset health summary (right)
  // ──────────────────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 29, 30, 0, 7), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 29, 30, 0, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A30`, values: [['ANNUAL COST TREND']] });

  fmt.push({ mergeCells: { range: gridRange(SID, 29, 30, 7, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 29, 30, 7, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!H30`, values: [['ASSET HEALTH SUMMARY']] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 29, endIndex: 30 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // Annual cost trend sub-header
  fmt.push({ repeatCell: { range: gridRange(SID, 30, 31, 0, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A31:G31`, values: [['Year','Total Paid','DIY Cost','Pro Cost','# Repairs','# Tasks Done','Avg Task Cost']] });

  // Asset health sub-header
  fmt.push({ repeatCell: { range: gridRange(SID, 30, 31, 7, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!H31:N31`, values: [['Asset Name','Type','Condition','Warranty','Last Svc','Next Svc','Repl Year']] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 30, endIndex: 31 }, properties: { pixelSize: 18 }, fields: 'pixelSize' } });

  // Annual cost trend data rows (5 years: CY-3 to CY+1)
  for (let yi = 0; yi < 5; yi++) {
    const yr = CY - 2 + yi;
    const rowIdx = 31 + yi;
    const bg = yi % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, rowIdx, rowIdx+1, 0, 7), cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
    vals.push({ range: `${S}!A${rowIdx+1}:G${rowIdx+1}`, values: [[
      yr,
      `=SUMPRODUCT((${COST}!O7:O1000="Paid")*(${COST}!P7:P1000=${yr})*(ISNUMBER(${COST}!M7:M1000))*${COST}!M7:M1000)`,
      `=SUMPRODUCT((${COST}!N7:N1000="DIY")*(${COST}!O7:O1000="Paid")*(${COST}!P7:P1000=${yr})*(ISNUMBER(${COST}!M7:M1000))*${COST}!M7:M1000)`,
      `=SUMPRODUCT((${COST}!N7:N1000="Professional")*(${COST}!O7:O1000="Paid")*(${COST}!P7:P1000=${yr})*(ISNUMBER(${COST}!M7:M1000))*${COST}!M7:M1000)`,
      `=SUMPRODUCT((${REP}!U7:U500=${yr})*(${REP}!I7:I500="Complete")*1)`,
      `=SUMPRODUCT((${TASK}!W7:W1000=${yr})*(${TASK}!H7:H1000="Complete")*1)`,
      `=IFERROR(SUMPRODUCT((${COST}!O7:O1000="Paid")*(${COST}!P7:P1000=${yr})*(ISNUMBER(${COST}!M7:M1000))*${COST}!M7:M1000)/SUMPRODUCT((${TASK}!W7:W1000=${yr})*(${TASK}!H7:H1000="Complete")*1),0)`,
    ]] });
  }
  fmt.push({ repeatCell: { range: gridRange(SID, 31, 36, 1, 4), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 31, 36, 6, 7), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Asset health rows (right panel) — pull top 10 assets from Appliances & Systems
  for (let ai = 0; ai < 10; ai++) {
    const rowIdx = 31 + ai;
    const assetRow = 8 + ai; // spreadsheet row for asset (row 8 = first asset data row)
    const bg = ai % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, rowIdx, rowIdx+1, 7, NC), cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
    vals.push({ range: `${S}!H${rowIdx+1}:N${rowIdx+1}`, values: [[
      `=IFERROR(${AST}!C${assetRow},"")`,  // Name
      `=IFERROR(${AST}!D${assetRow},"")`,  // Type
      `=IFERROR(${AST}!P${assetRow},"")`,  // Condition
      `=IFERROR(${AST}!M${assetRow},"")`,  // Warranty Status
      `=IFERROR(TEXT(${AST}!Q${assetRow},"M/d/yyyy"),"")`,  // Last Svc
      `=IFERROR(TEXT(${AST}!R${assetRow},"M/d/yyyy"),"")`,  // Next Svc
      `=IFERROR(${AST}!O${assetRow},"")`,  // Repl Year
    ]] });
  }

  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 31, endIndex: 36 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 36, endIndex: 41 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // ── Divider ───────────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 41, 42, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.wheat) } }, fields: 'userEnteredFormat.backgroundColor' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 41, endIndex: 42 }, properties: { pixelSize: 4 }, fields: 'pixelSize' } });

  // ──────────────────────────────────────────────────────────────────────────────
  // ROWS 42-52: Monthly task completion + Active repairs
  // ──────────────────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 42, 43, 0, 7), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 42, 43, 0, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A43`, values: [[`="MONTHLY TASK COMPLETION — " & TEXT(${SETUP}!C43,"0")`]] });

  fmt.push({ mergeCells: { range: gridRange(SID, 42, 43, 7, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 42, 43, 7, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!H43`, values: [['ACTIVE REPAIRS']] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 42, endIndex: 43 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // Monthly completion sub-header
  fmt.push({ repeatCell: { range: gridRange(SID, 43, 44, 0, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A44:G44`, values: [['Month','Tasks Done','Actual Cost','DIY Count','Pro Count','Avg Cost','Overdue Count']] });

  // Active repairs sub-header
  fmt.push({ repeatCell: { range: gridRange(SID, 43, 44, 7, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!H44:N44`, values: [['Repair / Issue','Asset','Category','Status','Priority','Reported','Total Cost']] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 43, endIndex: 44 }, properties: { pixelSize: 18 }, fields: 'pixelSize' } });

  const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const CY2 = `${SETUP}!C43`;
  for (let mi = 0; mi < 12; mi++) {
    const rowIdx = 44 + mi;
    const bg = mi % 2 === 0 ? C.panel : C.altRow;
    const mn = mi + 1;
    fmt.push({ repeatCell: { range: gridRange(SID, rowIdx, rowIdx+1, 0, 7), cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
    vals.push({ range: `${S}!A${rowIdx+1}:G${rowIdx+1}`, values: [[
      MONTHS[mi],
      // Tasks Done (completed in month/year)
      `=SUMPRODUCT((${TASK}!V7:V1000=${mn})*(${TASK}!W7:W1000=${CY2})*(${TASK}!H7:H1000="Complete")*1)`,
      // Actual cost paid in month/year (from Costs tab)
      `=SUMPRODUCT((${COST}!Q7:Q1000=${mn})*(${COST}!P7:P1000=${CY2})*(${COST}!O7:O1000="Paid")*(ISNUMBER(${COST}!M7:M1000))*${COST}!M7:M1000)`,
      // DIY count
      `=SUMPRODUCT((${TASK}!V7:V1000=${mn})*(${TASK}!W7:W1000=${CY2})*(${TASK}!M7:M1000="DIY")*(${TASK}!H7:H1000="Complete")*1)`,
      // Pro count
      `=SUMPRODUCT((${TASK}!V7:V1000=${mn})*(${TASK}!W7:W1000=${CY2})*(${TASK}!M7:M1000="Professional")*(${TASK}!H7:H1000="Complete")*1)`,
      // Avg cost (cost / tasks done)
      `=IFERROR(SUMPRODUCT((${COST}!Q7:Q1000=${mn})*(${COST}!P7:P1000=${CY2})*(${COST}!O7:O1000="Paid")*(ISNUMBER(${COST}!M7:M1000))*${COST}!M7:M1000)/SUMPRODUCT((${TASK}!V7:V1000=${mn})*(${TASK}!W7:W1000=${CY2})*(${TASK}!H7:H1000="Complete")*1),0)`,
      // Overdue count (tasks with due date in this month that were not completed)
      `=SUMPRODUCT((${TASK}!V7:V1000=${mn})*(${TASK}!W7:W1000=${CY2})*(${TASK}!T7:T1000=TRUE)*1)`,
    ]] });
  }
  fmt.push({ repeatCell: { range: gridRange(SID, 44, 56, 2, 3), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 44, 56, 5, 6), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Active repairs right panel (up to 12 rows)
  const activeRepairFilter = `(${REP}!I7:I500<>"Complete")*(${REP}!I7:I500<>"Cancelled")*(${REP}!I7:I500<>"")`;
  for (let ri = 0; ri < 12; ri++) {
    const rowIdx = 44 + ri;
    const n = ri + 1;
    const bg = ri % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, rowIdx, rowIdx+1, 7, NC), cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
    vals.push({ range: `${S}!H${rowIdx+1}:N${rowIdx+1}`, values: [[
      `=IFERROR(INDEX(FILTER(${REP}!D7:D500,${activeRepairFilter}),${n}),"")`,
      `=IFERROR(INDEX(FILTER(${REP}!C7:C500,${activeRepairFilter}),${n}),"")`,
      `=IFERROR(INDEX(FILTER(${REP}!E7:E500,${activeRepairFilter}),${n}),"")`,
      `=IFERROR(INDEX(FILTER(${REP}!I7:I500,${activeRepairFilter}),${n}),"")`,
      `=IFERROR(INDEX(FILTER(${REP}!H7:H500,${activeRepairFilter}),${n}),"")`,
      `=IFERROR(TEXT(INDEX(FILTER(${REP}!J7:J500,${activeRepairFilter}),${n}),"M/d/yy"),"")`,
      `=IFERROR(INDEX(FILTER(${REP}!Q7:Q500,${activeRepairFilter}),${n}),"")`,
    ]] });
  }
  fmt.push({ repeatCell: { range: gridRange(SID, 44, 56, 13, 14), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });

  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 44, endIndex: 56 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // ── Column widths ─────────────────────────────────────────────────────────────
  const colWidths = [55,200,100,70,90,90,90,180,100,90,80,90,75,65];
  colWidths.forEach((px, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  });

  // ── Tab background ────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 100, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // ── Freeze top rows ───────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 4 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'dash-fmt');
  await valuesBatchUpdate(id, vals, 'dash-vals');

  console.log('✓ Home Maintenance Dashboard complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
