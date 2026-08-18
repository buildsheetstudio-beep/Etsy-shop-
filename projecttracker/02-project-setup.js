'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Project Setup'];
const S = "'Project Setup'";
const MTL = "'Master Task Log'";
const MIL = "'Milestones & Progress'";
const NC = 26;

// ── Layout: Row 0 title | Row 1 subtitle | Rows 2-3 controls | Rows 4-5 KPI cards
//            Row 6 col headers | Rows 7-506 data (1-indexed rows 8-507)

(async () => {
  const fmt = [];
  const vals = [];

  // Title banner
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['📋 PROJECT SETUP & PORTFOLIO OVERVIEW']] });

  // Subtitle
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.teal), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [['Register projects, view portfolio health, and configure view controls for all tabs']] });

  // ── Controls (rows 2-3, 0-indexed) ────────────────────────────────────────
  const controls = [
    ['Selected Project ID',    ''],
    ['Selected Owner / All',   'All'],
    ['Selected Status / All',  'All'],
    ['Reporting Year',         '2026'],
    ['Reporting Month / Full Year', 'Full Year'],
    ['Week Start Date',        '=DATE(2026,8,18)'],
    ['Calendar Month',         '8'],
    ['Calendar Year',          '2026'],
  ];
  controls.forEach(([label, value], i) => {
    const col = (i % 4) * 6;
    const labelRow = 2 + Math.floor(i / 4) * 2;
    const valueRow = labelRow + 1;
    fmt.push({ mergeCells: { range: gridRange(SID, labelRow, labelRow+1, col, col+6), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, valueRow, valueRow+1, col, col+6), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, labelRow, labelRow+1, col, col+6), cell: { userEnteredFormat: { backgroundColor: hex(C.terra), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ repeatCell: { range: gridRange(SID, valueRow, valueRow+1, col, col+6), cell: { userEnteredFormat: { backgroundColor: hex(C.input), textFormat: { fontSize: 10, fontFamily: 'Arial' }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    const colLetter = String.fromCharCode(65 + col);
    vals.push({ range: `${S}!${colLetter}${labelRow+1}`, values: [[label]] });
    vals.push({ range: `${S}!${colLetter}${valueRow+1}`, values: [[value]] });
  });
  // Date format for Week Start Date
  fmt.push({ repeatCell: { range: gridRange(SID, 5, 6, 18, 24), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // ── Column Headers (row 6, 0-indexed) ─────────────────────────────────────
  const COLS = ['Project ID','Project Name','Category','Project Owner','Client/Stakeholder',
    'Priority','Status','Start Date','Target End Date','Actual End Date',
    'Total Tasks','Completed Tasks','Task Compl%','Total Milestones','Compl Milestones',
    'MS Compl%','Overdue Tasks','Blocked Tasks','Est Hours','Actual Hours',
    'Days Remaining','Schedule Variance','Project Health','Next Milestone','Next Task Due','Notes'];
  fmt.push({ repeatCell: { range: gridRange(SID, 6, 7, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)' } });
  vals.push({ range: `${S}!A7`, values: [COLS] });

  // ── Freeze rows 1-6 and columns A:C ───────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Data Validation ────────────────────────────────────────────────────────
  const rdRef = (listName) => `'Reference Data'!$A$1:$A$100`; // placeholder — use actual range
  const dv = (col, source) => ({
    setDataValidation: {
      range: gridRange(SID, 7, 507, col, col+1),
      rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: source }] },
        showCustomUi: true, strict: true,
      },
    },
  });
  // Col 2 = Category, Col 3 = Owner (assignees), Col 5 = Priority, Col 6 = Status
  // Compute reference ranges from the LISTS in 01-reference.js
  // Since reference data is written in positional columns, use named range or known positions.
  // Safer: just use ONE_OF_LIST with values
  const dvList = (col, items) => ({
    setDataValidation: {
      range: gridRange(SID, 7, 507, col, col+1),
      rule: {
        condition: { type: 'ONE_OF_LIST', values: items.map(v => ({ userEnteredValue: v })) },
        showCustomUi: true, strict: false,
      },
    },
  });
  fmt.push(dvList(2, ['Client Project','Business Operations','Marketing','Product Development','Finance','Administrative','Personal','Creative','Event','Education','Home','Other']));
  fmt.push(dvList(3, ['Sarah Chen','Marcus Webb','Olivia Park','Tyler Banks','Zara Osei','Self','TBD']));
  fmt.push(dvList(5, ['Low','Medium','High','Critical']));
  fmt.push(dvList(6, ['Not Started','Planning','Active','On Hold','At Risk','Complete','Cancelled','Archived']));

  // ── Formula rows ──────────────────────────────────────────────────────────
  const healthFml = (r) =>
    `=IFERROR(IF(G${r}="Complete","Complete",IF(K${r}=0,"Insufficient Data",IF(OR(IFERROR(R${r}/K${r},0)>=0.4,R${r}>=3),"At Risk",IF(AND(Q${r}>0,I${r}<TODAY()),"Delayed",IF(OR(Q${r}>0,R${r}>0),"Monitor","On Track"))))),"Insufficient Data")`;

  const fmlCols = [
    // [colIdx, formula generator]
    [0,  r => `=IF(B${r}="","","PRJ-"&TEXT(ROW()-7,"000"))`],
    [10, r => `=IFERROR(COUNTIFS(${MTL}!$B$8:$B$3007,$A${r}),0)`],
    [11, r => `=IFERROR(COUNTIFS(${MTL}!$B$8:$B$3007,$A${r},${MTL}!$N$8:$N$3007,"Complete"),0)`],
    [12, r => `=IFERROR(L${r}/K${r},0)`],
    [13, r => `=IFERROR(COUNTIFS(${MIL}!$B$8:$B$1007,$A${r}),0)`],
    [14, r => `=IFERROR(COUNTIFS(${MIL}!$B$8:$B$1007,$A${r},${MIL}!$H$8:$H$1007,"Achieved"),0)`],
    [15, r => `=IFERROR(O${r}/N${r},0)`],
    [16, r => `=IFERROR(COUNTIFS(${MTL}!$B$8:$B$3007,$A${r},${MTL}!$M$8:$M$3007,"<"&TODAY(),${MTL}!$N$8:$N$3007,"<>Complete",${MTL}!$N$8:$N$3007,"<>Cancelled"),0)`],
    [17, r => `=IFERROR(COUNTIFS(${MTL}!$B$8:$B$3007,$A${r},${MTL}!$N$8:$N$3007,"Blocked"),0)`],
    [18, r => `=IFERROR(SUMIFS(${MTL}!$P$8:$P$3007,${MTL}!$B$8:$B$3007,$A${r}),0)`],
    [19, r => `=IFERROR(SUMIFS(${MTL}!$Q$8:$Q$3007,${MTL}!$B$8:$B$3007,$A${r}),0)`],
    [20, r => `=IF(I${r}="","",IF(G${r}="Complete",0,I${r}-TODAY()))`],
    [21, r => `=IFERROR(IF(G${r}="Complete",J${r}-I${r},IF(I${r}<TODAY(),TODAY()-I${r},"")),"—")`],
    [22, r => healthFml(r)],
    [23, r => `=IFERROR(INDEX(${MIL}!$D$8:$D$1007,MATCH(MINIFS(${MIL}!$F$8:$F$1007,${MIL}!$B$8:$B$1007,$A${r},${MIL}!$H$8:$H$1007,"<>Achieved",${MIL}!$H$8:$H$1007,"<>Cancelled",${MIL}!$F$8:$F$1007,">="&TODAY()),${MIL}!$F$8:$F$1007,0)),"—")`],
    [24, r => `=IFERROR(INDEX(${MTL}!$F$8:$F$3007,MATCH(MINIFS(${MTL}!$M$8:$M$3007,${MTL}!$B$8:$B$3007,$A${r},${MTL}!$N$8:$N$3007,"<>Complete",${MTL}!$N$8:$N$3007,"<>Cancelled",${MTL}!$M$8:$M$3007,">="&TODAY()),${MTL}!$M$8:$M$3007,0)),"—")`],
  ];

  // Write formulas for all 500 rows
  fmlCols.forEach(([colIdx, fmlGen]) => {
    const colLetter = String.fromCharCode(65 + colIdx);
    const formulas = Array.from({ length: 500 }, (_, i) => [fmlGen(8 + i)]);
    vals.push({ range: `${S}!${colLetter}8:${colLetter}507`, values: formulas });
  });

  // ── Number formats ────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 507, 7, 10), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });  // dates cols H-J
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 507, 12, 13), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });  // Task compl%
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 507, 15, 16), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });  // MS compl%
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 507, 18, 20), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0' } } }, fields: 'userEnteredFormat.numberFormat' } });    // hours
  // Mark formula columns with formula background
  const fmlColIdxs = [0,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24];
  fmlColIdxs.forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID, 7, 507, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // Alternating row backgrounds
  for (let r = 0; r < 500; r++) {
    if (r % 2 !== 0) {
      fmt.push({ repeatCell: { range: gridRange(SID, 7+r, 8+r, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
    }
  }

  // Row heights
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 55 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // Column widths
  const colWidths = [80,200,120,100,120,70,90,90,90,90,70,80,70,70,80,65,70,65,65,65,75,75,100,160,160,180];
  colWidths.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, fmt, 'ps-fmt');

  // ── Sample Projects ────────────────────────────────────────────────────────
  // Columns: B(name), C(cat), D(owner), E(client), F(priority), G(status), H(start), I(end), J(actual), Z(notes)
  const projects = [
    ['Riverfront Apartment App','Client Project','Sarah Chen','Apex Realty','High','Active','1/15/2026','9/30/2026','','Feature build for a property listing mobile app'],
    ['Annual Marketing Rebrand','Marketing','Olivia Park','Internal','Medium','Planning','8/1/2026','12/31/2026','','Full brand refresh: logo, colors, tone of voice'],
    ['Nova X200 Product Launch','Product Development','Marcus Webb','Internal','Critical','Active','3/1/2026','10/15/2026','','Hardware launch including website, PR, and fulfillment'],
    ['HR Policy Update 2026','Administrative','Zara Osei','Internal','Medium','Complete','1/10/2026','5/30/2026','6/5/2026','Annual review and update of all HR policies'],
    ['E-commerce Platform Migration','Business Operations','Tyler Banks','Nexus Corp','High','At Risk','5/1/2026','8/31/2026','','Migrate from legacy platform to Shopify Plus'],
    ['Q3 Financial Audit','Finance','Sarah Chen','Internal','High','Active','7/1/2026','9/15/2026','','Quarterly audit of revenue, expenses, and cash flow'],
    ['Employee Wellness Day','Event','Olivia Park','Internal','Low','On Hold','6/1/2026','9/20/2026','','Half-day wellness event for all staff'],
    ['Mobile App Redesign','Product Development','Marcus Webb','Internal','High','Active','4/1/2026','11/30/2026','','Full UX/UI redesign of the customer-facing mobile app'],
    ['Client Onboarding System','Client Project','Tyler Banks','Vertex Solutions','High','Active','2/15/2026','9/30/2026','','Automated onboarding workflow for new enterprise clients'],
    ['Website SEO Overhaul','Marketing','Zara Osei','Internal','Medium','Planning','8/15/2026','12/15/2026','','Comprehensive SEO audit and site restructuring'],
    ['Home Office Renovation','Home','Self','Internal','Low','Active','7/20/2026','9/30/2026','','New desk, lighting, shelving, and cable management'],
    ['Online Course Development','Education','Sarah Chen','Internal','Medium','Planning','9/1/2026','3/31/2027','','Develop a 12-module project management certificate course'],
    ['Photography Portfolio Site','Creative','Self','Internal','Low','Active','7/1/2026','10/31/2026','','Build and launch a personal photography portfolio website'],
    ['Tax Filing 2025','Finance','Self','Internal','High','Complete','1/3/2026','4/15/2026','4/12/2026','Annual personal and business tax return preparation'],
    ['Team Training Workshop','Administrative','Olivia Park','Internal','Medium','Active','8/1/2026','9/15/2026','','Two-day project management skills workshop for team leads'],
    ['Product Catalog Update','Business Operations','Tyler Banks','Internal','Medium','Active','7/15/2026','9/30/2026','','Update 400+ product listings with new specs, images, prices'],
    ['Social Media Strategy Q3','Marketing','Zara Osei','Internal','Medium','Active','7/1/2026','9/30/2026','','Q3 content calendar, campaigns, and analytics setup'],
    ['Inventory Audit Q3','Business Operations','Marcus Webb','Internal','Low','On Hold','8/1/2026','8/31/2026','','Physical and system inventory count — paused pending ERP update'],
    ['Brand Video Production','Creative','Sarah Chen','Apex Realty','High','At Risk','5/1/2026','8/15/2026','','2-minute brand overview video; delayed due to location issues'],
    ['Supplier Contract Renewal','Administrative','Tyler Banks','Internal','Medium','Planning','9/1/2026','10/31/2026','','Renew and renegotiate contracts with 6 key suppliers'],
  ];

  const projectVals = projects.map((p, i) => {
    const row = Array(NC).fill('');
    // A (col 0) = formula already loaded; write B through Z
    row[1] = p[0]; // B name
    row[2] = p[1]; // C category
    row[3] = p[2]; // D owner
    row[4] = p[3]; // E client
    row[5] = p[4]; // F priority
    row[6] = p[5]; // G status
    row[7] = p[6]; // H start
    row[8] = p[7]; // I target end
    row[9] = p[8]; // J actual end
    row[25] = p[9]; // Z notes
    return row;
  });
  vals.push({ range: `${S}!A8`, values: projectVals });

  await valuesBatchUpdate(id, vals, 'ps-vals');
  console.log('✓ Project Setup complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
