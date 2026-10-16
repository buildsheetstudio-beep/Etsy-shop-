'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Meal Cost Tracker'];
const S   = "'Meal Cost Tracker'";
const MP  = "'Weekly Meal Planner'";
const RB  = "'Recipe Book'";

const NC = 16;

(async () => {
  const fmt = [];
  const vals = [];

  // Title
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lavender), textFormat: { bold: true, fontSize: 20, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  vals.push({ range: `${S}!A1`, values: [['💰 Meal Cost Tracker']] });

  // Subtitle
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.butter), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  vals.push({ range: `${S}!A2`, values: [['Track planned meal costs by recipe, day, and week — costs are estimates based on recipe ingredient data']] });

  // Controls (rows 3-6, 0-indexed 2-5)
  const controls = [
    ['Week Start Date',        `=DATE(2026,8,18)`],
    ['Recipe / All Recipes',   'All Recipes'],
    ['Recipe Category / All',  'All'],
    ['Meal Type / All',        'All'],
  ];
  controls.forEach(([label, value], i) => {
    const r = i + 2;
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 0, 3), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial' }, horizontalAlignment: 'RIGHT' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 3, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.input), textFormat: { fontSize: 10, fontFamily: 'Arial' }, horizontalAlignment: 'LEFT' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
    vals.push({ range: `${S}!A${r+1}`, values: [[label]] });
    vals.push({ range: `${S}!D${r+1}`, values: [[value]] });
  });

  // Date format
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 3, 4), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Summary Cards (rows 7-8, 0-indexed 6-7)
  const cardDefs = [
    ['Weekly Planned Meal Cost',   `=IFERROR(SUMIFS(${MP}!$Q$12:$Q$511,${MP}!$B$12:$B$511,$D$3),0)`],
    ['Estimated Grocery Cost',     `=IFERROR('Automated Grocery List'!E10,0)`],
    ['Cost per Planned Meal',      `=IFERROR(A8/COUNTIFS(${MP}!$B$12:$B$511,$D$3,${MP}!$F$12:$F$511,"REC-*"),0)`],
    ['Cost per Planned Serving',   `=IFERROR(A8/SUMIFS(${MP}!$L$12:$L$511,${MP}!$B$12:$B$511,$D$3),0)`],
    ['Lowest-Cost Planned Recipe', `=IFERROR(INDEX(${MP}!$G$12:$G$511,MATCH(MINIFS(${MP}!$Q$12:$Q$511,${MP}!$B$12:$B$511,$D$3),${MP}!$Q$12:$Q$511,0)),"—")`],
    ['Highest-Cost Planned Recipe',`=IFERROR(INDEX(${MP}!$G$12:$G$511,MATCH(MAXIFS(${MP}!$Q$12:$Q$511,${MP}!$B$12:$B$511,$D$3),${MP}!$Q$12:$Q$511,0)),"—")`],
    ['Weekly Grocery Budget',      `=IFERROR('Household Setup'!F8,200)`],
    ['Budget Variance',            `=IFERROR(G8-B8,0)`],
  ];

  cardDefs.forEach(([label, formula], i) => {
    const col = (i % 4) * 4;
    const row = 6 + Math.floor(i / 4) * 2;
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, col, col+4), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, row+1, row+2, col, col+4), mergeType: 'MERGE_ALL' } });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, col, col+4),
        cell: { userEnteredFormat: { backgroundColor: hex(C.lavender), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row+1, row+2, col, col+4),
        cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { bold: true, fontSize: 18, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+1}`, values: [[label]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+2}`, values: [[formula]] });
  });

  // Currency formats for cost cards
  [[6,8,0,4],[6,8,4,8],[6,8,8,12],[6,8,12,16],[8,10,0,4],[8,10,4,8],[8,10,8,12],[8,10,12,16]].forEach(([r1,r2,c1,c2]) => {
    if (c1 < 12) {
      fmt.push({ repeatCell: { range: gridRange(SID, r1+1, r2, c1, c2), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
    }
  });

  // Separator row 10 (0-indexed)
  fmt.push({ mergeCells: { range: gridRange(SID, 10, 11, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 10, 11, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.lavender) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // ─── Weekly Cost Table (rows 11-18, 0-indexed 11-18) ─────────────────────
  const WCT_HDR = 11;
  const WCT_D0  = 12;
  fmt.push({
    repeatCell: {
      range: gridRange(SID, WCT_HDR, WCT_HDR+1, 0, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(C.text), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)',
    },
  });
  vals.push({ range: `${S}!A12`, values: [['Day','Breakfast Cost','Lunch Cost','Dinner Cost','Snacks/Dessert Cost','Total Planned Meal Cost','Notes']] });

  const DAYS = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
  const WEEK_START = `$D$3`;
  DAYS.forEach((day, di) => {
    const r = WCT_D0 + di;
    const dateExpr = `(${WEEK_START}+${di})`;
    const mealCostFml = (types) => `IFERROR(SUMPRODUCT((IFERROR(${MP}!$C$12:$C$511,0)=${dateExpr})*(IFERROR(MATCH(${MP}!$E$12:$E$511,{"${types.join('","')}"},0),0)>0)*(${MP}!$Q$12:$Q$511)),0)`;
    const bCost = mealCostFml(['Breakfast','Morning Snack']);
    const lCost = mealCostFml(['Lunch']);
    const dCost = mealCostFml(['Dinner']);
    const sCost = mealCostFml(['Afternoon Snack','Dessert']);
    const tCost = `IFERROR(B${r+1}+C${r+1}+D${r+1}+E${r+1},0)`;

    fmt.push({
      repeatCell: {
        range: gridRange(SID, r, r+1, 0, 7),
        cell: { userEnteredFormat: { backgroundColor: di % 2 === 0 ? hex(C.panel) : hex(C.altRow), textFormat: { fontSize: 10, fontFamily: 'Arial' } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
    vals.push({ range: `${S}!A${r+1}`, values: [[day, `=${bCost}`, `=${lCost}`, `=${dCost}`, `=${sCost}`, `=${tCost}`, '']] });
  });

  // Weekly total row
  const WTOT_ROW = WCT_D0 + 7;
  fmt.push({
    repeatCell: {
      range: gridRange(SID, WTOT_ROW, WTOT_ROW+1, 0, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lavender), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.text) } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  vals.push({ range: `${S}!A${WTOT_ROW+1}`, values: [['Weekly Total',
    `=SUM(B${WCT_D0+1}:B${WTOT_ROW})`,
    `=SUM(C${WCT_D0+1}:C${WTOT_ROW})`,
    `=SUM(D${WCT_D0+1}:D${WTOT_ROW})`,
    `=SUM(E${WCT_D0+1}:E${WTOT_ROW})`,
    `=SUM(F${WCT_D0+1}:F${WTOT_ROW})`,
    '',
  ]] });

  // Currency for columns B:F
  fmt.push({ repeatCell: { range: gridRange(SID, WCT_D0, WTOT_ROW+1, 1, 6), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Separator
  fmt.push({ mergeCells: { range: gridRange(SID, WTOT_ROW+1, WTOT_ROW+2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, WTOT_ROW+1, WTOT_ROW+2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.lavender) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // ─── Recipe Cost Table (rows after weekly table) ──────────────────────────
  const RCT_HDR = WTOT_ROW + 2;
  const RCT_D0  = RCT_HDR + 1;
  fmt.push({
    repeatCell: {
      range: gridRange(SID, RCT_HDR, RCT_HDR+1, 0, 10),
      cell: { userEnteredFormat: { backgroundColor: hex(C.text), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)',
    },
  });
  vals.push({ range: `${S}!A${RCT_HDR+1}`, values: [['Recipe ID','Recipe Name','Base\nServings','Base\nRecipe Cost','Cost/Serving','Times\nPlanned','Planned\nServings','Weekly\nPlanned Cost','Avg Cost/\nPlanned Serving','Grocery Cost\nContribution %']] });

  // Sample recipe cost rows (the recipes planned in week of Aug 18)
  const plannedRecipes = [
    'REC-0001','REC-0002','REC-0003','REC-0006','REC-0007','REC-0008',
    'REC-0009','REC-0010','REC-0013','REC-0015','REC-0017','REC-0018',
    'REC-0019','REC-0021','REC-0022','REC-0037','REC-0038','REC-0049',
  ];

  plannedRecipes.forEach((recId, i) => {
    const r1 = RCT_D0 + i;
    const weekStart = `$D$3`;
    fmt.push({
      repeatCell: {
        range: gridRange(SID, RCT_D0+i, RCT_D0+i+1, 0, 10),
        cell: { userEnteredFormat: { backgroundColor: i % 2 === 0 ? hex(C.panel) : hex(C.altRow), textFormat: { fontSize: 10, fontFamily: 'Arial' } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
    [0,3,4,7,8,9].forEach(col => {
      fmt.push({ repeatCell: { range: gridRange(SID, RCT_D0+i, RCT_D0+i+1, col, col+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
    });
    vals.push({ range: `${S}!A${r1+1}`, values: [[
      recId,
      `=IFERROR(VLOOKUP(A${r1+1},'Recipe Book'!$A$6:$B$1005,2,FALSE),"")`,  // B
      `=IFERROR(VLOOKUP(A${r1+1},'Recipe Book'!$A$6:$F$1005,6,FALSE),0)`,   // C Base servings
      `=IFERROR(VLOOKUP(A${r1+1},'Recipe Book'!$A$6:$K$1005,11,FALSE),0)`,  // D Base recipe cost
      `=IFERROR(D${r1+1}/C${r1+1},0)`,                                       // E Cost/serving
      `=IFERROR(COUNTIFS(${MP}!$F$12:$F$511,A${r1+1},${MP}!$B$12:$B$511,${weekStart}),0)`, // F Times planned
      `=IFERROR(SUMIFS(${MP}!$L$12:$L$511,${MP}!$F$12:$F$511,A${r1+1},${MP}!$B$12:$B$511,${weekStart}),0)`, // G Planned servings
      `=IFERROR(SUMIFS(${MP}!$Q$12:$Q$511,${MP}!$F$12:$F$511,A${r1+1},${MP}!$B$12:$B$511,${weekStart}),0)`, // H Weekly planned cost
      `=IFERROR(H${r1+1}/G${r1+1},0)`,                                       // I Avg cost/planned serving
      `=IFERROR(H${r1+1}/$A$8,0)`,                                           // J Grocery cost contribution %
    ]] });
  });

  // Currency for D, E, H, I
  fmt.push({ repeatCell: { range: gridRange(SID, RCT_D0, RCT_D0+plannedRecipes.length+5, 3, 5), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, RCT_D0, RCT_D0+plannedRecipes.length+5, 7, 9), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, RCT_D0, RCT_D0+plannedRecipes.length+5, 9, 10), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Column widths
  const colWidths = [80,170,70,90,80,70,80,90,90,90,90,90,90,90,90,90];
  colWidths.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });

  // Freeze rows 1:6
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 6 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'mealcost-fmt');
  await valuesBatchUpdate(id, vals, 'mealcost-vals');
  console.log('✓ Meal Cost Tracker complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
