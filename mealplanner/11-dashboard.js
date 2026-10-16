'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DSH = sheetMap['📊 Dashboard'];

(async () => {
  // ── Values ────────────────────────────────────────────────────────────────
  const data = [
    { range: "'📊 Dashboard'!A1", values: [['🌿 Ultimate Meal Planner+']] },
    { range: "'📊 Dashboard'!A2", values: [['Track meals, grocery budget, nutrition & batch prep — all in one place']] },
    // Section: Quick Stats
    { range: "'📊 Dashboard'!A4", values: [['📊 This Week at a Glance']] },
    // KPI labels row 1
    { range: "'📊 Dashboard'!A5", values: [['🍽 Meals Planned','','💰 Budget Left ($)','','🛒 Items to Buy','',' ⚠️ Expiring Soon','']] },
    // KPI values row 1
    { range: "'📊 Dashboard'!A6", values: [[
      "=COUNTIF('📅 Weekly Meal Plan'!D3:D23,\"<>\")", '',
      "=IFERROR('💰 Grocery Budget'!B8-'💰 Grocery Budget'!C8,0)", '',
      "=COUNTIFS('🛒 Grocery List'!D3:D27,FALSE,'🛒 Grocery List'!G3:G27,FALSE)", '',
      "=COUNTIF('🥗 Food Inventory'!G3:G22,\"Use Soon\")", '',
    ]] },
    // KPI labels row 2
    { range: "'📊 Dashboard'!A7", values: [['🔥 Avg Daily Calories','','🥗 Prep Portions Ready','','❌ Expired Items','','✅ Prepped in Advance','']] },
    // KPI values row 2
    { range: "'📊 Dashboard'!A8", values: [[
      "=IFERROR(ROUND(SUM('📅 Weekly Meal Plan'!E3:E23)/7,0),0)", '',
      "=IFERROR(SUM('🍳 Meal Prep Schedule'!F3:F12),0)", '',
      "=COUNTIF('🥗 Food Inventory'!G3:G22,\"Expired\")", '',
      "=COUNTIF('📅 Weekly Meal Plan'!F3:F23,TRUE)", '',
    ]] },

    // Section: Budget Summary
    { range: "'📊 Dashboard'!A10", values: [['💰 Budget & Grocery Summary']] },
    { range: "'📊 Dashboard'!A11", values: [['Week Of','Budget','Actual Spend','Cost Per Meal','Variance']] },
    { range: "'📊 Dashboard'!A12", values: [["='💰 Grocery Budget'!A8","='💰 Grocery Budget'!B8","='💰 Grocery Budget'!C8","='💰 Grocery Budget'!D8","='💰 Grocery Budget'!E8"]] },
    { range: "'📊 Dashboard'!A13", values: [["='💰 Grocery Budget'!A7","='💰 Grocery Budget'!B7","='💰 Grocery Budget'!C7","='💰 Grocery Budget'!D7","='💰 Grocery Budget'!E7"]] },
    { range: "'📊 Dashboard'!A14", values: [["='💰 Grocery Budget'!A6","='💰 Grocery Budget'!B6","='💰 Grocery Budget'!C6","='💰 Grocery Budget'!D6","='💰 Grocery Budget'!E6"]] },

    // Section: Inventory Alerts
    { range: "'📊 Dashboard'!A16", values: [['🥗 Food Inventory Alerts']] },
    { range: "'📊 Dashboard'!A17", values: [['Item','Status','Expiration Date','Days Until Exp']] },
    // Pull items that are Expired or Use Soon from Food Inventory (static view of top concern items)
    { range: "'📊 Dashboard'!A18", values: [
      ["Bell Peppers","Expired","2026-07-12",`=DATEVALUE("2026-07-12")-TODAY()`],
      ["Lemon","Expired","2026-07-07",`=DATEVALUE("2026-07-07")-TODAY()`],
      ["Spinach","Use Soon","2026-07-14",`=DATEVALUE("2026-07-14")-TODAY()`],
      ["Avocado","Use Soon","2026-07-14",`=DATEVALUE("2026-07-14")-TODAY()`],
      ["Salmon","Use Soon","2026-07-14",`=DATEVALUE("2026-07-14")-TODAY()`],
    ]},

    // Section: Nutrition Summary
    { range: "'📊 Dashboard'!A24", values: [['🎯 Nutrition Summary — This Week']] },
    { range: "'📊 Dashboard'!A25", values: [['Family Member','Calorie Goal','Actual Calories','Variance %','Protein Goal','Actual Protein']] },
    { range: "'📊 Dashboard'!A26", values: [
      ["='🎯 Nutrition Goals'!A6","='🎯 Nutrition Goals'!C6","='🎯 Nutrition Goals'!D6","='🎯 Nutrition Goals'!G6","='🎯 Nutrition Goals'!E6","='🎯 Nutrition Goals'!F6"],
      ["='🎯 Nutrition Goals'!A5","='🎯 Nutrition Goals'!C5","='🎯 Nutrition Goals'!D5","='🎯 Nutrition Goals'!G5","='🎯 Nutrition Goals'!E5","='🎯 Nutrition Goals'!F5"],
      ["='🎯 Nutrition Goals'!A4","='🎯 Nutrition Goals'!C4","='🎯 Nutrition Goals'!D4","='🎯 Nutrition Goals'!G4","='🎯 Nutrition Goals'!E4","='🎯 Nutrition Goals'!F4"],
    ]},
  ];
  await valuesBatchUpdate(id, data, 'dsh-values');

  // ── Formatting ────────────────────────────────────────────────────────────
  const reqs = [];

  // Merge A1:H1 (title)
  reqs.push({ mergeCells: { range: gridRange(DSH, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  // Merge A2:H2 (subtitle)
  reqs.push({ mergeCells: { range: gridRange(DSH, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } });
  // Merge A3:H3 (spacer)
  reqs.push({ mergeCells: { range: gridRange(DSH, 2, 3, 0, 8), mergeType: 'MERGE_ALL' } });
  // Merge A4:H4 (section header)
  reqs.push({ mergeCells: { range: gridRange(DSH, 3, 4, 0, 8), mergeType: 'MERGE_ALL' } });
  // KPI label/value cells: merge pairs A-B, C-D, E-F, G-H in rows 5-8
  for (let r = 4; r <= 7; r++) {
    for (let c = 0; c < 8; c += 2) {
      reqs.push({ mergeCells: { range: gridRange(DSH, r, r + 1, c, c + 2), mergeType: 'MERGE_ALL' } });
    }
  }
  // Merge A10:H10 (budget section)
  reqs.push({ mergeCells: { range: gridRange(DSH, 9, 10, 0, 8), mergeType: 'MERGE_ALL' } });
  // Merge A15:H15 (spacer)
  reqs.push({ mergeCells: { range: gridRange(DSH, 14, 15, 0, 8), mergeType: 'MERGE_ALL' } });
  // Merge A16:H16 (inventory section)
  reqs.push({ mergeCells: { range: gridRange(DSH, 15, 16, 0, 8), mergeType: 'MERGE_ALL' } });
  // Merge A23:H23 (spacer)
  reqs.push({ mergeCells: { range: gridRange(DSH, 22, 23, 0, 8), mergeType: 'MERGE_ALL' } });
  // Merge A24:H24 (nutrition section)
  reqs.push({ mergeCells: { range: gridRange(DSH, 23, 24, 0, 8), mergeType: 'MERGE_ALL' } });

  // Title row: deep basil bg, white bold large
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 0, 1, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepBasil), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});

  // Subtitle: goldenMustard bg
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 1, 2, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.goldenMustard), textFormat: { foregroundColor: hex(C.white), bold: false, fontSize: 11 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  // Spacer rows: warmCream
  [2, 8, 14, 22].forEach(r => reqs.push({ repeatCell: {
    range: gridRange(DSH, r, r + 1, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.warmCream) } },
    fields: 'userEnteredFormat.backgroundColor',
  }}));

  // Section header row 4: warmOrange
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 3, 4, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.warmOrange), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  // KPI label rows (5 and 7): lightAmber bg, dark bold
  [4, 6].forEach(r => reqs.push({ repeatCell: {
    range: gridRange(DSH, r, r + 1, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.lightAmber), textFormat: { foregroundColor: hex(C.darkText), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }}));

  // KPI value rows (6 and 8): white bg, deepBasil text, large bold
  [5, 7].forEach(r => reqs.push({ repeatCell: {
    range: gridRange(DSH, r, r + 1, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.white), textFormat: { foregroundColor: hex(C.deepBasil), bold: true, fontSize: 20 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }}));

  // KPI row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 4, endIndex: 9 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});
  // KPI value rows taller
  [5, 7].forEach(r => reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: r, endIndex: r + 1 },
    properties: { pixelSize: 50 }, fields: 'pixelSize',
  }}));

  // Budget section header (row 10)
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 9, 10, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.warmOrange), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 9, endIndex: 10 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  // Budget table header row 11
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 10, 11, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.darkText), textFormat: { foregroundColor: hex(C.white), bold: true }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  // Budget data rows 12-14
  for (let r = 0; r < 3; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(DSH, 11 + r, 12 + r, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.warmCream : C.altRow), horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)',
    }});
  }
  // Currency format budget cols B-E (idx 1-4) rows 12-14
  [1,2,3,4].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(DSH, 11, 14, c, c + 1),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } },
    fields: 'userEnteredFormat.numberFormat',
  }}));

  // Inventory section header (row 16)
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 15, 16, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.mutedSage), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 15, endIndex: 16 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  // Inventory table header row 17
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 16, 17, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.darkText), textFormat: { foregroundColor: hex(C.white), bold: true }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  // Inventory data rows 18-22
  for (let r = 0; r < 5; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(DSH, 17 + r, 18 + r, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.warmCream : C.altRow), horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)',
    }});
  }

  // Nutrition section header (row 24)
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 23, 24, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepBasil), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 23, endIndex: 24 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  // Nutrition table header row 25
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 24, 25, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.darkText), textFormat: { foregroundColor: hex(C.white), bold: true }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  // Nutrition data rows 26-28
  for (let r = 0; r < 3; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(DSH, 25 + r, 26 + r, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.warmCream : C.altRow), horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)',
    }});
  }
  // Percent format col D in nutrition rows (col index 3)
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 25, 28, 3, 4),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});

  // Background fill: all remaining cells warm cream
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 0, 1000, 0, 26),
    cell: { userEnteredFormat: { backgroundColor: hex(C.warmCream) } },
    fields: 'userEnteredFormat.backgroundColor',
  }});
  // Re-apply specific formatting on top (ordering matters — general first, specific after)
  // Actually just set background on the warmCream first overall... it's already first above.
  // That's fine because later repeatCell calls overwrite.

  // Column widths: 8 cols roughly equal
  const widths = [130, 110, 110, 110, 110, 110, 110, 110];
  widths.forEach((px, c) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    },
  }));

  await batchUpdate(id, reqs, 'dsh-format');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
