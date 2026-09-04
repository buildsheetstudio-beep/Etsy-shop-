'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH = sheetMap['Dashboard'];

(async () => {
  const reqs = [];
  const values = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(DASH, 0, 1, 0, 16), mergeType: 'MERGE_ALL' } });
  values.push({ range: 'Dashboard!A1', values: [['🌱 ALL-SEASON GARDEN PLANNER — Season Dashboard 2025']] });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 0, 1, 0, 16),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.terracotta),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 56 }, fields: 'pixelSize',
  }});

  // ── Subtitle ──────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(DASH, 1, 2, 0, 16), mergeType: 'MERGE_ALL' } });
  values.push({ range: 'Dashboard!A2', values: [['BuildSheetStudio Premium Template  ·  Terracotta & Olive Edition  ·  2025 Growing Season']] });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 1, 2, 0, 16),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.parchment),
          textFormat: { foregroundColor: hex(C.soil), italic: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // ── KPI Cards row (row 3–5, indices 2–4) ──────────────────────────────────
  // 4 KPI cards: Total Plants, Harvest Value, Active Issues, Budget Remaining
  const kpiHeaders = ['🌿 TOTAL PLANTS', '🌾 SEASON HARVEST VALUE', '🐛 ACTIVE PEST ISSUES', '💰 BUDGET REMAINING'];
  const kpiFormulas = [
    `=COUNTA('Plant Library'!A3:A37)`,
    `=IFERROR(SUM('Harvest & Yield Tracker'!E3:E14),"$0")`,
    `=IFERROR(COUNTIF('Pest & Disease Log'!H3:H8,"Treating"),"0")`,
    `=IFERROR('Garden Budget'!H2,"$500")`,
  ];
  const kpiColors = [C.oliveLight, C.wheat, C.warningRed, C.sage];
  const kpiBgLight = ['#D8EDD4', C.paleWheat, C.warningBg, C.paleSage];

  // 4 cards across cols A-D, E-H, I-L, M-P
  [0, 1, 2, 3].forEach(k => {
    const col0 = k * 4;
    reqs.push({ mergeCells: { range: gridRange(DASH, 2, 3, col0, col0 + 4), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DASH, 3, 5, col0, col0 + 4), mergeType: 'MERGE_ALL' } });
    values.push({ range: `Dashboard!${String.fromCharCode(65 + col0)}3`, values: [[kpiHeaders[k]]] });
    values.push({ range: `Dashboard!${String.fromCharCode(65 + col0)}4`, values: [[kpiFormulas[k]]] });
    reqs.push({
      repeatCell: {
        range: gridRange(DASH, 2, 3, col0, col0 + 4),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(kpiColors[k]),
            textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    reqs.push({
      repeatCell: {
        range: gridRange(DASH, 3, 5, col0, col0 + 4),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(kpiBgLight[k]),
            textFormat: { foregroundColor: hex(kpiColors[k]), bold: true, fontSize: 22 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 3, endIndex: 5 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Currency format for harvest value and budget
  [1, 3].forEach(k => {
    const col0 = k * 4;
    reqs.push({
      repeatCell: {
        range: gridRange(DASH, 3, 5, col0, col0 + 4),
        cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } },
        fields: 'userEnteredFormat.numberFormat',
      },
    });
  });

  // ── SEASON OVERVIEW section (rows 6–13, indices 5–12) ─────────────────────
  reqs.push({ mergeCells: { range: gridRange(DASH, 5, 6, 0, 16), mergeType: 'MERGE_ALL' } });
  values.push({ range: 'Dashboard!A6', values: [['📊 SEASON OVERVIEW']] });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 5, 6, 0, 16),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.olive),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 5, endIndex: 6 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});

  // Left column: Plant status summary (A-H, cols 0-7)
  reqs.push({ mergeCells: { range: gridRange(DASH, 6, 7, 0, 8), mergeType: 'MERGE_ALL' } });
  values.push({ range: 'Dashboard!A7', values: [['🌱 Plant Status Breakdown']] });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 6, 7, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.paleSage), textFormat: { bold: true, fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  const statusItems = [
    ['Planned',       `=COUNTIF('Plant Library'!L3:L37,"Planned")`],
    ['Seeded',        `=COUNTIF('Plant Library'!L3:L37,"Seeded")`],
    ['Germinated',    `=COUNTIF('Plant Library'!L3:L37,"Germinated")`],
    ['Transplanted',  `=COUNTIF('Plant Library'!L3:L37,"Transplanted")`],
    ['Growing',       `=COUNTIF('Plant Library'!L3:L37,"Growing")`],
    ['Harvest Ready', `=COUNTIF('Plant Library'!L3:L37,"Harvest Ready")`],
    ['Harvested',     `=COUNTIF('Plant Library'!L3:L37,"Harvested")`],
  ];
  statusItems.forEach(([label, formula], i) => {
    values.push({ range: `Dashboard!A${8 + i}`, values: [[label, formula, '', `=IFERROR(SPARKLINE({B${8+i},35-B${8+i}},{"charttype","bar";"color1","${C.olive}";"color2","${C.lightGray}"}),"")`, '', '', '', ''] ]});
  });

  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 7, 7 + statusItems.length, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.cream2), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 7, 7 + statusItems.length, 0, 1),
      cell: { userEnteredFormat: { textFormat: { bold: true } } },
      fields: 'userEnteredFormat.textFormat.bold',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 7, 7 + statusItems.length, 1, 2),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true, fontSize: 12 } } },
      fields: 'userEnteredFormat(horizontalAlignment,textFormat)',
    },
  });

  // Right column: Harvest & Budget summary (I-P, cols 8-15)
  reqs.push({ mergeCells: { range: gridRange(DASH, 6, 7, 8, 16), mergeType: 'MERGE_ALL' } });
  values.push({ range: 'Dashboard!I7', values: [['🌾 Season Harvest Summary']] });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 6, 7, 8, 16),
      cell: { userEnteredFormat: { backgroundColor: hex(C.paleWheat), textFormat: { bold: true, fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  const harvestSummary = [
    ['Total Harvests:',    `=IFERROR(COUNTA('Harvest & Yield Tracker'!A3:A14),0)`],
    ['Total Yield (lbs):', `=IFERROR(SUMIF('Harvest & Yield Tracker'!D3:D14,"lbs",'Harvest & Yield Tracker'!C3:C14),0)`],
    ['Highest Single Harvest:', `=IFERROR(MAX('Harvest & Yield Tracker'!C3:C14),0)`],
    ['Most Harvested Plant:', `=IFERROR(INDEX('Harvest & Yield Tracker'!A3:A14,MATCH(MAX(COUNTIF('Harvest & Yield Tracker'!A3:A14,'Harvest & Yield Tracker'!A3:A14)),'Harvest & Yield Tracker'!A3:A14,0)),"—")`],
    ['Budget Spent:',      `=IFERROR(SUM('Garden Budget'!F4:F13),"$0")`],
    ['Budget Remaining:',  `=IFERROR('Garden Budget'!H2,"$500")`],
    ['Active Pest Issues:', `=IFERROR(COUNTIF('Pest & Disease Log'!H3:H8,"Treating"),0)`],
  ];
  harvestSummary.forEach(([label, formula], i) => {
    values.push({ range: `Dashboard!I${8 + i}`, values: [[label, '', '', '', formula, '', '', '']]});
  });

  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 7, 7 + harvestSummary.length, 8, 16),
      cell: { userEnteredFormat: { backgroundColor: hex(C.cream2), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 7, 7 + harvestSummary.length, 8, 9),
      cell: { userEnteredFormat: { textFormat: { bold: true } } },
      fields: 'userEnteredFormat.textFormat.bold',
    },
  });
  // Merge label cells
  harvestSummary.forEach((_, i) => {
    reqs.push({ mergeCells: { range: gridRange(DASH, 7 + i, 8 + i, 8, 12), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DASH, 7 + i, 8 + i, 12, 16), mergeType: 'MERGE_ALL' } });
  });
  // Currency format for $$ cells
  [4, 5].forEach(i => {
    reqs.push({
      repeatCell: {
        range: gridRange(DASH, 7 + i, 8 + i, 12, 16),
        cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' }, textFormat: { bold: true, fontSize: 12 } } },
        fields: 'userEnteredFormat(numberFormat,textFormat)',
      },
    });
  });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 7, 7 + harvestSummary.length, 12, 16),
      cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 12 }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(textFormat,horizontalAlignment)',
    },
  });

  // Row heights for overview section
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 6, endIndex: 15 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // ── UPCOMING TASKS section (rows 15–21, indices 14–20) ────────────────────
  const uptRow = 14;
  reqs.push({ mergeCells: { range: gridRange(DASH, uptRow, uptRow + 1, 0, 16), mergeType: 'MERGE_ALL' } });
  values.push({ range: `Dashboard!A${uptRow + 1}`, values: [['💧 UPCOMING WATERING — Plants due in next 3 days']] });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, uptRow, uptRow + 1, 0, 16),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex('#2E6DA4'),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: uptRow, endIndex: uptRow + 1 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});

  values.push({ range: `Dashboard!A${uptRow + 2}`, values: [['Plant', '', '', 'Next Due', '', '', 'Days Until', '']]});
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, uptRow + 1, uptRow + 2, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex('#5B9BD5'), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  values.push({ range: `Dashboard!A${uptRow + 3}`, values: [
    [`=IFERROR(FILTER('Watering & Care Schedule'!A3:A14,'Watering & Care Schedule'!G3:G14<=3,'Watering & Care Schedule'!G3:G14>=0),"✅ All plants watered!")`]
  ]});
  values.push({ range: `Dashboard!D${uptRow + 3}`, values: [
    [`=IFERROR(FILTER('Watering & Care Schedule'!F3:F14,'Watering & Care Schedule'!G3:G14<=3,'Watering & Care Schedule'!G3:G14>=0),"")`]
  ]});
  values.push({ range: `Dashboard!G${uptRow + 3}`, values: [
    [`=IFERROR(FILTER('Watering & Care Schedule'!G3:G14,'Watering & Care Schedule'!G3:G14<=3,'Watering & Care Schedule'!G3:G14>=0),"")`]
  ]});

  reqs.push({
    repeatCell: {
      range: gridRange(DASH, uptRow + 2, uptRow + 8, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex('#EBF5FB'), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: uptRow + 1, endIndex: uptRow + 8 },
    properties: { pixelSize: 26 }, fields: 'pixelSize',
  }});

  // ── RECENT HARVESTS section ────────────────────────────────────────────────
  const harvRow = uptRow + 9;
  reqs.push({ mergeCells: { range: gridRange(DASH, harvRow, harvRow + 1, 0, 8), mergeType: 'MERGE_ALL' } });
  values.push({ range: `Dashboard!A${harvRow + 1}`, values: [['🌾 RECENT HARVESTS (Last 5)']] });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, harvRow, harvRow + 1, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.wheat),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  values.push({ range: `Dashboard!A${harvRow + 2}`, values: [['Plant', '', 'Date', '', 'Amount', '', 'Value', '']]});
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, harvRow + 1, harvRow + 2, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.harvestGold), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  for (let i = 0; i < 5; i++) {
    const refRow = 14 - i; // Last 5 entries (rows 10-14 in harvest tab = rows 12-16 with header)
    values.push({ range: `Dashboard!A${harvRow + 3 + i}`, values: [[
      `=IFERROR('Harvest & Yield Tracker'!A${refRow},"")`, '',
      `=IFERROR('Harvest & Yield Tracker'!B${refRow},"")`, '',
      `=IFERROR('Harvest & Yield Tracker'!C${refRow}&" "&'Harvest & Yield Tracker'!D${refRow},"")`, '',
      `=IFERROR('Harvest & Yield Tracker'!E${refRow},"")`, '',
    ]]});
  }

  reqs.push({
    repeatCell: {
      range: gridRange(DASH, harvRow + 2, harvRow + 8, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.cream2), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: harvRow, endIndex: harvRow + 8 },
    properties: { pixelSize: 26 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: harvRow, endIndex: harvRow + 1 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});

  // ── Chart helper table (rows 22–27, indices 21–26) ────────────────────────
  // Used for the 4 charts
  const chartHelperRow = harvRow + 9;
  reqs.push({ mergeCells: { range: gridRange(DASH, chartHelperRow, chartHelperRow + 1, 0, 8), mergeType: 'MERGE_ALL' } });
  values.push({ range: `Dashboard!A${chartHelperRow + 1}`, values: [['Chart Helper — Plant Status Distribution']] });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, chartHelperRow, chartHelperRow + 1, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lightGray), textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.warmGray) }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  values.push({ range: `Dashboard!A${chartHelperRow + 2}`, values: [['Status', 'Count']] });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, chartHelperRow + 1, chartHelperRow + 2, 0, 2),
      cell: { userEnteredFormat: { backgroundColor: hex(C.olive), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  const chartStatuses = ['Planned', 'Seeded', 'Germinated', 'Transplanted', 'Growing', 'Harvest Ready', 'Harvested'];
  const chartRows = chartStatuses.map(s => [s, `=COUNTIF('Plant Library'!L3:L37,"${s}")`]);
  values.push({ range: `Dashboard!A${chartHelperRow + 3}`, values: chartRows });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, chartHelperRow + 2, chartHelperRow + 2 + chartStatuses.length, 0, 2),
      cell: { userEnteredFormat: { backgroundColor: hex(C.cream2), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // ── Column widths ─────────────────────────────────────────────────────────
  for (let c = 0; c < 16; c++) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DASH, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: 100 }, fields: 'pixelSize',
    }});
  }

  // ── Freeze ────────────────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: DASH, gridProperties: { frozenRowCount: 2 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await valuesBatchUpdate(id, values, 'dashboard-values');
  await batchUpdate(id, reqs, 'dashboard-format');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
