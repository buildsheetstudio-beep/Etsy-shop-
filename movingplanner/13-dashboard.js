'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH = sheetMap['🏠 Dashboard'];

(async () => {
  const values = [];
  const reqs = [];

  // ── Title row ─────────────────────────────────────────────────────────────
  values.push({ range: "'🏠 Dashboard'!A1", values: [['🏠 All-in-One Moving Planner']] });
  values.push({ range: "'🏠 Dashboard'!I1", values: [['BuildSheetStudio']] });

  // ── Move Details input block (rows 3–8) ───────────────────────────────────
  values.push({ range: "'🏠 Dashboard'!A3", values: [['MOVE DETAILS']] });
  values.push({ range: "'🏠 Dashboard'!A4", values: [
    ['Moving From:'], ['Moving To:'], ['Move Date:'], ['Move Type:'], ['Moving Company:'],
  ]});
  values.push({ range: "'🏠 Dashboard'!B4", values: [
    ['42 Birchwood Lane, Springfield, IL 62704'],
    ['15 Oakmont Drive, Riverside, IL 60546'],
    ['2025-08-30'],
    ['Local'],
    ['=IFERROR(\'🚛 Moving Company Quotes\'!B6,"TBD")'],
  ]});

  // ── Countdown (row 10) ────────────────────────────────────────────────────
  values.push({ range: "'🏠 Dashboard'!A10", values: [
    ['Days Until Move:'],
  ]});
  values.push({ range: "'🏠 Dashboard'!B10", values: [
    ['=MAX(DATEVALUE(TEXT(B6,"YYYY-MM-DD"))-TODAY(),0)'],
  ]});
  values.push({ range: "'🏠 Dashboard'!C10", values: [['days']] });

  // ── KPI block header (row 12) ─────────────────────────────────────────────
  values.push({ range: "'🏠 Dashboard'!A12", values: [['MOVING SNAPSHOT']] });

  // ── KPI labels row 13 ────────────────────────────────────────────────────
  values.push({ range: "'🏠 Dashboard'!A13", values: [[
    'Tasks Done', 'Tasks Remaining', 'Boxes Packed', 'Budget Spent',
    'Budget Remaining', 'Address Changes Done', 'Utilities Connected', '% Tasks Complete',
  ]]});

  // ── KPI values row 14 ────────────────────────────────────────────────────
  values.push({ range: "'🏠 Dashboard'!A14", values: [[
    '=IFERROR(COUNTIF(\'✅ Task Tracker\'!E6:E65,"Done"),0)',
    '=IFERROR(COUNTIF(\'✅ Task Tracker\'!E6:E65,"Not Started")+COUNTIF(\'✅ Task Tracker\'!E6:E65,"In Progress")+COUNTIF(\'✅ Task Tracker\'!E6:E65,"Blocked"),0)',
    '=IFERROR(COUNTIF(\'📦 Packing Tracker\'!E3:E22,"Packed")+COUNTIF(\'📦 Packing Tracker\'!E3:E22,"Loaded")+COUNTIF(\'📦 Packing Tracker\'!E3:E22,"Arrived")+COUNTIF(\'📦 Packing Tracker\'!E3:E22,"Unpacked"),0)',
    '=IFERROR(\'💰 Moving Budget\'!B3,0)',
    '=IFERROR(\'💰 Moving Budget\'!C3,0)',
    '=IFERROR(COUNTA(\'📋 Change of Address\'!G6:G65)-COUNTBLANK(\'📋 Change of Address\'!G6:G65),0)',
    '=IFERROR(COUNTIF(\'⚡ Utilities & Services\'!K6:K25,"Connected")+COUNTIF(\'⚡ Utilities & Services\'!K6:K25,"Active"),0)',
    '=IFERROR(\'✅ Task Tracker\'!F3,0)',
  ]]});

  // ── Task Status mini table (rows 16–21) ───────────────────────────────────
  values.push({ range: "'🏠 Dashboard'!A16", values: [['TASK STATUS BREAKDOWN']] });
  values.push({ range: "'🏠 Dashboard'!A17", values: [
    ['Status'], ['Not Started'], ['In Progress'], ['Done'], ['Blocked'],
  ]});
  values.push({ range: "'🏠 Dashboard'!B17", values: [
    ['Count'],
    ['=IFERROR(COUNTIF(\'✅ Task Tracker\'!E6:E65,"Not Started"),0)'],
    ['=IFERROR(COUNTIF(\'✅ Task Tracker\'!E6:E65,"In Progress"),0)'],
    ['=IFERROR(COUNTIF(\'✅ Task Tracker\'!E6:E65,"Done"),0)'],
    ['=IFERROR(COUNTIF(\'✅ Task Tracker\'!E6:E65,"Blocked"),0)'],
  ]});

  // ── Packing progress table (rows 16–21, col D–F) ──────────────────────────
  values.push({ range: "'🏠 Dashboard'!D16", values: [['PACKING PROGRESS']] });
  values.push({ range: "'🏠 Dashboard'!D17", values: [
    ['Status'], ['To Pack'], ['Packed'], ['Loaded'], ['Arrived/Unpacked'],
  ]});
  values.push({ range: "'🏠 Dashboard'!E17", values: [
    ['Boxes'],
    ['=IFERROR(COUNTIF(\'📦 Packing Tracker\'!E3:E22,"To Pack"),0)'],
    ['=IFERROR(COUNTIF(\'📦 Packing Tracker\'!E3:E22,"Packed"),0)'],
    ['=IFERROR(COUNTIF(\'📦 Packing Tracker\'!E3:E22,"Loaded"),0)'],
    ['=IFERROR(COUNTIF(\'📦 Packing Tracker\'!E3:E22,"Arrived")+COUNTIF(\'📦 Packing Tracker\'!E3:E22,"Unpacked"),0)'],
  ]});

  // ── Upcoming tasks (rows 23–30) ────────────────────────────────────────────
  values.push({ range: "'🏠 Dashboard'!A23", values: [['UPCOMING TASKS (Next 14 Days)']] });
  values.push({ range: "'🏠 Dashboard'!A24", values: [['Task', 'Due Date', 'Priority', 'Status']] });
  // FILTER formula — tasks due within 14 days
  values.push({ range: "'🏠 Dashboard'!A25", values: [[
    '=IFERROR(FILTER(\'✅ Task Tracker\'!C6:C65,\'✅ Task Tracker\'!F6:F65>=TODAY(),\'✅ Task Tracker\'!F6:F65<=TODAY()+14,\'✅ Task Tracker\'!E6:E65<>"Done"),"No upcoming tasks in next 14 days")',
  ]]});
  values.push({ range: "'🏠 Dashboard'!B25", values: [[
    '=IFERROR(FILTER(\'✅ Task Tracker\'!F6:F65,\'✅ Task Tracker\'!F6:F65>=TODAY(),\'✅ Task Tracker\'!F6:F65<=TODAY()+14,\'✅ Task Tracker\'!E6:E65<>"Done"),"")',
  ]]});
  values.push({ range: "'🏠 Dashboard'!C25", values: [[
    '=IFERROR(FILTER(\'✅ Task Tracker\'!D6:D65,\'✅ Task Tracker\'!F6:F65>=TODAY(),\'✅ Task Tracker\'!F6:F65<=TODAY()+14,\'✅ Task Tracker\'!E6:E65<>"Done"),"")',
  ]]});
  values.push({ range: "'🏠 Dashboard'!D25", values: [[
    '=IFERROR(FILTER(\'✅ Task Tracker\'!E6:E65,\'✅ Task Tracker\'!F6:F65>=TODAY(),\'✅ Task Tracker\'!F6:F65<=TODAY()+14,\'✅ Task Tracker\'!E6:E65<>"Done"),"")',
  ]]});

  // ── Budget summary block (rows 37–42, col G–I) ────────────────────────────
  values.push({ range: "'🏠 Dashboard'!G16", values: [['BUDGET SNAPSHOT']] });
  values.push({ range: "'🏠 Dashboard'!G17", values: [
    ['Total Budget'], ['Total Spent'], ['Remaining'], ['Over Budget?'], ['Biggest Expense'],
  ]});
  values.push({ range: "'🏠 Dashboard'!H17", values: [
    ['=IFERROR(\'💰 Moving Budget\'!A3,0)'],
    ['=IFERROR(\'💰 Moving Budget\'!B3,0)'],
    ['=IFERROR(\'💰 Moving Budget\'!C3,0)'],
    ['=IFERROR(\'💰 Moving Budget\'!D3,"—")'],
    ['=IFERROR(INDEX(\'💰 Moving Budget\'!B6:B15,MATCH(MAX(\'💰 Moving Budget\'!E6:E15),\'💰 Moving Budget\'!E6:E15,0)),"—")'],
  ]});

  await valuesBatchUpdate(id, values, 'dashboard-values');

  // ── Merge title A1:H1 ────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(DASH, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 0, 1, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepForest),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  // Brand name col I1 (index 8)
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 0, 1, 8, 9),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepForest),
          textFormat: { foregroundColor: hex(C.antiqueBrass), bold: true, fontSize: 10, italic: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── MOVE DETAILS section header ───────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(DASH, 2, 3, 0, 4), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 2, 3, 0, 4),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.antiqueBrass),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Move detail label cells (A4:A8) ───────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 3, 8, 0, 1),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.linen),
          textFormat: { foregroundColor: hex(C.nearBlack), bold: true, fontSize: 10 },
          horizontalAlignment: 'RIGHT',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Move detail value cells (B4:B8) ───────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 3, 8, 1, 4),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.ivory),
          textFormat: { foregroundColor: hex(C.nearBlack), fontSize: 10 },
          horizontalAlignment: 'LEFT',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // Date format for B6 (Move Date)
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 5, 6, 1, 2),
      cell: {
        userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } },
      },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // ── Countdown ─────────────────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 9, 10, 0, 1),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.linen),
          textFormat: { foregroundColor: hex(C.nearBlack), bold: true, fontSize: 11 },
          horizontalAlignment: 'RIGHT',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 9, 10, 1, 2),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.antiqueBrass),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 20 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── SNAPSHOT section header ────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(DASH, 11, 12, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 11, 12, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepForest),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── KPI label row 13 ─────────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 12, 13, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.mossGreen),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── KPI value row 14 ─────────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 13, 14, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.ivory),
          textFormat: { foregroundColor: hex(C.nearBlack), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  // Currency format for Budget Spent (D14) and Budget Remaining (E14) — indices 3,4
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 13, 14, 3, 5),
      cell: {
        userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } },
      },
      fields: 'userEnteredFormat.numberFormat',
    },
  });
  // Percent for col H14 (index 7)
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 13, 14, 7, 8),
      cell: {
        userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } },
      },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // ── Task Status table header row 16 ───────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(DASH, 15, 16, 0, 2), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 15, 16, 0, 2),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.antiqueBrass),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // Task status data rows 17–21
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 16, 17, 0, 2),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepForest),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  for (let r = 0; r < 4; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(DASH, 17 + r, 18 + r, 0, 2),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(C.nearBlack), fontSize: 10 },
            horizontalAlignment: 'CENTER',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
      },
    });
  }

  // ── Packing Progress table header row 16 col D ────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(DASH, 15, 16, 3, 6), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 15, 16, 3, 6),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.mossGreen),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 16, 17, 3, 5),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepForest),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  for (let r = 0; r < 4; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(DASH, 17 + r, 18 + r, 3, 5),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(C.nearBlack), fontSize: 10 },
            horizontalAlignment: 'CENTER',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
      },
    });
  }

  // ── Budget Snapshot table (col G–H) ──────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(DASH, 15, 16, 6, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 15, 16, 6, 9),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.forestGreen),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  for (let r = 0; r < 5; r++) {
    const bg = r % 2 === 0 ? C.linen : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(DASH, 16 + r, 17 + r, 6, 7),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(C.nearBlack), bold: true, fontSize: 10 },
            horizontalAlignment: 'RIGHT',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
      },
    });
    reqs.push({
      repeatCell: {
        range: gridRange(DASH, 16 + r, 17 + r, 7, 9),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(C.nearBlack), fontSize: 10 },
            horizontalAlignment: 'CENTER',
            numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' },
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,numberFormat)',
      },
    });
  }
  // Over Budget? and Biggest Expense don't need currency
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 19, 21, 7, 9),
      cell: {
        userEnteredFormat: {
          numberFormat: { type: 'TEXT' },
        },
      },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // ── UPCOMING TASKS section header ─────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(DASH, 22, 23, 0, 4), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 22, 23, 0, 4),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepForest),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  // Column headers row 24
  reqs.push({
    repeatCell: {
      range: gridRange(DASH, 23, 24, 0, 4),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.mossGreen),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  // Upcoming task rows
  for (let r = 0; r < 10; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(DASH, 24 + r, 25 + r, 0, 4),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(C.nearBlack), fontSize: 10 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [200, 180, 100, 160, 10, 100, 10, 130, 130];
  colWidths.forEach((w, i) => {
    reqs.push({
      updateDimensionProperties: {
        range: { sheetId: DASH, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
        properties: { pixelSize: w },
        fields: 'pixelSize',
      },
    });
  });

  // ── Row heights ───────────────────────────────────────────────────────────
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: DASH, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
      properties: { pixelSize: 50 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: DASH, dimension: 'ROWS', startIndex: 1, endIndex: 15 },
      properties: { pixelSize: 28 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: DASH, dimension: 'ROWS', startIndex: 15, endIndex: 40 },
      properties: { pixelSize: 24 },
      fields: 'pixelSize',
    },
  });

  // ── Freeze rows 1–2 ───────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: DASH, gridProperties: { frozenRowCount: 2 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await batchUpdate(id, reqs, 'dashboard-format');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
