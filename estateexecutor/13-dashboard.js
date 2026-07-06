'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DB = sheetMap['📋 Dashboard'];

(async () => {
  const reqs = [];

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: DB, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 70 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 3, endIndex: 55 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // Col widths: A=220 B=180 C=220 D=180 E=220 F=180 G=220 H=180 I=220 J=180
  [220, 180, 220, 180, 220, 180, 220, 180, 220, 180].forEach((px, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: DB, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: px }, fields: 'pixelSize',
  }}));

  // ── Row 0: Main title ─────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(DB, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 0, 1, 0, 10), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.deepCharcoal),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 22, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // ── Row 1: Disclaimer ─────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(DB, 1, 2, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 1, 2, 0, 10), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.disclaimerBg),
      textFormat: { foregroundColor: hex(C.warmGrey), italic: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
    },
  }, fields: 'userEnteredFormat' }});

  // ── Row 2: Estate details bar ──────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(DB, 2, 3, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 2, 3, 0, 10), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.mutedSage),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // ── Rows 3-4: Estate info panel (labels + values, 2 rows × 5 pairs) ────────
  // Row 3: label row (dark)  Row 4: value row (light)
  const infoPairs = [
    [0, 2], // A:B
    [2, 4], // C:D
    [4, 6], // E:F
    [6, 8], // G:H
    [8, 10], // I:J
  ];
  infoPairs.forEach(([c1, c2]) => {
    reqs.push({ mergeCells: { range: gridRange(DB, 3, 4, c1, c2), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DB, 4, 5, c1, c2), mergeType: 'MERGE_ALL' } });
  });
  reqs.push({ repeatCell: { range: gridRange(DB, 3, 4, 0, 10), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.deepCharcoal),
      textFormat: { foregroundColor: hex('#A8C4B8'), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  reqs.push({ repeatCell: { range: gridRange(DB, 4, 5, 0, 10), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.charcoal),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // ── Spacer row 5 ──────────────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 14 }, fields: 'pixelSize' }});

  // ── KPI blocks rows 6-11 (2 rows each, 5 KPIs across) ────────────────────
  // KPI 1: Tasks Completed     A:B
  // KPI 2: Milestones Done     C:D
  // KPI 3: Total Assets        E:F
  // KPI 4: Total Debts         G:H
  // KPI 5: Net Estate          I:J
  const kpiCols = [[0,2],[2,4],[4,6],[6,8],[8,10]];
  const kpiColors = [C.mutedSage, C.sageAccent, C.warmTaupe, C.deepRust, C.deepCharcoal];

  kpiCols.forEach(([c1, c2], i) => {
    // Label row 6
    reqs.push({ mergeCells: { range: gridRange(DB, 6, 7, c1, c2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(DB, 6, 7, c1, c2), cell: {
      userEnteredFormat: {
        backgroundColor: hex(kpiColors[i]),
        textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      },
    }, fields: 'userEnteredFormat' }});
    // Value row 7-9 (merge 3 rows for big number)
    reqs.push({ mergeCells: { range: gridRange(DB, 7, 10, c1, c2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(DB, 7, 10, c1, c2), cell: {
      userEnteredFormat: {
        backgroundColor: hex(C.warmStone),
        textFormat: { foregroundColor: hex(kpiColors[i]), bold: true, fontSize: 22 },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        borders: {
          bottom: { style: 'SOLID', color: hex(kpiColors[i]), width: 3 },
          left: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
          right: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
        },
      },
    }, fields: 'userEnteredFormat' }});
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 7, endIndex: 10 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // ── Spacer row 10 ─────────────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 14 }, fields: 'pixelSize' }});

  // ── Phase Progress table rows 11-18 (left A:E) ────────────────────────────
  // Section header row 11
  reqs.push({ mergeCells: { range: gridRange(DB, 11, 12, 0, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 11, 12, 0, 5), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.mutedSage),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      padding: { left: 8 },
    },
  }, fields: 'userEnteredFormat' }});

  // Col headers row 12
  reqs.push({ repeatCell: { range: gridRange(DB, 12, 13, 0, 5), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.sageAccent),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    },
  }, fields: 'userEnteredFormat' }});

  // Phase rows 13-18
  for (let r = 0; r < 6; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmStone;
    reqs.push({ repeatCell: { range: gridRange(DB, r+13, r+14, 0, 5), cell: {
      userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.charcoal), fontSize: 10 },
        verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
      },
    }, fields: 'userEnteredFormat' }});
  }

  // ── Estate Summary table rows 11-18 (right F:J) ───────────────────────────
  reqs.push({ mergeCells: { range: gridRange(DB, 11, 12, 5, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 11, 12, 5, 10), cell: {
    userEnteredFormat: {
      backgroundColor: hex(C.mutedSage),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11, fontFamily: 'Playfair Display' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      padding: { left: 8 },
    },
  }, fields: 'userEnteredFormat' }});

  for (let r = 0; r < 7; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmStone;
    reqs.push({ repeatCell: { range: gridRange(DB, r+12, r+13, 5, 10), cell: {
      userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.charcoal), fontSize: 11 },
        verticalAlignment: 'MIDDLE',
      },
    }, fields: 'userEnteredFormat' }});
  }

  // ── Borders ───────────────────────────────────────────────────────────────
  reqs.push({ updateBorders: { range: gridRange(DB, 11, 19, 0, 5),
    innerHorizontal: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    innerVertical: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    left: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    right: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    top: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
  }});
  reqs.push({ updateBorders: { range: gridRange(DB, 11, 19, 5, 10),
    innerHorizontal: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    innerVertical: { style: 'SOLID', color: hex(C.stoneBorder), width: 1 },
    left: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    right: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    top: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
  }});

  // Currency format on summary values F12:J19
  reqs.push({ repeatCell: { range: gridRange(DB, 12, 19, 6, 10), cell: {
    userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } },
  }, fields: 'userEnteredFormat.numberFormat' }});

  await batchUpdate(id, reqs, 'dash-fmt');

  // ── Values ────────────────────────────────────────────────────────────────
  const vals = [];

  vals.push({ range: `'📋 Dashboard'!A1`, values: [['ESTATE EXECUTOR ORGANISER']] });
  vals.push({ range: `'📋 Dashboard'!A2`, values: [['IMPORTANT DISCLAIMER: This tool is provided as an administrative aid only and does not constitute legal or financial advice. Estate administration involves complex legal and tax obligations. Consult a qualified solicitor and accountant before taking any action in relation to the estate.']] });
  vals.push({ range: `'📋 Dashboard'!A3`, values: [['Estate of Margaret Eleanor Williams | Executor: Robert James Williams | Ref: PROB-2025-0442 | Jurisdiction: New South Wales, Australia']] });

  // Estate info panel — row 4 labels
  vals.push({ range: `'📋 Dashboard'!A4:J4`, values: [[
    'DECEASED', 'DATE OF DEATH', 'ESTATE REFERENCE', 'EXECUTOR', 'JURISDICTION',
  ].flatMap(l => [l, ''])] });
  // row 5 values (B3 is the Date of Death input cell)
  vals.push({ range: `'📋 Dashboard'!A5`, values: [['Margaret Eleanor Williams']] });
  vals.push({ range: `'📋 Dashboard'!C5`, values: [['2025-04-12']] });
  vals.push({ range: `'📋 Dashboard'!E5`, values: [['PROB-2025-0442']] });
  vals.push({ range: `'📋 Dashboard'!G5`, values: [['Robert James Williams']] });
  vals.push({ range: `'📋 Dashboard'!I5`, values: [['New South Wales, AU']] });
  // Fix: B3 is the "Date of Death" cell referenced by Timeline. Let's put DoD at C5 (col 2, row 4).
  // The timeline formula references Dashboard!$B$3 — we need to align. Let's set B3 = DoD.
  vals.push({ range: `'📋 Dashboard'!B3`, values: [['2025-04-12']] });

  // Date format for C5
  reqs.push({ repeatCell: { range: gridRange(DB, 4, 5, 2, 3), cell: {
    userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } },
  }, fields: 'userEnteredFormat.numberFormat' }});
  // Date format B3
  reqs.push({ repeatCell: { range: gridRange(DB, 2, 3, 1, 2), cell: {
    userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } },
  }, fields: 'userEnteredFormat.numberFormat' }});

  await batchUpdate(id, reqs, 'dash-fmt2');

  // KPI values
  vals.push({ range: `'📋 Dashboard'!A7`, values: [['TASKS COMPLETED']] });
  vals.push({ range: `'📋 Dashboard'!C7`, values: [['MILESTONES DONE']] });
  vals.push({ range: `'📋 Dashboard'!E7`, values: [['TOTAL ASSETS']] });
  vals.push({ range: `'📋 Dashboard'!G7`, values: [['TOTAL DEBTS']] });
  vals.push({ range: `'📋 Dashboard'!I7`, values: [['NET ESTATE VALUE']] });

  // KPI big numbers
  vals.push({ range: `'📋 Dashboard'!A8`, values: [[`=COUNTIF('✅ Executor Checklist'!E4:E53,"Completed")&"/"&COUNTA('✅ Executor Checklist'!C4:C53)`]] });
  vals.push({ range: `'📋 Dashboard'!C8`, values: [[`=COUNTIF('📅 Estate Timeline & Probate Progress'!F4:F23,"Completed")&"/20"`]] });
  vals.push({ range: `'📋 Dashboard'!E8`, values: [[`=IFERROR(TEXT('💼 Assets & Debts'!F17,"$#,##0"),"—")`]] });
  vals.push({ range: `'📋 Dashboard'!G8`, values: [[`=IFERROR(TEXT('💼 Assets & Debts'!E26,"$#,##0"),"—")`]] });
  vals.push({ range: `'📋 Dashboard'!I8`, values: [[`=IFERROR(TEXT('💼 Assets & Debts'!F17-'💼 Assets & Debts'!E26,"$#,##0"),"—")`]] });

  // Phase progress table
  vals.push({ range: `'📋 Dashboard'!A12`, values: [['📊 EXECUTOR TASK PROGRESS BY PHASE']] });
  vals.push({ range: `'📋 Dashboard'!A13:E13`, values: [['Phase', 'Total Tasks', 'Completed', '% Complete', 'Progress Bar']] });

  const phases = [
    ['Phase 1: Immediate', '=COUNTIF(\'✅ Executor Checklist\'!B4:B53,"Phase 1: Immediate")', '=COUNTIFS(\'✅ Executor Checklist\'!B4:B53,"Phase 1: Immediate",\'✅ Executor Checklist\'!E4:E53,"Completed")'],
    ['Phase 2: First Month', '=COUNTIF(\'✅ Executor Checklist\'!B4:B53,"Phase 2: First Month")', '=COUNTIFS(\'✅ Executor Checklist\'!B4:B53,"Phase 2: First Month",\'✅ Executor Checklist\'!E4:E53,"Completed")'],
    ['Phase 3: 1-3 Months', '=COUNTIF(\'✅ Executor Checklist\'!B4:B53,"Phase 3: 1-3 Months")', '=COUNTIFS(\'✅ Executor Checklist\'!B4:B53,"Phase 3: 1-3 Months",\'✅ Executor Checklist\'!E4:E53,"Completed")'],
    ['Phase 4: 3-6 Months', '=COUNTIF(\'✅ Executor Checklist\'!B4:B53,"Phase 4: 3-6 Months")', '=COUNTIFS(\'✅ Executor Checklist\'!B4:B53,"Phase 4: 3-6 Months",\'✅ Executor Checklist\'!E4:E53,"Completed")'],
    ['Phase 5: 6-12 Months', '=COUNTIF(\'✅ Executor Checklist\'!B4:B53,"Phase 5: 6-12 Months")', '=COUNTIFS(\'✅ Executor Checklist\'!B4:B53,"Phase 5: 6-12 Months",\'✅ Executor Checklist\'!E4:E53,"Completed")'],
    ['Phase 6: Finalisation', '=COUNTIF(\'✅ Executor Checklist\'!B4:B53,"Phase 6: Finalisation")', '=COUNTIFS(\'✅ Executor Checklist\'!B4:B53,"Phase 6: Finalisation",\'✅ Executor Checklist\'!E4:E53,"Completed")'],
  ];

  phases.forEach(([phase, total, done], i) => {
    const r = i + 14; // row index 14-19
    vals.push({ range: `'📋 Dashboard'!A${r}`, values: [[phase]] });
    vals.push({ range: `'📋 Dashboard'!B${r}`, values: [[total]] });
    vals.push({ range: `'📋 Dashboard'!C${r}`, values: [[done]] });
    vals.push({ range: `'📋 Dashboard'!D${r}`, values: [[`=IFERROR(C${r}/B${r},"—")`]] });
    vals.push({ range: `'📋 Dashboard'!E${r}`, values: [[`=IFERROR(SPARKLINE({C${r},B${r}-C${r}},{"charttype","bar";"color1","#3D5A4A";"color2","#D5CFC4"}),"—")`]] });
  });

  // Estate Summary table (right side F:J)
  vals.push({ range: `'📋 Dashboard'!F12`, values: [['📋 ESTATE FINANCIAL SUMMARY']] });
  const summaryRows = [
    ['Total Gross Assets (Date of Death)', '', '=IFERROR(\'💼 Assets & Debts\'!E17,0)'],
    ['Total Current Asset Values', '', '=IFERROR(\'💼 Assets & Debts\'!F17,0)'],
    ['Total Debts & Liabilities', '', '=IFERROR(\'💼 Assets & Debts\'!E26,0)'],
    ['Net Estate Value (Current)', '', '=IFERROR(\'💼 Assets & Debts\'!F17-\'💼 Assets & Debts\'!E26,0)'],
    ['Total Professional Fees Paid', '', '=IFERROR(\'🏛️ Professional Services & Fees\'!K9,0)'],
    ['Total Fees Outstanding', '', '=IFERROR(\'🏛️ Professional Services & Fees\'!L9,0)'],
    ['Total Distributed to Beneficiaries', '', '=IFERROR(\'🏦 Estate Distribution\'!G8,0)'],
  ];
  summaryRows.forEach(([label, blank, formula], i) => {
    const r = i + 13;
    vals.push({ range: `'📋 Dashboard'!F${r}`, values: [[label]] });
    vals.push({ range: `'📋 Dashboard'!H${r}`, values: [[formula]] });
  });

  // Percentage format for D col
  reqs.push({ repeatCell: { range: gridRange(DB, 13, 20, 3, 4), cell: {
    userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } },
  }, fields: 'userEnteredFormat.numberFormat' }});

  // Label col F right-align
  reqs.push({ repeatCell: { range: gridRange(DB, 12, 20, 5, 9), cell: {
    userEnteredFormat: { horizontalAlignment: 'RIGHT' },
  }, fields: 'userEnteredFormat.horizontalAlignment' }});

  await batchUpdate(id, reqs, 'dash-fmt3');
  await valuesBatchUpdate(id, vals, 'dash-vals');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
