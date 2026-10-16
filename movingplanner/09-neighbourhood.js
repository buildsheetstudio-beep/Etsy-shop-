'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const NB = sheetMap['🏘️ Neighbourhood Comparison'];

(async () => {
  const values = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  values.push({ range: "'🏘️ Neighbourhood Comparison'!A1", values: [['🏘️ All-in-One Moving Planner — Neighbourhood Comparison']] });

  // ── Instructions row ──────────────────────────────────────────────────────
  values.push({ range: "'🏘️ Neighbourhood Comparison'!A2", values: [[
    'Rate each neighbourhood 1–5 per criterion. Set Weight (1–3) for importance. Weighted scores are auto-calculated.',
  ]]});

  // ── Column headers (row 3) ────────────────────────────────────────────────
  values.push({ range: "'🏘️ Neighbourhood Comparison'!A3", values: [[
    'Criterion', 'Weight (1–3)', 'Riverside (Target)', 'Oak Park', 'Thornton', 'Brookfield',
    'Best Score', 'WINNER',
  ]]});

  // ── 12 criteria rows ─────────────────────────────────────────────────────
  const criteria = [
    ['School Quality',       3, 4, 4, 3, 5],
    ['Commute Time',         3, 5, 3, 4, 4],
    ['Safety / Crime Rate',  3, 4, 5, 3, 4],
    ['Proximity to Shops',   2, 5, 4, 4, 3],
    ['Parks & Green Space',  2, 4, 3, 5, 4],
    ['Medical Facilities',   2, 4, 4, 3, 4],
    ['Property Value Trend', 2, 4, 3, 4, 5],
    ['Public Transport',     2, 5, 3, 3, 4],
    ['Childcare Options',    2, 4, 4, 3, 4],
    ['Community Feel',       1, 4, 4, 5, 4],
    ['Restaurant / Cafes',   1, 5, 4, 3, 3],
    ['Noise / Traffic',      2, 4, 3, 4, 4],
  ];

  const dataRows = criteria.map((row, i) => {
    const r = 4 + i; // row index 4 = row 5 in sheets
    return [
      row[0], row[1], row[2], row[3], row[4], row[5],
      `=MAX(C${r + 1}*$B${r + 1},D${r + 1}*$B${r + 1},E${r + 1}*$B${r + 1},F${r + 1}*$B${r + 1})`,
      `=IFERROR(INDEX($C$3:$F$3,MATCH(G${r + 1},C${r + 1}*$B${r + 1}&D${r + 1}*$B${r + 1}&E${r + 1}*$B${r + 1}&F${r + 1}*$B${r + 1},0)),"—")`,
    ];
  });

  // Fix — use simpler winner formula: INDEX of column headers based on max weighted score
  const dataRowsFixed = criteria.map((row, i) => {
    const r = 5 + i; // 1-based row number in sheet
    const wRow = row[1]; // weight value
    return [
      row[0],
      row[1],
      row[2], row[3], row[4], row[5],
      `=MAX(C${r}*B${r},D${r}*B${r},E${r}*B${r},F${r}*B${r})`,
      `=IFERROR(INDEX({"Riverside";"Oak Park";"Thornton";"Brookfield"},MATCH(MAX(C${r}*B${r},D${r}*B${r},E${r}*B${r},F${r}*B${r}),{C${r}*B${r};D${r}*B${r};E${r}*B${r};F${r}*B${r}},0)),"—")`,
    ];
  });

  values.push({ range: "'🏘️ Neighbourhood Comparison'!A4", values: dataRowsFixed });

  // ── Weighted total row ─────────────────────────────────────────────────────
  values.push({ range: "'🏘️ Neighbourhood Comparison'!A16", values: [[
    'WEIGHTED TOTAL', '=SUM(B4:B15)',
    '=SUMPRODUCT(C4:C15,B4:B15)',
    '=SUMPRODUCT(D4:D15,B4:B15)',
    '=SUMPRODUCT(E4:E15,B4:B15)',
    '=SUMPRODUCT(F4:F15,B4:B15)',
    '', '',
  ]]});

  // ── Overall winner row ────────────────────────────────────────────────────
  values.push({ range: "'🏘️ Neighbourhood Comparison'!A17", values: [[
    'OVERALL WINNER', '',
    `=IFERROR(INDEX({"Riverside";"Oak Park";"Thornton";"Brookfield"},MATCH(MAX(C16,D16,E16,F16),{C16;D16;E16;F16},0)),"—")`,
    '', '', '', '', '',
  ]]});

  // ── SPARKLINE score bars (row 18 onwards) ─────────────────────────────────
  values.push({ range: "'🏘️ Neighbourhood Comparison'!A19", values: [['Score Bars (vs. max possible)']] });
  const nbNames = ['Riverside', 'Oak Park', 'Thornton', 'Brookfield'];
  const nbCols  = ['C', 'D', 'E', 'F'];
  const sparkRows = nbNames.map((name, i) => [
    name,
    `=SUMPRODUCT(${nbCols[i]}4:${nbCols[i]}15,$B$4:$B$15)`,
    `=IFERROR(SPARKLINE({${nbCols[i]}16,MAX($C$16:$F$16)-${nbCols[i]}16},{"charttype","bar";"color1","#2C4A35";"color2","#FBF5E6"}),"-")`,
  ]);
  values.push({ range: "'🏘️ Neighbourhood Comparison'!A20", values: sparkRows });

  await valuesBatchUpdate(id, values, 'neighbourhood-values');

  // ── Merge title ───────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(NB, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(NB, 0, 1, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.slateBlue),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Instructions row 2 ───────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(NB, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(NB, 1, 2, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.linen),
          textFormat: { foregroundColor: hex(C.nearBlack), fontSize: 10, italic: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Column headers row 3 ──────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(NB, 2, 3, 0, 8),
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

  // ── Zebra stripe criteria rows 4–15 ──────────────────────────────────────
  for (let r = 0; r < 12; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(NB, 3 + r, 4 + r, 0, 8),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(C.nearBlack), fontSize: 10 },
            verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
      },
    });
    // Bold criterion name in col A
    reqs.push({
      repeatCell: {
        range: gridRange(NB, 3 + r, 4 + r, 0, 1),
        cell: {
          userEnteredFormat: {
            textFormat: { bold: true, fontSize: 10 },
            horizontalAlignment: 'LEFT',
          },
        },
        fields: 'userEnteredFormat(textFormat,horizontalAlignment)',
      },
    });
  }

  // ── Weighted Total row ────────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(NB, 15, 16, 0, 8),
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

  // ── Overall Winner row ────────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(NB, 16, 17, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepForest),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Score bars section ────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(NB, 18, 19, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(NB, 18, 19, 0, 8),
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
  for (let r = 0; r < 4; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(NB, 19 + r, 20 + r, 0, 3),
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
  const colWidths = [180, 80, 100, 100, 100, 100, 90, 120];
  colWidths.forEach((w, i) => {
    reqs.push({
      updateDimensionProperties: {
        range: { sheetId: NB, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
        properties: { pixelSize: w },
        fields: 'pixelSize',
      },
    });
  });

  // ── Row heights ───────────────────────────────────────────────────────────
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: NB, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
      properties: { pixelSize: 40 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: NB, dimension: 'ROWS', startIndex: 1, endIndex: 3 },
      properties: { pixelSize: 28 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: NB, dimension: 'ROWS', startIndex: 3, endIndex: 20 },
      properties: { pixelSize: 24 },
      fields: 'pixelSize',
    },
  });

  // ── Freeze 3 rows (no col freeze — full-width title merge conflicts) ─────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: NB, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
      fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
    },
  });

  await batchUpdate(id, reqs, 'neighbourhood-format');
  console.log('Neighbourhood Comparison complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
