'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard'];

// Layout (0-based idx):
//   idx 0: Title merged A1:H1
//   idx 1: Subtitle merged A2:H2
//   idx 2: Inputs — A3="Total Budget ($):", B3=1200, C3="Holiday Date:", D3=2025-12-25
//   idx 3: blank
//   idx 4-5: KPI label row 1 (4 stats: Total Spent, Left to Spend, Gifts Bought %, Gifts Wrapped %)
//   idx 6-7: KPI value row 1
//   idx 8-9: KPI label row 2 (3 stats: Gifts Sent %, Cards Sent %, Days Until Christmas)
//   idx 10-11: KPI value row 2
//   idx 12: blank spacer
//   idx 13-19: chart area row A (2 donuts)
//   idx 20-26: chart area row B (1 donut + budget chart)
//   idx 27: blank
//   idx 28: "— Chart Data (Internal) —"
//   idx 29: blank
//   idx 30-31: Bought donut data (label | count)
//   idx 32-33: Wrapped donut data
//   idx 34-35: Sent donut data

const GRT_RANGE = "'🎁 Gift Recipient Tracker'";
const CDL_RANGE = "'💌 Holiday Card & Mailing List'";

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'📊 Dashboard'!A1", values: [['🎄 ULTIMATE CHRISTMAS GIFT TRACKER 🎁']] },
    { range: "'📊 Dashboard'!A2", values: [['Track Every Gift, Card & Return — Reusable Every Year']] },
    { range: "'📊 Dashboard'!A3", values: [['Total Budget ($):', 1200, 'Holiday Date:', '2025-12-25']] },
    // KPI labels row 1 (row 5 = idx 4)
    { range: "'📊 Dashboard'!B5", values: [['❄️ Total Spent ($)', 'Left to Spend ($)', 'Gifts Bought %', 'Gifts Wrapped %']] },
    // KPI values row 1 (row 7 = idx 6)
    { range: "'📊 Dashboard'!B7", values: [[
      `=IFERROR(SUM(${GRT_RANGE}!F2:F26),0)`,
      `=IFERROR(B3-SUM(${GRT_RANGE}!F2:F26),0)`,
      `=IFERROR(COUNTIF(${GRT_RANGE}!H2:H26,TRUE)/COUNTA(${GRT_RANGE}!A2:A26),"")`,
      `=IFERROR(COUNTIF(${GRT_RANGE}!I2:I26,TRUE)/COUNTA(${GRT_RANGE}!A2:A26),"")`,
    ]] },
    // KPI labels row 2 (row 9 = idx 8)
    { range: "'📊 Dashboard'!B9", values: [['🎁 Gifts Sent %', '✉️ Cards Sent %', '🎄 Days Until Christmas']] },
    // KPI values row 2 (row 11 = idx 10)
    { range: "'📊 Dashboard'!B11", values: [[
      `=IFERROR(COUNTIF(${GRT_RANGE}!J2:J26,TRUE)/COUNTA(${GRT_RANGE}!A2:A26),"")`,
      `=IFERROR(COUNTIF(${CDL_RANGE}!C2:C21,"Sent")/COUNTA(${CDL_RANGE}!A2:A21),"")`,
      `=IF(DATE(YEAR(TODAY()),12,25)>=TODAY(),DATE(YEAR(TODAY()),12,25)-TODAY(),DATE(YEAR(TODAY())+1,12,25)-TODAY())`,
    ]] },
    // Chart data section header (row 29 = idx 28)
    { range: "'📊 Dashboard'!A29", values: [['— Chart Data (Internal) —']] },
    // Bought donut data (rows 31-32 = idx 30-31)
    { range: "'📊 Dashboard'!A31", values: [
      ['Bought',     `=IFERROR(COUNTIF(${GRT_RANGE}!H2:H26,TRUE),0)`],
      ['Not Bought', `=IFERROR(COUNTA(${GRT_RANGE}!A2:A26)-COUNTIF(${GRT_RANGE}!H2:H26,TRUE),0)`],
    ] },
    // Wrapped donut data (rows 33-34 = idx 32-33)
    { range: "'📊 Dashboard'!A33", values: [
      ['Wrapped',     `=IFERROR(COUNTIF(${GRT_RANGE}!I2:I26,TRUE),0)`],
      ['Not Wrapped', `=IFERROR(COUNTA(${GRT_RANGE}!A2:A26)-COUNTIF(${GRT_RANGE}!I2:I26,TRUE),0)`],
    ] },
    // Sent donut data (rows 35-36 = idx 34-35)
    { range: "'📊 Dashboard'!A35", values: [
      ['Sent',     `=IFERROR(COUNTIF(${GRT_RANGE}!J2:J26,TRUE),0)`],
      ['Not Sent', `=IFERROR(COUNTA(${GRT_RANGE}!A2:A26)-COUNTIF(${GRT_RANGE}!J2:J26,TRUE),0)`],
    ] },
  ], 'dashboard-values');

  const reqs = [];

  // ── Merges ────────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(DSH, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } }); // Title
  reqs.push({ mergeCells: { range: gridRange(DSH, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } }); // Subtitle
  // KPI row 1 merges (idx 4 labels, idx 6 values): B:C, D:E, F:G, H
  for (const ri of [4, 6]) {
    reqs.push({ mergeCells: { range: gridRange(DSH, ri, ri + 1, 1, 3), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DSH, ri, ri + 1, 3, 5), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DSH, ri, ri + 1, 5, 7), mergeType: 'MERGE_ALL' } });
  }
  // KPI row 2 merges (idx 8 labels, idx 10 values): B:C, D:E, F:H (wider for Days)
  for (const ri of [8, 10]) {
    reqs.push({ mergeCells: { range: gridRange(DSH, ri, ri + 1, 1, 3), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DSH, ri, ri + 1, 3, 5), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DSH, ri, ri + 1, 5, 8), mergeType: 'MERGE_ALL' } });
  }

  // ── Title ─────────────────────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 0, 1, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepCranberry),
          textFormat: { foregroundColor: hex(C.white), fontSize: 18, bold: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Subtitle ──────────────────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 1, 2, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.trueRed),
          textFormat: { foregroundColor: hex(C.lightAmber), fontSize: 11, italic: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Input row ─────────────────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 2, 3, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.lightAmber),
          textFormat: { foregroundColor: hex(C.darkText), fontSize: 11 },
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    },
  });
  // B3 — budget value
  reqs.push({ repeatCell: { range: gridRange(DSH, 2, 3, 1, 2), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 12 }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment,numberFormat)' } });
  // D3 — date value
  reqs.push({ repeatCell: { range: gridRange(DSH, 2, 3, 3, 4), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 11 }, horizontalAlignment: 'CENTER', numberFormat: { type: 'DATE', pattern: 'MMM d, yyyy' } } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment,numberFormat)' } });

  // ── KPI label rows (idx 4 & 8) ────────────────────────────────────────────
  for (const ri of [4, 8]) {
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, ri, ri + 1, 1, 8),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.warmGold),
            textFormat: { foregroundColor: hex(C.darkText), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
            wrapStrategy: 'WRAP',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
      },
    });
  }

  // ── KPI value rows (idx 6 & 10) ───────────────────────────────────────────
  for (const ri of [6, 10]) {
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, ri, ri + 1, 1, 8),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.lightAmber),
            textFormat: { foregroundColor: hex(C.deepCranberry), bold: true, fontSize: 26 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
  }

  // Currency format for B7:C7 (Total Spent, Left to Spend)
  reqs.push({ repeatCell: { range: gridRange(DSH, 6, 7, 1, 5), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat(numberFormat)' } });
  // Percentage format for D7:H7 (Bought%, Wrapped%)
  reqs.push({ repeatCell: { range: gridRange(DSH, 6, 7, 5, 8), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0%' } } }, fields: 'userEnteredFormat(numberFormat)' } });
  // Percentage format for B11:E11 (Sent%, Cards%)
  reqs.push({ repeatCell: { range: gridRange(DSH, 10, 11, 1, 5), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0%' } } }, fields: 'userEnteredFormat(numberFormat)' } });
  // Days value — plain number, no decimal
  reqs.push({ repeatCell: { range: gridRange(DSH, 10, 11, 5, 8), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0' } } }, fields: 'userEnteredFormat(numberFormat)' } });

  // ── Chart data area — dimmed text ─────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 28, 36, 0, 2),
      cell: { userEnteredFormat: { textFormat: { foregroundColor: hex(C.mediumGray), fontSize: 9 } } },
      fields: 'userEnteredFormat(textFormat)',
    },
  });

  // ── Row heights ───────────────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 34 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 8, endIndex: 9 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });

  // ── Column widths ─────────────────────────────────────────────────────────
  [[0,160],[1,140],[2,140],[3,140],[4,140],[5,140],[6,140],[7,140]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'dashboard-format');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
