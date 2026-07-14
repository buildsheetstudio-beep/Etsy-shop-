'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH  = sheetMap['📊 Dashboard'];
const DSH2 = sheetMap['🌙 Dashboard — Dark Theme'];
const GRT  = sheetMap['🎁 Gift Recipient Tracker'];

// Dark palette — deep pine theme
const D = {
  pine:       '#1F2B22',
  deepPine:   '#162019',
  midPine:    '#2A3C2D',
  cream:      '#F7F8F7',
  gold:       '#C9A227',
  cranberry:  '#8C1F2B',
  lightGold:  '#FFF3CC',
  amber:      '#D9A23C',
  darkGold:   '#A07C16',
  mutedCream: '#C8C9C8',
};

function hexD(h) {
  const r = parseInt(h.slice(1, 3), 16) / 255;
  const g = parseInt(h.slice(3, 5), 16) / 255;
  const b = parseInt(h.slice(5, 7), 16) / 255;
  return { red: r, green: g, blue: b };
}

const GRT_RANGE = "'🎁 Gift Recipient Tracker'";
const CDL_RANGE = "'💌 Holiday Card & Mailing List'";
// Reference light dashboard chart data tables for donuts
const DSH_RANGE = "'📊 Dashboard'";

(async () => {
  // ── Same formulas as light dashboard ─────────────────────────────────────
  await valuesBatchUpdate(id, [
    { range: "'🌙 Dashboard — Dark Theme'!A1", values: [['🎄 ULTIMATE CHRISTMAS GIFT TRACKER 🎁']] },
    { range: "'🌙 Dashboard — Dark Theme'!A2", values: [['Track Every Gift, Card & Return — Reusable Every Year  ✨ DARK MODE']] },
    { range: "'🌙 Dashboard — Dark Theme'!A3", values: [['Total Budget ($):', `=${DSH_RANGE}!B3`, 'Holiday Date:', `=${DSH_RANGE}!D3`]] },
    // KPI labels
    { range: "'🌙 Dashboard — Dark Theme'!B5", values: [['❄️ Total Spent ($)', 'Left to Spend ($)', 'Gifts Bought %', 'Gifts Wrapped %']] },
    { range: "'🌙 Dashboard — Dark Theme'!B7", values: [[
      `=IFERROR(SUM(${GRT_RANGE}!F2:F26),0)`,
      `=IFERROR(${DSH_RANGE}!B3-SUM(${GRT_RANGE}!F2:F26),0)`,
      `=IFERROR(COUNTIF(${GRT_RANGE}!H2:H26,TRUE)/COUNTA(${GRT_RANGE}!A2:A26),"")`,
      `=IFERROR(COUNTIF(${GRT_RANGE}!I2:I26,TRUE)/COUNTA(${GRT_RANGE}!A2:A26),"")`,
    ]] },
    { range: "'🌙 Dashboard — Dark Theme'!B9", values: [['🎁 Gifts Sent %', '✉️ Cards Sent %', '🎄 Days Until Christmas']] },
    { range: "'🌙 Dashboard — Dark Theme'!B11", values: [[
      `=IFERROR(COUNTIF(${GRT_RANGE}!J2:J26,TRUE)/COUNTA(${GRT_RANGE}!A2:A26),"")`,
      `=IFERROR(COUNTIF(${CDL_RANGE}!C2:C21,"Sent")/COUNTA(${CDL_RANGE}!A2:A21),"")`,
      `=IF(DATE(YEAR(TODAY()),12,25)>=TODAY(),DATE(YEAR(TODAY()),12,25)-TODAY(),DATE(YEAR(TODAY())+1,12,25)-TODAY())`,
    ]] },
  ], 'dark-values');

  // ── Formatting ────────────────────────────────────────────────────────────
  const reqs = [];

  // Fill entire visible area with deep pine
  reqs.push({
    repeatCell: {
      range: gridRange(DSH2, 0, 50, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hexD(D.pine), textFormat: { foregroundColor: hexD(D.cream) } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // Merges (identical to light dashboard)
  reqs.push({ mergeCells: { range: gridRange(DSH2, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(DSH2, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } });
  for (const ri of [4, 6]) {
    reqs.push({ mergeCells: { range: gridRange(DSH2, ri, ri + 1, 1, 3), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DSH2, ri, ri + 1, 3, 5), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DSH2, ri, ri + 1, 5, 7), mergeType: 'MERGE_ALL' } });
  }
  for (const ri of [8, 10]) {
    reqs.push({ mergeCells: { range: gridRange(DSH2, ri, ri + 1, 1, 3), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DSH2, ri, ri + 1, 3, 5), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DSH2, ri, ri + 1, 5, 8), mergeType: 'MERGE_ALL' } });
  }

  // Title — cranberry accent on pine
  reqs.push({
    repeatCell: {
      range: gridRange(DSH2, 0, 1, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hexD(D.cranberry),
          textFormat: { foregroundColor: hexD(D.cream), fontSize: 18, bold: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Subtitle — darker cranberry
  reqs.push({
    repeatCell: {
      range: gridRange(DSH2, 1, 2, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hexD('#6B1520'),
          textFormat: { foregroundColor: hexD(D.lightGold), fontSize: 11, italic: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Input row — mid-pine
  reqs.push({
    repeatCell: {
      range: gridRange(DSH2, 2, 3, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hexD(D.midPine),
          textFormat: { foregroundColor: hexD(D.cream), fontSize: 11 },
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    },
  });
  reqs.push({ repeatCell: { range: gridRange(DSH2, 2, 3, 1, 2), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 12, foregroundColor: hexD(D.gold) }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '"$"#,##0' } } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment,numberFormat)' } });
  reqs.push({ repeatCell: { range: gridRange(DSH2, 2, 3, 3, 4), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 11, foregroundColor: hexD(D.gold) }, horizontalAlignment: 'CENTER', numberFormat: { type: 'DATE', pattern: 'MMM d, yyyy' } } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment,numberFormat)' } });

  // KPI label rows (idx 4 & 8) — gold on dark pine
  for (const ri of [4, 8]) {
    reqs.push({
      repeatCell: {
        range: gridRange(DSH2, ri, ri + 1, 1, 8),
        cell: {
          userEnteredFormat: {
            backgroundColor: hexD(D.darkGold),
            textFormat: { foregroundColor: hexD(D.cream), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
            wrapStrategy: 'WRAP',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
      },
    });
  }

  // KPI value rows (idx 6 & 10) — deep pine bg, gold text
  for (const ri of [6, 10]) {
    reqs.push({
      repeatCell: {
        range: gridRange(DSH2, ri, ri + 1, 1, 8),
        cell: {
          userEnteredFormat: {
            backgroundColor: hexD(D.deepPine),
            textFormat: { foregroundColor: hexD(D.gold), bold: true, fontSize: 26 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
  }

  // Number formats
  reqs.push({ repeatCell: { range: gridRange(DSH2, 6, 7, 1, 5), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat(numberFormat)' } });
  reqs.push({ repeatCell: { range: gridRange(DSH2, 6, 7, 5, 8), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0%' } } }, fields: 'userEnteredFormat(numberFormat)' } });
  reqs.push({ repeatCell: { range: gridRange(DSH2, 10, 11, 1, 5), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0%' } } }, fields: 'userEnteredFormat(numberFormat)' } });
  reqs.push({ repeatCell: { range: gridRange(DSH2, 10, 11, 5, 8), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0' } } }, fields: 'userEnteredFormat(numberFormat)' } });

  // Row heights (identical to light)
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH2, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH2, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH2, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 34 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH2, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH2, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH2, dimension: 'ROWS', startIndex: 8, endIndex: 9 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH2, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });

  // Column widths
  [[0,160],[1,140],[2,140],[3,140],[4,140],[5,140],[6,140],[7,140]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: DSH2, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'dark-format');

  // ── Charts on dark theme sheet ────────────────────────────────────────────
  function srcRange(sheetId, r1, r2, c1, c2) {
    return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
  }
  function overlayPos(row, col, w, h) {
    return { overlayPosition: { anchorCell: { sheetId: DSH2, rowIndex: row, columnIndex: col }, offsetXPixels: 4, offsetYPixels: 4, widthPixels: w, heightPixels: h } };
  }

  const darkBg = hexD(D.midPine);
  const darkText = hexD(D.cream);

  const chartReqs = [
    {
      addChart: {
        chart: {
          spec: {
            title: '🎁 Gifts Bought',
            titleTextFormat: { bold: true, fontSize: 11, foregroundColor: darkText },
            pieChart: {
              legendPosition: 'RIGHT_LEGEND', pieHole: 0.5,
              domain: { sourceRange: { sources: [srcRange(DSH, 30, 32, 0, 1)] } },
              series: { sourceRange: { sources: [srcRange(DSH, 30, 32, 1, 2)] } },
            },
            backgroundColor: darkBg,
          },
          position: overlayPos(13, 0, 350, 270),
        },
      },
    },
    {
      addChart: {
        chart: {
          spec: {
            title: '🎀 Gifts Wrapped',
            titleTextFormat: { bold: true, fontSize: 11, foregroundColor: darkText },
            pieChart: {
              legendPosition: 'RIGHT_LEGEND', pieHole: 0.5,
              domain: { sourceRange: { sources: [srcRange(DSH, 32, 34, 0, 1)] } },
              series: { sourceRange: { sources: [srcRange(DSH, 32, 34, 1, 2)] } },
            },
            backgroundColor: darkBg,
          },
          position: overlayPos(13, 4, 350, 270),
        },
      },
    },
    {
      addChart: {
        chart: {
          spec: {
            title: '📦 Gifts Delivered',
            titleTextFormat: { bold: true, fontSize: 11, foregroundColor: darkText },
            pieChart: {
              legendPosition: 'RIGHT_LEGEND', pieHole: 0.5,
              domain: { sourceRange: { sources: [srcRange(DSH, 34, 36, 0, 1)] } },
              series: { sourceRange: { sources: [srcRange(DSH, 34, 36, 1, 2)] } },
            },
            backgroundColor: darkBg,
          },
          position: overlayPos(20, 0, 350, 270),
        },
      },
    },
    {
      addChart: {
        chart: {
          spec: {
            title: '💰 Budget vs. Actual Spend by Recipient',
            titleTextFormat: { bold: true, fontSize: 11, foregroundColor: darkText },
            basicChart: {
              chartType: 'COLUMN',
              legendPosition: 'BOTTOM_LEGEND',
              axis: [
                { position: 'BOTTOM_AXIS', title: 'Recipient' },
                { position: 'LEFT_AXIS', title: 'Amount ($)' },
              ],
              domains: [{ domain: { sourceRange: { sources: [srcRange(GRT, 1, 26, 0, 1)] } } }],
              series: [
                { series: { sourceRange: { sources: [srcRange(GRT, 1, 26, 4, 5)] } }, targetAxis: 'LEFT_AXIS', color: hexD(D.gold) },
                { series: { sourceRange: { sources: [srcRange(GRT, 1, 26, 5, 6)] } }, targetAxis: 'LEFT_AXIS', color: hexD(D.cranberry) },
              ],
              stackedType: 'NOT_STACKED',
              headerCount: 1,
            },
            backgroundColor: darkBg,
          },
          position: overlayPos(20, 4, 550, 290),
        },
      },
    },
  ];

  await batchUpdate(id, chartReqs, 'dark-charts');
  console.log('Dark Theme Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
