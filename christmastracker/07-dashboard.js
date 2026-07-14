'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard'];

// Layout:
//   Row 1 (idx 0): Title merged A1:H1
//   Row 2 (idx 1): Subtitle merged A2:H2
//   Row 3 (idx 2): Inputs — A3="Total Budget ($):", B3=value, C3="🎄 Christmas:", D3=date
//   Row 4 (idx 3): Days Until Christmas (formula) merged A4:H4
//   Row 5 (idx 4): Blank spacer
//   Rows 6-7 (idx 5-6): KPI label row — 4 KPIs
//   Rows 8-9 (idx 7-8): KPI value row
//   Rows 10-11 (idx 9-10): KPI label row 2 — 3 KPIs
//   Rows 12-13 (idx 11-12): KPI value row 2
//   Row 14 (idx 13): blank spacer
//   Rows 15-20 (idx 14-19): Chart area (charts overlay)
//   Row 21 (idx 20): blank
//   Row 22 (idx 21): Chart data section header
//   Row 23 (idx 22): Bought header
//   Rows 24-25 (idx 23-24): Bought data (Bought / Not Bought)
//   Row 26 (idx 25): blank
//   Row 27 (idx 26): Wrapped header
//   Rows 28-29 (idx 27-28): Wrapped data
//   Row 30 (idx 29): blank
//   Row 31 (idx 30): Sent header
//   Rows 32-33 (idx 31-32): Sent data
//   Row 34 (idx 33): blank
//   Row 35 (idx 34): Card Status header
//   Rows 36-38 (idx 35-37): Card Status data (Sent/Not Sent/Received)
//   Row 39 (idx 38): blank
//   Row 40 (idx 39): Budget summary header
//   Rows 41-43 (idx 40-42): Budget data rows

(async () => {
  // ── Static text & values ─────────────────────────────────────────────────
  await valuesBatchUpdate(id, [
    { range: "'📊 Dashboard'!A1", values: [['🎄 Ultimate Christmas Gift Tracker']] },
    { range: "'📊 Dashboard'!A2", values: [['Track Every Gift, Card & Return — Reusable Every Year']] },
    { range: "'📊 Dashboard'!A3", values: [['Total Budget ($):',1200,'🎄 Christmas:','2025-12-25']] },
    { range: "'📊 Dashboard'!A4", values: [[
      `=IF(DATE(YEAR(TODAY()),12,25)>=TODAY(),DATE(YEAR(TODAY()),12,25)-TODAY(),DATE(YEAR(TODAY())+1,12,25)-TODAY())&" days until Christmas 🎄"`,
    ]] },
    // KPI row 1 labels (row 6)
    { range: "'📊 Dashboard'!B6", values: [['Gifts Bought','Gifts Wrapped','Cards Sent','Cards Received']] },
    // KPI row 1 values (row 8)
    { range: "'📊 Dashboard'!B8", values: [[
      `=COUNTIF('🎁 Gift Recipient Tracker'!H:H,TRUE)`,
      `=COUNTIF('🎁 Gift Recipient Tracker'!I:I,TRUE)`,
      `=COUNTIF('💌 Holiday Card & Mailing List'!C:C,"Sent")`,
      `=COUNTIF('💌 Holiday Card & Mailing List'!D:D,TRUE)`,
    ]] },
    // KPI row 2 labels (row 10)
    { range: "'📊 Dashboard'!B10", values: [['Total Spent ($)','Budget Left ($)','Returns Pending']] },
    // KPI row 2 values (row 12)
    { range: "'📊 Dashboard'!B12", values: [[
      `=IFERROR(SUM('🎁 Gift Recipient Tracker'!F:F),0)`,
      `=IFERROR(B3-SUM('🎁 Gift Recipient Tracker'!F:F),0)`,
      `=COUNTIFS('🔄 Return & Exchange Tracker'!H:H,"<>Completed",'🔄 Return & Exchange Tracker'!H:H,"<>Denied",'🔄 Return & Exchange Tracker'!H:H,"<>")`,
    ]] },
    // Chart data section header
    { range: "'📊 Dashboard'!A22", values: [['— Chart Data (Internal) —']] },
    // Bought donut data
    { range: "'📊 Dashboard'!A23", values: [['Gift Status','Count']] },
    { range: "'📊 Dashboard'!A24", values: [
      ['Bought', `=COUNTIF('🎁 Gift Recipient Tracker'!H:H,TRUE)`],
      ['Not Bought', `=COUNTIF('🎁 Gift Recipient Tracker'!A:A,"<>")-COUNTIF('🎁 Gift Recipient Tracker'!H:H,TRUE)-1`],
    ] },
    // Wrapped donut data
    { range: "'📊 Dashboard'!A27", values: [['Wrap Status','Count']] },
    { range: "'📊 Dashboard'!A28", values: [
      ['Wrapped', `=COUNTIF('🎁 Gift Recipient Tracker'!I:I,TRUE)`],
      ['Not Wrapped', `=COUNTIF('🎁 Gift Recipient Tracker'!A:A,"<>")-COUNTIF('🎁 Gift Recipient Tracker'!I:I,TRUE)-1`],
    ] },
    // Sent donut data
    { range: "'📊 Dashboard'!A31", values: [['Delivery Status','Count']] },
    { range: "'📊 Dashboard'!A32", values: [
      ['Delivered', `=COUNTIF('🎁 Gift Recipient Tracker'!J:J,TRUE)`],
      ['Not Delivered', `=COUNTIF('🎁 Gift Recipient Tracker'!A:A,"<>")-COUNTIF('🎁 Gift Recipient Tracker'!J:J,TRUE)-1`],
    ] },
    // Card status column chart data
    { range: "'📊 Dashboard'!A35", values: [['Card Status','Count']] },
    { range: "'📊 Dashboard'!A36", values: [
      ['Sent', `=COUNTIF('💌 Holiday Card & Mailing List'!C:C,"Sent")`],
      ['Not Sent', `=COUNTIF('💌 Holiday Card & Mailing List'!C:C,"Not Sent")`],
      ['Received Back', `=COUNTIF('💌 Holiday Card & Mailing List'!D:D,TRUE)`],
    ] },
    // Budget data
    { range: "'📊 Dashboard'!A40", values: [['Budget Summary','Amount ($)']] },
    { range: "'📊 Dashboard'!A41", values: [
      ['Total Budget', `=B3`],
      ['Total Spent', `=IFERROR(SUM('🎁 Gift Recipient Tracker'!F:F),0)`],
      ['Remaining', `=IFERROR(B3-SUM('🎁 Gift Recipient Tracker'!F:F),0)`],
    ] },
  ], 'dashboard-data');

  // ── Formatting ────────────────────────────────────────────────────────────
  const reqs = [];

  // Merges
  reqs.push({ mergeCells: { range: gridRange(DSH, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } }); // Title
  reqs.push({ mergeCells: { range: gridRange(DSH, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } }); // Subtitle
  reqs.push({ mergeCells: { range: gridRange(DSH, 3, 4, 0, 8), mergeType: 'MERGE_ALL' } }); // Days until Xmas
  // KPI row 1 merges: B6:C6 for Gifts Bought, D6:E6 for Gifts Wrapped, F6:G6 for Cards Sent, H merges self
  reqs.push({ mergeCells: { range: gridRange(DSH, 5, 6, 1, 3), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(DSH, 5, 6, 3, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(DSH, 5, 6, 5, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(DSH, 7, 8, 1, 3), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(DSH, 7, 8, 3, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(DSH, 7, 8, 5, 7), mergeType: 'MERGE_ALL' } });
  // KPI row 2 merges: B10:C10, D10:E10, F10:G10
  reqs.push({ mergeCells: { range: gridRange(DSH, 9, 10, 1, 3), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(DSH, 9, 10, 3, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(DSH, 9, 10, 5, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(DSH, 11, 12, 1, 3), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(DSH, 11, 12, 3, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(DSH, 11, 12, 5, 7), mergeType: 'MERGE_ALL' } });

  // Title row
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

  // Subtitle row
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

  // Input row A3
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
  // Input values B3, D3 bold
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 2, 3, 1, 2),
      cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 12 }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '"$"#,##0' } } },
      fields: 'userEnteredFormat(textFormat,horizontalAlignment,numberFormat)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 2, 3, 3, 4),
      cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 11 }, horizontalAlignment: 'CENTER', numberFormat: { type: 'DATE', pattern: 'MMM d, yyyy' } } },
      fields: 'userEnteredFormat(textFormat,horizontalAlignment,numberFormat)',
    },
  });

  // Days until Christmas row (A4)
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 3, 4, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepGreen),
          textFormat: { foregroundColor: hex(C.lightAmber), fontSize: 13, bold: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // KPI label rows (rows 6,10 — idx 5,9)
  for (const labelRow of [5, 9]) {
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, labelRow, labelRow + 1, 1, 8),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.warmGold),
            textFormat: { foregroundColor: hex(C.darkText), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
  }

  // KPI value rows (rows 8,12 — idx 7,11)
  for (const valRow of [7, 11]) {
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, valRow, valRow + 1, 1, 8),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.lightAmber),
            textFormat: { foregroundColor: hex(C.deepCranberry), bold: true, fontSize: 24 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
  }

  // Currency format for Total Spent & Budget Left KPIs (idx 11, cols D:E and F:G)
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 11, 12, 3, 7),
      cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '"$"#,##0.00' } } },
      fields: 'userEnteredFormat(numberFormat)',
    },
  });

  // Chart data section — dim text
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 21, 44, 0, 2),
      cell: {
        userEnteredFormat: {
          textFormat: { foregroundColor: hex(C.mediumGray), fontSize: 9 },
        },
      },
      fields: 'userEnteredFormat(textFormat)',
    },
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 34 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 34 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 7, endIndex: 8 }, properties: { pixelSize: 54 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 9, endIndex: 10 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 11, endIndex: 12 }, properties: { pixelSize: 54 }, fields: 'pixelSize' } });

  // Column widths
  [[0,160],[1,130],[2,130],[3,130],[4,130],[5,130],[6,130],[7,130]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'dashboard-format');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
