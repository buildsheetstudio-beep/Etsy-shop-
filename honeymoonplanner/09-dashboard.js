'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const OVW = sheetMap['🌅 Honeymoon Overview'];
const BGT = sheetMap['💰 Budget & Expense Tracker'];
const CHK = sheetMap['✅ Research & Reservation Checklist'];
const PKG = sheetMap['🧳 Packing List'];
const VLT = sheetMap['🔐 Travel Document & Confirmation Vault'];
const WIS = sheetMap['✨ Activity & Excursion Wishlist'];

const BGT_RANGE = "'💰 Budget & Expense Tracker'";
const CHK_RANGE = "'✅ Research & Reservation Checklist'";
const PKG_RANGE = "'🧳 Packing List'";
const VLT_RANGE = "'🔐 Travel Document & Confirmation Vault'";
const WIS_RANGE = "'✨ Activity & Excursion Wishlist'";

(async () => {
  // ── Values ────────────────────────────────────────────────────────────────
  await valuesBatchUpdate(id, [
    // Title
    { range: "'🌅 Honeymoon Overview'!A1", values: [['🌅 ULTIMATE HONEYMOON PLANNER ✈️']] },
    { range: "'🌅 Honeymoon Overview'!A2", values: [['Track Every Flight, Hotel & Activity — Your Dream Honeymoon, Perfectly Organised 💕']] },

    // Input row (row 3)
    { range: "'🌅 Honeymoon Overview'!A3", values: [['Destination(s):', '', 'Departure Date:', '', 'Return Date:', '', 'Total Budget ($):', '']] },

    // KPI label row 1 (row 5)
    { range: "'🌅 Honeymoon Overview'!B5", values: [['💰 Total Estimated ($)', 'Total Actual Spent ($)', '📊 Budget Remaining ($)', '✅ Checklist Done %']] },
    // KPI value row 1 (row 7)
    { range: "'🌅 Honeymoon Overview'!B7", values: [[
      `=IFERROR(SUM(${BGT_RANGE}!C2:C26),0)`,
      `=IFERROR(SUM(${BGT_RANGE}!D2:D26),0)`,
      `=IFERROR(H3-SUM(${BGT_RANGE}!D2:D26),0)`,
      `=IFERROR(COUNTIF(${CHK_RANGE}!E2:E26,TRUE)/COUNTA(${CHK_RANGE}!A2:A26),"")`,
    ]] },

    // KPI label row 2 (row 9)
    { range: "'🌅 Honeymoon Overview'!B9", values: [['🧳 Items Packed %', '✈️ Days Until Departure', '🏨 Bookings Confirmed', '📄 Documents Confirmed']] },
    // KPI value row 2 (row 11)
    { range: "'🌅 Honeymoon Overview'!B11", values: [[
      `=IFERROR(COUNTIF(${PKG_RANGE}!D2:D34,TRUE)/COUNTA(${PKG_RANGE}!A2:A34),"")`,
      `=IF(D3="","",IF(D3>=TODAY(),D3-TODAY(),0))`,
      `=IFERROR(COUNTIF(${VLT_RANGE}!E2:E16,"Confirmed"),0)`,
      `=IFERROR(COUNTIF(${VLT_RANGE}!E2:E16,"Confirmed"),0)`,
    ]] },

    // Currency converter block (rows 13-16)
    { range: "'🌅 Honeymoon Overview'!A13", values: [['💱 CURRENCY CONVERTER']] },
    { range: "'🌅 Honeymoon Overview'!A14", values: [['Amount:', '', 'From Currency:', '', 'Exchange Rate (manual):', '', 'Converted Amount:', '']] },
    { range: "'🌅 Honeymoon Overview'!B15", values: [[1000]] },
    { range: "'🌅 Honeymoon Overview'!D15", values: [['USD']] },
    { range: "'🌅 Honeymoon Overview'!F15", values: [[0.92]] },
    { range: "'🌅 Honeymoon Overview'!H15", values: [['=IFERROR(B15*F15,"")']] },
    { range: "'🌅 Honeymoon Overview'!A16", values: [['ℹ️ Enter your exchange rate manually above — update it daily for accuracy.']] },

    // Budget breakdown helper table for chart (rows 32-39, cols A-B)
    { range: "'🌅 Honeymoon Overview'!A32", values: [['Category', 'Total Actual ($)']] },
    { range: "'🌅 Honeymoon Overview'!A33", values: [[
      'Flights',          `=IFERROR(SUMIF(${BGT_RANGE}!A2:A26,"Flights",${BGT_RANGE}!D2:D26),0)`,
    ]] },
    { range: "'🌅 Honeymoon Overview'!A34", values: [[
      'Accommodation',    `=IFERROR(SUMIF(${BGT_RANGE}!A2:A26,"Accommodation",${BGT_RANGE}!D2:D26),0)`,
    ]] },
    { range: "'🌅 Honeymoon Overview'!A35", values: [[
      'Food & Dining',    `=IFERROR(SUMIF(${BGT_RANGE}!A2:A26,"Food & Dining",${BGT_RANGE}!D2:D26),0)`,
    ]] },
    { range: "'🌅 Honeymoon Overview'!A36", values: [[
      'Activities & Tours',`=IFERROR(SUMIF(${BGT_RANGE}!A2:A26,"Activities & Tours",${BGT_RANGE}!D2:D26),0)`,
    ]] },
    { range: "'🌅 Honeymoon Overview'!A37", values: [[
      'Transport',        `=IFERROR(SUMIF(${BGT_RANGE}!A2:A26,"Transport",${BGT_RANGE}!D2:D26),0)`,
    ]] },
    { range: "'🌅 Honeymoon Overview'!A38", values: [[
      'Shopping',         `=IFERROR(SUMIF(${BGT_RANGE}!A2:A26,"Shopping",${BGT_RANGE}!D2:D26),0)`,
    ]] },
    { range: "'🌅 Honeymoon Overview'!A39", values: [[
      'Miscellaneous',    `=IFERROR(SUMIF(${BGT_RANGE}!A2:A26,"Miscellaneous",${BGT_RANGE}!D2:D26),0)`,
    ]] },
  ], 'ovw-values');

  // ── Formatting ────────────────────────────────────────────────────────────
  const reqs = [];

  // Base fill — warm ivory
  reqs.push({
    repeatCell: {
      range: gridRange(OVW, 0, 50, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.warmIvory), textFormat: { foregroundColor: hex(C.warmCharcoal) } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // Merges
  reqs.push({ mergeCells: { range: gridRange(OVW, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(OVW, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } });
  // KPI label rows
  for (const ri of [4, 8]) {
    reqs.push({ mergeCells: { range: gridRange(OVW, ri, ri+1, 1, 3), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(OVW, ri, ri+1, 3, 5), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(OVW, ri, ri+1, 5, 7), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(OVW, ri, ri+1, 7, 8), mergeType: 'MERGE_ALL' } });
  }
  for (const ri of [6, 10]) {
    reqs.push({ mergeCells: { range: gridRange(OVW, ri, ri+1, 1, 3), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(OVW, ri, ri+1, 3, 5), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(OVW, ri, ri+1, 5, 7), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(OVW, ri, ri+1, 7, 8), mergeType: 'MERGE_ALL' } });
  }
  // Currency converter
  reqs.push({ mergeCells: { range: gridRange(OVW, 12, 13, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(OVW, 15, 16, 0, 8), mergeType: 'MERGE_ALL' } });

  // Title — dusty peach
  reqs.push({
    repeatCell: {
      range: gridRange(OVW, 0, 1, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.dustyPeach),
          textFormat: { foregroundColor: hex(C.warmIvory), fontSize: 20, bold: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Subtitle — deeper peach
  reqs.push({
    repeatCell: {
      range: gridRange(OVW, 1, 2, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex('#C8855A'),
          textFormat: { foregroundColor: hex(C.warmIvory), fontSize: 11, italic: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Input row 3
  reqs.push({
    repeatCell: {
      range: gridRange(OVW, 2, 3, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.altRow),
          textFormat: { foregroundColor: hex(C.warmCharcoal), fontSize: 11 },
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    },
  });
  // Input cells: B3 (destination text), D3 (start date), F3 (end date), H3 (budget)
  reqs.push({ repeatCell: { range: gridRange(OVW, 2, 3, 1, 2), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.dustyPeach) }, horizontalAlignment: 'CENTER', backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment,backgroundColor)' } });
  reqs.push({ repeatCell: { range: gridRange(OVW, 2, 3, 3, 4), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.slateTeal) }, horizontalAlignment: 'CENTER', numberFormat: { type: 'DATE', pattern: 'MMM d, yyyy' }, backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment,numberFormat,backgroundColor)' } });
  reqs.push({ repeatCell: { range: gridRange(OVW, 2, 3, 5, 6), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.slateTeal) }, horizontalAlignment: 'CENTER', numberFormat: { type: 'DATE', pattern: 'MMM d, yyyy' }, backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment,numberFormat,backgroundColor)' } });
  reqs.push({ repeatCell: { range: gridRange(OVW, 2, 3, 7, 8), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.amber) }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '"$"#,##0' }, backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment,numberFormat,backgroundColor)' } });

  // KPI label rows (4 & 8) — slate teal
  for (const ri of [4, 8]) {
    reqs.push({
      repeatCell: {
        range: gridRange(OVW, ri, ri+1, 1, 8),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.slateTeal),
            textFormat: { foregroundColor: hex(C.warmIvory), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
            wrapStrategy: 'WRAP',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
      },
    });
  }

  // KPI value rows (6 & 10) — warm charcoal bg, dusty peach text
  for (const ri of [6, 10]) {
    reqs.push({
      repeatCell: {
        range: gridRange(OVW, ri, ri+1, 1, 8),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.warmCharcoal),
            textFormat: { foregroundColor: hex(C.dustyPeach), bold: true, fontSize: 26 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
  }

  // Number formats for KPI values
  reqs.push({ repeatCell: { range: gridRange(OVW, 6, 7, 1, 7), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat(numberFormat)' } });
  reqs.push({ repeatCell: { range: gridRange(OVW, 6, 7, 7, 8), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0%' } } }, fields: 'userEnteredFormat(numberFormat)' } });
  reqs.push({ repeatCell: { range: gridRange(OVW, 10, 11, 1, 2), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0%' } } }, fields: 'userEnteredFormat(numberFormat)' } });
  reqs.push({ repeatCell: { range: gridRange(OVW, 10, 11, 3, 4), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0' } } }, fields: 'userEnteredFormat(numberFormat)' } });
  reqs.push({ repeatCell: { range: gridRange(OVW, 10, 11, 5, 8), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0' } } }, fields: 'userEnteredFormat(numberFormat)' } });

  // Currency converter header
  reqs.push({
    repeatCell: {
      range: gridRange(OVW, 12, 13, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.amber),
          textFormat: { foregroundColor: hex(C.warmCharcoal), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Currency converter input row (row 14, idx 13)
  reqs.push({
    repeatCell: {
      range: gridRange(OVW, 13, 14, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { foregroundColor: hex(C.warmCharcoal) }, verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    },
  });
  // Input cells in row 14 (idx 13): B=Amount, D=Currency, F=Rate, H=Result
  reqs.push({ repeatCell: { range: gridRange(OVW, 14, 15, 1, 2), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.amber) }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '#,##0.00' }, backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment,numberFormat,backgroundColor)' } });
  reqs.push({ repeatCell: { range: gridRange(OVW, 14, 15, 3, 4), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.slateTeal) }, horizontalAlignment: 'CENTER', backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment,backgroundColor)' } });
  reqs.push({ repeatCell: { range: gridRange(OVW, 14, 15, 5, 6), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.warmCharcoal) }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '0.0000' }, backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment,numberFormat,backgroundColor)' } });
  reqs.push({ repeatCell: { range: gridRange(OVW, 14, 15, 7, 8), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.dustyPeach) }, horizontalAlignment: 'CENTER', numberFormat: { type: 'NUMBER', pattern: '#,##0.00' }, backgroundColor: hex(C.formulaBg) } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment,numberFormat,backgroundColor)' } });

  // Info row (row 16, idx 15)
  reqs.push({
    repeatCell: {
      range: gridRange(OVW, 15, 16, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lightTeal), textFormat: { foregroundColor: hex(C.deepTeal), fontSize: 10, italic: true }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: OVW, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 60 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: OVW, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: OVW, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: OVW, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: OVW, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 60 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: OVW, dimension: 'ROWS', startIndex: 8, endIndex: 9 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: OVW, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 60 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: OVW, dimension: 'ROWS', startIndex: 12, endIndex: 13 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: OVW, dimension: 'ROWS', startIndex: 13, endIndex: 14 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: OVW, dimension: 'ROWS', startIndex: 14, endIndex: 15 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: OVW, dimension: 'ROWS', startIndex: 15, endIndex: 16 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths
  [[0,160],[1,150],[2,150],[3,150],[4,150],[5,150],[6,150],[7,150]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: OVW, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'ovw-format');
  console.log('Dashboard Overview complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
