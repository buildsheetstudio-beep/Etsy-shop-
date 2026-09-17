const { hex, COLOR, gridRange, batchUpdate, valuesBatchUpdate, vr } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = 1614915889; // Currency Converter

// Exchange rates vs GBP (1 GBP = X foreign)
const rates = [
  ['GBP', 1.0000, '🇬🇧 British Pound', '£'],
  ['USD', 1.2720, '🇺🇸 US Dollar', '$'],
  ['EUR', 1.1890, '🇪🇺 Euro', '€'],
  ['AUD', 1.9580, '🇦🇺 Australian Dollar', 'A$'],
  ['JPY', 197.45, '🇯🇵 Japanese Yen', '¥'],
  ['THB', 44.82, '🇹🇭 Thai Baht', '฿'],
  ['AED', 4.6710, '🇦🇪 UAE Dirham', 'د.إ'],
  ['MXN', 21.93, '🇲🇽 Mexican Peso', '$'],
  ['SGD', 1.7140, '🇸🇬 Singapore Dollar', 'S$'],
];

(async () => {
  const reqs = [];

  // ── Title row (A1:H1) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 8),
      cell: {
        userEnteredValue: { stringValue: '✈ CURRENCY CONVERTER' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.deepNavy),
          textFormat: { foregroundColor: hex(COLOR.antiqueGold), bold: true, fontSize: 16 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // ── Subtitle row (A2:H2) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, 8),
      cell: {
        userEnteredValue: { stringValue: 'Live rates relative to GBP — update the rate column to refresh all conversions' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.midnightBlue),
          textFormat: { foregroundColor: hex(COLOR.champagne), italic: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // ── Section A: Exchange Rate Table (rows 3-12) ──
  // Headers row 3: Code | Currency | Symbol | Rate (per 1 GBP) | 1 GBP → | 10 GBP → | 100 GBP →
  const rateHeaders = ['Code','Currency Name','Symbol','Rate (per £1)','£1 →','£10 →','£100 →','£500 →'];
  rateHeaders.forEach((h, c) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 3, 4, c, c+1),
        cell: {
          userEnteredValue: { stringValue: h },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.midnightBlue),
            textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });

  // Rate data rows 4-12
  rates.forEach((r, i) => {
    const row = 4 + i;
    const isGBP = r[0] === 'GBP';
    const bg = isGBP ? hex(COLOR.lightGold) : (i % 2 === 0 ? hex(COLOR.white) : hex(COLOR.altRow));

    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 0, 8),
        cell: { userEnteredFormat: { backgroundColor: bg } },
        fields: 'userEnteredFormat.backgroundColor',
      },
    });
  });

  // Column widths
  const colWidths = [70, 200, 70, 130, 100, 100, 100, 100];
  colWidths.forEach((w, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: c, endIndex: c+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 20 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 14 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // ── Section B: Quick Converter (rows 14-22) ──
  // Spacer row 13
  reqs.push({ mergeCells: { range: gridRange(SID, 14, 15, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 14, 15, 0, 8),
      cell: {
        userEnteredValue: { stringValue: '⚡ QUICK CONVERTER' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.midnightBlue),
          textFormat: { foregroundColor: hex(COLOR.antiqueGold), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // Input row headers (row 15)
  const converterLabels = ['Amount (GBP)','','','From','','','To','Result'];
  converterLabels.forEach((h, c) => {
    if (!h) return;
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 15, 16, c, c+1),
        cell: {
          userEnteredValue: { stringValue: h },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.champagne),
            textFormat: { foregroundColor: hex(COLOR.deepNavy), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });

  // Input cells (row 16) - bg = inputBg with gold border
  [0, 3, 6].forEach(c => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 16, 17, c, c+1),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(COLOR.inputBg),
            borders: {
              bottom: { style: 'SOLID', color: hex(COLOR.antiqueGold) },
              top: { style: 'SOLID', color: hex(COLOR.antiqueGold) },
              left: { style: 'SOLID', color: hex(COLOR.antiqueGold) },
              right: { style: 'SOLID', color: hex(COLOR.antiqueGold) },
            },
          },
        },
        fields: 'userEnteredFormat',
      },
    });
  });

  // Result cell (row 16, col 7) - formula bg
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 16, 17, 7, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(COLOR.formulaBg),
          textFormat: { bold: true, fontSize: 12, foregroundColor: hex(COLOR.forestGreen) },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat',
    },
  });

  // All currencies quick view - rows 18-27 (amount in each currency)
  reqs.push({ mergeCells: { range: gridRange(SID, 18, 19, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 18, 19, 0, 8),
      cell: {
        userEnteredValue: { stringValue: 'Value of your £1 GBP input in all currencies' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.champagne),
          textFormat: { foregroundColor: hex(COLOR.deepNavy), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  await batchUpdate(id, reqs, 'currency-structure');

  // ── Values ──
  const data = [];

  // Exchange rate table values with formulas
  const rateRows = rates.map((r, i) => {
    const rowNum = 5 + i; // 1-indexed for formulas
    // Cols: Code | Name | Symbol | Rate | £1→ | £10→ | £100→ | £500→
    return [
      r[0],
      r[2],
      r[3],
      r[1],
      `=D${rowNum}*1`,
      `=D${rowNum}*10`,
      `=D${rowNum}*100`,
      `=D${rowNum}*500`,
    ];
  });
  data.push(vr(`'Currency Converter'!A5`, rateRows));

  // Quick converter inputs
  data.push(vr(`'Currency Converter'!A17`, [[100]]));        // Amount
  data.push(vr(`'Currency Converter'!D17`, [['GBP']]));      // From
  data.push(vr(`'Currency Converter'!G17`, [['EUR']]));      // To

  // Quick converter result formula
  // =A17 * VLOOKUP(G17, A5:D13, 4, FALSE) / VLOOKUP(D17, A5:D13, 4, FALSE)
  data.push(vr(`'Currency Converter'!H17`, [['=IFERROR(A17*VLOOKUP(G17,$A$5:$D$13,4,FALSE)/VLOOKUP(D17,$A$5:$D$13,4,FALSE),"—")' ]]));

  // All-currencies view (rows 19-27): each currency equivalent of A17 GBP
  const currRows = rates.map((r, i) => {
    const rateCellRow = 5 + i;
    return [`${r[3]} ${r[2]}`, `=IFERROR($A$17*D${rateCellRow},"—")`];
  });
  data.push(vr(`'Currency Converter'!A19`, currRows));

  await valuesBatchUpdate(id, data, 'currency-values');

  // Number formats for rate table
  const numReqs = [];
  rates.forEach((r, i) => {
    const row = 4 + i;
    // Rate column (D): 4 decimal places
    numReqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 3, 4),
        cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0.0000' } } },
        fields: 'userEnteredFormat.numberFormat',
      },
    });
    // £1/£10/£100/£500 columns: 2 decimal (or integer for JPY/THB)
    const pattern = (r[0] === 'JPY' || r[0] === 'THB') ? '#,##0' : '#,##0.00';
    numReqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 4, 8),
        cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern } } },
        fields: 'userEnteredFormat.numberFormat',
      },
    });
  });

  // Result cell number format
  numReqs.push({
    repeatCell: {
      range: gridRange(SID, 16, 17, 7, 8),
      cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  await batchUpdate(id, numReqs, 'currency-numfmt');

  console.log('Currency Converter built with', rates.length, 'currencies');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
