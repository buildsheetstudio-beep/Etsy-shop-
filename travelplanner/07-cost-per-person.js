const { hex, COLOR, gridRange, batchUpdate, valuesBatchUpdate, vr } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = 1289386811; // Cost Per Person Calculator

(async () => {
  const reqs = [];

  // ── Row 1: Title ──
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 10),
      cell: {
        userEnteredValue: { stringValue: '👥 COST PER PERSON CALCULATOR' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.deepNavy),
          textFormat: { foregroundColor: hex(COLOR.antiqueGold), bold: true, fontSize: 18 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // ── Row 2: Subtitle ──
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, 10),
      cell: {
        userEnteredValue: { stringValue: 'Enter group details to see how costs split across travellers' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.midnightBlue),
          textFormat: { foregroundColor: hex(COLOR.champagne), italic: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // ── Section A: Trip Settings (rows 3-8) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 2, 3, 0, 10),
      cell: {
        userEnteredValue: { stringValue: '⚙  TRIP SETTINGS' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.midnightBlue),
          textFormat: { foregroundColor: hex(COLOR.antiqueGold), bold: true, fontSize: 12 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  const settingLabel = (r, c, text) => reqs.push({
    repeatCell: {
      range: gridRange(SID, r, r+1, c, c+1),
      cell: {
        userEnteredValue: { stringValue: text },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.champagne),
          textFormat: { foregroundColor: hex(COLOR.deepNavy), bold: true, fontSize: 10 },
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  const inputCell = (r, c, merge = 1) => {
    if (merge > 1) reqs.push({ mergeCells: { range: gridRange(SID, r, r+1, c, c+merge), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r+1, c, c+merge),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(COLOR.inputBg),
            borders: {
              bottom: { style: 'SOLID', color: hex(COLOR.antiqueGold) },
              top: { style: 'SOLID', color: hex(COLOR.antiqueGold) },
              left: { style: 'SOLID', color: hex(COLOR.antiqueGold) },
              right: { style: 'SOLID', color: hex(COLOR.antiqueGold) },
            },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat',
      },
    });
  };

  settingLabel(3, 0, 'Trip Name');
  inputCell(3, 1, 3);
  settingLabel(3, 5, 'Currency');
  inputCell(3, 6, 1);
  settingLabel(3, 8, 'Total Travellers');
  inputCell(3, 9, 1);

  settingLabel(4, 0, 'Total Budget (£)');
  inputCell(4, 1, 3);
  settingLabel(4, 5, 'Total Days');
  inputCell(4, 6, 1);
  settingLabel(4, 8, 'Sharing Rooms?');
  inputCell(4, 9, 1);

  // ── Section B: Cost Breakdown (rows 6-16) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 5, 6, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 5, 6, 0, 10),
      cell: {
        userEnteredValue: { stringValue: '💰  COST BREAKDOWN PER PERSON' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.midnightBlue),
          textFormat: { foregroundColor: hex(COLOR.antiqueGold), bold: true, fontSize: 12 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // Headers
  ['Cost Category','Total Cost (£)','Shared? (Y/N)','Cost Per Person (£)','% of Per-Person Total','Notes'].forEach((h, c) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 6, 7, c === 0 ? 0 : c+1, c === 0 ? 2 : c+2),
        cell: {
          userEnteredValue: { stringValue: h },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.deepNavy),
            textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 9 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });
  if (true) {
    reqs.push({ mergeCells: { range: gridRange(SID, 6, 7, 0, 2), mergeType: 'MERGE_ALL' } });
  }

  const categories = [
    ['Accommodation','£2,560','Y',null,'Pre-booked hotels & Airbnb'],
    ['Transport (Flights/Rail)','£636','Y',null,'Flights + Eurostar divided'],
    ['Local Transport','£85','Y',null,'Metro/bus passes'],
    ['Food & Drink','£803','Y',null,'Restaurants, cafes, markets'],
    ['Activities & Entry Fees','£191','Y',null,'Museums, landmarks, tours'],
    ['Shopping','£220','N',null,'Personal spend'],
    ['Insurance','£85','N',null,'Personal insurance'],
    ['Visa & Fees','£0','Y',null,'No visa required (EU/UK)'],
    ['Miscellaneous','£50','Y',null,'Tips, extras'],
  ];

  categories.forEach((cat, i) => {
    const row = 7 + i;
    const bg = i % 2 === 0 ? hex(COLOR.white) : hex(COLOR.altRow);
    reqs.push({ mergeCells: { range: gridRange(SID, row, row+1, 0, 2), mergeType: 'MERGE_ALL' } });
    // Cat name
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 0, 2),
        cell: {
          userEnteredFormat: { backgroundColor: bg, verticalAlignment: 'MIDDLE' },
        },
        fields: 'userEnteredFormat',
      },
    });
    // Total cost
    reqs.push({ repeatCell: { range: gridRange(SID, row, row+1, 2, 3), cell: { userEnteredFormat: { backgroundColor: bg, numberFormat: { type: 'NUMBER', pattern: '£#,##0.00' }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat' } });
    // Shared Y/N
    reqs.push({ repeatCell: { range: gridRange(SID, row, row+1, 3, 4), cell: { userEnteredFormat: { backgroundColor: bg, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat' } });
    // Per person (formula)
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 4, 5),
        cell: {
          userEnteredValue: { formulaValue: `=IF(C${row+1}="","",IF(D${row+1}="Y",C${row+1}/MAX($J$4,1),C${row+1}))` },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.formulaBg),
            numberFormat: { type: 'NUMBER', pattern: '£#,##0.00' },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
    // % formula
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 5, 6),
        cell: {
          userEnteredValue: { formulaValue: `=IF(E${row+1}="","",E${row+1}/SUM(E$8:E$16))` },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.formulaBg),
            numberFormat: { type: 'NUMBER', pattern: '0.0%' },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
    // Notes
    reqs.push({ repeatCell: { range: gridRange(SID, row, row+1, 6, 10), cell: { userEnteredFormat: { backgroundColor: bg, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat' } });
    if (true) reqs.push({ mergeCells: { range: gridRange(SID, row, row+1, 6, 10), mergeType: 'MERGE_ALL' } });
  });

  // Totals row
  const totRow = 7 + categories.length;
  reqs.push({ mergeCells: { range: gridRange(SID, totRow, totRow+1, 0, 2), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, totRow, totRow+1, 0, 2),
      cell: {
        userEnteredValue: { stringValue: 'TOTAL' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.antiqueGold),
          textFormat: { foregroundColor: hex(COLOR.deepNavy), bold: true, fontSize: 11 },
          horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  [2, 4].forEach(c => {
    const letter = ['C','','E'][c-2];
    reqs.push({
      repeatCell: {
        range: gridRange(SID, totRow, totRow+1, c, c+1),
        cell: {
          userEnteredValue: { formulaValue: `=SUM(${letter}8:${letter}${totRow})` },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.antiqueGold),
            textFormat: { foregroundColor: hex(COLOR.deepNavy), bold: true, fontSize: 12 },
            numberFormat: { type: 'NUMBER', pattern: '£#,##0.00' },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });

  // ── Section C: Per-Person Summary KPIs (rows 19-22) ──
  const kpiRow = totRow + 2;
  reqs.push({ mergeCells: { range: gridRange(SID, kpiRow, kpiRow+1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, kpiRow, kpiRow+1, 0, 10),
      cell: {
        userEnteredValue: { stringValue: '📊  PER-PERSON SUMMARY' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.midnightBlue),
          textFormat: { foregroundColor: hex(COLOR.antiqueGold), bold: true, fontSize: 12 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  const summaryKPIs = [
    { label: 'Cost Per Person', formula: `=SUM(E8:E${totRow})`, fmt: '£#,##0.00' },
    { label: 'Daily Budget Per Person', formula: `=IFERROR(SUM(E8:E${totRow})/MAX($J$5,1),"—")`, fmt: '£#,##0.00' },
    { label: 'Total Travellers', formula: '=$J$4', fmt: '0' },
    { label: 'Sharing Rooms', formula: '=$J$5', fmt: '@' },
    { label: 'Total Group Cost', formula: `=SUM(C8:C${totRow})`, fmt: '£#,##0.00' },
  ];

  summaryKPIs.forEach((k, i) => {
    const r = kpiRow + 1;
    const c = i * 2;
    if (i < 5) {
      reqs.push({ mergeCells: { range: gridRange(SID, r, r+1, c, c+1), mergeType: 'MERGE_ALL' } });
      reqs.push({
        repeatCell: {
          range: gridRange(SID, r, r+1, c, c+1),
          cell: {
            userEnteredValue: { stringValue: k.label },
            userEnteredFormat: {
              backgroundColor: hex(COLOR.champagne),
              textFormat: { foregroundColor: hex(COLOR.deepNavy), bold: true, fontSize: 9 },
              horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM',
            },
          },
          fields: 'userEnteredValue,userEnteredFormat',
        },
      });
      reqs.push({ mergeCells: { range: gridRange(SID, r+1, r+2, c, c+1), mergeType: 'MERGE_ALL' } });
      reqs.push({
        repeatCell: {
          range: gridRange(SID, r+1, r+2, c, c+1),
          cell: {
            userEnteredValue: { formulaValue: k.formula },
            userEnteredFormat: {
              backgroundColor: hex(COLOR.formulaBg),
              textFormat: { foregroundColor: hex(COLOR.midnightBlue), bold: true, fontSize: 14 },
              numberFormat: { type: 'NUMBER', pattern: k.fmt },
              horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
            },
          },
          fields: 'userEnteredValue,userEnteredFormat',
        },
      });
    }
  });

  // Column widths
  const widths = [120,120,110,80,120,80,130,100,100,80];
  widths.forEach((w, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: c, endIndex: c+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 25 }, properties: { pixelSize: 34 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'cost-per-person-structure');

  // ── Values ──
  const data = [];
  data.push(vr(`'Cost Per Person Calculator'!B4`, [['European Summer Escape']]));
  data.push(vr(`'Cost Per Person Calculator'!G4`, [['GBP']]));
  data.push(vr(`'Cost Per Person Calculator'!J4`, [[2]])); // travellers
  data.push(vr(`'Cost Per Person Calculator'!B5`, [[4500]])); // budget
  data.push(vr(`'Cost Per Person Calculator'!G5`, [[22]])); // days
  data.push(vr(`'Cost Per Person Calculator'!J5`, [['Yes']])); // sharing rooms

  const catRows = categories.map(cat => [cat[0], cat[1].replace('£',''), cat[2], '', cat[4]]);
  data.push(vr(`'Cost Per Person Calculator'!A8`, catRows.map(r => [r[0]])));
  data.push(vr(`'Cost Per Person Calculator'!C8`, catRows.map(r => [r[1]])));
  data.push(vr(`'Cost Per Person Calculator'!D8`, catRows.map(r => [r[2]])));
  data.push(vr(`'Cost Per Person Calculator'!G8`, catRows.map(r => [r[4]])));

  await valuesBatchUpdate(id, data, 'cost-per-person-values');

  console.log('Cost Per Person Calculator built');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
