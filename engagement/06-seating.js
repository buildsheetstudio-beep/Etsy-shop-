'use strict';
const { C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SEA = sheetMap['🪑 Seating Chart'];

// Cols A:O (15 cols, indices 0-14)
// A=Table# B=Table Name C=Shape D=Capacity E=Assigned(COUNTA) F-O=Guest 1-10
// Row 0: standalone A + merged B:O (col A freeze requires no A merge in any header row)
// Row 1: standalone A + stats cells
// Row 2: headers, Rows 3-10: 8 tables

const TABLES = [
  [1, 'Enchanted Evening', 'Round',       8, ['Emma Thompson','James Wilson','Olivia Chen','Noah Martinez','','','','','','']],
  [2, 'Blossom',           'Round',       8, ['Sophia Kim','Liam Johnson','Isabella Williams','Mason Davis','','','','','','']],
  [3, 'Champagne Toast',   'Round',       6, ['Ava Rodriguez','Ethan Lee','Mia Anderson','Alexander Taylor','','','','','','']],
  [4, 'Starlight',         'Rectangular', 8, ['Charlotte Brown','Amelia Miller','Daniel Moore','Harper Jackson','','','','','','']],
  [5, 'Rosebud',           'Round',       6, ['Lucas White','Henry Thomas','Abigail Martin','Emily Clark','','','','','','']],
  [6, 'Golden Hour',       'Round',       6, ['Jack Lewis','Elizabeth Robinson','Sofia Hall','','','','','','','']],
  [7, 'Sunset',            'Round',       6, ['Aiden Young','Michael Wright','','','','','','','','']],
  [8, 'Moonrise',          'Rectangular', 8, ['','','','','','','','','','']],
];

function fmtCell(bg, textColor, bold, size, align) {
  return {
    backgroundColor: hex(bg),
    textFormat: { foregroundColor: hex(textColor), bold: !!bold, fontSize: size || 10 },
    horizontalAlignment: align || 'LEFT',
    verticalAlignment: 'MIDDLE',
  };
}

(async () => {
  const reqs = [];

  // ── Row 0: standalone A + merged B:O ──
  reqs.push({ repeatCell: {
    range: gridRange(SEA, 0, 1, 0, 1),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 16, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ mergeCells: { range: gridRange(SEA, 0, 1, 1, 15), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(SEA, 0, 1, 1, 15),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 15, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // ── Row 1: standalone A + stats ──
  reqs.push({ repeatCell: {
    range: gridRange(SEA, 1, 2, 0, 1),
    cell: { userEnteredFormat: fmtCell(C.warmSand, C.deepPlum, true, 9, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(SEA, 1, 2, 1, 15),
    cell: { userEnteredFormat: fmtCell(C.warmSand, C.deepPlum, true, 9, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // ── Row 2: headers ──
  reqs.push({ repeatCell: {
    range: gridRange(SEA, 2, 3, 0, 15),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 10, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // ── Data rows 3-10 ──
  for (let r = 0; r < 8; r++) {
    const bg = r % 2 === 0 ? C.ivoryCream : C.altRow;
    reqs.push({ repeatCell: {
      range: gridRange(SEA, r + 3, r + 4, 0, 15),
      cell: { userEnteredFormat: fmtCell(bg, C.bodyText, false, 10, 'LEFT') },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    // Center cols A-E (indices 0-4)
    reqs.push({ repeatCell: {
      range: gridRange(SEA, r + 3, r + 4, 0, 5),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
  }

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SEA, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SEA, dimension: 'ROWS', startIndex: 1, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Col widths: A=60 B=160 C=120 D=80 E=80 F-O=130 each
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SEA, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 60 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SEA, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 160 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SEA, dimension: 'COLUMNS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 120 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SEA, dimension: 'COLUMNS', startIndex: 3, endIndex: 5 },
    properties: { pixelSize: 80 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SEA, dimension: 'COLUMNS', startIndex: 5, endIndex: 15 },
    properties: { pixelSize: 130 }, fields: 'pixelSize',
  }});

  // Borders
  const bs = { style: 'SOLID', color: hex(C.border) };
  reqs.push({ updateBorders: {
    range: gridRange(SEA, 2, 11, 0, 15),
    top: bs, bottom: bs, left: bs, right: bs, innerHorizontal: bs, innerVertical: bs,
  }});

  await batchUpdate(id, reqs, 'seating-format');

  // Values
  const data = [];
  data.push({ range: `'🪑 Seating Chart'!A1`, values: [['🪑']] });
  data.push({ range: `'🪑 Seating Chart'!B1`, values: [['SEATING CHART']] });

  data.push({ range: `'🪑 Seating Chart'!A2`, values: [['STATS']] });
  data.push({
    range: `'🪑 Seating Chart'!B2`,
    values: [[`=IFERROR("Tables: "&COUNTA(B4:B200)&"  |  Guests Seated: "&COUNTA(F4:O200)&"  |  Available Seats: "&(SUM(D4:D200)-COUNTA(F4:O200)),"")`]],
  });

  data.push({
    range: `'🪑 Seating Chart'!A3:O3`,
    values: [['Table #','Table Name','Shape','Capacity','Assigned','Guest 1','Guest 2','Guest 3','Guest 4','Guest 5','Guest 6','Guest 7','Guest 8','Guest 9','Guest 10']],
  });

  const rows = TABLES.map((t, i) => {
    const row = i + 4;
    const assigned = `=COUNTA(F${row}:O${row})`;
    return [t[0], t[1], t[2], t[3], assigned, ...t[4]];
  });
  data.push({ range: `'🪑 Seating Chart'!A4:O11`, values: rows });

  await valuesBatchUpdate(id, data, 'seating-values');
  console.log('Seating Chart complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
