'use strict';
const { C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BUD = sheetMap['💰 Budget & Vendor Tracker'];

// Cols A:M (13 cols, indices 0-12)
// A=# B=Vendor/Item C=Category D=Budgeted E=Actual F=Variance(D-E) G=Payment Status
// H=Contact Name I=Phone/Email J=Contract K=Deposit Due L=Final Due M=Notes
// Row 0: merge A:M (full-width ok, no col A freeze needed)
// Row 1: stats/summary bar
// Row 2: headers, Row 3-12: 10 vendors

const VENDORS = [
  ['Grand Ballroom Events', 'Venue',          3000, 3000,  'Fully Paid',   'Mark Stevens',    '555-2001', 'TRUE',  '2025-11-01','2025-11-01','Historic venue, 5pm setup'],
  ['Delicious Bites Co',    'Catering',        4500, 4200,  'Deposit Paid', 'Angela Rivera',   '555-2002', 'TRUE',  '2025-10-15','2025-12-01','Menu tasting Nov 10'],
  ['Moments Photography',   'Photography',     2000, 1800,  'Deposit Paid', 'Chris Laurent',   '555-2003', 'TRUE',  '2025-10-01','2025-12-05','4hr coverage + editing'],
  ['CineMemories Studio',   'Videography',     1500, 1500,  'Fully Paid',   'Dana Park',       '555-2004', 'TRUE',  '2025-10-01','2025-10-01','Highlight reel incl.'],
  ['Bloom & Blossom',       'Florals',          800,  750,  'Deposit Paid', 'Sophie Green',    '555-2005', 'TRUE',  '2025-11-15','2025-12-01','Centerpieces + bouquet'],
  ['SoundWave DJ',          'Music/DJ',        1000,  900,  'Not Paid',     'Jake Morrison',   '555-2006', 'FALSE', '2025-11-20','2025-12-10','5hr set, MC services'],
  ['Glam Studio',           'Hair & Makeup',    600,    0,  'Not Paid',     'Rachel Torres',   '555-2007', 'FALSE', '2025-12-01','2025-12-10','2 stylists, 3 hr setup'],
  ['Sweet Endings Bakery',  'Cake/Desserts',    400,  200,  'Deposit Paid', 'Lily Chen',       '555-2008', 'TRUE',  '2025-11-01','2025-12-08','3-tier cake + petit fours'],
  ['Paper & Craft Studio',  'Invitations',      200,  185,  'Fully Paid',   'Tom Bradley',     '555-2009', 'TRUE',  '2025-09-01','2025-09-15','50 invites, 60 thank-yous'],
  ['Party Supplies Plus',   'Other',            300,  150,  'Deposit Paid', 'Lisa Hammond',    '555-2010', 'TRUE',  '2025-11-25','2025-12-05','Balloons, linens, candles'],
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

  // ── Row 0: Full-width title merge ──
  reqs.push({ mergeCells: { range: gridRange(BUD, 0, 1, 0, 13), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(BUD, 0, 1, 0, 13),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 15, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // ── Row 1: Stats bar (full merge) ──
  reqs.push({ mergeCells: { range: gridRange(BUD, 1, 2, 0, 13), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(BUD, 1, 2, 0, 13),
    cell: { userEnteredFormat: fmtCell(C.mutedGold, C.deepPlum, true, 10, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // ── Row 2: Headers ──
  reqs.push({ repeatCell: {
    range: gridRange(BUD, 2, 3, 0, 13),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 10, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // ── Data rows ──
  for (let r = 0; r < 10; r++) {
    const bg = r % 2 === 0 ? C.ivoryCream : C.altRow;
    reqs.push({ repeatCell: {
      range: gridRange(BUD, r + 3, r + 4, 0, 13),
      cell: { userEnteredFormat: fmtCell(bg, C.bodyText, false, 10, 'LEFT') },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    // Right-align numeric cols D, E, F (indices 3,4,5)
    reqs.push({ repeatCell: {
      range: gridRange(BUD, r + 3, r + 4, 3, 6),
      cell: { userEnteredFormat: { horizontalAlignment: 'RIGHT' } },
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
    // Center J (contract checkbox)
    reqs.push({ repeatCell: {
      range: gridRange(BUD, r + 3, r + 4, 9, 10),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
  }

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BUD, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BUD, dimension: 'ROWS', startIndex: 1, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Column widths: A=40 B=200 C=130 D=100 E=100 F=110 G=120 H=160 I=160 J=90 K=110 L=110 M=180
  const cw = [40, 200, 130, 100, 100, 110, 120, 160, 160, 90, 110, 110, 180];
  cw.forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: BUD, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Borders
  const bs = { style: 'SOLID', color: hex(C.border) };
  reqs.push({ updateBorders: {
    range: gridRange(BUD, 2, 13, 0, 13),
    top: bs, bottom: bs, left: bs, right: bs, innerHorizontal: bs, innerVertical: bs,
  }});

  // Currency format for D, E, F cols
  const currencyFmt = { type: 'CURRENCY', pattern: '"$"#,##0.00' };
  for (let r = 3; r < 13; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(BUD, r, r + 1, 3, 6),
      cell: { userEnteredFormat: { numberFormat: currencyFmt } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  }

  await batchUpdate(id, reqs, 'budget-format');

  // ── Values ──
  const data = [];
  data.push({ range: `'💰 Budget & Vendor Tracker'!A1`, values: [['💰 BUDGET & VENDOR TRACKER']] });
  data.push({
    range: `'💰 Budget & Vendor Tracker'!A2`,
    values: [[`=IFERROR("Total Budget: $"&TEXT(SUM(D4:D200),"#,##0")&"  |  Total Spent: $"&TEXT(SUM(E4:E200),"#,##0")&"  |  Remaining: $"&TEXT(SUM(D4:D200)-SUM(E4:E200),"#,##0")&"  |  Vendors Fully Paid: "&COUNTIF(G4:G200,"Fully Paid")&" of "&COUNTA(B4:B200),"")`]],
  });
  data.push({
    range: `'💰 Budget & Vendor Tracker'!A3:M3`,
    values: [['#','Vendor / Item','Category','Budgeted ($)','Actual ($)','Variance ($)','Payment Status','Contact Name','Phone / Email','Contract?','Deposit Due','Final Due','Notes']],
  });

  const rows = VENDORS.map((v, i) => {
    const row = i + 4;
    return [i + 1, v[0], v[1], v[2], v[3], `=D${row}-E${row}`, v[4], v[5], v[6], v[7], v[8], v[9], v[10]];
  });
  data.push({ range: `'💰 Budget & Vendor Tracker'!A4:M13`, values: rows });

  await valuesBatchUpdate(id, data, 'budget-values');
  console.log('Budget & Vendor Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
