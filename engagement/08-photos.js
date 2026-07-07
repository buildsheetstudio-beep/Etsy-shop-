'use strict';
const { C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const PHO = sheetMap['📸 Photo & Shot List'];

// Cols A:H (8 cols, indices 0-7)
// A=# B=Shot Description C=Type D=Priority E=Location F=People Involved G=Status H=Notes
// Row 0: full merge, Row 1: stats, Row 2: headers, Rows 3-17: 15 shots

const SHOTS = [
  ['Couple portrait — golden hour outdoors',      'Portrait',   'Must Have',     'Venue Garden',    'The Couple',               'Must Have',   'Shoot at 4:00 PM for best light'],
  ['Full group — all 30 guests',                  'Group',      'Must Have',     'Main Hall',       'All Guests',               'Must Have',   'After dinner, before dancing'],
  ['Wedding party group shot',                    'Group',      'Must Have',     'Main Hall',       'Best Man, MOH, Couple',    'Must Have',   'While everyone is still tidy'],
  ['Partner 1 family portrait',                   'Group',      'Must Have',     'Main Hall',       'Partner 1 Family',         'Must Have',   'Before dinner service'],
  ['Partner 2 family portrait',                   'Group',      'Must Have',     'Main Hall',       'Partner 2 Family',         'Must Have',   'Before dinner service'],
  ['Best friends candid at welcome drinks',       'Candid',     'Must Have',     'Cocktail Lounge', 'Close Friends',            'Must Have',   'Natural, laughing shots'],
  ['Ring detail close-up',                        'Detail',     'Must Have',     'Table',           '',                         'Must Have',   'Both rings on velvet/flowers'],
  ['Couple candid during reception',              'Candid',     'Must Have',     'Main Hall',       'The Couple',               'Must Have',   'Unposed, genuine moments'],
  ['Toast moment — glasses raised',               'Ceremony',   'Must Have',     'Main Hall',       'All Guests',               'Must Have',   'Wide shot of room during toast'],
  ['Cake cutting',                                'Reception',  'Must Have',     'Main Hall',       'The Couple',               'Must Have',   'Classic + candid versions'],
  ['First dance',                                 'Reception',  'Must Have',     'Dance Floor',     'The Couple',               'Must Have',   'Wide + detail shots'],
  ['Venue exterior / facade',                     'Venue',      'Nice to Have',  'Outside',         '',                         'Nice to Have','Arrival & detail architecture'],
  ['Table décor detail shots',                    'Detail',     'Nice to Have',  'Main Hall',       '',                         'Nice to Have','Centerpieces, place settings'],
  ['Floral arrangement close-ups',                'Detail',     'Nice to Have',  'All Tables',      '',                         'Nice to Have','Each arrangement'],
  ['Guest candids throughout evening',            'Candid',     'Nice to Have',  'All Areas',       'Various Guests',           'Nice to Have','Laughter, conversation, dancing'],
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

  // Row 0: full-width title
  reqs.push({ mergeCells: { range: gridRange(PHO, 0, 1, 0, 8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(PHO, 0, 1, 0, 8),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 15, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Row 1: stats bar
  reqs.push({ mergeCells: { range: gridRange(PHO, 1, 2, 0, 8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(PHO, 1, 2, 0, 8),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 10, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Row 2: headers
  reqs.push({ repeatCell: {
    range: gridRange(PHO, 2, 3, 0, 8),
    cell: { userEnteredFormat: fmtCell(C.mutedRust, C.white, true, 10, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Data rows
  for (let r = 0; r < 15; r++) {
    const bg = r % 2 === 0 ? C.ivoryCream : C.altRow;
    reqs.push({ repeatCell: {
      range: gridRange(PHO, r + 3, r + 4, 0, 8),
      cell: { userEnteredFormat: fmtCell(bg, C.bodyText, false, 10, 'LEFT') },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    // Center A, C, D, G
    [0, 2, 3, 6].forEach(ci => {
      reqs.push({ repeatCell: {
        range: gridRange(PHO, r + 3, r + 4, ci, ci + 1),
        cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
        fields: 'userEnteredFormat.horizontalAlignment',
      }});
    });
  }

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: PHO, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: PHO, dimension: 'ROWS', startIndex: 1, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Col widths: A=40 B=280 C=110 D=110 E=150 F=200 G=110 H=200
  const cw = [40, 280, 110, 110, 150, 200, 110, 200];
  cw.forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: PHO, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Borders
  const bs = { style: 'SOLID', color: hex(C.border) };
  reqs.push({ updateBorders: {
    range: gridRange(PHO, 2, 18, 0, 8),
    top: bs, bottom: bs, left: bs, right: bs, innerHorizontal: bs, innerVertical: bs,
  }});

  await batchUpdate(id, reqs, 'photos-format');

  // Values
  const data = [];
  data.push({ range: `'📸 Photo & Shot List'!A1`, values: [['📸 PHOTO & SHOT LIST']] });
  data.push({
    range: `'📸 Photo & Shot List'!A2`,
    values: [[`=IFERROR("Total Shots: "&COUNTA(B4:B200)&"  |  Must Have: "&COUNTIF(D4:D200,"Must Have")&"  |  Nice to Have: "&COUNTIF(D4:D200,"Nice to Have")&"  |  Completed: "&COUNTIF(G4:G200,"Completed"),"")`]],
  });
  data.push({
    range: `'📸 Photo & Shot List'!A3:H3`,
    values: [['#','Shot Description','Type','Priority','Location','People Involved','Status','Notes']],
  });

  const rows = SHOTS.map((s, i) => [i + 1, s[0], s[1], s[2], s[3], s[4], s[5], s[6]]);
  data.push({ range: `'📸 Photo & Shot List'!A4:H18`, values: rows });

  await valuesBatchUpdate(id, data, 'photos-values');
  console.log('Photo & Shot List complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
