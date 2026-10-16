'use strict';
const { C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TML = sheetMap['⏰ Day-Of Timeline'];

// Cols A:J (10 cols, indices 0-9)
// A=# B=Time C=Duration(min) D=Activity E=Type F=Responsible G=Vendor H=Location I=Notes J=Done
// Row 0: full merge, Row 1: stats, Row 2: headers, Rows 3-17: 15 blocks

const BLOCKS = [
  ['2:00 PM', 120, 'Venue Setup & Decoration',    'Setup',          'Venue Staff',     'Grand Ballroom Events', 'Main Hall',       'Set tables, linens, centerpieces'],
  ['3:30 PM',  60, 'Florals & Final Touches',      'Setup',          'Couple / Planner','Bloom & Blossom',       'Main Hall',       'Flower arrangements in place'],
  ['4:00 PM',  30, 'Couple Arrival & Prep Photos', 'Photography',    'Photographer',    'Moments Photography',   'Venue Garden',    'Golden hour portraits'],
  ['4:30 PM',  30, 'Early Guest Arrivals',          'Reception',      'Greeter',         '',                      'Venue Entrance',  'Welcome guests, seat them'],
  ['5:00 PM',  60, 'Champagne Welcome Reception',  'Reception',      'Caterer',         'Delicious Bites Co',    'Cocktail Lounge', 'Canapes & welcome drinks'],
  ['5:30 PM',  30, 'Canapés & Mingling',            'Catering',       'Caterer',         'Delicious Bites Co',    'Cocktail Lounge', 'Passed apps + soft drinks'],
  ['6:00 PM',  15, 'Guest Seating & Welcome',       'Ceremony',       'MC / DJ',         'SoundWave DJ',          'Main Hall',       'DJ announces dinner'],
  ['6:15 PM',  20, 'Welcome & Blessing Speech',     'Ceremony',       'Host Family',     '',                      'Main Hall',       'Parent opening remarks'],
  ['6:35 PM',  15, 'First Toast',                   'Ceremony',       'Best Man',        '',                      'Main Hall',       'Jake Morrison toasts'],
  ['6:50 PM',  10, 'Couple\'s Toast',               'Ceremony',       'The Couple',      '',                      'Main Hall',       'Partners speak'],
  ['7:00 PM',  45, 'Dinner Service – Main Course',  'Catering',       'Caterer',         'Delicious Bites Co',    'Main Hall',       'Plated 3-course dinner'],
  ['7:45 PM',  30, 'Toasts & Speeches',             'Ceremony',       'MC',              'SoundWave DJ',          'Main Hall',       'MOH + parent speeches'],
  ['8:15 PM',  20, 'Cake Cutting & Desserts',       'Catering',       'Caterer + Baker', 'Sweet Endings Bakery',  'Main Hall',       'Dessert table opens'],
  ['8:35 PM',  90, 'Open Dancing & Music',           'Entertainment',  'DJ',              'SoundWave DJ',          'Dance Floor',     'Open dance floor, music set'],
  ['10:05 PM', 25, 'Farewell & Cleanup',             'Cleanup',        'Couple / Venue',  'Grand Ballroom Events', 'All Areas',       'Guest send-off, venue reset'],
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
  reqs.push({ mergeCells: { range: gridRange(TML, 0, 1, 0, 10), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(TML, 0, 1, 0, 10),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 15, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Row 1: stats bar
  reqs.push({ mergeCells: { range: gridRange(TML, 1, 2, 0, 10), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(TML, 1, 2, 0, 10),
    cell: { userEnteredFormat: fmtCell(C.sageGreen, C.white, true, 10, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Row 2: headers
  reqs.push({ repeatCell: {
    range: gridRange(TML, 2, 3, 0, 10),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 10, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Data rows
  for (let r = 0; r < 15; r++) {
    const bg = r % 2 === 0 ? C.ivoryCream : C.altRow;
    reqs.push({ repeatCell: {
      range: gridRange(TML, r + 3, r + 4, 0, 10),
      cell: { userEnteredFormat: fmtCell(bg, C.bodyText, false, 10, 'LEFT') },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    // Center: A(#), B(time), C(dur), J(done)
    reqs.push({ repeatCell: {
      range: gridRange(TML, r + 3, r + 4, 0, 3),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(TML, r + 3, r + 4, 9, 10),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
  }

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TML, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TML, dimension: 'ROWS', startIndex: 1, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Col widths: A=40 B=90 C=90 D=230 E=130 F=160 G=160 H=140 I=220 J=60
  const cw = [40, 90, 90, 230, 130, 160, 160, 140, 220, 60];
  cw.forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: TML, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Borders
  const bs = { style: 'SOLID', color: hex(C.border) };
  reqs.push({ updateBorders: {
    range: gridRange(TML, 2, 18, 0, 10),
    top: bs, bottom: bs, left: bs, right: bs, innerHorizontal: bs, innerVertical: bs,
  }});

  await batchUpdate(id, reqs, 'timeline-format');

  // Values
  const data = [];
  data.push({ range: `'⏰ Day-Of Timeline'!A1`, values: [['⏰ DAY-OF TIMELINE']] });
  data.push({
    range: `'⏰ Day-Of Timeline'!A2`,
    values: [[`=IFERROR("Total Activities: "&COUNTA(D4:D200)&"  |  Completed: "&COUNTIF(J4:J200,TRUE)&"  |  Remaining: "&(COUNTA(D4:D200)-COUNTIF(J4:J200,TRUE))&"  |  Total Duration: "&TEXT(SUM(C4:C200)/60,"0.0")&" hrs","")`]],
  });
  data.push({
    range: `'⏰ Day-Of Timeline'!A3:J3`,
    values: [['#','Time','Dur (min)','Activity / Task','Type','Responsible Person','Vendor Involved','Location','Notes','Done?']],
  });

  const rows = BLOCKS.map((b, i) => [i + 1, b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7], 'FALSE']);
  data.push({ range: `'⏰ Day-Of Timeline'!A4:J18`, values: rows });

  await valuesBatchUpdate(id, data, 'timeline-values');
  console.log('Day-Of Timeline complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
