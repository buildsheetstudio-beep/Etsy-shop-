'use strict';
const { C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CHK = sheetMap['✅ Party Checklist'];

// Cols A:I (9 cols, indices 0-8)
// A=# B=Task C=Category D=Priority E=Status F=Owner G=Due Date H=Done I=Notes
// Row 0: full merge, Row 1: stats/progress, Row 2: headers, Rows 3-22: 20 tasks

const TASKS = [
  ['Book and confirm venue',                 'Venue',       'High',   'Complete',     'The Couple',  '2025-09-01', 'TRUE',  'Grand Ballroom Events — contract signed'],
  ['Finalize guest list (30 guests)',        'Invitations', 'High',   'Complete',     'The Couple',  '2025-09-15', 'TRUE',  'Final count: 30 + 6 plus ones'],
  ['Send invitations (4 wks before)',       'Invitations', 'High',   'Complete',     'Partner 1',   '2025-11-01', 'TRUE',  'Sent Oct 28 via post & email'],
  ['Collect RSVPs and finalize headcount',   'Invitations', 'High',   'In Progress',  'Partner 2',   '2025-11-20', 'FALSE', '26/30 received'],
  ['Book caterer and finalize menu',         'Catering',    'High',   'Complete',     'The Couple',  '2025-10-01', 'TRUE',  'Delicious Bites Co — tasting Nov 10'],
  ['Confirm dietary requirements',           'Catering',    'High',   'In Progress',  'Partner 1',   '2025-11-25', 'FALSE', 'Collecting from RSVPs'],
  ['Book florist and choose arrangements',   'Decor',       'High',   'Complete',     'Partner 2',   '2025-10-15', 'TRUE',  'Bloom & Blossom — roses & eucalyptus'],
  ['Order engagement cake',                  'Catering',    'High',   'In Progress',  'The Couple',  '2025-11-10', 'FALSE', 'Sweet Endings — confirm by Nov 10'],
  ['Book photographer',                      'Day-Of',      'High',   'Complete',     'The Couple',  '2025-09-20', 'TRUE',  'Moments Photography — 4 hr package'],
  ['Book videographer',                      'Day-Of',      'High',   'Complete',     'The Couple',  '2025-09-25', 'TRUE',  'CineMemories — highlight reel'],
  ['Book DJ / music',                        'Day-Of',      'High',   'Complete',     'Partner 1',   '2025-10-05', 'TRUE',  'SoundWave DJ — 5 hr set + MC'],
  ['Create seating plan',                    'Venue',       'Medium', 'In Progress',  'Partner 2',   '2025-11-22', 'FALSE', 'Draft done, waiting on final RSVPs'],
  ['Confirm all vendors (1 week before)',   'Venue',       'High',   'Not Started',  'The Couple',  '2025-12-01', 'FALSE', 'Call each vendor to reconfirm'],
  ['Hair & makeup trial',                    'Beauty',      'Medium', 'Not Started',  'Partner 2',   '2025-11-28', 'FALSE', 'Glam Studio — 2 hr trial'],
  ['Write vows / couple speech',             'Day-Of',      'High',   'Not Started',  'The Couple',  '2025-11-30', 'FALSE', 'Keep to 2-3 min each'],
  ['Prepare toasters / speakers',           'Day-Of',      'Medium', 'In Progress',  'Best Man',    '2025-11-25', 'FALSE', 'Jake & Rachel working on speeches'],
  ['Purchase/prepare party favors',          'Decor',       'Medium', 'Delegated',    'Partner 1',   '2025-11-15', 'FALSE', 'Candle favors — Etsy order placed'],
  ['Create & print place cards',             'Decor',       'Medium', 'Not Started',  'Partner 2',   '2025-12-01', 'FALSE', 'Paper & Craft Studio calligraphy'],
  ['Prepare photo booth',                    'Day-Of',      'Medium', 'Not Started',  'Best Man',    '2025-12-05', 'FALSE', 'Props + backdrop + polaroid camera'],
  ['Day-of emergency kit',                   'Day-Of',      'Low',    'Not Started',  'MOH',         '2025-12-07', 'FALSE', 'Safety pins, stain remover, lipstick, mints'],
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
  reqs.push({ mergeCells: { range: gridRange(CHK, 0, 1, 0, 9), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(CHK, 0, 1, 0, 9),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 15, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Row 1: stats + progress bar
  reqs.push({ mergeCells: { range: gridRange(CHK, 1, 2, 0, 5), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(CHK, 1, 2, 0, 5),
    cell: { userEnteredFormat: fmtCell(C.sageGreen, C.white, true, 10, 'LEFT') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ mergeCells: { range: gridRange(CHK, 1, 2, 5, 9), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(CHK, 1, 2, 5, 9),
    cell: { userEnteredFormat: fmtCell(C.sageGreen, C.white, true, 10, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Row 2: headers
  reqs.push({ repeatCell: {
    range: gridRange(CHK, 2, 3, 0, 9),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 10, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Data rows
  for (let r = 0; r < 20; r++) {
    const bg = r % 2 === 0 ? C.ivoryCream : C.altRow;
    reqs.push({ repeatCell: {
      range: gridRange(CHK, r + 3, r + 4, 0, 9),
      cell: { userEnteredFormat: fmtCell(bg, C.bodyText, false, 10, 'LEFT') },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    // Center A, C, D, E, H
    [0, 2, 3, 4, 7].forEach(ci => {
      reqs.push({ repeatCell: {
        range: gridRange(CHK, r + 3, r + 4, ci, ci + 1),
        cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
        fields: 'userEnteredFormat.horizontalAlignment',
      }});
    });
  }

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CHK, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CHK, dimension: 'ROWS', startIndex: 1, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Col widths: A=40 B=250 C=110 D=80 E=110 F=140 G=100 H=60 I=220
  const cw = [40, 250, 110, 80, 110, 140, 100, 60, 220];
  cw.forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: CHK, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Borders
  const bs = { style: 'SOLID', color: hex(C.border) };
  reqs.push({ updateBorders: {
    range: gridRange(CHK, 2, 23, 0, 9),
    top: bs, bottom: bs, left: bs, right: bs, innerHorizontal: bs, innerVertical: bs,
  }});

  await batchUpdate(id, reqs, 'checklist-format');

  // Values
  const data = [];
  data.push({ range: `'✅ Party Checklist'!A1`, values: [['✅ PARTY CHECKLIST']] });
  data.push({
    range: `'✅ Party Checklist'!A2`,
    values: [[`=IFERROR("Tasks: "&COUNTA(B4:B200)&"  |  Complete: "&COUNTIF(H4:H200,TRUE)&"  |  In Progress: "&COUNTIF(E4:E200,"In Progress")&"  |  Not Started: "&COUNTIF(E4:E200,"Not Started"),"")`]],
  });
  data.push({
    range: `'✅ Party Checklist'!F2`,
    values: [[`=IFERROR("Progress: "&TEXT(COUNTIF(H4:H200,TRUE)/COUNTA(B4:B200),"0%"),"")`]],
  });
  data.push({
    range: `'✅ Party Checklist'!A3:I3`,
    values: [['#','Task','Category','Priority','Status','Owner','Due Date','Done?','Notes']],
  });

  const rows = TASKS.map((t, i) => [i + 1, t[0], t[1], t[2], t[3], t[4], t[5], t[6], t[7]]);
  data.push({ range: `'✅ Party Checklist'!A4:I23`, values: rows });

  await valuesBatchUpdate(id, data, 'checklist-values');
  console.log('Party Checklist complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
