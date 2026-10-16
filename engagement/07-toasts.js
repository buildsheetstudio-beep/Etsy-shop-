'use strict';
const { C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TST = sheetMap['🥂 Toast & Speech Planner'];

// Cols A:H (8 cols, indices 0-7)
// A=# B=Speaker C=Relationship D=Speech Title E=Status F=Duration(min) G=Key Points H=Notes
// Row 0: full merge, Row 1: stats, Row 2: headers, Rows 3-7: 5 toasts

const TOASTS = [
  ['Jake Morrison',    'Best Man',            'A Brotherhood Celebrated',    'Speech Ready',  5, 'Friendship story; how they met; inside jokes; heartfelt wishes', 'Has microphone reserved'],
  ['Rachel Green',     'Maid of Honor',       'Sisters by Choice',           'Speech Ready',  4, 'Years of friendship; growth together; love story witness', 'Will cry — have tissues ready'],
  ['Robert Thompson',  'Father of Partner 1', 'A Parent\'s Greatest Wish',   'Draft Written', 3, 'Childhood memories; proud moment; welcome to family', 'First speech of the evening'],
  ['Maria Sanchez',    'Mother of Partner 2', 'Welcome to Our Family',       'Confirmed',     3, 'Family values; seeing child happy; welcoming partner 1', 'Bilingual — Spanish & English'],
  ['Emma & David',     'The Couple',          'To Us & Everyone Here',       'Pending',       2, 'Thank guests; acknowledge family; look to the future together', 'Closing toast, raise glasses'],
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
  reqs.push({ mergeCells: { range: gridRange(TST, 0, 1, 0, 8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(TST, 0, 1, 0, 8),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 15, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Row 1: stats bar
  reqs.push({ mergeCells: { range: gridRange(TST, 1, 2, 0, 8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(TST, 1, 2, 0, 8),
    cell: { userEnteredFormat: fmtCell(C.mutedRust, C.white, true, 10, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Row 2: headers
  reqs.push({ repeatCell: {
    range: gridRange(TST, 2, 3, 0, 8),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 10, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Data rows
  for (let r = 0; r < 5; r++) {
    const bg = r % 2 === 0 ? C.ivoryCream : C.altRow;
    reqs.push({ repeatCell: {
      range: gridRange(TST, r + 3, r + 4, 0, 8),
      cell: { userEnteredFormat: fmtCell(bg, C.bodyText, false, 10, 'LEFT') },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    // Center A, F
    reqs.push({ repeatCell: {
      range: gridRange(TST, r + 3, r + 4, 0, 1),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(TST, r + 3, r + 4, 5, 6),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
  }

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TST, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TST, dimension: 'ROWS', startIndex: 1, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TST, dimension: 'ROWS', startIndex: 3, endIndex: 8 },
    properties: { pixelSize: 60 }, fields: 'pixelSize',
  }});

  // Col widths: A=40 B=170 C=160 D=200 E=130 F=90 G=280 H=200
  const cw = [40, 170, 160, 200, 130, 90, 280, 200];
  cw.forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: TST, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Borders
  const bs = { style: 'SOLID', color: hex(C.border) };
  reqs.push({ updateBorders: {
    range: gridRange(TST, 2, 8, 0, 8),
    top: bs, bottom: bs, left: bs, right: bs, innerHorizontal: bs, innerVertical: bs,
  }});

  // Wrap text for G (Key Points), H (Notes)
  reqs.push({ repeatCell: {
    range: gridRange(TST, 3, 8, 6, 8),
    cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat.wrapStrategy',
  }});

  await batchUpdate(id, reqs, 'toasts-format');

  // Values
  const data = [];
  data.push({ range: `'🥂 Toast & Speech Planner'!A1`, values: [['🥂 TOAST & SPEECH PLANNER']] });
  data.push({
    range: `'🥂 Toast & Speech Planner'!A2`,
    values: [[`=IFERROR("Total Speeches: "&COUNTA(B4:B200)&"  |  Ready: "&COUNTIF(E4:E200,"Speech Ready")&"  |  Confirmed: "&COUNTIF(E4:E200,"Confirmed")&"  |  Total Duration: "&SUM(F4:F200)&" min","")`]],
  });
  data.push({
    range: `'🥂 Toast & Speech Planner'!A3:H3`,
    values: [['#','Speaker','Relationship','Speech Title','Status','Duration (min)','Key Points / Talking Points','Notes']],
  });

  const rows = TOASTS.map((t, i) => [i + 1, t[0], t[1], t[2], t[3], t[4], t[5], t[6]]);
  data.push({ range: `'🥂 Toast & Speech Planner'!A4:H8`, values: rows });

  await valuesBatchUpdate(id, data, 'toasts-values');
  console.log('Toast & Speech Planner complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
