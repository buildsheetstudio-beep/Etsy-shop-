'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BN = sheetMap['👶 Baby Names'];

// 11 cols, 55 rows
// Row 0: Title (B1:K1 — A1 blank for col A freeze)
// Row 1: Col headers
// Row 2+: name entries

(async () => {
  const reqs = [];

  // Leave A1 blank, merge B1:K1
  reqs.push({ mergeCells: { range: gridRange(BN, 0, 1, 1, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(BN, 0, 1, 0, 11),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: BN, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });

  // Summary rows
  reqs.push({ repeatCell: {
    range: gridRange(BN, 1, 2, 0, 11),
    cell: { userEnteredFormat: { backgroundColor: hex(C.whisperLav), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(BN, 2, 3, 0, 11),
    cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: BN, dimension: 'ROWS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Header row 3
  reqs.push({ repeatCell: {
    range: gridRange(BN, 3, 4, 0, 11),
    cell: { userEnteredFormat: { backgroundColor: hex(C.softLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: BN, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });

  // Alt rows 4-54
  for (let r = 4; r < 55; r += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(BN, r, r+1, 0, 11),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderMist) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }

  // Wrap
  reqs.push({ repeatCell: { range: gridRange(BN, 4, 55, 0, 11), cell: { userEnteredFormat: { wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(wrapStrategy,verticalAlignment)' }});

  // Formula col (K = combined score) — pale lavender
  reqs.push({ repeatCell: { range: gridRange(BN, 4, 55, 10, 11), cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)' }});

  // Col widths
  const colW = [50, 120, 80, 90, 80, 80, 80, 130, 130, 180, 90];
  colW.forEach((w, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: BN, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: w }, fields: 'pixelSize',
  }}));

  const border = { style: 'SOLID', color: hex(C.lavenderBorder) };
  reqs.push({ updateBorders: { range: gridRange(BN, 3, 55, 0, 11), innerHorizontal: border, innerVertical: border, top: border, bottom: border, left: border, right: border }});

  await batchUpdate(id, reqs, 'bn-format');

  const vdata = [];
  vdata.push({ range: `'👶 Baby Names'!B1`, values: [['👶 Baby Names']] });

  vdata.push({ range: `'👶 Baby Names'!A2:K2`, values: [['Total Names','Girl Names','Boy/Neutral Names','Top Score','Favourite','Partner Fav','','','','','']] });
  vdata.push({ range: `'👶 Baby Names'!A3:K3`, values: [[
    '=COUNTA(B5:B54)',
    '=COUNTIF(C5:C54,"Girl")',
    '=COUNTIF(C5:C54,"Boy")+COUNTIF(C5:C54,"Neutral")',
    '=IFERROR(MAX(K5:K54),"—")',
    '=IFERROR(INDEX(B5:B54,MATCH(MAX(K5:K54),K5:K54,0)),"—")',
    '=IFERROR(INDEX(B5:B54,MATCH(MAX(H5:H54),H5:H54,0)),"—")',
    '','','','','',
  ]]});

  vdata.push({ range: `'👶 Baby Names'!A4:K4`, values: [[
    '#','Name','Gender','Origin','Meaning (brief)','Vibe','Your Score (1-10)','Partner Score (1-10)','Family Score (1-10)','Notes','Combined Score',
  ]]});

  // 20 names (10 girl, 10 boy/neutral) with ratings
  // Combined score = AVERAGE of non-blank scores
  const names = [
    // Girl names
    [1,'Aurelia','Girl','Latin','Golden, the golden one','Classic',9,8,7,'My absolute favourite — Auri for short','=IFERROR(ROUND(AVERAGE(G5:I5),1),"")'],
    [2,'Celeste','Girl','Latin','Heavenly, of the sky','Classic',8,9,8,'Cele or Cece — so elegant','=IFERROR(ROUND(AVERAGE(G6:I6),1),"")'],
    [3,'Juniper','Girl','Latin','Evergreen shrub','Nature',7,7,6,'Juni — modern but grounded','=IFERROR(ROUND(AVERAGE(G7:I7),1),"")'],
    [4,'Isadora','Girl','Greek','Gift of Isis','Vintage',8,6,7,'Izzy or Dora — strong and romantic','=IFERROR(ROUND(AVERAGE(G8:I8),1),"")'],
    [5,'Elara','Girl','Greek','Bright, shining','Modern',9,9,8,'Short, strong, beautiful','=IFERROR(ROUND(AVERAGE(G9:I9),1),"")'],
    [6,'Seraphina','Girl','Hebrew','Fiery, ardent','Spiritual',7,8,6,'Sera or Fina — angelic','=IFERROR(ROUND(AVERAGE(G10:I10),1),"")'],
    [7,'Rosalind','Girl','Germanic','Gentle horse, lovely','Vintage',6,7,8,'Roz or Lindy — timeless','=IFERROR(ROUND(AVERAGE(G11:I11),1),"")'],
    [8,'Elowen','Girl','Cornish','Elm tree','Nature',8,7,5,'Ellie — ethereal and rare','=IFERROR(ROUND(AVERAGE(G12:I12),1),"")'],
    [9,'Vivienne','Girl','Latin','Life, alive','Classic',7,8,7,'Viv — sophisticated','=IFERROR(ROUND(AVERAGE(G13:I13),1),"")'],
    [10,'Margot','Girl','French/Greek','Pearl','Vintage',9,7,9,'Love the French feel — Margo spelling','=IFERROR(ROUND(AVERAGE(G14:I14),1),"")'],
    // Boy / Neutral names
    [11,'Jasper','Boy','Persian','Treasurer','Nature',8,9,8,'Jazz — warm and strong','=IFERROR(ROUND(AVERAGE(G15:I15),1),"")'],
    [12,'Felix','Boy','Latin','Happy, fortunate','Classic',9,8,9,'Simple, joyful, perfect','=IFERROR(ROUND(AVERAGE(G16:I16),1),"")'],
    [13,'Caspian','Boy','Latin/Literary','Of the Caspian Sea','Whimsical',8,7,6,'Cas — adventurous, Narnia vibes','=IFERROR(ROUND(AVERAGE(G17:I17),1),"")'],
    [14,'Orion','Boy','Greek','Rising in sky, hunter','Nature',7,9,7,'Rion — celestial and strong','=IFERROR(ROUND(AVERAGE(G18:I18),1),"")'],
    [15,'Theodore','Boy','Greek','Gift of God','Classic',9,8,8,'Theo — absolutely classic','=IFERROR(ROUND(AVERAGE(G19:I19),1),"")'],
    [16,'Rowan','Neutral','Gaelic','Little red one, tree','Nature',8,8,8,'Rowan — beautiful either way','=IFERROR(ROUND(AVERAGE(G20:I20),1),"")'],
    [17,'Emmett','Boy','Germanic','Entire, whole','Modern',7,9,7,'Em — strong and friendly','=IFERROR(ROUND(AVERAGE(G21:I21),1),"")'],
    [18,'Soren','Neutral','Scandinavian','Stern','Unique',8,7,5,'Sophisticated Scandinavian feel','=IFERROR(ROUND(AVERAGE(G22:I22),1),"")'],
    [19,'Callum','Boy','Scottish','Dove','Classic',7,8,6,'Cal — Scottish heritage ties','=IFERROR(ROUND(AVERAGE(G23:I23),1),"")'],
    [20,'Everett','Boy','Old English','Wild boar, brave','Modern',8,8,7,'Ev or Rhett — Southern charm','=IFERROR(ROUND(AVERAGE(G24:I24),1),"")'],
  ];
  vdata.push({ range: `'👶 Baby Names'!A5:K24`, values: names });

  await valuesBatchUpdate(id, vdata, 'bn-values');
  console.log('Baby Names complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
