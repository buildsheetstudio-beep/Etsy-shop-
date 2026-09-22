'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const NP = sheetMap['🛏️ Nursery Planner'];

// 10 cols, 55 rows
// Row 0: Title, Row 1: Summary labels, Row 2: Summary values
// Row 3: Section header — Nursery Items, Row 4: Col headers, Row 5-36: ~30 items
// Row 37: Section header — Mood Board / Notes, Row 38-54: Notes section

(async () => {
  const reqs = [];

  // Title
  reqs.push({ mergeCells: { range: gridRange(NP, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(NP, 0, 1, 0, 10),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: NP, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });

  // Summary rows 1-2
  [1,2].forEach((r, i) => {
    reqs.push({ repeatCell: {
      range: gridRange(NP, r, r+1, 0, 10),
      cell: { userEnteredFormat: { backgroundColor: hex(i === 0 ? C.whisperLav : C.paleLavender), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: NP, dimension: 'ROWS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Section headers rows 3 and 37
  [3, 37].forEach(r => {
    reqs.push({ mergeCells: { range: gridRange(NP, r, r+1, 0, 10), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(NP, r, r+1, 0, 10),
      cell: { userEnteredFormat: { backgroundColor: hex(C.softLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'LEFT' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
    }});
    reqs.push({ updateDimensionProperties: { range: { sheetId: NP, dimension: 'ROWS', startIndex: r, endIndex: r+1 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  });

  // Col headers row 4
  reqs.push({ repeatCell: {
    range: gridRange(NP, 4, 5, 0, 10),
    cell: { userEnteredFormat: { backgroundColor: hex(C.taupeGold), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: NP, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });

  // Alt rows 5-36
  for (let r = 5; r < 37; r += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(NP, r, r+1, 0, 10),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderMist) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }

  // Notes section header row 37 col headers row 38
  reqs.push({ repeatCell: {
    range: gridRange(NP, 38, 39, 0, 10),
    cell: { userEnteredFormat: { backgroundColor: hex(C.taupeGold), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: NP, dimension: 'ROWS', startIndex: 38, endIndex: 39 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Notes alt rows 39-54
  for (let r = 39; r < 55; r += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(NP, r, r+1, 0, 10),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderMist) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }

  // Wrap
  reqs.push({ repeatCell: { range: gridRange(NP, 5, 55, 0, 10), cell: { userEnteredFormat: { wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(wrapStrategy,verticalAlignment)' }});

  // Column widths
  const colW = [30, 170, 130, 100, 100, 110, 90, 110, 130, 180];
  colW.forEach((w, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: NP, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: w }, fields: 'pixelSize',
  }}));

  // Currency format for price cols (D=3, E=4)
  const currFmt = { type: 'CURRENCY', pattern: '"$"#,##0.00' };
  [3,4].forEach(c => reqs.push({ repeatCell: { range: gridRange(NP, 5, 37, c, c+1), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields: 'userEnteredFormat.numberFormat' }}));

  const border = { style: 'SOLID', color: hex(C.lavenderBorder) };
  reqs.push({ updateBorders: { range: gridRange(NP, 4, 37, 0, 10), innerHorizontal: border, innerVertical: border, top: border, bottom: border, left: border, right: border }});
  reqs.push({ updateBorders: { range: gridRange(NP, 38, 55, 0, 10), innerHorizontal: border, innerVertical: border, top: border, bottom: border, left: border, right: border }});

  await batchUpdate(id, reqs, 'np-format');

  const vdata = [];
  vdata.push({ range: `'🛏️ Nursery Planner'!A1`, values: [['🛏️ Nursery Planner']] });

  vdata.push({ range: `'🛏️ Nursery Planner'!A2:J2`, values: [['Total Items','Purchased','Still Needed','Total Budget','Total Spent','Remaining','% Complete','','','']] });
  vdata.push({ range: `'🛏️ Nursery Planner'!A3:J3`, values: [[
    '=COUNTA(B6:B37)',
    '=COUNTIF(G6:G37,TRUE)',
    '=A3-B3',
    '=SUM(D6:D37)',
    '=SUM(E6:E37)',
    '=D3-E3',
    '=IFERROR(TEXT(B3/A3,"0%"),"—")',
    '','','',
  ]]});

  vdata.push({ range: `'🛏️ Nursery Planner'!A4`, values: [['  🛏️ Nursery Items Checklist']] });
  vdata.push({ range: `'🛏️ Nursery Planner'!A5:J5`, values: [['#','Item','Category','Budgeted','Actual Cost','Priority','Purchased','Store','Notes','Link / SKU']] });

  const items = [
    [1,'Crib (Babyletto Hudson)','Furniture',400,380,'High',true,'Buy Buy Baby','White finish',''],
    [2,'Crib mattress','Bedding & Textiles',120,110,'High',true,'Amazon','Newton organic',''],
    [3,'Fitted crib sheets x3','Bedding & Textiles',45,42,'High',true,'Target','Muslin blend',''],
    [4,'Waterproof mattress cover x2','Bedding & Textiles',35,32,'High',true,'Amazon','',''],
    [5,'Swaddle blankets x4','Bedding & Textiles',40,38,'High',true,'HALO','SleepSack',''],
    [6,'Glider chair','Furniture',550,550,'High',true,'Pottery Barn','Cream bouclé',''],
    [7,'Glider ottoman','Furniture',200,190,'Medium',true,'Pottery Barn','Matching set',''],
    [8,'Dresser / changing table','Furniture',320,299,'High',true,'IKEA','Hemnes',''],
    [9,'Changing pad & cover','Changing Area',35,33,'High',true,'Amazon','Waterproof',''],
    [10,'Diaper pail','Changing Area',40,38,'High',true,'Ubbi','Steel',''],
    [11,'Baby monitor (Nanit Pro)','Electronics',300,279,'High',false,'Amazon','With breathing band',''],
    [12,'White noise machine','Electronics',50,48,'Medium',false,'Hatch Rest','',''],
    [13,'Blackout curtains','Lighting',80,75,'High',false,'Target','Pale lavender',''],
    [14,'Nightlight (dimmable)','Lighting',25,0,'Low',false,'IKEA','',''],
    [15,'Bookshelf','Furniture',80,75,'Medium',false,'IKEA','Floating shelves',''],
    [16,'Storage baskets x4','Storage & Organisation',60,0,'Medium',false,'Target','Woven seagrass',''],
    [17,'Drawer organisers','Storage & Organisation',25,0,'Low',false,'Amazon','',''],
    [18,'Hamper','Storage & Organisation',30,0,'Low',false,'Target','',''],
    [19,'Baby monitor mount','Electronics',20,0,'Low',false,'Amazon','',''],
    [20,'Mobile (for crib)','Décor & Wall Art',50,0,'Medium',false,'Etsy','Custom felt',''],
    [21,'Wall art x3','Décor & Wall Art',80,0,'Low',false,'Etsy','Botanical prints',''],
    [22,'Name sign / letters','Décor & Wall Art',45,0,'Low',false,'Etsy','Wooden letters',''],
    [23,'Rug (non-slip)','Bedding & Textiles',90,0,'Medium',false,'Rugs USA','Cream round',''],
    [24,'Baby gate (hall)','Safety',50,0,'High',false,'Graco','',''],
    [25,'Outlet covers','Safety',10,0,'High',false,'Amazon','',''],
    [26,'Furniture anchors','Safety',15,0,'High',false,'Amazon','',''],
    [27,'Nursing pillow (Boppy)','Feeding Station',45,43,'High',true,'Amazon','',''],
    [28,'Bottle warmer','Feeding Station',35,0,'Medium',false,'Philips Avent','',''],
    [29,'Rocking bassinet (bedside)','Furniture',200,0,'High',false,'SNOO/Halo','',''],
    [30,'Baby clothes organiser','Clothing Storage',25,0,'Low',false,'Amazon','Drawer dividers',''],
  ];
  vdata.push({ range: `'🛏️ Nursery Planner'!A6:J35`, values: items });

  vdata.push({ range: `'🛏️ Nursery Planner'!A38`, values: [['  🎨 Mood Board & Design Notes']] });
  vdata.push({ range: `'🛏️ Nursery Planner'!A39:J39`, values: [['Topic','Notes','Colour/Ref','Inspiration Link','Priority','','','','','']] });

  const notes = [
    ['Theme','Botanical garden — soft greens, lavenders, creams','#F5F2F9 + sage + warm cream','Pinterest board','High'],
    ['Wall colour','Sherwin Williams — Attitude Gray SW 6104 (very soft)','SW 6104','Paint swatch','High'],
    ['Accent wall','Removable wallpaper — soft botanical leaves in pastel','Hygge & West "Botanical"','Link TBD','Medium'],
    ['Lighting plan','Dimmable ceiling + floor lamp + Hatch nightlight','2700K warm white','','Medium'],
    ['Flooring','Existing hardwood + round cream rug','Rugs USA — Pearl Round','','Low'],
    ['Storage philosophy','Everything within reach from glider — baskets at eye level','','','Low'],
    ['Fragrance / sensory','Unscented nursery — pure essential oils diffuser only','','','Low'],
    ['Closet','PAX IKEA system — mix of hanging + shelf for 0-12M','','','Medium'],
  ];
  vdata.push({ range: `'🛏️ Nursery Planner'!A40:E47`, values: notes });

  await valuesBatchUpdate(id, vdata, 'np-values');
  console.log('Nursery Planner complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
