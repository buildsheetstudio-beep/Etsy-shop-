'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const ANA = sheetMap['📈 Episode Analytics'];

// A=Episode # | B=Episode Title (formula) | C=Publish Date | D=Downloads Week 1
// E=Downloads Month 1 | F=Total Downloads to Date | G=Avg Listen Duration (min) | H=Top Platform

const HEADERS = ['Episode #','Episode Title','Publish Date','Downloads Week 1','Downloads Month 1','Total Downloads to Date','Avg Listen Duration (min)','Top Platform'];
const WIDTHS  = [90, 260, 110, 120, 120, 140, 140, 120];

// 20 rows — realistic growth pattern (newer shows = lower cumulative)
const sampleData = [
  [1,'2024-01-08',312,680,2840,34,'Spotify'],
  [2,'2024-01-15',428,910,3100,41,'Apple Podcasts'],
  [3,'2024-01-22',385,820,2720,28,'Spotify'],
  [4,'2024-01-29',510,1080,3490,37,'Apple Podcasts'],
  [5,'2024-02-05',467,980,3210,32,'Spotify'],
  [6,'2024-02-12',390,820,2580,29,'YouTube'],
  [7,'2024-02-19',445,930,2940,36,'Spotify'],
  [8,'2024-02-26',620,1340,4120,44,'Apple Podcasts'],
  [9,'2024-03-04',408,870,2640,31,'Spotify'],
  [10,'2024-03-11',540,1150,3380,38,'Apple Podcasts'],
  [11,'2024-03-18',590,1270,3720,43,'Spotify'],
  [12,'2024-03-25',360,760,2140,26,'YouTube'],
  [13,'2024-04-01',670,1450,4260,45,'Spotify'],
  [14,'2024-04-08',420,880,2430,29,'Apple Podcasts'],
  [15,'2024-04-15',580,1240,3380,37,'Spotify'],
  [16,'2024-04-22',640,1380,3700,40,'Apple Podcasts'],
  [17,'2024-04-29',510,1090,2760,34,'Spotify'],
  [18,'2024-05-06',480,1020,2440,31,'YouTube'],
  [19,'2024-05-13',550,1180,2650,36,'Spotify'],
  [20,'2024-05-20',720,1560,3210,47,'Apple Podcasts'],
];

const CAL = "'🎙 Content Calendar'";

(async () => {
  const reqs = [];

  // Row 0: Title (col A standalone for freeze compat)
  reqs.push({ mergeCells: { range: gridRange(ANA, 0, 1, 1, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(ANA, 0, 1, 0, 1),
      cell: {
        userEnteredValue: { stringValue: '📈' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepCharcoal),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(ANA, 0, 1, 1, 8),
      cell: {
        userEnteredValue: { stringValue: 'Episode Analytics' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepCharcoal),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ANA, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Row 1: Headers
  reqs.push({
    repeatCell: {
      range: gridRange(ANA, 1, 2, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepCharcoal),
          textFormat: { foregroundColor: hex(C.burntOrange), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ANA, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});

  // Data rows
  for (let r = 2; r < 22; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(ANA, r, r+1, 0, 8),
        cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.inputBg : C.altRow), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: ANA, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
  }
  reqs.push({
    repeatCell: {
      range: gridRange(ANA, 22, 500, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat',
    },
  });

  // Col B (index 1): formula bg
  reqs.push({ repeatCell: {
    range: gridRange(ANA, 2, 500, 1, 2),
    cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg) } },
    fields: 'userEnteredFormat',
  }});

  // Col widths
  for (const [i, px] of WIDTHS.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: ANA, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  // Number formats
  reqs.push({ repeatCell: {
    range: gridRange(ANA, 2, 500, 2, 3),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  for (const col of [3, 4, 5]) {
    reqs.push({ repeatCell: {
      range: gridRange(ANA, 2, 500, col, col+1),
      cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat',
    }});
  }
  reqs.push({ repeatCell: {
    range: gridRange(ANA, 2, 500, 6, 7),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0.0' }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(ANA, 2, 500, 0, 1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true } } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(ANA, 2, 500, 7, 8),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});

  await batchUpdate(id, reqs, 'ana-format');

  // Values
  const data = [];
  data.push({ range: `'📈 Episode Analytics'!A2:H2`, values: [HEADERS] });

  for (let i = 0; i < sampleData.length; i++) {
    const row = i + 3;
    const [epNum, publishDate, w1, m1, total, duration, platform] = sampleData[i];
    const titleFormula = `=IFERROR(INDEX(${CAL}!$E:$E,MATCH(A${row},${CAL}!$A:$A,0)),"")`;
    data.push({
      range: `'📈 Episode Analytics'!A${row}:H${row}`,
      values: [[epNum, titleFormula, publishDate, w1, m1, total, duration, platform]],
    });
  }

  // Empty row formulas
  const emptyRows = [];
  for (let row = 23; row <= 200; row++) {
    emptyRows.push(['', `=IFERROR(INDEX(${CAL}!$E:$E,MATCH(A${row},${CAL}!$A:$A,0)),"")`, '', '', '', '', '', '']);
  }
  data.push({ range: `'📈 Episode Analytics'!A23`, values: emptyRows });

  await valuesBatchUpdate(id, data, 'ana-values');
  console.log('Episode Analytics complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
