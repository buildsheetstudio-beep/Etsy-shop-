'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Reference Data'];
const S = "'Reference Data'";

const LISTS = {
  A: { header: 'Genre',         items: ['Fantasy','Mystery','Science Fiction','Literary Fiction','Romance','Historical Fiction','Thriller','Horror','Non-Fiction','Biography','Memoir','Self-Help','Graphic Novel','Young Adult','Poetry','Short Stories','Children\'s','Other'] },
  B: { header: 'Status',        items: ['To Read','Reading','Finished','DNF','Paused','Abandoned'] },
  C: { header: 'Format',        items: ['Hardcover','Paperback','eBook','Audiobook','Library Book','Digital Library','Other'] },
  D: { header: 'Language',      items: ['English','Spanish','French','German','Italian','Portuguese','Japanese','Chinese','Korean','Other'] },
  E: { header: 'Priority',      items: ['High','Medium','Low'] },
  F: { header: 'Source',        items: ['Personal Discovery','Friend Recommendation','Book Club','Online Review','Best-of List','Author Follow','Social Media','Bookstore Browse','Other'] },
  G: { header: 'Goal Type',     items: ['Annual Book Count','Genre Goal','Author Goal','Series Goal','Page Count Goal','Reading Challenge','Other'] },
  H: { header: 'Goal Status',   items: ['Not Started','In Progress','Achieved','Behind Schedule','Paused','Abandoned'] },
  I: { header: 'Challenge Type',items: ['Reading Bingo','Genre Challenge','Seasonal Read','Award Nominees','Series Marathon','Author Deep Dive','Decade Challenge','Other'] },
  J: { header: 'Yes / No',      items: ['Yes','No'] },
  K: { header: 'Shelf Name',    items: ['Favorites','To Read Soon','Fiction','Non-Fiction','Series','Classics','Reference','Borrowed','E-Reader','Loaned Out'] },
  L: { header: 'Rating Stars',  items: ['1 ★','2 ★★','3 ★★★','4 ★★★★','5 ★★★★★'] },
};

(async () => {
  const vals = [];
  const fmt  = [];

  // Background wash
  fmt.push({ repeatCell: { range: gridRange(SID,0,200,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Section header
  vals.push({ range: `${S}!A1`, values: [['DROPDOWN REFERENCE LISTS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  const colLetters = Object.keys(LISTS);
  for (const col of colLetters) {
    const ci = col.charCodeAt(0) - 65;
    const list = LISTS[col];
    vals.push({ range: `${S}!${col}2`, values: [[list.header]] });
    fmt.push({ repeatCell: { range: gridRange(SID,1,2,ci,ci+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat' }});
    list.items.forEach((item, ii) => {
      vals.push({ range: `${S}!${col}${3+ii}`, values: [[item]] });
    });
    fmt.push({ repeatCell: { range: gridRange(SID,2,2+list.items.length,ci,ci+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.white), textFormat: { fontSize: 8, fontFamily: 'Arial' },
    }}, fields: 'userEnteredFormat' }});
  }

  // Section 2: Reading speed reference
  const speedRow = 23;
  vals.push({ range: `${S}!A${speedRow}`, values: [['READING SPEED REFERENCE (words per minute → illustrative averages)']] });
  fmt.push({ mergeCells: { range: gridRange(SID,speedRow-1,speedRow,0,6), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,speedRow-1,speedRow,0,6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.accent), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});

  const speedHdr = ['Reader Type','Avg WPM','Pages/Hour (250 WPM/page)','Hours for 300-page book','Notes'];
  vals.push({ range: `${S}!A${speedRow+1}`, values: [speedHdr] });
  fmt.push({ repeatCell: { range: gridRange(SID,speedRow,speedRow+1,0,5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.greenTint), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial' }, horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});

  const speedData = [
    ['Slow (below average)', 150, 36, 8.3, 'Takes time to fully absorb content'],
    ['Average',              250, 60, 5.0, 'Most adult readers'],
    ['Above Average',        350, 84, 3.6, 'Avid readers'],
    ['Speed Reader',         600,144, 2.1, 'With comprehension techniques'],
  ];
  speedData.forEach((row, ri) => {
    vals.push({ range: `${S}!A${speedRow+2+ri}`, values: [row] });
    fmt.push({ repeatCell: { range: gridRange(SID,speedRow+1+ri,speedRow+2+ri,0,5), cell: { userEnteredFormat: {
      backgroundColor: hex(ri % 2 === 0 ? C.white : C.altRow), textFormat: { fontSize: 8, fontFamily: 'Arial' },
    }}, fields: 'userEnteredFormat' }});
  });

  // Column widths
  [120,100,80,80,70,80,110,100,110,70,110,100].forEach((px,ci) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  await valuesBatchUpdate(id, vals, '02-reference values');
  await batchUpdate(id, fmt, '02-reference format');

  await batchUpdate(id, [{
    updateSheetProperties: {
      properties: { sheetId: SID, hidden: true },
      fields: 'hidden'
    }
  }], '02-reference hide');
  console.log('✅  Reference Data done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
