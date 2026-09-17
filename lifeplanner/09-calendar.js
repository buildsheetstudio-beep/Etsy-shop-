'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CAL = sheetMap['📅 Monthly Calendar View'];

const TSK_NAME = "'✅ Task Manager'";
const BUD_NAME = "'💰 Budget & Expense Tracker'";
const DOW = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];

(async () => {
  const data = [];
  data.push({ range: "'📅 Monthly Calendar View'!A1", values: [['📅 MONTHLY CALENDAR VIEW']] });
  data.push({ range: "'📅 Monthly Calendar View'!A2:B2", values: [['Month (1-12)', 1]] });
  data.push({ range: "'📅 Monthly Calendar View'!C2:D2", values: [['Year', 2025]] });
  data.push({ range: "'📅 Monthly Calendar View'!A3:G3", values: [DOW] });

  for (let week = 0; week < 6; week++) {
    const rowIdx = 4 + week;
    const rowData = [];
    for (let dow = 0; dow < 7; dow++) {
      const offset = week * 7 + dow;
      const dayRef = `DATE($D$2,$B$2,1)+${offset}-WEEKDAY(DATE($D$2,$B$2,1),2)+1`;
      const inMonth = `IF(MONTH(${dayRef})=$B$2,${dayRef},"")`;
      const formula = `=IFERROR(IF(IFERROR(${inMonth},"")="","",TEXT(${inMonth},"D")&CHAR(10)&IFERROR(TEXTJOIN(", ",TRUE,FILTER(${TSK_NAME}!A:A,${TSK_NAME}!E:E=${dayRef})),"")&IFERROR(IF(COUNTIF(${BUD_NAME}!I:I,${dayRef})>0,CHAR(10)&"Bill: "&TEXTJOIN(", ",TRUE,FILTER(${BUD_NAME}!H:H,${BUD_NAME}!I:I=${dayRef})),""),"")),"")`;
      rowData.push(formula);
    }
    data.push({ range: `'📅 Monthly Calendar View'!A${rowIdx}:G${rowIdx}`, values: [rowData] });
  }
  await valuesBatchUpdate(id, data, 'cal-values');

  const reqs = [];
  reqs.push({ mergeCells: { range: gridRange(CAL,0,1,0,7), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(CAL,0,1,0,7), cell: { userEnteredFormat: { backgroundColor: hex(C.raspberry), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  // Selector row
  reqs.push({ repeatCell: { range: gridRange(CAL,1,2,0,1), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 9 }, horizontalAlignment: 'RIGHT' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  reqs.push({ repeatCell: { range: gridRange(CAL,1,2,1,2), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.raspberry) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  reqs.push({ repeatCell: { range: gridRange(CAL,1,2,2,3), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 9 }, horizontalAlignment: 'RIGHT' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  reqs.push({ repeatCell: { range: gridRange(CAL,1,2,3,4), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.raspberry) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // DOW headers
  reqs.push({ repeatCell: { range: gridRange(CAL,2,3,0,7), cell: { userEnteredFormat: { backgroundColor: hex(C.mint), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  for (let week = 0; week < 6; week++) {
    const ri = 3 + week;
    reqs.push({ repeatCell: { range: gridRange(CAL,ri,ri+1,0,7), cell: { userEnteredFormat: { backgroundColor: hex(week%2===0?C.altRow:C.white), textFormat: { fontSize: 8, foregroundColor: hex(C.darkText) }, wrapStrategy: 'WRAP', verticalAlignment: 'TOP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });
  }
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'ROWS', startIndex: 3, endIndex: 9 }, properties: { pixelSize: 90 }, fields: 'pixelSize' } });
  for (let ci = 0; ci < 7; ci++) {
    reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: 130 }, fields: 'pixelSize' } });
  }
  await batchUpdate(id, reqs, 'cal-format');
  console.log('Monthly Calendar View complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
