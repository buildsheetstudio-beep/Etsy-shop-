'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['✨ Dashboard'];

const BUD = "'💰 Budget & Expense Tracker'";
const TSK = "'✅ Task Manager'";
const HAB = "'🔁 Habit Tracker'";
const WHL = "'⚖️ Life Balance Wheel'";
const JRN = "'📓 Journal & Gratitude Log'";

(async () => {
  const data = [];

  // Row 1: Title
  data.push({ range: "'✨ Dashboard'!A1", values: [['✨ MY LIFE PLANNER']] });
  // Row 3: Name input
  data.push({ range: "'✨ Dashboard'!A3:B3", values: [['Your Name', '']] });
  // Row 4: Greeting
  data.push({ range: "'✨ Dashboard'!A4", values: [['=IFERROR("Good "&IF(HOUR(NOW())<12,"morning",IF(HOUR(NOW())<18,"afternoon","evening"))&", "&B3&"! Here\'s your life at a glance.","")']] });

  // TODAY SECTION (rows 6-14)
  data.push({ range: "'✨ Dashboard'!A6", values: [['🗓️ TODAY']] });
  data.push({ range: "'✨ Dashboard'!A7:B7", values: [["Today's Date", '=TEXT(TODAY(),"MMMM D, YYYY")']] });
  data.push({ range: "'✨ Dashboard'!A8", values: [['🍽️ MEAL TODAY']] });
  data.push({ range: "'✨ Dashboard'!A9:B9",  values: [['Breakfast', '']] });
  data.push({ range: "'✨ Dashboard'!A10:B10", values: [['Lunch', '']] });
  data.push({ range: "'✨ Dashboard'!A11:B11", values: [['Dinner', '']] });
  data.push({ range: "'✨ Dashboard'!A12:B12", values: [['Snack', '']] });
  data.push({ range: "'✨ Dashboard'!A13", values: [['📌 TOP TASKS DUE TODAY']] });
  data.push({ range: "'✨ Dashboard'!A14", values: [[`=IFERROR(TEXTJOIN(CHAR(10),TRUE,FILTER(${TSK}!A:A,(${TSK}!E:E=TODAY())*(${TSK}!D:D<>"Done"))),"✅ Nothing due today — enjoy!")`]] });

  // FINANCE SECTION (rows 16-21)
  data.push({ range: "'✨ Dashboard'!A16", values: [['💰 FINANCES']] });
  data.push({ range: "'✨ Dashboard'!A17:B17", values: [['Total Income This Month', `=IFERROR(SUMIFS(${BUD}!E:E,${BUD}!D:D,"Income",${BUD}!A:A,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),${BUD}!A:A,"<"&DATE(YEAR(TODAY()),MONTH(TODAY())+1,1)),0)`]] });
  data.push({ range: "'✨ Dashboard'!A18:B18", values: [['Total Spent This Month', `=IFERROR(SUMIFS(${BUD}!E:E,${BUD}!D:D,"Expense",${BUD}!A:A,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),${BUD}!A:A,"<"&DATE(YEAR(TODAY()),MONTH(TODAY())+1,1)),0)`]] });
  data.push({ range: "'✨ Dashboard'!A19:B19", values: [['Left to Budget', '=IFERROR(B17-B18,0)']] });
  data.push({ range: "'✨ Dashboard'!A20", values: [['📋 BILLS DUE THIS WEEK']] });
  data.push({ range: "'✨ Dashboard'!A21", values: [[`=IFERROR(IF(COUNTIFS(${BUD}!I:I,">="&TODAY(),${BUD}!I:I,"<="&TODAY()+7,${BUD}!K:K,FALSE)=0,"✅ No unpaid bills due this week",TEXTJOIN(CHAR(10),TRUE,FILTER(${BUD}!H:H,(${BUD}!I:I>=TODAY())*(${BUD}!I:I<=TODAY()+7)*(${BUD}!K:K=FALSE)))),"")`]] });

  // ORGANIZATION SECTION (rows 23-33)
  data.push({ range: "'✨ Dashboard'!A23", values: [['🎯 ORGANIZATION']] });
  data.push({ range: "'✨ Dashboard'!A24", values: [['LIFE BALANCE (Current → Goal)']] });
  const cats = ['Personal Growth','Health & Wellness','Career & Business','Home & Household','Relationships & Social','Finance'];
  cats.forEach((cat, i) => {
    const r = 25 + i;
    data.push({ range: `'✨ Dashboard'!A${r}:C${r}`, values: [[cat, `=IFERROR(INDEX(${WHL}!B:B,MATCH("${cat}",${WHL}!A:A,0)),"—")`, `=IFERROR(INDEX(${WHL}!C:C,MATCH("${cat}",${WHL}!A:A,0)),"—")`]] });
  });
  data.push({ range: "'✨ Dashboard'!A31", values: [['TASK PRIORITY COUNTS']] });
  data.push({ range: "'✨ Dashboard'!A32:D32", values: [['High/To Do','High/In Prog','Med Remaining','Done Total']] });
  data.push({ range: "'✨ Dashboard'!A33:D33", values: [[
    `=IFERROR(COUNTIFS(${TSK}!C:C,"High",${TSK}!D:D,"To Do"),0)`,
    `=IFERROR(COUNTIFS(${TSK}!C:C,"High",${TSK}!D:D,"In Progress"),0)`,
    `=IFERROR(COUNTIFS(${TSK}!D:D,"<>Done",${TSK}!C:C,"Medium"),0)`,
    `=IFERROR(COUNTIF(${TSK}!D:D,"Done"),0)`,
  ]] });

  // WELLNESS SECTION (rows 35-40)
  data.push({ range: "'✨ Dashboard'!A35", values: [['💚 WELLNESS']] });
  data.push({ range: "'✨ Dashboard'!A36:B36", values: [['Habit Completion % This Week', `=IFERROR(TEXT(COUNTIFS(${HAB}!A:A,">="&TODAY()-WEEKDAY(TODAY(),2)+1,${HAB}!A:A,"<="&TODAY(),${HAB}!C:C,TRUE)/COUNTIFS(${HAB}!A:A,">="&TODAY()-WEEKDAY(TODAY(),2)+1,${HAB}!A:A,"<="&TODAY()),"0%"),"—")`]] });
  data.push({ range: "'✨ Dashboard'!A37:B37", values: [['Habits Logged This Week', `=IFERROR(COUNTIFS(${HAB}!A:A,">="&TODAY()-WEEKDAY(TODAY(),2)+1,${HAB}!A:A,"<="&TODAY(),${HAB}!C:C,TRUE)&" completed","—")`]] });
  data.push({ range: "'✨ Dashboard'!A38", values: [['MOOD THIS WEEK (last 7 days)']] });
  data.push({ range: "'✨ Dashboard'!A39", values: [[`=IFERROR(TEXTJOIN(" · ",TRUE,FILTER(${JRN}!F:F,${JRN}!A:A>=TODAY()-6)),"No mood entries this week.")`]] });
  data.push({ range: "'✨ Dashboard'!A40:B40", values: [['Weekly Reflection Reminder','=IF(WEEKDAY(TODAY())=1,"📝 Weekly reflection due today!","")']] });

  await valuesBatchUpdate(id, data, 'dash-values');

  const reqs = [];

  // Title banner
  reqs.push({ mergeCells: { range: gridRange(DSH,0,1,0,6), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,0,1,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.raspberry), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 60 }, fields: 'pixelSize' } });

  // Row 2 spacer
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 10 }, fields: 'pixelSize' } });

  // Row 3 name input
  reqs.push({ repeatCell: { range: gridRange(DSH,2,3,0,1), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 10 }, horizontalAlignment: 'RIGHT' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  reqs.push({ mergeCells: { range: gridRange(DSH,2,3,1,4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,2,3,1,4), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.raspberry) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // Row 4 greeting
  reqs.push({ mergeCells: { range: gridRange(DSH,3,4,0,6), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,3,4,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { italic: true, fontSize: 11, foregroundColor: hex(C.raspberry) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // Spacer row 5
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 8 }, fields: 'pixelSize' } });

  // TODAY section header (idx 5)
  reqs.push({ mergeCells: { range: gridRange(DSH,5,6,0,6), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,5,6,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.raspberry), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // TODAY section bg (idx 6-13)
  reqs.push({ repeatCell: { range: gridRange(DSH,6,14,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.todayTint) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Date row (idx 6)
  reqs.push({ repeatCell: { range: gridRange(DSH,6,7,0,1), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 10 }, horizontalAlignment: 'RIGHT' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });
  reqs.push({ mergeCells: { range: gridRange(DSH,6,7,1,4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,6,7,1,4), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.raspberry) } } }, fields: 'userEnteredFormat(textFormat)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });

  // Meal sub-header (idx 7)
  reqs.push({ mergeCells: { range: gridRange(DSH,7,8,0,4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,7,8,0,4), cell: { userEnteredFormat: { backgroundColor: hex(C.mint), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 7, endIndex: 8 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // Meal inputs (idx 8-11)
  reqs.push({ repeatCell: { range: gridRange(DSH,8,12,0,1), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 9 }, horizontalAlignment: 'RIGHT' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,8,12,1,4), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { fontSize: 10 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Top tasks sub-header (idx 12)
  reqs.push({ mergeCells: { range: gridRange(DSH,12,13,0,4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,12,13,0,4), cell: { userEnteredFormat: { backgroundColor: hex(C.mint), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 12, endIndex: 13 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // Top tasks formula (idx 13)
  reqs.push({ mergeCells: { range: gridRange(DSH,13,14,0,4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,13,14,0,4), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { fontSize: 9 }, wrapStrategy: 'WRAP', verticalAlignment: 'TOP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 13, endIndex: 14 }, properties: { pixelSize: 60 }, fields: 'pixelSize' } });

  // Spacer (idx 14)
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 14, endIndex: 15 }, properties: { pixelSize: 8 }, fields: 'pixelSize' } });

  // FINANCE section header (idx 15)
  reqs.push({ mergeCells: { range: gridRange(DSH,15,16,0,6), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,15,16,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.raspberry), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 15, endIndex: 16 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Finance bg (idx 16-21)
  reqs.push({ repeatCell: { range: gridRange(DSH,16,22,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.financeTint) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // KPI rows (idx 16-18)
  [16,17,18].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(DSH,ri,ri+1,0,1), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 10 }, horizontalAlignment: 'RIGHT' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });
    reqs.push({ mergeCells: { range: gridRange(DSH,ri,ri+1,1,3), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(DSH,ri,ri+1,1,3), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.raspberry) }, numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat)' } });
    reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  });

  // Bills sub-header (idx 19)
  reqs.push({ mergeCells: { range: gridRange(DSH,19,20,0,4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,19,20,0,4), cell: { userEnteredFormat: { backgroundColor: hex(C.mint), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 19, endIndex: 20 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // Bills formula (idx 20)
  reqs.push({ mergeCells: { range: gridRange(DSH,20,21,0,4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,20,21,0,4), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { fontSize: 9 }, wrapStrategy: 'WRAP', verticalAlignment: 'TOP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 20, endIndex: 21 }, properties: { pixelSize: 60 }, fields: 'pixelSize' } });

  // Spacer (idx 21)
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 21, endIndex: 22 }, properties: { pixelSize: 8 }, fields: 'pixelSize' } });

  // ORGANIZATION section header (idx 22)
  reqs.push({ mergeCells: { range: gridRange(DSH,22,23,0,6), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,22,23,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.raspberry), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 22, endIndex: 23 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Org bg (idx 23-32)
  reqs.push({ repeatCell: { range: gridRange(DSH,23,33,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.orgTint) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Life balance sub-header (idx 23)
  reqs.push({ mergeCells: { range: gridRange(DSH,23,24,0,4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,23,24,0,4), cell: { userEnteredFormat: { backgroundColor: hex(C.mint), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 23, endIndex: 24 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // Life balance rows (idx 24-29)
  for (let ri = 24; ri <= 29; ri++) {
    reqs.push({ repeatCell: { range: gridRange(DSH,ri,ri+1,0,1), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 9 }, horizontalAlignment: 'RIGHT' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });
    reqs.push({ repeatCell: { range: gridRange(DSH,ri,ri+1,1,2), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.mint) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
    reqs.push({ repeatCell: { range: gridRange(DSH,ri,ri+1,2,3), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.raspberry) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
    reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  }

  // Task priority sub-header (idx 30)
  reqs.push({ mergeCells: { range: gridRange(DSH,30,31,0,4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,30,31,0,4), cell: { userEnteredFormat: { backgroundColor: hex(C.mint), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 30, endIndex: 31 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // Task counts (idx 31-32)
  reqs.push({ repeatCell: { range: gridRange(DSH,31,32,0,4), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 8 }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,32,33,0,4), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.raspberry) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  // Spacer (idx 33)
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 33, endIndex: 34 }, properties: { pixelSize: 8 }, fields: 'pixelSize' } });

  // WELLNESS section header (idx 34)
  reqs.push({ mergeCells: { range: gridRange(DSH,34,35,0,6), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,34,35,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.raspberry), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 34, endIndex: 35 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Wellness bg (idx 35-40)
  reqs.push({ repeatCell: { range: gridRange(DSH,35,41,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.wellTint) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Habit % (idx 35)
  reqs.push({ repeatCell: { range: gridRange(DSH,35,36,0,1), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 9 }, horizontalAlignment: 'RIGHT' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });
  reqs.push({ mergeCells: { range: gridRange(DSH,35,36,1,3), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,35,36,1,3), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.raspberry) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 35, endIndex: 36 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });

  // Habits logged (idx 36)
  reqs.push({ repeatCell: { range: gridRange(DSH,36,37,0,1), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 9 }, horizontalAlignment: 'RIGHT' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });
  reqs.push({ mergeCells: { range: gridRange(DSH,36,37,1,3), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,36,37,1,3), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { fontSize: 10, foregroundColor: hex(C.darkText) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Mood sub-header (idx 37)
  reqs.push({ mergeCells: { range: gridRange(DSH,37,38,0,4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,37,38,0,4), cell: { userEnteredFormat: { backgroundColor: hex(C.mint), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 37, endIndex: 38 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // Mood formula (idx 38)
  reqs.push({ mergeCells: { range: gridRange(DSH,38,39,0,4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,38,39,0,4), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { fontSize: 9 }, wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 38, endIndex: 39 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });

  // Reflection reminder (idx 39)
  reqs.push({ mergeCells: { range: gridRange(DSH,39,40,0,4), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,39,40,0,4), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.amber) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Column widths
  [[0,180],[1,120],[2,100],[3,120],[4,180],[5,100]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'dash-format');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
