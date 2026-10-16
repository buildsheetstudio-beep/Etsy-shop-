'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Monthly Calendar'];
const S = "'Monthly Calendar'";

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,60,0,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Title row
  vals.push({ range: `${S}!A1`, values: [['MONTHLY CALENDAR — ADHD FOCUS FLOW PLANNER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,10), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Month/year selector row 2
  vals.push({ range: `${S}!A2`, values: [['Month:','','Year:','','','','','Monthly Summary:','','Tasks:']] });
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.coralTint), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  vals.push({ range: `${S}!B2`, values: [['September']] });
  vals.push({ range: `${S}!D2`, values: [['2026']] });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,1,2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.white), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.coral) },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,3,4), cell: { userEnteredFormat: {
    backgroundColor: hex(C.white), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.coral) },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});

  // Day of week headers row 3
  vals.push({ range: `${S}!A3`, values: [['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday','','','']] });
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,5,7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.sand), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // September 2026 calendar
  // Sep 1 = Tuesday (so Mon=blank, Tue=1)
  // Week 1: blank, 1, 2, 3, 4, 5, 6
  // Week 2: 7, 8, 9, 10, 11, 12, 13
  // Week 3: 14, 15, 16, 17, 18, 19, 20
  // Week 4: 21, 22, 23, 24, 25, 26, 27
  // Week 5: 28, 29, 30, blank, blank, blank, blank

  const WEEK_DATES = [
    ['', 1, 2, 3, 4, 5, 6],
    [7, 8, 9, 10, 11, 12, 13],
    [14, 15, 16, 17, 18, 19, 20],
    [21, 22, 23, 24, 25, 26, 27],
    [28, 29, 30, '', '', '', ''],
  ];

  // Events/notes for specific dates in September 2026
  const DATE_NOTES = {
    1:  '',
    14: 'TODAY — Weekly review\nQ3 report draft',
    15: 'Q3 report due\nTeam standup 9am',
    17: 'Presentation Friday\nPresent Q3 slides',
    21: 'Dentist appointment\nCar registration due',
    24: 'Order Mom gift\nBook restaurant',
    25: "MOM'S BIRTHDAY\nFamily dinner",
    28: 'Performance review\nMonthly budget review',
  };

  // Each calendar week = 2 rows: row 1 = date number, row 2 = event notes
  // Rows 4-13 = 5 weeks × 2 rows
  let calRowStart = 3; // 0-indexed row 4
  WEEK_DATES.forEach((week, wi) => {
    const dateRow = calRowStart + wi * 2;
    const noteRow = dateRow + 1;

    // Date number row
    const dateVals = week.map(d => d === '' ? '' : `${d}`);
    vals.push({ range: `${S}!A${dateRow+1}`, values: [dateVals.concat(['','',''])] });

    // Style date number row
    week.forEach((d, ci) => {
      if (ci >= 7) return;
      const isWeekend = ci >= 5;
      const isToday = d === 14;
      const bg = isToday ? C.coral : (isWeekend ? C.sandTint : C.white);
      const textColor = isToday ? C.white : (d === '' ? C.medGray : C.slate);
      fmt.push({ repeatCell: { range: gridRange(SID,dateRow,dateRow+1,ci,ci+1), cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { bold: isToday, fontSize: 11, foregroundColor: hex(textColor), fontFamily: 'Arial' },
        horizontalAlignment: 'LEFT', verticalAlignment: 'TOP',
        padding: { top: 4, left: 4 },
      }}, fields: 'userEnteredFormat' }});
    });
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: dateRow, endIndex: dateRow+1 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

    // Note row
    const noteVals = week.map(d => DATE_NOTES[d] || '');
    vals.push({ range: `${S}!A${noteRow+1}`, values: [noteVals.concat(['','',''])] });

    // Style note row
    week.forEach((d, ci) => {
      if (ci >= 7) return;
      const isWeekend = ci >= 5;
      const isToday = d === 14;
      const bg = isToday ? C.coralTint : (isWeekend ? C.sandTint : C.white);
      fmt.push({ repeatCell: { range: gridRange(SID,noteRow,noteRow+1,ci,ci+1), cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { fontSize: 8, foregroundColor: hex(C.text), fontFamily: 'Arial' },
        horizontalAlignment: 'LEFT', verticalAlignment: 'TOP',
        padding: { top: 2, left: 4 },
        wrapStrategy: 'WRAP',
      }}, fields: 'userEnteredFormat' }});
    });
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: noteRow, endIndex: noteRow+1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' }});

    // Cell borders for calendar grid
    week.forEach((d, ci) => {
      if (ci >= 7) return;
      fmt.push({ updateBorders: { range: gridRange(SID,dateRow,noteRow+1,ci,ci+1),
        top:    { style: 'SOLID', width: 1, color: hex(C.medGray) },
        bottom: { style: 'SOLID', width: 1, color: hex(C.medGray) },
        left:   { style: 'SOLID', width: 1, color: hex(C.medGray) },
        right:  { style: 'SOLID', width: 1, color: hex(C.medGray) },
      }});
    });
  });

  // Summary sidebar (col H-J) — monthly stats
  const SUM_LABELS = [
    ['','Tasks Completed',''],
    ['','Tasks Planned',''],
    ['','Completion Rate',''],
    ['','High Priority Done',''],
    ['','Events This Month',''],
    ['','Focus Days',''],
    ['','Low Cap Days',''],
  ];
  vals.push({ range: `${S}!H4`, values: SUM_LABELS });
  const SUM_VALS = [
    ['','=COUNTIFS(\'Master Tasks\'!E8:E5007,"Completed",\'Master Tasks\'!F8:F5007,">="&DATE(2026,9,1),\'Master Tasks\'!F8:F5007,"<="&DATE(2026,9,30))',''],
    ['','=COUNTA(\'Master Tasks\'!C8:C5007)',''],
    ['','=IFERROR(TEXT(H5/H6,"0%"),"0%")',''],
    ['','=COUNTIFS(\'Master Tasks\'!E8:E5007,"Completed",\'Master Tasks\'!D8:D5007,"Must Do")',''],
    ['','=COUNTIF(\'Master Tasks\'!B8:B5007,"Event")',''],
    ['','',''],
    ['','',''],
  ];
  vals.push({ range: `${S}!H4`, values: SUM_LABELS });
  vals.push({ range: `${S}!H11`, values: SUM_VALS });
  fmt.push({ repeatCell: { range: gridRange(SID,3,13,7,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.blueTint), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});

  // Column widths for 7 calendar columns + sidebar
  const WIDTHS = [100,100,100,100,100,100,100,80,120,80];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1-3
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '08-monthlycal values');
  await batchUpdate(id, fmt, '08-monthlycal format');
  console.log('✅  Monthly Calendar done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
