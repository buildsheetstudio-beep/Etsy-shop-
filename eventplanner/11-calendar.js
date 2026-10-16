'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CAL = sheetMap['🗓️ Events Calendar'];

(async () => {
  const reqs = [];

  // Row heights: title(52), nav(36), day-names(28), calendar weeks 4-9 (70px each)
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'ROWS', startIndex: 3, endIndex: 10 }, properties: { pixelSize: 70 }, fields: 'pixelSize' } });

  // Column widths: A(index col, 40px), B-H (Mon-Sun, 130px each)
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });
  for (let c = 1; c <= 7; c++) {
    reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'COLUMNS', startIndex: c, endIndex: c+1 }, properties: { pixelSize: 130 }, fields: 'pixelSize' } });
  }
  // Extra col I for padding
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'COLUMNS', startIndex: 8, endIndex: 9 }, properties: { pixelSize: 150 }, fields: 'pixelSize' } });

  // Title row A1:I1 merged
  reqs.push({ mergeCells: { range: gridRange(CAL, 0, 1, 0, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(CAL, 0, 1, 0, 9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.plum),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Nav row: A2=label, B2=Month(input), C2=label, D2=Year(input), E2:I2=title display
  reqs.push({ mergeCells: { range: gridRange(CAL, 1, 2, 4, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(CAL, 1, 2, 0, 9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.darkPlum), bold: true, fontSize: 11 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  // Month/Year input cells — ivory + gold border
  [1, 3].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(CAL, 1, 2, c, c+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.ivory),
        textFormat: { foregroundColor: hex(C.antiqueGold), bold: true, fontSize: 13 },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});
    reqs.push({ updateBorders: {
      range: gridRange(CAL, 1, 2, c, c+1),
      top: { style: 'SOLID', color: hex(C.antiqueGold), width: 2 },
      bottom: { style: 'SOLID', color: hex(C.antiqueGold), width: 2 },
      left: { style: 'SOLID', color: hex(C.antiqueGold), width: 2 },
      right: { style: 'SOLID', color: hex(C.antiqueGold), width: 2 },
    }});
  });
  // Calendar title (Month Year) in E2
  reqs.push({ repeatCell: {
    range: gridRange(CAL, 1, 2, 4, 9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.deepBlush), bold: true, fontSize: 15 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});

  // Day name row (Mon–Sun) B3:H3
  reqs.push({ repeatCell: {
    range: gridRange(CAL, 2, 3, 1, 8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlush),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Calendar grid cells B4:H9 (rows 3-8, cols 1-7) — warmWhite bg, wrap, top-align
  for (let r = 3; r < 10; r++) {
    for (let c = 1; c <= 7; c++) {
      reqs.push({ repeatCell: {
        range: gridRange(CAL, r, r+1, c, c+1),
        cell: { userEnteredFormat: {
          backgroundColor: hex(C.warmWhite),
          textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 9 },
          verticalAlignment: 'TOP',
          wrapStrategy: 'WRAP',
        }},
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,wrapStrategy)',
      }});
    }
  }

  // Row index col A4:A9
  for (let r = 3; r < 10; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(CAL, r, r+1, 0, 1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.champagne),
        textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 9 },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
  }

  // Grid borders
  reqs.push({ updateBorders: {
    range: gridRange(CAL, 2, 10, 1, 8),
    top: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    left: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    right: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    innerHorizontal: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
  }});

  // Legend area row 11 onwards
  reqs.push({ mergeCells: { range: gridRange(CAL, 10, 11, 1, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(CAL, 10, 11, 1, 5),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.darkPlum), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  await batchUpdate(id, reqs, 'calendar-format');

  const data = [];
  data.push({ range: `'🗓️ Events Calendar'!A1`, values: [['🗓️ Events Calendar']] });
  data.push({ range: `'🗓️ Events Calendar'!A2`, values: [['Month:']] });
  data.push({ range: `'🗓️ Events Calendar'!B2`, values: [[9]] }); // September
  data.push({ range: `'🗓️ Events Calendar'!C2`, values: [['Year:']] });
  data.push({ range: `'🗓️ Events Calendar'!D2`, values: [[2025]] });
  data.push({ range: `'🗓️ Events Calendar'!E2`, values: [['=TEXT(DATE($D$2,$B$2,1),"MMMM YYYY")']] });

  // Day names
  data.push({ range: `'🗓️ Events Calendar'!B3:H3`, values: [['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']] });

  // Row labels
  data.push({ range: `'🗓️ Events Calendar'!A4:A9`, values: [['W1'],['W2'],['W3'],['W4'],['W5'],['W6']] });

  // Calendar cell formula — shows day number + events on that day
  // Each cell: LET(weekStart, DATE($D$2,$B$2,1) - WEEKDAY(DATE($D$2,$B$2,1),2) + 1, dayDate, weekStart + (ROW()-4)*7 + (COL()-2), IF(MONTH(dayDate)<>$B$2, "", TEXT(DAY(dayDate),"0") & CHAR(10) & IFERROR(TEXTJOIN(CHAR(10),TRUE,FILTER('📋 Events Log'!B3:B65,IFERROR(DATEVALUE(TEXT('📋 Events Log'!E3:E65,"yyyy-mm-dd"))=dayDate,FALSE))),"") ))
  // We'll generate the 42 formulas for rows 4-9, cols B-H
  const calFormulas = [];
  for (let week = 0; week < 6; week++) {
    const row = [];
    for (let dow = 0; dow < 7; dow++) {
      // ROW() for row 4 is 4, etc. COL() for col B is 2, etc.
      const rowRef = 4 + week;
      const colRef = 2 + dow;
      const f = `=LET(weekStart,DATE($D$2,$B$2,1)-WEEKDAY(DATE($D$2,$B$2,1),2)+1,dayDate,weekStart+(${rowRef}-4)*7+(${colRef}-2),IF(MONTH(dayDate)<>$B$2,"",TEXT(DAY(dayDate),"0")&CHAR(10)&IFERROR(TEXTJOIN(CHAR(10),TRUE,FILTER('📋 Events Log'!$B$3:$B$65,IFERROR(INT('📋 Events Log'!$E$3:$E$65)=dayDate,FALSE))),"⠀")))`;
      row.push(f);
    }
    calFormulas.push(row);
  }
  data.push({ range: `'🗓️ Events Calendar'!B4:H9`, values: calFormulas });

  data.push({ range: `'🗓️ Events Calendar'!B11`, values: [['💡 Events are auto-filled from the Events Log tab. Enter month (1-12) and year above to navigate.']] });

  await valuesBatchUpdate(id, data, 'calendar-values');
  console.log('Events Calendar complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
