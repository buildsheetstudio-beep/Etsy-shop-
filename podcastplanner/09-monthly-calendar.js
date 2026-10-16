'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CAL = sheetMap['📅 Monthly Calendar'];

// Layout:
// Row 0: "📅 Monthly Calendar View" title banner
// Row 1: B1=Year (input), B2=Month (input) label row
// Row 2: Day-of-week headers (Sun Mon Tue Wed Thu Fri Sat) cols A-G
// Rows 3-8: 6 calendar grid rows (6 weeks × 7 days)
// Default view: January 2024 (year=2024, month=1)
// Ep 1 on Jan 8 (Mon), Ep 2 on Jan 15 (Mon), Ep 3 on Jan 22 (Mon), Ep 4 on Jan 29 (Mon)
// Jan 2024 starts on Monday (day 1 = Monday, offset col B)

const DAYS = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

// Jan 2024: 1=Mon(col B idx 1), so Sun col empty for week 1
// week 1: _, 1, 2, 3, 4, 5, 6
// week 2: 7, 8, 9, 10, 11, 12, 13
// week 3: 14, 15, 16, 17, 18, 19, 20
// week 4: 21, 22, 23, 24, 25, 26, 27
// week 5: 28, 29, 30, 31, _, _, _

// Grid: rows 3-8 (0-based), cols A-G (0-6)
// Day numbers: 7 cols × 6 rows = 42 cells
// Jan 2024 starts on Monday → offset = 1 (Sun=0)
const JAN_OFFSET = 1;
const JAN_DAYS = 31;

function dayCell(dayNum, row) {
  // Formula: pull episode titles from Content Calendar where date matches
  // B1=year, B2=month
  return `=IFERROR(TEXTJOIN(", ",TRUE,FILTER('🎙 Content Calendar'!E:E,'🎙 Content Calendar'!F:F=DATE(B1,B2,${dayNum}))),"")`;
}

(async () => {
  const reqs = [];

  // Row 0: Title
  reqs.push({ mergeCells: { range: gridRange(CAL, 0, 1, 0, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(CAL, 0, 1, 0, 7),
      cell: {
        userEnteredValue: { stringValue: '📅 Monthly Calendar View' },
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
    range: { sheetId: CAL, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Row 1: Year/Month selector
  reqs.push({
    repeatCell: {
      range: gridRange(CAL, 1, 2, 0, 7),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepCharcoal),
          textFormat: { foregroundColor: hex(C.warmCream), bold: true, fontSize: 10 },
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  // Input cells B1 (col 1) and D1 (col 3) with burnt orange bg
  for (const col of [1, 3]) {
    reqs.push({
      repeatCell: {
        range: gridRange(CAL, 1, 2, col, col+1),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.burntOrange),
            textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat',
      },
    });
  }
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CAL, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});

  // Row 2: Day-of-week headers
  reqs.push({
    repeatCell: {
      range: gridRange(CAL, 2, 3, 0, 7),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.burntOrange),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CAL, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});

  // Calendar grid rows 3-8: alternating light bg
  for (let r = 3; r < 9; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(CAL, r, r+1, 0, 7),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(r % 2 === 0 ? C.inputBg : C.altRow),
            verticalAlignment: 'TOP',
            wrapStrategy: 'WRAP',
            textFormat: { fontSize: 9 },
          },
        },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: CAL, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 70 }, fields: 'pixelSize',
    }});
  }

  // Col widths: 7 equal columns
  for (let i = 0; i < 7; i++) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: CAL, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: 140 }, fields: 'pixelSize',
    }});
  }

  // Borders on calendar grid cells
  const calGridRange = gridRange(CAL, 3, 9, 0, 7);
  reqs.push({
    updateBorders: {
      range: calGridRange,
      innerHorizontal: { style: 'SOLID', color: hex(C.gray), width: 1 },
      innerVertical: { style: 'SOLID', color: hex(C.gray), width: 1 },
      top: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
      bottom: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
      left: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
      right: { style: 'SOLID', color: hex(C.deepCharcoal), width: 2 },
    },
  });

  await batchUpdate(id, reqs, 'monthly-format');

  // Values
  const data = [];

  // Row 1 labels and inputs
  data.push({
    range: `'📅 Monthly Calendar'!A2:G2`,
    values: [['Year:', 2024, 'Month:', 1, '', 'Auto-updates when you change Year/Month above', '']],
  });

  // Day-of-week headers
  data.push({
    range: `'📅 Monthly Calendar'!A3:G3`,
    values: [DAYS],
  });

  // Build 6-row × 7-col grid
  // Jan 2024 starts Monday → offset 1 (0=Sun, 1=Mon)
  // cells: position 0..41
  // day number = cellIndex - offset + 1; valid if 1..31

  const gridValues = [];
  for (let week = 0; week < 6; week++) {
    const weekRow = [];
    for (let dow = 0; dow < 7; dow++) {
      const cellIndex = week * 7 + dow;
      const dayNum = cellIndex - JAN_OFFSET + 1;
      if (dayNum >= 1 && dayNum <= JAN_DAYS) {
        weekRow.push(dayCell(dayNum, 4 + week));
      } else {
        weekRow.push('');
      }
    }
    gridValues.push(weekRow);
  }

  // Day numbers in top-left of each cell (write as text prefix in a separate pass)
  // We'll encode day number in front of the formula text by writing to a second column range
  // Actually, the simplest approach: each calendar cell shows the FILTER result.
  // Let's add day numbers separately as the first cell value — but Google Sheets doesn't support
  // multi-line cell content easily. Instead, we'll prefix day number as a helper via formula:
  // =TEXT(day,"0")&IF(IFERROR(FILTER(...))="","",CHAR(10)&FILTER(...))
  // Let's rewrite the formula with day number prefix:
  const gridValuesWithDayNum = [];
  for (let week = 0; week < 6; week++) {
    const weekRow = [];
    for (let dow = 0; dow < 7; dow++) {
      const cellIndex = week * 7 + dow;
      const dayNum = cellIndex - JAN_OFFSET + 1;
      if (dayNum >= 1 && dayNum <= JAN_DAYS) {
        const filterPart = `IFERROR(TEXTJOIN(", ",TRUE,FILTER('🎙 Content Calendar'!E:E,'🎙 Content Calendar'!F:F=DATE(B2,D2,${dayNum}))),"")`;
        weekRow.push(`="${dayNum}"&IF(${filterPart}="","",CHAR(10)&${filterPart})`);
      } else {
        weekRow.push('');
      }
    }
    gridValuesWithDayNum.push(weekRow);
  }

  data.push({ range: `'📅 Monthly Calendar'!A4:G9`, values: gridValuesWithDayNum });

  await valuesBatchUpdate(id, data, 'monthly-values');

  // CF: cells that have episode content (len > 2 chars = has episode) → light accent fill
  // Use CUSTOM_FORMULA: =LEN(A4)>3 (day number alone is 1-2 chars, with episode is longer)
  const cfReqs = [];
  for (let r = 3; r < 9; r++) {
    for (let c = 0; c < 7; c++) {
      const colLetter = String.fromCharCode(65 + c);
      cfReqs.push({
        addConditionalFormatRule: {
          rule: {
            ranges: [gridRange(CAL, r, r+1, c, c+1)],
            booleanRule: {
              condition: {
                type: 'CUSTOM_FORMULA',
                values: [{ userEnteredValue: `=LEN(${colLetter}${r+1})>3` }],
              },
              format: {
                backgroundColor: hex(C.lightAmber),
                textFormat: { bold: true, foregroundColor: hex(C.darkText) },
              },
            },
          },
          index: 0,
        },
      });
    }
  }
  await batchUpdate(id, cfReqs, 'monthly-cf');

  console.log('Monthly Calendar complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
