'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const HM = sheetMap['🌈 52-Week Severity Heatmap'];
const DSL = sheetMap['📅 Daily Symptom Log'];

// Grid: row 1 = title, row 2 = Year Start Date input + week headers (Week 1 ... Week 52)
//       row 3 = day-of-week label + formula grid rows (7 days × 52 weeks)
// Col A = day label (Sun-Sat), cols B-BA = weeks 1-52
// Formula per cell: pull severity from Daily Symptom Log by date
// date = $B$2 + (col_offset)*7 + row_offset
//   where col_offset = COLUMN()-2 (week index, 0-based), row_offset = ROW()-3 (day index 0=Sun)

const DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

(async () => {
  const data = [];

  // Row 1: title
  data.push({ range: "'🌈 52-Week Severity Heatmap'!A1", values: [['🌈 52-WEEK SEVERITY HEATMAP']] });

  // Row 2: "Year Start Date" label in A2, date input in B2, week headers Week 1 .. Week 52 in C2..BB2
  data.push({ range: "'🌈 52-Week Severity Heatmap'!A2", values: [['Year Start Date →']] });
  data.push({ range: "'🌈 52-Week Severity Heatmap'!B2", values: [['2025-01-01']] });

  const weekHeaders = [['Week']]; // B2 is the start date; C2 onward = week numbers
  // Actually: B2 = date input. C2..BB2 = Week 1..52 (cols 2..53 in 0-index = C..BB)
  // We use B2 as anchor. Week labels go in C2..BB2
  const weekLabelRow = [['Year Start Date →', '2025-01-01']];
  for (let w = 1; w <= 52; w++) weekLabelRow[0].push(`Wk ${w}`);
  data.push({ range: "'🌈 52-Week Severity Heatmap'!A2", values: weekLabelRow });

  // Rows 3-9: DOW labels in col A, formulas in cols B-BB (53 cols: B=week1-day0 ... BB=week52-day6)
  // Cell at row r (0-indexed from row 3 = ri 2), col c (0-indexed from col B = ci 1):
  // date = $B$2 + (ci-1)*7 + (ri-2)  where ri = ROW()-1 (1-indexed row), ci = COLUMN()-1
  // In A1 notation: date = $B$2 + (COLUMN()-2)*7 + (ROW()-3)
  for (let dow = 0; dow < 7; dow++) {
    const sheetRow = dow + 3; // rows 3-9 in A1 notation
    const rowData = [DOW[dow]]; // col A label
    for (let week = 1; week <= 52; week++) {
      // COLUMN() for week 1 starts at col B = column index 2 in A1 terms
      // date for (week, dow): $B$2 + (week-1)*7 + dow
      // Formula using COLUMN() and ROW():
      const formula = `=IFERROR(INDEX('📅 Daily Symptom Log'!$B:$B,MATCH($B$2+(COLUMN()-2)*7+(ROW()-3),'📅 Daily Symptom Log'!$A:$A,0)),"")`;
      rowData.push(formula);
    }
    data.push({ range: `'🌈 52-Week Severity Heatmap'!A${sheetRow}:BB${sheetRow}`, values: [rowData] });
  }

  await valuesBatchUpdate(id, data, 'hm-values');

  const reqs = [];

  // Title banner row 1 — merge A1:BB1
  reqs.push({ mergeCells: { range: gridRange(HM, 0, 1, 0, 54), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(HM, 0, 1, 0, 54),
      cell: { userEnteredFormat: { backgroundColor: hex(C.periwinkle), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: HM, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });

  // Row 2: header row — label cells + week number cells
  reqs.push({
    repeatCell: {
      range: gridRange(HM, 1, 2, 0, 2),
      cell: { userEnteredFormat: { backgroundColor: hex(C.periwinkle), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(HM, 1, 2, 2, 54),
      cell: { userEnteredFormat: { backgroundColor: hex(C.periwinkle), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 8 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: HM, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // DOW label col A (rows 3-9)
  reqs.push({
    repeatCell: {
      range: gridRange(HM, 2, 9, 0, 1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { foregroundColor: hex(C.darkText), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Heatmap cell grid (rows 3-9, cols B-BB) — base formatting
  reqs.push({
    repeatCell: {
      range: gridRange(HM, 2, 9, 1, 54),
      cell: { userEnteredFormat: { backgroundColor: hex(C.fog), textFormat: { foregroundColor: hex(C.darkText), fontSize: 8 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', numberFormat: { type: 'NUMBER', pattern: '0' } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,numberFormat)',
    },
  });

  // Column widths: col A=40 (DOW), col B=90 (date input), cols C-BB=22 each (weeks)
  reqs.push({ updateDimensionProperties: { range: { sheetId: HM, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: HM, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 90 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: HM, dimension: 'COLUMNS', startIndex: 2, endIndex: 54 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // Row heights for data rows 3-9
  reqs.push({ updateDimensionProperties: { range: { sheetId: HM, dimension: 'ROWS', startIndex: 2, endIndex: 9 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Input cell B2 — white background
  reqs.push({
    repeatCell: {
      range: gridRange(HM, 1, 2, 1, 2),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { foregroundColor: hex(C.darkText), bold: true, fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // Color scale CF for heatmap grid — calm blue, not danger-red
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(HM, 2, 9, 2, 54)],
        gradientRule: {
          minpoint: { color: hex(C.heatMin), type: 'NUMBER', value: '0' },
          midpoint: { color: hex(C.heatMid), type: 'NUMBER', value: '50' },
          maxpoint: { color: hex(C.heatMax), type: 'NUMBER', value: '100' },
        },
      },
      index: 0,
    },
  });

  await batchUpdate(id, reqs, 'hm-format');
  console.log('52-Week Severity Heatmap complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
