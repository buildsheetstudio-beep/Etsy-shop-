'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const AN = sheetMap['📊 Event Analytics'];

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: AN, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: AN, dimension: 'ROWS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: AN, dimension: 'ROWS', startIndex: 3, endIndex: 35 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // Column widths: A(30) B(Event 200) C(Category 120) D(Type 100) E(Date 100) F(Capacity 90) G(Attendees 90) H(Attendance% 100) I(Revenue 110) J(Expenses 110) K(Net P/L 110) L(Rating 70) M(Satisfaction 80) N(Notes 200)
  const colW = [30, 200, 120, 100, 100, 90, 90, 100, 110, 110, 110, 70, 80, 200];
  colW.forEach((w,i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: AN, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title row
  reqs.push({ mergeCells: { range: gridRange(AN, 0, 1, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(AN, 0, 1, 0, 14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.tealSage),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // KPI row 2 labels
  reqs.push({ repeatCell: {
    range: gridRange(AN, 1, 2, 0, 14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.darkPlum), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // KPI row 3 values
  reqs.push({ repeatCell: {
    range: gridRange(AN, 2, 3, 0, 14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.paleRose),
      textFormat: { foregroundColor: hex(C.darkPlum), bold: true, fontSize: 11 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Header row 4
  reqs.push({ repeatCell: {
    range: gridRange(AN, 3, 4, 0, 14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.tealSage),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Data rows alternating
  for (let r = 4; r < 35; r++) {
    const bg = r % 2 === 0 ? C.champagne : C.ivory;
    reqs.push({ repeatCell: {
      range: gridRange(AN, r, r+1, 0, 14),
      cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 10 }, verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // Percent format H (attendance%)
  reqs.push({ repeatCell: {
    range: gridRange(AN, 4, 35, 7, 8),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});

  // Currency format I, J, K
  [8, 9, 10].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(AN, 4, 35, c, c+1),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  });

  // Date format E
  reqs.push({ repeatCell: {
    range: gridRange(AN, 4, 35, 4, 5),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});

  // Formula cols H, K — paleRose
  [7, 10].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(AN, 4, 35, c, c+1), cell: { userEnteredFormat: { backgroundColor: hex(C.paleRose) } }, fields: 'userEnteredFormat.backgroundColor' }});
  });

  // Center cols A, C, D, F, G, H, L, M
  [0, 2, 3, 5, 6, 7, 11, 12].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(AN, 3, 35, c, c+1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat.horizontalAlignment' }});
  });

  // Borders
  reqs.push({ updateBorders: {
    range: gridRange(AN, 3, 35, 0, 14),
    innerHorizontal: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
  }});

  await batchUpdate(id, reqs, 'analytics-format');

  const data = [];
  data.push({ range: `'📊 Event Analytics'!A1`, values: [['📊 Event Analytics — Performance Insights']] });

  // KPI row 2 labels
  data.push({ range: `'📊 Event Analytics'!A2:N2`, values: [['Events Completed','Avg Attendance %','Total Revenue','Total Expenses','Net P/L','Avg Rating','Avg Satisfaction','Events In Planning','','','','','','']] });
  // KPI row 3 values
  data.push({ range: `'📊 Event Analytics'!A3:H3`, values: [[
    '=COUNTIF(\'📋 Events Log\'!L3:L65,"Completed")',
    '=IFERROR(AVERAGE(H5:H35),0)',
    '=SUM(I5:I35)',
    '=SUM(J5:J35)',
    '=C3-D3',
    '=IFERROR(AVERAGE(L5:L35),0)',
    '=IFERROR(AVERAGE(M5:M35),0)',
    '=COUNTIF(\'📋 Events Log\'!L3:L65,"Planning")',
  ]] });

  data.push({ range: `'📊 Event Analytics'!A4:N4`, values: [['#','Event Name','Category','Type','Date','Capacity','Attendees','Attendance %','Revenue','Expenses','Net P/L','Rating /5','Satisfaction /5','Notes']] });

  const analytics = [
    [1,'Tech Summit 2025','Conference','In-Person','2025-09-15',150,142,'=IFERROR(G5/F5,0)',12500,15700,'=I5-J5',4.5,4.2,'Strong attendance; AV issue minor'],
    [2,'Virtual Workshop — AI Trends','Workshop','Virtual','2025-09-03',200,186,'=IFERROR(G6/F6,0)',3800,980,'=I6-J6',4.7,4.5,'High engagement, excellent feedback'],
    [3,'Networking Drinks — Q4','Social','In-Person','2025-10-09',80,0,'=IFERROR(G7/F7,0)',0,0,'=I7-J7',0,0,'Upcoming — no data yet'],
  ];

  data.push({ range: `'📊 Event Analytics'!A5:N7`, values: analytics });

  await valuesBatchUpdate(id, data, 'analytics-values');
  console.log('Event Analytics complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
