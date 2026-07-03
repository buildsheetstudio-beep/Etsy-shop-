'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DASH = sheetMap['🎉 Dashboard'];

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: DASH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 70 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DASH, dimension: 'ROWS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DASH, dimension: 'ROWS', startIndex: 3, endIndex: 55 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Column widths
  const colW = [30, 200, 120, 100, 110, 110, 110, 110, 110, 110, 110, 20, 160, 120, 20, 200];
  colW.forEach((w,i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: DASH, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title row A1:P1 merged
  reqs.push({ mergeCells: { range: gridRange(DASH, 0, 1, 0, 16), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 0, 1, 0, 16),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlush),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 22 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // KPI section label row 2 (A2:E2) + (G2:K2)
  const kpiSectionStyle = {
    backgroundColor: hex(C.rose),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  };
  [[0,5],[6,11]].forEach(([c0,c1]) => {
    reqs.push({ mergeCells: { range: gridRange(DASH, 1, 2, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(DASH, 1, 2, c0, c1),
      cell: { userEnteredFormat: kpiSectionStyle },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
  });

  // KPI label row 3 - two groups A-E and G-K
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 2, 3, 0, 11),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.darkPlum), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // KPI value row 4 A-E and G-K
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 3, 4, 0, 11),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.paleRose),
      textFormat: { foregroundColor: hex(C.antiqueGold), bold: true, fontSize: 15 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Divider col F (index 5) and L (index 11) — blank, narrow
  [5, 11].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(DASH, 0, 55, c, c+1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.white) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  });

  // Upcoming Events section header row 6 (index 5)
  reqs.push({ mergeCells: { range: gridRange(DASH, 5, 6, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 5, 6, 0, 11),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.rose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Upcoming events header row 7 (index 6)
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 6, 7, 0, 11),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlush),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Upcoming events data rows 8-18 (indices 7-17)
  for (let r = 7; r < 18; r++) {
    const bg = r % 2 === 0 ? C.champagne : C.ivory;
    reqs.push({ repeatCell: {
      range: gridRange(DASH, r, r+1, 0, 11),
      cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 10 }, verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // Quick Stats section (right panel M-P rows 6+)
  reqs.push({ mergeCells: { range: gridRange(DASH, 5, 6, 12, 16), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 5, 6, 12, 16),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.rose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Quick stats rows 7-18 (indices 6-17)
  for (let r = 6; r < 18; r++) {
    const bg = r % 2 === 0 ? C.champagne : C.ivory;
    reqs.push({ repeatCell: {
      range: gridRange(DASH, r, r+1, 12, 16),
      cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 10 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
    }});
  }

  // Tasks section header row 20 (index 19)
  reqs.push({ mergeCells: { range: gridRange(DASH, 19, 20, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 19, 20, 0, 11),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.rose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Task status chart data label header row 21 (index 20) — for chart source
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 20, 21, 0, 5),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.darkPlum), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Status data rows (indices 21-25)
  for (let r = 21; r < 26; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(DASH, r, r+1, 0, 5),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.ivory),
        textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 10 },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
  }

  // Events by status chart data header row 29 (index 28)
  reqs.push({ mergeCells: { range: gridRange(DASH, 27, 28, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 27, 28, 0, 11),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.rose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 28, 29, 0, 4),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.darkPlum), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  for (let r = 29; r < 36; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(DASH, r, r+1, 0, 4),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.ivory),
        textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 10 },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
  }

  // Date format for upcoming events col E (index 4)
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 7, 18, 4, 5),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});

  // Currency for budget KPI (col D value, index 3)
  reqs.push({ repeatCell: {
    range: gridRange(DASH, 3, 4, 3, 4),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});

  // Borders on upcoming events
  reqs.push({ updateBorders: {
    range: gridRange(DASH, 6, 18, 0, 11),
    innerHorizontal: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
  }});

  await batchUpdate(id, reqs, 'dashboard-format');

  const data = [];
  data.push({ range: `'🎉 Dashboard'!A1`, values: [['🎉 All-in-One Event Planner — Command Centre']] });

  // KPI section labels
  data.push({ range: `'🎉 Dashboard'!A2`, values: [['  📊 Event Overview']] });
  data.push({ range: `'🎉 Dashboard'!G2`, values: [['  ✅ Task & Budget Summary']] });

  // KPI labels row 3
  data.push({ range: `'🎉 Dashboard'!A3:E3`, values: [['Total Events','Confirmed','Completed','Planning','Cancelled']] });
  data.push({ range: `'🎉 Dashboard'!G3:K3`, values: [['Total Tasks','Done','In Progress','Blocked','Total Budget']] });

  // KPI values row 4
  data.push({ range: `'🎉 Dashboard'!A4:E4`, values: [[
    '=COUNTA(\'📋 Events Log\'!B3:B65)',
    '=COUNTIF(\'📋 Events Log\'!L3:L65,"Confirmed")',
    '=COUNTIF(\'📋 Events Log\'!L3:L65,"Completed")',
    '=COUNTIF(\'📋 Events Log\'!L3:L65,"Planning")',
    '=COUNTIF(\'📋 Events Log\'!L3:L65,"Cancelled")',
  ]] });
  data.push({ range: `'🎉 Dashboard'!G4:K4`, values: [[
    '=COUNTA(\'✅ Task Tracker\'!B5:B65)',
    '=COUNTIF(\'✅ Task Tracker\'!F5:F65,"Done")',
    '=COUNTIF(\'✅ Task Tracker\'!F5:F65,"In Progress")',
    '=COUNTIF(\'✅ Task Tracker\'!F5:F65,"Blocked")',
    '=SUM(\'💰 Event Budget\'!E4:E15)',
  ]] });

  // Upcoming Events section
  data.push({ range: `'🎉 Dashboard'!A6`, values: [['  📅 Upcoming Events (Next 90 Days)']] });
  data.push({ range: `'🎉 Dashboard'!A7:K7`, values: [['#','Event Name','Category','Type','Start Date','Status','Priority','Venue','Confirmed Guests','Budget','Notes']] });

  // Upcoming events — SORT+FILTER from Events Log (sorted by date)
  data.push({ range: `'🎉 Dashboard'!A8:K8`, values: [[
    1,
    '=IFERROR(INDEX(SORT(FILTER(\'📋 Events Log\'!B3:Q65,\'📋 Events Log\'!E3:E65>=TODAY(),\'📋 Events Log\'!E3:E65<=TODAY()+90),4,1),1,1),"")',
    '=IFERROR(INDEX(SORT(FILTER(\'📋 Events Log\'!B3:Q65,\'📋 Events Log\'!E3:E65>=TODAY(),\'📋 Events Log\'!E3:E65<=TODAY()+90),4,1),1,2),"")',
    '=IFERROR(INDEX(SORT(FILTER(\'📋 Events Log\'!B3:Q65,\'📋 Events Log\'!E3:E65>=TODAY(),\'📋 Events Log\'!E3:E65<=TODAY()+90),4,1),1,3),"")',
    '=IFERROR(INDEX(SORT(FILTER(\'📋 Events Log\'!B3:Q65,\'📋 Events Log\'!E3:E65>=TODAY(),\'📋 Events Log\'!E3:E65<=TODAY()+90),4,1),1,4),"")',
    '=IFERROR(INDEX(SORT(FILTER(\'📋 Events Log\'!B3:Q65,\'📋 Events Log\'!E3:E65>=TODAY(),\'📋 Events Log\'!E3:E65<=TODAY()+90),4,1),1,11),"")',
    '=IFERROR(INDEX(SORT(FILTER(\'📋 Events Log\'!B3:Q65,\'📋 Events Log\'!E3:E65>=TODAY(),\'📋 Events Log\'!E3:E65<=TODAY()+90),4,1),1,14),"")',
    '=IFERROR(INDEX(SORT(FILTER(\'📋 Events Log\'!B3:Q65,\'📋 Events Log\'!E3:E65>=TODAY(),\'📋 Events Log\'!E3:E65<=TODAY()+90),4,1),1,9),"")',
    '=IFERROR(INDEX(SORT(FILTER(\'📋 Events Log\'!B3:Q65,\'📋 Events Log\'!E3:E65>=TODAY(),\'📋 Events Log\'!E3:E65<=TODAY()+90),4,1),1,8),"")',
    '=IFERROR(INDEX(SORT(FILTER(\'📋 Events Log\'!B3:Q65,\'📋 Events Log\'!E3:E65>=TODAY(),\'📋 Events Log\'!E3:E65<=TODAY()+90),4,1),1,10),"")',
    '=IFERROR(INDEX(SORT(FILTER(\'📋 Events Log\'!B3:Q65,\'📋 Events Log\'!E3:E65>=TODAY(),\'📋 Events Log\'!E3:E65<=TODAY()+90),4,1),1,16),"")',
  ]] });

  // Rows 9-17 (indices 8-16) — same pattern with row offset 2-10
  for (let i = 1; i < 10; i++) {
    const rowData = [i+1];
    const n = i + 1;
    const cols = [1,2,3,4,11,14,9,8,10,16];
    cols.forEach(col => {
      rowData.push(`=IFERROR(INDEX(SORT(FILTER('📋 Events Log'!B$3:Q$65,'📋 Events Log'!E$3:E$65>=TODAY(),'📋 Events Log'!E$3:E$65<=TODAY()+90),4,1),${n},${col}),"")`);
    });
    data.push({ range: `'🎉 Dashboard'!A${8+i}:K${8+i}`, values: [rowData] });
  }

  // Quick Stats right panel
  data.push({ range: `'🎉 Dashboard'!M6`, values: [['  📈 Quick Stats']] });
  const quickStats = [
    ['Guests Confirmed','=COUNTIF(\'👥 Guest List & RSVP\'!F5:F65,"Confirmed")'],
    ['Guests Pending','=COUNTIF(\'👥 Guest List & RSVP\'!F5:F65,"Pending")'],
    ['Vendors Booked','=COUNTA(\'🏪 Vendor & Supplier Tracker\'!B5:B45)'],
    ['Vendors Paid','=COUNTIF(\'🏪 Vendor & Supplier Tracker\'!L5:L45,"Paid in Full")'],
    ['Total Budget','=SUM(\'💰 Event Budget\'!E4:E15)'],
    ['Total Spent','=SUM(\'💰 Event Budget\'!F4:F15)'],
    ['Budget Remaining','=M12-M13'],
    ['Tasks Done','=COUNTIF(\'✅ Task Tracker\'!F5:F65,"Done")'],
    ['Tasks Overdue','=COUNTIF(\'✅ Task Tracker\'!F5:F65,"Blocked")'],
    ['Avg Event Rating','=IFERROR(AVERAGE(\'📊 Event Analytics\'!L5:L35),0)'],
    ['Avg Satisfaction','=IFERROR(AVERAGE(\'📊 Event Analytics\'!M5:M35),0)'],
  ];
  quickStats.forEach((row, i) => {
    data.push({ range: `'🎉 Dashboard'!M${7+i}:N${7+i}`, values: [row] });
  });

  // Task status chart data
  data.push({ range: `'🎉 Dashboard'!A20`, values: [['  📊 Task Status Breakdown (Chart Source)']] });
  data.push({ range: `'🎉 Dashboard'!A21:E21`, values: [['Status','Done','In Progress','Not Started','Blocked']] });
  data.push({ range: `'🎉 Dashboard'!A22:E22`, values: [['Count',
    '=COUNTIF(\'✅ Task Tracker\'!F5:F65,"Done")',
    '=COUNTIF(\'✅ Task Tracker\'!F5:F65,"In Progress")',
    '=COUNTIF(\'✅ Task Tracker\'!F5:F65,"Not Started")',
    '=COUNTIF(\'✅ Task Tracker\'!F5:F65,"Blocked")',
  ]] });

  // Events by status chart data
  data.push({ range: `'🎉 Dashboard'!A28`, values: [['  📊 Events by Status (Chart Source)']] });
  data.push({ range: `'🎉 Dashboard'!A29:D29`, values: [['Status','Count','','Chart Data']] });
  const statuses = ['Planning','Confirmed','In Progress','Completed','Cancelled','Postponed'];
  statuses.forEach((s, i) => {
    data.push({ range: `'🎉 Dashboard'!A${30+i}:B${30+i}`, values: [[s, `=COUNTIF('📋 Events Log'!L$3:L$65,"${s}")`]] });
  });

  await valuesBatchUpdate(id, data, 'dashboard-values');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
