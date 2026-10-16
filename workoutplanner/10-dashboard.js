'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard'];

// Layout:
//   Row 1 (idx 0): Title merged A1:H1
//   Row 2 (idx 1): Subtitle merged A2:H2
//   Row 3 (idx 2): Spacer / input row — A3="Year:", B3=2026, C3="Month:", D3=7
//   Row 4 (idx 3): Safety note merged A4:H4
//   Row 5 (idx 4): blank spacer
//   Rows 6-7 (idx 5-6): KPI label row — 5 KPIs across B6:H6
//   Rows 8-9 (idx 7-8): KPI value row
//   Rows 10-11 (idx 9-10): blank spacer
//   Row 12 (idx 11): "📊 Key Stats" section header
//   Rows 13-17 (idx 12-16): Quick-glance stats
//   Rows 18-28 (idx 17-27): blank / chart area (charts overlay here)
//   Row 29 (idx 28): Chart data section header
//   Row 30 (idx 29): Muscle Group table header
//   Rows 31-40 (idx 30-39): Muscle Group COUNTIF rows (10 groups)
//   Row 41 (idx 40): blank
//   Row 42 (idx 41): Cardio Distance table header
//   Rows 43-48 (idx 42-47): Cardio Type SUMIF rows (6 types)
//   Row 49 (idx 48): blank
//   Row 50 (idx 49): Sleep Quality table header
//   Rows 51-54 (idx 50-53): Sleep Quality COUNTIF rows (4 levels)

const muscleGroups = ['Chest','Back','Shoulders','Arms','Core','Legs','Glutes','Full Body','Cardio','Mobility'];
const cardioTypes  = ['Running','Cycling','Swimming','Rowing','Elliptical','Jump Rope'];
const sleepQuality = ['Excellent','Good','Fair','Poor'];

(async () => {
  // ── Static text & input values ─────────────────────────────────────────────
  await valuesBatchUpdate(id, [
    { range: "'📊 Dashboard'!A1", values: [['🏋️ Ultimate Workout Planner']] },
    { range: "'📊 Dashboard'!A2", values: [['Track Every Rep, Every Run, Every Recovery']] },
    { range: "'📊 Dashboard'!A3", values: [['Year:',2026,'Month:',7]] },
    { range: "'📊 Dashboard'!A4", values: [['⚕️  SAFETY NOTICE: Consult a physician before starting a new exercise program, especially if you have any pre-existing health conditions.']] },
    // KPI Labels row 6
    { range: "'📊 Dashboard'!B6", values: [['Workouts Completed','','PRs Tracked','Avg Sleep (hrs)','Cardio Distance (mi)']] },
    // KPI formulas row 8
    { range: "'📊 Dashboard'!B8", values: [[
      `=COUNTIFS('📝 Workout Log'!$A$3:$A$32,">="&DATE(B3,D3,1),'📝 Workout Log'!$A$3:$A$32,"<="&EOMONTH(DATE(B3,D3,1),0),'📝 Workout Log'!$J$3:$J$32,TRUE)`,
      '',
      `=COUNTIF('🏆 PR Tracker'!$B$3:$B$14,">0")`,
      `=IFERROR(AVERAGEIFS('😴 Rest & Recovery'!$B$3:$B$32,'😴 Rest & Recovery'!$A$3:$A$32,">="&DATE(B3,D3,1),'😴 Rest & Recovery'!$A$3:$A$32,"<="&EOMONTH(DATE(B3,D3,1),0)),0)`,
      `=IFERROR(SUMIFS('🏃 Cardio Tracker'!$C$3:$C$12,'🏃 Cardio Tracker'!$A$3:$A$12,">="&DATE(B3,D3,1),'🏃 Cardio Tracker'!$A$3:$A$12,"<="&EOMONTH(DATE(B3,D3,1),0)),0)`,
    ]] },
    // Secondary stats row 12-16
    { range: "'📊 Dashboard'!A12", values: [['📊 Key Stats']] },
    { range: "'📊 Dashboard'!A13", values: [
      ['Total Workouts (All Time)', `=COUNTIF('📝 Workout Log'!$J$3:$J$32,TRUE)`],
      ['Current Body Weight (lbs)', `=IFERROR(INDEX('📏 Body Measurements'!$B$3:$B$10,COUNTA('📏 Body Measurements'!$B$3:$B$10)),"")`],
      ['Body Fat % Change', `=IFERROR(INDEX('📏 Body Measurements'!$C$3:$C$10,COUNTA('📏 Body Measurements'!$C$3:$C$10))-INDEX('📏 Body Measurements'!$C$3:$C$10,1),"")`],
      ['Rest Days (This Month)', `=COUNTIFS('😴 Rest & Recovery'!$A$3:$A$32,">="&DATE(B3,D3,1),'😴 Rest & Recovery'!$A$3:$A$32,"<="&EOMONTH(DATE(B3,D3,1),0),'😴 Rest & Recovery'!$F$3:$F$32,TRUE)`],
      ['Most Trained Muscle Group', `=IFERROR(INDEX('📋 Reference Data'!$A$2:$A$11,MATCH(MAX(COUNTIF('📝 Workout Log'!$C$3:$C$32,'📋 Reference Data'!$A$2:$A$11)),COUNTIF('📝 Workout Log'!$C$3:$C$32,'📋 Reference Data'!$A$2:$A$11),0)),"")`],
    ] },
    // Chart data header
    { range: "'📊 Dashboard'!A29", values: [['— Chart Data (Internal) —']] },
    // Muscle Group table
    { range: "'📊 Dashboard'!A30", values: [['Muscle Group','Workout Count']] },
    { range: "'📊 Dashboard'!A31", values: muscleGroups.map(mg => [mg, `=COUNTIF('📝 Workout Log'!$C$3:$C$32,"${mg}")`]) },
    // Cardio table
    { range: "'📊 Dashboard'!A42", values: [['Cardio Type','Total Distance (mi)']] },
    { range: "'📊 Dashboard'!A43", values: cardioTypes.map(ct => [ct, `=IFERROR(SUMIF('🏃 Cardio Tracker'!$B$3:$B$12,"${ct}",'🏃 Cardio Tracker'!$C$3:$C$12),0)`]) },
    // Sleep Quality table
    { range: "'📊 Dashboard'!A50", values: [['Sleep Quality','Count']] },
    { range: "'📊 Dashboard'!A51", values: sleepQuality.map(sq => [sq, `=COUNTIF('😴 Rest & Recovery'!$C$3:$C$32,"${sq}")`]) },
  ], 'dashboard-data');

  // ── Link calendar inputs to dashboard ─────────────────────────────────────
  await valuesBatchUpdate(id, [
    { range: "'📅 Monthly Calendar'!B1", values: [["='📊 Dashboard'!B3"]] },
    { range: "'📅 Monthly Calendar'!D1", values: [["='📊 Dashboard'!D3"]] },
  ], 'calendar-link');

  // ── Formatting ────────────────────────────────────────────────────────────
  const reqs = [];

  // Merge rows 1-4
  reqs.push({ mergeCells: { range: gridRange(DSH, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } }); // Title
  reqs.push({ mergeCells: { range: gridRange(DSH, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } }); // Subtitle
  reqs.push({ mergeCells: { range: gridRange(DSH, 3, 4, 0, 8), mergeType: 'MERGE_ALL' } }); // Safety note

  // Title row A1
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 0, 1, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.headerBlue),
          textFormat: { foregroundColor: hex(C.white), fontSize: 16, bold: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Subtitle row A2
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 1, 2, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.cobaltBlue),
          textFormat: { foregroundColor: hex(C.lightBlue), fontSize: 11, italic: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Input row A3:D3
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 2, 3, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.lightBlue),
          textFormat: { foregroundColor: hex(C.headerBlue), fontSize: 11 },
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    },
  });
  // Input values B3, D3 — bold
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 2, 3, 1, 2),
      cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 12 }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(textFormat,horizontalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 2, 3, 3, 4),
      cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 12 }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(textFormat,horizontalAlignment)',
    },
  });

  // Safety note A4 — amber background
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 3, 4, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.lightAmber),
          textFormat: { foregroundColor: hex(C.darkText), fontSize: 10, bold: true },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });

  // KPI section
  // Merge B6:C6 for "Workouts Completed" (wider)
  reqs.push({ mergeCells: { range: gridRange(DSH, 5, 6, 1, 3), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(DSH, 7, 8, 1, 3), mergeType: 'MERGE_ALL' } });

  // KPI label cells (row 6, idx 5)
  const kpiLabelBg = hex(C.mutedBlue);
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 5, 6, 1, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: kpiLabelBg,
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // KPI value cells (row 8, idx 7)
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 7, 8, 1, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.lightBlue),
          textFormat: { foregroundColor: hex(C.headerBlue), bold: true, fontSize: 22 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Avg Sleep format 0.0
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 7, 8, 4, 5),
      cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0' } } },
      fields: 'userEnteredFormat(numberFormat)',
    },
  });

  // Key Stats header row 12 (idx 11)
  reqs.push({ mergeCells: { range: gridRange(DSH, 11, 12, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 11, 12, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.cobaltBlue),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Stats rows 13-17 (idx 12-16)
  for (let r = 12; r < 17; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, r, r + 1, 0, 1),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.lightBlue),
            textFormat: { foregroundColor: hex(C.headerBlue), bold: true, fontSize: 10 },
            horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, r, r + 1, 1, 3),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(r % 2 === 0 ? C.offWhite : C.altRow),
            textFormat: { foregroundColor: hex(C.cobaltBlue), bold: true, fontSize: 11 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
  }

  // Chart data area (dim text so it doesn't distract)
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 28, 55, 0, 2),
      cell: {
        userEnteredFormat: {
          textFormat: { foregroundColor: hex(C.mediumGray), fontSize: 9 },
        },
      },
      fields: 'userEnteredFormat(textFormat)',
    },
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 7, endIndex: 8 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });

  // Column widths
  [[0,155],[1,130],[2,130],[3,130],[4,130],[5,130],[6,130],[7,130]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'dashboard-format');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
