'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange, colLetter } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = sheetMap['Habit Tracker']; // 4

// 12 habits: cols A-G (info), I-AM (days 1-31, skipping H as spacer)
// Col layout: A=#, B=Habit, C=Life Area, D=Frequency, E=Goal, F=Month Total, G=Completion%, H=spacer, I=Day1 ... AM=Day31
// Actually: A=# B=Habit C=Area D=Freq E=Goal/target F=spacer, then I-AM = 31 days
// Let's do: A=# B=Habit C=Area D=Freq E=Goal F=Done G=% H=sparkline trend I...AM=days 1-31

// Col I = index 8, AM = index 38 (I+30 = 8+30 = 38)
// Day col n = colIndex 8 + (n-1)

(async () => {
  const reqs = [];

  // ── Column widths ─────────────────────────────────────────────────────────
  // A(28), B(180), C(90), D(80), E(50), F(50), G(70), H(80/sparkline), then I-AM: 28px each
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 185 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 90 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 80 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 65 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 7, endIndex: 8 }, properties: { pixelSize: 80 }, fields: 'pixelSize' } });
  // Day columns I(8) through AM(38) = 31 cols of 28px
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 8, endIndex: 39 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // ── Row heights ────────────────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  // Habit rows 4-15 (12 habits)
  for (let r = 3; r <= 14; r++) {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r + 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  }
  // Summary row 16
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 15, endIndex: 16 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  // Totals row 17
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 16, endIndex: 17 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  // Streak row 18
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 17, endIndex: 18 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 39), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 39),
      cell: { userEnteredFormat: { backgroundColor: hex(C.dustyLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Row 2: Subtitle ───────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { foregroundColor: hex(C.deepPurple), italic: true, fontSize: 10 }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  // Day-of-week mini header in row 2 cols I-AM
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 8, 39),
      cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { foregroundColor: hex(C.mutedText), fontSize: 7 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Row 3: Column headers ─────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 2, 3, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.deepPurple), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 8 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });
  // Day number headers (1-31) in row 3 cols I-AM
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 2, 3, 8, 39),
      cell: { userEnteredFormat: { backgroundColor: hex(C.medLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 8 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Life area colors for habit rows ──────────────────────────────────────
  const habitAreaColors = [
    C.paleSage,   // 1 Morning meditation (Health)
    C.paleSage,   // 2 Daily run/workout (Health)
    C.paleSage,   // 3 Drink 2L water (Health)
    C.paleSky,    // 4 Work on course (Career)
    C.paleSky,    // 5 LinkedIn post (Career)
    C.paleGold,   // 6 No unnecessary spend (Finance)
    C.paleGold,   // 7 Track spending (Finance)
    C.paleBlush,  // 8 Gratitude to partner (Relationships)
    C.paleLavender, // 9 Read 30 min (Personal)
    C.paleLavender, // 10 Spanish practice (Personal)
    C.paleLavender, // 11 Journal (Personal)
    '#FED7AA',    // 12 Something fun/creative (Fun)
  ];

  const areaAccents = [
    C.dustySage, C.dustySage, C.dustySage,
    C.skyBlue, C.skyBlue,
    C.softGold, C.softGold,
    C.blush,
    C.dustyLavender, C.dustyLavender, C.dustyLavender,
    '#F97316',
  ];

  for (let r = 3; r <= 14; r++) {
    const i = r - 3;
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 0, 8),
        cell: { userEnteredFormat: { backgroundColor: hex(habitAreaColors[i]), textFormat: { fontSize: 9 }, verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
    // # col
    reqs.push({ repeatCell: { range: gridRange(SID, r, r + 1, 0, 1), cell: { userEnteredFormat: { textFormat: { bold: true }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });
    // Checkbox area (days 1-31) background
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 8, 39),
        cell: { userEnteredFormat: { backgroundColor: hex(habitAreaColors[i] === '#FED7AA' ? '#FFF8F3' : C.lavenderWhite) } },
        fields: 'userEnteredFormat.backgroundColor',
      },
    });
    // Area accent stripe on col A
    reqs.push({ repeatCell: { range: gridRange(SID, r, r + 1, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(areaAccents[i]), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  }

  // ── Summary rows (15-17) ──────────────────────────────────────────────────
  // Row 15: "TOTALS" header
  reqs.push({ mergeCells: { range: gridRange(SID, 15, 16, 0, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 15, 16, 0, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.deepPurple), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  // Row 15 day totals (COUNTIF)
  reqs.push({ repeatCell: { range: gridRange(SID, 15, 16, 8, 39), cell: { userEnteredFormat: { backgroundColor: hex(C.deepPurple), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 8 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // Row 16: Daily completion % row
  reqs.push({ mergeCells: { range: gridRange(SID, 16, 17, 0, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 16, 17, 0, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.medLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 16, 17, 8, 39), cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { fontSize: 8 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,numberFormat)' } });

  await batchUpdate(id, reqs, 'habit-format');

  // ── Values — headers ──────────────────────────────────────────────────────
  const headerVals = [
    { range: 'Habit Tracker!A1', values: [['✦  June 2025 Habit Tracker  ✦']] },
    { range: 'Habit Tracker!A2', values: [['Jordan  |  Word of Year: FLOURISH  |  June 2025']] },
    // Day-of-week row 2 (row index 1)
    { range: 'Habit Tracker!I2:AM2', values: [['Su','Mo','Tu','We','Th','Fr','Sa','Su','Mo','Tu','We','Th','Fr','Sa','Su','Mo','Tu','We','Th','Fr','Sa','Su','Mo','Tu','We','Th','Fr','Sa','Mo','Tu','Mo']] },
    // Column headers row 3
    { range: 'Habit Tracker!A3:H3', values: [['#', 'Habit', 'Life Area', 'Frequency', 'Days Done', 'Days Tried', '% Done', 'Trend']] },
    // Day numbers row 3
    { range: 'Habit Tracker!I3:AM3', values: [[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31]] },
  ];

  // 12 habits
  const habits = [
    ['1','Morning Meditation (20 min)','Health & Fitness','Daily'],
    ['2','Daily Run or Workout','Health & Fitness','5x/week'],
    ['3','Drink 2L Water','Health & Fitness','Daily'],
    ['4','Work on Course (1 hr)','Career & Purpose','5x/week'],
    ['5','LinkedIn Post or Engagement','Career & Purpose','3x/week'],
    ['6','No Unnecessary Spend','Finance & Wealth','Daily'],
    ['7','Track Spending in App','Finance & Wealth','Daily'],
    ['8','Express Gratitude to Partner','Relationships','Daily'],
    ['9','Read 30 Minutes','Personal Growth','Daily'],
    ['10','Spanish Practice (Duolingo)','Personal Growth','Daily'],
    ['11','Journal','Personal Growth','Daily'],
    ['12','Something Fun / Creative','Fun & Adventure','As Needed'],
  ];

  habits.forEach((h, i) => {
    const row = 4 + i;
    const r = 3 + i; // 0-indexed
    const colF = `F${row}`;
    const colE = `E${row}`;
    const colG = `G${row}`;
    const colH = `H${row}`;
    // First 4 columns
    headerVals.push({ range: `Habit Tracker!A${row}:D${row}`, values: [h] });
    // E = COUNTIF(I{row}:AM{row}, TRUE)
    headerVals.push({ range: colE, values: [[`=COUNTIF(I${row}:AM${row},TRUE)`]] });
    // F = days in range so far (since Jun 1-15 filled, = 15)
    headerVals.push({ range: colF, values: [[15]] });
    // G = E/F %
    headerVals.push({ range: colG, values: [[`=IFERROR(E${row}/F${row},0)`]] });
    // H = sparkline trend (last 15 days filled)
    headerVals.push({ range: colH, values: [[`=SPARKLINE(I${row}:W${row},{"charttype","bar";"color1","#9B8EC4";"color2","#EDE9F5"})`]] });
  });

  // Summary rows
  headerVals.push({ range: 'Habit Tracker!A16', values: [['✦  DAILY TOTALS']] });
  headerVals.push({ range: 'Habit Tracker!A17', values: [['✦  COMPLETION %']] });

  // Day TOTALS row 16: COUNTIF across 12 habits for each day
  const totalFormulas = [];
  for (let d = 1; d <= 31; d++) {
    const col = colLetter(8 + (d - 1)); // I=col9=idx8
    totalFormulas.push(`=COUNTIF(${col}4:${col}15,TRUE)`);
  }
  headerVals.push({ range: 'Habit Tracker!I16:AM16', values: [totalFormulas] });

  // Day % row 17: total/12
  const pctFormulas = [];
  for (let d = 1; d <= 31; d++) {
    const col = colLetter(8 + (d - 1));
    pctFormulas.push(`=IFERROR(${col}16/12,0)`);
  }
  headerVals.push({ range: 'Habit Tracker!I17:AM17', values: [pctFormulas] });

  await valuesBatchUpdate(id, headerVals, 'habit-headers');

  // ── Checkbox validation for day cells I4:AM15 ────────────────────────────
  const checkReqs = [];
  checkReqs.push({
    setDataValidation: {
      range: gridRange(SID, 3, 15, 8, 39),
      rule: {
        condition: { type: 'BOOLEAN' },
        strict: true,
        showCustomUi: true,
      },
    },
  });

  await batchUpdate(id, checkReqs, 'habit-checkboxes');

  // ── Pre-fill June 1–15 data ───────────────────────────────────────────────
  // Habit data: rows 4-15, days 1-15 (cols I-W = indices 8-22)
  // true = checked, false = unchecked. Each inner array is one habit's 15-day data.
  const habitData = [
    // Meditation - mostly true
    [true,true,true,true,false,true,true,true,true,true,false,true,true,true,true],
    // Daily run 5x/week
    [true,true,false,true,true,false,false,true,true,false,true,true,false,false,true],
    // Drink 2L water
    [true,false,true,true,true,true,true,false,true,true,true,true,true,false,true],
    // Work on course
    [true,true,true,false,false,true,true,true,true,true,false,false,true,true,true],
    // LinkedIn
    [false,false,true,false,false,true,false,false,false,true,false,false,true,false,false],
    // No unnecessary spend
    [true,true,false,true,true,true,true,true,true,false,true,true,true,true,true],
    // Track spending
    [true,true,true,true,true,true,false,true,true,true,true,true,true,false,true],
    // Gratitude to partner
    [true,true,true,false,true,true,true,true,false,true,true,true,true,true,true],
    // Read 30 min
    [true,false,true,true,true,true,true,false,true,true,false,true,true,true,true],
    // Spanish practice
    [true,true,true,true,true,true,true,true,true,true,true,true,true,true,false],
    // Journal
    [true,false,true,true,false,true,true,true,false,true,true,false,true,true,true],
    // Fun/creative
    [false,true,false,false,false,true,false,false,true,false,false,true,false,false,true],
  ];

  const checkVals = [];
  habitData.forEach((row, hi) => {
    const sheetRow = 4 + hi; // rows 4-15
    // Each day's value
    const dayRange = `Habit Tracker!I${sheetRow}:W${sheetRow}`;
    checkVals.push({ range: dayRange, values: [row] });
  });

  await valuesBatchUpdate(id, checkVals, 'habit-data');

  // % format for col G (rows 4-15)
  const fmtReqs = [];
  for (let r = 3; r <= 14; r++) {
    fmtReqs.push({ repeatCell: { range: gridRange(SID, r, r + 1, 6, 7), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });
  }
  await batchUpdate(id, fmtReqs, 'habit-pct-fmt');

  console.log('Habit Tracker tab complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
