'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = sheetMap['Monthly Milestones']; // 3

function sectionHdr(row, label, bg, textColor) {
  return [
    { mergeCells: { range: gridRange(SID, row, row + 1, 0, 12), mergeType: 'MERGE_ALL' } },
    {
      repeatCell: {
        range: gridRange(SID, row, row + 1, 0, 12),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(textColor || C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    },
  ];
}

(async () => {
  const reqs = [];

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [32, 220, 110, 110, 110, 130, 10, 200, 160, 10, 160, 160];
  colWidths.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 12),
      cell: { userEnteredFormat: { backgroundColor: hex(C.blush), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Row 2: Month/Year banner ───────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, 12),
      cell: { userEnteredFormat: { backgroundColor: hex(C.paleBlush), textFormat: { foregroundColor: hex(C.deepPurple), bold: false, fontSize: 11, italic: true }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Section 1: Monthly Goals (rows 2-11) ──────────────────────────────────
  // Header row 3
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push(...sectionHdr(2, '  ❤  JUNE GOALS', C.blush));

  // Column headers row 4
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 3, 4, 0, 6),
      cell: { userEnteredFormat: { backgroundColor: hex(C.deepPurple), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Goal rows 5-10 (6 goals)
  for (let r = 4; r <= 9; r++) {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r + 1 }, properties: { pixelSize: 46 }, fields: 'pixelSize' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 0, 6),
        cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.paleBlush : C.white), wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE', textFormat: { fontSize: 9 } } },
        fields: 'userEnteredFormat(backgroundColor,wrapStrategy,verticalAlignment,textFormat)',
      },
    });
  }

  // ── Section 2: Weekly Priorities (rows 11-20) ─────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push(...sectionHdr(10, '  📅  WEEKLY PRIORITIES', C.skyBlue));

  // Week headers (row 12)
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 11, endIndex: 12 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 11, 12, 0, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.deepPurple), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // 4 weeks × 3 priority rows = rows 12-23
  const weekBgs = [C.paleSky, C.paleSky, C.paleSky, C.paleSage, C.paleSage, C.paleSage, C.paleBlush, C.paleBlush, C.paleBlush, C.paleLavender, C.paleLavender, C.paleLavender];
  for (let r = 12; r <= 23; r++) {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r + 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 0, 5),
        cell: { userEnteredFormat: { backgroundColor: hex(weekBgs[r - 12]), wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE', textFormat: { fontSize: 9 } } },
        fields: 'userEnteredFormat(backgroundColor,wrapStrategy,verticalAlignment,textFormat)',
      },
    });
  }

  // Week label col A merged across 3 rows each
  [[12,15,'Week 1\nJun 1–7', C.skyBlue],[15,18,'Week 2\nJun 8–14', C.skyBlue],[18,21,'Week 3\nJun 15–21', C.blush],[21,24,'Week 4\nJun 22–30', C.blush]].forEach(([r0, r1, label, c]) => {
    reqs.push({ mergeCells: { range: gridRange(SID, r0, r1, 0, 1), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r0, r1, 0, 1),
        cell: { userEnteredFormat: { backgroundColor: hex(c), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
      },
    });
  });

  // ── Section 3: Habit Focus (rows 24-32) ───────────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 24, endIndex: 25 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push(...sectionHdr(24, '  🌱  HABIT FOCUS', C.dustySage));

  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 25, endIndex: 26 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 25, 26, 0, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.deepPurple), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  for (let r = 26; r <= 32; r++) {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r + 1 }, properties: { pixelSize: 42 }, fields: 'pixelSize' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 0, 5),
        cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.paleSage : C.white), wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE', textFormat: { fontSize: 9 } } },
        fields: 'userEnteredFormat(backgroundColor,wrapStrategy,verticalAlignment,textFormat)',
      },
    });
  }

  // ── Section 4: Monthly Reflection (rows 33-40) ────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 33, endIndex: 34 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push(...sectionHdr(33, '  ✍  MONTH-END REFLECTION', C.softGold, C.deepPurple));

  for (let r = 34; r <= 39; r++) {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r + 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 0, 12),
        cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.paleGold : C.white), wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE', textFormat: { fontSize: 10 } } },
        fields: 'userEnteredFormat(backgroundColor,wrapStrategy,verticalAlignment,textFormat)',
      },
    });
    // Prompt col A bold
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 0, 1),
        cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.deepPurple) }, wrapStrategy: 'WRAP' } },
        fields: 'userEnteredFormat(textFormat,wrapStrategy)',
      },
    });
  }

  // ── Separator/chart area (col G-L) → monthly progress chart helper ────────
  // Chart helper at A37:B42 — use rows 41-47 col A-B for month/progress series
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 40, endIndex: 48 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 40, 48, 0, 2),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderWhite), textFormat: { fontSize: 9 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  await batchUpdate(id, reqs, 'monthly-format');

  // ── Values ────────────────────────────────────────────────────────────────
  const vals = [
    { range: 'Monthly Milestones!A1', values: [['✦  June 2025 Monthly Milestones  ✦']] },
    { range: 'Monthly Milestones!A2', values: [['Month of FLOURISH  |  Q2: MOMENTUM  |  Jordan — Year of 2025']] },

    // Goals section headers
    { range: 'Monthly Milestones!A4', values: [['#']] },
    { range: 'Monthly Milestones!B4', values: [['June Goal']] },
    { range: 'Monthly Milestones!C4', values: [['Life Area']] },
    { range: 'Monthly Milestones!D4', values: [['Status']] },
    { range: 'Monthly Milestones!E4', values: [['% Done']] },
    { range: 'Monthly Milestones!F4', values: [['Notes / Progress']] },

    // 6 June goals
    { range: 'Monthly Milestones!A5:F5', values: [['1', 'Complete 30 meditation sessions in June', 'Health & Fitness', 'In Progress', 0.5, '15/30 days done — on track!']] },
    { range: 'Monthly Milestones!A6:F6', values: [['2', 'Run 3× per week — total 60km in June', 'Health & Fitness', 'In Progress', 0.45, '27km logged, 3 more weeks']] },
    { range: 'Monthly Milestones!A7:F7', values: [['3', 'Record 3 course modules & send to beta', 'Career & Purpose', 'In Progress', 0.67, '2 modules done, 1 left']] },
    { range: 'Monthly Milestones!A8:F8', values: [['4', 'Save $1,500 towards emergency fund', 'Finance & Wealth', 'In Progress', 0.6, '$900 saved so far']] },
    { range: 'Monthly Milestones!A9:F9', values: [['5', 'Book family trip abroad (flights + hotel)', 'Relationships', 'Complete', 1.0, 'Booked! Barcelona Aug 10–17 🎉']] },
    { range: 'Monthly Milestones!A10:F10', values: [['6', 'Read 2 books', 'Personal Growth', 'Complete', 1.0, 'Atomic Habits + The Alchemist ✓']] },

    // % format handled in pct section
    // Weekly priorities
    { range: 'Monthly Milestones!B12', values: [['Priority 1']] },
    { range: 'Monthly Milestones!C12', values: [['Priority 2']] },
    { range: 'Monthly Milestones!D12', values: [['Priority 3']] },
    { range: 'Monthly Milestones!E12', values: [['Big Win']] },

    // Week 1 (rows 13-15, col A merged)
    { range: 'Monthly Milestones!A13', values: [['Week 1\nJun 1–7']] },
    { range: 'Monthly Milestones!B13', values: [['Run 3× + gym']] },
    { range: 'Monthly Milestones!C13', values: [['Draft course module 1']] },
    { range: 'Monthly Milestones!D13', values: [['$500 to savings']] },
    { range: 'Monthly Milestones!E13', values: [['Submitted module 1!']] },
    { range: 'Monthly Milestones!B14', values: [['Meditate daily']] },
    { range: 'Monthly Milestones!C14', values: [['LinkedIn post × 2']] },
    { range: 'Monthly Milestones!D14', values: [['Review budget']] },
    { range: 'Monthly Milestones!E14', values: [['14-day streak hit']] },
    { range: 'Monthly Milestones!B15', values: [['Sleep by 10:30pm']] },
    { range: 'Monthly Milestones!C15', values: [['Reach out to mentor']] },
    { range: 'Monthly Milestones!D15', values: [['Cut dining budget']] },
    { range: 'Monthly Milestones!E15', values: [['Mentor call booked']] },

    // Week 2
    { range: 'Monthly Milestones!A16', values: [['Week 2\nJun 8–14']] },
    { range: 'Monthly Milestones!B16', values: [['8-mile long run']] },
    { range: 'Monthly Milestones!C16', values: [['Record module 2']] },
    { range: 'Monthly Milestones!D16', values: [['Invest $500']] },
    { range: 'Monthly Milestones!E16', values: [['Longest run ever! 8mi']] },
    { range: 'Monthly Milestones!B17', values: [['Meditate + journal']] },
    { range: 'Monthly Milestones!C17', values: [['Perf. review prep']] },
    { range: 'Monthly Milestones!D17', values: [['Check net worth']] },
    { range: 'Monthly Milestones!E17', values: [['Director 1:1 positive']] },
    { range: 'Monthly Milestones!B18', values: [['Book trip deposits']] },
    { range: 'Monthly Milestones!C18', values: [['Spanish lesson 2×']] },
    { range: 'Monthly Milestones!D18', values: [['Trip budget locked']] },
    { range: 'Monthly Milestones!E18', values: [['TRIP BOOKED! 🎉']] },

    // Week 3
    { range: 'Monthly Milestones!A19', values: [['Week 3\nJun 15–21']] },
    { range: 'Monthly Milestones!B19', values: [['Beach run + swim']] },
    { range: 'Monthly Milestones!C19', values: [['Beta course review']] },
    { range: 'Monthly Milestones!D19', values: [['$400 debt payment']] },
    { range: 'Monthly Milestones!E19', values: [['Course beta sent!']] },
    { range: 'Monthly Milestones!B20', values: [['Yoga 3×']] },
    { range: 'Monthly Milestones!C20', values: [['Feedback reviews']] },
    { range: 'Monthly Milestones!D20', values: [['Side income ideas']] },
    { range: 'Monthly Milestones!E20', values: [['5 beta signups!']] },
    { range: 'Monthly Milestones!B21', values: [['Self-care Sunday']] },
    { range: 'Monthly Milestones!C21', values: [['Spanish practice']] },
    { range: 'Monthly Milestones!D21', values: [['Frugal week']] },
    { range: 'Monthly Milestones!E21', values: [['Most relaxed week']] },

    // Week 4
    { range: 'Monthly Milestones!A22', values: [['Week 4\nJun 22–30']] },
    { range: 'Monthly Milestones!B22', values: [['10-mile run attempt']] },
    { range: 'Monthly Milestones!C22', values: [['Promo doc submit']] },
    { range: 'Monthly Milestones!D22', values: [['June savings tally']] },
    { range: 'Monthly Milestones!E22', values: [['10-mile run success!']] },
    { range: 'Monthly Milestones!B23', values: [['Monthly reflection']] },
    { range: 'Monthly Milestones!C23', values: [['Q2 review']] },
    { range: 'Monthly Milestones!D23', values: [['Q3 budget plan']] },
    { range: 'Monthly Milestones!E23', values: [['Promo doc in!']] },
    { range: 'Monthly Milestones!B24', values: [['Celebration dinner']] },
    { range: 'Monthly Milestones!C24', values: [['Set Q3 goals']] },
    { range: 'Monthly Milestones!D24', values: [['Q3 financial goals']] },
    { range: 'Monthly Milestones!E24', values: [['Best month ever!']] },

    // Habit Focus
    { range: 'Monthly Milestones!A26', values: [['Habit']] },
    { range: 'Monthly Milestones!B26', values: [['Daily Target']] },
    { range: 'Monthly Milestones!C26', values: [['Jun 1–15 Avg']] },
    { range: 'Monthly Milestones!D26', values: [['Status']] },
    { range: 'Monthly Milestones!E26', values: [['Notes']] },

    { range: 'Monthly Milestones!A27:E27', values: [['Morning meditation', '20 min', '18/15 ✓', 'In Progress', 'Averaging 18 min — exceeding goal!']] },
    { range: 'Monthly Milestones!A28:E28', values: [['Daily run or workout', '45 min', '3×/wk', 'In Progress', '3 runs + 2 strength sessions wk 1']] },
    { range: 'Monthly Milestones!A29:E29', values: [['Read before bed', '30 min', '22/15 ✓', 'In Progress', 'Averaging 22 min — loving it']] },
    { range: 'Monthly Milestones!A30:E30', values: [['Duolingo Spanish', '15 min', '15/15 ✓', 'In Progress', 'Perfect streak first 15 days!']] },
    { range: 'Monthly Milestones!A31:E31', values: [['Journal', '10 min', '12/15', 'In Progress', 'Missed 3 days — adjust to morning']] },
    { range: 'Monthly Milestones!A32:E32', values: [['No-phone first hour', '60 min', '10/15', 'In Progress', 'Hard but improving — 67%']] },
    { range: 'Monthly Milestones!A33:E33', values: [['Meal prep Sunday', 'Weekly', '3/3 wks', 'In Progress', 'On track — prepped every Sunday']] },

    // Reflection
    { range: 'Monthly Milestones!A35', values: [['🏆 Top 3 June Wins:']] },
    { range: 'Monthly Milestones!B35', values: [['1. Booked Barcelona trip! 2. 10-mile run! 3. Course beta launched!']] },
    { range: 'Monthly Milestones!A36', values: [['💭 What challenged me?']] },
    { range: 'Monthly Milestones!B36', values: [['Energy dips mid-month. Need better sleep hygiene.']] },
    { range: 'Monthly Milestones!A37', values: [['📚 What did I learn?']] },
    { range: 'Monthly Milestones!B37', values: [['Celebrating small wins keeps me going. Reward loops work!']] },
    { range: 'Monthly Milestones!A38', values: [['💡 What will I do differently in July?']] },
    { range: 'Monthly Milestones!B38', values: [['Protect Sundays for recovery. Add 1 non-screen evening per week.']] },
    { range: 'Monthly Milestones!A39', values: [['🎯 Carry-forward goals to July:']] },
    { range: 'Monthly Milestones!B39', values: [['Continue 10-mile+ runs, hit $12K emergency fund, record 2 more course modules']] },
    { range: 'Monthly Milestones!A40', values: [['⭐ June score (1–10):']] },
    { range: 'Monthly Milestones!B40', values: [['9.5 — FLOURISHED! 🌟']] },

    // Chart helper for monthly progress bar
    { range: 'Monthly Milestones!A42', values: [['Goal']] },
    { range: 'Monthly Milestones!B42', values: [['Progress']] },
    { range: 'Monthly Milestones!A43:B48', values: [
      ['Meditation', 0.5], ['Running', 0.45], ['Course', 0.67], ['Savings', 0.6], ['Family Trip', 1.0], ['Reading', 1.0],
    ]},
  ];

  await valuesBatchUpdate(id, vals, 'monthly-values');

  // % format for E5:E10 and B43:B48
  const pctReqs = [];
  for (let r = 4; r <= 9; r++) {
    pctReqs.push({ repeatCell: { range: gridRange(SID, r, r + 1, 4, 5), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });
  }
  for (let r = 42; r <= 47; r++) {
    pctReqs.push({ repeatCell: { range: gridRange(SID, r, r + 1, 1, 2), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });
  }
  await batchUpdate(id, pctReqs, 'monthly-pct');

  console.log('Monthly Milestones tab complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
