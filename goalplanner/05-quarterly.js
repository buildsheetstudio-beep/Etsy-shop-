'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = sheetMap['Quarterly Planner']; // 2

(async () => {
  const reqs = [];

  // ── Column widths ─────────────────────────────────────────────────────────
  // A=week#, B=dates, C=theme/focus, D=health, E=career, F=finance, G=relationships, H=personal, I=fun, J=big win, K=reflection, L=chart helper
  const colWidths = [40, 130, 150, 120, 120, 120, 120, 120, 100, 160, 160, 10, 120, 120];
  colWidths.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // ── Row heights ────────────────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  // Weekly rows (rows 4-16 = 13 weeks)
  for (let r = 4; r <= 16; r++) {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r + 1 }, properties: { pixelSize: 60 }, fields: 'pixelSize' } });
  }
  // Month label rows
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 17, endIndex: 18 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 18, endIndex: 19 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 11),
      cell: { userEnteredFormat: { backgroundColor: hex(C.skyBlue), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Row 2: Subtitle ───────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, 11),
      cell: { userEnteredFormat: { backgroundColor: hex(C.paleSky), textFormat: { foregroundColor: hex(C.deepPurple), bold: false, fontSize: 11, italic: true }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Row 3: Q2 Summary Inputs ──────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 2, 3, 0, 11),
      cell: { userEnteredFormat: { backgroundColor: hex(C.deepPurple), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Row 4: Column headers ─────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 3, 4, 0, 11),
      cell: { userEnteredFormat: { backgroundColor: hex(C.medLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });

  // ── Month dividers ────────────────────────────────────────────────────────
  // April header before week 1 (row 4 = already header), we insert month label at appropriate spot
  // April = weeks 1-4 (rows 4-7), May = weeks 5-8 (rows 8-11), June = weeks 9-13 (rows 12-16)
  const monthLabels = [
    { row: 4, label: 'APRIL', color: C.skyBlue },
    { row: 8, label: 'MAY', color: C.dustySage },
    { row: 12, label: 'JUNE', color: C.blush },
  ];

  // Actually we won't insert month divider rows — we'll use background color to differentiate
  // April weeks: paleSky bg, May weeks: paleSage bg, June weeks: paleBlush bg
  const weekBgs = [
    C.paleSky, C.paleSky, C.paleSky, C.paleSky,   // weeks 1-4 April
    C.paleSage, C.paleSage, C.paleSage, C.paleSage, // weeks 5-8 May
    C.paleBlush, C.paleBlush, C.paleBlush, C.paleBlush, C.paleBlush, // weeks 9-13 June
  ];

  for (let r = 4; r <= 16; r++) {
    const bg = weekBgs[r - 4];
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 0, 11),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP', textFormat: { fontSize: 9 } } },
        fields: 'userEnteredFormat(backgroundColor,verticalAlignment,wrapStrategy,textFormat)',
      },
    });
    // Week # cell - bold center
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 0, 1),
        cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER' } },
        fields: 'userEnteredFormat(textFormat,horizontalAlignment)',
      },
    });
    // Date cell - slightly muted
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 1, 2),
        cell: { userEnteredFormat: { textFormat: { fontSize: 9, foregroundColor: hex(C.mutedText) }, horizontalAlignment: 'CENTER' } },
        fields: 'userEnteredFormat(textFormat,horizontalAlignment)',
      },
    });
    // Focus column - italic
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 2, 3),
        cell: { userEnteredFormat: { textFormat: { italic: true, bold: true, fontSize: 9 }, horizontalAlignment: 'LEFT' } },
        fields: 'userEnteredFormat(textFormat,horizontalAlignment)',
      },
    });
  }

  // Month color bands on col A (accent stripe)
  [4,5,6,7].forEach(r => reqs.push({ repeatCell: { range: gridRange(SID, r, r+1, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(C.skyBlue), textFormat: { foregroundColor: hex(C.white), bold: true } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } }));
  [8,9,10,11].forEach(r => reqs.push({ repeatCell: { range: gridRange(SID, r, r+1, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(C.dustySage), textFormat: { foregroundColor: hex(C.white), bold: true } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } }));
  [12,13,14,15,16].forEach(r => reqs.push({ repeatCell: { range: gridRange(SID, r, r+1, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(C.blush), textFormat: { foregroundColor: hex(C.white), bold: true } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } }));

  // ── Row 18: Q2 Reflection section header ─────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 17, 18, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 17, 18, 0, 11), cell: { userEnteredFormat: { backgroundColor: hex(C.deepPurple), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // Reflection rows (19-23)
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 17, endIndex: 24 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
  for (let r = 18; r <= 22; r++) {
    reqs.push({ repeatCell: { range: gridRange(SID, r, r+1, 0, 11), cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.lavenderWhite : C.paleLavender), wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE', textFormat: { fontSize: 10 } } }, fields: 'userEnteredFormat(backgroundColor,wrapStrategy,verticalAlignment,textFormat)' } });
    reqs.push({ repeatCell: { range: gridRange(SID, r, r+1, 0, 1), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.deepPurple) }, wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(textFormat,wrapStrategy)' } });
  }

  // ── Chart helper (L col, rows 12-24) — line chart: week vs completion ────
  reqs.push({ repeatCell: { range: gridRange(SID, 11, 29, 12, 14), cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderWhite), textFormat: { fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  await batchUpdate(id, reqs, 'quarterly-format');

  // ── Values ────────────────────────────────────────────────────────────────
  const weeks = [
    // [week#, dates, focus, health, career, finance, relationships, personal, fun, big_win, reflection]
    ['1','Apr 1–6','Build momentum','3 runs + gym','Draft promo plan','Review budget','Date night','Read Ch.1-3','New café brunch','Booked half marathon race','Felt energized'],
    ['2','Apr 7–13','Consistency','5 runs','LinkedIn post ×3','$500 to savings','Phone call mom','Duolingo 7 days','Art gallery visit','20-day meditation streak','Best week yet'],
    ['3','Apr 14–20','Deep work','Rest & recover','Course outline draft','Debt payment $500','Park walk w/partner','Journal daily','Hiking trail new','Finished course outline','Stayed focused'],
    ['4','Apr 21–27','Wrap April','Yoga 4×','Mentor coffee chat','Review Q2 goals','Friends game night','Read book #5','Live music show','Paid off $1K debt','April strongest month'],
    ['5','Apr 28–May 4','Skill up','Bike 3×','LinkedIn article','Open index acct.','Sunday brunch date','Spanish class week 1','New cuisine: Thai','Opened investment account','Excited about Spanish'],
    ['6','May 5–11','Momentum','Gym 4×','Course recording x2','Auto-invest set','Anniversary dinner','Read book #6','Pottery class visit','First course module recorded','Feeling creative'],
    ['7','May 12–18','Stretch week','10K run!','Course recording x3','$600 to savings','Reach out old friend','Journal + meditate','Farmers market','Best run distance yet','Exhausted but proud'],
    ['8','May 19–25','Recovery','Stretch & walk','Q2 review prep','Emergency fund check','Video call 3 friends','Book #7 finished','Staycation day','Emergency fund 70%!','Needed the rest'],
    ['9','May 26–Jun 1','Ignite June','Gym 5×','Job perf. prep','Budget audit','Date night picnic','Duolingo 14-day streak','New restaurant: Peruvian','2.5K LinkedIn followers','June is MOMENTUM month'],
    ['10','Jun 2–8','Mid-quarter push','Run 8mi','1:1 with director','Invest $500','Weekend trip planning','Spanish grammar deep dive','Comedy show','First 8mi run complete','Clear mind after comedy'],
    ['11','Jun 9–15','Reflection & adjust','Yoga 5×','Peer feedback given','Review goals vs plan','Trip deposits booked','Book #8 done','Food truck festival','Booked family trip!','So grateful'],
    ['12','Jun 16–22','Final push','Swim 3×','Course beta review','Savings milestone','Partner beach day','Affirmations daily','New recipe night','Course ready for beta!','Beaches > everything'],
    ['13','Jun 23–30','Q2 close + celebrate','Rest week','Q2 promo doc submitted','Q2 financial review','Family dinner','Q2 reflection journal','Celebration dinner out','Promo doc submitted!!','Best quarter of my life'],
  ];

  const weekVals = [];
  weeks.forEach((w, i) => {
    weekVals.push({ range: `Quarterly Planner!A${5+i}:K${5+i}`, values: [w] });
  });

  const mainVals = [
    { range: 'Quarterly Planner!A1', values: [['✦  Q2 Quarterly Planner — Apr–Jun 2025  ✦']] },
    { range: 'Quarterly Planner!A2', values: [['Theme: MOMENTUM  |  Focus: Build, create, connect']] },
    { range: 'Quarterly Planner!A3', values: [['Quarter', 'Dates', 'Theme', '❤ Health', '💼 Career', '💰 Finance', '💞 Relationships', '🌱 Personal', '🎉 Fun', '🏆 Big Win', '✍ Reflection']] },
    { range: 'Quarterly Planner!A4', values: [['Wk', 'Dates', 'Weekly Focus', 'Health', 'Career', 'Finance', 'Relationships', 'Personal', 'Fun', 'Big Win of the Week', 'End-of-Week Reflection']] },

    // Q2 Reflection
    { range: 'Quarterly Planner!A18', values: [['✦  Q2 END-OF-QUARTER REFLECTION  ✦']] },
    { range: 'Quarterly Planner!A19:B19', values: [['🏆 3 Biggest Q2 Wins:']] },
    { range: 'Quarterly Planner!C19', values: [['Booked family trip abroad, submitted promo doc, launched course beta!']] },
    { range: 'Quarterly Planner!A20:B20', values: [['💭 What held me back?']] },
    { range: 'Quarterly Planner!C20', values: [['Over-scheduled in May — learned to protect rest days']] },
    { range: 'Quarterly Planner!A21:B21', values: [['📚 Biggest lessons learned:']] },
    { range: 'Quarterly Planner!C21', values: [['Consistency beats intensity. Small daily actions compound massively.']] },
    { range: 'Quarterly Planner!A22:B22', values: [['🎯 What to carry into Q3?']] },
    { range: 'Quarterly Planner!C22', values: [['Keep the course momentum, maintain date nights, double down on savings']] },
    { range: 'Quarterly Planner!A23:B23', values: [['⭐ Q2 overall score (1-10):']] },
    { range: 'Quarterly Planner!C23', values: [['9 — MOMENTUM was real!']] },

    // Chart helper: week completion % (estimated)
    { range: 'Quarterly Planner!M12', values: [['Week']] },
    { range: 'Quarterly Planner!N12', values: [['Completion %']] },
    { range: 'Quarterly Planner!M13:N25', values: [
      [1, 0.72], [2, 0.85], [3, 0.78], [4, 0.82],
      [5, 0.68], [6, 0.80], [7, 0.90], [8, 0.75],
      [9, 0.88], [10, 0.83], [11, 0.91], [12, 0.86], [13, 0.95],
    ]},
  ];

  await valuesBatchUpdate(id, [...mainVals, ...weekVals], 'quarterly-values');

  console.log('Quarterly Planner tab complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
