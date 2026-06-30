'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = sheetMap['Life Dashboard']; // 0

(async () => {
  const reqs = [];

  // ── Column widths ──
  const widths = [[0,160],[1,130],[2,130],[3,80],[4,180],[5,220],[6,30],[7,30],[8,30],[9,30],[10,30],[11,30]];
  widths.forEach(([ci,px]) => reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: px }, fields: 'pixelSize' } }));

  // ── Row heights ──
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 4 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 6 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 16 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 16, endIndex: 17 }, properties: { pixelSize: 6 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 17, endIndex: 25 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 25, endIndex: 34 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 26, endIndex: 31 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });

  // ── Row 1: Title ──
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, 12), cell: { userEnteredValue: { stringValue: '✨ MY ULTIMATE GOAL PLANNER' }, userEnteredFormat: { backgroundColor: hex(C.dustyLavender), textFormat: { foregroundColor: hex(C.paleGoldText), bold: true, fontSize: 24 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // ── Row 2: Subtitle ──
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, 12), cell: { userEnteredValue: { stringValue: 'Dream It. Plan It. Live It.' }, userEnteredFormat: { backgroundColor: hex(C.medLavender), textFormat: { foregroundColor: hex(C.white), italic: true, fontSize: 12 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // ── Row 3: Identity inputs — colored cells ──
  // A3: label
  reqs.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(C.bodyText), textFormat: { foregroundColor: hex(C.paleGoldText), bold: true, fontSize: 10 }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
  // B3: year input
  reqs.push({ repeatCell: { range: gridRange(SID, 2, 3, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { foregroundColor: hex(C.bodyText), bold: true, fontSize: 12 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  // D3: My Name label
  reqs.push({ repeatCell: { range: gridRange(SID, 2, 3, 3, 4), cell: { userEnteredFormat: { backgroundColor: hex(C.veryPaleLav), textFormat: { foregroundColor: hex(C.mutedText), bold: true, fontSize: 10 }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
  // E3: name input
  reqs.push({ repeatCell: { range: gridRange(SID, 2, 3, 4, 5), cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { foregroundColor: hex(C.bodyText), bold: true, fontSize: 12 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  // G3: Word of year label
  reqs.push({ repeatCell: { range: gridRange(SID, 2, 3, 6, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.paleBlush), textFormat: { foregroundColor: hex(C.mutedText), bold: true, fontSize: 10 }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
  // H3: word input
  reqs.push({ repeatCell: { range: gridRange(SID, 2, 3, 7, 8), cell: { userEnteredFormat: { backgroundColor: hex(C.blush), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  // J3: Quarterly theme label
  reqs.push({ repeatCell: { range: gridRange(SID, 2, 3, 9, 10), cell: { userEnteredFormat: { backgroundColor: hex(C.veryPaleLav), textFormat: { foregroundColor: hex(C.mutedText), bold: true, fontSize: 10 }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
  // K3: Q theme value
  reqs.push({ repeatCell: { range: gridRange(SID, 2, 3, 10, 11), cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { foregroundColor: hex(C.deepPurple), bold: true, fontSize: 11 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // ── Row 4: Identity banner ──
  reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 3, 4, 0, 12), cell: { userEnteredFormat: { backgroundColor: hex(C.veryPaleLav), textFormat: { foregroundColor: hex(C.mutedText), italic: true, fontSize: 11 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // ── Row 5: Spacer ──
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 12 }, fields: 'pixelSize' } });

  // ── Row 6: Wheel section header ──
  reqs.push({ mergeCells: { range: gridRange(SID, 5, 6, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 5, 6, 0, 12), cell: { userEnteredValue: { stringValue: '🌸 WHEEL OF LIFE — HOW FULL IS YOUR LIFE?' }, userEnteredFormat: { backgroundColor: hex(C.veryPaleLav), textFormat: { foregroundColor: hex(C.dustyLavender), bold: true, fontSize: 13 }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // ── Row 7: Intro text ──
  reqs.push({ mergeCells: { range: gridRange(SID, 6, 7, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 6, 7, 0, 12), cell: { userEnteredValue: { stringValue: 'Rate each life area 1–10. 10 = thriving, 1 = needs urgent attention.' }, userEnteredFormat: { textFormat: { foregroundColor: hex(C.mutedText), italic: true, fontSize: 9 }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredValue,userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)' } });

  // ── Row 8: Column headers ──
  reqs.push({ repeatCell: { range: gridRange(SID, 7, 8, 0, 6), cell: { userEnteredFormat: { backgroundColor: hex(C.dustyLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // ── Rows 9-14: Life area data rows ──
  const lifeBgs = [C.paleSage, C.paleSky, C.paleGold, C.paleBlush, C.paleLavender, '#FFEDD5'];
  for (let r = 8; r < 14; r++) {
    reqs.push({ repeatCell: { range: gridRange(SID, r, r+1, 0, 6), cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.white : C.lavenderWhite), textFormat: { fontSize: 10 }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
    // B col (Current score) — pale lavender input
    reqs.push({ repeatCell: { range: gridRange(SID, r, r+1, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { bold: true, fontSize: 12 }, horizontalAlignment: 'CENTER' }, userEnteredFormat: { backgroundColor: hex(C.paleLavender), horizontalAlignment: 'CENTER', textFormat: { bold: true, fontSize: 12 } } }, fields: 'userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)' } });
    // C col (Target score)
    reqs.push({ repeatCell: { range: gridRange(SID, r, r+1, 2, 3), cell: { userEnteredFormat: { backgroundColor: hex(C.veryPaleLav), horizontalAlignment: 'CENTER', textFormat: { fontSize: 11 } } }, fields: 'userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)' } });
    // D col (Gap) — centered
    reqs.push({ repeatCell: { range: gridRange(SID, r, r+1, 3, 4), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true, fontSize: 11 } } }, fields: 'userEnteredFormat(horizontalAlignment,textFormat)' } });
    // F col (Key focus) — italic
    reqs.push({ repeatCell: { range: gridRange(SID, r, r+1, 5, 6), cell: { userEnteredFormat: { textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.mutedText) } } }, fields: 'userEnteredFormat.textFormat' } });
  }

  // ── Row 15: Average row ──
  reqs.push({ repeatCell: { range: gridRange(SID, 14, 15, 0, 6), cell: { userEnteredFormat: { backgroundColor: hex(C.medLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // ── Row 16: Divider ──
  reqs.push({ mergeCells: { range: gridRange(SID, 15, 16, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 15, 16, 0, 12), cell: { userEnteredFormat: { backgroundColor: hex(C.dustyLavender) } }, fields: 'userEnteredFormat.backgroundColor' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 15, endIndex: 16 }, properties: { pixelSize: 4 }, fields: 'pixelSize' } });

  // ── Row 17: Top Goals header ──
  reqs.push({ mergeCells: { range: gridRange(SID, 16, 17, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 16, 17, 0, 12), cell: { userEnteredValue: { stringValue: '🎯 MY TOP GOALS FOR 2025' }, userEnteredFormat: { backgroundColor: hex(C.veryPaleLav), textFormat: { foregroundColor: hex(C.dustyLavender), bold: true, fontSize: 12 }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // ── Row 18: Top Goals col headers ──
  reqs.push({ repeatCell: { range: gridRange(SID, 17, 18, 0, 8), cell: { userEnteredFormat: { backgroundColor: hex(C.dustyLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // ── Rows 19-24: Top goal rows ──
  for (let r = 18; r < 24; r++) {
    reqs.push({ repeatCell: { range: gridRange(SID, r, r+1, 0, 8), cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.white : C.lavenderWhite), textFormat: { fontSize: 10 }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
  }
  // Status col (E=4) — centered
  reqs.push({ repeatCell: { range: gridRange(SID, 18, 24, 4, 5), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true, fontSize: 9 } } }, fields: 'userEnteredFormat(horizontalAlignment,textFormat)' } });
  // % col (F=5) — centered
  reqs.push({ repeatCell: { range: gridRange(SID, 18, 24, 5, 6), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true } } }, fields: 'userEnteredFormat(horizontalAlignment,textFormat)' } });
  // # col (A=0) — centered small
  reqs.push({ repeatCell: { range: gridRange(SID, 18, 24, 0, 1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true, foregroundColor: hex(C.mutedText) } } }, fields: 'userEnteredFormat(horizontalAlignment,textFormat)' } });

  // ── Row 25: Spacer ──
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 24, endIndex: 25 }, properties: { pixelSize: 12 }, fields: 'pixelSize' } });

  // ── Rows 26-34: Motivation & KPI section ──
  // A26:F26 — Quote header
  reqs.push({ mergeCells: { range: gridRange(SID, 25, 26, 0, 6), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 25, 26, 0, 6), cell: { userEnteredValue: { stringValue: '💭 QUOTE OF THE YEAR' }, userEnteredFormat: { backgroundColor: hex(C.veryPaleLav), textFormat: { foregroundColor: hex(C.dustyLavender), bold: true, fontSize: 11 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  // G26:L26 — Year at a glance header
  reqs.push({ mergeCells: { range: gridRange(SID, 25, 26, 6, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 25, 26, 6, 12), cell: { userEnteredValue: { stringValue: '🌟 YEAR AT A GLANCE' }, userEnteredFormat: { backgroundColor: hex(C.veryPaleLav), textFormat: { foregroundColor: hex(C.dustyLavender), bold: true, fontSize: 11 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  // A27:F30 — Quote text area
  reqs.push({ mergeCells: { range: gridRange(SID, 26, 31, 0, 6), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 26, 31, 0, 6), cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { foregroundColor: hex(C.dustyLavender), italic: true, fontSize: 11 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' } });
  // KPI cards: 2x2 grid in G27:L30
  // G27:H28 — Goals This Year
  reqs.push({ mergeCells: { range: gridRange(SID, 26, 28, 6, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 26, 28, 6, 8), cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { foregroundColor: hex(C.deepPurple), bold: true, fontSize: 16 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  // I27:J28 — Complete
  reqs.push({ mergeCells: { range: gridRange(SID, 26, 28, 8, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 26, 28, 8, 10), cell: { userEnteredFormat: { backgroundColor: hex(C.paleSage), textFormat: { foregroundColor: hex('#2D6A4F'), bold: true, fontSize: 16 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  // K27:L28 — Overall Progress
  reqs.push({ mergeCells: { range: gridRange(SID, 26, 28, 10, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 26, 28, 10, 12), cell: { userEnteredFormat: { backgroundColor: hex(C.paleSky), textFormat: { foregroundColor: hex('#1D4ED8'), bold: true, fontSize: 16 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  // G29:H30 — Habits Tracked
  reqs.push({ mergeCells: { range: gridRange(SID, 28, 31, 6, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 28, 31, 6, 9), cell: { userEnteredFormat: { backgroundColor: hex(C.paleBlush), textFormat: { foregroundColor: hex('#9B1B30'), bold: true, fontSize: 16 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  // I29:L30 — KPI labels
  reqs.push({ mergeCells: { range: gridRange(SID, 28, 31, 9, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 28, 31, 9, 12), cell: { userEnteredFormat: { backgroundColor: hex(C.paleGold), textFormat: { foregroundColor: hex('#92400E'), bold: true, fontSize: 14 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // ── Borders ──
  reqs.push({ updateBorders: { range: gridRange(SID, 7, 15, 0, 6), innerHorizontal: { style: 'SOLID', color: hex(C.border) }, innerVertical: { style: 'SOLID', color: hex(C.border) } } });
  reqs.push({ updateBorders: { range: gridRange(SID, 17, 24, 0, 8), innerHorizontal: { style: 'SOLID', color: hex(C.border) }, innerVertical: { style: 'SOLID', color: hex(C.border) } } });

  await batchUpdate(id, reqs, 'dashboard-format');

  // ── Values ──
  const wheel = [
    ['🌿 Health & Fitness', 6, 9, '=C9-B9', '=SPARKLINE({B9/10,1-B9/10},{"charttype","bar";"color1","#9B8EC4";"color2","#EDE9F5"})', 'Build consistent morning workout habit'],
    ['💼 Career & Purpose', 7, 9, '=C10-B10','=SPARKLINE({B10/10,1-B10/10},{"charttype","bar";"color1","#9B8EC4";"color2","#EDE9F5"})','Get promoted to senior level or launch side project'],
    ['💰 Finance & Wealth', 5, 8, '=C11-B11','=SPARKLINE({B11/10,1-B11/10},{"charttype","bar";"color1","#9B8EC4";"color2","#EDE9F5"})','Build 6-month emergency fund and invest consistently'],
    ['💕 Relationships',    8, 9, '=C12-B12','=SPARKLINE({B12/10,1-B12/10},{"charttype","bar";"color1","#9B8EC4";"color2","#EDE9F5"})','Deepen close friendships and be more present with family'],
    ['🌱 Personal Growth',  7,10, '=C13-B13','=SPARKLINE({B13/10,1-B13/10},{"charttype","bar";"color1","#9B8EC4";"color2","#EDE9F5"})','Read 24 books and finish online certification'],
    ['🎉 Fun & Adventure',  4, 8, '=C14-B14','=SPARKLINE({B14/10,1-B14/10},{"charttype","bar";"color1","#9B8EC4";"color2","#EDE9F5"})','Take 2 trips and try 12 new experiences this year'],
  ];

  const topGoals = [
    [1,'Run a half marathon (13.1 miles)','Health & Fitness','Jun 30, 2025','In Progress','45%','High'],
    [2,'Get promoted to Senior Role by Q3','Career & Purpose','Sep 30, 2025','In Progress','35%','High'],
    [3,'Build 6-month emergency fund ($15k)','Finance & Wealth','Dec 31, 2025','In Progress','40%','High'],
    [4,'Host monthly dinner parties','Relationships','Dec 31, 2025','In Progress','50%','Medium'],
    [5,'Read 24 books this year','Personal Growth','Dec 31, 2025','In Progress','62%','High'],
    [6,'Take 2 international trips','Fun & Adventure','Dec 31, 2025','In Progress','50%','Medium'],
  ];

  await valuesBatchUpdate(id, [
    { range: 'Life Dashboard!A3', values: [['Planning Year:']] },
    { range: 'Life Dashboard!B3', values: [[2025]] },
    { range: 'Life Dashboard!D3', values: [['My Name:']] },
    { range: 'Life Dashboard!E3', values: [['Jordan']] },
    { range: 'Life Dashboard!G3', values: [['Word of the Year:']] },
    { range: 'Life Dashboard!H3', values: [['FLOURISH']] },
    { range: 'Life Dashboard!J3', values: [['Quarterly Theme:']] },
    { range: 'Life Dashboard!K3', values: [["='Quarterly Planner'!E2"]] },
    { range: 'Life Dashboard!A4', values: [['="""✨ """&H3&""" ✨ — """&E3&"""\'s """&B3&""" Life Design System"""']] },
    // Wheel headers
    { range: 'Life Dashboard!A8:F8', values: [['Life Area', 'Current Score (1–10)', 'Target Score', 'Gap', 'Satisfaction Bar', 'Key Focus']] },
    // Wheel data
    { range: 'Life Dashboard!A9:F14', values: wheel },
    // Average row
    { range: 'Life Dashboard!A15:F15', values: [['AVERAGE LIFE SCORE','=AVERAGE(B9:B14)','=AVERAGE(C9:C14)','=C15-B15','=SPARKLINE({B15/10,1-B15/10},{"charttype","bar";"color1","#9B8EC4";"color2","#EDE9F5"})','Overall life satisfaction']] },
    // Top Goals headers
    { range: 'Life Dashboard!A18:G18', values: [['#','Goal','Life Area','Target Date','Status','% Complete','Priority']] },
    // Top Goals data
    { range: 'Life Dashboard!A19:G24', values: topGoals },
    // Quote
    { range: 'Life Dashboard!A27', values: [['"The secret of getting ahead is getting started. The secret of getting started is breaking your complex overwhelming tasks into small manageable tasks, and then starting on the first one." — Mark Twain']] },
    // KPI cards
    { range: 'Life Dashboard!G27', values: [['=COUNTA(\'Annual Goals\'!C8:C100)&" goals"']] },
    { range: 'Life Dashboard!I27', values: [['=COUNTIF(\'Annual Goals\'!H8:H100,"Complete")&" complete"']] },
    { range: 'Life Dashboard!K27', values: [['=IFERROR(TEXT(COUNTIF(\'Annual Goals\'!H8:H100,"Complete")/COUNTA(\'Annual Goals\'!C8:C100),"0%")&" done","0% done")']] },
    { range: 'Life Dashboard!G29', values: [['=COUNTA(\'Habit Tracker\'!A5:A20)&" habits tracked"']] },
    { range: 'Life Dashboard!J29', values: [['="Life Score: "&TEXT(AVERAGE(B9:B14),"0.0")&" / 10"']] },
  ], 'dashboard-values');

  console.log('Life Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
