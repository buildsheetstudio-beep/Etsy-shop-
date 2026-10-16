'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = sheetMap['Vision & Affirmations']; // 5

function banner(row, label, bg, textColor) {
  return [
    { mergeCells: { range: gridRange(SID, row, row + 1, 0, 10), mergeType: 'MERGE_ALL' } },
    { repeatCell: { range: gridRange(SID, row, row + 1, 0, 10), cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(textColor || C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } },
  ];
}

(async () => {
  const reqs = [];

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [28, 200, 220, 220, 220, 10, 220, 220, 220, 220];
  colWidths.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 64 }, fields: 'pixelSize' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, 10), cell: { userEnteredFormat: { backgroundColor: hex(C.softGold), textFormat: { foregroundColor: hex(C.deepPurple), bold: true, fontSize: 20 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // ── Row 2: Word of Year + theme ───────────────────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, 10), cell: { userEnteredFormat: { backgroundColor: hex(C.paleGold), textFormat: { foregroundColor: hex(C.deepPurple), bold: false, fontSize: 13, italic: true }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // ── Section 1: Core Values (rows 2-8) ─────────────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push(...banner(2, '  💛  CORE VALUES', C.softGold, C.deepPurple));

  // 2 rows × 5 values each = row 3 and row 4
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 5 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  // Row 3: 5 value cards A-E (no spacer)
  const valueColors = [C.paleLavender, C.paleSky, C.paleBlush, C.paleSage, C.paleGold, C.paleLavender, C.paleSky, C.paleBlush, C.paleSage, C.paleGold];
  const accentColors = [C.dustyLavender, C.skyBlue, C.blush, C.dustySage, C.softGold, C.dustyLavender, C.skyBlue, C.blush, C.dustySage, C.softGold];
  [3, 4].forEach(r => {
    [0,1,2,3,4].forEach(c => {
      const idx = (r - 3) * 5 + c;
      reqs.push({ repeatCell: { range: gridRange(SID, r, r + 1, c, c + 1), cell: { userEnteredFormat: { backgroundColor: hex(valueColors[idx] || C.paleLavender), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(accentColors[idx] || C.dustyLavender) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' } });
    });
    // Right 5 cols (5-9)
    [5,6,7,8,9].forEach(c => {
      const idx = (r - 3) * 5 + (c - 5);
      reqs.push({ repeatCell: { range: gridRange(SID, r, r + 1, c, c + 1), cell: { userEnteredFormat: { backgroundColor: hex(valueColors[idx] || C.paleLavender), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(accentColors[idx] || C.dustyLavender) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' } });
    });
  });

  // ── Section 2: Dream Life Vision (rows 5-16) ──────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push(...banner(5, '  🌟  DREAM LIFE VISION — 5 YEARS', C.dustyLavender));
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  // 6 life area vision cards in 2 columns × 3 rows
  // Rows 6 (headers), 7-9 (3 pairs)
  const visionAreas = ['Health & Fitness','Career & Purpose','Finance & Wealth','Relationships','Personal Growth','Fun & Adventure'];
  const visionBgs = [C.paleSage, C.paleSky, C.paleGold, C.paleBlush, C.paleLavender, '#FED7AA'];
  const visionAccents = [C.dustySage, C.skyBlue, C.softGold, C.blush, C.dustyLavender, '#F97316'];

  // Column headers for vision section
  reqs.push({ repeatCell: { range: gridRange(SID, 6, 7, 0, 5), cell: { userEnteredFormat: { backgroundColor: hex(C.deepPurple), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 6, 7, 5, 10), cell: { userEnteredFormat: { backgroundColor: hex(C.deepPurple), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  for (let v = 0; v < 6; v++) {
    const row = 7 + Math.floor(v / 2);
    const startCol = (v % 2 === 0) ? 0 : 5;
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: row, endIndex: row + 1 }, properties: { pixelSize: 80 }, fields: 'pixelSize' } });
    reqs.push({ mergeCells: { range: gridRange(SID, row, row + 1, startCol, startCol + 5), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SID, row, row + 1, startCol, startCol + 5), cell: { userEnteredFormat: { backgroundColor: hex(visionBgs[v]), textFormat: { fontSize: 10, foregroundColor: hex(visionAccents[v]) }, wrapStrategy: 'WRAP', verticalAlignment: 'TOP', horizontalAlignment: 'LEFT' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment,horizontalAlignment)' } });
  }

  // ── Section 3: Affirmations (rows 10-20) ─────────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push(...banner(10, '  ✨  DAILY AFFIRMATIONS', C.blush));
  // 10 affirmations in 2 columns
  for (let r = 11; r <= 20; r++) {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r + 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });
    const isLeft = (r - 11) % 2 === 0;
    const bg = r % 2 === 0 ? C.paleBlush : C.white;
    reqs.push({ repeatCell: { range: gridRange(SID, r, r + 1, 0, 10), cell: { userEnteredFormat: { backgroundColor: hex(bg), wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,wrapStrategy,verticalAlignment)' } });
    // Number col
    reqs.push({ repeatCell: { range: gridRange(SID, r, r + 1, 0, 1), cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.blush) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });
    // Affirmation text
    reqs.push({ repeatCell: { range: gridRange(SID, r, r + 1, 1, 10), cell: { userEnteredFormat: { textFormat: { fontSize: 10, italic: true, foregroundColor: hex(C.bodyText) } } }, fields: 'userEnteredFormat.textFormat' } });
  }

  // ── Section 4: Gratitude (rows 21-27) ────────────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 21, endIndex: 22 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push(...banner(21, '  🙏  GRATITUDE PRACTICE', C.dustySage));
  // 3 gratitude prompts × 2 responses = rows 22-27
  for (let r = 22; r <= 27; r++) {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r + 1 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
    reqs.push({ repeatCell: { range: gridRange(SID, r, r + 1, 0, 10), cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.paleSage : C.white), wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE', textFormat: { fontSize: 10 } } }, fields: 'userEnteredFormat(backgroundColor,wrapStrategy,verticalAlignment,textFormat)' } });
    reqs.push({ repeatCell: { range: gridRange(SID, r, r + 1, 0, 1), cell: { userEnteredFormat: { textFormat: { bold: true, foregroundColor: hex(C.dustySage) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });
  }

  // ── Section 5: Manifestation / intention (rows 28-32) ────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 28, endIndex: 29 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push(...banner(28, '  🌙  MANIFESTATION INTENTIONS', C.medLavender));
  reqs.push({ mergeCells: { range: gridRange(SID, 29, 33, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 29, endIndex: 33 }, properties: { pixelSize: 45 }, fields: 'pixelSize' } });
  reqs.push({ repeatCell: { range: gridRange(SID, 29, 33, 0, 10), cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { fontSize: 11, italic: true, foregroundColor: hex(C.deepPurple) }, wrapStrategy: 'WRAP', verticalAlignment: 'TOP', horizontalAlignment: 'LEFT' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment,horizontalAlignment)' } });

  await batchUpdate(id, reqs, 'vision-format');

  // ── Values ────────────────────────────────────────────────────────────────
  const vals = [
    { range: 'Vision & Affirmations!A1', values: [['✦  Vision Board & Affirmations  ✦']] },
    { range: 'Vision & Affirmations!A2', values: [['Word of Year: FLOURISH  |  2025  |  Jordan  |  "I am growing into my fullest self"']] },

    // Core values (row 4 and 5 = index 3 and 4) - 5 per row, cols A-E and F-J but layout is A-J with 10 cols
    // Row 4 (index 3): 5 values
    { range: 'Vision & Affirmations!A4:E4', values: [['🌱 GROWTH', '❤️ AUTHENTICITY', '⚡ COURAGE', '🌿 BALANCE', '✨ EXCELLENCE']] },
    // Row 5 (index 4): second set - same row but cols F-J
    { range: 'Vision & Affirmations!A5:E5', values: [['🤝 CONNECTION', '💡 CURIOSITY', '🧘 PEACE', '🎯 PURPOSE', '🙏 GRATITUDE']] },

    // Vision area column headers (row 7)
    { range: 'Vision & Affirmations!A7', values: [['Life Area']] },
    { range: 'Vision & Affirmations!F7', values: [['Life Area']] },

    // 5-year visions
    { range: 'Vision & Affirmations!A8', values: [['💪 HEALTH & FITNESS\n\nI wake up energized, run half marathons for fun, cook nourishing meals from scratch, and feel proud in my body every day.']] },
    { range: 'Vision & Affirmations!F8', values: [['💼 CAREER & PURPOSE\n\nI lead a team I inspire, my online course helps thousands, and I earn well doing work that lights me up.']] },
    { range: 'Vision & Affirmations!A9', values: [['💰 FINANCE & WEALTH\n\n$50K emergency fund, mortgage on our dream home, fully funded retirement accounts, and money anxiety is a memory.']] },
    { range: 'Vision & Affirmations!F9', values: [['💞 RELATIONSHIPS\n\nDeep love with my partner, quality time with family every week, and friendships that feed my soul.']] },
    { range: 'Vision & Affirmations!A10', values: [['🌱 PERSONAL GROWTH\n\nFluent in Spanish, reading 2 books/month, journaling daily, and always becoming a better version of myself.']] },
    { range: 'Vision & Affirmations!F10', values: [['🎉 FUN & ADVENTURE\n\nTravel 3 countries per year, try something new monthly, and live a life I\'m excited to talk about.']] },

    // Affirmations
    { range: 'Vision & Affirmations!A12', values: [['1']] },
    { range: 'Vision & Affirmations!B12', values: [['I am becoming the healthiest, strongest, and most vibrant version of myself every single day.']] },
    { range: 'Vision & Affirmations!A13', values: [['2']] },
    { range: 'Vision & Affirmations!B13', values: [['I attract opportunities that align with my purpose and bring out my greatest gifts.']] },
    { range: 'Vision & Affirmations!A14', values: [['3']] },
    { range: 'Vision & Affirmations!B14', values: [['Money flows to me easily and abundantly. I manage it with wisdom and joy.']] },
    { range: 'Vision & Affirmations!A15', values: [['4']] },
    { range: 'Vision & Affirmations!B15', values: [['I am deeply loved and I love deeply. My relationships are nourishing and full of joy.']] },
    { range: 'Vision & Affirmations!A16', values: [['5']] },
    { range: 'Vision & Affirmations!B16', values: [['I show up consistently. Small actions taken daily create extraordinary results.']] },
    { range: 'Vision & Affirmations!A17', values: [['6']] },
    { range: 'Vision & Affirmations!B17', values: [['I am worthy of everything I dream of. My desires are valid and achievable.']] },
    { range: 'Vision & Affirmations!A18', values: [['7']] },
    { range: 'Vision & Affirmations!B18', values: [['I face challenges with courage and curiosity. Every obstacle is my teacher.']] },
    { range: 'Vision & Affirmations!A19', values: [['8']] },
    { range: 'Vision & Affirmations!B19', values: [['I am at peace. I trust the timing of my life and release the need to control everything.']] },
    { range: 'Vision & Affirmations!A20', values: [['9']] },
    { range: 'Vision & Affirmations!B20', values: [['I am building a life I don\'t need a vacation from. Every day is a gift I choose to use well.']] },
    { range: 'Vision & Affirmations!A21', values: [['10']] },
    { range: 'Vision & Affirmations!B21', values: [['I FLOURISH. I grow, I create, I connect, I contribute — this is my year.']] },

    // Gratitude
    { range: 'Vision & Affirmations!A23', values: [['🙏']] },
    { range: 'Vision & Affirmations!B23', values: [['3 things I\'m grateful for today:']] },
    { range: 'Vision & Affirmations!C23', values: [['My health, my partner\'s support, the fact that I get to pursue big dreams']] },
    { range: 'Vision & Affirmations!A24', values: [['💫']] },
    { range: 'Vision & Affirmations!B24', values: [['Something I\'m proud of this week:']] },
    { range: 'Vision & Affirmations!C24', values: [['Stayed consistent with meditation even on hard days']] },
    { range: 'Vision & Affirmations!A25', values: [['🌱']] },
    { range: 'Vision & Affirmations!B25', values: [['A person I\'m grateful for:']] },
    { range: 'Vision & Affirmations!C25', values: [['My partner — always in my corner']] },
    { range: 'Vision & Affirmations!A26', values: [['✨']] },
    { range: 'Vision & Affirmations!B26', values: [['Something beautiful I noticed today:']] },
    { range: 'Vision & Affirmations!C26', values: [['Morning light through the window while meditating']] },
    { range: 'Vision & Affirmations!A27', values: [['🎯']] },
    { range: 'Vision & Affirmations!B27', values: [['One intention I\'m setting for this week:']] },
    { range: 'Vision & Affirmations!C27', values: [['Show up fully — no half-measures this week']] },
    { range: 'Vision & Affirmations!A28', values: [['🌟']] },
    { range: 'Vision & Affirmations!B28', values: [['My affirmation for this moment:']] },
    { range: 'Vision & Affirmations!C28', values: [['I am exactly where I need to be, and I\'m growing every day.']] },

    // Manifestation
    { range: 'Vision & Affirmations!A30', values: [['I am living my dream life. I wake up excited every morning, surrounded by love, doing work that matters, with the health to enjoy it all. I have more than enough — money, time, joy, and purpose. I am FLOURISHING. This is already mine. I am grateful. I am ready. I am it.']] },
  ];

  await valuesBatchUpdate(id, vals, 'vision-values');

  console.log('Vision & Affirmations tab complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
