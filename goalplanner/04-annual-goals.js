'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = sheetMap['Annual Goals']; // 1

function sectionHeader(row, label, bg) {
  return {
    repeatCell: {
      range: gridRange(SID, row, row + 1, 0, 14),
      cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'LEFT' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  };
}

(async () => {
  const reqs = [];

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [30, 280, 120, 80, 80, 80, 130, 130, 100, 80, 80, 80, 20, 80, 80];
  colWidths.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // ── Row heights ────────────────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  reqs.push({
    mergeCells: { range: gridRange(SID, 0, 1, 0, 14), mergeType: 'MERGE_ALL' },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 14),
      cell: { userEnteredFormat: { backgroundColor: hex(C.dustyLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Row 2: subtitle ───────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, 14),
      cell: { userEnteredFormat: { backgroundColor: hex(C.medLavender), textFormat: { foregroundColor: hex(C.white), bold: false, fontSize: 11, italic: true }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Row 3: Column headers ─────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 2, 3, 0, 14),
      cell: { userEnteredFormat: { backgroundColor: hex(C.deepPurple), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Life area sections (18 goals: 3 per area × 6 areas) ──────────────────
  // Section layout: header row, then 3 goal rows, then blank separator
  // Area colors: Health(dustySage), Career(skyBlue), Finance(softGold), Relationships(blush), Personal(dustyLavender), Fun(softOrange/FED7AA)
  const areas = [
    { name: 'Health & Fitness', color: C.dustySage, paleBg: C.paleSage, startRow: 3 },
    { name: 'Career & Purpose', color: C.skyBlue, paleBg: C.paleSky, startRow: 9 },
    { name: 'Finance & Wealth', color: C.softGold, paleBg: C.paleGold, startRow: 15 },
    { name: 'Relationships',    color: C.blush,    paleBg: C.paleBlush, startRow: 21 },
    { name: 'Personal Growth',  color: C.dustyLavender, paleBg: C.paleLavender, startRow: 27 },
    { name: 'Fun & Adventure',  color: '#F97316', paleBg: '#FED7AA', startRow: 33 },
  ];

  areas.forEach(area => {
    const hRow = area.startRow;
    // Section header
    reqs.push(sectionHeader(hRow, area.name, area.color));
    reqs.push({ mergeCells: { range: gridRange(SID, hRow, hRow + 1, 1, 7), mergeType: 'MERGE_ALL' } });

    // 3 goal rows
    for (let r = hRow + 1; r <= hRow + 3; r++) {
      reqs.push({
        repeatCell: {
          range: gridRange(SID, r, r + 1, 0, 14),
          cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? area.paleBg : C.white) } },
          fields: 'userEnteredFormat.backgroundColor',
        },
      });
      // Goal cell (B) - wrap text
      reqs.push({
        repeatCell: {
          range: gridRange(SID, r, r + 1, 1, 2),
          cell: { userEnteredFormat: { wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } },
          fields: 'userEnteredFormat(wrapStrategy,verticalAlignment)',
        },
      });
    }

    // Row heights for goal rows
    for (let r = hRow + 1; r <= hRow + 3; r++) {
      reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r + 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
    }
    // Separator row height
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: hRow + 4, endIndex: hRow + 5 }, properties: { pixelSize: 8 }, fields: 'pixelSize' } });
    reqs.push({ repeatCell: { range: gridRange(SID, hRow + 4, hRow + 5, 0, 14), cell: { userEnteredFormat: { backgroundColor: hex(C.border) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // ── Section header row heights ────────────────────────────────────────────
  areas.forEach(area => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: area.startRow, endIndex: area.startRow + 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  });

  // ── Chart helper table (N5:O9) for goal status donut ─────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 3, 22, 13, 15),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderWhite), textFormat: { fontSize: 9 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  await batchUpdate(id, reqs, 'goals-format');

  // ── Values ────────────────────────────────────────────────────────────────
  const vals = [
    { range: 'Annual Goals!A1', values: [['✦  2025 Annual Goals  ✦']] },
    { range: 'Annual Goals!A2', values: [['Dream big, plan smart — your year of FLOURISH']] },
    { range: 'Annual Goals!A3', values: [['#', 'Goal', 'Life Area', 'Priority', 'Status', '% Done', 'Q1 Milestone', 'Q2 Milestone', 'Q3 Milestone', 'Q4 Milestone', 'Target Date', 'Actual Date', '', 'Status', 'Count']] },

    // ── Health & Fitness (row 4=header, 5-7=goals, 8=sep) ─────────────────
    { range: 'Annual Goals!A4', values: [['', '  HEALTH & FITNESS', '', '', '', '', '', '', '', '', '', '', '', '', '']] },
    { range: 'Annual Goals!A5:L5', values: [['1', 'Run a half marathon (13.1 miles)', 'Health & Fitness', 'High', 'In Progress', '0.35', 'Build to 5K', 'Build to 10K', 'Race day prep', 'Done / Next race', '2025-11-15', '']] },
    { range: 'Annual Goals!A6:L6', values: [['2', 'Meditate daily for 30 consecutive days', 'Health & Fitness', 'High', 'In Progress', '0.50', '10-day streak', '30-day streak', 'Maintain habit', 'Year review', '2025-06-30', '2025-06-10']] },
    { range: 'Annual Goals!A7:L7', values: [['3', 'Reduce sugar intake & meal prep weekly', 'Health & Fitness', 'Medium', 'In Progress', '0.25', 'Plan meals', 'Consistent 4 wks', 'New recipes', 'Year habit', '2025-12-31', '']] },

    // ── Career & Purpose (row 9=header, 10-12=goals, 13=sep) ──────────────
    { range: 'Annual Goals!A9', values: [['', '  CAREER & PURPOSE', '', '', '', '', '', '', '', '', '', '', '', '', '']] },
    { range: 'Annual Goals!A10:L10', values: [['4', 'Get promoted to Senior Manager', 'Career & Purpose', 'High', 'In Progress', '0.40', 'Mid-year review', 'Case prep', 'Promo pitch', 'New role', '2025-12-01', '']] },
    { range: 'Annual Goals!A11:L11', values: [['5', 'Launch online course or side project', 'Career & Purpose', 'High', 'Not Started', '0.10', 'Topic research', 'Outline & record', 'Beta launch', 'Full launch', '2025-10-01', '']] },
    { range: 'Annual Goals!A12:L12', values: [['6', 'Build personal brand on LinkedIn (5K followers)', 'Career & Purpose', 'Medium', 'In Progress', '0.20', '1K followers', '2.5K followers', '4K followers', '5K followers', '2025-12-31', '']] },

    // ── Finance & Wealth (row 15=header, 16-18=goals, 19=sep) ─────────────
    { range: 'Annual Goals!A15', values: [['', '  FINANCE & WEALTH', '', '', '', '', '', '', '', '', '', '', '', '', '']] },
    { range: 'Annual Goals!A16:L16', values: [['7', 'Save $15,000 emergency fund (6 months)', 'Finance & Wealth', 'High', 'In Progress', '0.60', 'Save $6K', 'Save $9K', 'Save $12K', 'Save $15K', '2025-12-31', '']] },
    { range: 'Annual Goals!A17:L17', values: [['8', 'Pay off $8,000 credit card debt', 'Finance & Wealth', 'High', 'In Progress', '0.45', 'Pay $2K', 'Pay $4K', 'Pay $6K', 'Pay off!', '2025-12-31', '']] },
    { range: 'Annual Goals!A18:L18', values: [['9', 'Invest $500/month in index funds', 'Finance & Wealth', 'Medium', 'In Progress', '0.33', 'Open account', '$1,500 in', '$3,000 in', '$6,000 in', '2025-12-31', '']] },

    // ── Relationships (row 21=header, 22-24=goals, 25=sep) ────────────────
    { range: 'Annual Goals!A21', values: [['', '  RELATIONSHIPS', '', '', '', '', '', '', '', '', '', '', '', '', '']] },
    { range: 'Annual Goals!A22:L22', values: [['10', 'Plan & take a family trip abroad', 'Relationships', 'High', 'In Progress', '0.30', 'Pick destination', 'Book flights', 'Prep itinerary', 'Travel!', '2025-08-31', '']] },
    { range: 'Annual Goals!A23:L23', values: [['11', 'Weekly date nights with partner (no phones)', 'Relationships', 'High', 'In Progress', '0.50', 'Start habit', 'Maintain 8 wks', 'Maintain 16 wks', 'Year habit', '2025-12-31', '']] },
    { range: 'Annual Goals!A24:L24', values: [['12', 'Reconnect with 3 old friends (meet IRL)', 'Relationships', 'Medium', 'Not Started', '0.00', 'Reach out all 3', 'Meet friend 1', 'Meet friend 2', 'Meet friend 3', '2025-11-30', '']] },

    // ── Personal Growth (row 27=header, 28-30=goals, 31=sep) ──────────────
    { range: 'Annual Goals!A27', values: [['', '  PERSONAL GROWTH', '', '', '', '', '', '', '', '', '', '', '', '', '']] },
    { range: 'Annual Goals!A28:L28', values: [['13', 'Read 24 books this year (2/month)', 'Personal Growth', 'High', 'In Progress', '0.42', '6 books', '12 books', '18 books', '24 books', '2025-12-31', '']] },
    { range: 'Annual Goals!A29:L29', values: [['14', 'Complete a 30-day journaling challenge', 'Personal Growth', 'Medium', 'Complete', '1.00', 'Week 1 done', 'Week 2 done', 'Month done!', 'Build habit', '2025-03-31', '2025-03-31']] },
    { range: 'Annual Goals!A30:L30', values: [['15', 'Learn conversational Spanish (B1 level)', 'Personal Growth', 'Medium', 'In Progress', '0.15', '50 Duolingo days', '100 days', '150 days + class', 'B1 test', '2025-12-31', '']] },

    // ── Fun & Adventure (row 33=header, 34-36=goals, 37=sep) ──────────────
    { range: 'Annual Goals!A33', values: [['', '  FUN & ADVENTURE', '', '', '', '', '', '', '', '', '', '', '', '', '']] },
    { range: 'Annual Goals!A34:L34', values: [['16', 'Take a solo weekend trip (somewhere new)', 'Fun & Adventure', 'Medium', 'Not Started', '0.00', 'Pick destination', 'Book & plan', 'Travel solo!', 'Write about it', '2025-09-30', '']] },
    { range: 'Annual Goals!A35:L35', values: [['17', 'Try 12 new restaurants or cuisines', 'Fun & Adventure', 'Low', 'In Progress', '0.42', '3 restaurants', '6 restaurants', '9 restaurants', '12 restaurants', '2025-12-31', '']] },
    { range: 'Annual Goals!A36:L36', values: [['18', 'Take a pottery or art class', 'Fun & Adventure', 'Low', 'Not Started', '0.00', 'Research classes', 'Enroll', 'First class!', 'Complete series', '2025-10-31', '']] },

    // ── Chart helper data ─────────────────────────────────────────────────
    { range: 'Annual Goals!N4', values: [['Status']] },
    { range: 'Annual Goals!O4', values: [['Count']] },
    { range: 'Annual Goals!N5:O9', values: [['Complete', 1], ['In Progress', 11], ['Not Started', 5], ['On Hold', 1], ['Abandoned', 0]] },
  ];

  await valuesBatchUpdate(id, vals, 'goals-values');

  // ── % Done number format ──────────────────────────────────────────────────
  const pctReqs = [];
  const pctRows = [5,6,7,10,11,12,16,17,18,22,23,24,28,29,30,34,35,36].map(r => r - 1);
  pctRows.forEach(r => {
    pctReqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 5, 6),
        cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } },
        fields: 'userEnteredFormat.numberFormat',
      },
    });
  });
  // Sparkline-style % bars via SPARKLINE formula
  // We'll add them in the next step via values

  await batchUpdate(id, pctReqs, 'goals-pct-format');

  console.log('Annual Goals tab complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
