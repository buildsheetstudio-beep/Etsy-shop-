'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard'];

// Layout (0-based rows):
// Row 0: Title banner (full width, cols A-J)
// Row 1: Subtitle
// Row 2: spacer/divider
// Rows 3-4: Stat card labels (row 3) + values (row 4) — 5 cards × 2 cols
// Row 5: spacer
// Row 6: Disclaimer / note row
// Row 7: spacer
// Row 8: "📊 Chart Data Tables" divider header
//
// Chart data tables (sourced by chart scripts):
// Row 9:  "By Content Pillar" header
// Rows 10-19: Content Pillar COUNTIFS (10 pillars)
// Row 20: spacer
// Row 21: "By Episode Status" header
// Rows 22-27: Status COUNTIFS (6 statuses)
// Row 28: spacer
// Row 29: "Downloads — Top 10 Episodes" header
// Rows 30-39: Top 10 ep analytics (episode #, title, total downloads)
// Row 40: spacer
// Row 41: "Sponsorship Revenue by Month" header
// Rows 42-53: Monthly sponsorship SUMIFS (Jan-Dec 2024)

const CAL   = "'🎙 Content Calendar'";
const SPO   = "'💰 Sponsorship Tracker'";
const ANA   = "'📈 Episode Analytics'";

const PILLARS = [
  'Personal Growth','Content Creation','Productivity','Freelancing Tips',
  'Online Business','Mindset & Motivation','Tools & Workflows','Audience Building',
  'Work-Life Balance','Monetization',
];
const STATUSES = ['Published','Scheduled','Recording','Editing','On Hold','Ready to Record'];
const MONTHS   = ['January','February','March','April','May','June','July','August','September','October','November','December'];

(async () => {
  const reqs = [];

  // Row 0: Title banner
  reqs.push({ mergeCells: { range: gridRange(DSH, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 0, 1, 0, 10),
      cell: {
        userEnteredValue: { stringValue: '🎙 ULTIMATE PODCAST PLANNER' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepCharcoal),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});

  // Row 1: Subtitle
  reqs.push({ mergeCells: { range: gridRange(DSH, 1, 2, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 1, 2, 0, 10),
      cell: {
        userEnteredValue: { stringValue: 'Plan, track, grow — your complete podcast command center' },
        userEnteredFormat: {
          backgroundColor: hex(C.burntOrange),
          textFormat: { foregroundColor: hex(C.white), bold: false, fontSize: 11, italic: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});

  // Row 2: Spacer/divider
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 2, 3, 0, 10),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});

  // KPI cards: 5 cards, each 2 cols wide (A-J = 10 cols)
  // Row 3: KPI labels
  const cardColors = [C.deepCharcoal, C.burntOrange, C.mutedSage, C.amber, C.deepCharcoal];
  for (let i = 0; i < 5; i++) {
    const c = i * 2;
    reqs.push({ mergeCells: { range: gridRange(DSH, 3, 4, c, c+2), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, 3, 4, c, c+2),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(cardColors[i]),
            textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat',
      },
    });
  }
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Row 4: KPI values
  for (let i = 0; i < 5; i++) {
    const c = i * 2;
    reqs.push({ mergeCells: { range: gridRange(DSH, 4, 5, c, c+2), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, 4, 5, c, c+2),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.inputBg),
            textFormat: { foregroundColor: hex(cardColors[i]), bold: true, fontSize: 22 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat',
      },
    });
  }
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 4, endIndex: 5 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});

  // Row 5: spacer
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 5, endIndex: 6 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});

  // Row 6: Disclaimer
  reqs.push({ mergeCells: { range: gridRange(DSH, 6, 7, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 6, 7, 0, 10),
      cell: {
        userEnteredValue: { stringValue: '⚠️  All stats above are automatically calculated from your Content Calendar, Sponsorship Tracker, and Episode Analytics tabs. Update those tabs to keep this dashboard current.' },
        userEnteredFormat: {
          backgroundColor: hex(C.lightAmber),
          textFormat: { foregroundColor: hex(C.darkText), fontSize: 9, italic: true },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 6, endIndex: 7 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Row 7: spacer
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 7, endIndex: 8 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});

  // Row 8: chart data section header
  reqs.push({ mergeCells: { range: gridRange(DSH, 8, 9, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 8, 9, 0, 10),
      cell: {
        userEnteredValue: { stringValue: '📊  Dashboard Chart Data (auto-calculated)' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepCharcoal),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DSH, dimension: 'ROWS', startIndex: 8, endIndex: 9 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Sub-table headers (rows 9, 21, 29, 41)
  for (const [r, label] of [[9, '🎯  Episodes by Content Pillar'], [21, '📌  Episode Status Breakdown'], [29, '📥  Top 10 Episodes by Downloads'], [41, '💰  Sponsorship Revenue by Month (2024)']]) {
    reqs.push({ mergeCells: { range: gridRange(DSH, r, r+1, 0, 4), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, r, r+1, 0, 4),
        cell: {
          userEnteredValue: { stringValue: label },
          userEnteredFormat: {
            backgroundColor: hex(C.burntOrange),
            textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
            horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 26 }, fields: 'pixelSize',
    }});
  }

  // Data rows: alternating bg
  const dataRanges = [[10, 20], [22, 28], [30, 40], [42, 54]];
  for (const [start, end] of dataRanges) {
    for (let r = start; r < end; r++) {
      reqs.push({
        repeatCell: {
          range: gridRange(DSH, r, r+1, 0, 4),
          cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.inputBg : C.altRow), verticalAlignment: 'MIDDLE' } },
          fields: 'userEnteredFormat',
        },
      });
      reqs.push({ updateDimensionProperties: {
        range: { sheetId: DSH, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
        properties: { pixelSize: 24 }, fields: 'pixelSize',
      }});
    }
  }

  // Col widths: A=160, B=100, C=260, D=120, E-J=100
  const colWidths = [160, 100, 260, 120, 100, 100, 100, 100, 100, 100];
  for (const [i, px] of colWidths.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  // Formula bg on value cols
  reqs.push({ repeatCell: {
    range: gridRange(DSH, 4, 5, 0, 10),
    cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg) } },
    fields: 'userEnteredFormat',
  }});
  // Spacer rows
  for (const r of [20, 28, 40]) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DSH, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 8 }, fields: 'pixelSize',
    }});
  }

  await batchUpdate(id, reqs, 'dash-format');

  // ── Values ───────────────────────────────────────────────────────────────
  const data = [];

  // KPI labels (row 4 = sheet row 4)
  data.push({
    range: `'📊 Dashboard'!A4:J4`,
    values: [['Episodes Published','','In Pipeline','','Total Sponsor Revenue','','Avg Downloads/Ep','','Next Episode (days)','']],
  });

  // KPI values (row 5 = sheet row 5)
  data.push({
    range: `'📊 Dashboard'!A5:J5`,
    values: [[
      `=IFERROR(COUNTIF(${CAL}!G:G,"Published"),"")`,
      '',
      `=IFERROR(COUNTIFS(${CAL}!G:G,"<>Published",${CAL}!G:G,"<>"),"")`,
      '',
      `=IFERROR(SUMIFS(${SPO}!D:D,${SPO}!E:E,"Paid"),"")`,
      '',
      `=IFERROR(AVERAGE(${ANA}!F:F),"")`,
      '',
      `=IFERROR(MIN(IF(${CAL}!F:F>TODAY(),${CAL}!F:F))-TODAY(),"")`,
      '',
    ]],
  });
  // Array formula for Next Episode (entered as array formula)
  data.push({
    range: `'📊 Dashboard'!I5`,
    values: [[`=IFERROR(MIN(IF(${CAL}!F:F>TODAY(),${CAL}!F:F))-TODAY(),"")`]],
  });

  // By Content Pillar (rows 11-20 = 0-based 10-19)
  for (let i = 0; i < PILLARS.length; i++) {
    const row = i + 11; // sheet row
    data.push({
      range: `'📊 Dashboard'!A${row}:B${row}`,
      values: [[PILLARS[i], `=IFERROR(COUNTIF(${CAL}!B:B,"${PILLARS[i]}"),0)`]],
    });
  }

  // By Status (rows 23-28 = 0-based 22-27)
  for (let i = 0; i < STATUSES.length; i++) {
    const row = i + 23;
    data.push({
      range: `'📊 Dashboard'!A${row}:B${row}`,
      values: [[STATUSES[i], `=IFERROR(COUNTIF(${CAL}!G:G,"${STATUSES[i]}"),0)`]],
    });
  }

  // Top 10 Episodes by Downloads (rows 31-40 = 0-based 30-39)
  // Use LARGE/INDEX pattern to rank: just use the actual data rows since we have 20 fixed
  // Simple approach: LARGE on col F of Analytics, then pull episode # and title
  for (let i = 0; i < 10; i++) {
    const rank = i + 1;
    const row = i + 31;
    data.push({
      range: `'📊 Dashboard'!A${row}:D${row}`,
      values: [[
        `=IFERROR(INDEX(${ANA}!A:A,MATCH(LARGE(${ANA}!F$3:F$22,${rank}),${ANA}!F$3:F$22,0)+2),"")`,
        '',
        `=IFERROR(INDEX(${ANA}!B:B,MATCH(LARGE(${ANA}!F$3:F$22,${rank}),${ANA}!F$3:F$22,0)+2),"")`,
        `=IFERROR(LARGE(${ANA}!F$3:F$22,${rank}),0)`,
      ]],
    });
  }

  // Sponsorship Revenue by Month (rows 43-54 = 0-based 42-53)
  for (let i = 0; i < 12; i++) {
    const monthNum = i + 1;
    const row = i + 43;
    data.push({
      range: `'📊 Dashboard'!A${row}:B${row}`,
      values: [[
        MONTHS[i],
        `=IFERROR(SUMPRODUCT((MONTH(${SPO}!F$3:F$200)=${monthNum})*(YEAR(${SPO}!F$3:F$200)=2024)*(${SPO}!E$3:E$200="Paid")*(${SPO}!D$3:D$200)),0)`,
      ]],
    });
  }

  // Number formats for KPI values
  await batchUpdate(id, [
    { repeatCell: {
      range: gridRange(DSH, 4, 5, 4, 6),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    }},
    { repeatCell: {
      range: gridRange(DSH, 4, 5, 6, 8),
      cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' } } },
      fields: 'userEnteredFormat.numberFormat',
    }},
    { repeatCell: {
      range: gridRange(DSH, 41, 54, 1, 2),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    }},
    { repeatCell: {
      range: gridRange(DSH, 29, 40, 3, 4),
      cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' } } },
      fields: 'userEnteredFormat.numberFormat',
    }},
  ], 'dash-numfmt');

  await valuesBatchUpdate(id, data, 'dash-values');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
