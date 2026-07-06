'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DASH = sheetMap['🏆 Goal Dashboard'];

// Goal header GSheets rows (1-indexed): 3,9,15,21,27,33,39,45,51,57
const GOAL_GS_ROWS = [3,9,15,21,27,33,39,45,51,57];

const LIFE_AREAS = [
  'Career & Work','Health & Fitness','Finance & Wealth','Relationships',
  'Personal Growth','Education & Skills','Travel & Adventure','Creativity & Hobbies',
  'Spirituality & Wellbeing','Home & Family',
];
const STATUS_LIST = ['Not Started','In Progress','On Track','At Risk','Achieved','Paused'];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'🏆 Goal Dashboard'";
  const GL  = "'🎯 Goals & Milestones'";

  // Column widths
  [30,180,200,30,120,160,30,180,200].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DASH, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // ── ROW 0: Hero title banner (full width) ──
  reqs.push({ mergeCells: { range: gridRange(DASH,0,1,0,9), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,0,1,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midnightBlue),
      textFormat: { foregroundColor: hex(C.electricGold), bold: true, fontSize: 18 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 70 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A1`, values: [['🏆 Ultimate Annual Goal Planner']] });

  // ── ROW 1 (GSheets row 2): Name input + KPI labels ──
  reqs.push({ repeatCell: {
    range: gridRange(DASH,1,7,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.paleGold),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});

  // Input area labels (col A)
  reqs.push({ repeatCell: {
    range: gridRange(DASH,1,6,1,2),
    cell: { userEnteredFormat: {
      textFormat: { bold: true, foregroundColor: hex(C.medSlate), fontSize: 9 },
    }},
    fields: 'userEnteredFormat.textFormat',
  }});

  // KPI area labels (col D)
  reqs.push({ repeatCell: {
    range: gridRange(DASH,1,6,4,5),
    cell: { userEnteredFormat: {
      textFormat: { bold: true, foregroundColor: hex(C.medSlate), fontSize: 9 },
    }},
    fields: 'userEnteredFormat.textFormat',
  }});

  // KPI values (col E)
  reqs.push({ repeatCell: {
    range: gridRange(DASH,1,6,5,6),
    cell: { userEnteredFormat: {
      textFormat: { bold: true, foregroundColor: hex(C.electricGold), fontSize: 14 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(textFormat,horizontalAlignment)',
  }});

  vals.push({ range: `${TAB}!B2`, values: [['YOUR NAME']] });
  vals.push({ range: `${TAB}!C2`, values: [['Alex']] });
  vals.push({ range: `${TAB}!B3`, values: [['YEAR']] });
  vals.push({ range: `${TAB}!C3`, values: [[2025]] });

  // Title display formula (B4) — hero styled
  reqs.push({ repeatCell: {
    range: gridRange(DASH,3,4,1,3),
    cell: { userEnteredFormat: {
      textFormat: { bold: true, foregroundColor: hex(C.midnightBlue), fontSize: 16 },
    }},
    fields: 'userEnteredFormat.textFormat',
  }});
  vals.push({ range: `${TAB}!B4`, values: [[`=C2&"'s "&C3&" Goal Planner"`]] });

  // Quote row (B6, GSheets row 6 = 0-indexed 5)
  reqs.push({ repeatCell: {
    range: gridRange(DASH,5,6,1,5),
    cell: { userEnteredFormat: {
      textFormat: { italic: true, foregroundColor: hex(C.medSlate), fontSize: 9 },
    }},
    fields: 'userEnteredFormat.textFormat',
  }});
  reqs.push({ mergeCells: { range: gridRange(DASH,5,6,1,5), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${TAB}!B6`, values: [['\"A goal without a plan is just a wish.\" — Antoine de Saint-Exupéry']] });

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 1, endIndex: 7 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  // KPI formulas (E2:E6 = GSheets rows 2-6 = 0-indexed 1-5)
  const goalHRef = GOAL_GS_ROWS.map(r => `${GL}!H${r}`).join(',');
  const kpiFormulas = [
    `=IFERROR(COUNTA(${GOAL_GS_ROWS.map(r => `${GL}!A${r}`).join(',')}),"—")`,
    `=IFERROR(COUNTIF(${GL}!F:F,"Achieved"),"—")`,
    `=IFERROR(COUNTIF(${GL}!F:F,"In Progress"),"—")`,
    `=IFERROR(COUNTIF(${GL}!F:F,"Not Started"),"—")`,
    `=IFERROR(TEXT(AVERAGE(${goalHRef})/100,"0%"),"—")`,
  ];
  const kpiLabels = ['TOTAL GOALS','GOALS ACHIEVED','IN PROGRESS','NOT STARTED','OVERALL PROGRESS'];
  for (let i = 0; i < 5; i++) {
    vals.push({ range: `${TAB}!E${2+i}`, values: [[kpiLabels[i], kpiFormulas[i]]] });
  }

  // ── ROW 6 (0-indexed): Gold divider ──
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 6, endIndex: 7 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,6,7,0,9),
    cell: { userEnteredFormat: { backgroundColor: hex(C.electricGold) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});

  // ── ROW 7 (0-indexed): Goal Progress section header ──
  reqs.push({ mergeCells: { range: gridRange(DASH,7,8,0,9), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,7,8,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepNavy),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 7, endIndex: 8 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A8`, values: [['  GOAL PROGRESS OVERVIEW']] });

  // ── ROW 8 (0-indexed): Table column headers (GSheets row 9) ──
  reqs.push({ repeatCell: {
    range: gridRange(DASH,8,9,1,6),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepNavy),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 8, endIndex: 9 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!B9`, values: [['GOAL','LIFE AREA','% COMPLETE','PROGRESS BAR','STATUS']] });

  // ── ROWS 9-18 (0-indexed): 10 goal summary rows (GSheets 10-19) ──
  for (let i = 0; i < 10; i++) {
    const ri  = 9 + i;
    const row = 10 + i;
    const gs  = GOAL_GS_ROWS[i];
    const bg  = i % 2 === 0 ? C.white : C.paleGold;

    reqs.push({ repeatCell: {
      range: gridRange(DASH, ri, ri+1, 0, 9),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DASH, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 },
      properties: { pixelSize: 32 }, fields: 'pixelSize',
    }});
    // % col format
    reqs.push({ repeatCell: {
      range: gridRange(DASH, ri, ri+1, 3, 4),
      cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' }, horizontalAlignment: 'CENTER', backgroundColor: hex(bg) }},
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment,backgroundColor)',
    }});

    const pctRef = `${GL}!H${gs}`;
    vals.push({ range: `${TAB}!B${row}`, values: [[
      `=IFERROR(${GL}!A${gs},"")`,
      `=IFERROR(${GL}!B${gs},"")`,
      `=IFERROR(${pctRef}/100,0)`,
      `=IFERROR(SPARKLINE(${pctRef}/100,{"charttype","bar";"color1",IF(${pctRef}=100,"#2ECC71",IF(${pctRef}>=50,"#F4D03F","#C9A800"));"max",1}),"")`,
      `=IFERROR(${GL}!F${gs},"")`,
    ]]});
  }

  // ── ROW 19: Divider ──
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: DASH, dimension: 'ROWS', startIndex: 19, endIndex: 20 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,19,20,0,9),
    cell: { userEnteredFormat: { backgroundColor: hex(C.electricGold) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});

  // ── Chart data blocks: cols H-I (indices 7-8) ──
  // Status summary (rows 20-27 = GSheets 21-28)
  reqs.push({ repeatCell: {
    range: gridRange(DASH,19,28,7,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.paleBlue),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,19,20,7,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepNavy),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!H20`, values: [['STATUS','COUNT']] });
  STATUS_LIST.forEach((s, i) => {
    vals.push({ range: `${TAB}!H${21+i}`, values: [[s, `=COUNTIF(${GL}!F:F,"${s}")`]] });
  });

  // Life area avg progress (rows 29-40 = GSheets 30-41)
  reqs.push({ repeatCell: {
    range: gridRange(DASH,28,41,7,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.paleGold),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(DASH,28,29,7,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepNavy),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!H29`, values: [['LIFE AREA','AVG % COMPLETE']] });
  LIFE_AREAS.forEach((la, i) => {
    vals.push({ range: `${TAB}!H${30+i}`, values: [[la, `=IFERROR(AVERAGEIF(${GL}!B:B,"${la}",${GL}!H:H),"")`]] });
  });

  await batchUpdate(id, reqs, 'dashboard-format');
  await valuesBatchUpdate(id, vals, 'dashboard-values');
  console.log('Goal Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
