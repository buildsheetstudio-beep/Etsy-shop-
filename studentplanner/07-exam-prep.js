'use strict';
const { batchUpdate, valuesBatchUpdate, hex, COLOR, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = sheetMap['Exam Prep']; // 5

// 5 exams
const EXAMS = [
  ['English 101',    'Final Essay Portfolio',         '12/10/2025', 40, 3, 'Bring printed copies + USB backup'],
  ['Calculus II',    'Final Exam',                    '12/12/2025', 30, 2, 'Cumulative — all chapters'],
  ['World History',  'Final Exam',                    '12/15/2025', 30, 2, 'Essay format, 3 questions'],
  ['Biology 201',    'Lab Practical + Final',         '12/17/2025', 30, 3, 'Practical first then written'],
  ['Economics 101',  'Final Exam + Project Due',      '12/19/2025', 35, 4, 'Bring research project report'],
];

// 3 topic checklist blocks
const CHECKLISTS = [
  {
    course: 'English 101',
    title: 'English 101 — Essay Portfolio Prep',
    topics: [
      'Read all feedback from Essay 1 & 2',
      'Create outline for portfolio introduction',
      'Revise Essay 1 — Literary Analysis',
      'Revise Essay 2 — Research Paper draft',
      'Write reflective commentary (500 words)',
      'Proofread all essays for grammar',
      'Format per MLA guidelines',
      'Print + backup digital copy',
    ],
  },
  {
    course: 'Calculus II',
    title: 'Calculus II — Final Exam Prep',
    topics: [
      'Review integration techniques (u-sub, by parts)',
      'Practice Riemann sums (10 problems)',
      'Series and sequences — convergence tests',
      'Differential equations — separation of variables',
      'Taylor and Maclaurin series',
      'Past exam review — Problem Set 1-5',
      'Visit office hours for weak areas',
      'Create formula cheat sheet (if allowed)',
    ],
  },
  {
    course: 'Biology 201',
    title: 'Biology 201 — Lab Practical Prep',
    topics: [
      'Review lab reports 1, 2, and 3',
      'Memorize cell organelle functions',
      'Review DNA replication steps',
      'Photosynthesis light & dark reactions',
      'Practice microscope identification',
      'Study enzyme kinetics graphs',
      'Review ecosystem interactions',
      'Mock practical with lab partner',
    ],
  },
];

(async () => {
  const reqs = [];

  // ── Column widths ──
  const widths = [[0,160],[1,200],[2,100],[3,80],[4,80],[5,200],[6,50],[7,40],[8,160],[9,220]];
  widths.forEach(([ci,px]) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    },
  }));

  // ── Row heights ──
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 4 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });

  // ── Title row ──
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 10),
      cell: {
        userEnteredValue: { stringValue: '📝 Exam Prep Command Center' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.peach),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 18 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Subtitle ──
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, 10),
      cell: {
        userEnteredValue: { stringValue: 'Final Exams — December 2025 | Alex Rivera' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.palePeach),
          textFormat: { foregroundColor: hex(COLOR.orange), bold: false, italic: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Stats row (row 2) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 5, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 2, 3, 0, 5),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(COLOR.palePeach),
          textFormat: { foregroundColor: hex(COLOR.orange), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 2, 3, 5, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(COLOR.paleGold),
          textFormat: { foregroundColor: hex(COLOR.amber), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Exam overview section header (row 3) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 3, 4, 0, 10),
      cell: {
        userEnteredValue: { stringValue: '── EXAM SCHEDULE ──' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.orange),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Exam header row (row 4) ──
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 4, 5, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(COLOR.peach),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Exam data rows 5-9 ──
  for (let r = 5; r < 10; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 0, 10),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(r % 2 === 0 ? COLOR.palePeach : COLOR.white),
            textFormat: { fontSize: 10 }, verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }
  [2,3,4,5,6,7].forEach(ci => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 5, 10, ci, ci + 1),
        cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
        fields: 'userEnteredFormat.horizontalAlignment',
      },
    });
  });

  // ── Separator row between exams and checklists (row 10) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 10, 11, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 10, 11, 0, 10),
      cell: {
        userEnteredValue: { stringValue: '── STUDY TOPIC CHECKLISTS ──' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.orange),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });

  // ── 3 checklist blocks side by side ──
  // Layout: Each checklist in its own column group
  // Block 1: cols A-C (0-2), Block 2: cols D-F (3-5 → with gap), Block 3: cols G-I (6-8)
  // Actually lay them out vertically stacked since we only have 10 cols
  // Better: 3 blocks stacked (rows 11 + onward)
  let checkRow = 11;
  CHECKLISTS.forEach((cl, bi) => {
    const bgColors = [COLOR.paleSkyBlue, COLOR.paleMint, COLOR.palePeach];
    const hdrColors = [COLOR.darkBlue, COLOR.emerald, COLOR.orange];
    const bg = bgColors[bi];
    const hdr = hdrColors[bi];

    // Block title
    reqs.push({ mergeCells: { range: gridRange(SID, checkRow, checkRow + 1, 0, 10), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, checkRow, checkRow + 1, 0, 10),
        cell: {
          userEnteredValue: { stringValue: cl.title },
          userEnteredFormat: {
            backgroundColor: hex(hdr),
            textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 11 },
            horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
            padding: { left: 10 },
          },
        },
        fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: checkRow, endIndex: checkRow + 1 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
    checkRow++;

    // Column headers for checklist
    reqs.push({ mergeCells: { range: gridRange(SID, checkRow, checkRow + 1, 1, 7), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, checkRow, checkRow + 1, 0, 10),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(COLOR.nearBlack), bold: true, fontSize: 9 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    checkRow++;

    // Topic rows
    cl.topics.forEach((topic, ti) => {
      const rowBg = ti % 2 === 0 ? bg : COLOR.white;
      reqs.push({
        repeatCell: {
          range: gridRange(SID, checkRow, checkRow + 1, 0, 10),
          cell: {
            userEnteredFormat: {
              backgroundColor: hex(rowBg),
              textFormat: { fontSize: 10 }, verticalAlignment: 'MIDDLE',
            },
          },
          fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
        },
      });
      // Col H (7) = checkbox Done, Col I (8) = Notes
      reqs.push({
        repeatCell: {
          range: gridRange(SID, checkRow, checkRow + 1, 7, 8),
          cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
          fields: 'userEnteredFormat.horizontalAlignment',
        },
      });
      checkRow++;
    });

    // Progress row
    reqs.push({ mergeCells: { range: gridRange(SID, checkRow, checkRow + 1, 0, 7), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(SID, checkRow, checkRow + 1, 7, 10), mergeType: 'MERGE_ALL' } });
    const topicCount = cl.topics.length;
    const checkStartRow = checkRow - topicCount + 1; // 1-indexed
    reqs.push({
      repeatCell: {
        range: gridRange(SID, checkRow, checkRow + 1, 0, 10),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(hdr),
            textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    checkRow++;

    // Gap row
    checkRow++;
  });

  // ── Borders ──
  reqs.push({
    updateBorders: {
      range: gridRange(SID, 4, 10, 0, 10),
      innerHorizontal: { style: 'SOLID', color: hex(COLOR.palePeach) },
      innerVertical: { style: 'SOLID', color: hex(COLOR.palePeach) },
    },
  });

  await batchUpdate(id, reqs, 'examprep-format');

  // ── Values ──
  const examHeaders = ['Course','Exam Name','Date','Weight (%)','Confidence (1-5)','Days Until','','','Location/Notes'];
  const examData = EXAMS.map(([course, name, date, weight, conf, notes], i) => {
    const r = i + 6; // 1-indexed row
    return [course, name, date, weight, conf,
      `=IFERROR(DATEDIF(TODAY(),DATEVALUE(C${r}),"D"),"Past")`,
      '', '', notes];
  });

  // Checklist values
  let vCheckRow = 11;
  const checklistValues = [];
  CHECKLISTS.forEach((cl, bi) => {
    vCheckRow++; // title row
    // header row
    checklistValues.push({
      range: `Exam Prep!A${vCheckRow + 1}:J${vCheckRow + 1}`,
      values: [['#', 'Topic / Task', '', '', '', '', '', 'Done?', 'Notes', '']],
    });
    vCheckRow++; // header row

    // topic rows
    cl.topics.forEach((topic, ti) => {
      checklistValues.push({
        range: `Exam Prep!A${vCheckRow + 1}:J${vCheckRow + 1}`,
        values: [[ti + 1, topic, '', '', '', '', '', false, '', '']],
      });
      vCheckRow++;
    });

    // progress row
    const topicStartRow = vCheckRow - cl.topics.length + 1;
    checklistValues.push({
      range: `Exam Prep!A${vCheckRow + 1}:J${vCheckRow + 1}`,
      values: [[
        `Progress: `,
        `=COUNTIF(H${topicStartRow}:H${vCheckRow},TRUE)&"/"&COUNTA(B${topicStartRow}:B${vCheckRow})&" topics complete"`,
        '', '', '', '', '',
        `=REPT("▰",ROUND(COUNTIF(H${topicStartRow}:H${vCheckRow},TRUE)/COUNTA(B${topicStartRow}:B${vCheckRow})*10,0))&REPT("▱",10-ROUND(COUNTIF(H${topicStartRow}:H${vCheckRow},TRUE)/COUNTA(B${topicStartRow}:B${vCheckRow})*10,0))`,
        '', '',
      ]],
    });
    vCheckRow++;
    vCheckRow++; // gap
  });

  await valuesBatchUpdate(id, [
    { range: 'Exam Prep!A3', values: [['="""Exams Remaining: """&COUNTIF(F6:F10,">0")']] },
    { range: 'Exam Prep!F3', values: [['="""Next Exam: """&IFERROR(INDEX(B6:B10,MATCH(MIN(IF(ISNUMBER(F6:F10),F6:F10)),F6:F10,0)),"None")']] },
    { range: 'Exam Prep!A5:J5', values: [examHeaders] },
    { range: 'Exam Prep!A6:J10', values: examData },
    ...checklistValues,
  ], 'examprep-values');

  console.log('Exam Prep tab complete');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
