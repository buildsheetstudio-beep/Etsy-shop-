'use strict';
const { batchUpdate, valuesBatchUpdate, hex, COLOR, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = sheetMap['Grades & GPA']; // 3

// 5 courses with component weights and scores
// Smart formula: (weight*(score>0)*score) / (weight*(score>0)) — excludes unscored finals
// Resulting grade tiers: English=B+, Calculus=C-, History=A, Biology=C+, Econ=B-
const COURSES = [
  //  name              cred  HWw HWs  Qw   Qs   Midw Mids Partw Parts  Finw Fins
  ['English 101',        3,   20, 85,  15,  82,  25,  88,  0,   0,    40,  0 ],
  ['Calculus II',        4,   20, 72,  20,  68,  30,  74,  0,   0,    30,  0 ],
  ['World History',      3,   20, 96,  20,  93,  30,  91,  0,   0,    30,  0 ],
  ['Biology 201',        4,   15, 78,  20,  75,  25,  80,  10,  72,   30,  0 ],
  ['Economics 101',      3,   25, 84,  15,  80,  25,  83,  0,   0,    35,  0 ],
];
// cols: A=Course, B=Credits, C=HW Wt, D=HW Score, E=Quiz Wt, F=Quiz Score,
//        G=Mid Wt, H=Mid Score, I=Part Wt, J=Part Score, K=Final Wt, L=Final Score,
//        M=unused, N=unused, O=Current %, P=GPA pts, Q=Letter Grade

(async () => {
  const reqs = [];

  // ── Column widths ──
  const widths = [[0,170],[1,60],[2,60],[3,75],[4,65],[5,75],[6,65],[7,75],[8,65],[9,75],[10,65],[11,75],[12,40],[13,40],[14,90],[15,70],[16,90],[17,90]];
  widths.forEach(([ci,px]) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    },
  }));

  // ── Row heights ──
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 6 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 40 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // ── Title row ──
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 18), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 18),
      cell: {
        userEnteredValue: { stringValue: '🎓 Grades & GPA Calculator' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.lavender),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 18 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Subtitle ──
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 18), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, 18),
      cell: {
        userEnteredValue: { stringValue: 'Fall 2025 | Alex Rivera | State University' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.paleLavender),
          textFormat: { foregroundColor: hex(COLOR.violet), bold: false, italic: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── GPA summary block (rows 2-4) ──
  // Row 2: labels
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 5, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 10, 15), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 15, 18), mergeType: 'MERGE_ALL' } });

  reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, 0, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, 5, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, 10, 15), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, 15, 18), mergeType: 'MERGE_ALL' } });

  // Row 4: values
  reqs.push({ mergeCells: { range: gridRange(SID, 4, 5, 0, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 4, 5, 5, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 4, 5, 10, 15), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 4, 5, 15, 18), mergeType: 'MERGE_ALL' } });

  const gpaLabelCells = [
    { r: 2, label: 'Semester GPA', bg: COLOR.paleLavender },
    { r: 3, label: 'Cumulative GPA', bg: COLOR.subheaderBg },
    { r: 4, label: 'Total Credits', bg: COLOR.paleMint },
  ];
  // Apply label format to all 4 blocks per row
  gpaLabelCells.forEach(({ r, bg }) => {
    [[0,5],[5,10],[10,15],[15,18]].forEach(([c0,c1]) => {
      reqs.push({
        repeatCell: {
          range: gridRange(SID, r, r + 1, c0, c1),
          cell: {
            userEnteredFormat: {
              backgroundColor: hex(bg),
              textFormat: { foregroundColor: hex(COLOR.nearBlack), bold: true, fontSize: 11 },
              horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
            },
          },
          fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
        },
      });
    });
  });

  // ── Component header row (row 5) ──
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 5, 6, 0, 18),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(COLOR.violet),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 9 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Sub-header (row 5 pair headers) ──
  // Weight/Score pairs under each component
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 5, 7, 0, 18),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(COLOR.violet),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 9 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });

  // ── Course data rows (rows 7-11) ──
  for (let r = 6; r < 11; r++) {
    const bg = r % 2 === 0 ? COLOR.altRow : COLOR.white;
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 0, 18),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { fontSize: 10 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }
  // Center numeric cols
  [1,2,3,4,5,6,7,8,9,10,11,14,15].forEach(ci => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 6, 11, ci, ci + 1),
        cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
        fields: 'userEnteredFormat.horizontalAlignment',
      },
    });
  });

  // ── Totals row (row 12) ──
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 11, 12, 0, 18),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(COLOR.paleLavender),
          textFormat: { bold: true, fontSize: 10 },
          verticalAlignment: 'MIDDLE',
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
    },
  });

  // ── Section separator ──
  reqs.push({ mergeCells: { range: gridRange(SID, 12, 13, 0, 18), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 12, 13, 0, 18),
      cell: {
        userEnteredValue: { stringValue: '── GPA SCALE REFERENCE ──' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.lavender),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── GPA Scale reference (rows 13-22) ──
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 13, 23, 0, 4),
      cell: {
        userEnteredFormat: {
          textFormat: { fontSize: 10 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(textFormat,verticalAlignment,horizontalAlignment)',
    },
  });

  // ── Target Grade Calculator section (rows 13+ on right side) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 13, 14, 5, 18), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 13, 14, 5, 18),
      cell: {
        userEnteredValue: { stringValue: '── TARGET GRADE CALCULATOR ──' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.paleSkyBlue),
          textFormat: { foregroundColor: hex(COLOR.darkBlue), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Borders ──
  reqs.push({
    updateBorders: {
      range: gridRange(SID, 5, 12, 0, 18),
      innerHorizontal: { style: 'SOLID', color: hex(COLOR.paleLavender) },
      innerVertical: { style: 'SOLID', color: hex(COLOR.paleLavender) },
      bottom: { style: 'SOLID_MEDIUM', color: hex(COLOR.lavender) },
    },
  });

  await batchUpdate(id, reqs, 'grades-format');

  // ── Values ──
  // Header row 5+6
  const compHeaders = [
    'Course', 'Credits',
    'HW Wt', 'HW Score',
    'Quiz Wt', 'Quiz Score',
    'Midterm Wt', 'Midterm Score',
    'Participation Wt', 'Part. Score',
    'Final Wt', 'Final Score',
    '', '',
    'Current %', 'GPA Pts', 'Letter Grade', 'Sparkline',
  ];

  // Smart grade formula for each row
  // =IF(A7="","",IFERROR((C7*(D7>0)*D7+E7*(F7>0)*F7+G7*(H7>0)*H7+I7*(J7>0)*J7+K7*(L7>0)*L7)/(C7*(D7>0)+E7*(F7>0)+G7*(H7>0)+I7*(J7>0)+K7*(L7>0)),0))
  function gradeFormula(r) {
    return `=IF(A${r}="","",IFERROR((C${r}*(D${r}>0)*D${r}+E${r}*(F${r}>0)*F${r}+G${r}*(H${r}>0)*H${r}+I${r}*(J${r}>0)*J${r}+K${r}*(L${r}>0)*L${r})/(C${r}*(D${r}>0)+E${r}*(F${r}>0)+G${r}*(H${r}>0)+I${r}*(J${r}>0)+K${r}*(L${r}>0)),0))`;
  }
  function gpaFormula(r) {
    return `=IF(O${r}="","",IFS(O${r}>=93,4.0,O${r}>=90,3.7,O${r}>=87,3.3,O${r}>=83,3.0,O${r}>=80,2.7,O${r}>=77,2.3,O${r}>=73,2.0,O${r}>=70,1.7,O${r}>=67,1.3,O${r}>=63,1.0,O${r}>=60,0.7,TRUE,0.0))`;
  }
  function letterFormula(r) {
    return `=IF(O${r}="","",IFS(O${r}>=93,"A",O${r}>=90,"A-",O${r}>=87,"B+",O${r}>=83,"B",O${r}>=80,"B-",O${r}>=77,"C+",O${r}>=73,"C",O${r}>=70,"C-",O${r}>=67,"D+",O${r}>=63,"D",O${r}>=60,"D-",TRUE,"F"))`;
  }
  function sparkFormula(r) {
    return `=IF(O${r}="","",REPT("▰",ROUND(O${r}/10,0))&REPT("▱",10-ROUND(O${r}/10,0)))`;
  }

  const courseRows = COURSES.map(([name, cred, hwW, hwS, qW, qS, mW, mS, pW, pS, fW, fS], i) => {
    const r = i + 7;
    return [name, cred, hwW, hwS, qW, qS, mW, mS, pW, pS, fW, fS, '', '', gradeFormula(r), gpaFormula(r), letterFormula(r), sparkFormula(r)];
  });

  // GPA scale rows
  const gpaScale = [
    ['A','93-100','4.0'],['A-','90-92','3.7'],['B+','87-89','3.3'],
    ['B','83-86','3.0'],['B-','80-82','2.7'],['C+','77-79','2.3'],
    ['C','73-76','2.0'],['C-','70-72','1.7'],['D','60-69','1.0'],['F','<60','0.0'],
  ];

  // Target grade calculator inputs (rows 14-20, cols F-R)
  const targetCalcData = [
    ['Course Name:', '', 'English 101'],
    ['Current Grade %:', '', '=IFERROR(VLOOKUP(H15,A7:O11,15,0),"")'],
    ['Final Exam Weight (%):', '', '=IFERROR(VLOOKUP(H15,A7:K11,11,0),"")'],
    ['Target Grade %:', '', '90'],
    ['Required Final Score:', '', '=IFERROR(IF(H18="","",IF(H17=0,"N/A",(H18-(O7*(1-H17/100)))*100/H17)),"")'],
    ['Is it achievable?', '', '=IF(H19="N/A","No final exam",IF(H19>100,"Not achievable — need "&ROUND(H19,1)&"%",IF(H19<=100,"Yes — need "&ROUND(H19,1)&"%","—")))'],
  ];

  // Semester GPA = SUMPRODUCT(credits*gpapts)/SUM(credits) at C32 (row 31)
  // But first let's place it in the GPA block rows 2-4
  await valuesBatchUpdate(id, [
    { range: 'Grades & GPA!A3:A3', values: [['Semester GPA']] },
    { range: 'Grades & GPA!F3:F3', values: [['Cumulative GPA']] },
    { range: 'Grades & GPA!K3:K3', values: [['Total Credits']] },
    { range: 'Grades & GPA!P3:P3', values: [['Standing']] },
    // Values row (row 4, index 3)
    { range: 'Grades & GPA!A4:A4', values: [['=IFERROR(SUMPRODUCT(B7:B11,P7:P11)/SUM(B7:B11),0)']] },
    { range: 'Grades & GPA!F4:F4', values: [['=IFERROR((E2*H2+B3*K2)/(H2+K2),0)']] },  // cross-semester; K2=prior GPA pts
    { range: 'Grades & GPA!K4:K4', values: [['=SUM(B7:B11)']] },
    { range: 'Grades & GPA!P4:P4', values: [['=IF(A4>=3.5,"Dean\'s List",IF(A4>=2.0,"Good Standing","Academic Probation"))']] },
    // Prior GPA inputs (row 2)
    { range: 'Grades & GPA!A2:A2', values: [['Prior Credits:']] },
    { range: 'Grades & GPA!B2:B2', values: [[45]] },
    { range: 'Grades & GPA!H2:H2', values: [[45]] },
    { range: 'Grades & GPA!F2:F2', values: [['Prior GPA:']] },
    { range: 'Grades & GPA!G2:G2', values: [[3.1]] },
    { range: 'Grades & GPA!K2:K2', values: [[3.1]] }, // prior GPA pts for cumulative calc
    // Component headers
    { range: 'Grades & GPA!A6:R6', values: [compHeaders] },
    // Course data
    { range: 'Grades & GPA!A7:R11', values: courseRows },
    // Totals row (row 12)
    { range: 'Grades & GPA!A12:R12', values: [['TOTALS','=SUM(B7:B11)','','','','','','','','','','','','','=IFERROR(SUMPRODUCT(B7:B11,O7:O11)/SUM(B7:B11),0)','=IFERROR(SUMPRODUCT(B7:B11,P7:P11)/SUM(B7:B11),0)','','']] },
    // GPA scale
    { range: 'Grades & GPA!A14:C23', values: gpaScale },
    // Target calculator
    { range: 'Grades & GPA!F15:H20', values: targetCalcData },
    // C32 = Semester GPA (Dashboard cross-reference)
    { range: 'Grades & GPA!C32', values: [['=IFERROR(SUMPRODUCT(B7:B11,P7:P11)/SUM(B7:B11),0)']] },
  ], 'grades-values');

  console.log('Grades & GPA tab complete');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
