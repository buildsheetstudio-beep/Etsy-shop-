'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Gradebook'];
const S = "'Gradebook'";
const REF = "'Reference Data'";

// Column layout (headers row 5, data rows 6+):
// A: Grade ID  B: Student ID  C: Student Name  D: Class ID  E: Class Name
// F: Assignment ID  G: Assignment Name  H: Grade Category  I: Max Points  J: Score Type
// K: Due Date  L: Score  M: Percentage  N: Missing  O: Late  P: Excused
// Q: Graded  R: Notes  S: Date Graded  T: blank
// Assignments tab reads: F=AsnID, L=Score, N=Missing, Q=Graded

// CLS-001=STU-0001..0024, CLS-002=STU-0025..0034, CLS-003=STU-0035..0044
// CLS-004=STU-0045..0052, CLS-005=STU-0053..0063
const STU = {
  'CLS-001': [
    ['STU-0001',0.93],['STU-0002',0.88],['STU-0003',0.79],['STU-0004',0.84],
    ['STU-0005',0.76],['STU-0006',0.91],['STU-0007',0.72],['STU-0008',0.87],
    ['STU-0009',0.80],['STU-0010',0.95],['STU-0011',0.77],['STU-0012',0.89],
    ['STU-0013',0.75],['STU-0014',0.92],['STU-0015',0.68],['STU-0016',0.85],
    ['STU-0017',0.83],['STU-0018',0.78],['STU-0019',0.74],['STU-0020',0.82],
    ['STU-0021',0.90],['STU-0022',0.71],['STU-0023',0.96],['STU-0024',0.81],
  ],
  'CLS-002': [
    ['STU-0025',0.91],['STU-0026',0.86],['STU-0027',0.78],['STU-0028',0.83],
    ['STU-0029',0.75],['STU-0030',0.90],['STU-0031',0.70],['STU-0032',0.85],
    ['STU-0033',0.79],['STU-0034',0.94],
  ],
  'CLS-003': [
    ['STU-0035',0.77],['STU-0036',0.88],['STU-0037',0.74],['STU-0038',0.91],
    ['STU-0039',0.67],['STU-0040',0.84],['STU-0041',0.82],['STU-0042',0.76],
    ['STU-0043',0.73],['STU-0044',0.80],
  ],
  'CLS-004': [
    ['STU-0045',0.89],['STU-0046',0.70],['STU-0047',0.95],['STU-0048',0.80],
    ['STU-0049',0.92],['STU-0050',0.87],['STU-0051',0.78],['STU-0052',0.83],
  ],
  'CLS-005': [
    ['STU-0053',0.75],['STU-0054',0.90],['STU-0055',0.71],['STU-0056',0.86],
    ['STU-0057',0.79],['STU-0058',0.93],['STU-0059',0.76],['STU-0060',0.88],
    ['STU-0061',0.74],['STU-0062',0.91],['STU-0063',0.68],
  ],
};

// [asnId, clsId, maxPts, dateGraded, missingRate]
const GRADED = [
  // CLS-001 ELA — 12 graded assignments
  ['ASN-0001','CLS-001',10,'09/05/2025',0.04],
  ['ASN-0002','CLS-001',20,'09/10/2025',0.04],
  ['ASN-0003','CLS-001',25,'09/15/2025',0.08],
  ['ASN-0004','CLS-001',50,'10/06/2025',0.04],
  ['ASN-0005','CLS-001',40,'10/14/2025',0.08],
  ['ASN-0006','CLS-001',60,'10/29/2025',0.04],
  ['ASN-0007','CLS-001',10,'11/05/2025',0.04],
  ['ASN-0008','CLS-001',25,'11/14/2025',0.08],
  ['ASN-0009','CLS-001',10,'11/18/2025',0.04],
  ['ASN-0010','CLS-001',100,'01/22/2026',0.04],
  ['ASN-0011','CLS-001',5,'02/09/2026',0.04],
  ['ASN-0012','CLS-001',80,'03/11/2026',0.04],
  // CLS-002 Reading — 8 graded
  ['ASN-0016','CLS-002',20,'09/05/2025',0.04],
  ['ASN-0017','CLS-002',15,'09/12/2025',0.04],
  ['ASN-0018','CLS-002',20,'09/24/2025',0.08],
  ['ASN-0019','CLS-002',10,'10/03/2025',0.10],
  ['ASN-0020','CLS-002',5,'10/17/2025',0.04],
  ['ASN-0021','CLS-002',20,'11/07/2025',0.08],
  ['ASN-0022','CLS-002',50,'11/25/2025',0.04],
  ['ASN-0023','CLS-002',60,'01/28/2026',0.04],
  // CLS-003 Math — 10 graded
  ['ASN-0026','CLS-003',20,'09/04/2025',0.08],
  ['ASN-0027','CLS-003',15,'09/10/2025',0.04],
  ['ASN-0028','CLS-003',25,'09/17/2025',0.08],
  ['ASN-0029','CLS-003',50,'10/09/2025',0.04],
  ['ASN-0030','CLS-003',20,'10/15/2025',0.08],
  ['ASN-0031','CLS-003',10,'11/06/2025',0.04],
  ['ASN-0032','CLS-003',30,'11/20/2025',0.08],
  ['ASN-0033','CLS-003',50,'12/10/2025',0.04],
  ['ASN-0034','CLS-003',100,'01/22/2026',0.04],
  ['ASN-0035','CLS-003',20,'02/12/2026',0.04],
  // CLS-004 Science — 7 graded
  ['ASN-0042','CLS-004',10,'09/05/2025',0.04],
  ['ASN-0043','CLS-004',30,'09/19/2025',0.04],
  ['ASN-0044','CLS-004',25,'10/01/2025',0.08],
  ['ASN-0045','CLS-004',30,'10/24/2025',0.04],
  ['ASN-0046','CLS-004',60,'11/14/2025',0.04],
  ['ASN-0047','CLS-004',20,'12/04/2025',0.08],
  ['ASN-0048','CLS-004',60,'02/03/2026',0.04],
  // CLS-005 Social Studies — 7 graded
  ['ASN-0052','CLS-005',15,'09/08/2025',0.04],
  ['ASN-0053','CLS-005',20,'09/22/2025',0.04],
  ['ASN-0054','CLS-005',25,'10/09/2025',0.08],
  ['ASN-0055','CLS-005',50,'11/18/2025',0.04],
  ['ASN-0056','CLS-005',50,'12/10/2025',0.04],
  ['ASN-0057','CLS-005',15,'01/15/2026',0.08],
  ['ASN-0058','CLS-005',25,'02/05/2026',0.08],
];

function genScore(perf, maxPts, ai, si) {
  const noise = (Math.sin(si * 7.13 + ai * 13.7) + 1) / 2 * 0.2 - 0.1;
  return Math.round(maxPts * Math.min(1, Math.max(0.4, perf + noise)));
}
function isMiss(rate, ai, si) { return (Math.sin(si * 11.3 + ai * 17.9) + 1) / 2 < rate; }
function isLate(ai, si) { return (Math.sin(si * 5.7 + ai * 8.3) + 1) / 2 < 0.12; }

// Build flat grade records
const GRADES = [];
GRADED.forEach(([asnId, clsId, maxPts, dateGraded, missRate], ai) => {
  STU[clsId].forEach(([stuId, perf], si) => {
    const miss = isMiss(missRate, ai, si);
    const late = !miss && isLate(ai, si);
    GRADES.push([stuId, clsId, asnId,
      miss ? '' : genScore(perf, maxPts, ai, si),
      miss, late, false, !miss,
      miss ? 'Not submitted' : (late ? 'Late submission' : ''),
      miss ? '' : dateGraded]);
  });
});
// Total records
console.log(`Generating ${GRADES.length} grade records...`);

(async () => {
  const vals = [];
  const fmt  = [];

  // BG
  fmt.push({ repeatCell:{ range: gridRange(SID,0,5010,0,20), cell:{userEnteredFormat:{backgroundColor:hex(C.bg)}}, fields:'userEnteredFormat.backgroundColor' }});

  // Title row 1
  vals.push({ range:`${S}!A1`, values:[['GRADEBOOK']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,20), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,20), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3437'), textFormat:{bold:true,fontSize:16,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1}, properties:{pixelSize:44}, fields:'pixelSize' }});

  // Subtitle row 2
  vals.push({ range:`${S}!A2`, values:[['Enter scores in the Score column. Percentage auto-calculates. Check Missing for unsubmitted work; Graded when scored.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,1,2,0,20), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,20), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:9,foregroundColor:hex('#D5D8DB'),italic:true,fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});

  // Info row 3
  vals.push({ range:`${S}!A3`, values:[['Filter by Class ID or Student ID. Missing rows highlighted red. Late submissions highlighted amber. Grade summary on Dashboard tab.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,2,3,0,20), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,2,3,0,20), cell:{userEnteredFormat:{
    backgroundColor:hex(C.Math), textFormat:{fontSize:9,foregroundColor:hex(C.text),fontFamily:'Arial'},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:8},
  }}, fields:'userEnteredFormat' }});

  // Spacer row 4
  vals.push({ range:`${S}!A4`, values:[['']] });
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:1,endIndex:3}, properties:{pixelSize:28}, fields:'pixelSize' }});

  // Headers row 5
  const hdrs = ['Grade ID','Student ID','Student Name','Class ID','Class Name',
                'Assignment ID','Assignment Name','Grade Category','Max Points','Score Type',
                'Due Date','Score','Percentage','Missing?','Late?','Excused?',
                'Graded?','Notes','Date Graded',''];
  vals.push({ range:`${S}!A5`, values:[hdrs] });
  fmt.push({ repeatCell:{ range: gridRange(SID,4,5,0,20), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3437'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:4,endIndex:5}, properties:{pixelSize:30}, fields:'pixelSize' }});

  // Freeze rows 1-5
  fmt.push({ updateSheetProperties:{ properties:{ sheetId:SID, gridProperties:{ frozenRowCount:5 }}, fields:'gridProperties.frozenRowCount' }});

  // Column widths
  const colWidths = [80,90,150,80,130,95,180,110,80,90,90,60,80,65,55,65,65,160,90,40];
  colWidths.forEach((px, c) => {
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:c,endIndex:c+1}, properties:{pixelSize:px}, fields:'pixelSize' }});
  });

  await batchUpdate(id, fmt, '06-gradebook format');

  // Formula columns (all 1500 rows) — build as column arrays
  const ROWS = 1500;
  const aForms = [], cForms = [], eForms = [], gForms = [], hForms = [];
  const iForms = [], jForms = [], kForms = [], mForms = [];
  for (let i = 0; i < ROWS; i++) {
    const r = i + 6;
    aForms.push([`=IF(B${r}="","","GRD-"&TEXT(ROW()-5,"0000"))`]);
    cForms.push([`=IFERROR(INDEX('Student Roster'!B6:B1005,MATCH(B${r},'Student Roster'!A6:A1005,0)),"")`]);
    eForms.push([`=IFERROR(INDEX('School Year Setup'!B29:B48,MATCH(D${r},'School Year Setup'!A29:A48,0)),"")`]);
    gForms.push([`=IFERROR(INDEX('Assignments & Assessments'!D6:D1505,MATCH(F${r},'Assignments & Assessments'!A6:A1505,0)),"")`]);
    hForms.push([`=IFERROR(INDEX('Assignments & Assessments'!F6:F1505,MATCH(F${r},'Assignments & Assessments'!A6:A1505,0)),"")`]);
    iForms.push([`=IFERROR(INDEX('Assignments & Assessments'!L6:L1505,MATCH(F${r},'Assignments & Assessments'!A6:A1505,0)),"")`]);
    jForms.push([`=IFERROR(INDEX('Assignments & Assessments'!M6:M1505,MATCH(F${r},'Assignments & Assessments'!A6:A1505,0)),"")`]);
    kForms.push([`=IFERROR(INDEX('Assignments & Assessments'!K6:K1505,MATCH(F${r},'Assignments & Assessments'!A6:A1505,0)),"")`]);
    mForms.push([`=IFERROR(IF(I${r}=0,"",L${r}/I${r}),"")`]);
  }

  const formulaData = [
    { range:`${S}!A6`, values: aForms },
    { range:`${S}!C6`, values: cForms },
    { range:`${S}!E6`, values: eForms },
    { range:`${S}!G6`, values: gForms },
    { range:`${S}!H6`, values: hForms },
    { range:`${S}!I6`, values: iForms },
    { range:`${S}!J6`, values: jForms },
    { range:`${S}!K6`, values: kForms },
    { range:`${S}!M6`, values: mForms },
  ];
  await valuesBatchUpdate(id, formulaData, '06-gradebook formulas');

  // Grade data: write each non-formula column as a column array
  const bCol = GRADES.map(g => [g[0]]);  // Student ID
  const dCol = GRADES.map(g => [g[1]]);  // Class ID
  const fCol = GRADES.map(g => [g[2]]);  // Assignment ID
  const lCol = GRADES.map(g => [g[3]]);  // Score
  const nCol = GRADES.map(g => [g[4], g[5], g[6], g[7], g[8], g[9]]); // N:S (Missing..Date)

  const dataRanges = [
    { range:`${S}!B6`, values: bCol },
    { range:`${S}!D6`, values: dCol },
    { range:`${S}!F6`, values: fCol },
    { range:`${S}!L6`, values: lCol },
    { range:`${S}!N6`, values: nCol },
  ];
  await valuesBatchUpdate(id, dataRanges, '06-gradebook data');

  // Formatting for data rows
  const fmtData = [];

  // Alternate row tint
  fmtData.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,5005,0,20)],
    booleanRule:{ condition:{type:'CUSTOM_FORMULA',values:[{userEnteredValue:'=AND(MOD(ROW(),2)=0,B6<>"")'}]},
      format:{backgroundColor:hex(C.altRow)} },
  }, index:0}});

  // Missing rows = red tint (col N = TRUE)
  fmtData.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,5005,0,20)],
    booleanRule:{ condition:{type:'CUSTOM_FORMULA',values:[{userEnteredValue:'=$N6=TRUE'}]},
      format:{backgroundColor:hex('#F8DCDC')} },
  }, index:1}});

  // Late rows = amber tint (col O = TRUE)
  fmtData.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,5005,0,20)],
    booleanRule:{ condition:{type:'CUSTOM_FORMULA',values:[{userEnteredValue:'=$O6=TRUE'}]},
      format:{backgroundColor:hex('#FDF3DC')} },
  }, index:2}});

  // Data row height
  fmtData.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:5,endIndex:5005}, properties:{pixelSize:20}, fields:'pixelSize' }});

  // Percentage column number format (col M = index 12)
  fmtData.push({ repeatCell:{ range: gridRange(SID,5,5005,12,13), cell:{userEnteredFormat:{
    numberFormat:{type:'PERCENT',pattern:'0%'},
    horizontalAlignment:'CENTER',
  }}, fields:'userEnteredFormat(numberFormat,horizontalAlignment)' }});

  // Score column center-align (col L = index 11)
  fmtData.push({ repeatCell:{ range: gridRange(SID,5,5005,11,12), cell:{userEnteredFormat:{
    horizontalAlignment:'CENTER',
  }}, fields:'userEnteredFormat.horizontalAlignment' }});

  // Checkboxes for N (Missing), O (Late), P (Excused), Q (Graded)
  fmtData.push({ repeatCell:{ range: gridRange(SID,5,5005,13,17), cell:{userEnteredFormat:{
    horizontalAlignment:'CENTER',
  }}, fields:'userEnteredFormat.horizontalAlignment' }});
  fmtData.push({ setDataValidation:{ range: gridRange(SID,5,5005,13,17), rule:{
    condition:{ type:'BOOLEAN' }, showCustomUi:true,
  }}});

  // ID / code columns center-align (A,B,D,F)
  [0,1,3,5].forEach(c => {
    fmtData.push({ repeatCell:{ range: gridRange(SID,5,5005,c,c+1), cell:{userEnteredFormat:{
      horizontalAlignment:'CENTER',
    }}, fields:'userEnteredFormat.horizontalAlignment' }});
  });

  // Max Points + Score Type center
  fmtData.push({ repeatCell:{ range: gridRange(SID,5,5005,8,11), cell:{userEnteredFormat:{
    horizontalAlignment:'CENTER',
  }}, fields:'userEnteredFormat.horizontalAlignment' }});

  await batchUpdate(id, fmtData, '06-gradebook data format');
  console.log(`✅ Gradebook done. ${GRADES.length} grade records.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
