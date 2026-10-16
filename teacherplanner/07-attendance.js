'use strict';
const { sheets, batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Attendance'];
const S = "'Attendance'";
const REF = "'Reference Data'";

// Column layout (headers row 5, data rows 6+):
// A: Record ID  B: Date  C: Day  D: Class ID  E: Class Name
// F: Student ID  G: Student Name  H: Status  I: Arrival Time
// J: Notes  K: Parent Notified?  L: Excused by Admin?  M: Make-up Required?
// N: Make-up Completed?  O: Period / Block  P: blank

// School days Sept 2025 – Feb 2026 (Mon-Fri, skip holidays)
const SCHOOL_DAYS = [
  // September 2025 (school starts 09/02)
  '09/02/2025','09/03/2025','09/04/2025','09/05/2025',
  '09/08/2025','09/09/2025','09/10/2025','09/11/2025','09/12/2025',
  '09/15/2025','09/16/2025','09/17/2025','09/18/2025','09/19/2025',
  '09/22/2025','09/23/2025','09/24/2025','09/25/2025','09/26/2025',
  '09/29/2025','09/30/2025',
  // October 2025
  '10/01/2025','10/02/2025','10/03/2025',
  '10/06/2025','10/07/2025','10/08/2025','10/09/2025','10/10/2025',
  '10/13/2025','10/14/2025','10/15/2025','10/16/2025','10/17/2025',
  '10/20/2025','10/21/2025','10/22/2025','10/23/2025','10/24/2025',
  '10/27/2025','10/28/2025','10/29/2025','10/30/2025','10/31/2025',
  // November 2025
  '11/03/2025','11/04/2025','11/05/2025','11/06/2025','11/07/2025',
  '11/10/2025','11/12/2025','11/13/2025','11/14/2025', // 11/11 Veterans Day
  '11/17/2025','11/18/2025','11/19/2025','11/20/2025','11/21/2025',
  // Thanksgiving week off 11/24-11/28
  '12/01/2025','12/02/2025','12/03/2025','12/04/2025','12/05/2025',
  '12/08/2025','12/09/2025','12/10/2025','12/11/2025','12/12/2025',
  '12/15/2025','12/16/2025','12/17/2025','12/18/2025','12/19/2025',
  // Winter break 12/22 – 01/02
  '01/05/2026','01/06/2026','01/07/2026','01/08/2026','01/09/2026',
  '01/12/2026','01/13/2026','01/14/2026','01/15/2026','01/16/2026',
  // 01/19 MLK Day off
  '01/20/2026','01/21/2026','01/22/2026','01/23/2026',
  '01/26/2026','01/27/2026','01/28/2026','01/29/2026','01/30/2026',
  // February 2026
  '02/02/2026','02/03/2026','02/04/2026','02/05/2026','02/06/2026',
  '02/09/2026','02/10/2026','02/11/2026','02/12/2026','02/13/2026',
  '02/17/2026','02/18/2026','02/19/2026','02/20/2026', // 02/16 Presidents Day
  '02/23/2026','02/24/2026','02/25/2026','02/26/2026','02/27/2026',
];

const DAYS_OF_WEEK = ['Mon','Tue','Wed','Thu','Fri'];
function getDow(dateStr) {
  const [m, d, y] = dateStr.split('/').map(Number);
  const dt = new Date(y, m - 1, d);
  return DAYS_OF_WEEK[dt.getDay() - 1] || 'Mon';
}

// Students per class (same as gradebook)
const CLASS_STUDENTS = {
  'CLS-001': ['STU-0001','STU-0002','STU-0003','STU-0004','STU-0005','STU-0006',
              'STU-0007','STU-0008','STU-0009','STU-0010','STU-0011','STU-0012',
              'STU-0013','STU-0014','STU-0015','STU-0016','STU-0017','STU-0018',
              'STU-0019','STU-0020','STU-0021','STU-0022','STU-0023','STU-0024'],
  'CLS-002': ['STU-0025','STU-0026','STU-0027','STU-0028','STU-0029',
              'STU-0030','STU-0031','STU-0032','STU-0033','STU-0034'],
};

// Attendance status weights per student (index in CLS-001 STU list)
// Most students: 96% Present, 2% Absent, 1% Tardy, 1% Excused
// A few students have higher absence rates (Oliver Park STU-0015, Gabriel Santos STU-0007)
function getStatus(stuIdx, clsIdx, dayIdx) {
  const seed = (Math.sin(stuIdx * 7.3 + dayIdx * 13.1 + clsIdx * 3.7) + 1) / 2;
  const absSeed = (Math.sin(stuIdx * 5.9 + dayIdx * 17.3) + 1) / 2;
  // Students with higher absence rates: index 6 (Gabriel), 14 (Oliver)
  const isHighAbsence = stuIdx === 6 || stuIdx === 14;
  const absentRate = isHighAbsence ? 0.09 : 0.03;
  const tardyRate  = isHighAbsence ? 0.06 : 0.02;
  const excusedRate = 0.02;

  if (absSeed < absentRate) {
    return seed < 0.4 ? 'Excused' : 'Absent';
  }
  if (absSeed < absentRate + tardyRate) return 'Tardy';
  if (absSeed < absentRate + tardyRate + excusedRate) return 'Excused';
  return 'Present';
}

function getArrival(status) {
  if (status === 'Present') return '8:05 AM';
  if (status === 'Tardy') {
    const mins = [10, 15, 20, 25, 30][Math.floor(Math.random() * 5)];
    return `8:${mins < 10 ? '0'+mins : mins} AM`;
  }
  return '';
}

function getNotes(status, stuIdx, dayIdx) {
  const seed = (Math.sin(stuIdx * 9.1 + dayIdx * 11.7) + 1) / 2;
  if (status === 'Absent' && seed < 0.5) return 'Parent called';
  if (status === 'Absent' && seed >= 0.5) return 'No contact — follow up';
  if (status === 'Tardy') return 'Arrived late';
  if (status === 'Excused') return 'Doctor appointment';
  return '';
}

// Generate records: CLS-001 for all school days (24 students × ~100 days = 2400)
// + CLS-002 for first 40 school days (10 students × 40 days = 400)
// Total ≈ 2800 records — trim to 500+ by sampling fewer days for CLS-001
// Use first 22 days for CLS-001 full class (528 records) plus CLS-002 first 25 days (250 records)
// That gives 778 records, then add daily summaries for all 100+ days = 600+

// Strategy: CLS-001 — first 28 school days (28×24=672)
//           CLS-002 — first 28 school days (28×10=280)
// Total: 952 — slightly trim by using first 22 days each
// CLS-001: 22 days × 24 students = 528
// CLS-002: 22 days × 10 students = 220
// Total: 748 records ✓

const ATT_DAYS_CLS001 = SCHOOL_DAYS.slice(0, 22);
const ATT_DAYS_CLS002 = SCHOOL_DAYS.slice(0, 22);

const RECORDS = [];

function addRecords(days, clsId, students, clsOffset) {
  days.forEach((date, dayIdx) => {
    const dow = getDow(date);
    students.forEach((stuId, stuIdx) => {
      const status = getStatus(stuIdx, clsOffset, dayIdx);
      RECORDS.push([
        date, dow, clsId, stuId,
        status, status !== 'Absent' ? getArrival(status) : '',
        getNotes(status, stuIdx, dayIdx),
        status === 'Absent' ? true : false,  // Parent Notified
        false,                                // Excused by Admin
        status === 'Absent' || status === 'Tardy', // Make-up Required
        false,                                // Make-up Completed
        'All Day',                            // Period
      ]);
    });
  });
}

addRecords(ATT_DAYS_CLS001, 'CLS-001', CLASS_STUDENTS['CLS-001'], 0);
addRecords(ATT_DAYS_CLS002, 'CLS-002', CLASS_STUDENTS['CLS-002'], 1);

console.log(`Generating ${RECORDS.length} attendance records...`);

(async () => {
  // Dynamic pre-pass: unmerge existing merges in first 5 rows to prevent re-run errors
  const spInfo = await sheets.spreadsheets.get({ spreadsheetId: id, fields: 'sheets(properties.sheetId,merges)' });
  const thisSheet = (spInfo.data.sheets || []).find(sh => sh.properties && sh.properties.sheetId === SID);
  const existingMerges = (thisSheet && thisSheet.merges) ? thisSheet.merges.filter(m => m.startRowIndex < 5) : [];
  if (existingMerges.length > 0) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: id, requestBody: { requests: existingMerges.map(range => ({ unmergeCells: { range } })) } });
  }

  const vals = [];
  const fmt  = [];

  // BG
  fmt.push({ repeatCell:{ range: gridRange(SID,0,5010,0,16), cell:{userEnteredFormat:{backgroundColor:hex(C.bg)}}, fields:'userEnteredFormat.backgroundColor' }});

  // Title row 1
  vals.push({ range:`${S}!A1`, values:[['ATTENDANCE']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,16), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,16), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3437'), textFormat:{bold:true,fontSize:16,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1}, properties:{pixelSize:44}, fields:'pixelSize' }});

  // Filter controls row 2: Month (B2)
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,1), cell:{userEnteredFormat:{
    backgroundColor:hex(C.altRow), textFormat:{bold:true,fontSize:9,foregroundColor:hex(C.text),fontFamily:'Arial'},
    horizontalAlignment:'RIGHT', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,1,2), cell:{userEnteredFormat:{
    backgroundColor:hex(C.input), textFormat:{fontSize:9,foregroundColor:hex(C.text),fontFamily:'Arial'},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ setDataValidation:{ range: gridRange(SID,1,2,1,2), rule:{
    condition:{ type:'ONE_OF_LIST', values:[
      {userEnteredValue:'All'},
      {userEnteredValue:'September 2025'},
      {userEnteredValue:'October 2025'},
    ]},
    showCustomUi:true, strict:false,
  }}});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,2,16), cell:{userEnteredFormat:{
    backgroundColor:hex(C.bg),
  }}, fields:'userEnteredFormat.backgroundColor' }});

  // Alert row 3
  vals.push({ range:`${S}!A3`, values:[['Use the Month filter (B2) above to view attendance by month. Absent rows are red; Tardy rows are amber.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,2,3,0,16), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,2,3,0,16), cell:{userEnteredFormat:{
    backgroundColor:hex(C.Science), textFormat:{fontSize:9,foregroundColor:hex(C.text),fontFamily:'Arial'},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:8},
  }}, fields:'userEnteredFormat' }});

  vals.push({ range:`${S}!A4`, values:[['']] });
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:1,endIndex:3}, properties:{pixelSize:28}, fields:'pixelSize' }});

  // Headers row 5
  const hdrs = ['Record ID','Date','Day','Class ID','Class Name',
                'Student ID','Student Name','Status','Arrival Time',
                'Notes','Parent Notified?','Excused by Admin?',
                'Make-up Required?','Make-up Completed?','Period / Block',''];
  vals.push({ range:`${S}!A5`, values:[hdrs] });
  fmt.push({ repeatCell:{ range: gridRange(SID,4,5,0,16), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3437'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:4,endIndex:5}, properties:{pixelSize:30}, fields:'pixelSize' }});

  // Freeze rows 1-5
  fmt.push({ updateSheetProperties:{ properties:{ sheetId:SID, gridProperties:{ frozenRowCount:5 }}, fields:'gridProperties.frozenRowCount' }});

  // Column widths
  const colWidths = [80,90,50,80,130,90,140,80,90,180,100,100,110,110,100,40];
  colWidths.forEach((px, c) => {
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:c,endIndex:c+1}, properties:{pixelSize:px}, fields:'pixelSize' }});
  });

  await batchUpdate(id, fmt, '07-attendance format');
  // Write filter control labels and defaults to row 2
  await valuesBatchUpdate(id, [
    { range:`${S}!A2`, values:[['Month:']] },
    { range:`${S}!B2`, values:[['All']] },
  ], '07-attendance filter controls');

  // Formula columns for 1500 rows: A (Record ID), E (Class Name), G (Student Name)
  const ROWS = 1500;
  const aForms = [], eForms = [], gForms = [];
  for (let i = 0; i < ROWS; i++) {
    const r = i + 6;
    aForms.push([`=IF(B${r}="","","ATT-"&TEXT(ROW()-5,"0000"))`]);
    eForms.push([`=IFERROR(INDEX('School Year Setup'!B29:B48,MATCH(D${r},'School Year Setup'!A29:A48,0)),"")`]);
    gForms.push([`=IFERROR(INDEX('Student Roster'!B6:B1005,MATCH(F${r},'Student Roster'!A6:A1005,0)),"")`]);
  }
  await valuesBatchUpdate(id, [
    { range:`${S}!A6`, values: aForms },
    { range:`${S}!E6`, values: eForms },
    { range:`${S}!G6`, values: gForms },
  ], '07-attendance formulas');

  // Grade data as column arrays
  const bCol = RECORDS.map(r => [r[0]]);   // Date
  const cCol = RECORDS.map(r => [r[1]]);   // Day
  const dCol = RECORDS.map(r => [r[2]]);   // Class ID
  const fCol = RECORDS.map(r => [r[3]]);   // Student ID
  // H:P (Status through Period) — H=index 7, P=index 15 = 8 columns
  const hCol = RECORDS.map(r => [r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11]]);

  await valuesBatchUpdate(id, [
    { range:`${S}!B6`, values: bCol },
    { range:`${S}!C6`, values: cCol },
    { range:`${S}!D6`, values: dCol },
    { range:`${S}!F6`, values: fCol },
    { range:`${S}!H6`, values: hCol },
  ], '07-attendance data');

  // Data formatting
  const fmtData = [];

  // Data row height
  fmtData.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:5,endIndex:5005}, properties:{pixelSize:20}, fields:'pixelSize' }});

  // Alternate row tint
  fmtData.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,5005,0,16)],
    booleanRule:{ condition:{type:'CUSTOM_FORMULA',values:[{userEnteredValue:'=AND(MOD(ROW(),2)=0,B6<>"")'}]},
      format:{backgroundColor:hex(C.altRow)} },
  }, index:0}});

  // Absent = red tint
  fmtData.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,5005,0,16)],
    booleanRule:{ condition:{type:'CUSTOM_FORMULA',values:[{userEnteredValue:'=$H6="Absent"'}]},
      format:{backgroundColor:hex('#F8DCDC')} },
  }, index:1}});

  // Tardy = amber tint
  fmtData.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,5005,0,16)],
    booleanRule:{ condition:{type:'CUSTOM_FORMULA',values:[{userEnteredValue:'=$H6="Tardy"'}]},
      format:{backgroundColor:hex('#FDF3DC')} },
  }, index:2}});

  // Checkboxes for K (Parent Notified), L (Excused by Admin), M (Make-up Required), N (Make-up Completed)
  fmtData.push({ setDataValidation:{ range: gridRange(SID,5,5005,10,14), rule:{
    condition:{ type:'BOOLEAN' }, showCustomUi:true,
  }}});

  // Status dropdown
  fmtData.push({ setDataValidation:{ range: gridRange(SID,5,5005,7,8), rule:{
    condition:{ type:'ONE_OF_RANGE', values:[{userEnteredValue:`=${REF}!I2:I9`}] }, showCustomUi:true, strict:false,
  }}});

  // Center ID columns
  [0,1,2,3,5,7,8,14].forEach(c => {
    fmtData.push({ repeatCell:{ range: gridRange(SID,5,5005,c,c+1), cell:{userEnteredFormat:{
      horizontalAlignment:'CENTER',
    }}, fields:'userEnteredFormat.horizontalAlignment' }});
  });

  // Center checkbox columns
  fmtData.push({ repeatCell:{ range: gridRange(SID,5,5005,10,14), cell:{userEnteredFormat:{
    horizontalAlignment:'CENTER',
  }}, fields:'userEnteredFormat.horizontalAlignment' }});

  // Gray out rows not matching the selected month (highest priority — index 0)
  fmtData.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,5005,0,16)],
    booleanRule:{ condition:{type:'CUSTOM_FORMULA', values:[{userEnteredValue:'=AND($D6<>"",$B$2<>"All",TEXT($B6,"MMMM YYYY")<>$B$2)'}]},
      format:{ textFormat:{ foregroundColor:hex('#CCCCCC') }, backgroundColor:hex('#EDEDED') },
    },
  }, index:0}});

  await batchUpdate(id, fmtData, '07-attendance data format');
  console.log(`✅ Attendance done. ${RECORDS.length} records.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
