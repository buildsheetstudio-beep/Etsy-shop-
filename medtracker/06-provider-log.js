'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const PRV = sheetMap['🏥 Provider Log'];

// A=Date | B=Provider Name | C=Specialty | D=Participant | E=Reason for Visit
// F=Follow-Up Needed (checkbox) | G=Next Appointment Date | H=Notes

const HEADERS = ['Date','Provider Name','Specialty','Participant','Reason for Visit','Follow-Up?','Next Appt. Date','Notes'];
const WIDTHS  = [90, 180, 130, 90, 200, 80, 110, 220];

const sampleData = [
  ['2024-01-08','Dr. Sarah Smith','Primary Care','Self','Annual wellness exam',false,'2025-01-08','All labs normal; continue current medications'],
  ['2024-01-15','Dr. James Jones','Dentist','Spouse','Routine cleaning + x-rays',true,'2024-07-15','Gum inflammation noted; return for deep cleaning'],
  ['2024-01-22','Dr. Patricia Lee','Pediatrics','Child 1','Annual wellness check',false,'2025-01-22','Growth and development on track; vaccines up to date'],
  ['2024-02-12','Dr. Kim','Pediatrics','Child 2','Ear infection – sick visit',true,'2024-03-01','Recheck if no improvement after antibiotic course'],
  ['2024-03-04','Dr. Wright','Optometrist','Spouse','Vision exam',false,'2025-03-04','Prescription updated; progressive lenses ordered'],
  ['2024-03-22','Dr. Lisa Chen','Therapist','Self','Ongoing therapy – anxiety',true,'2024-04-22','Monthly sessions continuing; making good progress'],
  ['2024-04-01','City ER Staff','Urgent Care','Self','Hand laceration – 5 stitches',false,'','Follow up with PCP if signs of infection develop'],
  ['2024-05-09','Dr. Lisa Chen','Therapist','Spouse','Couples counseling',true,'2024-06-09','Bi-monthly sessions scheduled'],
  ['2024-06-10','VisionPro','Optometrist','Spouse','Progressive lenses fitting',false,'','Adjustment period expected for 2–3 weeks'],
  ['2024-07-15','Dr. Patel','Specialist','Spouse','Pre-op consultation',true,'2024-07-22','Knee procedure scheduled; anesthesia consent obtained'],
  ['2024-08-05','Dr. Martinez','Orthopedics','Self','MRI results review',true,'2024-09-15','PT recommended 3×/week for 6 weeks; cortisone discussed'],
  ['2024-09-17','Urgent Care Plus','Urgent Care','Child 1','Broken finger evaluation',true,'2024-10-17','X-ray review with orthopedist in 4 weeks'],
  ['2024-10-01','Dr. Brown','Gastroenterology','Self','Colonoscopy',false,'2029-10-01','No polyps found; next scope in 5 years'],
  ['2024-11-20','Dr. Kim','Pediatrics','Child 2','Allergy follow-up',true,'2025-02-20','Continued allergen avoidance; retest in 3 months'],
  ['2024-12-16','Dr. Martinez','Orthopedics','Self','Orthopedic follow-up',true,'2025-03-16','PT completed; considering cortisone injection'],
];

(async () => {
  const reqs = [];

  // Row 0: Title
  reqs.push({ mergeCells: { range: gridRange(PRV, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(PRV, 0, 1, 0, 8),
      cell: {
        userEnteredValue: { stringValue: '🏥 Provider & Appointment Log' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepTeal),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: PRV, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Row 1: Headers
  reqs.push({
    repeatCell: {
      range: gridRange(PRV, 1, 2, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepTeal),
          textFormat: { foregroundColor: hex(C.warmCoral), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: PRV, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});

  // Data rows
  for (let r = 2; r < 17; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(PRV, r, r+1, 0, 8),
        cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.inputBg : C.altRow), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: PRV, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
  }
  reqs.push({
    repeatCell: {
      range: gridRange(PRV, 17, 200, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat',
    },
  });

  // Col widths
  for (const [i, px] of WIDTHS.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: PRV, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  // Number formats
  reqs.push({ repeatCell: {
    range: gridRange(PRV, 2, 200, 0, 1),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(PRV, 2, 200, 6, 7),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(PRV, 2, 200, 5, 6),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(PRV, 2, 200, 3, 4),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});

  await batchUpdate(id, reqs, 'provider-format');

  const data = [];
  data.push({ range: `'🏥 Provider Log'!A2:H2`, values: [HEADERS] });

  for (let i = 0; i < sampleData.length; i++) {
    const row = i + 3;
    data.push({ range: `'🏥 Provider Log'!A${row}:H${row}`, values: [sampleData[i]] });
  }

  await valuesBatchUpdate(id, data, 'provider-values');
  console.log('Provider Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
