'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const INS = sheetMap['📋 Insurance Claims'];

// A=Claim # | B=Date Submitted | C=Provider | D=Participant | E=Amount Billed
// F=Amount Covered by Insurance | G=Patient Responsibility (formula) | H=Status
// I=Date Resolved | J=Notes

const HEADERS = ['Claim #','Date Submitted','Provider','Participant','Amount Billed','Amount Covered','Patient Responsibility','Status','Date Resolved','Notes'];
const WIDTHS  = [100, 100, 170, 90, 110, 120, 130, 100, 100, 220];

// 12 rows covering every status at least once
const sampleData = [
  ['CLM-2024-001','2024-01-25','Dr. Smith','Self',1400,1120,'','Paid','2024-02-28','Annual physical + labs – fully processed'],
  ['CLM-2024-002','2024-02-05','Valley Dental','Spouse',650,420,'','Paid','2024-03-15','Dental cleaning and x-rays'],
  ['CLM-2024-003','2024-03-18','Pediatric Dental Group','Child 1',280,140,'','Processing','','Cavity filling – awaiting EOB'],
  ['CLM-2024-004','2024-04-05','City ER','Self',2800,1950,'','Approved','2024-05-10','ER visit – laceration treatment approved; check pending'],
  ['CLM-2024-005','2024-05-15','MindCare Therapy','Spouse',800,480,'','Submitted','','Monthly therapy sessions Jan-Apr – manually submitted'],
  ['CLM-2024-006','2024-06-20','Allergy Associates','Child 2',850,0,'','Denied','2024-07-30','Allergy testing denied – filing appeal with insurer'],
  ['CLM-2024-007','2024-07-05','Memorial Hospital','Spouse',6500,5200,'','Paid','2024-08-15','Outpatient knee procedure – EOB confirmed'],
  ['CLM-2024-008','2024-08-10','Radiology Partners','Self',2200,1760,'','Paid','2024-09-20','MRI left knee – 80% covered'],
  ['CLM-2024-009','2024-09-25','Urgent Care Plus','Child 1',375,225,'','Processing','','Broken finger urgent care + x-rays'],
  ['CLM-2024-010','2024-10-08','Gastro Partners','Self',3200,2880,'','Paid','2024-11-12','Colonoscopy – preventive; 90% covered'],
  ['CLM-2024-011','2024-11-15','ClearVision Optometry','Child 2',280,140,'','Approved','2024-12-05','Pediatric eye exam + lenses approved'],
  ['CLM-2024-012','2024-12-10','Sleep Solutions Center','Spouse',1450,725,'','Submitted','','CPAP machine – durable medical equipment claim'],
];

(async () => {
  const reqs = [];

  // Row 0: Title (col A standalone for col-freeze compatibility)
  reqs.push({ mergeCells: { range: gridRange(INS, 0, 1, 1, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(INS, 0, 1, 0, 1),
      cell: {
        userEnteredValue: { stringValue: '📋' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepTeal),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(INS, 0, 1, 1, 10),
      cell: {
        userEnteredValue: { stringValue: 'Insurance Claim Tracker' },
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
    range: { sheetId: INS, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Row 1: Headers
  reqs.push({
    repeatCell: {
      range: gridRange(INS, 1, 2, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.amber),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: INS, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});

  // Data rows 2-13
  for (let r = 2; r < 14; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(INS, r, r+1, 0, 10),
        cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.inputBg : C.altRow), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: INS, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
  }
  reqs.push({
    repeatCell: {
      range: gridRange(INS, 14, 200, 0, 10),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat',
    },
  });

  // Patient Responsibility col G (index 6): formula bg
  reqs.push({ repeatCell: {
    range: gridRange(INS, 2, 200, 6, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg) } },
    fields: 'userEnteredFormat',
  }});

  // Col widths
  for (const [i, px] of WIDTHS.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: INS, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  // Number formats
  reqs.push({ repeatCell: {
    range: gridRange(INS, 2, 200, 1, 2),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(INS, 2, 200, 8, 9),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  for (const col of [4, 5, 6]) {
    reqs.push({ repeatCell: {
      range: gridRange(INS, 2, 200, col, col+1),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  }
  reqs.push({ repeatCell: {
    range: gridRange(INS, 2, 200, 3, 4),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(INS, 2, 200, 7, 8),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});

  await batchUpdate(id, reqs, 'ins-format');

  // Values
  const data = [];
  data.push({ range: `'📋 Insurance Claims'!A2:J2`, values: [HEADERS] });

  for (let i = 0; i < sampleData.length; i++) {
    const row = i + 3;
    const [claimNo, dateSubmit, provider, participant, billed, covered, , status, dateResolved, notes] = sampleData[i];
    const respFormula = `=IFERROR(E${row}-F${row},"")`;
    data.push({
      range: `'📋 Insurance Claims'!A${row}:J${row}`,
      values: [[claimNo, dateSubmit, provider, participant, billed, covered, respFormula, status, dateResolved, notes]],
    });
  }

  // Formula for remaining empty rows
  const formulaRows = [];
  for (let row = 15; row <= 200; row++) {
    formulaRows.push(['','','','','','', `=IFERROR(IF(E${row}="","",E${row}-F${row}),"")`, '','','']);
  }
  data.push({ range: `'📋 Insurance Claims'!A15`, values: formulaRows });

  await valuesBatchUpdate(id, data, 'ins-values');
  console.log('Insurance Claims complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
