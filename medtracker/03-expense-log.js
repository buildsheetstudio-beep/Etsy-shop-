'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const EXP = sheetMap['🧾 Expense Log'];

// A=Date | B=Participant | C=Category | D=Provider/Description | E=Amount
// F=Payment Method | G=Receipt/Doc Ref | H=Claim Status | I=Notes

const HEADERS = ['Date','Participant','Category','Provider / Description','Amount','Payment Method','Receipt / Doc Ref','Claim / Reimb. Status','Notes'];
const WIDTHS  = [90, 90, 150, 210, 90, 120, 110, 130, 180];

// 40 sample rows: 30 in 2024, 7 in 2025, 3 in 2026
// date, participant, category, description, amount, paymentMethod, receipt, status, notes
const sampleData = [
  // 2024 — 30 rows
  ['2024-01-08','Self','Doctor Visits','Dr. Smith – Annual physical',250,'HSA','REC-2024-001','Paid',''],
  ['2024-01-15','Spouse','Dental Care','Dr. Jones Dental – Routine cleaning',180,'FSA','REC-2024-002','Paid',''],
  ['2024-01-22','Child 1','Preventive Care','Pediatric wellness check',0,'Insurance Direct Pay','REC-2024-003','Paid','No cost with insurance'],
  ['2024-02-05','Self','Prescription Medications','Metformin 90-day supply',45,'HSA','REC-2024-004','Paid',''],
  ['2024-02-12','Child 2','Doctor Visits','Sick visit – ear infection',120,'Credit Card','REC-2024-005','Submitted','Pending EOB'],
  ['2024-02-20','Self','Lab/Diagnostic Tests','Annual blood panel',180,'HSA','REC-2024-006','Paid',''],
  ['2024-03-04','Spouse','Vision Care','Eye exam and new lenses',320,'FSA','REC-2024-007','Paid',''],
  ['2024-03-15','Child 1','Dental Care','Cavity filling',280,'Credit Card','REC-2024-008','Processing','Awaiting EOB'],
  ['2024-03-22','Self','Therapy/Mental Health','Monthly therapy session',150,'HSA','REC-2024-009','Paid',''],
  ['2024-04-01','Self','Emergency Services','ER visit – hand laceration',850,'Credit Card','REC-2024-010','Approved','Awaiting payment check'],
  ['2024-04-08','Spouse','Prescription Medications','Cetirizine – allergy medication',65,'FSA','REC-2024-011','Paid',''],
  ['2024-04-16','Child 2','Lab/Diagnostic Tests','Allergy testing panel',395,'Insurance Direct Pay','REC-2024-012','Processing','Complex multi-allergen test'],
  ['2024-05-02','Self','Preventive Care','Mammogram screening',0,'Insurance Direct Pay','REC-2024-013','Paid','Fully covered preventive care'],
  ['2024-05-09','Spouse','Therapy/Mental Health','Couples counseling session',200,'Credit Card','REC-2024-014','Submitted','Out-of-network; submitting manually'],
  ['2024-05-20','Child 1','Doctor Visits','Follow-up visit',90,'Check','REC-2024-015','Paid',''],
  ['2024-06-03','Self','Prescription Medications','Lisinopril – blood pressure',55,'HSA','REC-2024-016','Paid',''],
  ['2024-06-10','Spouse','Vision Care','Progressive lenses – new glasses',450,'FSA','REC-2024-017','Paid',''],
  ['2024-06-17','Self','Dental Care','Crown replacement',1200,'Credit Card','REC-2024-018','Denied','Cosmetic exclusion – appealing'],
  ['2024-06-25','Child 2','Doctor Visits','Sick visit – strep throat',130,'Debit Card','REC-2024-019','Paid',''],
  ['2024-07-08','Self','Medical Equipment','Blood pressure monitor',85,'HSA','REC-2024-020','Paid',''],
  ['2024-07-15','Spouse','Hospital Services','Outpatient knee procedure',2400,'Insurance Direct Pay','REC-2024-021','Paid',''],
  ['2024-07-22','Child 1','Prescription Medications','ADHD medication monthly',180,'Credit Card','REC-2024-022','Processing','Prior auth required'],
  ['2024-08-05','Self','Lab/Diagnostic Tests','MRI – left knee',1400,'Credit Card','REC-2024-023','Approved','Approved; EOB mailed'],
  ['2024-08-12','Spouse','Preventive Care','Flu vaccination',0,'Insurance Direct Pay','REC-2024-024','Paid',''],
  ['2024-08-19','Child 2','Dental Care','Orthodontist consultation',150,'FSA','REC-2024-025','Paid',''],
  ['2024-09-03','Self','Doctor Visits','Dermatologist – skin cancer screening',175,'HSA','REC-2024-026','Paid',''],
  ['2024-09-10','Spouse','Medical Equipment','CPAP machine',950,'FSA','REC-2024-027','Paid',''],
  ['2024-09-17','Child 1','Emergency Services','Urgent care – broken finger',375,'Credit Card','REC-2024-028','Submitted','X-rays included in bill'],
  ['2024-10-01','Self','Hospital Services','Colonoscopy (preventive)',1800,'Insurance Direct Pay','REC-2024-029','Paid',''],
  ['2024-10-08','Spouse','Therapy/Mental Health','Individual therapy session',150,'Cash','REC-2024-030','Not Submitted','Out-of-network; out-of-pocket'],
  // 2025 — 7 rows
  ['2025-01-10','Self','Doctor Visits','Annual physical – Dr. Smith',275,'HSA','REC-2025-001','Paid',''],
  ['2025-02-14','Child 2','Dental Care','Braces installation – phase 1',1800,'Credit Card','REC-2025-002','Processing','Orthodontic treatment plan'],
  ['2025-03-22','Spouse','Prescription Medications','New prescription refill',78,'FSA','REC-2025-003','Paid',''],
  ['2025-05-08','Self','Lab/Diagnostic Tests','Thyroid panel bloodwork',220,'HSA','REC-2025-004','Paid',''],
  ['2025-07-14','Child 1','Preventive Care','Annual sports physical',0,'Insurance Direct Pay','REC-2025-005','Paid','Required for school sports'],
  ['2025-09-05','Self','Therapy/Mental Health','Quarterly therapy session',150,'HSA','REC-2025-006','Paid',''],
  ['2025-11-20','Spouse','Vision Care','Annual eye exam',280,'FSA','REC-2025-007','Paid',''],
  // 2026 — 3 rows
  ['2026-01-15','Self','Doctor Visits','Annual physical – Dr. Smith',290,'HSA','REC-2026-001','Paid',''],
  ['2026-03-10','Child 2','Prescription Medications','Asthma inhaler refill',105,'Debit Card','REC-2026-002','Paid',''],
  ['2026-05-22','Spouse','Dental Care','Semi-annual cleaning',195,'FSA','REC-2026-003','Submitted',''],
];

(async () => {
  const reqs = [];

  // Row 0: Title — col A standalone for col-freeze compatibility, B:I merged
  reqs.push({ mergeCells: { range: gridRange(EXP, 0, 1, 1, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(EXP, 0, 1, 0, 1),
      cell: {
        userEnteredValue: { stringValue: '🧾' },
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
      range: gridRange(EXP, 0, 1, 1, 9),
      cell: {
        userEnteredValue: { stringValue: 'Medical Expense Log' },
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
    range: { sheetId: EXP, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Row 1: Headers
  reqs.push({
    repeatCell: {
      range: gridRange(EXP, 1, 2, 0, 9),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.warmCoral),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: EXP, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});

  // Alternating data rows
  for (let r = 2; r < 42; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(EXP, r, r+1, 0, 9),
        cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.inputBg : C.altRow), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
  }
  reqs.push({
    repeatCell: {
      range: gridRange(EXP, 42, 200, 0, 9),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat',
    },
  });

  // Col widths
  for (const [i, px] of WIDTHS.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: EXP, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  // Number formats
  reqs.push({ repeatCell: {
    range: gridRange(EXP, 2, 200, 0, 1),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(EXP, 2, 200, 4, 5),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  // Formula bg for no cols here — all are input; just status col get coral border
  const coralBorder = { style: 'SOLID', color: hex(C.border) };
  for (const col of [0,1,2,3,4,5,6,7,8]) {
    reqs.push({ repeatCell: {
      range: gridRange(EXP, 2, 200, col, col+1),
      cell: { userEnteredFormat: { borders: { bottom: coralBorder } } },
      fields: 'userEnteredFormat.borders',
    }});
  }
  // Input cells: ivory bg with coral border for A,B,C,D,E,F,G,H,I
  // Formula bg for G (Receipt) is already input bg; just style H (status) differently
  // H (claim status) col: center align
  reqs.push({ repeatCell: {
    range: gridRange(EXP, 2, 200, 7, 8),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(EXP, 2, 200, 1, 2),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});

  await batchUpdate(id, reqs, 'expense-format');

  // Values
  const data = [];
  data.push({ range: `'🧾 Expense Log'!A2:I2`, values: [HEADERS] });

  for (let i = 0; i < sampleData.length; i++) {
    const row = i + 3;
    data.push({ range: `'🧾 Expense Log'!A${row}:I${row}`, values: [sampleData[i]] });
  }

  await valuesBatchUpdate(id, data, 'expense-values');
  console.log('Expense Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
