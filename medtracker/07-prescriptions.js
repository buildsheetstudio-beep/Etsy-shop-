'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const RX = sheetMap['💊 Prescription Tracker'];

// A=Medication Name | B=Participant | C=Prescribing Doctor | D=Dosage
// E=Refill Frequency (days) | F=Last Filled Date | G=Next Refill Due (formula: F+E)
// H=Pharmacy | I=Cost Per Fill | J=Insurance Covered (checkbox)

// Today = 2026-07-08
// CF: amber when G-TODAY()<=7 AND G>=TODAY()  (due within 7 days)
//     red   when G<TODAY()                     (overdue)
// Need at least one of each:
//   Overdue: last=2026-04-01, freq=90 → next=2026-06-30 < today ✓ (row 3)
//            last=2026-06-25, freq=10 → next=2026-07-05 < today ✓ (row 8)
//   Within 7 days: last=2026-06-09, freq=30 → next=2026-07-09 ✓ (row 2)
//                  last=2026-06-12, freq=30 → next=2026-07-12 ✓ (row 7)
//   OK (not due yet): others

const HEADERS = ['Medication Name','Participant','Prescribing Doctor','Dosage','Refill Frequency (days)','Last Filled Date','Next Refill Due','Pharmacy','Cost Per Fill','Ins. Covered?'];
const WIDTHS  = [170, 90, 140, 120, 120, 110, 110, 130, 100, 90];

const sampleData = [
  ['Metformin 500mg','Self','Dr. Smith','500mg twice daily',90,'2026-05-01','','CVS Pharmacy',45,true],
  ['Lisinopril 10mg','Self','Dr. Smith','10mg once daily',30,'2026-06-09','','CVS Pharmacy',18,true],
  ['Atorvastatin 20mg','Self','Dr. Smith','20mg nightly',90,'2026-04-01','','CVS Pharmacy',55,true],
  ['Cetirizine 10mg','Spouse','Dr. Lee','10mg once daily',30,'2026-07-03','','Walgreens',12,false],
  ['Escitalopram 10mg','Spouse','Dr. Chen','10mg once daily',30,'2026-06-20','','MindRx Pharmacy',35,true],
  ['Methylphenidate 20mg','Child 1','Dr. Kim','20mg once daily (school days)',30,'2026-07-01','','CVS Pharmacy',85,true],
  ['Montelukast 5mg','Child 2','Dr. Lee','5mg nightly',30,'2026-06-12','','Walgreens',28,true],
  ['Amoxicillin 250mg','Child 2','Dr. Kim','250mg twice daily x 10 days',10,'2026-06-25','','CVS Pharmacy',15,true],
  ['Levothyroxine 75mcg','Spouse','Dr. Patel','75mcg before breakfast',90,'2026-05-15','','Walgreens',22,true],
  ['Ibuprofen 800mg','Self','Dr. Martinez','800mg as needed (max 3/day)',30,'2026-06-30','','CVS Pharmacy',8,false],
];

(async () => {
  const reqs = [];

  // Row 0: Title
  reqs.push({ mergeCells: { range: gridRange(RX, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(RX, 0, 1, 0, 10),
      cell: {
        userEnteredValue: { stringValue: '💊 Prescription Refill Tracker' },
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
    range: { sheetId: RX, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Row 1: Headers
  reqs.push({
    repeatCell: {
      range: gridRange(RX, 1, 2, 0, 10),
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
    range: { sheetId: RX, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});

  // Data rows
  for (let r = 2; r < 12; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(RX, r, r+1, 0, 10),
        cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.inputBg : C.altRow), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: RX, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
  }
  reqs.push({
    repeatCell: {
      range: gridRange(RX, 12, 200, 0, 10),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat',
    },
  });

  // Col G (index 6): formula bg
  reqs.push({ repeatCell: {
    range: gridRange(RX, 2, 200, 6, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg) } },
    fields: 'userEnteredFormat',
  }});

  // Col widths
  for (const [i, px] of WIDTHS.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: RX, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  // Number formats
  reqs.push({ repeatCell: {
    range: gridRange(RX, 2, 200, 5, 7),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(RX, 2, 200, 8, 9),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(RX, 2, 200, 4, 5),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(RX, 2, 200, 1, 2),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(RX, 2, 200, 9, 10),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});

  await batchUpdate(id, reqs, 'rx-format');

  // Values
  const data = [];
  data.push({ range: `'💊 Prescription Tracker'!A2:J2`, values: [HEADERS] });

  for (let i = 0; i < sampleData.length; i++) {
    const row = i + 3;
    const rowData = [...sampleData[i]];
    // Formula for col G (index 6)
    rowData[6] = `=IFERROR(F${row}+E${row},"")`;
    data.push({ range: `'💊 Prescription Tracker'!A${row}:J${row}`, values: [rowData] });
  }

  // Formulas for empty rows
  const formulaRows = [];
  for (let row = 13; row <= 100; row++) {
    formulaRows.push(['','','','','','', `=IFERROR(IF(F${row}="","",F${row}+E${row}),"")`, '','',false]);
  }
  data.push({ range: `'💊 Prescription Tracker'!A13`, values: formulaRows });

  await valuesBatchUpdate(id, data, 'rx-values');
  console.log('Prescription Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
