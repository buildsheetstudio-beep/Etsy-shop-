'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const MIL = sheetMap['🚗 Mileage Log'];

// Cols: A=Date | B=Purpose | C=Business Miles | D=Tax Year | E=IRS Rate | F=Deductible | G=Notes

const HEADERS = ['Date','Purpose','Business Miles','Tax Year','IRS Rate $/mi','Deductible Amount','Notes'];
const WIDTHS  = [90, 220, 100, 80, 90, 110, 180];

const sampleData = [
  ['2024-01-15','Client meeting - downtown office',34.2,2024],
  ['2024-01-29','Supply run - office depot',18.7,2024],
  ['2024-02-08','Networking event - chamber of commerce',22.4,2024],
  ['2024-02-20','Client site visit - project kickoff',41.0,2024],
  ['2024-03-05','Bank deposit + post office',12.3,2024],
  ['2024-03-19','Industry meetup - evening event',28.6,2024],
  ['2024-04-10','Accountant office - tax prep meeting',15.8,2024],
  ['2024-04-24','Client presentation - uptown',38.5,2024],
  ['2024-05-07','Coworking space commute (client-required)',0,2024],
  ['2024-05-21','Conference day 1 - venue travel',19.2,2024],
  ['2024-06-03','Conference day 2 - venue travel',19.2,2024],
  ['2024-06-17','New client site tour',45.7,2024],
  ['2024-07-09','Vendor meeting - supplier facility',52.3,2024],
  ['2024-08-14','Training workshop - business district',24.1,2024],
  ['2024-09-23','Year-end client review meetings',67.8,2024],
];

(async () => {
  const reqs = [];

  // Row 0: Title banner
  reqs.push({ mergeCells: { range: gridRange(MIL, 0, 1, 0, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(MIL, 0, 1, 0, 7),
      cell: {
        userEnteredValue: { stringValue: '🚗 Mileage Log' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepSapphire),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: MIL, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Row 1: Headers
  reqs.push({
    repeatCell: {
      range: gridRange(MIL, 1, 2, 0, 7),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.mutedSage),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: MIL, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  // Alternating rows
  for (let r = 2; r < 17; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(MIL, r, r+1, 0, 7),
        cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.inputBg : C.altRow), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
  }
  reqs.push({
    repeatCell: {
      range: gridRange(MIL, 17, 200, 0, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat',
    },
  });

  // Col widths
  for (const [i, px] of WIDTHS.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: MIL, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  // Number formats
  reqs.push({ repeatCell: {
    range: gridRange(MIL, 2, 200, 0, 1),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(MIL, 2, 200, 2, 3),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0.0' }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(MIL, 2, 200, 3, 4),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(MIL, 2, 200, 4, 5),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '$0.000' }, backgroundColor: hex(C.formulaBg), horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(MIL, 2, 200, 5, 6),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' }, backgroundColor: hex(C.formulaBg) } },
    fields: 'userEnteredFormat',
  }});

  await batchUpdate(id, reqs, 'mileage-format');

  // Values
  const data = [];
  data.push({ range: `'🚗 Mileage Log'!A2:G2`, values: [HEADERS] });

  for (let i = 0; i < sampleData.length; i++) {
    const row = i + 3;
    const [date, purpose, miles, year] = sampleData[i];
    const rateFormula = `=IFERROR(VLOOKUP(D${row},'📋 Reference Data'!$C:$D,2,FALSE),0)`;
    const dedFormula  = `=IFERROR(IF(C${row}="","",C${row}*E${row}),"")`;
    data.push({
      range: `'🚗 Mileage Log'!A${row}:G${row}`,
      values: [[date, purpose, miles === 0 ? '' : miles, year, rateFormula, dedFormula, '']],
    });
  }

  // Formulas for empty rows
  const formulaRows = [];
  for (let row = sampleData.length + 3; row <= 200; row++) {
    formulaRows.push([
      '', '', '', '',
      `=IFERROR(IF(D${row}="","",VLOOKUP(D${row},'📋 Reference Data'!$C:$D,2,FALSE)),"")`,
      `=IFERROR(IF(C${row}="","",C${row}*E${row}),"")`,
      '',
    ]);
  }
  if (formulaRows.length) {
    data.push({ range: `'🚗 Mileage Log'!A${sampleData.length+3}`, values: formulaRows });
  }

  await valuesBatchUpdate(id, data, 'mileage-values');

  // Summary box (rows 0-1, cols I-L already outside data)
  const summaryReqs = [];
  summaryReqs.push({ mergeCells: { range: gridRange(MIL, 0, 1, 8, 12), mergeType: 'MERGE_ALL' } });
  summaryReqs.push({
    repeatCell: {
      range: gridRange(MIL, 0, 1, 8, 12),
      cell: {
        userEnteredValue: { stringValue: '2024 Mileage Summary' },
        userEnteredFormat: {
          backgroundColor: hex(C.mutedSage),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  const summaryData = [
    ['Total Miles', `=IFERROR(SUMIF('🚗 Mileage Log'!$D$3:$D$200,2024,'🚗 Mileage Log'!$C$3:$C$200),0)`],
    ['IRS Rate', `=IFERROR(VLOOKUP(2024,'📋 Reference Data'!$C:$D,2,FALSE),0)`],
    ['Total Deductible', `=IFERROR(SUMIF('🚗 Mileage Log'!$D$3:$D$200,2024,'🚗 Mileage Log'!$F$3:$F$200),0)`],
  ];
  for (let r = 0; r < summaryData.length; r++) {
    summaryReqs.push({
      repeatCell: {
        range: gridRange(MIL, r+1, r+2, 8, 10),
        cell: { userEnteredFormat: { backgroundColor: hex(C.lightSage) } },
        fields: 'userEnteredFormat',
      },
    });
  }
  await batchUpdate(id, summaryReqs, 'mileage-summary-fmt');

  const sumVals = [];
  sumVals.push({ range: `'🚗 Mileage Log'!I2:J2`, values: [['Label', 'Value']] });
  for (const [i, [label, formula]] of summaryData.entries()) {
    sumVals.push({ range: `'🚗 Mileage Log'!I${i+3}:J${i+3}`, values: [[label, formula]] });
  }
  await valuesBatchUpdate(id, sumVals, 'mileage-summary-vals');

  console.log('Mileage Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
