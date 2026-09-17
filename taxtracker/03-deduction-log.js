'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const LOG = sheetMap['📝 Deduction Log'];

// Cols: A=Date | B=Category | C=Description | D=Amount | E=Limit% | F=Doc? | G=Net Deductible | H=Notes

const HEADERS = ['Date','Category','Description','Amount','Limit %','Doc?','Net Deductible','Notes'];
const COLS    = ['A','B','C','D','E','F','G','H'];
const WIDTHS  = [90, 160, 220, 90, 70, 55, 110, 180];

// Sample data: 50 rows across 2024 (35), 2025 (10), 2026 (5)
const sampleData = [
  // 2024 data (35 rows)
  ['2024-01-08','Office Expenses','Printer paper and ink cartridges',89.50,true],
  ['2024-01-15','Meals & Entertainment','Client lunch - new contract discussion',67.20,true],
  ['2024-01-22','Professional Services','Accountant consultation - Q4 filing',250.00,true],
  ['2024-02-03','Subscriptions & Software','Adobe Creative Cloud annual plan',599.88,true],
  ['2024-02-10','Marketing & Advertising','Facebook ad campaign - product launch',340.00,true],
  ['2024-02-18','Meals & Entertainment','Team dinner - project celebration',148.60,false],
  ['2024-02-28','Equipment & Technology','External hard drive for backups',129.99,true],
  ['2024-03-05','Office Expenses','Desk organizer and cable management',54.30,true],
  ['2024-03-12','Professional Services','Legal review of client contract',400.00,true],
  ['2024-03-20','Travel','Conference hotel - 2 nights',389.00,true],
  ['2024-04-02','Insurance','Business liability insurance - Q2',312.50,true],
  ['2024-04-09','Education & Training','Online course - digital marketing',197.00,true],
  ['2024-04-17','Meals & Entertainment','Prospect coffee meeting',18.40,true],
  ['2024-04-25','Subscriptions & Software','Zoom Pro annual subscription',149.90,true],
  ['2024-05-01','Office Expenses','Wireless keyboard and mouse',79.95,true],
  ['2024-05-08','Marketing & Advertising','Google Ads campaign - summer push',520.00,true],
  ['2024-05-15','Travel','Flight to industry conference',418.00,true],
  ['2024-05-22','Professional Services','Bookkeeper - monthly fee',300.00,true],
  ['2024-06-04','Equipment & Technology','Ring light for video calls',89.00,false],
  ['2024-06-11','Meals & Entertainment','Client dinner - contract renewal',189.75,true],
  ['2024-06-19','Rent & Utilities','Coworking space membership - June',450.00,true],
  ['2024-07-02','Office Expenses','Whiteboard and markers',67.49,true],
  ['2024-07-09','Insurance','Business liability insurance - Q3',312.50,true],
  ['2024-07-16','Education & Training','Workshop - advanced Excel',95.00,true],
  ['2024-07-24','Subscriptions & Software','QuickBooks Self-Employed monthly',15.00,true],
  ['2024-08-05','Marketing & Advertising','LinkedIn Premium - 3 months',359.97,true],
  ['2024-08-12','Travel','Mileage - client site visits Aug',0,false],
  ['2024-08-20','Professional Services','Bookkeeper - monthly fee',300.00,true],
  ['2024-09-03','Office Expenses','Laptop stand and monitor arm',149.99,true],
  ['2024-09-10','Meals & Entertainment','Networking lunch - industry group',42.00,true],
  ['2024-09-18','Equipment & Technology','Webcam upgrade for client calls',109.00,true],
  ['2024-10-02','Insurance','Business liability insurance - Q4',312.50,true],
  ['2024-10-14','Marketing & Advertising','Trade show booth deposit',750.00,true],
  ['2024-11-06','Professional Services','Accountant - year-end planning',350.00,true],
  ['2024-12-19','Office Expenses','Office supplies - year-end restock',122.30,false],
  // 2025 data (10 rows)
  ['2025-01-10','Office Expenses','Standing desk converter',289.00,true],
  ['2025-01-20','Subscriptions & Software','Notion team plan - annual',192.00,true],
  ['2025-02-05','Meals & Entertainment','Client strategy lunch',94.30,true],
  ['2025-02-18','Professional Services','Bookkeeper - monthly fee',325.00,true],
  ['2025-03-03','Marketing & Advertising','Instagram ads - spring campaign',480.00,true],
  ['2025-04-07','Education & Training','Masterclass annual membership',180.00,false],
  ['2025-05-12','Equipment & Technology','Portable SSD - 2TB',149.99,true],
  ['2025-06-20','Travel','Conference airfare + hotel',892.00,true],
  ['2025-09-15','Insurance','Business liability insurance - Q3',325.00,true],
  ['2025-11-28','Office Expenses','End-of-year office refresh',215.40,true],
  // 2026 data (5 rows)
  ['2026-01-05','Professional Services','Accountant - annual filing prep',450.00,true],
  ['2026-02-14','Subscriptions & Software','Design tool suite - annual',348.00,true],
  ['2026-03-10','Meals & Entertainment','Client dinner - new partnership',156.00,true],
  ['2026-04-02','Marketing & Advertising','Q2 ad budget - Google + Meta',640.00,true],
  ['2026-05-18','Office Expenses','Ergonomic chair upgrade',389.99,true],
];

(async () => {
  const reqs = [];

  // Row 0: Title banner
  reqs.push({
    mergeCells: { range: gridRange(LOG, 0, 1, 0, 8), mergeType: 'MERGE_ALL' },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(LOG, 0, 1, 0, 8),
      cell: {
        userEnteredValue: { stringValue: '📝 Deduction Log' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepSapphire),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER',
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: LOG, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 },
    fields: 'pixelSize',
  }});

  // Row 1: Headers
  reqs.push({
    repeatCell: {
      range: gridRange(LOG, 1, 2, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.warmGold),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: LOG, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 30 },
    fields: 'pixelSize',
  }});

  // Data rows: alternating background
  for (let r = 2; r < 52; r++) {
    const bg = r % 2 === 0 ? C.inputBg : C.altRow;
    reqs.push({
      repeatCell: {
        range: gridRange(LOG, r, r+1, 0, 8),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
  }
  // Empty rows beyond data
  reqs.push({
    repeatCell: {
      range: gridRange(LOG, 52, 200, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat',
    },
  });

  // Col widths
  for (const [i, px] of WIDTHS.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: LOG, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px },
      fields: 'pixelSize',
    }});
  }

  // Format D (Amount) col C: date, D: currency, E: percent, G: currency
  reqs.push({ repeatCell: {
    range: gridRange(LOG, 2, 200, 0, 1),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(LOG, 2, 200, 3, 4),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(LOG, 2, 200, 4, 5),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(LOG, 2, 200, 5, 6),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(LOG, 2, 200, 6, 7),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' }, backgroundColor: hex(C.formulaBg) } },
    fields: 'userEnteredFormat',
  }});

  // Gold border on input cells (A-D, F, H)
  const borderStyle = { style: 'SOLID', color: hex(C.border) };
  for (const col of [0,1,2,3,5,7]) {
    reqs.push({ repeatCell: {
      range: gridRange(LOG, 2, 200, col, col+1),
      cell: { userEnteredFormat: { borders: { bottom: borderStyle } } },
      fields: 'userEnteredFormat.borders',
    }});
  }

  await batchUpdate(id, reqs, 'log-format');

  // Values: headers
  const data = [];
  data.push({ range: `'📝 Deduction Log'!A2:H2`, values: [HEADERS] });

  // Sample data rows (start at row 3, index 2)
  for (let i = 0; i < sampleData.length; i++) {
    const row = i + 3; // 1-based sheet row
    const [date, cat, desc, amount, documented] = sampleData[i];
    // E = limit lookup formula, G = net deductible formula
    const limitFormula = `=IFERROR(VLOOKUP(B${row},'📋 Reference Data'!$A:$B,2,FALSE),1)`;
    const netFormula = `=IFERROR(D${row}*E${row},"")`;
    data.push({
      range: `'📝 Deduction Log'!A${row}:H${row}`,
      values: [[date, cat, desc, amount === 0 ? '' : amount, limitFormula, documented, netFormula, '']],
    });
  }

  // Formulas for empty rows 53-200
  const formulaRows = [];
  for (let row = sampleData.length + 3; row <= 200; row++) {
    formulaRows.push([
      '', '', '', '',
      `=IFERROR(IF(B${row}="","",VLOOKUP(B${row},'📋 Reference Data'!$A:$B,2,FALSE)),"")`,
      '',
      `=IFERROR(IF(D${row}="","",D${row}*E${row}),"")`,
      '',
    ]);
  }
  if (formulaRows.length) {
    data.push({ range: `'📝 Deduction Log'!A${sampleData.length+3}`, values: formulaRows });
  }

  await valuesBatchUpdate(id, data, 'log-values');
  console.log('Deduction Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
