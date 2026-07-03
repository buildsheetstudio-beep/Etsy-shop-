'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BUD = sheetMap['💰 Event Budget'];

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: BUD, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: BUD, dimension: 'ROWS', startIndex: 1, endIndex: 55 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // Column widths: A(30) B(EventName 200) C(Category 120) D(Vendor 180) E(Budgeted 110) F(Actual 110) G(Variance 110) H(Paid 80) I(PayStatus 120) J(Invoice# 100) K(Notes 200) + KPI L-M
  const colW = [30, 200, 120, 180, 110, 110, 110, 80, 120, 100, 200, 20, 160, 120];
  colW.forEach((w,i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: BUD, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title row A1:K1 merged
  reqs.push({ mergeCells: { range: gridRange(BUD, 0, 1, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(BUD, 0, 1, 0, 11),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.sageGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // KPI block L1:N1 header
  reqs.push({ mergeCells: { range: gridRange(BUD, 0, 1, 11, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(BUD, 0, 1, 11, 14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.antiqueGold),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Section header — Allocation Table row 2
  reqs.push({ mergeCells: { range: gridRange(BUD, 1, 2, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(BUD, 1, 2, 0, 11),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.rose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Header row 3
  reqs.push({ repeatCell: {
    range: gridRange(BUD, 2, 3, 0, 11),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlush),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Allocation data rows 4-16
  for (let r = 3; r < 16; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.champagne;
    reqs.push({ repeatCell: {
      range: gridRange(BUD, r, r+1, 0, 11),
      cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 10 }, verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // Totals row 17 (index 16)
  reqs.push({ repeatCell: {
    range: gridRange(BUD, 16, 17, 0, 11),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.antiqueGold),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Expense Log section header row 18 (index 17)
  reqs.push({ mergeCells: { range: gridRange(BUD, 17, 18, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(BUD, 17, 18, 0, 11),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.rose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Expense log header row 19 (index 18)
  reqs.push({ repeatCell: {
    range: gridRange(BUD, 18, 19, 0, 11),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlush),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Expense log data rows 20-55
  for (let r = 19; r < 55; r++) {
    const bg = r % 2 === 0 ? C.champagne : C.ivory;
    reqs.push({ repeatCell: {
      range: gridRange(BUD, r, r+1, 0, 11),
      cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 10 }, verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // Currency format — E, F, G (budgeted, actual, variance) in allocation (rows 4-17)
  [4, 5, 6].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(BUD, 3, 17, c, c+1),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(BUD, 19, 55, c, c+1),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  });

  // Variance col G — formula col (paleRose)
  reqs.push({ repeatCell: { range: gridRange(BUD, 3, 17, 6, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.paleRose) } }, fields: 'userEnteredFormat.backgroundColor' }});
  reqs.push({ repeatCell: { range: gridRange(BUD, 19, 55, 6, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.paleRose) } }, fields: 'userEnteredFormat.backgroundColor' }});

  // Date format col J in expense log (index 9)
  reqs.push({ repeatCell: {
    range: gridRange(BUD, 19, 55, 9, 10),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});

  // Center align columns
  [0, 7, 8, 9].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(BUD, 2, 55, c, c+1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat.horizontalAlignment' }});
  });

  // KPI block styling L2:N10
  const kpiLabels = [
    ['Total Budget','',''],
    ['Total Spent','',''],
    ['Remaining','',''],
    ['% Spent','',''],
    ['# Events','',''],
    ['Paid in Full','',''],
    ['Outstanding','',''],
    ['Overdue','',''],
  ];
  for (let r = 1; r < 9; r++) {
    const bg = r % 2 === 0 ? C.champagne : C.ivory;
    // Label col L (index 11)
    reqs.push({ repeatCell: {
      range: gridRange(BUD, r, r+1, 11, 12),
      cell: { userEnteredFormat: { backgroundColor: hex(C.dustyRose), textFormat: { foregroundColor: hex(C.darkPlum), bold: true, fontSize: 9 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'LEFT' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
    }});
    // Value col M (index 12)
    reqs.push({ repeatCell: {
      range: gridRange(BUD, r, r+1, 12, 13),
      cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.antiqueGold), bold: true, fontSize: 11 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
    }});
  }

  // Borders
  reqs.push({ updateBorders: {
    range: gridRange(BUD, 2, 17, 0, 11),
    innerHorizontal: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
  }});
  reqs.push({ updateBorders: {
    range: gridRange(BUD, 18, 55, 0, 11),
    innerHorizontal: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
  }});

  await batchUpdate(id, reqs, 'budget-format');

  // Values
  const data = [];
  data.push({ range: `'💰 Event Budget'!A1`, values: [['💰 Event Budget — Allocations & Expense Log']] });
  data.push({ range: `'💰 Event Budget'!L1`, values: [['📊 Budget KPIs']] });
  data.push({ range: `'💰 Event Budget'!A2`, values: [['  📁 Budget Allocation by Category']] });
  data.push({ range: `'💰 Event Budget'!A3:K3`, values: [['#','Event Name','Category','Vendor / Description','Budgeted','Actual','Variance','Paid?','Payment Status','Invoice #','Notes']] });

  // Budget allocation rows (rows 4-16 = index 3-15)
  const allocations = [
    [1,'Tech Summit 2025','Venue','Grand Convention Centre',8500,8500,'=E4-F4',true,'Paid in Full','INV-001','Venue hire + AV included'],
    [2,'Tech Summit 2025','Catering','Fresh Table Catering Co.',4200,4200,'=E5-F5',true,'Paid in Full','INV-002','Lunch + morning tea 150 pax'],
    [3,'Tech Summit 2025','Marketing','Print & Digital Agency',2000,1850,'=E6-F6',true,'Paid in Full','INV-003','Banners, programs, social'],
    [4,'Tech Summit 2025','Audio/Visual','ProAV Solutions',3500,0,'=E7-F7',false,'Not Paid','INV-004','Tech day-of setup'],
    [5,'Charity Gala 2025','Venue','Riverside Ballroom',12000,6000,'=E8-F8',false,'Deposit Paid','INV-010','50% deposit paid'],
    [6,'Charity Gala 2025','Catering','Prestige Catering',9000,0,'=E9-F9',false,'Not Paid','','Full dinner service 200 pax'],
    [7,'Charity Gala 2025','Entertainment','The Jazz Collective',4500,0,'=E10-F10',false,'Not Paid','','Live music 4 hours'],
    [8,'Charity Gala 2025','Decor','Bloom & Drape Florist',3500,1750,'=E11-F11',false,'Deposit Paid','INV-015','50% deposit'],
    [9,'Holiday Party 2025','Venue','The Loft Event Space',5500,0,'=E12-F12',false,'Not Paid','','End of year party'],
    [10,'Holiday Party 2025','Catering','TBC',3500,0,'=E13-F13',false,'Not Paid','',''],
    [11,'Investor Day 2025','Venue','HQ Boardroom',0,0,'=E14-F14',true,'Paid in Full','','Internal venue - free'],
    [12,'Annual Awards Night','Venue','The Grand Hotel',9500,4750,'=E15-F15',false,'Deposit Paid','INV-020','50% deposit paid'],
  ];

  data.push({ range: `'💰 Event Budget'!A4:K15`, values: allocations });

  // Totals row (row 17 = index 16)
  data.push({ range: `'💰 Event Budget'!A17:K17`, values: [['','TOTALS','','','=SUM(E4:E15)','=SUM(F4:F15)','=E17-F17','','','','']] });

  // Expense log section header
  data.push({ range: `'💰 Event Budget'!A18`, values: [['  📄 Expense Log — Individual Transactions']] });
  data.push({ range: `'💰 Event Budget'!A19:K19`, values: [['#','Event Name','Category','Vendor / Description','Amount Budgeted','Amount Paid','Variance','Paid?','Payment Status','Date Paid','Notes']] });

  const expenses = [
    [1,'Tech Summit 2025','Venue','Grand Convention Centre deposit',8500,8500,'=E20-F20',true,'Paid in Full','2025-06-01','Final invoice settled'],
    [2,'Tech Summit 2025','Catering','Fresh Table Catering Co.',4200,4200,'=E21-F21',true,'Paid in Full','2025-07-15','150 pax lunch + tea'],
    [3,'Tech Summit 2025','Marketing','Print & Digital Agency',2000,1850,'=E22-F22',true,'Paid in Full','2025-07-20','Under budget'],
    [4,'Charity Gala 2025','Venue','Riverside Ballroom — 50% deposit',12000,6000,'=E23-F23',false,'Deposit Paid','2025-07-01','Balance due 30 Sep'],
    [5,'Charity Gala 2025','Decor','Bloom & Drape Florist deposit',3500,1750,'=E24-F24',false,'Deposit Paid','2025-07-05',''],
    [6,'Product Launch — Nova X','Marketing','Social media ad spend',1500,1480,'=E25-F25',true,'Paid in Full','2025-07-28',''],
    [7,'Product Launch — Nova X','Audio/Visual','Livestream setup',800,850,'=E26-F26',true,'Paid in Full','2025-07-28','$50 over budget'],
    [8,'Team Building Day','Venue','Outdoor Adventure Park',2800,2800,'=E27-F27',true,'Paid in Full','2025-07-10',''],
    [9,'Marketing Workshop','Venue','Training Centre B',800,800,'=E28-F28',true,'Paid in Full','2025-08-05',''],
    [10,'Virtual Workshop — AI Trends','Audio/Visual','Zoom webinar + recording',500,480,'=E29-F29',true,'Paid in Full','2025-08-20',''],
    [11,'CEO Birthday Celebration','Catering','Private dining room catering',2200,2200,'=E30-F30',true,'Paid in Full','2025-07-25',''],
    [12,'Annual Awards Night','Venue','The Grand Hotel 50% deposit',9500,4750,'=E31-F31',false,'Deposit Paid','2025-08-15',''],
    [13,'5-Year Anniversary Dinner','Venue','Harbour View Restaurant deposit',4800,2400,'=E32-F32',false,'Deposit Paid','2025-08-20',''],
    [14,'5-Year Anniversary Dinner','Catering','Custom dinner menu',2500,0,'=E33-F33',false,'Not Paid','',''],
    [15,'Networking Drinks — Q4','Catering','Rooftop Bar drinks package',1800,0,'=E34-F34',false,'Not Paid','','Due 1 Oct'],
    [16,'Holiday Party 2025','Venue','The Loft Event Space deposit',5500,0,'=E35-F35',false,'Not Paid','','Not yet invoiced'],
    [17,'Investor Day 2025','Catering','Boardroom catering 20 pax',800,0,'=E36-F36',false,'Not Paid','',''],
    [18,'Tech Summit 2025','Audio/Visual','ProAV Solutions',3500,0,'=E37-F37',false,'Not Paid','','Due on day'],
  ];

  data.push({ range: `'💰 Event Budget'!A20:K37`, values: expenses });

  // KPI values
  data.push({ range: `'💰 Event Budget'!L2:M9`, values: [
    ['Total Budget','=SUM(E4:E15)'],
    ['Total Spent','=SUM(F4:F15)'],
    ['Remaining','=M2-M3'],
    ['% Spent','=IFERROR(M3/M2,0)'],
    ['# Events','=COUNTA(UNIQUE(B4:B15))-1'],
    ['Paid in Full','=COUNTIF(I20:I37,"Paid in Full")'],
    ['Outstanding','=COUNTIF(I20:I37,"Not Paid")'],
    ['Overdue','=COUNTIF(I20:I37,"Overdue")'],
  ]});

  await valuesBatchUpdate(id, data, 'budget-values');

  // Format KPI currency/pct
  const kpiReqs = [];
  kpiReqs.push({ repeatCell: {
    range: gridRange(BUD, 1, 4, 12, 13),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  kpiReqs.push({ repeatCell: {
    range: gridRange(BUD, 4, 5, 12, 13),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  await batchUpdate(id, kpiReqs, 'budget-kpi-fmt');

  console.log('Event Budget complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
