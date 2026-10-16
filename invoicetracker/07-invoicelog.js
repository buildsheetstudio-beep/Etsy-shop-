'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const IL = sheetMap['Invoice Log'];
const S  = "'Invoice Log'";

const headers = [
  'Invoice ID','Invoice Number','Client ID','Client Name','Invoice Date','Due Date',
  'Year','Month','Project / Job','Currency','Subtotal','Discount','Shipping','Tax',
  'Invoice Total','Deposit Required','Payments Received','Balance Due','Payment Status',
  'Invoice Status','Days Until / Overdue','Last Payment Date','Next Follow-Up Date',
  'Follow-Up Status','Delivery Method','Sent Date','Paid Date','Notes'
];

// [InvoiceID, InvoiceNum, ClientID, ClientName, InvDate, DueDate, Project, Currency,
//  Subtotal, Discount, Shipping, Tax, Total, DepositReq, PaymentsRec, BalanceDue,
//  PayStatus, InvStatus, LastPayDate, NextFollowUp, FollowUpStatus, DeliveryMethod, SentDate, PaidDate, Notes]
const invoices = [
  // Paid invoices
  ['INV-2026-0001','INV-2026-0001','CL-001','Apex Digital Agency','Jan 5, 2026','Feb 4, 2026','Q1 Brand Strategy',  'USD',1950,0,   0,170.63,2120.63,0,    2120.63,0,    'Paid',           'Paid',          'Feb 2, 2026', '',            'Not Needed','Email','Jan 5, 2026', 'Feb 2, 2026', 'Paid on time'],
  ['INV-2026-0002','INV-2026-0002','CL-002','BlueSky Architecture','Jan 8, 2026','Jan 22, 2026','Office Logo Redesign','USD',850,  0,   0,0,     850,    0,    850,    0,    'Paid',           'Paid',          'Jan 20, 2026','',            'Not Needed','Email','Jan 8, 2026', 'Jan 20, 2026','Quick payment'],
  ['INV-2026-0003','INV-2026-0003','CL-004','Driftwood Publishing','Jan 12, 2026','Feb 11, 2026','Book Cover Series',  'USD',2550,0,   0,210.38,2760.38,0,    2760.38,0,    'Paid',           'Paid',          'Feb 10, 2026','',            'Not Needed','Email','Jan 12, 2026','Feb 10, 2026','3-book cover project'],
  ['INV-2026-0004','INV-2026-0004','CL-005','Emerald Coast Realty','Jan 15, 2026','Jan 30, 2026','Q1 Marketing Pack',  'USD',480,  24,  0,31.92, 487.92, 0,    487.92, 0,    'Paid',           'Paid',          'Jan 28, 2026','',            'Not Needed','Email','Jan 15, 2026','Jan 28, 2026','5% early pay discount applied'],
  ['INV-2026-0005','INV-2026-0005','CL-007','Gilded Thread Boutique','Jan 20, 2026','Jan 20, 2026','Winter Lookbook',    'USD',980,  0,   45,86.63,1111.63,0,    1111.63,0,    'Paid',           'Paid',          'Jan 20, 2026','',            'Not Needed','In Person','Jan 20, 2026','Jan 20, 2026','Due on receipt — paid on delivery'],
  ['INV-2026-0006','INV-2026-0006','CL-010','Juniper Lane Interiors','Feb 1, 2026','Mar 3, 2026', 'Staging Photo Session','USD',900,  0,   0,81,    981,    0,    981,    0,    'Paid',           'Paid',          'Mar 1, 2026', '',            'Not Needed','Email','Feb 1, 2026', 'Mar 1, 2026', 'Paid early'],
  ['INV-2026-0007','INV-2026-0007','CL-012','Luminary Learning Co.','Feb 3, 2026','Mar 5, 2026', 'Course Slide Decks',  'USD',1300,130, 0,84.68, 1254.68,0,    1254.68,0,    'Paid',           'Paid',          'Mar 4, 2026', '',            'Not Needed','Email','Feb 3, 2026', 'Mar 4, 2026', '10% discount on multi-deck order'],
  ['INV-2026-0008','INV-2026-0008','CL-014','Northgate Consulting','Feb 10, 2026','Mar 12, 2026','Q1 Report Design',   'USD',650,  0,   0,39,    689,    0,    689,    0,    'Paid',           'Paid',          'Mar 10, 2026','',            'Not Needed','Email','Feb 10, 2026','Mar 10, 2026','Paid net 30'],
  // Partially Paid
  ['INV-2026-0009','INV-2026-0009','CL-001','Apex Digital Agency','Feb 15, 2026','Mar 17, 2026','Brand Identity Refresh','USD',4200,0,   0,367.5, 4567.5, 1000, 1500,   3067.5,'Partially Paid','Partially Paid','Feb 20, 2026','Apr 1, 2026', 'Upcoming',  'Email','Feb 15, 2026','',            'Deposit + first partial payment received'],
  ['INV-2026-0010','INV-2026-0010','CL-008','Harbor View Hotels','Feb 20, 2026','Apr 6, 2026', 'Annual Brand Refresh', 'USD',5800,0,   0,362.5, 6162.5, 2000, 3500,   2662.5,'Deposit Paid',  'Partially Paid','Mar 1, 2026', 'Apr 10, 2026','Upcoming',  'Email','Feb 20, 2026','',            'Deposit paid — balance due after deliverables'],
  // Sent / Viewed
  ['INV-2026-0011','INV-2026-0011','CL-003','Cascade Wellness Spa','Mar 1, 2026', 'Mar 1, 2026', 'March Newsletter',    'USD',240,  0,   0,24,    264,    0,    0,      264,  'Unpaid',        'Sent',          '',            'Mar 15, 2026','Upcoming',  'Email','Mar 1, 2026', '',            'Monthly newsletter invoice'],
  ['INV-2026-0012','INV-2026-0012','CL-006','Flint & Steel Brewing Co.','Mar 5, 2026','Apr 4, 2026','Tap Room Spring Events','USD',760, 0, 0,22.04, 782.04, 0,    0,      782.04,'Unpaid',       'Viewed',        '',            'Apr 6, 2026', 'Upcoming',  'Printed','Mar 5, 2026','',            'Client viewed invoice Mar 8'],
  ['INV-2026-0013','INV-2026-0013','CL-015','Opal Events Co.','Mar 10, 2026','Mar 24, 2026','Wedding Stationery',   'USD',1150,0,   55,117.38,1322.38,400,  400,    922.38,'Deposit Paid',  'Sent',          'Mar 12, 2026','Mar 26, 2026','Upcoming',  'Email','Mar 10, 2026','',            'Deposit received at booking'],
  ['INV-2026-0014','INV-2026-0014','CL-009','Ironclad Fitness','Mar 15, 2026','Mar 22, 2026','Social Media Pack Q1', 'USD',420,  0,   0,23.52, 443.52, 0,    0,      443.52,'Unpaid',        'Sent',          '',            'Mar 24, 2026','Upcoming',  'Email','Mar 15, 2026','',            ''],
  // Overdue invoices
  ['INV-2026-0015','INV-2026-0015','CL-011','Keystone Law Group','Jan 28, 2026','Feb 12, 2026','Letterhead & Templates','USD',320, 0,  0,0,     320,    0,    0,      320,  'Unpaid',        'Overdue',       '',            'Mar 1, 2026', 'Overdue',   'Email','Jan 28, 2026','',            '2nd reminder sent Feb 20'],
  ['INV-2026-0016','INV-2026-0016','CL-017','Quartz Ridge Analytics','Feb 2, 2026','Apr 3, 2026','Dashboard Consulting',  'USD',2340,0,  0,0,     2340,   0,    0,      2340, 'Unpaid',        'Overdue',       '',            'Apr 10, 2026','Overdue',   'Email','Feb 2, 2026', '',            'Net 60 — past due, follow up scheduled'],
  ['INV-2026-0017','INV-2026-0017','CL-019','Sunrise Solar Solutions','Jan 10, 2026','Feb 9, 2026','Annual Report 2025',   'USD',1800,180, 0,142.56,1762.56,0,   0,      1762.56,'Unpaid',       'Overdue',       '',            'Mar 15, 2026','Overdue',   'Email','Jan 10, 2026','',            '10% discount — overdue 45+ days'],
  // Refunded
  ['INV-2026-0018','INV-2026-0018','CL-013','Maple & Birch Bakery','Dec 5, 2025','Dec 5, 2025', 'Logo Design Deposit',  'USD',425,  0,   0,37.13, 462.13, 0,    0,      0,    'Refunded',      'Refunded',      'Dec 8, 2025', '',            'Not Needed','In Person','Dec 5, 2025','',            'Client cancelled — full refund issued Dec 15'],
  // Draft
  ['INV-2026-0019','INV-2026-0019','CL-012','Luminary Learning Co.','Mar 20, 2026','Apr 19, 2026','Q2 Course Content',    'USD',1950,0,   0,141.38,2091.38,0,    0,      2091.38,'Unpaid',       'Draft',         '',            '',            'Not Needed','Email','',            '',            'Pending scope confirmation'],
  // Cancelled
  ['INV-2026-0020','INV-2026-0020','CL-020','Tidewater Creative','Nov 15, 2025','Dec 15, 2025','Rebranding Project',   'USD',3200,0,   0,169.6, 3369.6, 0,    0,      0,    'Cancelled',     'Cancelled',     '',            '',            'Not Needed','Email','',            '',            'Client paused — invoice voided'],
  // More paid/overdue/partial for aging coverage
  ['INV-2026-0021','INV-2026-0021','CL-016','Prairie Wind Photography','Feb 25, 2026','Mar 12, 2026','March Editing Sub',    'USD',320,  0,   0,17.6,  337.6,  0,    337.6,  0,    'Paid',          'Paid',          'Mar 10, 2026','',            'Not Needed','In Person','Feb 25, 2026','Mar 10, 2026','Monthly subscription'],
  ['INV-2026-0022','INV-2026-0022','CL-018','River Oak Counseling','Mar 1, 2026', 'Mar 1, 2026', 'Practice Brochure',    'USD',380,  0,   18,33.82,431.82, 0,    431.82, 0,    'Paid',          'Paid',          'Mar 3, 2026', '',            'Not Needed','Email','Mar 1, 2026', 'Mar 3, 2026', 'Paid same day'],
  ['INV-2026-0023','INV-2026-0023','CL-004','Driftwood Publishing','Mar 5, 2026', 'Apr 4, 2026', 'Spring Catalog',       'USD',720,  0,   0,59.4,  779.4,  0,    0,      779.4,'Unpaid',        'Sent',          '',            'Apr 6, 2026', 'Upcoming',  'Email','Mar 5, 2026', '',            ''],
  ['INV-2026-0024','INV-2026-0024','CL-005','Emerald Coast Realty','Mar 8, 2026', 'Mar 23, 2026','Q2 Property Flyers',   'USD',360,  0,   0,25.2,  385.2,  0,    385.2,  0,    'Paid',          'Paid',          'Mar 22, 2026','',            'Not Needed','Email','Mar 8, 2026', 'Mar 22, 2026',''],
  ['INV-2026-0025','INV-2026-0025','CL-001','Apex Digital Agency','Mar 12, 2026','Apr 11, 2026','Social Campaign Q2',   'USD',2100,0,   0,183.75,2283.75,0,    0,      2283.75,'Unpaid',       'Viewed',        '',            'Apr 13, 2026','Upcoming',  'Email','Mar 12, 2026','',            ''],
  ['INV-2026-0026','INV-2026-0026','CL-010','Juniper Lane Interiors','Mar 15, 2026','Apr 14, 2026','Spring Staging Pack',  'USD',1200,0,  0,108,   1308,   0,    0,      1308, 'Unpaid',        'Sent',          '',            'Apr 16, 2026','Upcoming',  'Email','Mar 15, 2026','',            ''],
  // Overpayment scenario
  ['INV-2026-0027','INV-2026-0027','CL-009','Ironclad Fitness','Mar 18, 2026','Mar 25, 2026','Q2 Social Pack',       'USD',420,  0,   0,23.52, 443.52, 0,    500,    0,    'Overpaid',      'Paid',          'Mar 24, 2026','',            'Not Needed','Email','Mar 18, 2026','Mar 24, 2026','Client overpaid — credit applied to next invoice'],
  // Large outstanding
  ['INV-2026-0028','INV-2026-0028','CL-008','Harbor View Hotels','Mar 20, 2026','May 4, 2026', 'Website Redesign Copy','USD',4800,0,   0,300,   5100,   1500, 1500,   3600, 'Deposit Paid',  'Partially Paid','Mar 25, 2026','May 6, 2026', 'Upcoming',  'Email','Mar 20, 2026','',            'Deposit only — balance due on completion'],
  // Various aging buckets (overdue by different amounts)
  ['INV-2025-0085','INV-2025-0085','CL-007','Gilded Thread Boutique','Nov 1, 2025','Dec 1, 2025','Holiday Lookbook',     'USD',1100,0,   0,97.63, 1197.63,0,    0,      1197.63,'Unpaid',       'Overdue',       '',            'Jan 15, 2026','Overdue',   'Email','Nov 1, 2025','',            '90+ days overdue'],
  ['INV-2025-0086','INV-2025-0086','CL-002','BlueSky Architecture','Dec 15, 2025','Jan 14, 2026','Signage Design',       'USD',580,  0,   0,0,     580,    0,    0,      580,  'Unpaid',        'Overdue',       '',            'Feb 1, 2026', 'Overdue',   'Email','Dec 15, 2025','',            '61-90 days overdue'],
  ['INV-2026-0029','INV-2026-0029','CL-014','Northgate Consulting','Feb 18, 2026','Mar 20, 2026','Presentation Template', 'USD',480,  0,   0,28.8,  508.8,  0,    0,      508.8,'Unpaid',        'Overdue',       '',            'Mar 25, 2026','Overdue',   'Email','Feb 18, 2026','',            '31-60 days overdue'],
  ['INV-2026-0030','INV-2026-0030','CL-003','Cascade Wellness Spa','Mar 2, 2026', 'Mar 17, 2026','Feb Newsletter',       'USD',240,  0,   0,24,    264,    0,    0,      264,  'Unpaid',        'Overdue',       '',            'Mar 22, 2026','Overdue',   'Email','Mar 2, 2026', '',            '1-30 days overdue'],
  ['INV-2026-0031','INV-2026-0031','CL-006','Flint & Steel Brewing Co.','Feb 28, 2026','Mar 30, 2026','Spring Menu Design',   'USD',340,  0,   0,9.86,  349.86, 0,    349.86, 0,    'Paid',          'Paid',          'Mar 28, 2026','',            'Not Needed','Printed','Feb 28, 2026','Mar 28, 2026',''],
  ['INV-2026-0032','INV-2026-0032','CL-013','Maple & Birch Bakery','Mar 10, 2026','Mar 10, 2026','Easter Menu Design',   'USD',185,  0,   0,16.43, 201.43, 0,    201.43, 0,    'Paid',          'Paid',          'Mar 10, 2026','',            'Not Needed','In Person','Mar 10, 2026','Mar 10, 2026','Due on receipt'],
  ['INV-2026-0033','INV-2026-0033','CL-015','Opal Events Co.','Mar 18, 2026','Apr 1, 2026', 'Invitation Suite',     'USD',890,  0,   25,92.55,1007.55,0,    0,      1007.55,'Unpaid',       'Sent',          '',            'Apr 3, 2026', 'Upcoming',  'Email','Mar 18, 2026','',            ''],
  ['INV-2026-0034','INV-2026-0034','CL-022','Luminary Learning Co.','Mar 22, 2026','Apr 21, 2026','April Newsletter Design','USD',240, 0,  0,17.4,  257.4,  0,    0,      257.4,'Unpaid',       'Draft',         '',            '',            'Not Needed','Email','',            '',            'Recurring invoice template'],
  ['INV-2026-0035','INV-2026-0035','CL-019','Sunrise Solar Solutions','Mar 25, 2026','Apr 24, 2026','Pitch Deck Design',    'USD',650,  0,   0,36.4,  686.4,  0,    0,      686.4,'Unpaid',       'Ready to Send', '',            '',            'Not Needed','Email','',            '',            'Ready to send after client approval'],
];

(async () => {
  const data = [];

  data.push({ range: `${S}!A1`, values: [['INVOICE LOG']] });
  data.push({ range: `${S}!A2`, values: [['Archive every issued invoice here. Copy data from Invoice Generator. Payment status and balance due update automatically from Payment Tracker.']] });
  data.push({ range: `${S}!A5`, values: [headers] });

  for (let i = 0; i < invoices.length; i++) {
    const r = i + 6;
    const [invId, invNum, clientId, clientName, invDate, dueDate, project, currency,
           subtotal, discount, shipping, tax, total, depositReq, paymentsRec, balanceDue,
           payStatus, invStatus, lastPayDate, nextFollowUp, followUpStatus, deliveryMethod, sentDate, paidDate, notes] = invoices[i];

    data.push({ range: `${S}!A${r}`, values: [[invId]] });
    data.push({ range: `${S}!B${r}`, values: [[invNum]] });
    data.push({ range: `${S}!C${r}`, values: [[clientId]] });
    data.push({ range: `${S}!D${r}`, values: [[clientName]] });
    data.push({ range: `${S}!E${r}`, values: [[invDate]] });
    data.push({ range: `${S}!F${r}`, values: [[dueDate]] });
    // Year formula (col G)
    data.push({ range: `${S}!G${r}`, values: [[`=IFERROR(YEAR(E${r}),"")`]] });
    // Month formula (col H)
    data.push({ range: `${S}!H${r}`, values: [[`=IFERROR(TEXT(E${r},"mmmm"),"")`]] });
    data.push({ range: `${S}!I${r}`, values: [[project]] });
    data.push({ range: `${S}!J${r}`, values: [[currency]] });
    data.push({ range: `${S}!K${r}`, values: [[subtotal]] });
    data.push({ range: `${S}!L${r}`, values: [[discount]] });
    data.push({ range: `${S}!M${r}`, values: [[shipping]] });
    data.push({ range: `${S}!N${r}`, values: [[tax]] });
    data.push({ range: `${S}!O${r}`, values: [[total]] });
    data.push({ range: `${S}!P${r}`, values: [[depositReq]] });
    // Payments Received — live SUMIFS from Payment Tracker
    data.push({ range: `${S}!Q${r}`, values: [[`=IFERROR(SUMIFS('Payment Tracker'!$N$6:$N$1505,'Payment Tracker'!$C$6:$C$1505,A${r}),0)`]] });
    // Balance Due
    data.push({ range: `${S}!R${r}`, values: [[`=MAX(0,O${r}-Q${r})`]] });
    // Payment Status auto-logic
    data.push({ range: `${S}!S${r}`, values: [[
      `=IFERROR(IF(T${r}="Cancelled","Cancelled",IF(T${r}="Refunded","Refunded",IF(Q${r}=0,"Unpaid",IF(Q${r}>O${r},"Overpaid",IF(Q${r}=O${r},"Paid",IF(AND(P${r}>0,Q${r}>=P${r},Q${r}<O${r}),"Deposit Paid","Partially Paid")))))),"")`
    ]] });
    data.push({ range: `${S}!T${r}`, values: [[invStatus]] });
    // Days Until / Overdue
    data.push({ range: `${S}!U${r}`, values: [[`=IF(F${r}="","",IF(R${r}=0,0,F${r}-TODAY()))`]] });
    data.push({ range: `${S}!V${r}`, values: [[lastPayDate]] });
    data.push({ range: `${S}!W${r}`, values: [[nextFollowUp]] });
    data.push({ range: `${S}!X${r}`, values: [[followUpStatus]] });
    data.push({ range: `${S}!Y${r}`, values: [[deliveryMethod]] });
    data.push({ range: `${S}!Z${r}`, values: [[sentDate]] });
    data.push({ range: `${S}!AA${r}`, values: [[paidDate]] });
    data.push({ range: `${S}!AB${r}`, values: [[notes]] });
  }

  // Extend formula rows 41-1005
  for (let r = 41; r <= 1005; r++) {
    data.push({ range: `${S}!G${r}`, values: [[`=IFERROR(YEAR(E${r}),"")`]] });
    data.push({ range: `${S}!H${r}`, values: [[`=IFERROR(TEXT(E${r},"mmmm"),"")`]] });
    data.push({ range: `${S}!Q${r}`, values: [[`=IFERROR(SUMIFS('Payment Tracker'!$N$6:$N$1505,'Payment Tracker'!$C$6:$C$1505,A${r}),0)`]] });
    data.push({ range: `${S}!R${r}`, values: [[`=IF(O${r}="","",MAX(0,O${r}-Q${r}))`]] });
    data.push({ range: `${S}!S${r}`, values: [[`=IFERROR(IF(T${r}="Cancelled","Cancelled",IF(T${r}="Refunded","Refunded",IF(Q${r}=0,"Unpaid",IF(Q${r}>O${r},"Overpaid",IF(Q${r}=O${r},"Paid",IF(AND(P${r}>0,Q${r}>=P${r},Q${r}<O${r}),"Deposit Paid","Partially Paid")))))),"")`]] });
    data.push({ range: `${S}!U${r}`, values: [[`=IF(F${r}="","",IF(R${r}=0,0,F${r}-TODAY()))`]] });
  }

  await valuesBatchUpdate(id, data, 'invoicelog-values');

  const reqs = [];

  reqs.push({ repeatCell: { range: gridRange(IL, 0, 1010, 0, 28), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });

  reqs.push({ mergeCells: { range: gridRange(IL, 0, 1, 0, 28), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(IL, 0, 1, 0, 28), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 16 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: IL, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  reqs.push({ mergeCells: { range: gridRange(IL, 1, 2, 0, 28), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(IL, 1, 2, 0, 28), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { italic: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  reqs.push({ repeatCell: { range: gridRange(IL, 4, 5, 0, 28), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  reqs.push({ repeatCell: { range: gridRange(IL, 5, 1005, 0, 28), cell: { userEnteredFormat: { backgroundColor: hex(C.panel) } }, fields: 'userEnteredFormat.backgroundColor' } });
  for (let r = 1; r < 1000; r += 2) {
    reqs.push({ repeatCell: { range: gridRange(IL, 5 + r, 6 + r, 0, 28), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }

  // Formula cols: G,H,Q,R,S,U
  [6,7,16,17,18,20].forEach(ci => {
    reqs.push({ repeatCell: { range: gridRange(IL, 5, 1005, ci, ci + 1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // Currency format: K,L,M,N,O,P,Q,R
  [10,11,12,13,14,15,16,17].forEach(ci => {
    reqs.push({ repeatCell: { range: gridRange(IL, 5, 1005, ci, ci + 1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00;[Red]-$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });
  // Date format: E,F,V,W,Z,AA
  [4,5,21,22,25,26].forEach(ci => {
    reqs.push({ repeatCell: { range: gridRange(IL, 5, 1005, ci, ci + 1), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'MMM D, YYYY' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });

  // Column widths
  const colWidths = [130,100,70,160,90,90,50,80,160,50,90,80,70,70,90,80,90,80,110,110,80,90,90,110,80,90,80,200];
  colWidths.forEach((px, ci) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: IL, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  });

  reqs.push({ updateSheetProperties: { properties: { sheetId: IL, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'invoicelog-format');
  console.log('✓ Invoice Log');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
