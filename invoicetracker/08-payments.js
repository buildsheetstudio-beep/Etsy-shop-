'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const PT = sheetMap['Payment Tracker'];
const S  = "'Payment Tracker'";

const headers = [
  'Payment ID','Payment Date','Invoice ID','Invoice Number','Client ID','Client Name',
  'Payment Type','Gross Amount','Payment Method','Reference / Transaction ID',
  'Deposit?','Refund?','Processing Fee','Net Payment','Balance Before','Balance After','Notes'
];

// [Date, InvoiceID, InvNum, ClientID, ClientName, Type, Gross, Method, Ref, IsDeposit, IsRefund, ProcFee, Notes]
const payments = [
  // Invoice 0001 — Apex — Paid in full
  ['Jan 28, 2026','INV-2026-0001','INV-2026-0001','CL-001','Apex Digital Agency','Final Payment',2120.63,'Bank Transfer','ACH-2026-0001',false,false,0,'Full payment on Net 30'],
  // Invoice 0002 — BlueSky — Paid
  ['Jan 20, 2026','INV-2026-0002','INV-2026-0002','CL-002','BlueSky Architecture','Final Payment',850,'Check','CHK-4821',false,false,0,'Check received by mail'],
  // Invoice 0003 — Driftwood — Paid
  ['Feb 10, 2026','INV-2026-0003','INV-2026-0003','CL-004','Driftwood Publishing','Final Payment',2760.38,'Credit Card','CC-TXN-88821',false,false,27.6,'3% CC processing fee applied'],
  // Invoice 0004 — Emerald — Paid
  ['Jan 28, 2026','INV-2026-0004','INV-2026-0004','CL-005','Emerald Coast Realty','Final Payment',487.92,'PayPal','PP-2026-0041',false,false,14.88,'PayPal fee deducted from net'],
  // Invoice 0005 — Gilded — Paid on delivery
  ['Jan 20, 2026','INV-2026-0005','INV-2026-0005','CL-007','Gilded Thread Boutique','Final Payment',1111.63,'Cash','CASH-0005',false,false,0,'Cash on delivery'],
  // Invoice 0006 — Juniper Lane — Paid
  ['Mar 1, 2026','INV-2026-0006','INV-2026-0006','CL-010','Juniper Lane Interiors','Final Payment',981,'Bank Transfer','ACH-2026-0006',false,false,0,''],
  // Invoice 0007 — Luminary — Paid
  ['Mar 4, 2026','INV-2026-0007','INV-2026-0007','CL-012','Luminary Learning Co.','Final Payment',1254.68,'Bank Transfer','ACH-2026-0007',false,false,0,''],
  // Invoice 0008 — Northgate — Paid
  ['Mar 10, 2026','INV-2026-0008','INV-2026-0008','CL-014','Northgate Consulting','Final Payment',689,'Zelle','ZL-2026-0008',false,false,0,''],
  // Invoice 0009 — Apex Brand Refresh — Deposit + partial
  ['Feb 20, 2026','INV-2026-0009','INV-2026-0009','CL-001','Apex Digital Agency','Deposit',1000,'Bank Transfer','ACH-2026-0009D',true,false,0,'Deposit per contract'],
  ['Mar 5, 2026','INV-2026-0009','INV-2026-0009','CL-001','Apex Digital Agency','Partial Payment',500,'Bank Transfer','ACH-2026-0009P',false,false,0,'Partial — milestone 1 complete'],
  // Invoice 0010 — Harbor View — Deposit
  ['Mar 1, 2026','INV-2026-0010','INV-2026-0010','CL-008','Harbor View Hotels','Deposit',2000,'Bank Transfer','ACH-2026-0010D',true,false,0,'50% deposit at project kick-off'],
  ['Mar 20, 2026','INV-2026-0010','INV-2026-0010','CL-008','Harbor View Hotels','Partial Payment',1500,'Bank Transfer','ACH-2026-0010P',false,false,0,'Phase 1 completion payment'],
  // Invoice 0013 — Opal — Deposit
  ['Mar 12, 2026','INV-2026-0013','INV-2026-0013','CL-015','Opal Events Co.','Deposit',400,'Stripe','STR-0013D',true,false,11.9,'Stripe processing fee 2.9%+$0.30'],
  // Invoice 0021 — Prairie Wind — Paid
  ['Mar 10, 2026','INV-2026-0021','INV-2026-0021','CL-016','Prairie Wind Photography','Final Payment',337.6,'Venmo','VNM-2026-0021',false,false,0,'Monthly editing subscription'],
  // Invoice 0022 — River Oak — Paid same day
  ['Mar 3, 2026','INV-2026-0022','INV-2026-0022','CL-018','River Oak Counseling','Final Payment',431.82,'Credit Card','CC-TXN-22001',false,false,12.97,'CC processing fee 2.9%+$0.30'],
  // Invoice 0024 — Emerald — Paid
  ['Mar 22, 2026','INV-2026-0024','INV-2026-0024','CL-005','Emerald Coast Realty','Final Payment',385.2,'PayPal','PP-2026-0024',false,false,11.56,'PayPal fee applied'],
  // Invoice 0027 — Ironclad — OVERPAYMENT
  ['Mar 24, 2026','INV-2026-0027','INV-2026-0027','CL-009','Ironclad Fitness','Overpayment',500,'ACH','ACH-2026-0027',false,false,0,'Client paid $500 vs $443.52 — credit held'],
  // Invoice 0028 — Harbor View Website — Deposit
  ['Mar 25, 2026','INV-2026-0028','INV-2026-0028','CL-008','Harbor View Hotels','Deposit',1500,'Bank Transfer','ACH-2026-0028D',true,false,0,'Phase 1 deposit'],
  // Invoice 0031 — Flint Steel — Paid
  ['Mar 28, 2026','INV-2026-0031','INV-2026-0031','CL-006','Flint & Steel Brewing Co.','Final Payment',349.86,'Check','CHK-4902',false,false,0,'Check paid on time'],
  // Invoice 0032 — Maple & Birch — Paid on receipt
  ['Mar 10, 2026','INV-2026-0032','INV-2026-0032','CL-013','Maple & Birch Bakery','Final Payment',201.43,'Cash','CASH-0032',false,false,0,'Cash at pickup'],
  // Invoice 0018 — REFUND (Maple & Birch cancelled)
  ['Dec 15, 2025','INV-2026-0018','INV-2026-0018','CL-013','Maple & Birch Bakery','Refund',462.13,'Bank Transfer','REF-2025-0018',false,true,0,'Full refund — project cancelled'],
  // Prior year payments — Invoice 0085 (partially paid at one point, then fell through)
  ['Nov 15, 2025','INV-2025-0085','INV-2025-0085','CL-007','Gilded Thread Boutique','Deposit',300,'Credit Card','CC-TXN-85001',true,false,9,'Deposit — client later requested refund then no-show'],
  // Refund of deposit on 0085
  ['Dec 1, 2025','INV-2025-0085','INV-2025-0085','CL-007','Gilded Thread Boutique','Refund',300,'Credit Card','REF-2025-8501',false,true,0,'Refund of deposit — balance still outstanding'],
  // Invoice 0001 — credit adjustment
  ['Feb 5, 2026','INV-2026-0001','INV-2026-0001','CL-001','Apex Digital Agency','Adjustment',0,'','ADJ-2026-0001',false,false,0,'Courtesy credit — zero net impact'],
  // Additional payments for richer data
  ['Dec 20, 2025','INV-2025-0086','INV-2025-0086','CL-002','BlueSky Architecture','Partial Payment',200,'Check','CHK-4811',false,false,0,'Partial — 61-90 days overdue remainder still owed'],
  ['Feb 28, 2026','INV-2026-0011','INV-2026-0011','CL-003','Cascade Wellness Spa','Partial Payment',100,'Venmo','VNM-0011A',false,false,0,'Partial payment on newsletter invoice'],
  ['Mar 1, 2026','INV-2026-0003','INV-2026-0003','CL-004','Driftwood Publishing','Adjustment',-27.6,'','ADJ-0003-FEE',false,false,0,'Processing fee adjustment — client reimbursed fee'],
  // Harbor Hotels second partial on 0010
  ['Mar 22, 2026','INV-2026-0010','INV-2026-0010','CL-008','Harbor View Hotels','Partial Payment',0,'','ON-HOLD-0010',false,false,0,'On hold — awaiting deliverable approval'],
  // Ironclad Q1 original
  ['Feb 25, 2026','INV-2026-0014','INV-2026-0014','CL-009','Ironclad Fitness','Partial Payment',200,'Stripe','STR-0014A',false,false,5.9,'Partial on social pack'],
  // Large payment — Apex second
  ['Mar 18, 2026','INV-2026-0009','INV-2026-0009','CL-001','Apex Digital Agency','Partial Payment',0,'','PENDING-0009',false,false,0,'Milestone 2 pending approval'],
  // Cash payment
  ['Mar 15, 2026','INV-2026-0022','INV-2026-0022','CL-018','River Oak Counseling','Adjustment',0,'','ADJ-0022',false,false,0,'Minor rounding adjustment'],
  // Zelle
  ['Mar 8, 2026','INV-2026-0012','INV-2026-0012','CL-006','Flint & Steel Brewing Co.','Partial Payment',100,'Zelle','ZL-0012A',false,false,0,'Partial — client paying in installments'],
  // Opal second
  ['Mar 20, 2026','INV-2026-0013','INV-2026-0013','CL-015','Opal Events Co.','Partial Payment',400,'Stripe','STR-0013P',false,false,11.9,'Partial payment at milestone'],
  // Final large payment
  ['Mar 28, 2026','INV-2026-0006','INV-2026-0006','CL-010','Juniper Lane Interiors','Adjustment',0,'','ADJ-0006',false,false,0,'Account reconciliation'],
  // ACH payment Northgate Q2
  ['Mar 14, 2026','INV-2026-0029','INV-2026-0029','CL-014','Northgate Consulting','Partial Payment',200,'ACH','ACH-2026-0029P',false,false,0,'Partial on overdue invoice'],
  // Deposit on invoice 0033
  ['Mar 20, 2026','INV-2026-0033','INV-2026-0033','CL-015','Opal Events Co.','Deposit',350,'Stripe','STR-0033D',true,false,10.45,'Deposit on invitation suite'],
  // More adjustments and edge cases
  ['Jan 5, 2026','INV-2026-0004','INV-2026-0004','CL-005','Emerald Coast Realty','Credit',0,'','CREDIT-0004',false,false,0,'Courtesy credit for rescheduling'],
  ['Feb 15, 2026','INV-2026-0007','INV-2026-0007','CL-012','Luminary Learning Co.','Partial Payment',654.68,'Bank Transfer','ACH-2026-0007A',false,false,0,'First milestone payment'],
  ['Mar 4, 2026','INV-2026-0007','INV-2026-0007','CL-012','Luminary Learning Co.','Final Payment',600,'Bank Transfer','ACH-2026-0007B',false,false,0,'Final milestone payment'],
  // Final tallies
  ['Feb 12, 2026','INV-2026-0006','INV-2026-0006','CL-010','Juniper Lane Interiors','Deposit',400,'Bank Transfer','ACH-2026-0006D',true,false,0,'Deposit at project start'],
  ['Mar 1, 2026','INV-2026-0006','INV-2026-0006','CL-010','Juniper Lane Interiors','Final Payment',581,'Bank Transfer','ACH-2026-0006F',false,false,0,'Balance payment'],
  ['Mar 5, 2026','INV-2026-0025','INV-2026-0025','CL-001','Apex Digital Agency','Deposit',1000,'Bank Transfer','ACH-2026-0025D',true,false,0,'Retainer deposit'],
  ['Mar 24, 2026','INV-2026-0026','INV-2026-0026','CL-010','Juniper Lane Interiors','Deposit',500,'Bank Transfer','ACH-2026-0026D',true,false,0,'Staging deposit'],
  ['Mar 26, 2026','INV-2026-0023','INV-2026-0023','CL-004','Driftwood Publishing','Partial Payment',400,'Check','CHK-4924',false,false,0,'Partial payment on spring catalog'],
];

(async () => {
  // Extend Payment Tracker to 1510 rows (default grid is 1200)
  await batchUpdate(id, [{ appendDimension: { sheetId: PT, dimension: 'ROWS', length: 310 } }], 'extend-pt-rows');

  const data = [];

  data.push({ range: `${S}!A1`, values: [['PAYMENT TRACKER']] });
  data.push({ range: `${S}!A2`, values: [['Log every payment received here. Payment ID auto-generates. Net Payment updates Invoice Log automatically via SUMIFS. Use Refund? checkbox to enter refunds as negative values.']] });
  data.push({ range: `${S}!A5`, values: [headers] });

  for (let i = 0; i < payments.length; i++) {
    const r = i + 6;
    data.push({ range: `${S}!A${r}`, values: [[`=IF(B${r}="","","PAY-"&TEXT(ROW()-5,"0000"))`]] });
    const [date, invId, invNum, clientId, clientName, type, gross, method, ref, isDeposit, isRefund, procFee, notes] = payments[i];
    data.push({ range: `${S}!B${r}`, values: [[date]] });
    data.push({ range: `${S}!C${r}`, values: [[invId]] });
    data.push({ range: `${S}!D${r}`, values: [[invNum]] });
    data.push({ range: `${S}!E${r}`, values: [[clientId]] });
    data.push({ range: `${S}!F${r}`, values: [[clientName]] });
    data.push({ range: `${S}!G${r}`, values: [[type]] });
    data.push({ range: `${S}!H${r}`, values: [[gross]] });
    data.push({ range: `${S}!I${r}`, values: [[method]] });
    data.push({ range: `${S}!J${r}`, values: [[ref]] });
    data.push({ range: `${S}!K${r}`, values: [[isDeposit]] });
    data.push({ range: `${S}!L${r}`, values: [[isRefund]] });
    data.push({ range: `${S}!M${r}`, values: [[procFee]] });
    // Net Payment formula
    data.push({ range: `${S}!N${r}`, values: [[`=IF(L${r}=TRUE,-ABS(H${r})-M${r},H${r}-M${r})`]] });
    // Balance Before = Invoice Total - prior payments (simplified reference)
    data.push({ range: `${S}!O${r}`, values: [[`=IFERROR(VLOOKUP(C${r},'Invoice Log'!$A$6:$O$1005,15,0)-SUMIFS($N$6:N${r-1},$C$6:C${r-1},C${r}),"")`]] });
    // Balance After
    data.push({ range: `${S}!P${r}`, values: [[`=IFERROR(O${r}-N${r},"")`]] });
    data.push({ range: `${S}!Q${r}`, values: [[notes]] });
  }
  // Extend formula rows
  for (let r = 51; r <= 1505; r++) {
    data.push({ range: `${S}!A${r}`, values: [[`=IF(B${r}="","","PAY-"&TEXT(ROW()-5,"0000"))`]] });
    data.push({ range: `${S}!N${r}`, values: [[`=IF(L${r}=TRUE,-ABS(H${r})-M${r},H${r}-M${r})`]] });
    data.push({ range: `${S}!O${r}`, values: [[`=IFERROR(VLOOKUP(C${r},'Invoice Log'!$A$6:$O$1005,15,0)-SUMIFS($N$6:N${r-1},$C$6:C${r-1},C${r}),"")`]] });
    data.push({ range: `${S}!P${r}`, values: [[`=IFERROR(O${r}-N${r},"")`]] });
  }

  await valuesBatchUpdate(id, data, 'payments-values');

  const reqs = [];

  reqs.push({ repeatCell: { range: gridRange(PT, 0, 1510, 0, 18), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });

  reqs.push({ mergeCells: { range: gridRange(PT, 0, 1, 0, 17), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(PT, 0, 1, 0, 17), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 16 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: PT, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  reqs.push({ mergeCells: { range: gridRange(PT, 1, 2, 0, 17), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(PT, 1, 2, 0, 17), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { italic: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  reqs.push({ repeatCell: { range: gridRange(PT, 4, 5, 0, 17), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  reqs.push({ repeatCell: { range: gridRange(PT, 5, 1505, 0, 17), cell: { userEnteredFormat: { backgroundColor: hex(C.panel) } }, fields: 'userEnteredFormat.backgroundColor' } });
  for (let r = 1; r < 1500; r += 2) {
    reqs.push({ repeatCell: { range: gridRange(PT, 5 + r, 6 + r, 0, 17), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }

  // Formula cols: A,N,O,P
  [0,13,14,15].forEach(ci => {
    reqs.push({ repeatCell: { range: gridRange(PT, 5, 1505, ci, ci + 1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // Currency: H,M,N,O,P
  [7,12,13,14,15].forEach(ci => {
    reqs.push({ repeatCell: { range: gridRange(PT, 5, 1505, ci, ci + 1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00;[Red]-$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });
  // Date: B
  reqs.push({ repeatCell: { range: gridRange(PT, 5, 1505, 1, 2), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'MMM D, YYYY' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Column widths
  [90,90,130,100,70,160,110,90,110,140,70,60,90,90,100,100,200].forEach((px, ci) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: PT, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  });

  reqs.push({ updateSheetProperties: { properties: { sheetId: PT, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'payments-format');
  console.log('✓ Payment Tracker');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
