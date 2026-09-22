'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Contribution Log'];
const S   = "'Contribution Log'";
const NC  = 14;

// ── Transaction generator ────────────────────────────────────────────────────
const TXNS = [];
let txNum = 1;

const fmtDate = (d) => `${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}/${d.getFullYear()}`;

const addTx = (dateObj, ben, acc, contrib, type, src, amt, freq, notes='') => {
  TXNS.push([
    `TXN-${String(txNum++).padStart(5,'0')}`,
    fmtDate(dateObj), ben, acc, contrib, type, src, amt, freq, dateObj.getFullYear(), notes,
    '', '', ''  // L, M, N — formula columns set below
  ]);
};

const monthly = (ben, acc, contrib, src, amt, y0, m0, y1, m1, notes='') => {
  let y=y0, m=m0;
  while (y<y1 || (y===y1 && m<=m1)) {
    addTx(new Date(y,m-1,1), ben, acc, contrib, 'Contribution', src, amt, 'Monthly', notes);
    if (++m>12) { m=1; y++; }
  }
};
const quarterly = (ben, acc, contrib, src, amt, y0, m0, y1, m1) => {
  let y=y0, m=m0;
  while (y<y1 || (y===y1 && m<=m1)) {
    addTx(new Date(y,m-1,1), ben, acc, contrib, 'Contribution', src, amt, 'Quarterly');
    m+=3; if (m>12) { m-=12; y++; }
  }
};
const single = (dateStr, ben, acc, contrib, type, src, amt, freq='Irregular / Manual', notes='') =>
  addTx(new Date(dateStr), ben, acc, contrib, type, src, amt, freq, notes);

// ── ACC-001 (BEN-001): ScholarShare — Rachel Hartley ────────────────────────
single('2024-01-05','BEN-001','ACC-001','Rachel Hartley','Contribution','Parent / Guardian',5000,'Irregular / Manual','Initial deposit');
monthly('BEN-001','ACC-001','Rachel Hartley','Payroll / Automatic Deposit',300, 2024,1, 2026,9);

// ── ACC-002 (BEN-002): Vanguard — David Hartley ─────────────────────────────
single('2024-01-05','BEN-002','ACC-002','David Hartley','Contribution','Parent / Guardian',8000,'Irregular / Manual','Initial deposit');
monthly('BEN-002','ACC-002','David Hartley','Payroll / Automatic Deposit',400, 2024,1, 2026,9);

// ── ACC-003 (BEN-003): Fidelity UNIQUE — Carlos Delgado ─────────────────────
single('2024-01-10','BEN-003','ACC-003','Carlos Delgado','Contribution','Parent / Guardian',12000,'Irregular / Manual','Initial deposit');
monthly('BEN-003','ACC-003','Carlos Delgado','Payroll / Automatic Deposit',350, 2024,1, 2026,8);

// ── ACC-004 (BEN-004): CollegeAdvantage — enrolled; deposits then withdrawals
single('2024-01-15','BEN-004','ACC-004','James Washington','Contribution','Parent / Guardian',45000,'Irregular / Manual','Balance forward — prior years savings');
monthly('BEN-004','ACC-004','James Washington','Payroll / Automatic Deposit',350, 2024,1, 2024,9);
for (let m=1;m<=9;m++) single(`2025-${String(m).padStart(2,'0')}-15`,'BEN-004','ACC-004','Marcus Washington','Withdrawal','Beneficiary',2000,'Monthly','Tuition and living — enrolled');
for (let m=1;m<=9;m++) single(`2026-${String(m).padStart(2,'0')}-15`,'BEN-004','ACC-004','Marcus Washington','Withdrawal','Beneficiary',2000,'Monthly','Tuition and living — enrolled');

// ── ACC-005 (BEN-005): my529 — Thomas + Rose Beaumont ───────────────────────
single('2024-01-20','BEN-005','ACC-005','Thomas Beaumont','Contribution','Parent / Guardian',15000,'Irregular / Manual','Initial deposit');
monthly('BEN-005','ACC-005','Thomas Beaumont','Payroll / Automatic Deposit',350, 2024,1, 2026,9);
single('2024-12-20','BEN-005','ACC-005','Rose Beaumont','Contribution','Gift',5000,'Annual','Annual holiday gift contribution');
single('2025-12-20','BEN-005','ACC-005','Rose Beaumont','Contribution','Gift',5000,'Annual','Annual holiday gift contribution');

// ── ACC-006 (BEN-001): ScholarShare grandparent — Margaret Chen ─────────────
single('2024-01-05','BEN-001','ACC-006','Margaret Chen','Contribution','Grandparent',3000,'Irregular / Manual','Initial deposit');
quarterly('BEN-001','ACC-006','Margaret Chen','Grandparent',500, 2024,1, 2026,7);
// Birthday gift each year
single('2024-03-15','BEN-001','ACC-006','Margaret Chen','Contribution','Gift',500,'Irregular / Manual','Birthday gift — Emma');
single('2025-03-15','BEN-001','ACC-006','Margaret Chen','Contribution','Gift',500,'Irregular / Manual','Birthday gift — Emma');
single('2026-03-15','BEN-001','ACC-006','Margaret Chen','Contribution','Gift',500,'Irregular / Manual','Birthday gift — Emma');

// ── ACC-007 (BEN-002): College Savings Iowa — grandparent ───────────────────
single('2024-01-10','BEN-002','ACC-007','Rose Beaumont','Contribution','Grandparent',4000,'Irregular / Manual','Initial deposit');
quarterly('BEN-002','ACC-007','Rose Beaumont','Grandparent',500, 2024,1, 2026,7);
// Tax refund contribution
single('2024-04-10','BEN-002','ACC-007','David Hartley','Contribution','Tax Refund',1200,'Irregular / Manual','Federal tax refund directed to 529');
single('2025-04-08','BEN-002','ACC-007','David Hartley','Contribution','Tax Refund',1100,'Irregular / Manual','Federal tax refund directed to 529');

// ── ACC-008 (BEN-003): DreamAhead — opened, then transferred/closed ─────────
single('2023-01-15','BEN-003','ACC-008','Carlos Delgado','Contribution','Parent / Guardian',2000,'Irregular / Manual','Initial deposit');
monthly('BEN-003','ACC-008','Carlos Delgado','Payroll / Automatic Deposit',350, 2023,2, 2023,9,'');
single('2023-10-01','BEN-003','ACC-008','Carlos Delgado','Transfer Out','Parent / Guardian',4800,'Irregular / Manual','Full balance transferred to ACC-003 (Fidelity UNIQUE 529)');

// ── ACC-009 (BEN-004): MOST 529 — second account; balance forward + withdrawals
single('2024-01-15','BEN-004','ACC-009','Patricia Washington','Contribution','Parent / Guardian',50000,'Irregular / Manual','Balance forward — prior years savings');
monthly('BEN-004','ACC-009','Patricia Washington','Payroll / Automatic Deposit',350, 2024,1, 2024,9);
for (let m=10;m<=12;m++) single(`2024-${m}-15`,'BEN-004','ACC-009','Marcus Washington','Withdrawal','Beneficiary',2000,'Monthly','Room and board — enrolled');
for (let m=1;m<=9;m++) single(`2025-${String(m).padStart(2,'0')}-15`,'BEN-004','ACC-009','Marcus Washington','Withdrawal','Beneficiary',2000,'Monthly','Room and board — enrolled');
for (let m=1;m<=9;m++) single(`2026-${String(m).padStart(2,'0')}-15`,'BEN-004','ACC-009','Marcus Washington','Withdrawal','Beneficiary',2000,'Monthly','Room and board — enrolled');

// ── ACC-010 (BEN-005): NY 529 — Thomas Beaumont starting Jun 2024 ────────────
single('2024-06-01','BEN-005','ACC-010','Thomas Beaumont','Contribution','Parent / Guardian',10000,'Irregular / Manual','Initial deposit');
monthly('BEN-005','ACC-010','Thomas Beaumont','Payroll / Automatic Deposit',350, 2024,6, 2026,9);

// Bonus contributions
single('2025-01-10','BEN-001','ACC-001','Rachel Hartley','Contribution','Bonus',1500,'Irregular / Manual','Year-end bonus directed to 529');
single('2025-01-10','BEN-002','ACC-002','David Hartley','Contribution','Bonus',1500,'Irregular / Manual','Year-end bonus directed to 529');
single('2026-01-08','BEN-005','ACC-005','Thomas Beaumont','Contribution','Bonus',2000,'Irregular / Manual','Year-end bonus contribution');

// ── Build formula rows for L, M, N ──────────────────────────────────────────
const numTxns = TXNS.length;

// Replace the blank L, M, N placeholders with actual formulas
TXNS.forEach((row, i) => {
  const r = i + 6; // 1-indexed row in sheet (header is row 5, data starts row 6)
  row[11] = `=IFERROR(VLOOKUP(D${r},'529 Accounts'!$A$8:$C$1007,3,FALSE),"")`;
  row[12] = `=IFERROR(VLOOKUP(C${r},'Beneficiary Setup'!$A$8:$B$507,2,FALSE),"")`;
  row[13] = `=IF(B${r}="","",IF(YEAR(B${r})=YEAR(TODAY()),"Yes","No"))`;
});

(async () => {
  const vals = [];
  const fmt  = [];

  fmt.push({ repeatCell:{ range:gridRange(SID,0,5100,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.bg), textFormat:{ fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.text) }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title
  vals.push({ range:`${S}!A1`, values:[['Contribution Log']] });
  fmt.push({ mergeCells:{ range:gridRange(SID,0,1,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,0,1,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.primary), textFormat:{ bold:true, fontSize:16, foregroundColor:hex(C.white) },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:42 }, fields:'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range:`${S}!A2`, values:[["Log all contributions, withdrawals, and transfers. Account and beneficiary names auto-fill. Balances in 529 Accounts tab pull from this log."]] });
  fmt.push({ mergeCells:{ range:gridRange(SID,1,2,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,1,2,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.aubergTint), textFormat:{ fontSize:9, foregroundColor:hex(C.text), italic:true },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});

  // Rows 3-4: Summary Cards
  const CARDS = [
    { label:'Total Transactions', val:`=COUNTA(A6:A5005)`, cur:false },
    { label:'Total Contributions (All-Time)', val:`=IFERROR(SUMIF(F6:F5005,"Contribution",H6:H5005),"")`, cur:true },
    { label:'Total Withdrawals (All-Time)', val:`=IFERROR(SUMIF(F6:F5005,"Withdrawal",H6:H5005),"")`, cur:true },
    { label:'YTD Net Contributions', val:`=IFERROR(SUMPRODUCT((YEAR(B6:B5005)=YEAR(TODAY()))*(F6:F5005="Contribution")*H6:H5005)-SUMPRODUCT((YEAR(B6:B5005)=YEAR(TODAY()))*(F6:F5005="Withdrawal")*H6:H5005),"")`, cur:true },
  ];
  const cSpans = [[0,3],[3,7],[7,10],[10,14]];
  CARDS.forEach(({ label, val, cur }, i) => {
    const [cs,ce] = cSpans[i];
    const cl = String.fromCharCode(65+cs);
    vals.push({ range:`${S}!${cl}3`, values:[[label]] });
    vals.push({ range:`${S}!${cl}4`, values:[[val]] });
    fmt.push({ mergeCells:{ range:gridRange(SID,2,3,cs,ce), mergeType:'MERGE_ALL' }});
    fmt.push({ repeatCell:{ range:gridRange(SID,2,3,cs,ce), cell:{ userEnteredFormat:{
      backgroundColor:hex(C.eucalTint), textFormat:{ bold:true, fontSize:9, foregroundColor:hex(C.secText) },
      horizontalAlignment:'CENTER'
    }}, fields:'userEnteredFormat' }});
    fmt.push({ mergeCells:{ range:gridRange(SID,3,4,cs,ce), mergeType:'MERGE_ALL' }});
    const vf = { backgroundColor:hex(C.white), textFormat:{ bold:true, fontSize:20, foregroundColor:hex(C.primary) },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' };
    if (cur) vf.numberFormat = { type:'CURRENCY', pattern:'"$"#,##0' };
    fmt.push({ repeatCell:{ range:gridRange(SID,3,4,cs,ce), cell:{ userEnteredFormat:vf }, fields:'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:2, endIndex:5 }, properties:{ pixelSize:36 }, fields:'pixelSize' }});

  // Row 5: Column headers
  vals.push({ range:`${S}!A5`, values:[[
    'Transaction ID','Date','Beneficiary ID','Account ID','Contributor',
    'Transaction Type','Source','Amount','Frequency','Tax Year','Notes',
    'Account Name','Beneficiary Name','YTD?'
  ]] });
  fmt.push({ repeatCell:{ range:gridRange(SID,4,5,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.hdrDark), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9 },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:4, endIndex:5 }, properties:{ pixelSize:36 }, fields:'pixelSize' }});

  // Write all transaction data
  vals.push({ range:`${S}!A6`, values:TXNS });

  // Alternate row colors
  for (let i = 0; i < numTxns; i++) {
    fmt.push({ repeatCell:{ range:gridRange(SID,5+i,6+i,0,NC), cell:{ userEnteredFormat:{
      backgroundColor:hex(i%2===0 ? C.white : C.altRow)
    }}, fields:'userEnteredFormat.backgroundColor' }});
  }

  // Input columns A-K (0-10)
  fmt.push({ repeatCell:{ range:gridRange(SID,5,5+numTxns,0,11), cell:{ userEnteredFormat:{ backgroundColor:hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' }});
  // Amount (H=7): currency
  fmt.push({ repeatCell:{ range:gridRange(SID,5,5005,7,8), cell:{ userEnteredFormat:{ numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0.00' } } }, fields:'userEnteredFormat.numberFormat' }});
  // Date (B=1): date format
  fmt.push({ repeatCell:{ range:gridRange(SID,5,5005,1,2), cell:{ userEnteredFormat:{ numberFormat:{ type:'DATE', pattern:'MM/DD/YYYY' } } }, fields:'userEnteredFormat.numberFormat' }});
  // Formula columns L-N (11-13)
  fmt.push({ repeatCell:{ range:gridRange(SID,5,5+numTxns,11,14), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Freeze row 5
  fmt.push({ updateSheetProperties:{ properties:{ sheetId:SID, gridProperties:{ frozenRowCount:5 } }, fields:'gridProperties.frozenRowCount' }});

  // Column widths
  [90,100,90,80,150, 130,160,100,110,70,180, 160,140,60]
    .forEach((w,i) => fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'COLUMNS', startIndex:i, endIndex:i+1 }, properties:{ pixelSize:w }, fields:'pixelSize' }}));

  await valuesBatchUpdate(id, vals, '06-contribution-log');
  await batchUpdate(id, fmt, '06-contribution-log');
  console.log(`06-contribution-log done ✓  (${numTxns} transactions)`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
