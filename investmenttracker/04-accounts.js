'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Account Tracker'];
const S = "'Account Tracker'";
const REF = "'Reference Data'";

const HDR_ROW = 4;  // 0-indexed row 5 in sheet (frozen through row 5)
const DATA_START = 5; // row 6 in sheet

const HEADERS = [
  'Account ID','Account Name','Owner','Institution','Account Type',
  'Tax Treatment','Acct # Mask','Opening Balance','Current Cash','Holdings Value',
  'Total Value','Annual Fee %','Employer Match?','Beneficiary Reviewed?','Active?',
  'In Dashboard?','Last Updated','Notes'
];

const COL_W = [90,185,120,130,140,140,110,110,110,110,110,85,95,110,70,85,105,200];

// sheetId 2 accounts for Holdings Value formula
// Current Holdings Value = SUMPRODUCT on Holdings tab
const HV = (act) => `=IFERROR(SUMPRODUCT(('Holdings'!$C$6:$C$1005="${act}")*('Holdings'!$O$6:$O$1005)),0)`;

const ACCOUNTS = [
  // [Name, Owner, Institution, AcctType, TaxTreatment, AcctMask, OpenBal, CurCash, EmpMatch, BenRev, Active, InDash, LastUpd, Notes]
  ['Walsh Family Taxable','Daniel Walsh','Fidelity','Taxable Brokerage','Taxable','XXXX-4821',85000,2340.50,false,true,true,true,'2/15/2026','Core taxable brokerage account'],
  ["Daniel's Roth IRA",'Daniel Walsh','Vanguard','Roth IRA','Tax-Free / Roth','XXXX-2293',42000,125.00,false,true,true,true,'2/15/2026','Max contribution annually'],
  ["Daniel's 401(k)",'Daniel Walsh','Vanguard','401(k)','Tax-Deferred','XXXX-8847',115000,0,true,true,true,true,'1/31/2026','Employer 4% match'],
  ["Emily's Taxable Brokerage",'Emily Walsh','Schwab','Taxable Brokerage','Taxable','XXXX-6612',38000,860.00,false,true,true,true,'2/10/2026',''],
  ["Emily's Roth IRA",'Emily Walsh','Fidelity','Roth IRA','Tax-Free / Roth','XXXX-3358',55000,95.00,false,true,true,true,'2/10/2026','Max contribution annually'],
  ["Emily's 403(b)",'Emily Walsh','TIAA','403(b)','Tax-Deferred','XXXX-7741',98000,0,true,false,true,true,'1/31/2026','University employer plan'],
  ['Walsh Joint Brokerage','Joint Household','Schwab','Taxable Brokerage','Taxable','XXXX-5503',62000,1580.25,false,true,true,true,'2/10/2026','Joint account for shared goals'],
  ["Daniel's HSA",'Daniel Walsh','HSA Bank','HSA','Tax-Free / Roth','XXXX-0034',12500,800.00,false,true,true,true,'2/1/2026','Invested portion in VTI'],
  ["Emily's HSA",'Emily Walsh','Optum','HSA','Tax-Free / Roth','XXXX-1182',8800,650.00,false,true,true,true,'2/1/2026',''],
  ["College Savings 529",'Joint Household','Fidelity','Education Account','Tax-Free / Roth','XXXX-9920',18500,200.00,false,false,true,true,'1/15/2026','For child education'],
  ['Walsh Family Trust','Joint Household','Northern Trust','Trust Account','Taxable','XXXX-3370',210000,4200.00,false,true,true,true,'2/1/2026','Revocable living trust'],
  ["Daniel's Crypto",'Daniel Walsh','Coinbase','Crypto Account','Taxable','N/A',8500,0,false,false,true,true,'2/14/2026','BTC and ETH only'],
  ["Former SEP IRA (Rolled)",'Daniel Walsh','TD Ameritrade','SEP IRA','Tax-Deferred','XXXX-5599',0,0,false,false,false,false,'6/30/2023','Rolled over to 401k in 2023'],
];

(async () => {
  const vals = [];
  const fmt  = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,400,0,18), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // Column widths
  COL_W.forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Title banner (rows 1-2)
  fmt.push({ mergeCells: { range: gridRange(SID,0,2,0,18), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${S}!A1`, values: [['ACCOUNT TRACKER\nAll investment accounts • balances • ownership • status']] });
  fmt.push({ repeatCell: { range: gridRange(SID,0,2,0,18), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 46 }, fields: 'pixelSize' }});

  // Summary Cards (rows 3-4, 0-indexed 2-3)
  const totalVal   = `=IFERROR(SUMPRODUCT(($O$6:$O$305=TRUE)*($K$6:$K$305)),0)`;
  const taxableVal = `=IFERROR(SUMPRODUCT(($O$6:$O$305=TRUE)*($F$6:$F$305="Taxable")*($K$6:$K$305)),0)`;
  const deferVal   = `=IFERROR(SUMPRODUCT(($O$6:$O$305=TRUE)*($F$6:$F$305="Tax-Deferred")*($K$6:$K$305)),0)`;
  const rothVal    = `=IFERROR(SUMPRODUCT(($O$6:$O$305=TRUE)*($F$6:$F$305="Tax-Free / Roth")*($K$6:$K$305)),0)`;
  const cashBal    = `=IFERROR(SUMPRODUCT(($O$6:$O$305=TRUE)*($I$6:$I$305)),0)`;
  const activeAct  = `=IFERROR(COUNTIF($O$6:$O$305,TRUE),0)`;
  const avgFee     = `=IFERROR(AVERAGEIF($O$6:$O$305,TRUE,$L$6:$L$305),0)`;
  const noReview   = `=IFERROR(COUNTIFS($O$6:$O$305,TRUE,$N$6:$N$305,FALSE),0)`;

  const summCards = [
    ['Total Portfolio Value', totalVal, '$#,##0.00'],
    ['Taxable Value',         taxableVal,'$#,##0.00'],
    ['Tax-Deferred Value',    deferVal,  '$#,##0.00'],
    ['Tax-Free / Roth Value', rothVal,   '$#,##0.00'],
    ['Total Cash Balance',    cashBal,   '$#,##0.00'],
    ['Active Accounts',       activeAct, '0'],
    ['Avg Annual Fee %',      avgFee,    '0.00%'],
    ['Beneficiary Review Needed', noReview, '0'],
  ];

  summCards.forEach((card, i) => {
    const c = i * 2;
    if (c >= 16) return; // don't overflow
    fmt.push({ mergeCells: { range: gridRange(SID,2,3,c,c+2), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID,3,4,c,c+2), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}3`, values: [[card[0]]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}4`, values: [[card[1]]] });
    fmt.push({ repeatCell: { range: gridRange(SID,2,3,c,c+2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrB), textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,3,4,c,c+2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.highlight), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      numberFormat: { type: card[2].includes('%') ? 'PERCENT' : card[2].includes('$') ? 'CURRENCY' : 'NUMBER', pattern: card[2] },
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 4 },
      properties: { pixelSize: 30 }, fields: 'pixelSize' }});
  });

  // Column headers row (HDR_ROW = 4)
  vals.push({ range: `${S}!A${HDR_ROW+1}`, values: [HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID,HDR_ROW,HDR_ROW+1,0,18), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: HDR_ROW, endIndex: HDR_ROW+1 },
    properties: { pixelSize: 32 }, fields: 'pixelSize' }});

  // Data rows
  for (let i = 0; i < ACCOUNTS.length; i++) {
    const r = DATA_START + i;
    const [name, owner, inst, accType, taxTreat, mask, openBal, curCash, empMatch, benRev, active, inDash, lastUpd, notes] = ACCOUNTS[i];
    const actId = `ACT-${String(i+1).padStart(3,'0')}`;
    const bg = i % 2 === 0 ? C.panel : C.altRow;

    // Col A: Account ID formula
    vals.push({ range: `${S}!A${r+1}`, values: [[`=IF(B${r+1}="","","ACT-"&TEXT(ROW()-5,"000"))`]] });
    // Col B-Q: values
    vals.push({ range: `${S}!B${r+1}`, values: [[name]] });
    vals.push({ range: `${S}!C${r+1}`, values: [[owner]] });
    vals.push({ range: `${S}!D${r+1}`, values: [[inst]] });
    vals.push({ range: `${S}!E${r+1}`, values: [[accType]] });
    vals.push({ range: `${S}!F${r+1}`, values: [[taxTreat]] });
    vals.push({ range: `${S}!G${r+1}`, values: [[mask]] });
    vals.push({ range: `${S}!H${r+1}`, values: [[openBal]] });
    vals.push({ range: `${S}!I${r+1}`, values: [[curCash]] });
    // Col J: Holdings Value = SUMPRODUCT from Holdings tab (reference actId)
    vals.push({ range: `${S}!J${r+1}`, values: [[`=IFERROR(SUMPRODUCT(('Holdings'!$C$6:$C$1005=A${r+1})*('Holdings'!$O$6:$O$1005)),0)`]] });
    // Col K: Total Account Value
    vals.push({ range: `${S}!K${r+1}`, values: [[`=IFERROR(I${r+1}+J${r+1},0)`]] });
    vals.push({ range: `${S}!L${r+1}`, values: [[0.0015]] }); // Annual fee % default
    vals.push({ range: `${S}!M${r+1}`, values: [[empMatch]] });
    vals.push({ range: `${S}!N${r+1}`, values: [[benRev]] });
    vals.push({ range: `${S}!O${r+1}`, values: [[active]] });
    vals.push({ range: `${S}!P${r+1}`, values: [[inDash]] });
    vals.push({ range: `${S}!Q${r+1}`, values: [[lastUpd]] });
    vals.push({ range: `${S}!R${r+1}`, values: [[notes]] });

    // Row format
    fmt.push({ repeatCell: { range: gridRange(SID,r,r+1,0,18), cell: { userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { fontSize: 9, fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 20 }, fields: 'pixelSize' }});
  }

  // Override specific fees
  const fees = [0.0003,0.0003,0.0008,0.0003,0.0003,0.0010,0.0003,0.0005,0.0005,0.0008,0.0012,0.0025,0.0012];
  fees.forEach((fee, i) => {
    vals.push({ range: `${S}!L${DATA_START+i+1}`, values: [[fee]] });
  });

  // Column format
  // Currency columns: H, I, J, K (7,8,9,10)
  [7,8,9,10].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID,DATA_START,DATA_START+ACCOUNTS.length,c,c+1), cell: { userEnteredFormat: {
      numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' }
    }}, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Fee % column (L=11)
  fmt.push({ repeatCell: { range: gridRange(SID,DATA_START,DATA_START+ACCOUNTS.length,11,12), cell: { userEnteredFormat: {
    numberFormat: { type: 'PERCENT', pattern: '0.00%' }
  }}, fields: 'userEnteredFormat.numberFormat' }});
  // Date column Q (16)
  fmt.push({ repeatCell: { range: gridRange(SID,DATA_START,DATA_START+ACCOUNTS.length,16,17), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'MMM D, YYYY' }
  }}, fields: 'userEnteredFormat.numberFormat' }});
  // Formula tint for J (holdings value from formula)
  fmt.push({ repeatCell: { range: gridRange(SID,DATA_START,DATA_START+ACCOUNTS.length,9,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula),
  }}, fields: 'userEnteredFormat.backgroundColor' }});
  fmt.push({ repeatCell: { range: gridRange(SID,DATA_START,DATA_START+ACCOUNTS.length,10,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // Dropdowns
  fmt.push({ setDataValidation: { range: gridRange(SID,DATA_START,DATA_START+50,2,3),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$4:$A$6` }] }, showCustomUi: true, strict: false } }});
  fmt.push({ setDataValidation: { range: gridRange(SID,DATA_START,DATA_START+50,4,5),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$9:$A$23` }] }, showCustomUi: true, strict: false } }});
  fmt.push({ setDataValidation: { range: gridRange(SID,DATA_START,DATA_START+50,5,6),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$26:$A$30` }] }, showCustomUi: true, strict: false } }});

  // Checkboxes: M(12), N(13), O(14), P(15)
  [12,13,14,15].forEach(c => {
    fmt.push({ setDataValidation: { range: gridRange(SID,DATA_START,DATA_START+50,c,c+1),
      rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true } }});
  });

  // Conditional formatting
  // Active = TRUE → sage
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID,DATA_START,DATA_START+50,0,18)],
    booleanRule: { condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=AND($O6=TRUE,$O6<>"")` }] },
      format: { /* no override — default colors */ }
    }
  }, index: 0 }});
  // Inactive → gray
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID,DATA_START,DATA_START+50,0,18)],
    booleanRule: { condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$O6=FALSE` }] },
      format: { textFormat: { foregroundColor: hex(C.secText) }, backgroundColor: hex(C.altRow) }
    }
  }, index: 1 }});
  // Beneficiary not reviewed → amber
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID,DATA_START,DATA_START+50,13,14)],
    booleanRule: { condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=AND($N6=FALSE,$O6=TRUE)` }] },
      format: { backgroundColor: hex(C.warning) }
    }
  }, index: 2 }});

  // Freeze rows 1:5 and columns A:C
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 5, frozenColumnCount: 3 } }, fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount' }});

  await valuesBatchUpdate(id, vals, '04-accounts values');
  await batchUpdate(id, fmt, '04-accounts format');
  console.log('✅ Account Tracker done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
