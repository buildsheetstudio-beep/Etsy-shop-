'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Price Updates'];
const S = "'Price Updates'";
const REF = "'Reference Data'";

const HEADERS = [
  'Security ID','Ticker / Symbol','Security Name','Security Type','Asset Class',
  'Currency','Current Price','Prior Price','Price Change','Price Change %',
  'Last Updated','Dividend / Share','Div Frequency','Exchange Rate to Base','Price in Base',
  'Active?','Notes'
];

const COL_W = [90,90,200,120,140,70,100,100,90,90,105,100,110,130,100,65,200];

// [Ticker, Name, Type, AssetClass, Currency, CurPrice, PriorPrice, LastUpd, DivPerShare, DivFreq, ExchRate, Active, Notes]
const SECURITIES = [
  // ETFs
  ['VTI',    'Vanguard Total Stock Market ETF',      'ETF','U.S. Stocks',         'USD',238.45,235.20,'2/14/2026',0.88,'Quarterly',1,'',false],
  ['VXUS',   'Vanguard Total International ETF',    'ETF','International Stocks', 'USD',63.28, 62.10,'2/14/2026',0.72,'Quarterly',1,'',false],
  ['BND',    'Vanguard Total Bond Market ETF',       'ETF','Bonds',               'USD',74.12, 73.80,'2/14/2026',0.32,'Monthly',  1,'',false],
  ['VNQ',    'Vanguard Real Estate ETF',             'ETF','REITs',               'USD',89.55, 91.20,'2/14/2026',0.72,'Quarterly',1,'',false],
  ['QQQ',    'Invesco QQQ Trust',                   'ETF','U.S. Stocks',         'USD',492.30,485.00,'2/14/2026',0.57,'Quarterly',1,'',false],
  ['SCHD',   'Schwab US Dividend Equity ETF',        'ETF','U.S. Stocks',         'USD',28.92, 28.40,'2/14/2026',0.26,'Quarterly',1,'',false],
  ['AGG',    'iShares Core US Aggregate Bond ETF',   'ETF','Bonds',               'USD',97.84, 97.20,'2/14/2026',0.30,'Monthly',  1,'',false],
  ['EMB',    'iShares EM USD Bond ETF',              'ETF','Bonds',               'USD',88.40, 87.60,'2/14/2026',0.48,'Monthly',  1,'',false],
  ['VWO',    'Vanguard FTSE Emerging Markets ETF',   'ETF','Emerging Markets',    'USD',45.18, 44.50,'2/14/2026',0.55,'Semiannual',1,'',false],
  ['IEFA',   'iShares Core MSCI EAFE ETF',           'ETF','International Stocks','USD',78.65, 77.90,'2/14/2026',0.68,'Semiannual',1,'',false],
  ['GLD',    'SPDR Gold Shares ETF',                 'ETF','Gold / Precious Metals','USD',225.80,222.50,'2/14/2026',0,'None',     1,'No dividend',false],
  ['BITO',   'ProShares Bitcoin Strategy ETF',       'ETF','Cryptocurrency',      'USD',29.44, 31.20,'2/14/2026',0,'None',      1,'',false],
  // Mutual Funds
  ['VTSAX',  'Vanguard Total Stock Market Index Adm','Mutual Fund','U.S. Stocks', 'USD',129.88,127.40,'2/14/2026',0,'None',      1,'Annual distribution',false],
  ['VBTLX',  'Vanguard Total Bond Market Index Adm', 'Mutual Fund','Bonds',       'USD',10.82, 10.78,'2/14/2026',0,'None',      1,'Monthly dist.',false],
  ['FXAIX',  'Fidelity 500 Index Fund',             'Mutual Fund','U.S. Stocks', 'USD',205.64,201.80,'2/14/2026',0,'None',      1,'Annual distribution',false],
  ['SWTSX',  'Schwab Total Stock Market Index Fund', 'Mutual Fund','U.S. Stocks', 'USD',82.34, 80.80,'2/14/2026',0,'None',      1,'',false],
  ['VGHCX',  'Vanguard Health Care Fund Investor',  'Mutual Fund','U.S. Stocks', 'USD',248.90,245.10,'2/14/2026',0,'None',      1,'',false],
  // Individual Stocks
  ['AAPL',   'Apple Inc.',                           'Stock','U.S. Stocks',       'USD',228.50,225.80,'2/14/2026',0.25,'Quarterly',1,'',false],
  ['MSFT',   'Microsoft Corporation',                'Stock','U.S. Stocks',       'USD',415.20,410.50,'2/14/2026',0.83,'Quarterly',1,'',false],
  ['NVDA',   'NVIDIA Corporation',                   'Stock','U.S. Stocks',       'USD',142.80,138.90,'2/14/2026',0.01,'Quarterly',1,'',false],
  ['JNJ',    'Johnson & Johnson',                    'Stock','U.S. Stocks',       'USD',152.30,153.90,'2/14/2026',1.24,'Quarterly',1,'',false],
  ['V',      'Visa Inc.',                            'Stock','U.S. Stocks',       'USD',310.45,308.80,'2/14/2026',0.52,'Quarterly',1,'',false],
  // REITs
  ['SPG',    'Simon Property Group Inc.',            'REIT','REITs',              'USD',165.40,162.80,'2/14/2026',2.05,'Quarterly',1,'',false],
  ['PLD',    'Prologis Inc.',                        'REIT','REITs',              'USD',118.70,120.50,'2/14/2026',0.96,'Quarterly',1,'',false],
  ['AMT',    'American Tower Corp.',                 'REIT','REITs',              'USD',210.20,212.40,'2/14/2026',1.62,'Quarterly',1,'',false],
  // Bonds
  ['STXBND', 'US Treasury Note 4.75% 2029',         'Bond','Bonds',              'USD',98.50, 97.80,'2/14/2026',23.75,'Semiannual',1,'$1000 par',false],
  ['CBND',   'Corp Bond Fund 4.5% 2028',            'Bond','Bonds',              'USD',97.20, 96.40,'2/14/2026',22.50,'Semiannual',1,'$1000 par',false],
  ['CD5',    '12-Month CD 5.25% (Ally)',             'CD','Cash',                 'USD',1000,  1000,'10/15/2025',52.50,'Annual',    1,'Matures Oct 2026',false],
  // Crypto
  ['BTC-USD','Bitcoin',                              'Cryptocurrency','Cryptocurrency','USD',98500,102000,'2/14/2026',0,'None',   1,'Held on exchange',false],
  ['ETH-USD','Ethereum',                             'Cryptocurrency','Cryptocurrency','USD',3240, 3380,'2/14/2026',0,'None',    1,'Held on exchange',false],
  // Cash / Money Market
  ['MKTSETL','Money Market Settlement Fund',         'Money Market Fund','Cash',  'USD',1.00,  1.00,'2/14/2026',0.0042,'Monthly',1,'NAV always $1.00',false],
];

(async () => {
  const vals = [];
  const fmt  = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,1100,0,17), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // Column widths
  COL_W.forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Title banner
  fmt.push({ mergeCells: { range: gridRange(SID,0,2,0,17), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${S}!A1`, values: [['PRICE UPDATES\nUpdate security prices manually here • all tabs pull from this table']] });
  fmt.push({ repeatCell: { range: gridRange(SID,0,2,0,17), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 46 }, fields: 'pixelSize' }});

  // Summary cards (row 3-4)
  const cards = [
    ['Total Securities', `=COUNTA($B$6:$B$1005)`],
    ['Active Securities', `=COUNTIF($P$6:$P$1005,TRUE)`],
    ['Stale Prices (>30 days)', `=IFERROR(SUMPRODUCT(($P$6:$P$1005=TRUE)*(($K$6:$K$1005<>"")*(TODAY()-$K$6:$K$1005>30))),0)`],
    ['Zero / Missing Price',    `=SUMPRODUCT(($P$6:$P$1005=TRUE)*($G$6:$G$1005=0))`],
  ];
  cards.forEach((card, i) => {
    const c = i * 4;
    fmt.push({ mergeCells: { range: gridRange(SID,2,3,c,c+3), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID,3,4,c,c+3), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}3`, values: [[card[0]]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}4`, values: [[card[1]]] });
    fmt.push({ repeatCell: { range: gridRange(SID,2,3,c,c+3), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrB), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,3,4,c,c+3), cell: { userEnteredFormat: {
      backgroundColor: hex(C.highlight), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 4 },
      properties: { pixelSize: 30 }, fields: 'pixelSize' }});
  });

  // Column headers (row 5, 0-indexed 4)
  vals.push({ range: `${S}!A5`, values: [HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID,4,5,0,17), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 },
    properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Data rows (starting row 6, 0-indexed 5)
  const DATA_START = 5;
  for (let i = 0; i < SECURITIES.length; i++) {
    const r = DATA_START + i;
    const [tick, name, type, assetClass, curr, curPrice, priorPrice, lastUpd, divPerShr, divFreq, exchRate, , notes] = SECURITIES[i];
    const bg = i % 2 === 0 ? C.panel : C.altRow;

    vals.push({ range: `${S}!A${r+1}`, values: [[`=IF(C${r+1}="","","SEC-"&TEXT(ROW()-5,"0000"))`]] });
    vals.push({ range: `${S}!B${r+1}`, values: [[tick]] });
    vals.push({ range: `${S}!C${r+1}`, values: [[name]] });
    vals.push({ range: `${S}!D${r+1}`, values: [[type]] });
    vals.push({ range: `${S}!E${r+1}`, values: [[assetClass]] });
    vals.push({ range: `${S}!F${r+1}`, values: [[curr]] });
    vals.push({ range: `${S}!G${r+1}`, values: [[curPrice]] });
    vals.push({ range: `${S}!H${r+1}`, values: [[priorPrice]] });
    vals.push({ range: `${S}!I${r+1}`, values: [[`=IFERROR(G${r+1}-H${r+1},"")`]] });
    vals.push({ range: `${S}!J${r+1}`, values: [[`=IFERROR((G${r+1}-H${r+1})/H${r+1},"")`]] });
    vals.push({ range: `${S}!K${r+1}`, values: [[lastUpd]] });
    vals.push({ range: `${S}!L${r+1}`, values: [[divPerShr]] });
    vals.push({ range: `${S}!M${r+1}`, values: [[divFreq]] });
    vals.push({ range: `${S}!N${r+1}`, values: [[exchRate]] });
    vals.push({ range: `${S}!O${r+1}`, values: [[`=IFERROR(G${r+1}*N${r+1},"")`]] });
    vals.push({ range: `${S}!P${r+1}`, values: [[true]] }); // Active
    vals.push({ range: `${S}!Q${r+1}`, values: [[notes]] });

    fmt.push({ repeatCell: { range: gridRange(SID,r,r+1,0,17), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 20 }, fields: 'pixelSize' }});
  }

  // Set one stale price for sample (CD5, index 27 → row DATA_START+27+1 = row 33)
  vals.push({ range: `${S}!K${DATA_START+27+1}`, values: [['10/15/2025']] }); // stale

  // Formula tint for calculated cols I(8),J(9),O(14)
  [8,9,14].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID,DATA_START,DATA_START+60,c,c+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula),
    }}, fields: 'userEnteredFormat.backgroundColor' }});
  });

  // Number formats
  // Current Price, Prior Price, Div/Share, Price in Base: $0.0000
  [6,7,11,14].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID,DATA_START,DATA_START+60,c,c+1), cell: { userEnteredFormat: {
      numberFormat: { type: 'CURRENCY', pattern: '$0.0000' }
    }}, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Price Change: $0.0000
  fmt.push({ repeatCell: { range: gridRange(SID,DATA_START,DATA_START+60,8,9), cell: { userEnteredFormat: {
    numberFormat: { type: 'NUMBER', pattern: '$0.0000' }
  }}, fields: 'userEnteredFormat.numberFormat' }});
  // Price Change %: 0.0%
  fmt.push({ repeatCell: { range: gridRange(SID,DATA_START,DATA_START+60,9,10), cell: { userEnteredFormat: {
    numberFormat: { type: 'PERCENT', pattern: '0.0%' }
  }}, fields: 'userEnteredFormat.numberFormat' }});
  // Last Updated: date
  fmt.push({ repeatCell: { range: gridRange(SID,DATA_START,DATA_START+60,10,11), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'MMM D, YYYY' }
  }}, fields: 'userEnteredFormat.numberFormat' }});

  // Validations
  fmt.push({ setDataValidation: { range: gridRange(SID,DATA_START,DATA_START+200,3,4),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$48:$A$57` }] }, showCustomUi: true, strict: false } }});
  fmt.push({ setDataValidation: { range: gridRange(SID,DATA_START,DATA_START+200,4,5),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$33:$A$45` }] }, showCustomUi: true, strict: false } }});
  fmt.push({ setDataValidation: { range: gridRange(SID,DATA_START,DATA_START+200,5,6),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$164:$A$172` }] }, showCustomUi: true, strict: false } }});
  fmt.push({ setDataValidation: { range: gridRange(SID,DATA_START,DATA_START+200,12,13),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$76:$A$81` }] }, showCustomUi: true, strict: false } }});
  fmt.push({ setDataValidation: { range: gridRange(SID,DATA_START,DATA_START+200,15,16),
    rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true } }});

  // Conditional: stale price (>30 days) → amber on Last Updated
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID,DATA_START,DATA_START+200,10,11)],
    booleanRule: { condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=AND($P6=TRUE,$K6<>"",(TODAY()-$K6)>30)` }] },
      format: { backgroundColor: hex(C.warning) }
    }
  }, index: 0 }});
  // Zero or negative price → attention
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID,DATA_START,DATA_START+200,6,7)],
    booleanRule: { condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=AND($P6=TRUE,$G6<=0)` }] },
      format: { backgroundColor: hex(C.attention) }
    }
  }, index: 1 }});
  // Positive price change → sage
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID,DATA_START,DATA_START+200,8,9)],
    booleanRule: { condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0.001' }] },
      format: { textFormat: { foregroundColor: hex(C.success), bold: true } }
    }
  }, index: 2 }});
  // Negative price change → attention
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID,DATA_START,DATA_START+200,8,9)],
    booleanRule: { condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '-0.001' }] },
      format: { textFormat: { foregroundColor: hex(C.attention), bold: true } }
    }
  }, index: 3 }});

  // Freeze rows 1:5, cols A:C
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 5, frozenColumnCount: 3 } }, fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount' }});

  await valuesBatchUpdate(id, vals, '05-prices values');
  await batchUpdate(id, fmt, '05-prices format');
  console.log('✅ Price Updates done. 31 securities entered.');
  console.log('  Security ID range: A6:A36');
  console.log('  Ticker range: B6:B36');
  console.log('  Current Price range: G6:G36');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
