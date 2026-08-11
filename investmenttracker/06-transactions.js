'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Transaction Log'];
const S = "'Transaction Log'";
const REF = "'Reference Data'";

const HDR  = 4; // 0-indexed row 5 in sheet
const DATA = 5; // 0-indexed row 6 in sheet

const HEADERS = [
  'Transaction ID','Date','Year','Month','Owner','Account ID','Account Name',
  'Security ID','Ticker','Security Name','Asset Class','Transaction Type',
  'Quantity','Price / Unit','Gross Amount','Fees','Taxes / Withholding',
  'Net Cash Flow','Div Reinvested?','Contribution Type','Notes'
];
const COL_W = [90,100,55,55,110,80,160,90,70,160,120,140,90,100,110,70,110,110,80,120,200];

// [Date, Owner, AccountID, SecurityID, TxType, Qty, Price, Fees, Taxes, DivReinv, ContribType, Notes]
// AccountIDs: ACT-001 through ACT-012
// SecurityIDs: SEC-0001 through SEC-0031
const TXNS = [
// ── 2024 Q1 ────────────────────────────────────────────────────────
['1/2/2024', 'Daniel Walsh','ACT-002','SEC-0001','Contribution',0,7000,0,0,false,'One-Time','2024 Roth IRA max contribution'],
['1/2/2024', 'Emily Walsh', 'ACT-005','SEC-0001','Contribution',0,7000,0,0,false,'One-Time','2024 Roth IRA max contribution'],
['1/5/2024', 'Daniel Walsh','ACT-001','SEC-0001','Buy',25,197.80,2.95,0,false,'','VTI — taxable core'],
['1/5/2024', 'Daniel Walsh','ACT-001','SEC-0018','Buy',10,185.50,0,0,false,'','AAPL'],
['1/8/2024', 'Emily Walsh', 'ACT-004','SEC-0002','Buy',50,58.40,0,0,false,'','VXUS international exposure'],
['1/10/2024','Daniel Walsh','ACT-002','SEC-0001','Buy',35,197.80,0,0,false,'','VTI in Roth IRA'],
['1/10/2024','Emily Walsh', 'ACT-005','SEC-0006','Buy',120,25.80,0,0,false,'','SCHD dividend focus in Roth'],
['1/15/2024','Daniel Walsh','ACT-003','','Contribution',0,1800,0,0,false,'Recurring','Jan 401k contribution'],
['1/15/2024','Emily Walsh', 'ACT-006','','Contribution',0,1600,0,0,false,'Recurring','Jan 403b contribution'],
['1/15/2024','Joint Household','ACT-007','SEC-0001','Buy',30,197.80,0,0,false,'','Joint account VTI'],
['1/16/2024','Daniel Walsh','ACT-003','SEC-0013','Buy',13.86,129.88,0,0,false,'','VTSAX in 401k (new units)'],
['1/20/2024','Daniel Walsh','ACT-001','SEC-0019','Buy',5,375.20,0,0,false,'','MSFT taxable'],
['1/22/2024','Joint Household','ACT-011','SEC-0001','Buy',50,197.80,0,0,false,'','Trust VTI purchase'],
['1/25/2024','Daniel Walsh','ACT-008','SEC-0001','Buy',15,197.80,0,0,false,'','HSA VTI investment'],
['1/25/2024','Emily Walsh', 'ACT-009','SEC-0001','Buy',10,197.80,0,0,false,'','Emily HSA VTI'],
['2/1/2024', 'Daniel Walsh','ACT-003','','Contribution',0,1800,0,0,false,'Recurring','Feb 401k contribution'],
['2/1/2024', 'Emily Walsh', 'ACT-006','','Contribution',0,1600,0,0,false,'Recurring','Feb 403b contribution'],
['2/5/2024', 'Emily Walsh', 'ACT-004','SEC-0007','Buy',40,97.10,0,0,false,'','AGG bond allocation'],
['2/8/2024', 'Joint Household','ACT-007','SEC-0002','Buy',60,58.40,0,0,false,'','Joint VXUS'],
['2/12/2024','Daniel Walsh','ACT-001','SEC-0011','Buy',5,186.20,0,0,false,'','GLD hedge'],
['2/15/2024','Daniel Walsh','ACT-001','SEC-0006','Buy',80,25.80,0,0,false,'','SCHD dividend income'],
['2/20/2024','Emily Walsh', 'ACT-005','SEC-0009','Buy',40,43.20,0,0,false,'','VWO emerging markets'],
['2/22/2024','Joint Household','ACT-010','SEC-0013','Buy',50,127.40,0,0,false,'','529 VTSAX'],
['3/1/2024', 'Daniel Walsh','ACT-003','','Contribution',0,1800,0,0,false,'Recurring','Mar 401k contribution'],
['3/1/2024', 'Emily Walsh', 'ACT-006','','Contribution',0,1600,0,0,false,'Recurring','Mar 403b contribution'],
['3/5/2024', 'Daniel Walsh','ACT-001','SEC-0020','Buy',20,82.50,0,0,false,'','NVDA — tech growth (low entry)'],
['3/8/2024', 'Emily Walsh', 'ACT-004','SEC-0021','Buy',15,155.60,0,0,false,'','JNJ defensive'],
['3/12/2024','Joint Household','ACT-007','SEC-0024','Buy',20,108.40,0,0,false,'','PLD logistics REIT'],
['3/15/2024','Daniel Walsh','ACT-002','SEC-0005','Buy',8,422.50,0,0,false,'','QQQ in Roth IRA'],
['3/20/2024','Emily Walsh', 'ACT-005','SEC-0010','Buy',35,74.80,0,0,false,'','IEFA international developed'],
['3/25/2024','Joint Household','ACT-011','SEC-0019','Buy',15,375.20,0,0,false,'','MSFT in Trust'],
// ── 2024 Q1 Dividends ────────────────────────────────────────────────
['3/28/2024','Daniel Walsh','ACT-001','SEC-0006','Dividend',0,83.20,0,8.32,false,'','SCHD Q1 2024 dividend - 80 shares × $1.04'],
['3/28/2024','Emily Walsh', 'ACT-005','SEC-0006','Dividend',0,124.80,0,12.48,false,'','SCHD Q1 2024 dividend - 120 shares × $1.04'],
['3/28/2024','Daniel Walsh','ACT-001','SEC-0001','Dividend',0,25.05,0,0,false,'','VTI Q1 dividend - 25 shares × $1.00'],
// ── 2024 Q2 ────────────────────────────────────────────────────────
['4/1/2024', 'Daniel Walsh','ACT-003','','Contribution',0,1800,0,0,false,'Recurring','Apr 401k + employer match'],
['4/1/2024', 'Daniel Walsh','ACT-003','','Contribution',0,720,0,0,false,'Employer Match','Employer 4% match on $18,000 eff'],
['4/1/2024', 'Emily Walsh', 'ACT-006','','Contribution',0,1600,0,0,false,'Recurring','Apr 403b contribution'],
['4/5/2024', 'Daniel Walsh','ACT-001','SEC-0022','Buy',10,290.40,0,0,false,'','V Visa payment network'],
['4/10/2024','Emily Walsh', 'ACT-004','SEC-0008','Buy',30,86.20,0,0,false,'','EMB EM bond'],
['4/15/2024','Daniel Walsh','ACT-012','SEC-0029','Buy',0.05,64000,0,25,0,false,'','BTC purchase'],
['4/20/2024','Emily Walsh', 'ACT-005','SEC-0006','Buy',40,26.20,0,0,false,'','SCHD reinforce position'],
['4/25/2024','Joint Household','ACT-007','SEC-0003','Buy',80,73.20,0,0,false,'','BND bonds to joint account'],
['5/1/2024', 'Daniel Walsh','ACT-003','','Contribution',0,1800,0,0,false,'Recurring','May 401k contribution'],
['5/1/2024', 'Daniel Walsh','ACT-003','','Contribution',0,720,0,0,false,'Employer Match','Employer match May'],
['5/1/2024', 'Emily Walsh', 'ACT-006','','Contribution',0,1600,0,0,false,'Recurring','May 403b'],
['5/8/2024', 'Joint Household','ACT-011','SEC-0011','Buy',10,186.20,0,0,false,'','GLD in Trust inflation hedge'],
['5/15/2024','Daniel Walsh','ACT-001','SEC-0018','Buy',5,178.90,0,0,false,'','AAPL add to position'],
['5/20/2024','Emily Walsh', 'ACT-004','SEC-0022','Buy',5,290.40,0,0,false,'','V Visa'],
['5/25/2024','Joint Household','ACT-007','SEC-0023','Buy',10,152.60,0,0,false,'','SPG mall REIT'],
['6/1/2024', 'Daniel Walsh','ACT-003','','Contribution',0,1800,0,0,false,'Recurring','Jun 401k contribution'],
['6/1/2024', 'Daniel Walsh','ACT-003','','Contribution',0,720,0,0,false,'Employer Match','Employer match Jun'],
['6/1/2024', 'Emily Walsh', 'ACT-006','','Contribution',0,1600,0,0,false,'Recurring','Jun 403b'],
['6/15/2024','Daniel Walsh','ACT-012','SEC-0030','Buy',1.5,3450,0,15,0,false,'','ETH purchase'],
// ── 2024 Q2 Dividends ────────────────────────────────────────────────
['6/27/2024','Daniel Walsh','ACT-001','SEC-0006','Dividend',0,83.20,0,8.32,false,'','SCHD Q2 2024 dividend'],
['6/27/2024','Emily Walsh', 'ACT-005','SEC-0006','Dividend',0,166.40,0,16.64,false,'','SCHD Q2 2024 - 160 shares'],
['6/28/2024','Daniel Walsh','ACT-001','SEC-0001','Dividend',0,25.05,0,0,false,'','VTI Q2 dividend'],
['6/28/2024','Joint Household','ACT-007','SEC-0001','Dividend',0,30.06,0,0,false,'','VTI Q2 - 30 shares'],
['6/30/2024','Daniel Walsh','ACT-001','SEC-0018','Dividend',0,12.50,0,1.25,false,'','AAPL Q2 dividend - 15 shares'],
['6/30/2024','Joint Household','ACT-011','SEC-0019','Dividend',0,12.45,0,1.25,false,'','MSFT Trust Q2'],
// ── 2024 Q3 ────────────────────────────────────────────────────────
['7/1/2024', 'Daniel Walsh','ACT-003','','Contribution',0,1800,0,0,false,'Recurring','Jul 401k'],
['7/1/2024', 'Daniel Walsh','ACT-003','','Contribution',0,720,0,0,false,'Employer Match','Jul match'],
['7/1/2024', 'Emily Walsh', 'ACT-006','','Contribution',0,1600,0,0,false,'Recurring','Jul 403b'],
['7/5/2024', 'Daniel Walsh','ACT-002','SEC-0020','Buy',25,110.20,0,0,false,'','NVDA buy in Roth — tax-free growth'],
['7/10/2024','Emily Walsh', 'ACT-005','SEC-0002','Buy',40,60.80,0,0,false,'','VXUS international in Roth'],
['7/15/2024','Joint Household','ACT-007','SEC-0025','Buy',15,195.40,0,0,false,'','AMT tower REIT'],
['7/20/2024','Daniel Walsh','ACT-001','SEC-0019','Buy',3,425.80,0,0,false,'','MSFT additional purchase'],
['7/25/2024','Emily Walsh', 'ACT-004','SEC-0009','Buy',25,44.10,0,0,false,'','VWO more EM'],
['8/1/2024', 'Daniel Walsh','ACT-003','','Contribution',0,1800,0,0,false,'Recurring','Aug 401k'],
['8/1/2024', 'Daniel Walsh','ACT-003','','Contribution',0,720,0,0,false,'Employer Match','Aug match'],
['8/1/2024', 'Emily Walsh', 'ACT-006','','Contribution',0,1600,0,0,false,'Recurring','Aug 403b'],
['8/5/2024', 'Emily Walsh', 'ACT-006','SEC-0016','Buy',19.45,82.34,0,0,false,'','SWTSX in 403b'],
['8/10/2024','Joint Household','ACT-010','SEC-0013','Buy',25,127.40,0,0,false,'','529 add'],
['8/20/2024','Emily Walsh', 'ACT-006','SEC-0014','Buy',46.21,10.82,0,0,false,'','VBTLX bonds in 403b'],
['8/25/2024','Daniel Walsh','ACT-001','SEC-0006','Reinvestment',3.22,25.80,0,0,true,'','SCHD dividend reinvestment'],
['9/1/2024', 'Daniel Walsh','ACT-003','','Contribution',0,1800,0,0,false,'Recurring','Sep 401k'],
['9/1/2024', 'Daniel Walsh','ACT-003','','Contribution',0,720,0,0,false,'Employer Match','Sep match'],
['9/1/2024', 'Emily Walsh', 'ACT-006','','Contribution',0,1600,0,0,false,'Recurring','Sep 403b'],
['9/10/2024','Daniel Walsh','ACT-002','SEC-0005','Buy',4,470.20,0,0,false,'','QQQ add'],
['9/15/2024','Joint Household','ACT-007','SEC-0003','Buy',40,72.80,0,0,false,'','BND add to bonds'],
// ── 2024 Q3 Dividends ──────────────────────────────────────────────
['9/27/2024','Daniel Walsh','ACT-001','SEC-0006','Dividend',0,83.20,0,8.32,false,'','SCHD Q3 2024'],
['9/27/2024','Emily Walsh', 'ACT-005','SEC-0006','Dividend',0,166.40,0,16.64,false,'','SCHD Q3 2024'],
['9/28/2024','Daniel Walsh','ACT-001','SEC-0001','Dividend',0,25.05,0,0,false,'','VTI Q3'],
['9/28/2024','Joint Household','ACT-007','SEC-0001','Dividend',0,30.06,0,0,false,'','VTI Joint Q3'],
['9/30/2024','Daniel Walsh','ACT-001','SEC-0018','Dividend',0,12.50,0,1.25,false,'','AAPL Q3 dividend'],
['9/30/2024','Emily Walsh', 'ACT-004','SEC-0021','Dividend',0,37.20,0,3.72,false,'','JNJ Q3 - 15 shares × $2.48'],
// ── 2024 Q4 ────────────────────────────────────────────────────────
['10/1/2024','Daniel Walsh','ACT-003','','Contribution',0,1800,0,0,false,'Recurring','Oct 401k'],
['10/1/2024','Daniel Walsh','ACT-003','','Contribution',0,720,0,0,false,'Employer Match','Oct match'],
['10/1/2024','Emily Walsh', 'ACT-006','','Contribution',0,1600,0,0,false,'Recurring','Oct 403b'],
['10/5/2024','Joint Household','ACT-007','SEC-0024','Buy',15,112.60,0,0,false,'','PLD add to REIT'],
['10/10/2024','Daniel Walsh','ACT-001','SEC-0020','Buy',10,130.40,0,0,false,'','NVDA add - still bullish'],
['10/15/2024','Daniel Walsh','ACT-001','SEC-0026','Buy',3,98.00,0,0,false,'','Treasury bond'],
['10/20/2024','Emily Walsh', 'ACT-004','SEC-0022','Buy',3,310.80,0,0,false,'','Visa add'],
['10/25/2024','Joint Household','ACT-011','SEC-0018','Buy',5,222.80,0,0,false,'','AAPL Trust buy'],
['11/1/2024','Daniel Walsh','ACT-003','','Contribution',0,1800,0,0,false,'Recurring','Nov 401k'],
['11/1/2024','Daniel Walsh','ACT-003','','Contribution',0,720,0,0,false,'Employer Match','Nov match'],
['11/1/2024','Emily Walsh', 'ACT-006','','Contribution',0,1600,0,0,false,'Recurring','Nov 403b'],
['11/5/2024','Daniel Walsh','ACT-012','SEC-0029','Buy',0.03,84000,0,15,0,false,'','BTC add - near ATH'],
['11/10/2024','Emily Walsh', 'ACT-005','SEC-0009','Buy',20,44.50,0,0,false,'','VWO add'],
['11/15/2024','Joint Household','ACT-007','SEC-0002','Buy',25,62.10,0,0,false,'','VXUS add'],
['11/20/2024','Daniel Walsh','ACT-001','SEC-0012','Buy',30,35.20,0,0,false,'','BITO crypto ETF — bought near peak'],
['12/1/2024','Daniel Walsh','ACT-003','','Contribution',0,1800,0,0,false,'Recurring','Dec 401k'],
['12/1/2024','Daniel Walsh','ACT-003','','Contribution',0,720,0,0,false,'Employer Match','Dec match'],
['12/1/2024','Emily Walsh', 'ACT-006','','Contribution',0,1600,0,0,false,'Recurring','Dec 403b'],
['12/5/2024','Emily Walsh', 'ACT-006','SEC-0017','Buy',8.04,248.90,0,0,false,'','VGHCX healthcare'],
['12/10/2024','Daniel Walsh','ACT-001','SEC-0006','Sell',20,27.20,2.95,0,false,'','Tax-loss harvest SCHD partial'],
// ── 2024 Q4 Dividends & Cap Gains ─────────────────────────────────
['12/20/2024','Daniel Walsh','ACT-001','SEC-0006','Dividend',0,83.20,0,8.32,false,'','SCHD Q4 2024'],
['12/20/2024','Emily Walsh', 'ACT-005','SEC-0006','Dividend',0,166.40,0,16.64,false,'','SCHD Q4 2024'],
['12/20/2024','Joint Household','ACT-007','SEC-0001','Dividend',0,50.10,0,0,false,'','VTI Q4 joint'],
['12/27/2024','Daniel Walsh','ACT-001','SEC-0001','Dividend',0,25.05,0,0,false,'','VTI Q4'],
['12/27/2024','Daniel Walsh','ACT-001','SEC-0018','Dividend',0,12.50,0,1.25,false,'','AAPL Q4'],
['12/28/2024','Daniel Walsh','ACT-001','SEC-0019','Dividend',0,24.90,0,2.49,false,'','MSFT Q4'],
['12/28/2024','Joint Household','ACT-011','SEC-0019','Dividend',0,24.90,0,2.49,false,'','MSFT Trust Q4'],
['12/30/2024','Daniel Walsh','ACT-003','SEC-0013','Capital Gain Distribution',0,850,0,0,false,'','VTSAX cap gain dist 2024'],
['12/30/2024','Emily Walsh', 'ACT-006','SEC-0017','Capital Gain Distribution',0,420,0,0,false,'','VGHCX annual cap gain'],
['12/30/2024','Emily Walsh', 'ACT-004','SEC-0021','Dividend',0,37.20,0,3.72,false,'','JNJ Q4'],
// ── 2025 Q1 ────────────────────────────────────────────────────────
['1/2/2025', 'Daniel Walsh','ACT-002','SEC-0001','Contribution',0,7000,0,0,false,'One-Time','2025 Roth IRA max'],
['1/2/2025', 'Emily Walsh', 'ACT-005','SEC-0001','Contribution',0,7000,0,0,false,'One-Time','2025 Roth IRA max'],
['1/3/2025', 'Daniel Walsh','ACT-002','SEC-0001','Buy',29.01,241.30,0,0,false,'','VTI Roth — full $7k deployment'],
['1/3/2025', 'Emily Walsh', 'ACT-005','SEC-0006','Buy',268,26.10,0,0,false,'','SCHD Roth more shares'],
['1/5/2025', 'Daniel Walsh','ACT-001','SEC-0001','Buy',20,241.30,0,0,false,'','VTI taxable Jan 2025'],
['1/8/2025', 'Joint Household','ACT-007','SEC-0001','Buy',25,241.30,0,0,false,'','Joint VTI'],
['1/10/2025','Emily Walsh', 'ACT-004','SEC-0010','Buy',30,76.40,0,0,false,'','IEFA add'],
['1/15/2025','Daniel Walsh','ACT-003','','Contribution',0,1875,0,0,false,'Recurring','Jan 401k 2025'],
['1/15/2025','Daniel Walsh','ACT-003','','Contribution',0,750,0,0,false,'Employer Match','Jan match 2025'],
['1/15/2025','Emily Walsh', 'ACT-006','','Contribution',0,1667,0,0,false,'Recurring','Jan 403b 2025'],
['1/20/2025','Joint Household','ACT-011','SEC-0001','Buy',30,241.30,0,0,false,'','Trust VTI 2025'],
['1/22/2025','Daniel Walsh','ACT-008','SEC-0001','Buy',12,241.30,0,0,false,'','HSA VTI 2025'],
['1/25/2025','Emily Walsh', 'ACT-009','SEC-0001','Buy',8,241.30,0,0,false,'','Emily HSA 2025'],
['2/1/2025', 'Daniel Walsh','ACT-003','','Contribution',0,1875,0,0,false,'Recurring','Feb 401k'],
['2/1/2025', 'Daniel Walsh','ACT-003','','Contribution',0,750,0,0,false,'Employer Match','Feb match'],
['2/1/2025', 'Emily Walsh', 'ACT-006','','Contribution',0,1667,0,0,false,'Recurring','Feb 403b'],
['2/10/2025','Daniel Walsh','ACT-001','SEC-0019','Buy',2,432.10,0,0,false,'','MSFT add'],
['2/15/2025','Emily Walsh', 'ACT-004','SEC-0007','Buy',20,97.80,0,0,false,'','AGG add bonds'],
['2/20/2025','Joint Household','ACT-007','SEC-0023','Buy',5,159.80,0,0,false,'','SPG add REIT'],
['3/1/2025', 'Daniel Walsh','ACT-003','','Contribution',0,1875,0,0,false,'Recurring','Mar 401k'],
['3/1/2025', 'Daniel Walsh','ACT-003','','Contribution',0,750,0,0,false,'Employer Match','Mar match'],
['3/1/2025', 'Emily Walsh', 'ACT-006','','Contribution',0,1667,0,0,false,'Recurring','Mar 403b'],
['3/5/2025', 'Daniel Walsh','ACT-001','SEC-0020','Buy',5,115.40,0,0,false,'','NVDA add to position'],
// ── 2025 Q1 Dividends ────────────────────────────────────────────────
['3/28/2025','Daniel Walsh','ACT-001','SEC-0006','Dividend',0,72.80,0,7.28,false,'','SCHD Q1 2025 — 60 shares after harvest'],
['3/28/2025','Emily Walsh', 'ACT-005','SEC-0006','Dividend',0,432.00,0,43.20,false,'','SCHD Q1 2025 — 428 shares × $1.01'],
['3/28/2025','Daniel Walsh','ACT-001','SEC-0001','Dividend',0,45.09,0,0,false,'','VTI Q1 2025 — 45 shares × $1.00'],
['3/28/2025','Joint Household','ACT-007','SEC-0001','Dividend',0,55.11,0,0,false,'','VTI Joint Q1 2025'],
['3/31/2025','Daniel Walsh','ACT-001','SEC-0018','Dividend',0,18.75,0,1.88,false,'','AAPL Q1 2025 - 15 shares'],
['3/31/2025','Emily Walsh', 'ACT-004','SEC-0021','Dividend',0,37.20,0,3.72,false,'','JNJ Q1 2025'],
['3/31/2025','Daniel Walsh','ACT-001','SEC-0022','Dividend',0,26.00,0,2.60,false,'','Visa Q1 2025 - 10 shares × $2.60'],
// ── 2025 Q2 ────────────────────────────────────────────────────────
['4/1/2025', 'Daniel Walsh','ACT-003','','Contribution',0,1875,0,0,false,'Recurring','Apr 401k'],
['4/1/2025', 'Daniel Walsh','ACT-003','','Contribution',0,750,0,0,false,'Employer Match','Apr match'],
['4/1/2025', 'Emily Walsh', 'ACT-006','','Contribution',0,1667,0,0,false,'Recurring','Apr 403b'],
['4/5/2025', 'Emily Walsh', 'ACT-005','SEC-0010','Buy',15,77.90,0,0,false,'','IEFA Roth'],
['4/10/2025','Joint Household','ACT-010','SEC-0014','Buy',92.42,10.82,0,0,false,'','VBTLX in 529'],
['4/15/2025','Daniel Walsh','ACT-001','SEC-0018','Buy',3,196.40,0,0,false,'','AAPL buy on pullback'],
['4/20/2025','Emily Walsh', 'ACT-004','SEC-0008','Buy',15,87.20,0,0,false,'','EMB add'],
['5/1/2025', 'Daniel Walsh','ACT-003','','Contribution',0,1875,0,0,false,'Recurring','May 401k'],
['5/1/2025', 'Daniel Walsh','ACT-003','','Contribution',0,750,0,0,false,'Employer Match','May match'],
['5/1/2025', 'Emily Walsh', 'ACT-006','','Contribution',0,1667,0,0,false,'Recurring','May 403b'],
['5/12/2025','Daniel Walsh','ACT-012','SEC-0029','Sell',0.02,96800,0,15,0,false,'','BTC partial sale — took profit'],
['5/15/2025','Joint Household','ACT-007','SEC-0025','Buy',5,205.60,0,0,false,'','AMT add'],
['5/20/2025','Emily Walsh', 'ACT-004','SEC-0022','Buy',2,305.20,0,0,false,'','V add'],
['6/1/2025', 'Daniel Walsh','ACT-003','','Contribution',0,1875,0,0,false,'Recurring','Jun 401k'],
['6/1/2025', 'Daniel Walsh','ACT-003','','Contribution',0,750,0,0,false,'Employer Match','Jun match'],
['6/1/2025', 'Emily Walsh', 'ACT-006','','Contribution',0,1667,0,0,false,'Recurring','Jun 403b'],
['6/5/2025', 'Daniel Walsh','ACT-001','SEC-0006','Buy',20,27.10,0,0,false,'','SCHD replace sold shares'],
// ── 2025 Q2 Dividends ────────────────────────────────────────────────
['6/27/2025','Daniel Walsh','ACT-001','SEC-0006','Dividend',0,80.60,0,8.06,false,'','SCHD Q2 2025 — 80 shares'],
['6/27/2025','Emily Walsh', 'ACT-005','SEC-0006','Dividend',0,432.00,0,43.20,false,'','SCHD Q2 2025'],
['6/28/2025','Daniel Walsh','ACT-001','SEC-0001','Dividend',0,45.09,0,0,false,'','VTI Q2 2025'],
['6/28/2025','Joint Household','ACT-007','SEC-0001','Dividend',0,80.15,0,0,false,'','VTI Joint Q2 - 80 shares'],
['6/30/2025','Daniel Walsh','ACT-001','SEC-0018','Dividend',0,18.75,0,1.88,false,'','AAPL Q2 2025'],
['6/30/2025','Emily Walsh', 'ACT-004','SEC-0021','Dividend',0,37.20,0,3.72,false,'','JNJ Q2 2025'],
['6/30/2025','Joint Household','ACT-011','SEC-0023','Dividend',0,41.00,0,4.10,false,'','SPG Trust Q2 dividend'],
// ── 2025 Q3 ────────────────────────────────────────────────────────
['7/1/2025', 'Daniel Walsh','ACT-003','','Contribution',0,1875,0,0,false,'Recurring','Jul 401k'],
['7/1/2025', 'Daniel Walsh','ACT-003','','Contribution',0,750,0,0,false,'Employer Match','Jul match'],
['7/1/2025', 'Emily Walsh', 'ACT-006','','Contribution',0,1667,0,0,false,'Recurring','Jul 403b'],
['7/5/2025', 'Emily Walsh', 'ACT-005','SEC-0002','Buy',20,63.10,0,0,false,'','VXUS Roth add'],
['7/10/2025','Joint Household','ACT-011','SEC-0025','Buy',10,205.60,0,0,false,'','AMT Trust add'],
['7/15/2025','Daniel Walsh','ACT-001','SEC-0020','Sell',15,128.30,0,0,false,'','NVDA partial sell - lock gains at 55%'],
['7/20/2025','Emily Walsh', 'ACT-004','SEC-0007','Buy',10,97.50,0,0,false,'','AGG bonds'],
['8/1/2025', 'Daniel Walsh','ACT-003','','Contribution',0,1875,0,0,false,'Recurring','Aug 401k'],
['8/1/2025', 'Daniel Walsh','ACT-003','','Contribution',0,750,0,0,false,'Employer Match','Aug match'],
['8/1/2025', 'Emily Walsh', 'ACT-006','','Contribution',0,1667,0,0,false,'Recurring','Aug 403b'],
['8/5/2025', 'Daniel Walsh','ACT-001','SEC-0027','Buy',3,96.40,0,0,false,'','Corp bond add'],
['8/10/2025','Joint Household','ACT-007','SEC-0009','Buy',30,44.80,0,0,false,'','VWO emerging markets joint'],
['9/1/2025', 'Daniel Walsh','ACT-003','','Contribution',0,1875,0,0,false,'Recurring','Sep 401k'],
['9/1/2025', 'Daniel Walsh','ACT-003','','Contribution',0,750,0,0,false,'Employer Match','Sep match'],
['9/1/2025', 'Emily Walsh', 'ACT-006','','Contribution',0,1667,0,0,false,'Recurring','Sep 403b'],
['9/5/2025', 'Emily Walsh', 'ACT-009','SEC-0001','Buy',5,238.45,0,0,false,'','Emily HSA more VTI 2025'],
// ── 2025 Q3 Dividends ─────────────────────────────────────────────
['9/26/2025','Daniel Walsh','ACT-001','SEC-0006','Dividend',0,80.60,0,8.06,false,'','SCHD Q3 2025'],
['9/26/2025','Emily Walsh', 'ACT-005','SEC-0006','Dividend',0,432.00,0,43.20,false,'','SCHD Q3 2025'],
['9/27/2025','Daniel Walsh','ACT-001','SEC-0001','Dividend',0,45.09,0,0,false,'','VTI Q3 2025'],
['9/27/2025','Joint Household','ACT-007','SEC-0001','Dividend',0,80.15,0,0,false,'','VTI Joint Q3 2025'],
['9/30/2025','Joint Household','ACT-011','SEC-0024','Dividend',0,57.60,0,5.76,false,'','PLD dividend Q3 - 30 shares'],
['9/30/2025','Emily Walsh', 'ACT-004','SEC-0021','Dividend',0,37.20,0,3.72,false,'','JNJ Q3 2025'],
// ── 2025 Q4 ────────────────────────────────────────────────────────
['10/1/2025','Daniel Walsh','ACT-003','','Contribution',0,1875,0,0,false,'Recurring','Oct 401k'],
['10/1/2025','Daniel Walsh','ACT-003','','Contribution',0,750,0,0,false,'Employer Match','Oct match'],
['10/1/2025','Emily Walsh', 'ACT-006','','Contribution',0,1667,0,0,false,'Recurring','Oct 403b'],
['10/5/2025','Daniel Walsh','ACT-001','SEC-0020','Buy',10,138.60,0,0,false,'','NVDA re-enter after partial sale'],
['10/10/2025','Joint Household','ACT-007','SEC-0004','Buy',25,88.20,0,0,false,'','VNQ REIT add to joint'],
['10/15/2025','Daniel Walsh','ACT-001','SEC-0028','Buy',3,1000,0,0,false,'','12-month CD purchase 5.25%'],
['11/1/2025','Daniel Walsh','ACT-003','','Contribution',0,1875,0,0,false,'Recurring','Nov 401k'],
['11/1/2025','Daniel Walsh','ACT-003','','Contribution',0,750,0,0,false,'Employer Match','Nov match'],
['11/1/2025','Emily Walsh', 'ACT-006','','Contribution',0,1667,0,0,false,'Recurring','Nov 403b'],
['11/8/2025','Daniel Walsh','ACT-012','SEC-0030','Buy',0.5,3200,0,10,0,false,'','ETH add'],
['11/15/2025','Emily Walsh', 'ACT-004','SEC-0002','Buy',15,63.28,0,0,false,'','VXUS add international'],
['11/20/2025','Joint Household','ACT-010','SEC-0013','Buy',15,129.88,0,0,false,'','529 VTSAX annual add'],
['12/1/2025','Daniel Walsh','ACT-003','','Contribution',0,1875,0,0,false,'Recurring','Dec 401k'],
['12/1/2025','Daniel Walsh','ACT-003','','Contribution',0,750,0,0,false,'Employer Match','Dec match'],
['12/1/2025','Emily Walsh', 'ACT-006','','Contribution',0,1667,0,0,false,'Recurring','Dec 403b'],
['12/5/2025','Emily Walsh', 'ACT-005','SEC-0006','Dividend',0,432.00,0,43.20,false,'','SCHD Q4 2025'],
['12/5/2025','Daniel Walsh','ACT-001','SEC-0006','Dividend',0,80.60,0,8.06,false,'','SCHD Q4 2025'],
['12/10/2025','Daniel Walsh','ACT-001','SEC-0001','Dividend',0,45.09,0,0,false,'','VTI Q4 2025'],
['12/10/2025','Joint Household','ACT-007','SEC-0001','Dividend',0,80.15,0,0,false,'','VTI Joint Q4 2025'],
['12/15/2025','Daniel Walsh','ACT-001','SEC-0018','Dividend',0,18.75,0,1.88,false,'','AAPL Q4 2025'],
['12/20/2025','Daniel Walsh','ACT-003','SEC-0013','Capital Gain Distribution',0,1200,0,0,false,'','VTSAX cap gain 2025'],
['12/20/2025','Emily Walsh', 'ACT-006','SEC-0017','Capital Gain Distribution',0,580,0,0,false,'','VGHCX cap gain 2025'],
['12/28/2025','Joint Household','ACT-011','SEC-0019','Dividend',0,24.90,0,2.49,false,'','MSFT Trust Q4 2025'],
['12/28/2025','Daniel Walsh','ACT-001','SEC-0006','Fee',0,8.50,8.50,0,false,'','Fidelity advisory fee Q4'],
// ── 2026 (Jan–Feb) ─────────────────────────────────────────────────
['1/2/2026', 'Daniel Walsh','ACT-002','SEC-0001','Contribution',0,7000,0,0,false,'One-Time','2026 Roth IRA max'],
['1/2/2026', 'Emily Walsh', 'ACT-005','SEC-0001','Contribution',0,7000,0,0,false,'One-Time','2026 Roth IRA max'],
['1/5/2026', 'Daniel Walsh','ACT-001','SEC-0001','Buy',20,235.20,0,0,false,'','VTI Jan 2026'],
['1/5/2026', 'Emily Walsh', 'ACT-004','SEC-0001','Buy',15,235.20,0,0,false,'','VXUS → VTI for simplicity'],
['1/8/2026', 'Joint Household','ACT-007','SEC-0003','Buy',50,73.80,0,0,false,'','BND rebalance — more bonds'],
['1/10/2026','Daniel Walsh','ACT-002','SEC-0001','Buy',29.76,235.20,0,0,false,'','VTI Roth 2026 deployment'],
['1/15/2026','Daniel Walsh','ACT-003','','Contribution',0,1875,0,0,false,'Recurring','Jan 401k 2026'],
['1/15/2026','Daniel Walsh','ACT-003','','Contribution',0,750,0,0,false,'Employer Match','Jan match 2026'],
['1/15/2026','Emily Walsh', 'ACT-006','','Contribution',0,1667,0,0,false,'Recurring','Jan 403b 2026'],
['1/20/2026','Emily Walsh', 'ACT-005','SEC-0006','Buy',50,28.40,0,0,false,'','SCHD Roth 2026'],
['1/28/2026','Daniel Walsh','ACT-001','SEC-0006','Dividend',0,80.60,0,8.06,false,'','SCHD Q1 2026 early pay'],
['2/1/2026', 'Daniel Walsh','ACT-003','','Contribution',0,1875,0,0,false,'Recurring','Feb 401k'],
['2/1/2026', 'Daniel Walsh','ACT-003','','Contribution',0,750,0,0,false,'Employer Match','Feb match'],
['2/1/2026', 'Emily Walsh', 'ACT-006','','Contribution',0,1667,0,0,false,'Recurring','Feb 403b'],
['2/5/2026', 'Joint Household','ACT-007','SEC-0001','Buy',10,238.45,0,0,false,'','VTI add joint'],
['2/10/2026','Emily Walsh', 'ACT-009','SEC-0003','Buy',10,74.12,0,0,false,'','BND in Emily HSA'],
['2/14/2026','Daniel Walsh','ACT-001','SEC-0020','Buy',5,140.20,0,0,false,'','NVDA 2026 add'],
];

// Net Cash Flow logic
function netCashFlow(txType, gross, fees, taxes) {
  if (txType === 'Buy') return -(gross + fees);
  if (txType === 'Sell') return gross - fees - taxes;
  if (txType === 'Contribution') return gross;
  if (txType === 'Withdrawal') return -gross;
  if (txType === 'Dividend' || txType === 'Interest' || txType === 'Capital Gain Distribution') return gross - taxes;
  if (txType === 'Fee' || txType === 'Tax Withholding') return -(fees + taxes);
  if (txType === 'Transfer In' || txType === 'Transfer Out' || txType === 'Reinvestment' || txType === 'Stock Split' || txType === 'Adjustment') return 0;
  return 0;
}

(async () => {
  const vals = [];
  const fmt  = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,5100,0,21), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // Column widths
  COL_W.forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Title banner
  fmt.push({ mergeCells: { range: gridRange(SID,0,2,0,21), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${S}!A1`, values: [['TRANSACTION LOG\nRecord every buy, sell, contribution, withdrawal, dividend, fee, and transfer']] });
  fmt.push({ repeatCell: { range: gridRange(SID,0,2,0,21), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 46 }, fields: 'pixelSize' }});

  // Summary cards (rows 3-4)
  const cds = [
    ['Total Transactions', `=COUNTA($B$6:$B$5005)`],
    ['Buy Transactions',   `=COUNTIF($L$6:$L$5005,"Buy")`],
    ['Dividends',          `=COUNTIF($L$6:$L$5005,"Dividend")`],
    ['Contributions YTD',  `=IFERROR(SUMPRODUCT((YEAR($B$6:$B$5005)=YEAR(TODAY()))*($L$6:$L$5005="Contribution")*($O$6:$O$5005)),0)`],
    ['Total Fees Paid',    `=IFERROR(SUMIF($L$6:$L$5005,"Fee",$O$6:$O$5005)+SUMIF($P$6:$P$5005,">"&0,$P$6:$P$5005),0)`],
  ];
  cds.forEach((card, i) => {
    const c = i * 4;
    if (c >= 20) return;
    const span = i === 4 ? 1 : 4;
    fmt.push({ mergeCells: { range: gridRange(SID,2,3,c,c+span), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID,3,4,c,c+span), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}3`, values: [[card[0]]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}4`, values: [[card[1]]] });
    fmt.push({ repeatCell: { range: gridRange(SID,2,3,c,c+span), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrB), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,3,4,c,c+span), cell: { userEnteredFormat: {
      backgroundColor: hex(C.highlight), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 4 },
      properties: { pixelSize: 30 }, fields: 'pixelSize' }});
  });

  // Column headers
  vals.push({ range: `${S}!A5`, values: [HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID,HDR,HDR+1,0,21), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: HDR, endIndex: HDR+1 },
    properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Data rows
  for (let i = 0; i < TXNS.length; i++) {
    const r = DATA + i;
    const [date, owner, acctId, secId, txType, qty, price, fees, taxes, divReinv, contribType, notes] = TXNS[i];
    const gross = qty * price;
    const ncf   = netCashFlow(txType, gross, fees, taxes);
    const bg = i % 2 === 0 ? C.panel : C.altRow;

    // Formula cells
    vals.push({ range: `${S}!A${r+1}`, values: [[`=IF(B${r+1}="","","TXN-"&TEXT(ROW()-5,"00000"))`]] });
    vals.push({ range: `${S}!B${r+1}`, values: [[date]] });
    vals.push({ range: `${S}!C${r+1}`, values: [[`=IFERROR(YEAR(B${r+1}),"")`]] });
    vals.push({ range: `${S}!D${r+1}`, values: [[`=IFERROR(MONTH(B${r+1}),"")`]] });
    vals.push({ range: `${S}!E${r+1}`, values: [[owner]] });
    vals.push({ range: `${S}!F${r+1}`, values: [[acctId]] });
    vals.push({ range: `${S}!G${r+1}`, values: [[`=IFERROR(VLOOKUP(F${r+1},'Account Tracker'!$A$6:$B$305,2,FALSE),"")`]] });
    vals.push({ range: `${S}!H${r+1}`, values: [[secId]] });
    vals.push({ range: `${S}!I${r+1}`, values: [[`=IFERROR(VLOOKUP(H${r+1},'Price Updates'!$A$6:$B$1005,2,FALSE),"")`]] });
    vals.push({ range: `${S}!J${r+1}`, values: [[`=IFERROR(VLOOKUP(H${r+1},'Price Updates'!$A$6:$C$1005,3,FALSE),"")`]] });
    vals.push({ range: `${S}!K${r+1}`, values: [[`=IFERROR(VLOOKUP(H${r+1},'Price Updates'!$A$6:$E$1005,5,FALSE),"")`]] });
    vals.push({ range: `${S}!L${r+1}`, values: [[txType]] });
    vals.push({ range: `${S}!M${r+1}`, values: [[qty]] });
    vals.push({ range: `${S}!N${r+1}`, values: [[price]] });
    vals.push({ range: `${S}!O${r+1}`, values: [[`=IFERROR(M${r+1}*N${r+1},0)`]] });
    vals.push({ range: `${S}!P${r+1}`, values: [[fees]] });
    vals.push({ range: `${S}!Q${r+1}`, values: [[taxes]] });
    vals.push({ range: `${S}!R${r+1}`, values: [[`=IFERROR(IF(OR(L${r+1}="Buy",L${r+1}="Fee"),-(O${r+1}+P${r+1}),IF(L${r+1}="Sell",O${r+1}-P${r+1}-Q${r+1},IF(L${r+1}="Withdrawal",-O${r+1},IF(OR(L${r+1}="Transfer In",L${r+1}="Transfer Out",L${r+1}="Reinvestment",L${r+1}="Stock Split",L${r+1}="Adjustment"),0,O${r+1}-Q${r+1})))),"")`]] });
    vals.push({ range: `${S}!S${r+1}`, values: [[divReinv]] });
    vals.push({ range: `${S}!T${r+1}`, values: [[contribType]] });
    vals.push({ range: `${S}!U${r+1}`, values: [[notes]] });

    fmt.push({ repeatCell: { range: gridRange(SID,r,r+1,0,21), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 8, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 19 }, fields: 'pixelSize' }});
  }

  // Number formats on data range
  // Date col B
  fmt.push({ repeatCell: { range: gridRange(SID,DATA,DATA+TXNS.length,1,2), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'M/D/YYYY' }
  }}, fields: 'userEnteredFormat.numberFormat' }});
  // Currency cols N (price), O (gross), P (fees), Q (taxes), R (ncf)
  [13,14,15,16,17].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID,DATA,DATA+TXNS.length,c,c+1), cell: { userEnteredFormat: {
      numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' }
    }}, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Quantity col M: 0.0000
  fmt.push({ repeatCell: { range: gridRange(SID,DATA,DATA+TXNS.length,12,13), cell: { userEnteredFormat: {
    numberFormat: { type: 'NUMBER', pattern: '0.0000' }
  }}, fields: 'userEnteredFormat.numberFormat' }});
  // Formula tint cols A,C,D,G,I,J,K,O,R
  [0,2,3,6,8,9,10,14,17].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID,DATA,DATA+TXNS.length,c,c+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula),
    }}, fields: 'userEnteredFormat.backgroundColor' }});
  });

  // Validations
  fmt.push({ setDataValidation: { range: gridRange(SID,DATA,DATA+5000,4,5),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$4:$A$6` }] }, showCustomUi: true, strict: false } }});
  fmt.push({ setDataValidation: { range: gridRange(SID,DATA,DATA+5000,11,12),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$60:$A$73` }] }, showCustomUi: true, strict: false } }});
  fmt.push({ setDataValidation: { range: gridRange(SID,DATA,DATA+5000,18,19),
    rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true } }});
  fmt.push({ setDataValidation: { range: gridRange(SID,DATA,DATA+5000,19,20),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$138:$A$143` }] }, showCustomUi: true, strict: false } }});

  // Conditional: Dividend → info blue
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID,DATA,DATA+5000,0,21)],
    booleanRule: { condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$L6="Dividend"` }] },
      format: { backgroundColor: hex('#EEF3F8') }
    }
  }, index: 0 }});
  // Contribution → pale green
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID,DATA,DATA+5000,0,21)],
    booleanRule: { condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=OR($L6="Contribution",$L6="Employer Match")` }] },
      format: { backgroundColor: hex('#EBF5EE') }
    }
  }, index: 1 }});
  // Sell → pale amber
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID,DATA,DATA+5000,0,21)],
    booleanRule: { condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$L6="Sell"` }] },
      format: { backgroundColor: hex('#FBF4E0') }
    }
  }, index: 2 }});

  // Freeze rows 1:5 and cols A:D
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 5, frozenColumnCount: 4 } }, fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount' }});

  await valuesBatchUpdate(id, vals, '06-transactions values');
  await batchUpdate(id, fmt, '06-transactions format');
  console.log(`✅ Transaction Log done. ${TXNS.length} transactions entered.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
