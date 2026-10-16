'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BUD = sheetMap['💰 Budget & Expense Tracker'];

const TXN_HEADERS = ['Date','Category','Description','Type','Amount','Payment Method','Notes'];
const TXN_ROWS = [
  ['2025-01-03','Housing','January Rent','Expense',1850,'Bank Transfer',''],
  ['2025-01-05','Food','Weekly groceries - Trader Joe\'s','Expense',112.47,'Debit Card',''],
  ['2025-01-08','Transportation','Gas fill-up','Expense',58.20,'Credit Card',''],
  ['2025-01-10','Personal','Coffee & lunch out','Expense',24.80,'Credit Card',''],
  ['2025-01-15','Income','Paycheck','Income',3200,'Direct Deposit',''],
  ['2025-01-15','Utilities','Electric bill','Expense',94.50,'Auto-Pay',''],
  ['2025-01-16','Entertainment','Netflix subscription','Expense',15.99,'Credit Card',''],
  ['2025-01-18','Health','Gym membership','Expense',45.00,'Credit Card',''],
  ['2025-01-20','Food','Groceries','Expense',89.30,'Debit Card',''],
  ['2025-01-22','Savings','Transfer to savings','Expense',300,'Bank Transfer','Monthly savings goal'],
  ['2025-01-25','Personal','Haircut','Expense',45,'Cash',''],
  ['2025-01-28','Food','Takeout - sushi night','Expense',62.40,'Credit Card',''],
  ['2025-01-30','Income','Freelance payment','Income',850,'Bank Transfer','Web project'],
  ['2025-02-01','Housing','Renter\'s insurance','Expense',22.50,'Auto-Pay',''],
  ['2025-02-03','Transportation','Monthly bus pass','Expense',75,'Cash',''],
  ['2025-02-05','Food','Groceries','Expense',104.85,'Debit Card',''],
  ['2025-02-08','Debt Payments','Student loan','Expense',250,'Auto-Pay',''],
  ['2025-02-10','Personal','Amazon order','Expense',38.99,'Credit Card','Household supplies'],
  ['2025-02-14','Entertainment','Dinner out - Valentine\'s','Expense',95,'Credit Card',''],
  ['2025-02-15','Income','Paycheck','Income',3200,'Direct Deposit',''],
  ['2025-02-18','Health','Dentist co-pay','Expense',50,'Credit Card',''],
  ['2025-02-20','Food','Groceries','Expense',98.60,'Debit Card',''],
  ['2025-02-22','Savings','Transfer to savings','Expense',300,'Bank Transfer',''],
  ['2025-02-25','Utilities','Phone bill','Expense',65,'Auto-Pay',''],
  ['2025-02-27','Other','Car registration renewal','Expense',120,'Online','Annual fee'],
  ['2025-03-01','Housing','March Rent','Expense',1850,'Bank Transfer',''],
  ['2025-03-03','Food','Grocery run','Expense',91.20,'Debit Card',''],
  ['2025-03-07','Transportation','Uber rides this week','Expense',28.50,'Credit Card',''],
  ['2025-03-12','Entertainment','Concert tickets','Expense',110,'Credit Card','Local show'],
  ['2025-03-15','Income','Paycheck','Income',3200,'Direct Deposit',''],
];
const BILL_HEADERS = ['Bill Name','Amount','Due Date','Paid'];
const BILL_ROWS = [
  ['Rent',1850,'2025-04-01',false],
  ['Electric',94.50,'2025-04-05',false],
  ['Phone',65,'2025-04-10',true],
  ['Internet',59.99,'2025-04-10',false],
  ['Renter\'s Insurance',22.50,'2025-04-01',true],
  ['Gym Membership',45,'2025-04-15',false],
  ['Student Loan',250,'2025-04-08',true],
  ['Netflix',15.99,'2025-04-16',true],
];

(async () => {
  if (TXN_ROWS.length !== 30) throw new Error(`Expected 30 txn rows, got ${TXN_ROWS.length}`);
  const data = [
    { range: "'💰 Budget & Expense Tracker'!A1:G1", values: [TXN_HEADERS] },
    { range: "'💰 Budget & Expense Tracker'!A2:G31", values: TXN_ROWS },
    { range: "'💰 Budget & Expense Tracker'!H1:K1", values: [BILL_HEADERS] },
    { range: "'💰 Budget & Expense Tracker'!H2:K9", values: BILL_ROWS },
  ];
  await valuesBatchUpdate(id, data, 'bud-values');

  const reqs = [];
  reqs.push({ repeatCell: { range: gridRange(BUD,0,1,0,7), cell: { userEnteredFormat: { backgroundColor: hex(C.raspberry), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: BUD, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  for (let r = 1; r <= 30; r++) {
    reqs.push({ repeatCell: { range: gridRange(BUD,r,r+1,0,7), cell: { userEnteredFormat: { backgroundColor: hex(r%2===1?C.altRow:C.white), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  }
  reqs.push({ repeatCell: { range: gridRange(BUD,1,31,4,5), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  reqs.push({ repeatCell: { range: gridRange(BUD,0,1,7,11), cell: { userEnteredFormat: { backgroundColor: hex(C.mint), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  for (let r = 1; r <= 8; r++) {
    reqs.push({ repeatCell: { range: gridRange(BUD,r,r+1,7,11), cell: { userEnteredFormat: { backgroundColor: hex(r%2===1?C.altRow:C.white), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  }
  reqs.push({ repeatCell: { range: gridRange(BUD,1,9,8,9), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  [[0,100],[1,110],[2,200],[3,80],[4,90],[5,120],[6,180],[7,150],[8,90],[9,100],[10,60]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: BUD, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  await batchUpdate(id, reqs, 'bud-format');
  console.log('Budget & Expense Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
