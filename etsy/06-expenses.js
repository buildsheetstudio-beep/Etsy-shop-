'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C,
        EXPENSE_CATS, EXPENSE_STATUSES, TAX_GROUP_MAP, MONTHS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SH = sheetMap['Expense Log'];
const S  = "'Expense Log'";

// Columns:
// A=Date  B=Year(f)  C=Month(f)  D=Quarter(f)  E=Category  F=Sub-Category
// G=Description  H=Amount  I=Tax Deductible?  J=Tax Group(f)  K=Status
// L=Payment Method  M=Vendor  N=Receipt?  O=Notes

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 1, endIndex: 5 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 5, endIndex: 500 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // Column widths
  const colW = [100,50,90,55,160,130,210,100,90,150,110,120,140,80,180];
  colW.forEach((w,i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title A1:O1
  reqs.push({ mergeCells: { range: gridRange(SH, 0, 1, 0, 15), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 0, 1, 0, 15), cell: { userEnteredFormat: {
    backgroundColor: hex(C.attention),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 17 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Summary cards rows 2-4
  [[0,5],[5,10],[10,15]].forEach(([c0,c1], idx) => {
    const colors = [C.attention, C.warning, C.success];
    reqs.push({ mergeCells: { range: gridRange(SH, 1, 2, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 1, 2, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(colors[idx]),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    reqs.push({ mergeCells: { range: gridRange(SH, 2, 3, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 2, 3, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { foregroundColor: hex(colors[idx]), bold: true, fontSize: 20 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    reqs.push({ mergeCells: { range: gridRange(SH, 3, 4, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, 3, 4, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.altRow),
      textFormat: { foregroundColor: hex(C.secText), fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  });

  // Header row (index 4)
  reqs.push({ repeatCell: { range: gridRange(SH, 4, 5, 0, 15), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' }});

  // Data rows
  for (let r = 5; r < 500; r++) {
    const bg = r % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 15), cell: { userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { foregroundColor: hex(C.mainText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
  }

  // Formula cols B,C,D,J (1,2,3,9) — formula bg
  [1,2,3,9].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 5, 500, c, c+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' }});
  });

  // Date format A
  reqs.push({ repeatCell: { range: gridRange(SH, 5, 500, 0, 1), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } } }, fields: 'userEnteredFormat.numberFormat' }});

  // Currency H (7)
  reqs.push({ repeatCell: { range: gridRange(SH, 5, 500, 7, 8), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
  // Summary card currency
  reqs.push({ repeatCell: { range: gridRange(SH, 2, 3, 0, 15), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});

  // Center: A,B,C,D,I,K,N (0,1,2,3,8,10,13)
  [0,1,2,3,8,10,13].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(SH, 4, 500, c, c+1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat.horizontalAlignment' }});
  });

  // Borders
  reqs.push({ updateBorders: { range: gridRange(SH, 4, 500, 0, 15),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Freeze row 5
  reqs.push({ updateSheetProperties: { properties: { sheetId: SH, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'expense-format');

  // Values
  const data = [];
  data.push({ range: `${S}!A1`, values: [['💸 Expense Log — All Business Expenses']] });

  // Summary cards
  data.push({ range: `${S}!A2`, values: [['  Total Expenses (2026)']] });
  data.push({ range: `${S}!A3`, values: [['=SUMPRODUCT((YEAR(A$6:A$500)=2026)*H$6:H$500)']] });
  data.push({ range: `${S}!A4`, values: [['All categories · fiscal year to date']] });

  data.push({ range: `${S}!F2`, values: [['  Tax Deductible Expenses (2026)']] });
  data.push({ range: `${S}!F3`, values: [['=SUMPRODUCT((YEAR(A$6:A$500)=2026)*(I$6:I$500="Yes")*H$6:H$500)']] });
  data.push({ range: `${S}!F4`, values: [['Schedule C deductible amount']] });

  data.push({ range: `${S}!K2`, values: [['  Etsy Platform Fees (2026)']] });
  data.push({ range: `${S}!K3`, values: [['=SUMPRODUCT((YEAR(A$6:A$500)=2026)*(ISNUMBER(MATCH(E$6:E$500,{"Etsy Transaction Fees","Etsy Listing Fees","Etsy Payment Processing Fees","Etsy Ads","Offsite Ads"},0)))*H$6:H$500)']] });
  data.push({ range: `${S}!K4`, values: [['Etsy fees only']] });

  // Headers
  data.push({ range: `${S}!A5:O5`, values: [['Date','Year','Month','Quarter','Category','Sub-Category','Description','Amount','Tax Deductible?','Tax Group','Status','Payment Method','Vendor','Receipt?','Notes']] });

  // 60 sample expense rows across Jan-Jul 2026
  const sampleExpenses = [
    // Etsy Fees (auto-tracked)
    ['2026-01-03','Etsy Listing Fees','New Listings','Listing fees for 25 new items',5.00,'Yes','Etsy Platform Fees','Paid','Etsy','Yes','25 × $0.20'],
    ['2026-01-10','Etsy Transaction Fees','Jan Sales Fees','Transaction fees Q1 Jan',45.23,'Yes','Etsy Platform Fees','Paid','Etsy','Yes','6.5% of sales'],
    ['2026-01-10','Etsy Payment Processing Fees','Jan Processing','Payment processing fees Jan',22.17,'Yes','Etsy Platform Fees','Paid','Etsy','Yes','3%+$0.25 per order'],
    ['2026-01-15','Etsy Ads','Jan Ads','Etsy Ads campaign Jan',50.00,'Yes','Etsy Platform Fees','Paid','Etsy','Yes','$50 monthly ad budget'],
    ['2026-01-20','Software & Subscriptions','Canva Pro','Canva Pro monthly',14.99,'Yes','Software','Paid','Canva','Yes','Design tool'],
    ['2026-01-22','Software & Subscriptions','Adobe CC','Adobe Creative Cloud',59.99,'Yes','Software','Paid','Adobe','Yes','Monthly plan'],
    ['2026-01-25','Design Assets','Font Pack','Procreate font bundle',18.00,'Yes','Design Tools','Paid','Creative Market','Yes','Commercial license'],
    ['2026-01-28','Office Supplies','Ink Cartridges','HP ink for proofing',38.50,'Yes','Office','Paid','Staples','Yes',''],
    ['2026-02-02','Etsy Listing Fees','Feb Listings','Listing renewal fees Feb',4.40,'Yes','Etsy Platform Fees','Paid','Etsy','Yes','22 renewals'],
    ['2026-02-05','Etsy Transaction Fees','Feb Sales Fees','Transaction fees Feb',52.80,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-02-05','Etsy Payment Processing Fees','Feb Processing','Payment processing Feb',26.12,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-02-10','Etsy Ads','Feb Ads','Etsy Ads Feb',75.00,'Yes','Etsy Platform Fees','Paid','Etsy','Yes','Increased budget'],
    ['2026-02-14','Design Assets','Valentines Graphics','Valentine clipart bundle',24.00,'Yes','Design Tools','Paid','Etsy','Yes','Commercial use'],
    ['2026-02-18','Education & Training','Skillshare','Skillshare annual plan',99.00,'Yes','Education','Paid','Skillshare','Yes','Annual'],
    ['2026-02-22','Photography','Product Photos','Product photography session',120.00,'Yes','Photography','Paid','Studio Local','Yes','10 product images'],
    ['2026-02-25','Banking Fees','PayPal Fee','PayPal transfer fee',1.50,'Yes','Banking','Paid','PayPal','No',''],
    ['2026-03-01','Etsy Listing Fees','Mar Listings','Listing fees Mar',3.60,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-03-03','Etsy Transaction Fees','Mar Sales Fees','Transaction fees Mar',61.45,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-03-03','Etsy Payment Processing Fees','Mar Processing','Processing fees Mar',30.22,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-03-05','Etsy Ads','Mar Ads','Etsy Ads Mar',75.00,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-03-10','Software & Subscriptions','Planoly','Instagram scheduler',18.00,'Yes','Software','Paid','Planoly','Yes',''],
    ['2026-03-15','Contractors','VA Support','Virtual assistant 5hrs',75.00,'Yes','VA & Contractors','Paid','Freelancer','Yes','Order management'],
    ['2026-03-20','Print-on-Demand Costs','POD Jan-Mar','Printify production costs Q1',245.50,'Yes','Production','Paid','Printify','Yes','20 POD orders'],
    ['2026-03-28','Office Supplies','Paper & Supplies','Printing supplies',22.00,'Yes','Office','Paid','Staples','Yes',''],
    ['2026-04-02','Etsy Listing Fees','Apr Listings','Listing fees Apr',4.00,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-04-04','Etsy Transaction Fees','Apr Sales Fees','Transaction fees Apr',68.90,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-04-04','Etsy Payment Processing Fees','Apr Processing','Processing fees Apr',34.10,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-04-07','Etsy Ads','Apr Ads','Etsy Ads April',100.00,'Yes','Etsy Platform Fees','Paid','Etsy','Yes','Increased for spring'],
    ['2026-04-10','Offsite Ads','Offsite Fee','Offsite ads fee Apr',38.22,'Yes','Etsy Platform Fees','Paid','Etsy','Yes','15% on 3 orders'],
    ['2026-04-15','Design Assets','Spring Clipart','Spring/Easter graphics pack',32.00,'Yes','Design Tools','Paid','Creative Market','Yes','Commercial license'],
    ['2026-04-18','Software & Subscriptions','Canva Pro','Canva Pro Apr',14.99,'Yes','Software','Paid','Canva','Yes',''],
    ['2026-04-18','Software & Subscriptions','Adobe CC','Adobe Creative Cloud Apr',59.99,'Yes','Software','Paid','Adobe','Yes',''],
    ['2026-04-22','Contractors','Copywriter','Product description copy 10 items',80.00,'Yes','VA & Contractors','Paid','Freelancer','Yes',''],
    ['2026-04-28','Internet & Phone','Internet','Monthly internet bill',65.00,'Yes','Home Office','Paid','ISP','Yes','Business portion 60%'],
    ['2026-05-01','Etsy Listing Fees','May Listings','Listing fees May',5.00,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-05-03','Etsy Transaction Fees','May Sales Fees','Transaction fees May',74.20,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-05-03','Etsy Payment Processing Fees','May Processing','Processing fees May',37.50,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-05-05','Etsy Ads','May Ads','Etsy Ads May',100.00,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-05-10','Print-on-Demand Costs','POD Apr-May','Printify production costs',189.30,'Yes','Production','Paid','Printify','Yes','15 POD orders'],
    ['2026-05-15','Design Assets','Summer Graphics','Summer design bundle',28.00,'Yes','Design Tools','Paid','Etsy','Yes','Commercial license'],
    ['2026-05-20','Photography','Social Photos','Lifestyle photography',90.00,'Yes','Photography','Paid','Freelancer','Yes','Instagram content'],
    ['2026-05-28','Professional Fees','Accountant','Quarterly bookkeeping review',150.00,'Yes','Professional','Paid','CPA Services','Yes','Q1 review'],
    ['2026-06-02','Etsy Listing Fees','Jun Listings','Listing fees June',3.80,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-06-04','Etsy Transaction Fees','Jun Sales Fees','Transaction fees June',80.55,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-06-04','Etsy Payment Processing Fees','Jun Processing','Processing fees June',39.80,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-06-07','Etsy Ads','Jun Ads','Etsy Ads June',100.00,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-06-10','Software & Subscriptions','Canva Pro','Canva Pro Jun',14.99,'Yes','Software','Paid','Canva','Yes',''],
    ['2026-06-10','Software & Subscriptions','Adobe CC','Adobe Creative Cloud Jun',59.99,'Yes','Software','Paid','Adobe','Yes',''],
    ['2026-06-15','Design Assets','Fall Preview Pack','Early fall design assets',45.00,'Yes','Design Tools','Paid','Creative Market','Yes','Commercial license'],
    ['2026-06-22','Contractors','VA Support','VA June 8hrs',120.00,'Yes','VA & Contractors','Paid','Freelancer','Yes','Customer service'],
    ['2026-06-28','Banking Fees','Stripe Fee','Payment processing fee',3.20,'Yes','Banking','Paid','Stripe','No',''],
    ['2026-07-01','Etsy Listing Fees','Jul Listings','Listing fees July',4.20,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-07-03','Etsy Transaction Fees','Jul Sales Fees','Transaction fees July',85.10,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-07-03','Etsy Payment Processing Fees','Jul Processing','Processing fees July',42.20,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-07-05','Etsy Ads','Jul Ads','Etsy Ads July',100.00,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-07-08','Offsite Ads','Offsite Jul','Offsite ads fee July',52.40,'Yes','Etsy Platform Fees','Paid','Etsy','Yes',''],
    ['2026-07-10','Print-on-Demand Costs','POD Jun-Jul','Printify production Q2-Q3',220.80,'Yes','Production','Paid','Printify','Yes','18 POD orders'],
    ['2026-07-15','Education & Training','Course Purchase','Etsy SEO masterclass',49.00,'Yes','Education','Paid','Teachable','Yes','Online course'],
    ['2026-07-18','Software & Subscriptions','Everbee','Everbee research tool',19.99,'Yes','Software','Paid','Everbee','Yes','Keyword research'],
    ['2026-07-20','Design Assets','Back to School','Back to school clipart',22.00,'Yes','Design Tools','Paid','Etsy','Yes',''],
    ['2026-07-22','Taxes Paid','Q2 Estimated','Q2 estimated tax payment',1200.00,'Yes','Taxes','Paid','IRS','Yes','Form 1040-ES'],
  ];

  const expenseRows = sampleExpenses.map(([date, cat, subcat, desc, amount, taxDed, taxGroup, status, vendor, receipt, notes]) => [
    date, '', '', '', cat, subcat, desc, amount, taxDed, taxGroup, status, 'Bank Transfer', vendor, receipt, notes,
  ]);
  data.push({ range: `${S}!A6:O${5+expenseRows.length}`, values: expenseRows });

  // Formula columns for all data rows
  for (let r = 6; r <= 5 + sampleExpenses.length + 10; r++) {
    data.push({ range: `${S}!B${r}`, values: [[`=IFERROR(YEAR(A${r}),"")`]] });
    data.push({ range: `${S}!C${r}`, values: [[`=IFERROR(TEXT(A${r},"mmmm"),"")`]] });
    data.push({ range: `${S}!D${r}`, values: [[`=IFERROR(IF(MONTH(A${r})<=3,"Q1",IF(MONTH(A${r})<=6,"Q2",IF(MONTH(A${r})<=9,"Q3","Q4"))),"")`]] });
    data.push({ range: `${S}!J${r}`, values: [[`=IFERROR(VLOOKUP(E${r},'Reference Data'!$N$2:$O${1+EXPENSE_CATS.length},2,FALSE),"")`]] });
  }

  await valuesBatchUpdate(id, data, 'expense-values');

  // Validations
  const dvReqs = [];
  dvReqs.push({ setDataValidation: {
    range: gridRange(SH, 5, 500, 4, 5),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$C$2:$C${1+EXPENSE_CATS.length}` }] }, showCustomUi: true, strict: false },
  }});
  dvReqs.push({ setDataValidation: {
    range: gridRange(SH, 5, 500, 8, 9),
    rule: { condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'Yes' },{ userEnteredValue: 'No' }] }, showCustomUi: true, strict: true },
  }});
  dvReqs.push({ setDataValidation: {
    range: gridRange(SH, 5, 500, 10, 11),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$G$2:$G${1+EXPENSE_STATUSES.length}` }] }, showCustomUi: true, strict: true },
  }});
  dvReqs.push({ setDataValidation: {
    range: gridRange(SH, 5, 500, 13, 14),
    rule: { condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'Yes' },{ userEnteredValue: 'No' }] }, showCustomUi: true, strict: true },
  }});
  await batchUpdate(id, dvReqs, 'expense-validation');

  console.log('Expense Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
