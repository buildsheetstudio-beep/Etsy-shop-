'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, colL, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Goals & Milestones'];
const S = "'Goals & Milestones'";
const SETUP = "'Fund Setup & Goals'";
const LOG   = "'Contribution Log'";

(async () => {
  const fmt = [];
  const vals = [];

  // ── Tab header ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 14), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A1`, values: [['GOALS & MILESTONES']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, 14), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 14), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A2`, values: [['Track milestone achievements and goal progress across all sinking funds']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, 14), cell: { userEnteredFormat: { backgroundColor: hex(C.lightGray), textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.textMid) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  // ── Column headers (row 5) ───────────────────────────────────────────────────
  const HEADERS = [
    'Milestone ID','Fund ID','Fund Name','Milestone Type','Milestone Name',
    'Target Amount ($)','Target %','Target Date','Current Balance ($)',
    '% of Milestone','Status','Date Achieved','Notes','Auto-Check',
  ];
  vals.push({ range: `${S}!A5`, values: [HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, 0, HEADERS.length), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white) }, wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });

  // ── Sample milestones data ───────────────────────────────────────────────────
  // 60+ milestones covering all 26 funds with varied milestone types
  // Columns: MilestoneID, FundID, FundName(formula), Type, Name, TargetAmt, TargetPct, TargetDate, CurrBal(formula), PctOfMilestone(formula), Status(formula), DateAchieved, Notes, AutoCheck(formula)

  const MILESTONES = [
    // Vacation Fund (FUND-001) - Goal $3,500
    ['MST-001','FUND-001','10% Funded','First Step: $350 saved','350','10','2024-03-01','','Vacation kick-off savings'],
    ['MST-002','FUND-001','25% Funded','Quarter Way: $875 saved','875','25','2024-06-01','',''],
    ['MST-003','FUND-001','50% Funded','Halfway: $1,750 saved','1750','50','2024-09-01','2024-10-15','Ahead of schedule!'],
    ['MST-004','FUND-001','75% Funded','Almost There: $2,625 saved','2625','75','2025-01-01','',''],
    ['MST-005','FUND-001','100% Goal Reached','Vacation Fund Complete','3500','100','2025-06-01','','Ready to book!'],
    // Emergency Fund (FUND-002) - Goal $8,000
    ['MST-006','FUND-002','10% Funded','Emergency Starter: $800','800','10','2024-02-01','2024-01-31','Early achiever'],
    ['MST-007','FUND-002','25% Funded','$2,000 Milestone','2000','25','2024-05-01','2024-04-28',''],
    ['MST-008','FUND-002','50% Funded','$4,000 Halfway Point','4000','50','2024-10-01','','On track'],
    ['MST-009','FUND-002','75% Funded','$6,000 — Almost Safe','6000','75','2025-03-01','',''],
    ['MST-010','FUND-002','100% Goal Reached','Full 3-Month Emergency Fund','8000','100','2025-09-01','',''],
    // Home Repair Fund (FUND-003) - Goal $5,000
    ['MST-011','FUND-003','25% Funded','Home Repair: $1,250 saved','1250','25','2024-04-01','',''],
    ['MST-012','FUND-003','50% Funded','Home Repair Halfway: $2,500','2500','50','2024-10-01','',''],
    ['MST-013','FUND-003','Custom Amount','First $1,000 for Roof Fund','1000','20','2024-03-01','2024-03-20','Roof repair priority'],
    ['MST-014','FUND-003','75% Funded','$3,750 — Major Repairs Covered','3750','75','2025-03-01','',''],
    // Car Maintenance (FUND-004) - Goal $2,500
    ['MST-015','FUND-004','25% Funded','Car: $625 saved','625','25','2024-03-01','2024-02-28',''],
    ['MST-016','FUND-004','50% Funded','Car: $1,250 halfway','1250','50','2024-07-01','2024-07-10',''],
    ['MST-017','FUND-004','100% Goal Reached','Car Fund Complete: $2,500','2500','100','2025-01-01','',''],
    // Christmas Fund (FUND-005) - Goal $1,200
    ['MST-018','FUND-005','10% Funded','Christmas: $120','120','10','2024-02-01','2024-01-20','Start early!'],
    ['MST-019','FUND-005','50% Funded','Christmas Halfway: $600','600','50','2024-07-01','2024-07-31','On track for December'],
    ['MST-020','FUND-005','100% Goal Reached','Christmas Fund Ready!','1200','100','2024-11-01','2024-11-05','Ahead of Christmas'],
    // Medical Fund (FUND-006) - Goal $3,000
    ['MST-021','FUND-006','25% Funded','Medical: $750 saved','750','25','2024-04-01','',''],
    ['MST-022','FUND-006','50% Funded','Medical Deductible Covered: $1,500','1500','50','2024-09-01','',''],
    ['MST-023','FUND-006','Custom Amount','Cover Annual Deductible: $1,400','1400','47','2024-06-01','2024-05-30','Deductible met'],
    // New Vehicle (FUND-007) - Goal $15,000
    ['MST-024','FUND-007','10% Funded','Vehicle Down Payment Start: $1,500','1500','10','2024-06-01','',''],
    ['MST-025','FUND-007','25% Funded','$3,750 Saved for Vehicle','3750','25','2025-06-01','',''],
    ['MST-026','FUND-007','Custom Amount','Enough for Used Car Purchase: $8,000','8000','53','2026-01-01','','Could buy reliable used car'],
    ['MST-027','FUND-007','50% Funded','Vehicle Fund Halfway: $7,500','7500','50','2026-06-01','',''],
    // Travel Fund (FUND-008) - Goal $4,000
    ['MST-028','FUND-008','25% Funded','Travel: $1,000 saved','1000','25','2024-05-01','2024-04-30',''],
    ['MST-029','FUND-008','50% Funded','Travel: $2,000 halfway','2000','50','2024-10-01','',''],
    ['MST-030','FUND-008','Target Date','Europe Trip Savings Goal','4000','100','2025-06-01','','Planning summer trip'],
    // Home Renovation (FUND-009) - Goal $12,000
    ['MST-031','FUND-009','10% Funded','Renovation Seed: $1,200','1200','10','2024-06-01','',''],
    ['MST-032','FUND-009','Custom Amount','Kitchen Refresh Budget: $3,500','3500','29','2025-01-01','','Phase 1: kitchen'],
    ['MST-033','FUND-009','25% Funded','Renovation: $3,000 saved','3000','25','2025-06-01','',''],
    // Dental Fund (FUND-010) - Goal $2,000
    ['MST-034','FUND-010','50% Funded','Dental: $1,000 halfway','1000','50','2024-08-01','',''],
    ['MST-035','FUND-010','100% Goal Reached','Dental Fund Complete: $2,000','2000','100','2025-02-01','',''],
    // Property Tax (FUND-011) - Goal $4,800
    ['MST-036','FUND-011','25% Funded','Tax: $1,200 saved','1200','25','2024-04-01','2024-03-31','Quarterly milestone'],
    ['MST-037','FUND-011','50% Funded','Tax: $2,400 halfway','2400','50','2024-07-01','2024-06-28',''],
    ['MST-038','FUND-011','75% Funded','Tax: $3,600 — Almost ready','3600','75','2024-10-01','',''],
    ['MST-039','FUND-011','100% Goal Reached','Property Tax Fully Funded','4800','100','2024-12-01','2024-11-20','Paid on time!'],
    // Furniture Fund (FUND-012) - Goal $3,500
    ['MST-040','FUND-012','25% Funded','Furniture: $875 saved','875','25','2024-06-01','',''],
    ['MST-041','FUND-012','Custom Amount','Couch Fund: $1,200','1200','34','2024-09-01','2024-08-15','Bought the couch!'],
    ['MST-042','FUND-012','50% Funded','Furniture: $1,750 halfway','1750','50','2025-01-01','',''],
    // Pets Fund (FUND-013) - Goal $1,800
    ['MST-043','FUND-013','50% Funded','Pet Fund: $900 halfway','900','50','2024-08-01','','Annual vet visit ready'],
    ['MST-044','FUND-013','100% Goal Reached','Pet Fund Complete: $1,800','1800','100','2025-03-01','','Emergency vet coverage'],
    // Technology (FUND-014) - Goal $2,200
    ['MST-045','FUND-014','25% Funded','Tech: $550 saved','550','25','2024-05-01','2024-04-20',''],
    ['MST-046','FUND-014','Custom Amount','Laptop Replacement Fund: $1,200','1200','55','2024-10-01','2024-09-30','New laptop purchased'],
    // Annual Subscriptions (FUND-015) - Goal $1,500
    ['MST-047','FUND-015','100% Goal Reached','Subscriptions Fully Funded','1500','100','2024-12-01','2024-11-28','All renewals covered'],
    // Wedding (FUND-016) - Goal $20,000
    ['MST-048','FUND-016','10% Funded','Wedding: $2,000 saved','2000','10','2024-09-01','','Starting to plan'],
    ['MST-049','FUND-016','Custom Percentage','25% of Wedding Budget','5000','25','2025-06-01','','Venue deposit possible'],
    ['MST-050','FUND-016','Custom Amount','Venue Deposit: $3,500','3500','18','2025-03-01','','Book venue early'],
    // Education (FUND-017) - Goal $6,000
    ['MST-051','FUND-017','25% Funded','Education: $1,500 saved','1500','25','2024-06-01','',''],
    ['MST-052','FUND-017','Custom Amount','First Semester Tuition: $2,800','2800','47','2024-08-01','2024-07-31','Ready for fall semester'],
    // Baby Fund (FUND-018) - Goal $5,000
    ['MST-053','FUND-018','25% Funded','Baby Fund: $1,250 saved','1250','25','2024-07-01','',''],
    ['MST-054','FUND-018','50% Funded','Baby Fund Halfway: $2,500','2500','50','2025-01-01','',''],
    // Holidays (FUND-019) - Goal $800
    ['MST-055','FUND-019','100% Goal Reached','Holiday Fund Complete','800','100','2024-11-01','2024-10-30',''],
    // Birthday Gifts (FUND-020) - Goal $600
    ['MST-056','FUND-020','50% Funded','Birthdays: $300 halfway','300','50','2024-06-01','2024-05-28',''],
    ['MST-057','FUND-020','100% Goal Reached','Birthday Fund Complete','600','100','2024-12-01','2024-11-15',''],
    // Moving Fund (FUND-021) - Goal $4,500
    ['MST-058','FUND-021','25% Funded','Moving: $1,125 saved','1125','25','2025-03-01','',''],
    ['MST-059','FUND-021','50% Funded','Moving: $2,250 halfway','2250','50','2025-09-01','',''],
    // Insurance (FUND-022) - Goal $2,400
    ['MST-060','FUND-022','50% Funded','Insurance: $1,200 halfway','1200','50','2024-09-01','2024-09-05',''],
    ['MST-061','FUND-022','100% Goal Reached','Insurance Fund Complete','2400','100','2025-03-01','',''],
    // Memberships (FUND-023) - Goal $900
    ['MST-062','FUND-023','100% Goal Reached','Memberships Funded','900','100','2024-12-01','2024-11-25','Gym + clubs covered'],
    // School Expenses (FUND-024) - Goal $1,800
    ['MST-063','FUND-024','50% Funded','School: $900 halfway','900','50','2024-08-01','2024-07-20',''],
    ['MST-064','FUND-024','100% Goal Reached','School Expenses Complete','1800','100','2025-01-01','',''],
    // Business Fund (FUND-025) - Goal $3,000
    ['MST-065','FUND-025','25% Funded','Business: $750 saved','750','25','2024-06-01','',''],
    ['MST-066','FUND-025','50% Funded','Business: $1,500 halfway','1500','50','2025-01-01','',''],
    ['MST-067','FUND-025','Custom Amount','Equipment Budget: $2,000','2000','67','2025-06-01','','New camera gear'],
    // Personal Fund (FUND-026) - Goal $1,000
    ['MST-068','FUND-026','50% Funded','Personal: $500 halfway','500','50','2024-07-01','2024-07-01',''],
    ['MST-069','FUND-026','100% Goal Reached','Personal Fund Complete','1000','100','2025-01-01','','Self-care fully funded'],
  ];

  const rowData = [];
  MILESTONES.forEach(([mid, fid, mtype, mname, targetAmt, targetPct, targetDate, dateAchieved, notes], i) => {
    const r = 6 + i;
    const fundNameF = `=IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$B$33,2,0),"")`;
    const currBalF  = `=IFERROR(VLOOKUP("${fid}",${SETUP}!$A$8:$I$33,9,0),0)`;
    const pctOfMilF = `=IFERROR(IF(F${r}=0,0,MIN(1,I${r}/F${r})),0)`;
    // Status: if DateAchieved filled → Achieved; elif currbal >= target → auto-Achieved; elif past due → Delayed; else In Progress / Not Started
    const statusF = `=IFERROR(IF(L${r}<>"","Achieved",IF(I${r}>=F${r},"Achieved",IF(AND(H${r}<>"",H${r}<TODAY(),I${r}<F${r}),"Delayed",IF(I${r}>0,"In Progress","Not Started")))),"")`;
    // Auto-check: TRUE if current balance >= target amount
    const autoF = `=IFERROR(I${r}>=F${r},FALSE)`;

    rowData.push([
      mid, fid, fundNameF, mtype, mname,
      parseFloat(targetAmt), parseFloat(targetPct)/100,
      targetDate || '',
      currBalF, pctOfMilF, statusF,
      dateAchieved || '', notes, autoF,
    ]);
  });

  vals.push({ range: `${S}!A6`, values: rowData });

  // ── Number formats ────────────────────────────────────────────────────────────
  // Currency: F, I (cols 5, 8)
  [5, 8].forEach(col => {
    fmt.push({ repeatCell: { range: gridRange(SID, 5, 5+MILESTONES.length, col, col+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });
  // Percent: G, J (cols 6, 9)
  [6, 9].forEach(col => {
    fmt.push({ repeatCell: { range: gridRange(SID, 5, 5+MILESTONES.length, col, col+1), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });
  // Date: H, L (cols 7, 11)
  [7, 11].forEach(col => {
    fmt.push({ repeatCell: { range: gridRange(SID, 5, 5+MILESTONES.length, col, col+1), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });
  });

  // ── Alternating row stripes ──────────────────────────────────────────────────
  for (let i = 0; i < MILESTONES.length; i++) {
    const bg = i % 2 === 0 ? C.white : C.stripeBg;
    fmt.push({ repeatCell: { range: gridRange(SID, 5+i, 6+i, 0, 14), cell: { userEnteredFormat: { backgroundColor: hex(bg) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }

  // ── Data validation: Status dropdown (col K = index 10) ─────────────────────
  const REF = "'Reference Data'";
  fmt.push({
    setDataValidation: {
      range: gridRange(SID, 5, 5+MILESTONES.length, 11, 12),
      rule: {
        condition: { type: 'ONE_OF_LIST', values: [
          { userEnteredValue: 'Not Started' }, { userEnteredValue: 'In Progress' },
          { userEnteredValue: 'Achieved' }, { userEnteredValue: 'Delayed' }, { userEnteredValue: 'Reassess' },
        ] },
        showCustomUi: true, strict: false,
      },
    },
  });

  // ── Column widths ────────────────────────────────────────────────────────────
  const colWidths = [80,80,150,130,180,90,75,100,100,85,100,100,160,70];
  colWidths.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // ── Row heights ──────────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });

  // ── Freeze ───────────────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Borders ──────────────────────────────────────────────────────────────────
  fmt.push({ updateBorders: { range: gridRange(SID, 4, 5+MILESTONES.length, 0, 14), innerHorizontal: { style: 'SOLID', color: hex(C.borderLight) }, innerVertical: { style: 'SOLID', color: hex(C.borderLight) }, bottom: { style: 'SOLID', color: hex(C.border) }, top: { style: 'SOLID', color: hex(C.border) }, left: { style: 'SOLID', color: hex(C.border) }, right: { style: 'SOLID', color: hex(C.border) } } });

  // ── Summary stats (rows 3-4) ─────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, 2), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A3`, values: [['MILESTONE SUMMARY']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  const N = MILESTONES.length;
  vals.push({ range: `${S}!A4`, values: [[
    'Total Milestones', `=${N}`, '',
    'Achieved', `=IFERROR(COUNTIF(K6:K${5+N},"Achieved"),0)`, '',
    'In Progress', `=IFERROR(COUNTIF(K6:K${5+N},"In Progress"),0)`, '',
    'Delayed', `=IFERROR(COUNTIF(K6:K${5+N},"Delayed"),0)`, '',
    'Not Started', `=IFERROR(COUNTIF(K6:K${5+N},"Not Started"),0)`, '',
    '% Complete', `=IFERROR(COUNTIF(K6:K${5+N},"Achieved")/${N},0)`,
  ]] });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 0, 14), cell: { userEnteredFormat: { backgroundColor: hex(C.lightGray), textFormat: { bold: false, fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  // Bold the label cells
  [0,3,6,9,12].forEach(col => {
    fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, col, col+1), cell: { userEnteredFormat: { textFormat: { bold: true } } }, fields: 'userEnteredFormat.textFormat' } });
  });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 13, 14), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

  await batchUpdate(id, fmt, 'milestones-fmt');
  await valuesBatchUpdate(id, vals, 'milestones-vals');

  console.log('✓ Goals & Milestones complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
