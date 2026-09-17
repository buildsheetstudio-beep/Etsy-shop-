'use strict';
const { batchUpdate, valuesBatchUpdate, hex, COLOR, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = sheetMap['Student Budget']; // 4

// Budget entries: [Category, Item, Budgeted, Actual, Frequency, Notes]
const INCOME = [
  ['Part-time Job', 'Campus Bookstore', 800, 750, 'Monthly', 'Weekends only'],
  ['Scholarship', 'Merit Scholarship', 1000, 1000, 'Monthly', 'Fall semester disbursement'],
  ['Parents', 'Family Support', 500, 500, 'Monthly', 'Monthly transfer'],
  ['Financial Aid', 'FAFSA Grant', 600, 600, 'Monthly', 'Per semester / ~2mo'],
  ['Freelance', 'Tutoring Income', 150, 200, 'Monthly', 'Other students'],
];

const EXPENSES = [
  ['Housing', 'Dorm Room', 900, 920, 'Monthly', 'With meal plan discount'],
  ['Food', 'Dining Hall Meal Plan', 250, 280, 'Monthly', 'Unlimited swipes'],
  ['Food', 'Groceries', 100, 115, 'Monthly', 'Off-campus extras'],
  ['Food', 'Eating Out', 80, 125, 'Monthly', 'Weekends & social'],
  ['Transportation', 'Bus Pass', 40, 40, 'Monthly', 'City transit'],
  ['Transportation', 'Rideshare/Uber', 30, 55, 'Monthly', 'Late nights'],
  ['Books & Supplies', 'Textbooks', 120, 95, 'Monthly', 'Used + rented'],
  ['Books & Supplies', 'Lab Supplies', 40, 45, 'Monthly', 'Bio lab kit'],
  ['Entertainment', 'Streaming Services', 25, 25, 'Monthly', 'Netflix + Spotify'],
  ['Entertainment', 'Social Activities', 60, 85, 'Monthly', 'Concerts, events'],
  ['Health', 'Campus Health Fee', 50, 50, 'Monthly', 'Included in tuition'],
  ['Technology', 'Phone Bill', 45, 45, 'Monthly', 'Family plan share'],
  ['Personal Care', 'Toiletries & Misc', 35, 42, 'Monthly', 'Personal needs'],
];

const SAVINGS = [
  ['Savings', 'Emergency Fund', 100, 50, 'Monthly', 'Building to $500'],
  ['Savings', 'Spring Semester Fund', 150, 100, 'Monthly', 'Saving ahead'],
];

(async () => {
  const reqs = [];

  // ── Column widths ──
  const widths = [[0,160],[1,180],[2,100],[3,100],[4,90],[5,200]];
  widths.forEach(([ci,px]) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    },
  }));

  // ── Row heights ──
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 4 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });

  // ── Title row ──
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 6), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 6),
      cell: {
        userEnteredValue: { stringValue: '💰 Student Budget Tracker' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.gold),
          textFormat: { foregroundColor: hex(COLOR.nearBlack), bold: true, fontSize: 18 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Subtitle ──
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 6), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, 6),
      cell: {
        userEnteredValue: { stringValue: 'Fall 2025 Monthly Budget | Alex Rivera' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.paleGold),
          textFormat: { foregroundColor: hex(COLOR.amber), bold: false, italic: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── KPI summary row (row 2) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, 2), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 2, 4), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 4, 6), mergeType: 'MERGE_ALL' } });

  const kpiStyles = [
    { c0: 0, bg: COLOR.paleMint, fg: COLOR.emerald },
    { c0: 2, bg: COLOR.paleRed, fg: COLOR.alertRed },
    { c0: 4, bg: COLOR.paleSkyBlue, fg: COLOR.darkBlue },
  ];
  kpiStyles.forEach(({ c0, bg, fg }) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 2, 3, c0, c0 + 2),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(fg), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
  });

  // Helper function for section headers
  function sectionHeader(r, label, bg, fg) {
    reqs.push({ mergeCells: { range: gridRange(SID, r, r + 1, 0, 6), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 0, 6),
        cell: {
          userEnteredValue: { stringValue: label },
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(fg), bold: true, fontSize: 11 },
            horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
            padding: { left: 10 },
          },
        },
        fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r + 1 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  }

  function columnHeaders(r, bg, fg) {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 0, 6),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(fg), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
  }

  function dataRows(r0, r1, altBg) {
    for (let r = r0; r < r1; r++) {
      if ((r - r0) % 2 === 0) {
        reqs.push({
          repeatCell: {
            range: gridRange(SID, r, r + 1, 0, 6),
            cell: { userEnteredFormat: { backgroundColor: hex(altBg) } },
            fields: 'userEnteredFormat.backgroundColor',
          },
        });
      }
    }
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r0, r1, 0, 6),
        cell: {
          userEnteredFormat: {
            textFormat: { fontSize: 10 }, verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(textFormat,verticalAlignment)',
      },
    });
    [2,3].forEach(ci => {
      reqs.push({
        repeatCell: {
          range: gridRange(SID, r0, r1, ci, ci + 1),
          cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } },
          fields: 'userEnteredFormat(horizontalAlignment,numberFormat)',
        },
      });
    });
  }

  function totalRow(r, bg, fg) {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 0, 6),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(fg), bold: true, fontSize: 11 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    [2,3].forEach(ci => {
      reqs.push({
        repeatCell: {
          range: gridRange(SID, r, r + 1, ci, ci + 1),
          cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } },
          fields: 'userEnteredFormat.numberFormat',
        },
      });
    });
  }

  // ── INCOME section rows 3-14 ──
  // Row 3: col headers, Row 4-8: income, Row 9: section label (blank), Row 10: income total
  sectionHeader(3, '💵 INCOME', COLOR.paleMint, COLOR.emerald);
  columnHeaders(4, COLOR.emerald, COLOR.white);
  dataRows(5, 10, COLOR.altRow);
  totalRow(10, COLOR.paleMint, COLOR.emerald);

  // ── EXPENSES section rows 11-28 ──
  sectionHeader(11, '💸 EXPENSES', COLOR.paleRed, COLOR.alertRed);
  columnHeaders(12, COLOR.alertRed, COLOR.white);
  dataRows(13, 26, COLOR.altRow);
  totalRow(26, COLOR.paleRed, COLOR.alertRed);

  // ── SAVINGS section rows 27-33 ──
  sectionHeader(27, '🏦 SAVINGS', COLOR.paleSkyBlue, COLOR.darkBlue);
  columnHeaders(28, COLOR.darkBlue, COLOR.white);
  dataRows(29, 31, COLOR.altRow);
  totalRow(31, COLOR.paleSkyBlue, COLOR.darkBlue);

  // ── NET BUDGET row (row 32, index 31→ already used; let's use 32) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 32, 33, 0, 2), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 32, 33, 2, 4), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 32, 33, 4, 6), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 32, 33, 0, 6),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(COLOR.skyBlue),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  [2,4].forEach(ci => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 32, 33, ci, ci + 2),
        cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } },
        fields: 'userEnteredFormat.numberFormat',
      },
    });
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 32, endIndex: 33 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // ── Borders ──
  reqs.push({ updateBorders: { range: gridRange(SID, 3, 33, 0, 6), innerHorizontal: { style: 'SOLID', color: hex(COLOR.softGrey) }, innerVertical: { style: 'SOLID', color: hex(COLOR.midGrey) } } });

  await batchUpdate(id, reqs, 'budget-format');

  // ── Values ──
  const colHdr = ['Category', 'Item', 'Budgeted', 'Actual', 'Frequency', 'Notes'];

  // Income rows start at row 5 (index 4), data at 5-9 (5 items), total at row 11 (index 10)
  // Expenses rows 13-25 (13 items), total at 27 (index 26)
  // Savings rows 29-30 (2 items), total at 32 (index 31)
  // Net budget row 33 (index 32)

  await valuesBatchUpdate(id, [
    // KPI
    { range: 'Student Budget!A3', values: [['="""Total Income: $"""&TEXT(SUM(C6:C10),"#,##0.00")']] },
    { range: 'Student Budget!C3', values: [['="""Total Expenses: $"""&TEXT(SUM(C14:C26),"#,##0.00")']] },
    { range: 'Student Budget!E3', values: [['="""Net Balance: $"""&TEXT(C11-C27-C32,"#,##0.00")']] },
    // Income
    { range: 'Student Budget!A5:F5', values: [colHdr] },
    { range: 'Student Budget!A6:F10', values: INCOME },
    { range: 'Student Budget!A11:F11', values: [['TOTAL INCOME','','=SUM(C6:C10)','=SUM(D6:D10)','','=IF(D11>C11,"Over budget","On track")']] },
    // Expenses
    { range: 'Student Budget!A13:F13', values: [colHdr] },
    { range: 'Student Budget!A14:F26', values: EXPENSES },
    { range: 'Student Budget!A27:F27', values: [['TOTAL EXPENSES','','=SUM(C14:C26)','=SUM(D14:D26)','','=IF(D27>C27,"Over budget","On track")']] },
    // Savings
    { range: 'Student Budget!A29:F29', values: [colHdr] },
    { range: 'Student Budget!A30:F31', values: SAVINGS },
    { range: 'Student Budget!A32:F32', values: [['TOTAL SAVINGS','','=SUM(C30:C31)','=SUM(D30:D31)','','=IF(D32>C32,"On track","Behind")']] },
    // Net budget row
    { range: 'Student Budget!A33:F33', values: [['NET BUDGET','','=C11-C27-C32','=D11-D27-D32','','=IF(C33>=0,"Balanced","In Deficit")']] },
    // D34 = C36 override (Dashboard cross-reference cell) — actual leftover
    // Row 34 = index 33; use C33 (=net actual) for the dashboard reference
    // D34 cell = actual net balance (D11-D27-D32)
    { range: 'Student Budget!D34', values: [['=C33']] },
  ], 'budget-values');

  console.log('Student Budget tab complete');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
