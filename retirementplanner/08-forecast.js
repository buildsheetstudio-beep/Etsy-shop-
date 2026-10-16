'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Growth Forecast'];
const S = "'Growth Forecast'";
const SETUP = "'Personal & Household Setup'";
const ACCTS = "'Retirement Accounts'";
const INC   = "'Retirement Income'";

// Column layout (0-indexed):
// A(0)=Year  B(1)=Phase  C(2)=P1Age  D(3)=P2Age
// E(4)=BegBalance  F(5)=Contributions  G(6)=PortfolioReturn
// H(7)=ExternalIncome  I(8)=AnnualExpenses  J(9)=NetWithdrawal
// K(10)=EOYBalance  L(11)=CoverageRatio

const COL_WIDTHS = [72, 130, 72, 72, 140, 130, 130, 130, 130, 130, 150, 100];
const START_YEAR = 2026;
const END_YEAR   = 2065; // 40-year forecast
const YEARS = END_YEAR - START_YEAR + 1; // 40

// Retirement year = YEAR(P1 DOB) + P1 Retire Age = 1970+67 = 2037
// Life expectancy year for P1 = 1970+90 = 2060, P2 = 1972+90 = 2062

(async () => {
  const vals = [];
  const fmt  = [];

  COL_WIDTHS.forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  fmt.push({ repeatCell: { range: gridRange(SID,0,200,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // ===== TITLE =====
  vals.push({ range: `${S}!A1`, values: [['GROWTH FORECAST — 40-YEAR PROJECTION']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  vals.push({ range: `${S}!A2`, values: [['Annual portfolio projections using rates from Setup tab. Phase: Accumulation = contributions made; Distribution = withdrawals begin. All values in today\'s nominal dollars.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrB), textFormat: { fontSize: 9, foregroundColor: hex(C.primaryText), italic: true, fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // ===== SUMMARY CARDS (rows 3-4) =====
  const lastDataRow = 6 + YEARS; // last data row (1-indexed)
  const CARDS = [
    { label: 'Portfolio in 2037 (Retirement)', formula: `=IFERROR(VLOOKUP(2037,A7:K${lastDataRow},11,FALSE),"—")`, fmt: '"$"#,##0' },
    { label: 'Portfolio in 2050 (Mid-Retirement)', formula: `=IFERROR(VLOOKUP(2050,A7:K${lastDataRow},11,FALSE),"—")`, fmt: '"$"#,##0' },
    { label: 'Portfolio in 2060 (P1 @ 90)',   formula: `=IFERROR(VLOOKUP(2060,A7:K${lastDataRow},11,FALSE),"—")`, fmt: '"$"#,##0' },
    { label: 'Avg Annual External Income',     formula: `=IFERROR(AVERAGEIF(B7:B${lastDataRow},"Distribution",H7:H${lastDataRow}),"—")`, fmt: '"$"#,##0' },
    { label: 'Avg Annual Expenses (Ret.)',     formula: `=IFERROR(AVERAGEIF(B7:B${lastDataRow},"Distribution",I7:I${lastDataRow}),"—")`, fmt: '"$"#,##0' },
  ];
  const cardSpans = [[0,2],[2,4],[4,6],[6,8],[8,12]];
  CARDS.forEach((card, ci) => {
    const [c1,c2] = cardSpans[ci];
    const col = String.fromCharCode(65+c1);
    fmt.push({ mergeCells: { range: gridRange(SID,2,3,c1,c2), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,2,3,c1,c2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrB), textFormat: { bold: true, fontSize: 8, foregroundColor: hex('#D0D4E8'), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    vals.push({ range: `${S}!${col}3`, values: [[card.label]] });
    fmt.push({ mergeCells: { range: gridRange(SID,3,4,c1,c2), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,3,4,c1,c2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      numberFormat: { type: 'NUMBER', pattern: card.fmt },
      borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
    }}, fields: 'userEnteredFormat' }});
    vals.push({ range: `${S}!${col}4`, values: [[card.formula]] });
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // ===== COLUMN HEADERS (row 6, 0-idx 5) =====
  vals.push({ range: `${S}!A6`, values: [[
    'Year','Phase','P1\nAge','P2\nAge',
    'Beginning Balance','Annual Contributions','Portfolio Return ($)',
    'External Income','Annual Expenses','Net Withdrawal',
    'EOY Portfolio Balance','Coverage\nRatio',
  ]] });
  fmt.push({ repeatCell: { range: gridRange(SID,5,6,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
    borders: { bottom: { style: 'SOLID_MEDIUM', color: hex(C.primary) } },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 6 } }, fields: 'gridProperties.frozenRowCount' }});

  // ===== DATA ROWS =====
  const DATA_START = 6; // 0-indexed row of first data row
  const RETIRE_YEAR = 2037;
  const P1_BIRTH_YEAR = 1970;
  const P2_BIRTH_YEAR = 1972;

  // Formulas that are the same for all rows (referencing SETUP):
  const retireYrFn = `YEAR(${SETUP}!B8)+${SETUP}!B10`; // = 1970+67 = 2037

  // Phase formula
  const phaseFn = r => `=IF(A${r}>=(${retireYrFn}),"Distribution","Accumulation")`;

  // Contribution formula (only in accumulation)
  const contribFn = r => `=IFERROR(IF(B${r}="Accumulation",SUMIF(${ACCTS}!L7:L100,"Active",${ACCTS}!I7:I100),0),0)`;

  // Portfolio return: beginning balance × rate
  const returnFn = r => `=IFERROR(E${r}*${SETUP}!B25,0)`;

  // External income:
  // Accumulation: current Active income from income tab × inflation factor
  // Distribution: Active (ex. ending) + Projected income × inflation factor
  // Simplified approach: use base values and inflate from the transition year
  const incomePreRetBase  = `IFERROR(SUMIF(${INC}!L7:L200,"Active",${INC}!E7:E200),0)`;
  const incomePostRetBase = `IFERROR(SUMIF(${INC}!L7:L200,"Active",${INC}!E7:E200)+SUMIF(${INC}!L7:L200,"Projected",${INC}!E7:E200),0)`;
  const inflRate = `${SETUP}!B23`;

  const incomeFn = r => `=IFERROR(IF(B${r}="Distribution",(${incomePostRetBase})*(1+${inflRate})^(A${r}-(${retireYrFn})),(${incomePreRetBase})*(1+${inflRate})^(A${r}-${START_YEAR})),0)`;

  // Annual Expenses:
  // Pre-retirement: combined current expense from expense tab (for reference)
  // Post-retirement: Setup Annual Goal × inflation factor
  const annGoal = `${SETUP}!B30`;
  const expPreBase = `IFERROR(SUMIF('Retirement Expenses'!K7:K200,"Active",'Retirement Expenses'!F7:F200),0)`;
  const expensesFn = r => `=IFERROR(IF(B${r}="Distribution",${annGoal}*(1+${inflRate})^(A${r}-(${retireYrFn})),(${expPreBase})*(1+${inflRate})^(A${r}-${START_YEAR})),0)`;

  // Net withdrawal from portfolio: max(0, expenses - external income) in distribution
  const withdrawFn = r => `=IFERROR(IF(B${r}="Distribution",MAX(0,I${r}-H${r}),0),0)`;

  // EOY Balance = Beg + Contributions + Return - Withdrawal
  const eoyFn = r => `=IFERROR(E${r}+F${r}+G${r}-J${r},0)`;

  // Coverage ratio = External income / Expenses
  const coverageFn = r => `=IFERROR(H${r}/I${r},"")`;

  // Write all data rows in batches
  const dataRows = [];
  for (let yi = 0; yi < YEARS; yi++) {
    const yr = START_YEAR + yi;
    const rIdx = DATA_START + yi; // 0-indexed
    const r = rIdx + 1;           // 1-indexed
    const prevR = r - 1;

    // Beginning balance
    let begBalance;
    if (yi === 0) {
      begBalance = `=IFERROR(SUMIF(${ACCTS}!L7:L100,"Active",${ACCTS}!F7:F100),0)`;
    } else {
      begBalance = `=IFERROR(K${prevR},0)`;
    }

    dataRows.push({
      range: `${S}!A${r}`,
      values: [[
        yr,
        phaseFn(r),
        yr - P1_BIRTH_YEAR,
        yr - P2_BIRTH_YEAR,
        begBalance,
        contribFn(r),
        returnFn(r),
        incomeFn(r),
        expensesFn(r),
        withdrawFn(r),
        eoyFn(r),
        coverageFn(r),
      ]],
    });
  }
  // Batch the value writes (40 rows)
  for (let b = 0; b < dataRows.length; b += 40) {
    vals.push(...dataRows.slice(b, b+40));
  }

  // ===== FORMAT DATA ROWS =====
  // Year
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+YEARS, 0, 1), cell: { userEnteredFormat: {
    textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', backgroundColor: hex(C.altRow),
  }}, fields: 'userEnteredFormat' }});

  // Phase
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+YEARS, 1, 2), cell: { userEnteredFormat: {
    textFormat: { fontSize: 9, fontFamily: 'Arial' }, horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment' }});

  // Ages (C, D)
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+YEARS, 2, 4), cell: { userEnteredFormat: {
    textFormat: { fontSize: 9, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment' }});

  // Currency cols E, F, G, H, I, J (indices 4-9)
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+YEARS, 4, 11), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' },
    textFormat: { fontSize: 9, fontFamily: 'Arial' }, horizontalAlignment: 'RIGHT',
  }}, fields: 'userEnteredFormat' }});

  // Beginning balance (col E) — input bg
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+YEARS, 4, 5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // EOY Balance (col K) — bold, highlighted
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+YEARS, 10, 11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial' },
    numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' }, horizontalAlignment: 'RIGHT',
  }}, fields: 'userEnteredFormat' }});

  // Coverage ratio (col L) — percent
  fmt.push({ repeatCell: { range: gridRange(SID, DATA_START, DATA_START+YEARS, 11, 12), cell: { userEnteredFormat: {
    numberFormat: { type: 'PERCENT', pattern: '0%' },
    textFormat: { fontSize: 9, fontFamily: 'Arial' }, horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat' }});

  // Row heights
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: DATA_START, endIndex: DATA_START+YEARS },
    properties: { pixelSize: 20 }, fields: 'pixelSize' }});

  // Alternating row colors (every 5 years / every year)
  for (let yi = 0; yi < YEARS; yi++) {
    const rIdx = DATA_START + yi;
    const yr = START_YEAR + yi;
    const isRetire = yr === RETIRE_YEAR;
    const isDist   = yr >= RETIRE_YEAR;
    const isMilestone = yr % 5 === 0;
    const bgColor = isRetire ? C.warning :
                    isDist   ? (isMilestone ? '#E8F2EE' : C.panel) :
                               (isMilestone ? '#EEF0F8' : C.panel);

    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,0,12), cell: { userEnteredFormat: {
      backgroundColor: hex(bgColor),
    }}, fields: 'userEnteredFormat.backgroundColor' }});

    // Mark retirement year with special border
    if (isRetire) {
      fmt.push({ updateBorders: { range: gridRange(SID,rIdx,rIdx+1,0,12), top: { style: 'SOLID_MEDIUM', color: hex(C.warning) }}});
    }
    // Milestone years: bold year
    if (isMilestone || isRetire) {
      fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,0,1), cell: { userEnteredFormat: {
        textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primary) },
      }}, fields: 'userEnteredFormat.textFormat' }});
    }
  }

  // Conditional: EOY Balance > $3M → green; < $500K → red
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, DATA_START, DATA_START+YEARS, 10, 11)],
    booleanRule: {
      condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '3000000' }] },
      format: { textFormat: { foregroundColor: hex(C.success), bold: true } },
    },
  }, index: 0 }});
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, DATA_START, DATA_START+YEARS, 10, 11)],
    booleanRule: {
      condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '500000' }] },
      format: { textFormat: { foregroundColor: hex(C.attention), bold: true } },
    },
  }, index: 1 }});

  // ===== CHARTS =====
  const CHART_ANCHOR = (row, col, w, h) => ({
    overlayPosition: { anchorCell: { sheetId: SID, rowIndex: row, columnIndex: col }, widthPixels: w, heightPixels: h },
  });

  // Chart 1: Portfolio Balance Over Time (COLUMN chart, col K)
  const chart1 = { addChart: { chart: {
    spec: {
      title: 'Portfolio Balance Over 40 Years',
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'NO_LEGEND',
        headerCount: 1,
        domains: [{ domain: { sourceRange: { sources: [{
          sheetId: SID, startRowIndex: 5, endRowIndex: 5+YEARS+1, startColumnIndex: 0, endColumnIndex: 1,
        }] }}}],
        series: [{ series: { sourceRange: { sources: [{
          sheetId: SID, startRowIndex: 5, endRowIndex: 5+YEARS+1, startColumnIndex: 10, endColumnIndex: 11,
        }] }}, color: hex(C.secondary) }],
      },
    },
    position: CHART_ANCHOR(1, 12, 520, 300),
  }}};

  // Chart 2: External Income vs Annual Expenses (combo/column)
  const chart2 = { addChart: { chart: {
    spec: {
      title: 'External Income vs Annual Expenses (Retirement Phase)',
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'BOTTOM_LEGEND',
        headerCount: 1,
        domains: [{ domain: { sourceRange: { sources: [{
          sheetId: SID, startRowIndex: 5, endRowIndex: 5+YEARS+1, startColumnIndex: 0, endColumnIndex: 1,
        }] }}}],
        series: [
          { series: { sourceRange: { sources: [{
            sheetId: SID, startRowIndex: 5, endRowIndex: 5+YEARS+1, startColumnIndex: 7, endColumnIndex: 8,
          }] }}, color: hex(C.success) },
          { series: { sourceRange: { sources: [{
            sheetId: SID, startRowIndex: 5, endRowIndex: 5+YEARS+1, startColumnIndex: 8, endColumnIndex: 9,
          }] }}, color: hex(C.attention) },
        ],
      },
    },
    position: CHART_ANCHOR(1, 19, 520, 300),
  }}};

  // ===== DISCLAIMER =====
  const DISC_ROW = DATA_START + YEARS + 1; // 0-indexed
  vals.push({ range: `${S}!A${DISC_ROW+1}`, values: [[
    'DISCLAIMER: Projections are illustrative and based on simplified constant-rate assumptions. ' +
    'Actual investment returns, inflation, expenses, and income will vary annually. ' +
    'Tax implications, RMDs, Social Security adjustments, and market volatility are not modeled. ' +
    'Consult a licensed financial planner for personalized retirement planning.',
  ]] });
  fmt.push({ mergeCells: { range: gridRange(SID, DISC_ROW, DISC_ROW+1, 0, 12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID, DISC_ROW, DISC_ROW+1, 0, 12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow), textFormat: { fontSize: 8, foregroundColor: hex(C.secText), italic: true, fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: DISC_ROW, endIndex: DISC_ROW+1 },
    properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  await valuesBatchUpdate(id, vals, '08-forecast values');
  await batchUpdate(id, fmt, '08-forecast format');
  await batchUpdate(id, [chart1, chart2], '08-forecast charts');

  console.log(`✅ Growth Forecast done — ${YEARS} years (${START_YEAR}–${END_YEAR}) + 2 charts.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
