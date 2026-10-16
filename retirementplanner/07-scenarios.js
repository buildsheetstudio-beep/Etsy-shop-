'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Scenario Planner'];
const S = "'Scenario Planner'";
const SETUP = "'Personal & Household Setup'";
const ACCTS = "'Retirement Accounts'";

// Column layout (0-indexed), A-N = 14 cols:
// A(0)=ScnID  B(1)=Name  C(2)=Description
// D(3)=P1RetireAge  E(4)=P2RetireAge  F(5)=Inflation%  G(6)=ExpReturn%
// H(7)=SpendingAdj%  I(8)=LifeExpYrs
// J(9)=P1YrsToRetire  K(10)=EstPortfolio  L(11)=AnnualSWR
// M(12)=AnnualGoalAdj  N(13)=AnnualGap

const COL_WIDTHS = [80, 175, 210, 90, 90, 90, 90, 90, 90, 100, 150, 130, 140, 130];

// Scenario colors from lib
const SCN_COLORS = [C.SCN01, C.SCN02, C.SCN03, C.SCN04, C.SCN05, C.SCN06];

// 6 scenarios
const SCENARIOS = [
  { id:'SCN-01', name:'Baseline',           desc:'Current plan — no changes from Setup assumptions',         p1Age:67, p2Age:65, infl:0.03,  ret:0.07,  spend:0.00,  lifeExp:90 },
  { id:'SCN-02', name:'Earlier Retirement', desc:'Both retire 2 years sooner than baseline plan',            p1Age:65, p2Age:63, infl:0.03,  ret:0.07,  spend:0.00,  lifeExp:90 },
  { id:'SCN-03', name:'Later Retirement',   desc:'Both retire 2 years later — larger nest egg',              p1Age:69, p2Age:67, infl:0.03,  ret:0.07,  spend:0.00,  lifeExp:90 },
  { id:'SCN-04', name:'Higher Inflation',   desc:'Inflation rises to 5% — purchasing power erodes faster',  p1Age:67, p2Age:65, infl:0.05,  ret:0.07,  spend:0.00,  lifeExp:90 },
  { id:'SCN-05', name:'Lower Returns',      desc:'Portfolio earns 2% less than expected (bear market)',      p1Age:67, p2Age:65, infl:0.03,  ret:0.05,  spend:0.00,  lifeExp:90 },
  { id:'SCN-06', name:'Reduced Spending',   desc:'Retirement expenses cut 15% — core lifestyle maintained', p1Age:67, p2Age:65, infl:0.03,  ret:0.07,  spend:-0.15, lifeExp:90 },
];

// FV formula for estimated portfolio using scenario params at row r
// Uses accounts tab totals + scenario's return & retirement age
function portfolioFV(r) {
  const totBal = `IFERROR(SUMIF(${ACCTS}!L7:L100,"Active",${ACCTS}!F7:F100),0)`;
  const totCon = `IFERROR(SUMIF(${ACCTS}!L7:L100,"Active",${ACCTS}!I7:I100),0)`;
  const p1DOB  = `${SETUP}!B8`;
  const ytr    = `MAX(0,D${r}-IFERROR(DATEDIF(${p1DOB},TODAY(),"Y"),0))`;
  return `=IFERROR(IF(G${r}=0,(${totBal})+(${totCon})*${ytr},(${totBal})*(1+G${r})^(${ytr})+(${totCon})*((1+G${r})^(${ytr})-1)/G${r}),"")`;
}

(async () => {
  const vals = [];
  const fmt  = [];

  COL_WIDTHS.forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  fmt.push({ repeatCell: { range: gridRange(SID,0,400,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // ===== TITLE =====
  vals.push({ range: `${S}!A1`, values: [['SCENARIO PLANNER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  vals.push({ range: `${S}!A2`, values: [['Compare six "what-if" scenarios. Edit yellow cells to model different retirement timelines, returns, and spending. Columns J-N are calculated.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrB), textFormat: { fontSize: 9, foregroundColor: hex(C.primaryText), italic: true, fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // ===== BASELINE REFERENCE STRIP (row 3) =====
  vals.push({ range: `${S}!A3`, values: [[
    'Baseline from Setup →',
    `Retire Ages: P1=${`'${SETUP}'`}!B10, P2=${`'${SETUP}'`}!E10`,
    '', '',
    `=${SETUP}!B10`, `=${SETUP}!E10`,
    `=${SETUP}!B23`, `=${SETUP}!B25`, '', `=${SETUP}!B26`,
    '','','','',
  ]] });
  // Actually just show as static reference strip without formulas to avoid display issues
  // Use a label strip instead
  vals.push({ range: `${S}!A3`, values: [[
    'BASELINE REFERENCE (from Setup tab — read-only)',
    '','','',
    `=IFERROR(${SETUP}!B10,67)`,
    `=IFERROR(${SETUP}!E10,65)`,
    `=IFERROR(${SETUP}!B23,0.03)`,
    `=IFERROR(${SETUP}!B25,0.07)`,
    '0%',
    `=IFERROR(${SETUP}!B26,90)`,
    '','','','',
  ]] });
  fmt.push({ mergeCells: { range: gridRange(SID,2,3,0,4), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow),
    textFormat: { italic: true, fontSize: 8, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,4,10), cell: { userEnteredFormat: {
    textFormat: { bold: true, italic: false, fontSize: 9, foregroundColor: hex(C.primary) },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment' }});
  // Percent formats for baseline reference cells
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,6,8), cell: { userEnteredFormat: {
    numberFormat: { type: 'PERCENT', pattern: '0.0%' },
  }}, fields: 'userEnteredFormat.numberFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // ===== COLUMN HEADERS (row 5, 0-idx 4) =====
  vals.push({ range: `${S}!A5`, values: [[
    'Scenario','Scenario Name','Description',
    'P1 Retire\nAge','P2 Retire\nAge','Inflation\nRate','Exp. Return','Spending\nAdj.','Life\nExp.',
    'P1 Yrs to\nRetire','Est. Portfolio\n@ Retirement','Annual SWR\nIncome','Annual Income\nGoal (Adj.)','Annual Gap\n(+) Surplus',
  ]] });
  fmt.push({ repeatCell: { range: gridRange(SID,4,5,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
    borders: { bottom: { style: 'SOLID_MEDIUM', color: hex(C.primary) } },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 32 }, fields: 'pixelSize' }});
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' }});

  // ===== SCENARIO DATA ROWS (rows 6-11, 0-idx 5-10) =====
  const DATA_START = 5;

  SCENARIOS.forEach((scn, si) => {
    const rIdx = DATA_START + si;
    const rowNum = rIdx + 1;
    const scnColor = SCN_COLORS[si];

    // Years to retirement P1 for scenario
    const ytrP1 = `IFERROR(D${rowNum}-DATEDIF(${SETUP}!B8,TODAY(),"Y"),0)`;
    // Annual SWR: K × SWR rate from Setup
    const annualSWR = `=IFERROR(K${rowNum}*${SETUP}!B28,"")`;
    // Annual Goal adjusted by spending change: B30 from Setup is Annual Goal
    const annualGoalAdj = `=IFERROR(${SETUP}!B30*(1+H${rowNum}),"")`;
    // Annual Gap
    const annualGap = `=IFERROR(L${rowNum}-M${rowNum},"")`;

    vals.push({ range: `${S}!A${rowNum}`, values: [[
      scn.id, scn.name, scn.desc,
      scn.p1Age, scn.p2Age, scn.infl, scn.ret, scn.spend, scn.lifeExp,
      `=IFERROR(MAX(0,D${rowNum}-DATEDIF(${SETUP}!B8,TODAY(),"Y")),0)`,
      portfolioFV(rowNum),
      annualSWR, annualGoalAdj, annualGap,
    ]] });

    // Base row style
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,0,14), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});

    // Scenario ID — colored badge
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,0,1), cell: { userEnteredFormat: {
      backgroundColor: hex(scnColor), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text) },
      horizontalAlignment: 'CENTER',
      borders: { left: { style: 'SOLID_MEDIUM', color: hex(scnColor) } },
    }}, fields: 'userEnteredFormat' }});

    // Name (bold)
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,1,2), cell: { userEnteredFormat: {
      textFormat: { bold: true, fontSize: 9 }, padding: { left: 4 },
    }}, fields: 'userEnteredFormat.textFormat,userEnteredFormat.padding' }});

    // Description
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,2,3), cell: { userEnteredFormat: {
      textFormat: { fontSize: 8, foregroundColor: hex(C.secText) },
      padding: { left: 4 }, wrapStrategy: 'WRAP',
    }}, fields: 'userEnteredFormat' }});

    // Editable input cols D-I (retirement ages, rates)
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,3,9), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), horizontalAlignment: 'CENTER',
      textFormat: { bold: true, fontSize: 9 },
    }}, fields: 'userEnteredFormat' }});

    // % formats for F(infl), G(ret), H(spend)
    [5,6,7].forEach(ci => {
      fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,ci,ci+1), cell: { userEnteredFormat: {
        numberFormat: { type: 'PERCENT', pattern: '0.0%' },
      }}, fields: 'userEnteredFormat.numberFormat' }});
    });

    // Computed cols J-N (formula cells)
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,9,14), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula), textFormat: { bold: true, fontSize: 9 },
      horizontalAlignment: 'RIGHT',
    }}, fields: 'userEnteredFormat' }});

    // J: years to retire — number format
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,9,10), cell: { userEnteredFormat: {
      numberFormat: { type: 'NUMBER', pattern: '0 "yrs"' }, horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat.numberFormat,userEnteredFormat.horizontalAlignment' }});

    // K, L, M: currency
    [10,11,12].forEach(ci => {
      fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,ci,ci+1), cell: { userEnteredFormat: {
        numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' },
      }}, fields: 'userEnteredFormat.numberFormat' }});
    });

    // N (Gap): currency with color conditional inline
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,13,14), cell: { userEnteredFormat: {
      numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0;[RED]-"$"#,##0' },
    }}, fields: 'userEnteredFormat.numberFormat' }});

    fmt.push({ updateBorders: { range: gridRange(SID,rIdx,rIdx+1,0,14), bottom: { style: 'SOLID', color: hex(C.border) }}});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: rIdx, endIndex: rIdx+1 },
      properties: { pixelSize: 30 }, fields: 'pixelSize' }});
  });

  // ===== HELPER TABLE FOR CHARTS (cols P-Q, starting row 5) =====
  // Helper A: Portfolio at Retirement per scenario (P15:Q21)
  const HELP_COL = 15; // col P (0-indexed)
  const HELP_ROW = 14; // row 15 (0-indexed 14)

  // spacer rows 12-13
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 11, endIndex: 13 },
    properties: { pixelSize: 10 }, fields: 'pixelSize' }});

  // Section header for helper
  vals.push({ range: `${S}!A13`, values: [['SCENARIO OUTCOME SNAPSHOT']] });
  fmt.push({ mergeCells: { range: gridRange(SID,12,13,0,8), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,12,13,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 12, endIndex: 13 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Snapshot table header (row 14, 0-idx 13)
  vals.push({ range: `${S}!A14`, values: [[
    'Scenario','Name','P1 Retire Age','Years to Retire','Est. Portfolio','Annual SWR Income','Annual Goal','Gap (+Surplus)',
  ]] });
  fmt.push({ repeatCell: { range: gridRange(SID,13,14,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrB), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    borders: { bottom: { style: 'SOLID', color: hex(C.primary) } },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 13, endIndex: 14 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  SCENARIOS.forEach((scn, si) => {
    const srcRow = DATA_START + si + 1; // 1-indexed data row
    const snapRow = 14 + si + 1;       // snapshot table row (1-indexed row 15-20)
    const sIdx = 14 + si;              // 0-indexed
    const scnColor = SCN_COLORS[si];

    vals.push({ range: `${S}!A${snapRow}`, values: [[
      `=A${srcRow}`, `=B${srcRow}`, `=D${srcRow}`, `=J${srcRow}`,
      `=K${srcRow}`, `=L${srcRow}`, `=M${srcRow}`, `=N${srcRow}`,
    ]] });

    fmt.push({ repeatCell: { range: gridRange(SID,sIdx,sIdx+1,0,8), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel), textFormat: { fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE', borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
    }}, fields: 'userEnteredFormat' }});

    // ID badge
    fmt.push({ repeatCell: { range: gridRange(SID,sIdx,sIdx+1,0,1), cell: { userEnteredFormat: {
      backgroundColor: hex(scnColor), textFormat: { bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat' }});

    // Currency formats cols E-H (indices 4-7)
    [4,5,6,7].forEach(ci => {
      fmt.push({ repeatCell: { range: gridRange(SID,sIdx,sIdx+1,ci,ci+1), cell: { userEnteredFormat: {
        numberFormat: { type: ci === 3 ? 'NUMBER' : 'CURRENCY', pattern: ci === 3 ? '0 "yrs"' : '"$"#,##0' },
        horizontalAlignment: 'RIGHT', textFormat: { bold: true },
      }}, fields: 'userEnteredFormat.numberFormat,userEnteredFormat.horizontalAlignment,userEnteredFormat.textFormat' }});
    });
    // Col D (years) — number
    fmt.push({ repeatCell: { range: gridRange(SID,sIdx,sIdx+1,3,4), cell: { userEnteredFormat: {
      numberFormat: { type: 'NUMBER', pattern: '0 "yrs"' }, horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat.numberFormat,userEnteredFormat.horizontalAlignment' }});

    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: sIdx, endIndex: sIdx+1 },
      properties: { pixelSize: 24 }, fields: 'pixelSize' }});
  });

  // ===== HELPER TABLE FOR BAR CHART (off-screen cols P-Q) =====
  // P = scenario names, Q = estimated portfolio
  const hCol = 15; // 0-indexed col P
  const hStartRow = 4; // 0-indexed (row 5 = header)
  vals.push({ range: `${S}!P5`, values: [['Scenario','Est. Portfolio','Annual Gap']] });
  SCENARIOS.forEach((scn, si) => {
    const srcRow = DATA_START + si + 1;
    vals.push({ range: `${S}!P${5+si+1}`, values: [[
      `=B${srcRow}`, `=K${srcRow}`, `=N${srcRow}`,
    ]] });
  });
  // Format helper cols
  fmt.push({ repeatCell: { range: gridRange(SID,4,11,hCol,hCol+3), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow), textFormat: { fontSize: 8, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: hCol, endIndex: hCol+3 },
    properties: { pixelSize: 120 }, fields: 'pixelSize' }});

  // ===== CHARTS =====
  const CHART_ANCHOR = (row, col) => ({ overlayPosition: { anchorCell: { sheetId: SID, rowIndex: row, columnIndex: col }, widthPixels: 440, heightPixels: 260 }});

  // Chart 1: Estimated Portfolio at Retirement (bar)
  const chart1 = {
    addChart: { chart: {
      spec: {
        title: 'Estimated Portfolio at Retirement',
        basicChart: {
          chartType: 'BAR',
          legendPosition: 'NO_LEGEND',
          headerCount: 1,
          domains: [{ domain: { sourceRange: { sources: [{
            sheetId: SID, startRowIndex: 4, endRowIndex: 11, startColumnIndex: 15, endColumnIndex: 16,
          }] }}}],
          series: [{ series: { sourceRange: { sources: [{
            sheetId: SID, startRowIndex: 4, endRowIndex: 11, startColumnIndex: 16, endColumnIndex: 17,
          }] }}, color: hex(C.secondary) }],
        },
      },
      position: CHART_ANCHOR(1, 9),
    }},
  };

  // Chart 2: Annual Gap per scenario (column)
  const chart2 = {
    addChart: { chart: {
      spec: {
        title: 'Annual Income Gap by Scenario (+Surplus / −Shortfall)',
        basicChart: {
          chartType: 'COLUMN',
          legendPosition: 'NO_LEGEND',
          headerCount: 1,
          domains: [{ domain: { sourceRange: { sources: [{
            sheetId: SID, startRowIndex: 4, endRowIndex: 11, startColumnIndex: 15, endColumnIndex: 16,
          }] }}}],
          series: [{ series: { sourceRange: { sources: [{
            sheetId: SID, startRowIndex: 4, endRowIndex: 11, startColumnIndex: 17, endColumnIndex: 18,
          }] }}, color: hex(C.SCN01) }],
        },
      },
      position: { overlayPosition: { anchorCell: { sheetId: SID, rowIndex: 1, columnIndex: 12 }, widthPixels: 440, heightPixels: 260 }},
    }},
  };

  // ===== DISCLAIMER =====
  const DISC_ROW = 20;
  vals.push({ range: `${S}!A21`, values: [[
    'Note: Projected portfolio values use a simplified future-value formula (FV = PV·(1+r)ⁿ + PMT·((1+r)ⁿ−1)/r). ' +
    'Results are illustrative only — actual returns will vary. Consult a licensed financial advisor before making decisions.',
  ]] });
  fmt.push({ mergeCells: { range: gridRange(SID,20,21,0,14), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,20,21,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow), textFormat: { fontSize: 8, foregroundColor: hex(C.secText), italic: true, fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 20, endIndex: 21 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  await valuesBatchUpdate(id, vals, '07-scenarios values');
  await batchUpdate(id, fmt, '07-scenarios format');
  await batchUpdate(id, [chart1, chart2], '07-scenarios charts');

  console.log('✅ Scenario Planner done — 6 scenarios + 2 charts.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
