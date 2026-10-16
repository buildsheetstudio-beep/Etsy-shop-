'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID  = sheetMap['Dividend Income'];
const S    = "'Dividend Income'";
const PRC  = "'Price Updates'";
const ACT  = "'Account Tracker'";

// ─── Column layout (F=AcctID, G=SecID, O=Net, S=Status — must match Holdings formula) ──
// A:DivNum B:Date C:Ticker D:SecName E:AcctName F:AcctID G:SecID H:Owner I:AssetClass
// J:Shares K:DivPerShare L:Gross M:TaxWithheld N:DivType O:NetAmount P:Frequency Q:Qualified R:Reinvested S:Status

// ─── Row layout (data MUST start at 1-indexed row 6 per Holdings formula) ─────
const R_TITLE  = 0;  // title banner (rows 0-1 merged)
const R_CARDHDR = 2;
const R_CARD   = 3;
const R_DCOLS  = 4;  // column headers (1-indexed row 5)
const R_DD0    = 5;  // data start (1-indexed row 6) ← Holdings formula references this

// ─── Payer definitions ───────────────────────────────────────────────────────
const PAYERS = [
  // [ticker, acct, sec, shares, dps, freq, qualified, months(null=every month), costBasis]
  { t:'BND',  a:'ACT-007', s:'SEC-0003', sh:300, dps:0.32,  freq:'Monthly',   qual:false, mo:null,        cost:73.80  },
  { t:'VTI',  a:'ACT-001', s:'SEC-0001', sh:120, dps:0.88,  freq:'Quarterly', qual:true,  mo:[3,6,9,12],  cost:195.00 },
  { t:'VTI',  a:'ACT-007', s:'SEC-0001', sh:180, dps:0.88,  freq:'Quarterly', qual:true,  mo:[3,6,9,12],  cost:205.00 },
  { t:'SCHD', a:'ACT-001', s:'SEC-0006', sh:200, dps:0.26,  freq:'Quarterly', qual:true,  mo:[3,6,9,12],  cost:27.80  },
  { t:'VNQ',  a:'ACT-007', s:'SEC-0004', sh:120, dps:0.72,  freq:'Quarterly', qual:false, mo:[3,6,9,12],  cost:86.00  },
  { t:'AMT',  a:'ACT-007', s:'SEC-0025', sh:50,  dps:1.62,  freq:'Quarterly', qual:false, mo:[2,5,8,11],  cost:215.00 },
  { t:'SPG',  a:'ACT-007', s:'SEC-0023', sh:30,  dps:2.05,  freq:'Quarterly', qual:false, mo:[1,4,7,10],  cost:155.00 },
];

function generateDivRecords() {
  const recs = [];
  for (let m = 0; m < 26; m++) {
    const year  = 2024 + Math.floor(m / 12);
    const month = (m % 12) + 1;
    PAYERS.forEach(p => {
      if (p.mo !== null && !p.mo.includes(month)) return;
      const day   = p.freq === 'Monthly' ? 15 : 20;
      const gross = +(p.sh * p.dps).toFixed(2);
      recs.push({ date:`=DATE(${year},${month},${day})`, acct:p.a, sec:p.s, ticker:p.t,
        shares:p.sh, dps:p.dps, gross, tax:0, divType:'Regular',
        freq:p.freq, qual:p.qual, cost:p.cost, status:'Received' });
    });
  }
  return recs.sort((a, b) => new Date(a.date) - new Date(b.date));
}

(async () => {
  const fmt  = [];
  const vals = [];

  const recs = generateDivRecords();
  const ND   = recs.length;
  const DDN  = R_DD0 + ND;    // last data row (0-indexed, exclusive) = first row after data
  const DD1  = R_DD0 + 1;     // first data row 1-indexed (= 6)
  const DDNV = DDN;            // 1-indexed last data row for formulas

  // ── Background ────────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 0, DDN + 100, 0, 22),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat.backgroundColor' } });

  // ── Column widths (A-S + chart spacer) ───────────────────────────────────────
  const COL_W = [80,90,65,175,148,78,78,105,120,75,80,90,75,95,90,90,70,70,85, 20,20,20];
  COL_W.forEach((px, ci) => fmt.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
    properties: { pixelSize: px }, fields: 'pixelSize' } }));

  // ── Title banner (rows 0-1) ──────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 2, 0, 19), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A1`, values: [['DIVIDEND INCOME\nAll dividend payments • Yield tracking • Account & tax detail']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 2, 0, 19),
    cell: { userEnteredFormat: { backgroundColor: hex(C.primary),
      textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 46 }, fields: 'pixelSize' } });

  // ── Summary cards (rows 2-3) — each 4-5 cols wide across A-S ────────────────
  const total2024 = `=IFERROR(SUMPRODUCT((IFERROR(YEAR($B$${DD1}:$B$2005),0)=2024)*($S$${DD1}:$S$2005="Received")*($O$${DD1}:$O$2005)),0)`;
  const total2025 = `=IFERROR(SUMPRODUCT((IFERROR(YEAR($B$${DD1}:$B$2005),0)=2025)*($S$${DD1}:$S$2005="Received")*($O$${DD1}:$O$2005)),0)`;
  const total2026 = `=IFERROR(SUMPRODUCT((IFERROR(YEAR($B$${DD1}:$B$2005),0)=2026)*($S$${DD1}:$S$2005="Received")*($O$${DD1}:$O$2005)),0)`;
  const avgMo2025 = `=IFERROR(${total2025.slice(1)}/12,0)`;

  const cards = [
    { label: '2026 YTD DIVIDENDS', val: total2026,                            nf: '$#,##0.00', color: C.primary,   tcol: C.primaryText, cols: [0,5] },
    { label: '2025 ANNUAL TOTAL',  val: total2025,                            nf: '$#,##0.00', color: C.hdrC,      tcol: C.primaryText, cols: [5,10] },
    { label: '2024 ANNUAL TOTAL',  val: total2024,                            nf: '$#,##0.00', color: C.highlight, tcol: C.text,        cols: [10,14] },
    { label: 'AVG MONTHLY (2025)', val: `=IFERROR(${total2025.slice(1)}/12,0)`, nf: '$#,##0.00', color: C.info,      tcol: C.text,        cols: [14,19] },
  ];
  cards.forEach(({ label, val, nf, color, tcol, cols: [c1, c2] }) => {
    fmt.push({ mergeCells: { range: gridRange(SID, R_CARDHDR, R_CARDHDR + 1, c1, c2), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${cl(c1)}${R_CARDHDR + 1}`, values: [[label]] });
    fmt.push({ repeatCell: { range: gridRange(SID, R_CARDHDR, R_CARDHDR + 1, c1, c2),
      cell: { userEnteredFormat: { backgroundColor: hex(color),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(tcol), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } },
      fields: 'userEnteredFormat' } });
    fmt.push({ mergeCells: { range: gridRange(SID, R_CARD, R_CARD + 1, c1, c2), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${cl(c1)}${R_CARD + 1}`, values: [[val]] });
    fmt.push({ repeatCell: { range: gridRange(SID, R_CARD, R_CARD + 1, c1, c2),
      cell: { userEnteredFormat: { backgroundColor: hex(color),
        textFormat: { bold: true, fontSize: 18, foregroundColor: hex(tcol), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'TOP',
        numberFormat: { type: 'NUMBER', pattern: nf } } },
      fields: 'userEnteredFormat' } });
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_CARDHDR, endIndex: R_CARDHDR + 1 },
    properties: { pixelSize: 20 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_CARD, endIndex: R_CARD + 1 },
    properties: { pixelSize: 42 }, fields: 'pixelSize' } });

  // ── Column headers ────────────────────────────────────────────────────────────
  const HDRS = ['Div #','Date','Ticker','Security Name','Account Name',
    'Account ID','Security ID','Owner','Asset Class',
    'Shares','Div / Share','Gross','Tax Withheld','Div Type',
    'Net Amount','Frequency','Qualified?','Reinvested?','Status'];
  vals.push({ range: `${S}!A${R_DCOLS + 1}`, values: [HDRS] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_DCOLS, R_DCOLS + 1, 0, 19),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrB),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });

  // ── Data rows ─────────────────────────────────────────────────────────────────
  const dataRows = recs.map((rec, i) => {
    const r1 = DD1 + i;
    const divNumFml   = `=IF(B${r1}="","","DIV-"&TEXT(ROW()-5,"0000"))`;
    const secNameFml  = `=IFERROR(VLOOKUP(G${r1},${PRC}!$A$6:$C$1005,3,FALSE),"")`;
    const acctNameFml = `=IFERROR(VLOOKUP(F${r1},${ACT}!$A$6:$B$305,2,FALSE),"")`;
    const ownerFml    = `=IFERROR(VLOOKUP(F${r1},${ACT}!$A$6:$C$305,3,FALSE),"")`;
    const assetClsFml = `=IFERROR(VLOOKUP(G${r1},${PRC}!$A$6:$E$1005,5,FALSE),"")`;
    const grossFml    = `=IFERROR(J${r1}*K${r1},0)`;
    const netFml      = `=IFERROR(L${r1}-M${r1},0)`;
    return [
      divNumFml, rec.date, rec.ticker, secNameFml, acctNameFml,
      rec.acct, rec.sec, ownerFml, assetClsFml,
      rec.shares, rec.dps, grossFml, rec.tax, rec.divType,
      netFml, rec.freq, rec.qual, false, rec.status,
    ];
  });
  vals.push({ range: `${S}!A${DD1}`, values: dataRows });

  // ── Data formatting ───────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, R_DD0, DDN, 0, 19),
    cell: { userEnteredFormat: { backgroundColor: hex(C.panel),
      textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  for (let i = 0; i < ND; i += 2) {
    fmt.push({ repeatCell: { range: gridRange(SID, R_DD0 + i, R_DD0 + i + 1, 0, 19),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor' } });
  }
  // Date format
  fmt.push({ repeatCell: { range: gridRange(SID, R_DD0, DDN, 1, 2),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } },
    fields: 'userEnteredFormat.numberFormat' } });
  // Currency: Gross(L=11), TaxW/H(M=12), Net(O=14)
  [11, 12, 14].forEach(ci => fmt.push({ repeatCell: { range: gridRange(SID, R_DD0, DDN, ci, ci + 1),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '$#,##0.00' },
      horizontalAlignment: 'RIGHT' } },
    fields: 'userEnteredFormat' } }));
  // Div/Share (K=10): $0.000
  fmt.push({ repeatCell: { range: gridRange(SID, R_DD0, DDN, 10, 11),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '$0.000' },
      horizontalAlignment: 'RIGHT' } },
    fields: 'userEnteredFormat' } });
  // Checkboxes: Qualified (Q=16), Reinvested (R=17)
  [16, 17].forEach(ci => fmt.push({ setDataValidation: { range: gridRange(SID, R_DD0, DDN, ci, ci + 1),
    rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true } } }));

  // ── Status conditional formats ────────────────────────────────────────────────
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_DD0, DDN, 0, 19)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$S${DD1}="Received"` }] },
      format: { backgroundColor: hex(C.success) } }
  }, index: 0 } });
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_DD0, DDN, 0, 19)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$S${DD1}="Expected"` }] },
      format: { backgroundColor: hex(C.info) } }
  }, index: 1 } });

  // ── Data validation: Status ───────────────────────────────────────────────────
  fmt.push({ setDataValidation: { range: gridRange(SID, R_DD0, DDN, 18, 19),
    rule: { condition: { type: 'ONE_OF_LIST',
      values: [
        { userEnteredValue: 'Received' },
        { userEnteredValue: 'Expected' },
        { userEnteredValue: 'Scheduled' },
      ] }, showCustomUi: true, strict: true } } });

  // ── Freeze rows ───────────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID,
    gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' } });

  // ─── Apply fmt + data ─────────────────────────────────────────────────────────
  await batchUpdate(id, fmt, 'div-fmt');
  await valuesBatchUpdate(id, vals, 'div-vals');

  // ── Summary sections (for charts) ─────────────────────────────────────────────
  const summaryFmt  = [];
  const summaryVals = [];

  // Annual summary (3 years)
  const R_ANNHDR  = DDN + 1;
  const R_ANNCOLS = R_ANNHDR + 1;
  const R_ANNDATA = R_ANNCOLS + 1;

  summaryFmt.push({ mergeCells: { range: gridRange(SID, R_ANNHDR, R_ANNHDR + 1, 0, 6), mergeType: 'MERGE_ALL' } });
  summaryVals.push({ range: `${S}!A${R_ANNHDR + 1}`, values: [['ANNUAL SUMMARY']] });
  summaryFmt.push({ repeatCell: { range: gridRange(SID, R_ANNHDR, R_ANNHDR + 1, 0, 6),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrA),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' } } },
    fields: 'userEnteredFormat' } });

  summaryVals.push({ range: `${S}!A${R_ANNCOLS + 1}`, values: [['Year','# Payments','Gross','Tax W/H','Net Dividends','Avg / Month']] });
  summaryFmt.push({ repeatCell: { range: gridRange(SID, R_ANNCOLS, R_ANNCOLS + 1, 0, 6),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrB),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat' } });

  const annRows = [2024, 2025, 2026].map((yr, ai) => {
    const r1 = R_ANNDATA + ai + 1;
    const ct  = `=IFERROR(SUMPRODUCT((YEAR($B$${DD1}:$B$${DDN})=${yr})*($S$${DD1}:$S$${DDN}="Received")),0)`;
    const grs = `=IFERROR(SUMPRODUCT((YEAR($B$${DD1}:$B$${DDN})=${yr})*($S$${DD1}:$S$${DDN}="Received")*($L$${DD1}:$L$${DDN})),0)`;
    const tax = `=IFERROR(SUMPRODUCT((YEAR($B$${DD1}:$B$${DDN})=${yr})*($S$${DD1}:$S$${DDN}="Received")*($M$${DD1}:$M$${DDN})),0)`;
    const net = `=IFERROR(SUMPRODUCT((YEAR($B$${DD1}:$B$${DDN})=${yr})*($S$${DD1}:$S$${DDN}="Received")*($O$${DD1}:$O$${DDN})),0)`;
    const avg = `=IFERROR(E${r1}/12,0)`;
    return [yr, ct, grs, tax, net, avg];
  });
  summaryVals.push({ range: `${S}!A${R_ANNDATA + 1}`, values: annRows });
  summaryFmt.push({ repeatCell: { range: gridRange(SID, R_ANNDATA, R_ANNDATA + 3, 0, 6),
    cell: { userEnteredFormat: { backgroundColor: hex(C.panel),
      textFormat: { fontSize: 9, fontFamily: 'Arial' },
      numberFormat: { type: 'NUMBER', pattern: '$#,##0.00' } } },
    fields: 'userEnteredFormat' } });
  summaryFmt.push({ repeatCell: { range: gridRange(SID, R_ANNDATA, R_ANNDATA + 3, 0, 2),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0' } } },
    fields: 'userEnteredFormat.numberFormat' } });

  // Monthly totals for trend chart (26 months)
  const R_MONHDR  = R_ANNDATA + 5;
  const R_MONCOLS = R_MONHDR + 1;
  const R_MONDATA = R_MONCOLS + 1;

  summaryFmt.push({ mergeCells: { range: gridRange(SID, R_MONHDR, R_MONHDR + 1, 0, 4), mergeType: 'MERGE_ALL' } });
  summaryVals.push({ range: `${S}!A${R_MONHDR + 1}`, values: [['MONTHLY TOTALS (2024–2026)']] });
  summaryFmt.push({ repeatCell: { range: gridRange(SID, R_MONHDR, R_MONHDR + 1, 0, 4),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrA),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' } } },
    fields: 'userEnteredFormat' } });

  summaryVals.push({ range: `${S}!A${R_MONCOLS + 1}`, values: [['Month','# Payments','Net Received','Cumulative']] });
  summaryFmt.push({ repeatCell: { range: gridRange(SID, R_MONCOLS, R_MONCOLS + 1, 0, 4),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrB),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat' } });

  const monthRows = [];
  for (let m = 0; m < 26; m++) {
    const yr  = 2024 + Math.floor(m / 12);
    const mo  = (m % 12) + 1;
    const ri  = R_MONDATA + m + 1;
    const ct  = `=IFERROR(SUMPRODUCT((YEAR($B$${DD1}:$B$${DDN})=${yr})*(MONTH($B$${DD1}:$B$${DDN})=${mo})*($S$${DD1}:$S$${DDN}="Received")),0)`;
    const net = `=IFERROR(SUMPRODUCT((YEAR($B$${DD1}:$B$${DDN})=${yr})*(MONTH($B$${DD1}:$B$${DDN})=${mo})*($S$${DD1}:$S$${DDN}="Received")*($O$${DD1}:$O$${DDN})),0)`;
    const cum = m === 0
      ? `=IFERROR(C${ri},0)`
      : `=IFERROR(C${ri}+D${ri - 1},0)`;
    monthRows.push([`${mo}/1/${yr}`, ct, net, cum]);
  }
  summaryVals.push({ range: `${S}!A${R_MONDATA + 1}`, values: monthRows });
  summaryFmt.push({ repeatCell: { range: gridRange(SID, R_MONDATA, R_MONDATA + 26, 0, 4),
    cell: { userEnteredFormat: { backgroundColor: hex(C.panel),
      textFormat: { fontSize: 9, fontFamily: 'Arial' } } },
    fields: 'userEnteredFormat' } });
  summaryFmt.push({ repeatCell: { range: gridRange(SID, R_MONDATA, R_MONDATA + 26, 0, 1),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm yyyy' } } },
    fields: 'userEnteredFormat.numberFormat' } });
  summaryFmt.push({ repeatCell: { range: gridRange(SID, R_MONDATA, R_MONDATA + 26, 2, 4),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '$#,##0.00' },
      horizontalAlignment: 'RIGHT' } },
    fields: 'userEnteredFormat.numberFormat' } });

  await batchUpdate(id, summaryFmt, 'div-summary-fmt');
  await valuesBatchUpdate(id, summaryVals, 'div-summary-vals');

  // ── Charts ────────────────────────────────────────────────────────────────────
  const R_MEND = R_MONDATA + 26;  // exclusive end of monthly data
  const charts = [];

  // 1. Monthly dividend net income — COLUMN
  charts.push({ addChart: { chart: {
    spec: { title: 'Monthly Dividend Income', basicChart: {
      chartType: 'COLUMN', legendPosition: 'NO_LEGEND',
      axis: [{ position: 'BOTTOM_AXIS', title: 'Month' },
             { position: 'LEFT_AXIS',   title: 'Net Dividends ($)' }],
      domains: [{ domain: { sourceRange: { sources: [gridRange(SID, R_MONCOLS, R_MEND, 0, 1)] } } }],
      series: [{ series: { sourceRange: { sources: [gridRange(SID, R_MONCOLS, R_MEND, 2, 3)] } },
        targetAxis: 'LEFT_AXIS', color: hex(C.success) }],
      headerCount: 1,
    } },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_ANNHDR, columnIndex: 7 },
      widthPixels: 480, heightPixels: 260,
    } },
  } } });

  // 2. Cumulative dividends — LINE
  charts.push({ addChart: { chart: {
    spec: { title: 'Cumulative Dividends Received', basicChart: {
      chartType: 'LINE', legendPosition: 'NO_LEGEND',
      axis: [{ position: 'BOTTOM_AXIS', title: 'Month' },
             { position: 'LEFT_AXIS',   title: 'Cumulative ($)' }],
      domains: [{ domain: { sourceRange: { sources: [gridRange(SID, R_MONCOLS, R_MEND, 0, 1)] } } }],
      series: [{ series: { sourceRange: { sources: [gridRange(SID, R_MONCOLS, R_MEND, 3, 4)] } },
        targetAxis: 'LEFT_AXIS', color: hex(C.primary), lineStyle: { width: 2 } }],
      headerCount: 1,
    } },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_ANNHDR + 15, columnIndex: 7 },
      widthPixels: 480, heightPixels: 260,
    } },
  } } });

  // 3. Annual comparison — COLUMN
  charts.push({ addChart: { chart: {
    spec: { title: 'Annual Dividend Comparison', basicChart: {
      chartType: 'COLUMN', legendPosition: 'NO_LEGEND',
      axis: [{ position: 'BOTTOM_AXIS', title: 'Year' },
             { position: 'LEFT_AXIS',   title: 'Net Dividends ($)' }],
      domains: [{ domain: { sourceRange: { sources: [gridRange(SID, R_ANNCOLS, R_ANNDATA + 3, 0, 1)] } } }],
      series: [{ series: { sourceRange: { sources: [gridRange(SID, R_ANNCOLS, R_ANNDATA + 3, 4, 5)] } },
        targetAxis: 'LEFT_AXIS', color: hex(C.hdrC) }],
      headerCount: 1,
    } },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_MEND + 2, columnIndex: 7 },
      widthPixels: 300, heightPixels: 240,
    } },
  } } });

  await batchUpdate(id, charts, 'div-charts');
  console.log(`✓ Dividend Income complete (${ND} records)`);
})();

function cl(i) { return String.fromCharCode(65 + i); }
