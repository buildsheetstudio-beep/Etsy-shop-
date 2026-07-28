'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DB = sheetMap['Dashboard'];
const S  = "'Dashboard'";
const IL  = "'Invoice Log'";
const PT  = "'Payment Tracker'";

// Control cells: B5=Year, D5=Month(or "Full Year"), F5=Client(or "All"), H5=Status(or "All")
// All KPI formulas filter by these controls.

const currFmt  = { type: 'CURRENCY', pattern: '$#,##0.00;[Red]-$#,##0.00' };
const numFmt   = { type: 'NUMBER',   pattern: '#,##0' };
const pctFmt   = { type: 'PERCENT',  pattern: '0.0%' };
const dateFmt  = { type: 'DATE',     pattern: 'MMM D, YYYY' };

// Filter logic embedded in SUMPRODUCT/SUMIFS
// Year match: YEAR(IL!E:E)=B5 (or B5="All")
// Month match: TEXT(IL!E:E,"mmmm")=D5 OR D5="Full Year"
// Client match: IL!D:D=F5 OR F5="All Clients"
// Status match: IL!T:T=H5 OR H5="All Statuses"

function yFilter(yr)   { return `(YEAR('Invoice Log'!$E$6:$E$1005)=$B$5)`; }
function mFilter()     { return `(($D$5="Full Year")+(TEXT('Invoice Log'!$E$6:$E$1005,"mmmm")=$D$5)>0)`; }
function cFilter()     { return `(($F$5="All Clients")+('Invoice Log'!$D$6:$D$1005=$F$5)>0)`; }
function sFilter()     { return `(($H$5="All Statuses")+('Invoice Log'!$T$6:$T$1005=$H$5)>0)`; }
function notCancRef()  { return `('Invoice Log'!$T$6:$T$1005<>"Cancelled")*('Invoice Log'!$T$6:$T$1005<>"Refunded")`; }

const allF = `${yFilter()}*${mFilter()}*${cFilter()}*${sFilter()}`;
const allFnoS = `${yFilter()}*${mFilter()}*${cFilter()}`;

// KPI formulas
const kpis = {
  totalInvoiced:    `=IFERROR(SUMPRODUCT(${allF}*('Invoice Log'!$O$6:$O$1005)),0)`,
  paymentsRec:      `=IFERROR(SUMPRODUCT(${allFnoS}*('Invoice Log'!$Q$6:$Q$1005)),0)`,
  outstanding:      `=IFERROR(SUMPRODUCT(${allFnoS}*${notCancRef()}*('Invoice Log'!$R$6:$R$1005)),0)`,
  overdue:          `=IFERROR(SUMPRODUCT(${yFilter()}*${mFilter()}*${cFilter()}*('Invoice Log'!$T$6:$T$1005="Overdue")*('Invoice Log'!$R$6:$R$1005)),0)`,
  paidCount:        `=IFERROR(SUMPRODUCT(${allFnoS}*('Invoice Log'!$S$6:$S$1005="Paid")),0)`,
  unpaidCount:      `=IFERROR(SUMPRODUCT(${allFnoS}*${notCancRef()}*('Invoice Log'!$S$6:$S$1005<>"Paid")*('Invoice Log'!$S$6:$S$1005<>"")),0)`,
  avgInvoice:       `=IFERROR(SUMPRODUCT(${allF}*('Invoice Log'!$O$6:$O$1005))/MAX(1,SUMPRODUCT(${allF}*('Invoice Log'!$O$6:$O$1005>0))),0)`,
  collectionRate:   `=IFERROR(B10/B8,0)`,
  dueThisWeek:      `=IFERROR(SUMPRODUCT(('Invoice Log'!$F$6:$F$1005>=TODAY())*('Invoice Log'!$F$6:$F$1005<=TODAY()+7)*('Invoice Log'!$R$6:$R$1005>0)*${cFilter()}),0)`,
  dueThisMonth:     `=IFERROR(SUMPRODUCT((YEAR('Invoice Log'!$F$6:$F$1005)=YEAR(TODAY()))*(MONTH('Invoice Log'!$F$6:$F$1005)=MONTH(TODAY()))*('Invoice Log'!$R$6:$R$1005>0)*${cFilter()}),0)`,
  partialPaid:      `=IFERROR(SUMPRODUCT(${allFnoS}*('Invoice Log'!$S$6:$S$1005="Partially Paid")),0)`,
  overdueCount:     `=IFERROR(SUMPRODUCT(${allFnoS}*('Invoice Log'!$T$6:$T$1005="Overdue")),0)`,
  largestOutstanding:`=IFERROR(MAXIFS('Invoice Log'!$R$6:$R$1005,'Invoice Log'!$T$6:$T$1005,"<>Cancelled",'Invoice Log'!$T$6:$T$1005,"<>Paid"),0)`,
  topClient:        `=IFERROR(INDEX('Invoice Log'!$D$6:$D$1005,MATCH(MAXIFS('Invoice Log'!$O$6:$O$1005,'Invoice Log'!$D$6:$D$1005,'Invoice Log'!$D$6:$D$1005),'Invoice Log'!$O$6:$O$1005,0)),"")`,
  avgDaysToPay:     `=IFERROR(AVERAGEIFS('Invoice Log'!$AA$6:$AA$1005-'Invoice Log'!$E$6:$E$1005,'Invoice Log'!$S$6:$S$1005,"Paid"),0)`,
  depositsCollected:`=IFERROR(SUMPRODUCT(('Payment Tracker'!$K$6:$K$1505=TRUE)*('Payment Tracker'!$N$6:$N$1505)),0)`,
};

// Aging (based on Days Until/Overdue col U, negative = overdue)
const agingBuckets = [
  ['Current',      `=IFERROR(SUMPRODUCT(('Invoice Log'!$U$6:$U$1005>=0)*('Invoice Log'!$R$6:$R$1005>0)*${notCancRef()}),0)`,
                   `=IFERROR(SUMPRODUCT(('Invoice Log'!$U$6:$U$1005>=0)*('Invoice Log'!$R$6:$R$1005)),0)`],
  ['1–30 Days',    `=IFERROR(SUMPRODUCT(('Invoice Log'!$U$6:$U$1005>=-30)*('Invoice Log'!$U$6:$U$1005<0)*('Invoice Log'!$R$6:$R$1005>0)*${notCancRef()}),0)`,
                   `=IFERROR(SUMPRODUCT(('Invoice Log'!$U$6:$U$1005>=-30)*('Invoice Log'!$U$6:$U$1005<0)*('Invoice Log'!$R$6:$R$1005)),0)`],
  ['31–60 Days',   `=IFERROR(SUMPRODUCT(('Invoice Log'!$U$6:$U$1005>=-60)*('Invoice Log'!$U$6:$U$1005<-30)*('Invoice Log'!$R$6:$R$1005>0)*${notCancRef()}),0)`,
                   `=IFERROR(SUMPRODUCT(('Invoice Log'!$U$6:$U$1005>=-60)*('Invoice Log'!$U$6:$U$1005<-30)*('Invoice Log'!$R$6:$R$1005)),0)`],
  ['61–90 Days',   `=IFERROR(SUMPRODUCT(('Invoice Log'!$U$6:$U$1005>=-90)*('Invoice Log'!$U$6:$U$1005<-61)*('Invoice Log'!$R$6:$R$1005>0)*${notCancRef()}),0)`,
                   `=IFERROR(SUMPRODUCT(('Invoice Log'!$U$6:$U$1005>=-90)*('Invoice Log'!$U$6:$U$1005<-61)*('Invoice Log'!$R$6:$R$1005)),0)`],
  ['90+ Days',     `=IFERROR(SUMPRODUCT(('Invoice Log'!$U$6:$U$1005<-90)*('Invoice Log'!$R$6:$R$1005>0)*${notCancRef()}),0)`,
                   `=IFERROR(SUMPRODUCT(('Invoice Log'!$U$6:$U$1005<-90)*('Invoice Log'!$R$6:$R$1005)),0)`],
];

(async () => {
  const data = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['INVOICE TRACKER DASHBOARD']] });
  data.push({ range: `${S}!A2`, values: [['Invoicing • Payments • Aging • Revenue']] });
  data.push({ range: `${S}!A3`, values: [['For organizational purposes only. Not accounting, tax, legal, or financial advice. Verify invoice requirements, taxes, payment terms, and business records for your location and business.']] });

  reqs.push({ repeatCell: { range: gridRange(DB, 0, 500, 0, 12), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });

  reqs.push({ mergeCells: { range: gridRange(DB, 0, 1, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 0, 1, 0, 12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 18 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });

  reqs.push({ mergeCells: { range: gridRange(DB, 1, 2, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 1, 2, 0, 12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { foregroundColor: hex(C.white), fontSize: 10 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  reqs.push({ mergeCells: { range: gridRange(DB, 2, 3, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 2, 3, 0, 12), cell: { userEnteredFormat: {
    backgroundColor: hex('#F0EDE8'), textFormat: { italic: true, foregroundColor: hex(C.secText), fontSize: 8 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' } });

  // ── Controls (row 5) ──────────────────────────────────────────────────────
  data.push({ range: `${S}!A5:H5`, values: [['Reporting Year','','Month / Full Year','','Client','','Invoice Status','']] });
  data.push({ range: `${S}!B5`, values: [[2026]] });
  data.push({ range: `${S}!D5`, values: [['Full Year']] });
  data.push({ range: `${S}!F5`, values: [['All Clients']] });
  data.push({ range: `${S}!H5`, values: [['All Statuses']] });

  reqs.push({ repeatCell: { range: gridRange(DB, 4, 5, 0, 12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  // Input cells
  [1,3,5,7].forEach(ci => {
    reqs.push({ repeatCell: { range: gridRange(DB, 4, 5, ci, ci + 1), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // ── Primary KPI Cards (rows 7-11, 8 cards in 4 columns) ──────────────────
  const primaryKPIs = [
    ['Total Invoiced',       kpis.totalInvoiced,    currFmt,  'B8'],
    ['Payments Received',    kpis.paymentsRec,      currFmt,  'D8'],
    ['Outstanding Balance',  kpis.outstanding,      currFmt,  'F8'],
    ['Overdue Balance',      kpis.overdue,          currFmt,  'H8'],
    ['Paid Invoices',        kpis.paidCount,        numFmt,   'B10'],
    ['Unpaid Invoices',      kpis.unpaidCount,      numFmt,   'D10'],
    ['Avg Invoice Value',    kpis.avgInvoice,       currFmt,  'F10'],
    ['Collection Rate',      kpis.collectionRate,   pctFmt,   'H10'],
  ];

  primaryKPIs.forEach(([label, formula, fmt, cell], i) => {
    const col = ['A','C','E','G'][i % 4];
    const valCol = ['B','D','F','H'][i % 4];
    const baseRow = i < 4 ? 7 : 9;

    data.push({ range: `${S}!${col}${baseRow}`, values: [[label]] });
    data.push({ range: `${S}!${valCol}${baseRow + 1}`, values: [[formula]] });

    const ci = ['A','B','C','D','E','F','G','H'].indexOf(col);
    const vi = ci + 1;
    // Label
    reqs.push({ repeatCell: { range: gridRange(DB, baseRow - 1, baseRow, ci, ci + 2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
    reqs.push({ mergeCells: { range: gridRange(DB, baseRow - 1, baseRow, ci, ci + 2), mergeType: 'MERGE_ALL' } });
    // Value
    reqs.push({ repeatCell: { range: gridRange(DB, baseRow, baseRow + 1, ci, ci + 2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel), textFormat: { bold: true, foregroundColor: hex(C.mainText), fontSize: 14 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', numberFormat: fmt,
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,numberFormat)' } });
    reqs.push({ mergeCells: { range: gridRange(DB, baseRow, baseRow + 1, ci, ci + 2), mergeType: 'MERGE_ALL' } });
    reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'ROWS', startIndex: baseRow, endIndex: baseRow + 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });
  });

  // ── Secondary KPI Cards (rows 13-17) ─────────────────────────────────────
  const secondaryKPIs = [
    ['Due This Week',        kpis.dueThisWeek,      numFmt],
    ['Due This Month',       kpis.dueThisMonth,     numFmt],
    ['Partially Paid',       kpis.partialPaid,      numFmt],
    ['Overdue Invoices',     kpis.overdueCount,     numFmt],
    ['Largest Outstanding',  kpis.largestOutstanding,currFmt],
    ['Top Client by Revenue',kpis.topClient,        null],
    ['Avg Days to Pay',      kpis.avgDaysToPay,     numFmt],
    ['Deposits Collected',   kpis.depositsCollected,currFmt],
  ];

  secondaryKPIs.forEach(([label, formula, fmt], i) => {
    const col = ['A','C','E','G'][i % 4];
    const baseRow = i < 4 ? 13 : 15;
    const ci = ['A','B','C','D','E','F','G','H'].indexOf(col);

    data.push({ range: `${S}!${col}${baseRow}`, values: [[label]] });
    data.push({ range: `${S}!${col}${baseRow + 1}`, values: [[formula]] });

    reqs.push({ mergeCells: { range: gridRange(DB, baseRow - 1, baseRow, ci, ci + 2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(DB, baseRow - 1, baseRow, ci, ci + 2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
    reqs.push({ mergeCells: { range: gridRange(DB, baseRow, baseRow + 1, ci, ci + 2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(DB, baseRow, baseRow + 1, ci, ci + 2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel), textFormat: { bold: true, foregroundColor: hex(C.mainText), fontSize: 12 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      numberFormat: fmt || { type: 'TEXT' },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,numberFormat)' } });
  });

  // ── Aging Summary (rows 19-25) ─────────────────────────────────────────────
  data.push({ range: `${S}!A19`, values: [['AGING SUMMARY']] });
  data.push({ range: `${S}!A20:E20`, values: [['Bucket','Count','Balance','Share of Outstanding','']] });
  reqs.push({ mergeCells: { range: gridRange(DB, 18, 19, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 18, 19, 0, 12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 11 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 19, 20, 0, 5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  agingBuckets.forEach(([bucket, countFormula, balanceFormula], i) => {
    const r = 21 + i;
    data.push({ range: `${S}!A${r}`, values: [[bucket]] });
    data.push({ range: `${S}!B${r}`, values: [[countFormula]] });
    data.push({ range: `${S}!C${r}`, values: [[balanceFormula]] });
    data.push({ range: `${S}!D${r}`, values: [[`=IFERROR(C${r}/SUMPRODUCT(('Invoice Log'!$R$6:$R$1005)*${notCancRef()}),0)`]] });
    reqs.push({ repeatCell: { range: gridRange(DB, r - 1, r, 0, 5), cell: { userEnteredFormat: {
      backgroundColor: i % 2 === 0 ? hex(C.panel) : hex(C.altRow),
    }}, fields: 'userEnteredFormat.backgroundColor' } });
    reqs.push({ repeatCell: { range: gridRange(DB, r - 1, r, 2, 3), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields: 'userEnteredFormat.numberFormat' } });
    reqs.push({ repeatCell: { range: gridRange(DB, r - 1, r, 3, 4), cell: { userEnteredFormat: { numberFormat: pctFmt } }, fields: 'userEnteredFormat.numberFormat' } });
  });

  // ── Recent Payments Table (rows 28-35) ────────────────────────────────────
  data.push({ range: `${S}!A28`, values: [['RECENT PAYMENTS']] });
  reqs.push({ mergeCells: { range: gridRange(DB, 27, 28, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 27, 28, 0, 12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 11 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  data.push({ range: `${S}!A29:F29`, values: [['Date','Invoice ID','Client','Method','Gross','Net']] });
  reqs.push({ repeatCell: { range: gridRange(DB, 28, 29, 0, 6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  for (let i = 0; i < 7; i++) {
    const r = 30 + i;
    data.push({ range: `${S}!A${r}`, values: [[`=IFERROR(LARGE('Payment Tracker'!$B$6:$B$1505,${i+1}),"")`]] });
    data.push({ range: `${S}!B${r}`, values: [[`=IFERROR(INDEX('Payment Tracker'!$C$6:$C$1505,MATCH(LARGE('Payment Tracker'!$B$6:$B$1505,${i+1}),'Payment Tracker'!$B$6:$B$1505,0)),"")`]] });
    data.push({ range: `${S}!C${r}`, values: [[`=IFERROR(INDEX('Payment Tracker'!$F$6:$F$1505,MATCH(LARGE('Payment Tracker'!$B$6:$B$1505,${i+1}),'Payment Tracker'!$B$6:$B$1505,0)),"")`]] });
    data.push({ range: `${S}!D${r}`, values: [[`=IFERROR(INDEX('Payment Tracker'!$I$6:$I$1505,MATCH(LARGE('Payment Tracker'!$B$6:$B$1505,${i+1}),'Payment Tracker'!$B$6:$B$1505,0)),"")`]] });
    data.push({ range: `${S}!E${r}`, values: [[`=IFERROR(INDEX('Payment Tracker'!$H$6:$H$1505,MATCH(LARGE('Payment Tracker'!$B$6:$B$1505,${i+1}),'Payment Tracker'!$B$6:$B$1505,0)),"")`]] });
    data.push({ range: `${S}!F${r}`, values: [[`=IFERROR(INDEX('Payment Tracker'!$N$6:$N$1505,MATCH(LARGE('Payment Tracker'!$B$6:$B$1505,${i+1}),'Payment Tracker'!$B$6:$B$1505,0)),"")`]] });
    reqs.push({ repeatCell: { range: gridRange(DB, r - 1, r, 0, 7), cell: { userEnteredFormat: { backgroundColor: i % 2 === 0 ? hex(C.panel) : hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
    reqs.push({ repeatCell: { range: gridRange(DB, r - 1, r, 0, 1), cell: { userEnteredFormat: { numberFormat: dateFmt } }, fields: 'userEnteredFormat.numberFormat' } });
    reqs.push({ repeatCell: { range: gridRange(DB, r - 1, r, 4, 6), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields: 'userEnteredFormat.numberFormat' } });
  }

  // ── Upcoming + Overdue (rows 39-50) ───────────────────────────────────────
  data.push({ range: `${S}!A39`, values: [['UPCOMING DUE']] });
  reqs.push({ mergeCells: { range: gridRange(DB, 38, 39, 0, 6), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 38, 39, 0, 6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.warning), textFormat: { bold: true, foregroundColor: hex(C.mainText), fontSize: 11 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  data.push({ range: `${S}!A40:E40`, values: [['Invoice ID','Client','Due Date','Balance Due','Status']] });
  reqs.push({ repeatCell: { range: gridRange(DB, 39, 40, 0, 5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  for (let i = 0; i < 5; i++) {
    const r = 41 + i;
    data.push({ range: `${S}!A${r}`, values: [[`=IFERROR(INDEX('Invoice Log'!$A$6:$A$1005,MATCH(SMALL(IF(('Invoice Log'!$R$6:$R$1005>0)*('Invoice Log'!$U$6:$U$1005>=0)*('Invoice Log'!$U$6:$U$1005<=30),'Invoice Log'!$U$6:$U$1005,9999),${i+1}),'Invoice Log'!$U$6:$U$1005,0)),"")`]] });
    data.push({ range: `${S}!B${r}`, values: [[`=IFERROR(INDEX('Invoice Log'!$D$6:$D$1005,MATCH(SMALL(IF(('Invoice Log'!$R$6:$R$1005>0)*('Invoice Log'!$U$6:$U$1005>=0)*('Invoice Log'!$U$6:$U$1005<=30),'Invoice Log'!$U$6:$U$1005,9999),${i+1}),'Invoice Log'!$U$6:$U$1005,0)),"")`]] });
    data.push({ range: `${S}!C${r}`, values: [[`=IFERROR(INDEX('Invoice Log'!$F$6:$F$1005,MATCH(SMALL(IF(('Invoice Log'!$R$6:$R$1005>0)*('Invoice Log'!$U$6:$U$1005>=0)*('Invoice Log'!$U$6:$U$1005<=30),'Invoice Log'!$U$6:$U$1005,9999),${i+1}),'Invoice Log'!$U$6:$U$1005,0)),"")`]] });
    data.push({ range: `${S}!D${r}`, values: [[`=IFERROR(INDEX('Invoice Log'!$R$6:$R$1005,MATCH(SMALL(IF(('Invoice Log'!$R$6:$R$1005>0)*('Invoice Log'!$U$6:$U$1005>=0)*('Invoice Log'!$U$6:$U$1005<=30),'Invoice Log'!$U$6:$U$1005,9999),${i+1}),'Invoice Log'!$U$6:$U$1005,0)),"")`]] });
    data.push({ range: `${S}!E${r}`, values: [[`=IFERROR(INDEX('Invoice Log'!$T$6:$T$1005,MATCH(SMALL(IF(('Invoice Log'!$R$6:$R$1005>0)*('Invoice Log'!$U$6:$U$1005>=0)*('Invoice Log'!$U$6:$U$1005<=30),'Invoice Log'!$U$6:$U$1005,9999),${i+1}),'Invoice Log'!$U$6:$U$1005,0)),"")`]] });
    reqs.push({ repeatCell: { range: gridRange(DB, r - 1, r, 0, 5), cell: { userEnteredFormat: { backgroundColor: i % 2 === 0 ? hex(C.panel) : hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
    reqs.push({ repeatCell: { range: gridRange(DB, r - 1, r, 2, 3), cell: { userEnteredFormat: { numberFormat: dateFmt } }, fields: 'userEnteredFormat.numberFormat' } });
    reqs.push({ repeatCell: { range: gridRange(DB, r - 1, r, 3, 4), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields: 'userEnteredFormat.numberFormat' } });
  }

  data.push({ range: `${S}!G39`, values: [['OVERDUE INVOICES']] });
  reqs.push({ mergeCells: { range: gridRange(DB, 38, 39, 6, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 38, 39, 6, 12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.attention), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 11 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  data.push({ range: `${S}!G40:K40`, values: [['Invoice ID','Client','Days Overdue','Balance Due','Status']] });
  reqs.push({ repeatCell: { range: gridRange(DB, 39, 40, 6, 11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  for (let i = 0; i < 5; i++) {
    const r = 41 + i;
    // Overdue = U < 0, R > 0
    data.push({ range: `${S}!G${r}`, values: [[`=IFERROR(INDEX('Invoice Log'!$A$6:$A$1005,MATCH(LARGE(IF(('Invoice Log'!$R$6:$R$1005>0)*('Invoice Log'!$U$6:$U$1005<0),-'Invoice Log'!$U$6:$U$1005,0),${i+1}),-'Invoice Log'!$U$6:$U$1005,0)),"")`]] });
    data.push({ range: `${S}!H${r}`, values: [[`=IFERROR(INDEX('Invoice Log'!$D$6:$D$1005,MATCH(LARGE(IF(('Invoice Log'!$R$6:$R$1005>0)*('Invoice Log'!$U$6:$U$1005<0),-'Invoice Log'!$U$6:$U$1005,0),${i+1}),-'Invoice Log'!$U$6:$U$1005,0)),"")`]] });
    data.push({ range: `${S}!I${r}`, values: [[`=IFERROR(-INDEX('Invoice Log'!$U$6:$U$1005,MATCH(LARGE(IF(('Invoice Log'!$R$6:$R$1005>0)*('Invoice Log'!$U$6:$U$1005<0),-'Invoice Log'!$U$6:$U$1005,0),${i+1}),-'Invoice Log'!$U$6:$U$1005,0)),"")`]] });
    data.push({ range: `${S}!J${r}`, values: [[`=IFERROR(INDEX('Invoice Log'!$R$6:$R$1005,MATCH(LARGE(IF(('Invoice Log'!$R$6:$R$1005>0)*('Invoice Log'!$U$6:$U$1005<0),-'Invoice Log'!$U$6:$U$1005,0),${i+1}),-'Invoice Log'!$U$6:$U$1005,0)),"")`]] });
    data.push({ range: `${S}!K${r}`, values: [[`=IFERROR(INDEX('Invoice Log'!$T$6:$T$1005,MATCH(LARGE(IF(('Invoice Log'!$R$6:$R$1005>0)*('Invoice Log'!$U$6:$U$1005<0),-'Invoice Log'!$U$6:$U$1005,0),${i+1}),-'Invoice Log'!$U$6:$U$1005,0)),"")`]] });
    reqs.push({ repeatCell: { range: gridRange(DB, r - 1, r, 6, 11), cell: { userEnteredFormat: { backgroundColor: i % 2 === 0 ? hex(C.panel) : hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
    reqs.push({ repeatCell: { range: gridRange(DB, r - 1, r, 9, 10), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields: 'userEnteredFormat.numberFormat' } });
  }

  // ── Monthly summary table for charts (rows 53-65) ─────────────────────────
  data.push({ range: `${S}!A53`, values: [['MONTHLY SUMMARY (Chart Data)']] });
  reqs.push({ mergeCells: { range: gridRange(DB, 52, 53, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 52, 53, 0, 12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 10 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  data.push({ range: `${S}!A54:D54`, values: [['Month','Invoiced','Paid','Outstanding']] });
  reqs.push({ repeatCell: { range: gridRange(DB, 53, 54, 0, 4), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  months.forEach((mo, i) => {
    const r = 55 + i;
    data.push({ range: `${S}!A${r}`, values: [[mo]] });
    data.push({ range: `${S}!B${r}`, values: [[`=IFERROR(SUMPRODUCT((YEAR('Invoice Log'!$E$6:$E$1005)=$B$5)*(TEXT('Invoice Log'!$E$6:$E$1005,"mmmm")="${mo}")*('Invoice Log'!$O$6:$O$1005)),0)`]] });
    data.push({ range: `${S}!C${r}`, values: [[`=IFERROR(SUMPRODUCT((YEAR('Invoice Log'!$E$6:$E$1005)=$B$5)*(TEXT('Invoice Log'!$E$6:$E$1005,"mmmm")="${mo}")*('Invoice Log'!$Q$6:$Q$1005)),0)`]] });
    data.push({ range: `${S}!D${r}`, values: [[`=IFERROR(SUMPRODUCT((YEAR('Invoice Log'!$E$6:$E$1005)=$B$5)*(TEXT('Invoice Log'!$E$6:$E$1005,"mmmm")="${mo}")*('Invoice Log'!$R$6:$R$1005)),0)`]] });
    reqs.push({ repeatCell: { range: gridRange(DB, r - 1, r, 0, 4), cell: { userEnteredFormat: {
      backgroundColor: i % 2 === 0 ? hex(C.panel) : hex(C.altRow),
    }}, fields: 'userEnteredFormat.backgroundColor' } });
    reqs.push({ repeatCell: { range: gridRange(DB, r - 1, r, 1, 4), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields: 'userEnteredFormat.numberFormat' } });
  });

  // Invoice Status summary (for chart)
  data.push({ range: `${S}!F54:G54`, values: [['Invoice Status','Count']] });
  reqs.push({ repeatCell: { range: gridRange(DB, 53, 54, 5, 7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  const statuses = ['Draft','Ready to Send','Sent','Viewed','Partially Paid','Paid','Overdue','Cancelled','Refunded'];
  statuses.forEach((st, i) => {
    const r = 55 + i;
    data.push({ range: `${S}!F${r}`, values: [[st]] });
    data.push({ range: `${S}!G${r}`, values: [[`=IFERROR(COUNTIF('Invoice Log'!$T$6:$T$1005,"${st}"),0)`]] });
  });

  // Payment Method summary (for chart)
  data.push({ range: `${S}!I54:J54`, values: [['Payment Method','Net Received']] });
  reqs.push({ repeatCell: { range: gridRange(DB, 53, 54, 8, 10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  const methods = ['Cash','Check','Bank Transfer','ACH','Credit Card','Debit Card','PayPal','Stripe','Venmo','Zelle','Other'];
  methods.forEach((m, i) => {
    const r = 55 + i;
    data.push({ range: `${S}!I${r}`, values: [[m]] });
    data.push({ range: `${S}!J${r}`, values: [[`=IFERROR(SUMIF('Payment Tracker'!$I$6:$I$1505,"${m}",'Payment Tracker'!$N$6:$N$1505),0)`]] });
  });
  reqs.push({ repeatCell: { range: gridRange(DB, 54, 66, 9, 10), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields: 'userEnteredFormat.numberFormat' } });

  // Column widths
  [130,110,110,130,10,130,100,10,130,110].forEach((px, ci) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: DB, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  });

  // Freeze rows 1:5
  reqs.push({ updateSheetProperties: { properties: { sheetId: DB, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  await valuesBatchUpdate(id, data, 'dashboard-values');
  await batchUpdate(id, reqs, 'dashboard-format');
  console.log('✓ Dashboard');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
