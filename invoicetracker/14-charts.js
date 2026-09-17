'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DB = sheetMap['Dashboard'];
const S  = "'Dashboard'";

(async () => {
  const data = [];
  const reqs = [];

  // ── Client Revenue chart data (K54:L64) ───────────────────────────────────
  const topClients = [
    'Apex Digital Agency', 'Harbor Financial Group', 'Driftwood Publishing',
    'Luminary Education', 'Northgate Consulting', 'BlueSky Architecture',
    'Cascade Wellness Spa', 'Gilded Thread Boutique', 'Juniper Properties',
    'Emerald Coast Realty',
  ];

  data.push({ range: `${S}!K54:L54`, values: [['Client', 'Total Invoiced']] });
  reqs.push({ repeatCell: { range: gridRange(DB, 53, 54, 10, 12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  topClients.forEach((client, i) => {
    const r = 55 + i;
    data.push({ range: `${S}!K${r}`, values: [[client]] });
    data.push({ range: `${S}!L${r}`, values: [[`=IFERROR(SUMPRODUCT((YEAR('Invoice Log'!$E$6:$E$1005)=$B$5)*('Invoice Log'!$D$6:$D$1005="${client}")*('Invoice Log'!$O$6:$O$1005)),0)`]] });
  });
  reqs.push({ repeatCell: { range: gridRange(DB, 54, 65, 11, 12), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat.numberFormat' } });

  // ── Charts section header at row 67 ───────────────────────────────────────
  data.push({ range: `${S}!A67`, values: [['CHARTS & ANALYTICS']] });
  reqs.push({ mergeCells: { range: gridRange(DB, 66, 67, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(DB, 66, 67, 0, 12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 11 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  await valuesBatchUpdate(id, data, 'chart-data');
  await batchUpdate(id, reqs, 'chart-format');

  // ── Add Charts ────────────────────────────────────────────────────────────
  const chartReqs = [];

  function chartSrc(r1, r2, c1, c2) {
    return { sourceRange: { sources: [gridRange(DB, r1, r2, c1, c2)] } };
  }

  // 1. Monthly Invoiced vs Paid (COLUMN) — left, row 68
  chartReqs.push({ addChart: { chart: {
    spec: {
      title: 'Monthly Invoiced vs Paid',
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'BOTTOM_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Month' },
          { position: 'LEFT_AXIS',   title: 'Amount ($)' },
        ],
        domains: [{ domain: chartSrc(53, 66, 0, 1) }],
        series: [
          { series: chartSrc(53, 66, 1, 2), targetAxis: 'LEFT_AXIS', color: hex(C.primary) },
          { series: chartSrc(53, 66, 2, 3), targetAxis: 'LEFT_AXIS', color: hex(C.secondary) },
        ],
        headerCount: 1,
      },
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: DB, rowIndex: 67, columnIndex: 0 },
      offsetXPixels: 0, offsetYPixels: 0, widthPixels: 480, heightPixels: 270,
    }},
  }}});

  // 2. Monthly Outstanding Balance (LINE) — right, row 68
  chartReqs.push({ addChart: { chart: {
    spec: {
      title: 'Monthly Outstanding Balance',
      basicChart: {
        chartType: 'LINE',
        legendPosition: 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Month' },
          { position: 'LEFT_AXIS',   title: 'Outstanding ($)' },
        ],
        domains: [{ domain: chartSrc(53, 66, 0, 1) }],
        series: [
          { series: chartSrc(53, 66, 3, 4), targetAxis: 'LEFT_AXIS', color: hex(C.attention) },
        ],
        headerCount: 1,
      },
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: DB, rowIndex: 67, columnIndex: 6 },
      offsetXPixels: 0, offsetYPixels: 0, widthPixels: 480, heightPixels: 270,
    }},
  }}});

  // 3. Invoice Status Distribution (PIE) — left, row 82
  chartReqs.push({ addChart: { chart: {
    spec: {
      title: 'Invoice Status Distribution',
      pieChart: {
        legendPosition: 'LABELED_LEGEND',
        domain: chartSrc(53, 63, 5, 6),
        series: chartSrc(53, 63, 6, 7),
      },
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: DB, rowIndex: 81, columnIndex: 0 },
      offsetXPixels: 0, offsetYPixels: 0, widthPixels: 480, heightPixels: 270,
    }},
  }}});

  // 4. Payments by Method (BAR) — right, row 82
  chartReqs.push({ addChart: { chart: {
    spec: {
      title: 'Payments by Method',
      basicChart: {
        chartType: 'BAR',
        legendPosition: 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Net Received ($)' },
          { position: 'LEFT_AXIS',   title: 'Payment Method' },
        ],
        domains: [{ domain: chartSrc(53, 65, 8, 9) }],
        series: [
          { series: chartSrc(53, 65, 9, 10), targetAxis: 'BOTTOM_AXIS', color: hex(C.secondary) },
        ],
        headerCount: 1,
      },
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: DB, rowIndex: 81, columnIndex: 6 },
      offsetXPixels: 0, offsetYPixels: 0, widthPixels: 480, heightPixels: 270,
    }},
  }}});

  // 5. Outstanding Balance by Aging Bucket (COLUMN) — left, row 95
  // Data lives in Dashboard A20:C25 (header row 20, 5 buckets rows 21-25)
  chartReqs.push({ addChart: { chart: {
    spec: {
      title: 'Outstanding Balance by Aging Bucket',
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Aging Bucket' },
          { position: 'LEFT_AXIS',   title: 'Balance ($)' },
        ],
        domains: [{ domain: chartSrc(19, 25, 0, 1) }],
        series: [
          { series: chartSrc(19, 25, 2, 3), targetAxis: 'LEFT_AXIS', color: hex(C.attention) },
        ],
        headerCount: 1,
      },
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: DB, rowIndex: 94, columnIndex: 0 },
      offsetXPixels: 0, offsetYPixels: 0, widthPixels: 480, heightPixels: 270,
    }},
  }}});

  // 6. Revenue by Client (BAR) — right, row 95
  chartReqs.push({ addChart: { chart: {
    spec: {
      title: 'Revenue by Client',
      basicChart: {
        chartType: 'BAR',
        legendPosition: 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Total Invoiced ($)' },
          { position: 'LEFT_AXIS',   title: 'Client' },
        ],
        domains: [{ domain: chartSrc(53, 64, 10, 11) }],
        series: [
          { series: chartSrc(53, 64, 11, 12), targetAxis: 'BOTTOM_AXIS', color: hex(C.primary) },
        ],
        headerCount: 1,
      },
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: DB, rowIndex: 94, columnIndex: 6 },
      offsetXPixels: 0, offsetYPixels: 0, widthPixels: 480, heightPixels: 270,
    }},
  }}});

  await batchUpdate(id, chartReqs, 'charts');
  console.log('✓ Charts');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
