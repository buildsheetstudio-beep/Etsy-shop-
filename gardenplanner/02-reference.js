'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const REF = sheetMap['Reference Data'];

(async () => {
  const values = [];

  // ── Column A: Plant Types ─────────────────────────────────────────────────
  values.push({ range: 'Reference Data!A1', values: [['Plant Type']] });
  values.push({ range: 'Reference Data!A2', values: [
    ['Vegetable'],['Herb'],['Fruit'],['Flower'],['Root Vegetable'],
    ['Leafy Green'],['Vine'],['Perennial'],
  ]});

  // ── Column B: Sun Requirements ────────────────────────────────────────────
  values.push({ range: 'Reference Data!B1', values: [['Sun Requirement']] });
  values.push({ range: 'Reference Data!B2', values: [
    ['Full Sun'],['Partial Sun'],['Partial Shade'],['Full Shade'],
  ]});

  // ── Column C: Water Needs ─────────────────────────────────────────────────
  values.push({ range: 'Reference Data!C1', values: [['Water Needs']] });
  values.push({ range: 'Reference Data!C2', values: [
    ['Low'],['Moderate'],['High'],['Very High'],
  ]});

  // ── Column D: Plant Status ────────────────────────────────────────────────
  values.push({ range: 'Reference Data!D1', values: [['Plant Status']] });
  values.push({ range: 'Reference Data!D2', values: [
    ['Planned'],['Seeded'],['Germinated'],['Transplanted'],['Growing'],
    ['Flowering'],['Harvest Ready'],['Harvested'],['Dormant'],['Failed'],
  ]});

  // ── Column E: Season ─────────────────────────────────────────────────────
  values.push({ range: 'Reference Data!E1', values: [['Season']] });
  values.push({ range: 'Reference Data!E2', values: [
    ['Spring'],['Summer'],['Fall'],['Winter'],['All Season'],['Perennial'],
  ]});

  // ── Column F: Difficulty ─────────────────────────────────────────────────
  values.push({ range: 'Reference Data!F1', values: [['Difficulty']] });
  values.push({ range: 'Reference Data!F2', values: [
    ['Beginner'],['Intermediate'],['Advanced'],
  ]});

  // ── Column G: Harvest Unit ────────────────────────────────────────────────
  values.push({ range: 'Reference Data!G1', values: [['Harvest Unit']] });
  values.push({ range: 'Reference Data!G2', values: [
    ['lbs'],['oz'],['count'],['bunch'],['cloves'],['heads'],['ears'],['quarts'],
  ]});

  // ── Column H: Budget Category ─────────────────────────────────────────────
  values.push({ range: 'Reference Data!H1', values: [['Budget Category']] });
  values.push({ range: 'Reference Data!H2', values: [
    ['Seeds & Starts'],['Soil & Amendments'],['Tools & Equipment'],
    ['Irrigation'],['Pest Control'],['Structures & Supports'],
    ['Fertilizer'],['Other'],
  ]});

  // ── Column I: Pest/Disease Type ───────────────────────────────────────────
  values.push({ range: 'Reference Data!I1', values: [['Pest/Disease Type']] });
  values.push({ range: 'Reference Data!I2', values: [
    ['Insect Pest'],['Fungal Disease'],['Bacterial Disease'],
    ['Viral Disease'],['Animal Pest'],['Nutrient Deficiency'],['Other'],
  ]});

  // ── Column J: Severity ────────────────────────────────────────────────────
  values.push({ range: 'Reference Data!J1', values: [['Severity']] });
  values.push({ range: 'Reference Data!J2', values: [
    ['Low'],['Medium'],['High'],['Critical'],
  ]});

  // ── Column K: Treatment Status ────────────────────────────────────────────
  values.push({ range: 'Reference Data!K1', values: [['Treatment Status']] });
  values.push({ range: 'Reference Data!K2', values: [
    ['Monitoring'],['Treating'],['Resolved'],['Recurring'],
  ]});

  // ── Column L: Seed Source ─────────────────────────────────────────────────
  values.push({ range: 'Reference Data!L1', values: [['Seed Source']] });
  values.push({ range: 'Reference Data!L2', values: [
    ['Saved Seed'],['Purchased'],['Gift/Trade'],['Unknown'],
  ]});

  // ── Column M: Months ─────────────────────────────────────────────────────
  values.push({ range: 'Reference Data!M1', values: [['Month']] });
  values.push({ range: 'Reference Data!M2', values: [
    ['January'],['February'],['March'],['April'],['May'],['June'],
    ['July'],['August'],['September'],['October'],['November'],['December'],
  ]});

  // ── Column N: Watering Frequency ─────────────────────────────────────────
  values.push({ range: 'Reference Data!N1', values: [['Watering Frequency']] });
  values.push({ range: 'Reference Data!N2', values: [
    ['Daily'],['Every 2 Days'],['Every 3 Days'],['Twice a Week'],
    ['Weekly'],['Bi-weekly'],['As Needed'],
  ]});

  // ── Column O: Zone ────────────────────────────────────────────────────────
  values.push({ range: 'Reference Data!O1', values: [['Garden Zone']] });
  values.push({ range: 'Reference Data!O2', values: [
    ['Zone A — Herbs'],['Zone B — Tomatoes'],['Zone C — Root Veg'],
    ['Zone D — Greens'],['Zone E — Cucurbits'],['Zone F — Flowers'],
    ['Zone G — Perennials'],['Containers'],['Raised Bed'],
  ]});

  // ── Columns P–Q: Price Lookup Table ──────────────────────────────────────
  values.push({ range: 'Reference Data!P1', values: [['Category']] });
  values.push({ range: 'Reference Data!Q1', values: [['Avg Cost ($)']] });
  const priceData = [
    ['Seeds & Starts', 4.99],
    ['Soil & Amendments', 18.99],
    ['Tools & Equipment', 29.99],
    ['Irrigation', 24.99],
    ['Pest Control', 12.99],
    ['Structures & Supports', 19.99],
    ['Fertilizer', 14.99],
    ['Other', 9.99],
  ];
  values.push({ range: 'Reference Data!P2', values: priceData.map(r => [r[0]]) });
  values.push({ range: 'Reference Data!Q2', values: priceData.map(r => [r[1]]) });

  await valuesBatchUpdate(id, values, 'reference-values');

  // ── Format headers ────────────────────────────────────────────────────────
  const reqs = [];
  const headerCols = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16];
  headerCols.forEach(c => {
    reqs.push({
      repeatCell: {
        range: gridRange(REF, 0, 1, c, c+1),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.olive),
            textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
      },
    });
  });

  // Auto-resize columns
  reqs.push({
    autoResizeDimensions: {
      dimensions: { sheetId: REF, dimension: 'COLUMNS', startIndex: 0, endIndex: 18 },
    },
  });

  await batchUpdate(id, reqs, 'reference-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
