const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SIDs = {
  tripOverview: 732094577,
  itinerary: 353766069,
  budget: 1601326951,
  costPerPerson: 1289386811,
  currency: 1614915889,
  reservations: 128079104,
  sightseeing: 2144247023,
  foodDiary: 1291844518,
  visa: 1488564375,
  packing: 1128916394,
  reference: 2028483192,
};

(async () => {
  const reqs = [];

  // ── Chart 1: Budget Breakdown Donut (Travel Budget tab) ──
  // Data: N4:O12 — Category | Actual spend
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Spend by Category',
          titleTextFormat: { fontFamily: 'Georgia', bold: true, fontSize: 13 },
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            pieHole: 0.5,
            series: {
              sourceRange: {
                sources: [{
                  sheetId: SIDs.budget,
                  startRowIndex: 3, endRowIndex: 12,
                  startColumnIndex: 15, endColumnIndex: 16, // P4:P12
                }],
              },
            },
            domain: {
              sourceRange: {
                sources: [{
                  sheetId: SIDs.budget,
                  startRowIndex: 3, endRowIndex: 12,
                  startColumnIndex: 13, endColumnIndex: 14, // N4:N12
                }],
              },
            },
          },
          backgroundColor: { red: 0.941, green: 0.929, blue: 0.835 }, // champagne
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: SIDs.budget, rowIndex: 2, columnIndex: 9 },
            offsetXPixels: 0,
            offsetYPixels: 0,
            widthPixels: 420,
            heightPixels: 300,
          },
        },
      },
    },
  });

  // ── Chart 2: Sightseeing Bar Chart (Sightseeing & Activities tab) ──
  // Horizontal bar: activities by destination from M3:N5
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Activities by Destination',
          titleTextFormat: { fontFamily: 'Georgia', bold: true, fontSize: 13 },
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'No. of Activities' },
              { position: 'LEFT_AXIS', title: 'Destination' },
            ],
            domains: [{
              domain: {
                sourceRange: {
                  sources: [{
                    sheetId: SIDs.sightseeing,
                    startRowIndex: 2, endRowIndex: 5,
                    startColumnIndex: 12, endColumnIndex: 13, // M3:M5
                  }],
                },
              },
            }],
            series: [{
              series: {
                sourceRange: {
                  sources: [{
                    sheetId: SIDs.sightseeing,
                    startRowIndex: 2, endRowIndex: 5,
                    startColumnIndex: 13, endColumnIndex: 14, // N3:N5
                  }],
                },
              },
              targetAxis: 'BOTTOM_AXIS',
              color: { red: 0.176, green: 0.416, blue: 0.310 }, // forest green
            }],
          },
          backgroundColor: { red: 0.941, green: 0.929, blue: 0.835 },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: SIDs.sightseeing, rowIndex: 1, columnIndex: 12 },
            offsetXPixels: 0,
            offsetYPixels: 40,
            widthPixels: 380,
            heightPixels: 260,
          },
        },
      },
    },
  });

  // ── Chart 3: Food Diary Column Chart (Restaurant & Food Diary tab) ──
  // Column chart: spend by meal type from M3:O7
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Spend by Meal Type',
          titleTextFormat: { fontFamily: 'Georgia', bold: true, fontSize: 13 },
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Meal Type' },
              { position: 'LEFT_AXIS', title: 'Total Spend (£)' },
            ],
            domains: [{
              domain: {
                sourceRange: {
                  sources: [{
                    sheetId: SIDs.foodDiary,
                    startRowIndex: 2, endRowIndex: 7,
                    startColumnIndex: 12, endColumnIndex: 13, // M3:M7
                  }],
                },
              },
            }],
            series: [{
              series: {
                sourceRange: {
                  sources: [{
                    sheetId: SIDs.foodDiary,
                    startRowIndex: 2, endRowIndex: 7,
                    startColumnIndex: 14, endColumnIndex: 15, // O3:O7
                  }],
                },
              },
              targetAxis: 'LEFT_AXIS',
              color: { red: 0.545, green: 0.102, blue: 0.102 }, // deep crimson
            }],
          },
          backgroundColor: { red: 0.941, green: 0.929, blue: 0.835 },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: SIDs.foodDiary, rowIndex: 1, columnIndex: 12 },
            offsetXPixels: 0,
            offsetYPixels: 40,
            widthPixels: 380,
            heightPixels: 260,
          },
        },
      },
    },
  });

  await batchUpdate(id, reqs, 'charts');
  console.log('3 charts added: Budget donut, Sightseeing bar, Food Diary column');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
