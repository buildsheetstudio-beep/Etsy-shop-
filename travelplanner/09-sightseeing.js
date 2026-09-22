const { hex, COLOR, gridRange, batchUpdate, valuesBatchUpdate, vr } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = 2144247023; // Sightseeing & Activities

// 8 sample activities
const activities = [
  // Name, Type, Destination, Date, Duration(hrs), Cost(£), Rating(1-5), Booked?, MustSee?, Notes
  ['Louvre Museum','Museum / Gallery','Paris','2025-07-16',3,12.50,5,true,true,'Book online to skip queues. Go early! Mona Lisa gets crowded by 10am.'],
  ['Eiffel Tower (Summit)','Museum / Gallery','Paris','2025-07-17',2,15.00,5,true,true,'Evening slot (8pm) for stunning sunset views. Champagne bar at top!'],
  ['Palace of Versailles','Guided Tour','Paris','2025-07-18',6,25.00,5,true,true,'Take RER C from Champs de Mars. Arrive early – huge queues otherwise.'],
  ['Sacré-Cœur & Montmartre','Cultural Experience','Paris','2025-07-17',2,0,4,false,true,'Free entry to basilica. Explore the artists\' square nearby. Great views over Paris.'],
  ['Sagrada Família (Towers)','Museum / Gallery','Barcelona','2025-07-23',2.5,22.00,5,true,true,'Book 2+ weeks ahead. Morning light through stained glass is magical.'],
  ['Park Güell','Cultural Experience','Barcelona','2025-07-24',2,10.00,4,true,false,'Timed entry to main monument zone. Arrive 15 mins early.'],
  ['Vatican Museums & Sistine Chapel','Museum / Gallery','Rome','2025-07-30',4,28.00,5,true,true,'Early access 7:30am – nearly empty! Absolutely unmissable.'],
  ['Colosseum & Roman Forum','Guided Tour','Rome','2025-07-31',3.5,18.00,5,true,true,'Arena floor access worth the upgrade. Book guided tour for context.'],
];

(async () => {
  const reqs = [];

  // ── Row 1: Title ──
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 12),
      cell: {
        userEnteredValue: { stringValue: '🏛 SIGHTSEEING & ACTIVITIES PLANNER' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.deepNavy),
          textFormat: { foregroundColor: hex(COLOR.antiqueGold), bold: true, fontSize: 18 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // ── Row 2: Column headers ──
  const headers = ['#','Activity Name','Type','Destination','Date','Duration (hrs)','Cost (£)','Rating','Booked?','Must See?','Notes'];
  headers.forEach((h, c) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 1, 2, c, c+1),
        cell: {
          userEnteredValue: { stringValue: h },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.midnightBlue),
            textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 9 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
            wrapStrategy: 'WRAP',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });

  // Helper chart data (cols M-N) - destination activity count
  ['Destination','# Activities'].forEach((h, ci) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 1, 2, 12+ci, 13+ci),
        cell: {
          userEnteredValue: { stringValue: h },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.midnightBlue),
            textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 9 },
            horizontalAlignment: 'CENTER',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });

  // ── Data rows ──
  const totalRows = 100;
  for (let i = 0; i < totalRows; i++) {
    const row = 2 + i;
    const bg = i % 2 === 0 ? hex(COLOR.white) : hex(COLOR.altRow);

    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 0, 1),
        cell: {
          userEnteredValue: { formulaValue: `=IF(B${row+1}="","",ROW()-1)` },
          userEnteredFormat: { backgroundColor: hex(COLOR.champagne), textFormat: { bold: true }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });

    [1,2,3,4,5,6,7,8,9,10].forEach(c => {
      reqs.push({
        repeatCell: {
          range: gridRange(SID, row, row+1, c, c+1),
          cell: {
            userEnteredFormat: {
              backgroundColor: bg,
              numberFormat: c === 4 ? { type: 'DATE', pattern: 'dd mmm yyyy' } :
                            (c === 5 || c === 6) ? { type: 'NUMBER', pattern: c === 6 ? '£#,##0.00' : '0.0' } : undefined,
              horizontalAlignment: (c === 5 || c === 6 || c === 7 || c === 8 || c === 9) ? 'CENTER' : 'LEFT',
              verticalAlignment: 'MIDDLE',
              wrapStrategy: (c === 1 || c === 10) ? 'WRAP' : 'CLIP',
            },
          },
          fields: 'userEnteredFormat',
        },
      });
    });
  }

  // ── Chart helper (col M-N): dest activity counts ──
  const dests = ['Paris','Barcelona','Rome'];
  dests.forEach((d, i) => {
    const row = 2 + i;
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 12, 14),
        cell: {
          userEnteredFormat: { backgroundColor: hex(COLOR.altRow), verticalAlignment: 'MIDDLE' },
        },
        fields: 'userEnteredFormat',
      },
    });
  });

  // Column widths
  const widths = [35,200,140,100,100,80,80,80,75,75,220];
  widths.forEach((w, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: c, endIndex: c+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  // Helper cols M,N
  [120,100].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 12+i, endIndex: 13+i }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 34 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 50 }, properties: { pixelSize: 55 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 50, endIndex: 102 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'sightseeing-structure');

  // ── Values ──
  const data = [];
  const actRows = activities.map(a => [a[0],a[1],a[2],a[3],a[4],a[5],a[6],a[7],a[8],a[9]]);
  data.push(vr(`'Sightseeing & Activities'!B3`, actRows));

  // Destination chart helper
  const destRows = dests.map(d => [d, `=COUNTIF(D$3:D$102,"${d}")`]);
  data.push(vr(`'Sightseeing & Activities'!M3`, destRows));

  await valuesBatchUpdate(id, data, 'sightseeing-values');

  console.log('Sightseeing & Activities built with', activities.length, 'activities');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
