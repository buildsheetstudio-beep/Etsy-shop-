const { hex, COLOR, gridRange, batchUpdate, valuesBatchUpdate, vr } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = 128079104; // Research & Reservations

// 6 sample reservations
const reservations = [
  // Type, Name, Destination, CheckIn, CheckOut, BookingRef, Status, Cost, ConfirmationEmail, Website, Notes
  ['Accommodation','Hotel Duquesne Eiffel','Paris','2025-07-15','2025-07-21','BK-2024-PAR-001','Confirmed',975,'paris@duquesneeiffel.com','duquesneeiffel.com','Romantic boutique, near Arc de Triomphe. Breakfast not included.'],
  ['Accommodation','Airbnb Eixample Apartment','Barcelona','2025-07-21','2025-07-29','BK-2024-BCN-001','Confirmed',735,'airbnb@host.com','airbnb.com/rooms/12345','2-bed apartment, Gaudí neighbourhood. Check-in lockbox code: 4782'],
  ['Accommodation','Hotel Artemide','Rome','2025-07-29','2025-08-05','BK-2024-ROM-001','Confirmed',850,'info@artemide.com','hotelartemide.com','4-star near Termini. Breakfast included. Ask for upper floor room.'],
  ['Transport','Eurostar London → Paris (Standard Premier)','Paris','2025-07-15','2025-07-15','ES-7834-LDN-CDG','Confirmed',178,'eurostar@booking.com','eurostar.com','Departs St Pancras 09:01, arrives Gare du Nord 12:17. Car 12, Seats 23A & 23B'],
  ['Transport','Vueling BCN-ROM Flight','Rome','2025-07-21','2025-07-21','VY-1234-BCN-ROM','Confirmed',95,'vueling@booking.com','vueling.com','Departs 14:25, arrives 16:45. Hand luggage only. Check-in online 24hrs before.'],
  ['Activities','Vatican Museums & Sistine Chapel Early Access Tour','Rome','2025-07-30','2025-07-30','VAT-2025-0730','Confirmed',56,'info@vaticantickets.com','vaticantickets.com','07:30 early access, skip the line. 2 tickets @ £28 each. Dress code required.'],
];

(async () => {
  const reqs = [];

  // ── Row 1: Title ──
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 11),
      cell: {
        userEnteredValue: { stringValue: '📋 RESEARCH & RESERVATIONS TRACKER' },
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
  const headers = ['#','Type','Name / Description','Destination','Check-In / Date','Check-Out / End','Booking Ref','Status','Total Cost (£)','Contact / Email','Notes'];
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

  // ── Data rows ──
  const totalRows = 100;
  for (let i = 0; i < totalRows; i++) {
    const row = 2 + i;
    const sample = reservations[i] || null;
    const bg = i % 2 === 0 ? hex(COLOR.white) : hex(COLOR.altRow);

    // # auto
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 0, 1),
        cell: {
          userEnteredValue: { formulaValue: `=IF(C${row+1}="","",ROW()-1)` },
          userEnteredFormat: { backgroundColor: hex(COLOR.champagne), textFormat: { bold: true }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });

    // Input cells
    [1,2,3,4,5,6,7,8,9,10].forEach(c => {
      reqs.push({
        repeatCell: {
          range: gridRange(SID, row, row+1, c, c+1),
          cell: {
            userEnteredFormat: {
              backgroundColor: c === 8 ? hex(COLOR.formulaBg) : bg,
              numberFormat: (c === 4 || c === 5) ? { type: 'DATE', pattern: 'dd mmm yyyy' } :
                            c === 8 ? { type: 'NUMBER', pattern: '£#,##0.00' } : undefined,
              verticalAlignment: 'MIDDLE',
              wrapStrategy: (c === 2 || c === 10) ? 'WRAP' : 'CLIP',
            },
          },
          fields: 'userEnteredFormat',
        },
      });
    });
  }

  // Column widths
  const widths = [35,120,220,100,110,110,130,100,100,170,220];
  widths.forEach((w, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: c, endIndex: c+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 34 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 102 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'reservations-structure');

  // ── Values ──
  const data = [];
  const resRows = reservations.map(r => [r[0],r[1],r[2],r[3],r[4],r[5],r[6],r[7],r[9],r[10]]);
  data.push(vr(`'Research & Reservations'!B3`, resRows));
  await valuesBatchUpdate(id, data, 'reservations-values');

  console.log('Research & Reservations built with', reservations.length, 'bookings');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
