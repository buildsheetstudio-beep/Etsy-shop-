const { hex, COLOR, gridRange, batchUpdate, valuesBatchUpdate, vr } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = 732094577; // Trip Overview

// Sample trip: European Summer Escape
// Jul 15 – Aug 5 2025, 2 travellers, £4,500 budget, GBP
// 3 destinations: Paris (Jul 15-21, 7 nights), Barcelona (Jul 21-29, 8 nights), Rome (Jul 29-Aug 5, 7 nights)

(async () => {
  const reqs = [];

  // ── Row 1: Main title ──
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 12),
      cell: {
        userEnteredValue: { stringValue: '✈ ULTIMATE TRAVEL PLANNER' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.deepNavy),
          textFormat: { foregroundColor: hex(COLOR.antiqueGold), bold: true, fontSize: 20 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // ── Row 2: Subtitle ──
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, 12),
      cell: {
        userEnteredValue: { stringValue: 'TRIP OVERVIEW & SUMMARY' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.midnightBlue),
          textFormat: { foregroundColor: hex(COLOR.champagne), bold: true, fontSize: 13 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // ── Row 3: spacer ──
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 2, 3, 0, 12),
      cell: { userEnteredFormat: { backgroundColor: hex(COLOR.deepNavy) } },
      fields: 'userEnteredFormat.backgroundColor',
    },
  });

  // ── Section A: Trip Details (rows 3–13) ──
  // Left block (A4:F13): Trip Info
  // Right block (H4:L13): Quick Stats (formula-driven)

  const sectionHeader = (text, r, c0, c1, bg = COLOR.midnightBlue, fg = COLOR.antiqueGold) => {
    reqs.push({ mergeCells: { range: gridRange(SID, r, r+1, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r+1, c0, c1),
        cell: {
          userEnteredValue: { stringValue: text },
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(fg), bold: true, fontSize: 11 },
            horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  };

  sectionHeader('📋  TRIP DETAILS', 3, 0, 6);
  sectionHeader('📊  QUICK STATS', 3, 7, 12);

  // Label+input style cell pair helper
  const labelCell = (r, c, text) => reqs.push({
    repeatCell: {
      range: gridRange(SID, r, r+1, c, c+1),
      cell: {
        userEnteredValue: { stringValue: text },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.champagne),
          textFormat: { foregroundColor: hex(COLOR.deepNavy), bold: true, fontSize: 10 },
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  const inputCell = (r, c, merge = 1) => {
    if (merge > 1) reqs.push({ mergeCells: { range: gridRange(SID, r, r+1, c, c+merge), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r+1, c, c+merge),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(COLOR.inputBg),
            borders: {
              bottom: { style: 'SOLID', color: hex(COLOR.antiqueGold) },
              top: { style: 'SOLID', color: hex(COLOR.antiqueGold) },
              left: { style: 'SOLID', color: hex(COLOR.antiqueGold) },
              right: { style: 'SOLID', color: hex(COLOR.antiqueGold) },
            },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat',
      },
    });
  };

  const formulaCell = (r, c, merge = 1) => {
    if (merge > 1) reqs.push({ mergeCells: { range: gridRange(SID, r, r+1, c, c+merge), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r+1, c, c+merge),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(COLOR.formulaBg),
            textFormat: { foregroundColor: hex(COLOR.deepNavy) },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat',
      },
    });
  };

  // Trip Details labels & inputs (rows 4-12)
  // Row 4: Trip Name
  labelCell(4, 0, 'Trip Name');
  inputCell(4, 1, 4); // B4:E4
  // Row 5: Destination(s)
  labelCell(5, 0, 'Destination(s)');
  inputCell(5, 1, 4);
  // Row 6: Start Date / End Date
  labelCell(6, 0, 'Start Date');
  inputCell(6, 1, 1);
  labelCell(6, 3, 'End Date');
  inputCell(6, 4, 1);
  // Row 7: No. Travellers / Currency
  labelCell(7, 0, 'No. Travellers');
  inputCell(7, 1, 1);
  labelCell(7, 3, 'Currency');
  inputCell(7, 4, 1);
  // Row 8: Total Budget
  labelCell(8, 0, 'Total Budget');
  inputCell(8, 1, 1);
  labelCell(8, 3, 'Traveller Type');
  inputCell(8, 4, 1);
  // Row 9: Notes (merged)
  labelCell(9, 0, 'Notes');
  inputCell(9, 1, 4);
  // Row 10: Trip Style
  labelCell(10, 0, 'Trip Style / Theme');
  inputCell(10, 1, 4);
  // Row 11: emergency contact section header
  sectionHeader('🆘  EMERGENCY CONTACTS', 11, 0, 6, COLOR.deepCrimson, COLOR.white);
  // Row 12: Contact headers
  ['Name','Phone','Email','Relation'].forEach((h, ci) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 12, 13, ci === 0 ? 0 : ci+1, ci === 0 ? 1 : ci+2),
        cell: {
          userEnteredValue: { stringValue: h },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.champagne),
            textFormat: { foregroundColor: hex(COLOR.deepNavy), bold: true, fontSize: 9 },
            horizontalAlignment: 'CENTER',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });
  // Rows 13-15: Contact input rows
  [13, 14, 15].forEach(r => {
    for (let c = 0; c < 6; c++) {
      inputCell(r, c, 1);
    }
  });

  // ── Quick Stats (right side, rows 4-10) ──
  const statLabel = (r, text) => {
    reqs.push({ mergeCells: { range: gridRange(SID, r, r+1, 7, 9), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r+1, 7, 9),
        cell: {
          userEnteredValue: { stringValue: text },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.champagne),
            textFormat: { foregroundColor: hex(COLOR.deepNavy), bold: true, fontSize: 10 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  };
  const statValue = (r, formula, isBig = false) => {
    reqs.push({ mergeCells: { range: gridRange(SID, r, r+1, 9, 12), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r+1, 9, 12),
        cell: {
          userEnteredValue: { formulaValue: formula },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.formulaBg),
            textFormat: { foregroundColor: hex(COLOR.midnightBlue), bold: isBig, fontSize: isBig ? 14 : 11 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  };

  statLabel(4, '📅 Trip Duration');
  statValue(4, '=IFERROR(TEXT(C7-B7+1,"0")&" days","—")', true);
  statLabel(5, '👥 No. of Travellers');
  statValue(5, '=IFERROR(TEXT(B8,"0")&" travellers","—")');
  statLabel(6, '💰 Total Budget');
  statValue(6, '=IFERROR(TEXT(B9,"£#,##0"),"—")', true);
  statLabel(7, '💸 Budget Per Person');
  statValue(7, '=IFERROR(TEXT(B9/B8,"£#,##0"),"—")');
  statLabel(8, '💸 Budget Per Day');
  statValue(8, '=IFERROR(TEXT(B9/(C7-B7+1),"£#,##0.00"),"—")');
  statLabel(9, '📍 Destinations');
  statValue(9, '=IFERROR(COUNTA(\'Day-by-Day Itinerary\'!E:E)-1&" stops","—")');
  statLabel(10, '💳 Total Spent');
  statValue(10, '=IFERROR(TEXT(SUM(\'Travel Budget\'!E:E),"£#,##0.00"),"—")', true);

  // ── Section B: Destinations Table (rows 17-27) ──
  sectionHeader('📍  DESTINATIONS AT A GLANCE', 17, 0, 12);

  const destHeaders = ['#','City / Region','Country','Arrival Date','Departure Date','Nights','Accommodation','Booking Ref','Notes'];
  destHeaders.forEach((h, c) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 18, 19, c, c+1),
        cell: {
          userEnteredValue: { stringValue: h },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.midnightBlue),
            textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 9 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });

  // 5 destination data rows (rows 19-23) – 3 populated, 2 blank
  const destColors = [hex(COLOR.white), hex(COLOR.altRow), hex(COLOR.white), hex(COLOR.altRow), hex(COLOR.white)];
  destColors.forEach((bg, i) => {
    const row = 19 + i;
    // # column = formula
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 0, 1),
        cell: {
          userEnteredValue: { formulaValue: `=IF(B${row+1}="","",ROW()-18)` },
          userEnteredFormat: { backgroundColor: bg, textFormat: { bold: true }, horizontalAlignment: 'CENTER' },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
    for (let c = 1; c < 9; c++) {
      const isDate = c === 3 || c === 4;
      const isFormula = c === 5;
      reqs.push({
        repeatCell: {
          range: gridRange(SID, row, row+1, c, c+1),
          cell: {
            userEnteredFormat: {
              backgroundColor: isFormula ? hex(COLOR.formulaBg) : bg,
              numberFormat: isDate ? { type: 'DATE', pattern: 'dd mmm yyyy' } : undefined,
            },
          },
          fields: isDate ? 'userEnteredFormat.backgroundColor,userEnteredFormat.numberFormat' : 'userEnteredFormat.backgroundColor',
        },
      });
    }
    // Nights formula (col E = 5)
    reqs.push({
      updateCells: {
        range: gridRange(SID, row, row+1, 5, 6),
        rows: [{ values: [{ userEnteredValue: { formulaValue: `=IF(D${row+1}="","",E${row+1}-D${row+1})` } }] }],
        fields: 'userEnteredValue',
      },
    });
  });

  // ── Section C: Trip Progress Bar (rows 25-27) ──
  sectionHeader('⏱  TRIP PROGRESS', 25, 0, 12, COLOR.forestGreen, COLOR.white);

  reqs.push({ mergeCells: { range: gridRange(SID, 26, 27, 0, 3), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 26, 27, 0, 3),
      cell: {
        userEnteredValue: { formulaValue: '=IFERROR(IF(TODAY()<B7,"Trip hasn\'t started yet",IF(TODAY()>C7,"Trip completed! 🎉","Day "&(TODAY()-B7+1)&" of "&(C7-B7+1))),"Set dates above")' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.lightGold),
          textFormat: { foregroundColor: hex(COLOR.deepNavy), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  reqs.push({ mergeCells: { range: gridRange(SID, 26, 27, 3, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 26, 27, 3, 12),
      cell: {
        userEnteredValue: { formulaValue: '=IFERROR(REPT("▰",ROUND(MIN((TODAY()-B7)/(C7-B7)*10,10),0))&REPT("▱",10-ROUND(MIN((TODAY()-B7)/(C7-B7)*10,10),0))&"  "&TEXT(MIN((TODAY()-B7)/(C7-B7)*100,100),"0.0")&"% complete","——————————————  0.0% complete")' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.paleNavy),
          textFormat: { foregroundColor: hex(COLOR.midnightBlue), bold: true, fontSize: 12, fontFamily: 'Courier New' },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // Column widths for Trip Overview
  const colWidths = [90,140,120,120,120,80,160,120,180,120,120,120];
  colWidths.forEach((w, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: c, endIndex: c+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 54 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 28 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'trip-overview-structure');

  // ── Values ──
  const data = [];
  // Trip details
  data.push(vr(`'Trip Overview'!B5`, [['European Summer Escape']]));
  data.push(vr(`'Trip Overview'!B6`, [['Paris, Barcelona & Rome']]));
  data.push(vr(`'Trip Overview'!B7`, [['2025-07-15']]));
  data.push(vr(`'Trip Overview'!C7`, [['2025-08-05']]));  // End date in C7 (col C = index 2)
  data.push(vr(`'Trip Overview'!B8`, [[2]]));             // Travellers
  data.push(vr(`'Trip Overview'!E8`, [['GBP']]));        // Currency
  data.push(vr(`'Trip Overview'!B9`, [[4500]]));          // Budget
  data.push(vr(`'Trip Overview'!E9`, [['Couple']]));     // Traveller type
  data.push(vr(`'Trip Overview'!B10`, [['Romantic summer escape, mix of culture, food & history']]));
  data.push(vr(`'Trip Overview'!B11`, [['Two weeks+ in Europe — art, food, architecture & sunshine!']]));

  // Emergency contacts
  data.push(vr(`'Trip Overview'!A13`, [
    ['Sarah Johnson (Sister)','+44 7700 900123','sarah.j@email.com','Sister'],
    ['Dr. Michael Scott','0800 111 999','','GP (NHS Direct)'],
    ['Travel Insurance','0800 123 4567','claims@insuranceco.com','Insurance Co.'],
  ]));

  // Destinations table
  data.push(vr(`'Trip Overview'!B20`, [
    ['Paris','France','2025-07-15','2025-07-21','','Hotel Duquesne Eiffel','BK-2024-PAR-001','Romantic boutique hotel, Arc de Triomphe district'],
    ['Barcelona','Spain','2025-07-21','2025-07-29','','Airbnb Eixample','BK-2024-BCN-001','2-bed apt, Gaudí architecture nearby'],
    ['Rome','Italy','2025-07-29','2025-08-05','','Hotel Artemide','BK-2024-ROM-001','4-star near Termini station'],
  ]));

  await valuesBatchUpdate(id, data, 'trip-overview-values');

  // Date formats for Trip Details
  const dateReqs = [];
  // B7 and C7 (start/end dates) — row 6 (0-indexed)
  ['B7','C7'].forEach((cell, i) => {
    dateReqs.push({
      repeatCell: {
        range: gridRange(SID, 6, 7, i === 0 ? 1 : 4, i === 0 ? 2 : 5),
        cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } } },
        fields: 'userEnteredFormat.numberFormat',
      },
    });
  });

  await batchUpdate(id, dateReqs, 'trip-overview-dates');

  console.log('Trip Overview built with European Summer Escape sample data');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
