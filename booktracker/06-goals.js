'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Goals & Challenges'];
const S = "'Goals & Challenges'";
const LIB = "'Master Book Library'";

(async () => {
  const fmt  = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,300,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 10, fontFamily: 'Georgia', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Title ─────────────────────────────────────────────────────────────────
  vals.push({ range: `${S}!A1`, values: [['🎯 Goals & Challenges']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.accent), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 42 }, fields: 'pixelSize' }});

  // Subtitle
  vals.push({ range: `${S}!A2`, values: [['Set your annual reading goals, track reading challenges, and celebrate milestones. Progress updates automatically from your library.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.greenTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // ── Section 1: Annual Reading Goals ──────────────────────────────────────
  const SEC1 = 4;
  vals.push({ range: `${S}!A${SEC1}`, values: [['SECTION 1 — ANNUAL READING GOALS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,SEC1-1,SEC1,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,SEC1-1,SEC1,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: SEC1-1, endIndex: SEC1 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  const goalsHdr = ['Year','Goal Type','Target','Books Read (Auto)','Pages Read (Auto)','% Complete (Auto)','Status','Notes'];
  vals.push({ range: `${S}!A${SEC1+1}`, values: [goalsHdr] });
  fmt.push({ repeatCell: { range: gridRange(SID,SEC1,SEC1+1,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: SEC1, endIndex: SEC1+1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' }});

  // Annual goals data rows — Books Read and Pages Read auto-calculated
  // SUMPRODUCT for year-based counting (not COUNTIFS(YEAR(...)) which causes #VALUE)
  const ANNUAL_GOALS = [
    [2023,'Annual Book Count',48,null,null,null,'Achieved','Finished 50 books — exceeded the goal!'],
    [2024,'Annual Book Count',55,null,null,null,'Achieved','Finished 57 books. A new personal record.'],
    [2025,'Annual Book Count',52,null,null,null,'Achieved','Finished 54 books. Strong year.'],
    [2026,'Annual Book Count',50,null,null,null,'In Progress','On track — aiming to hit 50 by December.'],
    [2023,'Page Count Goal',15000,null,null,null,'Achieved','Read 16,842 pages in 2023.'],
    [2024,'Page Count Goal',18000,null,null,null,'Achieved','Read 19,340 pages in 2024.'],
    [2025,'Page Count Goal',17000,null,null,null,'Achieved','Read 18,200 pages in 2025.'],
    [2026,'Page Count Goal',15000,null,null,null,'In Progress','Currently at ~7,500 pages halfway through the year.'],
  ];

  ANNUAL_GOALS.forEach((row, ri) => {
    const dataRow = SEC1 + 2 + ri;
    // Year, GoalType, Target in cols A,B,C
    vals.push({ range: `${S}!A${dataRow}:C${dataRow}`, values: [[row[0], row[1], row[2]]] });
    vals.push({ range: `${S}!G${dataRow}:H${dataRow}`, values: [[row[6], row[7]]] });
    // Auto-calculated: D=books read that year, E=pages read that year, F=% complete
    const yr = row[0];
    fmt.push({ repeatCell: {
      range: gridRange(SID,dataRow-1,dataRow,3,4),
      cell: { userEnteredValue: { formulaValue: `=SUMPRODUCT((YEAR(IFERROR(DATEVALUE(${LIB}!$O$8:$O$1008),0))=${yr})*(${LIB}!$M$8:$M$1008="Finished"))` },
        userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }},
      fields: 'userEnteredValue,userEnteredFormat',
    }});
    fmt.push({ repeatCell: {
      range: gridRange(SID,dataRow-1,dataRow,4,5),
      cell: { userEnteredValue: { formulaValue: `=SUMPRODUCT((YEAR(IFERROR(DATEVALUE(${LIB}!$O$8:$O$1008),0))=${yr})*(${LIB}!$M$8:$M$1008="Finished")*${LIB}!$Q$8:$Q$1008)` },
        userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER',
          numberFormat: { type: 'NUMBER', pattern: '#,##0' } }},
      fields: 'userEnteredValue,userEnteredFormat',
    }});
    fmt.push({ repeatCell: {
      range: gridRange(SID,dataRow-1,dataRow,5,6),
      cell: { userEnteredValue: { formulaValue: `=IFERROR(D${dataRow}/C${dataRow},0)` },
        userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER',
          numberFormat: { type: 'PERCENT', pattern: '0%' } }},
      fields: 'userEnteredValue,userEnteredFormat',
    }});
    const bg = ri % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,dataRow-1,dataRow,0,8), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    [0,1,2,6,7].forEach(ci => {
      fmt.push({ repeatCell: { range: gridRange(SID,dataRow-1,dataRow,ci,ci+1), cell: { userEnteredFormat: {
        backgroundColor: hex(C.input),
      }}, fields: 'userEnteredFormat.backgroundColor' }});
    });
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: dataRow-1, endIndex: dataRow }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});
  });

  // ── Section 2: Reading Challenges ─────────────────────────────────────────
  const SEC2 = SEC1 + 2 + ANNUAL_GOALS.length + 2;
  vals.push({ range: `${S}!A${SEC2}`, values: [['SECTION 2 — READING CHALLENGES']] });
  fmt.push({ mergeCells: { range: gridRange(SID,SEC2-1,SEC2,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,SEC2-1,SEC2,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: SEC2-1, endIndex: SEC2 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  const chalHdr = ['Year','Challenge Name','Challenge Type','Target (books)','Completed','% Done','Status','Description','Notes'];
  vals.push({ range: `${S}!A${SEC2+1}`, values: [chalHdr] });
  fmt.push({ repeatCell: { range: gridRange(SID,SEC2,SEC2+1,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: SEC2, endIndex: SEC2+1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' }});

  const CHALLENGES = [
    [2023,'Genre Explorer 2023','Genre Challenge',12,12,null,'Achieved','Read at least one book in 12 different genres','Completed December 30th — final genre was Poetry.'],
    [2023,'Spooky October Reads','Seasonal Read',8,10,null,'Achieved','Read 8 horror or gothic books in October','Exceeded the goal — read 10 horror books in October.'],
    [2024,'Award Winners 2024','Award Nominees',10,10,null,'Achieved','Read 10 books from major award shortlists','Completed all 10 — Booker, Hugo, Nebula, and Edgar nominees.'],
    [2024,'Series Marathon','Series Marathon',6,6,null,'Achieved','Complete at least 2 full series in 2024','Completed The Throne Saga (3 books) and two standalones in related series.'],
    [2024,'Spooky October Reads','Seasonal Read',8,8,null,'Achieved','Read 8 horror or gothic books in October','Hit the target on October 29th — perfect timing.'],
    [2025,'Decade Dive: 2010s Best','Decade Challenge',12,12,null,'Achieved','Read 12 books published between 2010-2019','Discovered so many hidden gems from that decade.'],
    [2025,'Read the World','Genre Challenge',10,10,null,'Achieved','Read books by authors from 10 different countries','Reached authors from 12 countries — incredible range.'],
    [2025,'Spooky October Reads','Seasonal Read',8,9,null,'Achieved','Read 8 horror or gothic books in October','Read 9 — a new October record.'],
    [2026,'Genre Explorer 2026','Genre Challenge',15,7,null,'In Progress','Read at least 15 different genres across the year','7 genres covered so far: Fantasy, Mystery, Literary, Historical, SF, Romance, Horror.'],
    [2026,'Spooky October Reads','Seasonal Read',8,0,null,'Not Started','Read 8 horror or gothic books in October','Starting in October — have my reading pile ready.'],
    [2026,'Author Deep Dive: Vasquez','Author Deep Dive',3,3,null,'Achieved','Read all three Throne Saga books in 2026','Finished the trilogy! Elena Vasquez is an instant favourite.'],
    [2026,'2026 Debut Novels','Genre Challenge',10,5,null,'In Progress','Read 10 debut novels published in 2026','Tracking debutants carefully this year.'],
  ];

  CHALLENGES.forEach((row, ri) => {
    const dataRow = SEC2 + 2 + ri;
    vals.push({ range: `${S}!A${dataRow}:E${dataRow}`, values: [[row[0], row[1], row[2], row[3], row[4]]] });
    vals.push({ range: `${S}!G${dataRow}:I${dataRow}`, values: [[row[6], row[7], row[8]]] });
    fmt.push({ repeatCell: {
      range: gridRange(SID,dataRow-1,dataRow,5,6),
      cell: { userEnteredValue: { formulaValue: `=IFERROR(E${dataRow}/D${dataRow},0)` },
        userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER',
          numberFormat: { type: 'PERCENT', pattern: '0%' } }},
      fields: 'userEnteredValue,userEnteredFormat',
    }});
    const bg = ri % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,dataRow-1,dataRow,0,9), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    [0,1,2,3,4,6,7,8].forEach(ci => {
      fmt.push({ repeatCell: { range: gridRange(SID,dataRow-1,dataRow,ci,ci+1), cell: { userEnteredFormat: {
        backgroundColor: hex(C.input),
      }}, fields: 'userEnteredFormat.backgroundColor' }});
    });
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: dataRow-1, endIndex: dataRow }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});
  });

  // ── Section 3: Series Tracking ─────────────────────────────────────────────
  const SEC3 = SEC2 + 2 + CHALLENGES.length + 2;
  vals.push({ range: `${S}!A${SEC3}`, values: [['SECTION 3 — SERIES TRACKING']] });
  fmt.push({ mergeCells: { range: gridRange(SID,SEC3-1,SEC3,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,SEC3-1,SEC3,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.accent), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: SEC3-1, endIndex: SEC3 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  const serHdr = ['Series Name','Author','Genre','Total Books in Series','Books Read','Books Remaining (Auto)','Status','Next to Read','Notes'];
  vals.push({ range: `${S}!A${SEC3+1}`, values: [serHdr] });
  fmt.push({ repeatCell: { range: gridRange(SID,SEC3,SEC3+1,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.accent), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: SEC3, endIndex: SEC3+1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' }});

  const SERIES = [
    ['The Throne Saga','Elena Vasquez','Fantasy',3,3,null,'Completed','—','All three books read and loved. A perfect trilogy.'],
    ['Drift Series','Zara Ito','Science Fiction',4,2,null,'In Progress','Starfall Protocol (Book 2)','Currently reading Book 2. Cannot wait for the rest.'],
    ['The Tudor Files','Harriet Vance','Historical Fiction',4,2,null,'In Progress','Book 3 (forthcoming)','Book 3 not yet released — eagerly awaiting.'],
    ['The Accord Trilogy','Dion Larkin','Fantasy',3,0,null,'On Wishlist','The Phoenix Accord (Book 1)','On the wishlist — starting whenever I acquire it.'],
    ['The Ashlight Trilogy','Tarla Nieves','Young Adult',3,0,null,'Pre-Ordered','Kingdom of Ash and Starlight (Book 1)','Pre-ordered Book 1 — arrives October.'],
    ['The Mage Chronicles','Finn Kelley','Fantasy',5,0,null,'On Wishlist','The Mage\'s Burden (Book 1)','On wishlist — waiting to acquire.'],
    ['Mortuary Mysteries','Solomon Pierce','Mystery',4,1,null,'In Progress','Mortuary Mysteries Book 2','Enjoyed Book 1 — collecting the rest.'],
    ['The Wells Mysteries','Agatha Wells','Mystery',6,2,null,'In Progress','Book 3 (The Veiled Garden)','Working through the series gradually.'],
    ['Iron Gate Chronicles','Sofia Petrov','Fantasy',5,0,null,'In Progress','Book 1 (currently reading)','Currently reading Book 1 — very promising.'],
    ['The Iron Wars','Kieran Voss','Fantasy',5,0,null,'On Wishlist','Sovereign (Book 1)','On the wishlist.'],
    ['The Void Series','Zara Osei','Science Fiction',4,0,null,'Pre-Ordered','Void Walkers (Book 1)','Pre-ordered.'],
    ['Quinn Investigations','Harriet Quinn','Mystery',4,0,null,'On Wishlist','The Silent Evidence (Book 1)','On wishlist.'],
  ];

  SERIES.forEach((row, ri) => {
    const dataRow = SEC3 + 2 + ri;
    vals.push({ range: `${S}!A${dataRow}:E${dataRow}`, values: [[row[0], row[1], row[2], row[3], row[4]]] });
    vals.push({ range: `${S}!G${dataRow}:I${dataRow}`, values: [[row[6], row[7], row[8]]] });
    fmt.push({ repeatCell: {
      range: gridRange(SID,dataRow-1,dataRow,5,6),
      cell: { userEnteredValue: { formulaValue: `=IFERROR(D${dataRow}-E${dataRow},0)` },
        userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontSize: 9 }, horizontalAlignment: 'CENTER' }},
      fields: 'userEnteredValue,userEnteredFormat',
    }});
    const bg = ri % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,dataRow-1,dataRow,0,9), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    [0,1,2,3,4,6,7,8].forEach(ci => {
      fmt.push({ repeatCell: { range: gridRange(SID,dataRow-1,dataRow,ci,ci+1), cell: { userEnteredFormat: {
        backgroundColor: hex(C.input),
      }}, fields: 'userEnteredFormat.backgroundColor' }});
    });
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: dataRow-1, endIndex: dataRow }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});
  });

  // Column widths
  const colWidths = [120,180,110,80,80,100,110,160,200,90,90,90,90,90,90,90];
  colWidths.forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Freeze
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' }});

  await batchUpdate(id, fmt, '06-goals format');
  await valuesBatchUpdate(id, vals, '06-goals values');
  console.log('✅  Goals & Challenges done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
