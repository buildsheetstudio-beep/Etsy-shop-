'use strict';
const { sheets, batchUpdate, valuesBatchUpdate, gridRange, hex, colL, C } = require('./lib');
const { id, sheetMap } = JSON.parse(require('fs').readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Person Profile'];
const S   = 'Person Profile';

async function main() {
  const reqs = [];

  // Tab color — primary heritage blue
  reqs.push({
    updateSheetProperties: {
      properties: {
        sheetId: SID,
        tabColor: hex(C.primary),
        tabColorStyle: { rgbColor: hex(C.primary) },
        gridProperties: { frozenRowCount: 2 },
      },
      fields: 'tabColor,tabColorStyle,gridProperties.frozenRowCount',
    },
  });

  // Column widths: A-J (10 cols)
  // A=labels, B=left-values, C=spacer, D=right-labels, E=right-values, F-J=table cols 6-10
  const colWidths = [158, 208, 14, 158, 208, 105, 105, 105, 105, 105];
  colWidths.forEach((px, i) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: px },
      fields: 'pixelSize',
    },
  }));

  // Row heights
  // idx 0=title(52), 1=selector(38), 2=spacer(8), 3=BasicInfo hdr(30),
  // 4-18=field rows (28 each), 19=bio label(26), 20-21=bio content(50 each),
  // 22=spacer(8), 23=VitalRec hdr(30), 24=VR col hdrs(26), 25-42=VR data(20),
  // 43=spacer(8), 44=Rels hdr(30), 45=Rels col hdrs(26), 46-63=Rels data(20),
  // 64=spacer(8), 65=Locs hdr(30), 66=Locs col hdrs(26), 67-84=Locs data(20),
  // 85=spacer(8), 86=ResLog hdr(30), 87=ResLog col hdrs(26), 88-109=ResLog data(20)
  const rh = [
    [0,1,52],[1,2,38],[2,3,8],[3,4,30],[4,19,28],[19,20,26],[20,22,50],
    [22,23,8],[23,24,30],[24,25,26],[25,43,20],
    [43,44,8],[44,45,30],[45,46,26],[46,64,20],
    [64,65,8],[65,66,30],[66,67,26],[67,85,20],
    [85,86,8],[86,87,30],[87,88,26],[88,110,20],
  ];
  rh.forEach(([s,e,px]) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'ROWS', startIndex: s, endIndex: e },
      properties: { pixelSize: px },
      fields: 'pixelSize',
    },
  }));

  // ── Merges ──────────────────────────────────────────────────────────────────
  // Title row 1: A1:J1
  reqs.push({ mergeCells: { range: gridRange(SID,0,1,0,10), mergeType:'MERGE_ALL' } });
  // Selector row 2: B2:C2 (input), E2:J2 (name display)
  reqs.push({ mergeCells: { range: gridRange(SID,1,2,1,3), mergeType:'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID,1,2,4,10), mergeType:'MERGE_ALL' } });
  // Basic Info section header row 4: A4:J4
  reqs.push({ mergeCells: { range: gridRange(SID,3,4,0,10), mergeType:'MERGE_ALL' } });
  // Bio sub-label row 20: A20:J20
  reqs.push({ mergeCells: { range: gridRange(SID,19,20,0,10), mergeType:'MERGE_ALL' } });
  // Bio content rows 21-22: A21:J22
  reqs.push({ mergeCells: { range: gridRange(SID,20,22,0,10), mergeType:'MERGE_ALL' } });
  // Section headers for data tables
  for (const r of [23,44,65,86]) {
    reqs.push({ mergeCells: { range: gridRange(SID,r,r+1,0,10), mergeType:'MERGE_ALL' } });
  }

  // ── Formatting ──────────────────────────────────────────────────────────────
  // Title
  reqs.push({ repeatCell: {
    range: gridRange(SID,0,1,0,10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryDeep),
      textFormat: { foregroundColor: hex(C.white), fontSize: 16, bold: true },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat',
  }});

  // Selector bar (row 2) background
  reqs.push({ repeatCell: {
    range: gridRange(SID,1,2,0,10),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg), verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat',
  }});
  // A2: label
  reqs.push({ repeatCell: {
    range: gridRange(SID,1,2,0,1),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.bg),
      textFormat: { fontSize: 10, bold: true, foregroundColor: hex(C.text) },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE', padding: { right: 5 },
    }},
    fields: 'userEnteredFormat',
  }});
  // B2:C2 — input cell
  reqs.push({ repeatCell: {
    range: gridRange(SID,1,2,1,3),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.input),
      textFormat: { fontSize: 11, foregroundColor: hex(C.text) },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 5 },
    }},
    fields: 'userEnteredFormat',
  }});
  // D2: label
  reqs.push({ repeatCell: {
    range: gridRange(SID,1,2,3,4),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.bg),
      textFormat: { fontSize: 10, bold: true, foregroundColor: hex(C.text) },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE', padding: { right: 5 },
    }},
    fields: 'userEnteredFormat',
  }});
  // E2:J2 — name display
  reqs.push({ repeatCell: {
    range: gridRange(SID,1,2,4,10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.bg),
      textFormat: { fontSize: 13, bold: true, italic: true, foregroundColor: hex(C.primaryDeep) },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 6 },
    }},
    fields: 'userEnteredFormat',
  }});

  // Spacer rows
  for (const r of [2,22,43,64,85]) {
    reqs.push({ repeatCell: {
      range: gridRange(SID,r,r+1,0,10),
      cell: { userEnteredFormat: { backgroundColor: hex(C.border) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }

  // Section headers style
  const sectionHdr = {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), fontSize: 11, bold: true },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  };
  for (const r of [3,23,44,65,86]) {
    reqs.push({ repeatCell: {
      range: gridRange(SID,r,r+1,0,10),
      cell: { userEnteredFormat: sectionHdr },
      fields: 'userEnteredFormat',
    }});
  }

  // Bio sub-label (row 20)
  reqs.push({ repeatCell: {
    range: gridRange(SID,19,20,0,10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { fontSize: 10, bold: true, foregroundColor: hex(C.text) },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 6 },
    }},
    fields: 'userEnteredFormat',
  }});
  // Bio content (rows 21-22)
  reqs.push({ repeatCell: {
    range: gridRange(SID,20,22,0,10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { fontSize: 10, italic: true, foregroundColor: hex(C.text) },
      horizontalAlignment: 'LEFT', verticalAlignment: 'TOP',
      wrapStrategy: 'WRAP', padding: { left: 6, top: 4 },
    }},
    fields: 'userEnteredFormat',
  }});

  // Field label cells: col A rows 5-19 (idx 4-18)
  const labelFmt = {
    backgroundColor: hex(C.secondary),
    textFormat: { fontSize: 10, bold: true, foregroundColor: hex(C.text) },
    horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE', padding: { right: 6 },
  };
  reqs.push({ repeatCell: { range: gridRange(SID,4,19,0,1), cell: { userEnteredFormat: labelFmt }, fields: 'userEnteredFormat' } });
  reqs.push({ repeatCell: { range: gridRange(SID,4,19,3,4), cell: { userEnteredFormat: labelFmt }, fields: 'userEnteredFormat' } });

  // Value cells: col B and col E rows 5-19
  const valFmt = {
    backgroundColor: hex(C.panel),
    textFormat: { fontSize: 10, foregroundColor: hex(C.text) },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 5 },
  };
  reqs.push({ repeatCell: { range: gridRange(SID,4,19,1,2), cell: { userEnteredFormat: valFmt }, fields: 'userEnteredFormat' } });
  reqs.push({ repeatCell: { range: gridRange(SID,4,19,4,5), cell: { userEnteredFormat: valFmt }, fields: 'userEnteredFormat' } });

  // Spacer col C (rows 5-19): light border color
  reqs.push({ repeatCell: {
    range: gridRange(SID,4,19,2,3),
    cell: { userEnteredFormat: { backgroundColor: hex(C.border) } },
    fields: 'userEnteredFormat.backgroundColor',
  }});

  // Data table column headers (rows 25, 46, 67, 88)
  const tblHdr = {
    backgroundColor: hex(C.secondary),
    textFormat: { fontSize: 9, bold: true, foregroundColor: hex(C.text) },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  };
  for (const r of [24,45,66,87]) {
    reqs.push({ repeatCell: { range: gridRange(SID,r,r+1,0,10), cell: { userEnteredFormat: tblHdr }, fields: 'userEnteredFormat' } });
  }

  // Data table cell format
  const tblCell = {
    backgroundColor: hex(C.panel),
    textFormat: { fontSize: 9, foregroundColor: hex(C.text) },
    horizontalAlignment: 'LEFT', verticalAlignment: 'TOP',
    wrapStrategy: 'WRAP', padding: { left: 4 },
  };
  const altCell = { backgroundColor: hex(C.altRow) };

  // VR data rows 26-43 (idx 25-42)
  reqs.push({ repeatCell: { range: gridRange(SID,25,43,0,10), cell: { userEnteredFormat: tblCell }, fields: 'userEnteredFormat' } });
  for (let r = 25; r < 43; r += 2) reqs.push({ repeatCell: { range: gridRange(SID,r,r+1,0,10), cell: { userEnteredFormat: altCell }, fields: 'userEnteredFormat.backgroundColor' } });

  // Rels data rows 47-64 (idx 46-63)
  reqs.push({ repeatCell: { range: gridRange(SID,46,64,0,10), cell: { userEnteredFormat: tblCell }, fields: 'userEnteredFormat' } });
  for (let r = 46; r < 64; r += 2) reqs.push({ repeatCell: { range: gridRange(SID,r,r+1,0,10), cell: { userEnteredFormat: altCell }, fields: 'userEnteredFormat.backgroundColor' } });

  // Locs data rows 68-85 (idx 67-84)
  reqs.push({ repeatCell: { range: gridRange(SID,67,85,0,10), cell: { userEnteredFormat: tblCell }, fields: 'userEnteredFormat' } });
  for (let r = 67; r < 85; r += 2) reqs.push({ repeatCell: { range: gridRange(SID,r,r+1,0,10), cell: { userEnteredFormat: altCell }, fields: 'userEnteredFormat.backgroundColor' } });

  // ResLog data rows 89-109 (idx 88-108)
  reqs.push({ repeatCell: { range: gridRange(SID,88,110,0,10), cell: { userEnteredFormat: tblCell }, fields: 'userEnteredFormat' } });
  for (let r = 88; r < 110; r += 2) reqs.push({ repeatCell: { range: gridRange(SID,r,r+1,0,10), cell: { userEnteredFormat: altCell }, fields: 'userEnteredFormat.backgroundColor' } });

  // Borders around sections
  const thinGray = { style: 'SOLID', color: hex(C.border) };
  const dotted   = { style: 'DOTTED', color: hex(C.border) };
  // Basic info block
  reqs.push({ updateBorders: {
    range: gridRange(SID,4,22,0,10),
    innerHorizontal: { style: 'SOLID', color: hex(C.altRow) },
    bottom: thinGray, left: thinGray, right: thinGray,
  }});
  // Vertical divider between left and right panes (col C = spacer, already colored)
  // Table section borders
  for (const [s,e] of [[24,43],[44,64],[65,85],[86,110]]) {
    reqs.push({ updateBorders: {
      range: gridRange(SID,s,e,0,10),
      innerHorizontal: dotted, innerVertical: dotted,
      bottom: thinGray, left: thinGray, right: thinGray,
    }});
  }

  // Data validation: B2 — Person ID dropdown from Master People
  reqs.push({
    setDataValidation: {
      range: gridRange(SID,1,2,1,3),
      rule: {
        condition: {
          type: 'ONE_OF_RANGE',
          values: [{ userEnteredValue: `='Master People'!$A$8:$A$5107` }],
        },
        showCustomUi: true,
        strict: false,
      },
    },
  });

  await batchUpdate(id, reqs, 'profile-fmt');

  // ── Values ───────────────────────────────────────────────────────────────────
  // VLOOKUP helper: col index into Master People A:AE (1-indexed)
  const vl = (col) => `=IFERROR(VLOOKUP($B$2,'Master People'!$A$8:$AE$5007,${col},FALSE),"")`;

  // QUERY formulas — 0 header rows since source data rows start at row 8
  const guard = (inner) => `=IF($B$2="","← Select a Person ID to load profile",${inner})`;

  const vitalsQ = guard(`IFERROR(QUERY('Vital Records'!A8:M5007,"SELECT D,E,F,G,H,I,J,K,L,M WHERE B='"&$B$2&"' ORDER BY E",0),{"No vital records found","","","","","","","","",""})`);
  const relsQ   = guard(`IFERROR(QUERY('Relationships'!A8:M5007,"SELECT C,D,F,G,H,I,L,J,K,M WHERE B='"&$B$2&"' OR E='"&$B$2&"' ORDER BY H",0),{"No relationships found","","","","","","","","",""})`);
  const locsQ   = guard(`IFERROR(QUERY('Locations & Migration'!A8:N5007,"SELECT D,E,F,G,H,I,J,L,M,N WHERE B='"&$B$2&"' ORDER BY E",0),{"No location records found","","","","","","","","",""})`);
  const logQ    = guard(`IFERROR(QUERY('Research Log'!A8:L5007,"SELECT B,E,F,G,H,I,J,K,L WHERE C='"&$B$2&"' ORDER BY B",0),{"No research log entries found","","","","","","","",""})`);

  const vals = [
    // Row 1: title
    { range: `'${S}'!A1`, values: [['GENEALOGY PERSON PROFILE']] },
    // Row 2: selector bar
    { range: `'${S}'!A2`, values: [['PERSON ID:']] },
    { range: `'${S}'!B2`, values: [['P-00001']] }, // pre-filled with root person
    { range: `'${S}'!D2`, values: [['FULL NAME:']] },
    { range: `'${S}'!E2`, values: [[`=IFERROR(VLOOKUP($B$2,'Master People'!$A$8:$B$5007,2,FALSE),"")`]] },
    // Row 4: Basic Info header
    { range: `'${S}'!A4`, values: [['BASIC INFORMATION']] },
    // Field rows 5-19 (A=left label, B=left value, D=right label, E=right value)
    { range: `'${S}'!A5:E19`, values: [
      ['Person ID',        vl(1),  '', 'Preferred Name',        vl(2)  ],
      ['Given Name',       vl(3),  '', 'Middle Name',           vl(4)  ],
      ['Birth Surname',    vl(5),  '', 'Married Surname',       vl(6)  ],
      ['Prefix',           vl(8),  '', 'Suffix',                vl(9)  ],
      ['Sex / Gender',     vl(10), '', 'Living Status',         vl(11) ],
      ['Birth Date',       vl(12), '', 'Birth Date Confidence', vl(13) ],
      ['Birth Place',      vl(14), '', 'Burial Place',          vl(18) ],
      ['Death Date',       vl(15), '', 'Death Date Confidence', vl(16) ],
      ['Death Place',      vl(17), '', 'Occupation',            vl(19) ],
      ['Generation',       vl(26), '', 'Family Branch',         vl(27) ],
      ['Religion',         vl(21), '', 'Nationality',           vl(22) ],
      ['Language',         vl(20), '', 'Research Status',       vl(30) ],
      ['Evidence Conf.',   vl(25), '', 'Primary Source ID',     vl(24) ],
      ['Alternate Name',   vl(7),  '', 'Photo / Doc URL',       vl(23) ],
      ['Date Added',       vl(28), '', 'Last Updated',          vl(29) ],
    ]},
    // Row 20: bio sub-label
    { range: `'${S}'!A20`, values: [['Biography / Notes']] },
    // Rows 21-22: bio content
    { range: `'${S}'!A21`, values: [[vl(31)]] },
    // Vital Records section
    { range: `'${S}'!A24`, values: [['VITAL RECORDS']] },
    { range: `'${S}'!A25:J25`, values: [['Event Type','Event Date','Date Precision','Event Place','Place Detail','Witnesses / Officiants','Source ID','Evidence Conf.','Citation Status','Notes']] },
    { range: `'${S}'!A26`, values: [[vitalsQ]] },
    // Relationships section
    { range: `'${S}'!A45`, values: [['RELATIONSHIPS']] },
    { range: `'${S}'!A46:J46`, values: [['Person 1 Name','Relationship Type','Person 2 Name','Status','Start Date','End Date','Expected Reciprocal','Source ID','Confidence','Notes']] },
    { range: `'${S}'!A47`, values: [[relsQ]] },
    // Locations section
    { range: `'${S}'!A66`, values: [['LOCATIONS & MIGRATION']] },
    { range: `'${S}'!A67:J67`, values: [['Event Type','Date','Date Precision','Country','State / Province','City / Town','Address Detail','Source ID','Evidence Conf.','Notes']] },
    { range: `'${S}'!A68`, values: [[locsQ]] },
    // Research Log section
    { range: `'${S}'!A87`, values: [['RESEARCH LOG']] },
    { range: `'${S}'!A88:I88`, values: [['Date','Research Goal','Repository / Resource','Research Status','Research Outcome','Source Found','Evidence Conf.','Follow-Up Required','Next Steps / Notes']] },
    { range: `'${S}'!A89`, values: [[logQ]] },
  ];

  await valuesBatchUpdate(id, vals, 'profile-vals');
  console.log(`✓ Person Profile — lookup view for ${103} people`);
}

main().catch(e => { console.error(e); process.exit(1); });
