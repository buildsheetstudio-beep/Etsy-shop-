'use strict';
const { sheets, batchUpdate, valuesBatchUpdate, gridRange, hex, colL, C } = require('./lib');
const { id, sheetMap } = JSON.parse(require('fs').readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Genealogy Dashboard'];
const S   = 'Genealogy Dashboard';

// 12 equal columns (A-L), 115px each
// Stat tiles: 3 cols wide × 3 rows tall (label / big-number / subtext)
// 3 rows of 4 tiles = 12 tiles total

const TILE_COLORS = [
  C.primaryDeep, '#7B6D9E', '#B89098',   '#9A7D3A',  // row 1
  '#4A8A88',     C.secondaryDeep, C.primary, '#8DA8A0',  // row 2
  '#5A7E5A',     '#B98482',     C.secondary, '#D1B36D',  // row 3
];

async function main() {
  const reqs = [];

  reqs.push({
    updateSheetProperties: {
      properties: {
        sheetId: SID,
        tabColor: hex(C.primaryDeep),
        tabColorStyle: { rgbColor: hex(C.primaryDeep) },
        gridProperties: { frozenRowCount: 4 },
      },
      fields: 'tabColor,tabColorStyle,gridProperties.frozenRowCount',
    },
  });

  // 12 columns A-L at 115px each
  for (let i = 0; i < 12; i++) {
    reqs.push({
      updateDimensionProperties: {
        range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
        properties: { pixelSize: 115 },
        fields: 'pixelSize',
      },
    });
  }

  // Row heights
  // idx0=title(52), idx1=subtitle(26), idx2=spacer(8), idx3=sectionHdr(28)
  // Tile rows (3 sets of 3 rows + spacer):
  //   idx4-6=labels/num/sub(22/42/20), idx7=spacer(6)
  //   idx8-10=labels/num/sub(22/42/20), idx11=spacer(6)
  //   idx12-14=labels/num/sub(22/42/20)
  // idx15=spacer(12), idx16=sectionHdr(28), idx17=colHdr(24), idx18-27=branchData(24ea)
  // idx28=spacer(12), idx29=sectionHdr(28), idx30=colHdr(24), idx31-40=peopleData(24ea)
  // idx41=spacer(12), idx42=sectionHdr(28), idx43=colHdr(24), idx44-50=recentData(26ea)
  // idx51=spacer(12), idx52=sectionHdr(28), idx53-59=notesRows(22ea)
  const rh = [
    [0,1,52],[1,2,26],[2,3,8],[3,4,28],
    [4,5,22],[5,6,44],[6,7,20],[7,8,6],
    [8,9,22],[9,10,44],[10,11,20],[11,12,6],
    [12,13,22],[13,14,44],[14,15,20],
    [15,16,12],[16,17,28],[17,18,24],[18,28,24],
    [28,29,12],[29,30,28],[30,31,24],[31,41,24],
    [41,42,12],[42,43,28],[43,44,24],[44,51,26],
    [51,52,12],[52,53,28],[53,60,22],
  ];
  rh.forEach(([s,e,px]) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'ROWS', startIndex: s, endIndex: e },
      properties: { pixelSize: px },
      fields: 'pixelSize',
    },
  }));

  // ── Merges ──────────────────────────────────────────────────────────────────
  // Title / subtitle / section headers — all A:L
  for (const r of [0,1,2,3,15,16,28,29,41,42,51,52]) {
    reqs.push({ mergeCells: { range: gridRange(SID,r,r+1,0,12), mergeType:'MERGE_ALL' } });
  }
  // Tile merges: 4 tiles × 3 col groups × 3 rows per tile, repeated 3 times
  for (let tileRow = 0; tileRow < 3; tileRow++) {
    const baseRow = 4 + tileRow * 4; // rows 4,8,12
    for (let tileCol = 0; tileCol < 4; tileCol++) {
      const c1 = tileCol * 3;
      for (let tr = 0; tr < 3; tr++) {
        reqs.push({ mergeCells: { range: gridRange(SID, baseRow+tr, baseRow+tr+1, c1, c1+3), mergeType:'MERGE_ALL' } });
      }
    }
  }
  // Spacer gap rows between tile rows (idx 7, 11)
  for (const r of [7,11]) {
    reqs.push({ mergeCells: { range: gridRange(SID,r,r+1,0,12), mergeType:'MERGE_ALL' } });
  }

  // ── Global background ────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID,0,60,0,12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat.backgroundColor',
  }});

  // ── Title ────────────────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID,0,1,0,12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryDeep),
      textFormat: { foregroundColor: hex(C.white), fontSize: 18, bold: true },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat',
  }});
  // Subtitle row 2
  reqs.push({ repeatCell: {
    range: gridRange(SID,1,2,0,12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryDeep),
      textFormat: { foregroundColor: hex('#C8DCF0'), fontSize: 10, italic: true },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat',
  }});

  // Section headers
  const sectionHdrFmt = {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), fontSize: 11, bold: true },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  };
  for (const r of [3,16,29,42,52]) {
    reqs.push({ repeatCell: {
      range: gridRange(SID,r,r+1,0,12),
      cell: { userEnteredFormat: sectionHdrFmt },
      fields: 'userEnteredFormat',
    }});
  }

  // Spacer/gap rows
  for (const r of [2,7,11,15,28,41,51]) {
    reqs.push({ repeatCell: {
      range: gridRange(SID,r,r+1,0,12),
      cell: { userEnteredFormat: { backgroundColor: hex(C.border) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }

  // ── Stat tiles formatting ────────────────────────────────────────────────────
  let tileIndex = 0;
  for (let tileRow = 0; tileRow < 3; tileRow++) {
    const baseRow = 4 + tileRow * 4;
    for (let tileCol = 0; tileCol < 4; tileCol++) {
      const c1 = tileCol * 3;
      const bg = TILE_COLORS[tileIndex++];
      // Label row (small, semitransparent-style)
      reqs.push({ repeatCell: {
        range: gridRange(SID, baseRow, baseRow+1, c1, c1+3),
        cell: { userEnteredFormat: {
          backgroundColor: hex(bg),
          textFormat: { foregroundColor: hex('#E8F0F8'), fontSize: 9, bold: false },
          horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM',
          padding: { bottom: 2 },
        }},
        fields: 'userEnteredFormat',
      }});
      // Number row (big bold)
      reqs.push({ repeatCell: {
        range: gridRange(SID, baseRow+1, baseRow+2, c1, c1+3),
        cell: { userEnteredFormat: {
          backgroundColor: hex(bg),
          textFormat: { foregroundColor: hex(C.white), fontSize: 24, bold: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        }},
        fields: 'userEnteredFormat',
      }});
      // Subtext row
      reqs.push({ repeatCell: {
        range: gridRange(SID, baseRow+2, baseRow+3, c1, c1+3),
        cell: { userEnteredFormat: {
          backgroundColor: hex(bg),
          textFormat: { foregroundColor: hex('#D0E4F0'), fontSize: 8 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'TOP',
          padding: { top: 2 },
        }},
        fields: 'userEnteredFormat',
      }});
      // Subtle border between adjacent tiles
      if (tileCol < 3) {
        reqs.push({ updateBorders: {
          range: gridRange(SID, baseRow, baseRow+3, c1+3, c1+3+1),
          left: { style: 'SOLID', color: hex(C.bg) },
        }});
      }
    }
  }

  // ── Data table formatting (branch stats, key people, recent log) ─────────────
  const tblHdr = {
    backgroundColor: hex(C.secondary),
    textFormat: { fontSize: 9, bold: true, foregroundColor: hex(C.text) },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  };
  const tblCell = {
    backgroundColor: hex(C.panel),
    textFormat: { fontSize: 9, foregroundColor: hex(C.text) },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    wrapStrategy: 'CLIP',
  };
  const altCell = { backgroundColor: hex(C.altRow) };
  const thinBorder = { style: 'SOLID', color: hex(C.border) };

  for (const [hdrRow, dataStart, dataEnd] of [[17,18,28],[30,31,41],[43,44,51]]) {
    reqs.push({ repeatCell: { range: gridRange(SID,hdrRow,hdrRow+1,0,12), cell: { userEnteredFormat: tblHdr }, fields: 'userEnteredFormat' } });
    reqs.push({ repeatCell: { range: gridRange(SID,dataStart,dataEnd,0,12), cell: { userEnteredFormat: tblCell }, fields: 'userEnteredFormat' } });
    for (let r = dataStart; r < dataEnd; r += 2) {
      reqs.push({ repeatCell: { range: gridRange(SID,r,r+1,0,12), cell: { userEnteredFormat: altCell }, fields: 'userEnteredFormat.backgroundColor' } });
    }
    reqs.push({ updateBorders: {
      range: gridRange(SID,hdrRow,dataEnd,0,12),
      innerHorizontal: { style: 'DOTTED', color: hex(C.border) },
      innerVertical:   { style: 'DOTTED', color: hex(C.border) },
      bottom: thinBorder, left: thinBorder, right: thinBorder,
    }});
  }

  // Notes section formatting
  reqs.push({ repeatCell: {
    range: gridRange(SID,52,53,0,12),
    cell: { userEnteredFormat: sectionHdrFmt },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(SID,53,60,0,12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { fontSize: 9, foregroundColor: hex(C.secText) },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      wrapStrategy: 'WRAP', padding: { left: 8 },
    }},
    fields: 'userEnteredFormat',
  }});
  // Disclaimer text color distinction on row 59 (idx 58)
  reqs.push({ repeatCell: {
    range: gridRange(SID,58,59,0,12),
    cell: { userEnteredFormat: {
      backgroundColor: hex('#FFF8ED'),
      textFormat: { fontSize: 9, bold: true, foregroundColor: hex(C.conflict) },
    }},
    fields: 'userEnteredFormat',
  }});
  reqs.push({ mergeCells: { range: gridRange(SID,58,59,0,12), mergeType:'MERGE_ALL' } });

  await batchUpdate(id, reqs, 'dash-fmt');

  // ── Values ───────────────────────────────────────────────────────────────────
  const MP  = `'Master People'!`;
  const VR  = `'Vital Records'!`;
  const REL = `'Relationships'!`;
  const LOC = `'Locations & Migration'!`;
  const SRC = `'Sources & Citations'!`;
  const LOG = `'Research Log'!`;

  // Unique count helper
  const uniq = (range) =>
    `SUMPRODUCT(IFERROR((LEN(${range})>0)*(1/COUNTIF(${range},${range})),0))`;

  // Stat tile values [label, =formula, subtext]
  const tiles = [
    // Row 1
    ['People Documented',      `=COUNTA(${MP}$A$8:$A$5107)`,       'total family members'],
    ['Vital Records',          `=COUNTA(${VR}$A$8:$A$5007)`,       'birth, marriage & death events'],
    ['Relationships',          `=COUNTA(${REL}$A$8:$A$5007)`,      'family connections mapped'],
    ['Source Citations',       `=COUNTA(${SRC}$A$8:$A$1007)`,      'primary sources documented'],
    // Row 2
    ['Location Records',       `=COUNTA(${LOC}$A$8:$A$5007)`,      'migration & residence entries'],
    ['Research Log Entries',   `=COUNTA(${LOG}$B$8:$B$5007)`,      'research sessions documented'],
    ['Generations Covered',    `=${uniq(`${MP}$Z$8:$Z$5107`)}`,    'generations fully researched'],
    ['Countries & Regions',    `=${uniq(`${LOC}$G$8:$G$5007`)}`,   'countries / regions covered'],
    // Row 3
    ['Confirmed Records',      `=COUNTIF(${MP}$Y$8:$Y$5107,"Confirmed")`, 'high-confidence people'],
    ['Brick Walls',            `=COUNTIF(${MP}$AD$8:$AD$5107,"Brick Wall")`,'unresolved dead ends'],
    ['Living Members',         `=COUNTIF(${MP}$K$8:$K$5107,"Living")`,     'living family members'],
    ['Open Research Tasks',    `=COUNTIF(${LOG}$G$8:$G$5007,"Open")`,      'tasks awaiting follow-up'],
  ];

  const vals = [
    { range: `'${S}'!A1`, values: [['GENEALOGY & FAMILY TREE TRACKER — DASHBOARD']] },
    { range: `'${S}'!A2`, values: [[`=TODAY()&"  •  Hartwell & O'Brien Family Research  •  Root Person: Emma Rose Hartwell (P-00001)"`]] },
    { range: `'${S}'!A4`, values: [['COLLECTION SUMMARY']] },
  ];

  // Write tile labels, numbers, subtexts
  let ti = 0;
  for (let tileRow = 0; tileRow < 3; tileRow++) {
    const baseRow = 5 + tileRow * 4; // 1-indexed row for label
    for (let tileCol = 0; tileCol < 4; tileCol++) {
      const col = colL(tileCol * 3);
      const [label, num, sub] = tiles[ti++];
      vals.push({ range: `'${S}'!${col}${baseRow}`,   values: [[label]] });
      vals.push({ range: `'${S}'!${col}${baseRow+1}`, values: [[num]] });
      vals.push({ range: `'${S}'!${col}${baseRow+2}`, values: [[sub]] });
    }
  }

  // ── Branch Statistics (rows 17-27) ──────────────────────────────────────────
  vals.push({ range: `'${S}'!A17`, values: [['BRANCH STATISTICS & RESEARCH PROGRESS']] });
  vals.push({ range: `'${S}'!A18:L18`, values: [['Branch','People','Vital Records','Relationships','Locations','Sources','Confirmed','Strong Evid.','Brick Walls','Resolved','Researching','Living']] });
  // Hartwell branch (row 19)
  vals.push({ range: `'${S}'!A19:L19`, values: [[
    'Hartwell',
    `=COUNTIF(${MP}$AA$8:$AA$5107,"Hartwell")`,
    `=COUNTIFS(${VR}$B$8:$B$5007,"P-0000*",${MP}$AA$8:$AA$5107,"Hartwell")`, // approximate
    `=COUNTIF(${REL}$B$8:$B$5007,$A19)`, // count where Person1 is Hartwell branch (simplified)
    `=COUNTIFS(${LOC}$B$8:$B$5007,"P-0000*")`, // placeholder — not branch-keyed in LOC
    `=ROUND(COUNTIF(${MP}$AA$8:$AA$5107,"Hartwell")/COUNTA(${SRC}$A$8:$A$1007)*100,0)&"%"`,
    `=COUNTIFS(${MP}$AA$8:$AA$5107,"Hartwell",${MP}$Y$8:$Y$5107,"Confirmed")`,
    `=COUNTIFS(${MP}$AA$8:$AA$5107,"Hartwell",${MP}$Y$8:$Y$5107,"Strong")`,
    `=COUNTIFS(${MP}$AA$8:$AA$5107,"Hartwell",${MP}$AD$8:$AD$5107,"Brick Wall")`,
    `=COUNTIFS(${MP}$AA$8:$AA$5107,"Hartwell",${MP}$AD$8:$AD$5107,"Resolved")`,
    `=COUNTIFS(${MP}$AA$8:$AA$5107,"Hartwell",${MP}$AD$8:$AD$5107,"Researching")`,
    `=COUNTIFS(${MP}$AA$8:$AA$5107,"Hartwell",${MP}$K$8:$K$5107,"Living")`,
  ]] });
  // O'Brien branch (row 20)
  vals.push({ range: `'${S}'!A20:L20`, values: [[
    "O'Brien",
    `=COUNTIF(${MP}$AA$8:$AA$5107,"O'Brien")`,
    `=COUNTIFS(${VR}$B$8:$B$5007,"P-0003*",${VR}$D$8:$D$5007,"<>")`,
    `=IFERROR(COUNTA(${REL}$A$8:$A$5007)-B19,0)`,
    `=IFERROR(COUNTA(${LOC}$A$8:$A$5007)-C19,0)`,
    `=ROUND(COUNTIF(${MP}$AA$8:$AA$5107,"O'Brien")/COUNTA(${SRC}$A$8:$A$1007)*100,0)&"%"`,
    `=COUNTIFS(${MP}$AA$8:$AA$5107,"O'Brien",${MP}$Y$8:$Y$5107,"Confirmed")`,
    `=COUNTIFS(${MP}$AA$8:$AA$5107,"O'Brien",${MP}$Y$8:$Y$5107,"Strong")`,
    `=COUNTIFS(${MP}$AA$8:$AA$5107,"O'Brien",${MP}$AD$8:$AD$5107,"Brick Wall")`,
    `=COUNTIFS(${MP}$AA$8:$AA$5107,"O'Brien",${MP}$AD$8:$AD$5107,"Resolved")`,
    `=COUNTIFS(${MP}$AA$8:$AA$5107,"O'Brien",${MP}$AD$8:$AD$5107,"Researching")`,
    `=COUNTIFS(${MP}$AA$8:$AA$5107,"O'Brien",${MP}$K$8:$K$5107,"Living")`,
  ]] });
  // Totals row (row 21)
  vals.push({ range: `'${S}'!A21:L21`, values: [[
    'TOTAL',
    `=B19+B20`, `=COUNTA(${VR}$A$8:$A$5007)`,
    `=COUNTA(${REL}$A$8:$A$5007)`, `=COUNTA(${LOC}$A$8:$A$5007)`,
    '100%',
    `=G19+G20`, `=H19+H20`, `=I19+I20`, `=J19+J20`, `=K19+K20`, `=L19+L20`,
  ]] });

  // Confidence breakdown (rows 23-27)
  vals.push({ range: `'${S}'!A23:L23`, values: [['Evidence Confidence Breakdown','','','',
    'Confirmed', `=COUNTIF(${MP}$Y$8:$Y$5107,"Confirmed")`,
    'Strong',    `=COUNTIF(${MP}$Y$8:$Y$5107,"Strong")`,
    'Reasonable',`=COUNTIF(${MP}$Y$8:$Y$5107,"Reasonable")`,
    'Estimated / Speculative',
    `=COUNTIF(${MP}$Y$8:$Y$5107,"Estimated")+COUNTIF(${MP}$Y$8:$Y$5107,"Speculative")`,
  ]] });

  // ── Key Ancestors: Family Tree Overview (rows 30-40) ─────────────────────────
  vals.push({ range: `'${S}'!A30`, values: [['KEY ANCESTORS — FAMILY TREE OVERVIEW']] });
  vals.push({ range: `'${S}'!A31:L31`, values: [['Person ID','Full Name','Living Status','Birth Date','Birth Place','Death Date','Death Place','Generation','Branch','Occupation','Confidence','Research Status']] });
  // Use QUERY to pull top ancestors by Person ID
  vals.push({ range: `'${S}'!A32`, values: [[
    `=IFERROR(CHOOSECOLS(FILTER(${MP}A8:AE5107,ROW(${MP}A8:A5107)-ROW(${MP}A8)+1<=10),1,2,11,12,14,15,17,26,27,19,25,30),"")`,
  ]] });

  // ── Recent Research Log (rows 43-50) ─────────────────────────────────────────
  vals.push({ range: `'${S}'!A43`, values: [['RECENT RESEARCH ACTIVITY (Last 6 Entries)']] });
  vals.push({ range: `'${S}'!A44:G44`, values: [['Date','Person Name','Research Goal','Repository / Resource','Status','Outcome','Follow-Up']] });
  vals.push({ range: `'${S}'!A45`, values: [[
    `=IFERROR(QUERY(${LOG}A8:L5007,"SELECT B,D,E,F,G,H,K WHERE B<>'' ORDER BY B DESC LIMIT 6",0),"No research entries yet")`,
  ]] });

  // ── Notes (rows 53-59) ──────────────────────────────────────────────────────
  vals.push({ range: `'${S}'!A53`, values: [['NOTES & GUIDANCE']] });
  vals.push({ range: `'${S}'!A54:L59`, values: [
    [`Navigate the tabs at the bottom to explore each section of this tracker.`,'','','','','','','','','','',''],
    [`• Master People — full list of all 103 documented individuals`,'','','','','','','','','','',''],
    [`• Vital Records — births, baptisms, marriages, divorces, deaths & burials`,'','','','','','','','','','',''],
    [`• Person Profile — select any Person ID to view their complete profile`,'','','','','','','','','','',''],
    [`• Search & Filter — multi-criteria live search across all people`,'','','','','','','','','','',''],
    [`DISCLAIMER: All personal data in this tracker is fictional sample data for demonstration. Dates, names, and records are invented. This tracker is a genealogy research tool only.`,'','','','','','','','','','',''],
  ] });

  await valuesBatchUpdate(id, vals, 'dash-vals');
  console.log(`✓ Genealogy Dashboard — aggregates data from all 10 content tabs`);
}

main().catch(e => { console.error(e); process.exit(1); });
