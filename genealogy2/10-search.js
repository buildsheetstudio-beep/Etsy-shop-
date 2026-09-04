'use strict';
const { sheets, batchUpdate, valuesBatchUpdate, gridRange, hex, colL, C } = require('./lib');
const { id, sheetMap } = JSON.parse(require('fs').readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Search & Filter'];
const S   = 'Search & Filter';

async function main() {
  const reqs = [];

  reqs.push({
    updateSheetProperties: {
      properties: {
        sheetId: SID,
        tabColor: hex(C.info),
        tabColorStyle: { rgbColor: hex(C.info) },
        gridProperties: { frozenRowCount: 10, rowCount: 1100 },
      },
      fields: 'tabColor,tabColorStyle,gridProperties.frozenRowCount,gridProperties.rowCount',
    },
  });

  // 12 columns A-L
  // A(160px) label  |  B(200px) input  |  C(15px) spacer
  // D(160px) label  |  E(200px) input  |  F-L(110px each)
  const colWidths = [160, 200, 15, 160, 200, 110, 110, 110, 110, 110, 110, 110];
  colWidths.forEach((px,i) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px },
      fields: 'pixelSize',
    },
  }));

  // Row heights
  const rh = [
    [0,1,48], [1,2,28], [2,3,34], [3,4,34], [4,5,34], [5,6,34], [6,7,34],
    [7,8,10], // thin spacer
    [8,9,32], // results section header
    [9,10,26],[10,1010,22], // col headers + results data
  ];
  rh.forEach(([s,e,px]) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'ROWS', startIndex: s, endIndex: e },
      properties: { pixelSize: px },
      fields: 'pixelSize',
    },
  }));

  // ── Merges ──────────────────────────────────────────────────────────────────
  // Title A1:L1
  reqs.push({ mergeCells: { range: gridRange(SID,0,1,0,12), mergeType:'MERGE_ALL' } });
  // Row 2: subtitle A2:L2
  reqs.push({ mergeCells: { range: gridRange(SID,1,2,0,12), mergeType:'MERGE_ALL' } });
  // Row 3: Name search spans B3:E3 (wider input)
  reqs.push({ mergeCells: { range: gridRange(SID,2,3,1,5), mergeType:'MERGE_ALL' } });
  // Rows 4-7: B and E are inputs (single cells — no merge needed)
  // Row 8: spacer A8:L8
  reqs.push({ mergeCells: { range: gridRange(SID,7,8,0,12), mergeType:'MERGE_ALL' } });
  // Row 9: results header A9:L9
  reqs.push({ mergeCells: { range: gridRange(SID,8,9,0,12), mergeType:'MERGE_ALL' } });

  // ── Formatting ──────────────────────────────────────────────────────────────
  // Title
  reqs.push({ repeatCell: {
    range: gridRange(SID,0,1,0,12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryDeep),
      textFormat: { foregroundColor: hex(C.white), fontSize: 15, bold: true },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat',
  }});

  // Subtitle row 2
  reqs.push({ repeatCell: {
    range: gridRange(SID,1,2,0,12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.bg),
      textFormat: { foregroundColor: hex(C.secText), fontSize: 10, italic: true },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat',
  }});

  // Criteria area (rows 3-7): background
  reqs.push({ repeatCell: {
    range: gridRange(SID,2,7,0,12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat.backgroundColor',
  }});

  // Label cells: col A and col D, rows 3-7
  const lblFmt = {
    backgroundColor: hex(C.secondary),
    textFormat: { fontSize: 10, bold: true, foregroundColor: hex(C.text) },
    horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE', padding: { right: 6 },
  };
  reqs.push({ repeatCell: { range: gridRange(SID,2,7,0,1), cell: { userEnteredFormat: lblFmt }, fields: 'userEnteredFormat' } });
  reqs.push({ repeatCell: { range: gridRange(SID,2,7,3,4), cell: { userEnteredFormat: lblFmt }, fields: 'userEnteredFormat' } });

  // Input cells: col B rows 3-7, col E rows 4-7
  const inputFmt = {
    backgroundColor: hex(C.input),
    textFormat: { fontSize: 11, foregroundColor: hex(C.text) },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 5 },
  };
  reqs.push({ repeatCell: { range: gridRange(SID,2,7,1,2), cell: { userEnteredFormat: inputFmt }, fields: 'userEnteredFormat' } });
  reqs.push({ repeatCell: { range: gridRange(SID,3,7,4,5), cell: { userEnteredFormat: inputFmt }, fields: 'userEnteredFormat' } });

  // Spacer col C in criteria area
  reqs.push({ repeatCell: {
    range: gridRange(SID,2,7,2,3),
    cell: { userEnteredFormat: { backgroundColor: hex(C.border) } },
    fields: 'userEnteredFormat.backgroundColor',
  }});

  // Thin spacer row 8
  reqs.push({ repeatCell: {
    range: gridRange(SID,7,8,0,12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.border) } },
    fields: 'userEnteredFormat.backgroundColor',
  }});

  // Results section header row 9
  reqs.push({ repeatCell: {
    range: gridRange(SID,8,9,0,12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { foregroundColor: hex(C.white), fontSize: 11, bold: true },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat',
  }});

  // Results column headers row 10
  reqs.push({ repeatCell: {
    range: gridRange(SID,9,10,0,12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { fontSize: 9, bold: true, foregroundColor: hex(C.text) },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat',
  }});

  // Results data rows 11+ (alternating)
  reqs.push({ repeatCell: {
    range: gridRange(SID,10,1010,0,12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { fontSize: 9, foregroundColor: hex(C.text) },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      wrapStrategy: 'CLIP',
    }},
    fields: 'userEnteredFormat',
  }});
  for (let r = 10; r < 1010; r += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(SID,r,r+1,0,12),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }

  // Borders
  const thin = { style: 'SOLID', color: hex(C.border) };
  reqs.push({ updateBorders: {
    range: gridRange(SID,2,7,0,12),
    bottom: thin, innerHorizontal: { style: 'SOLID', color: hex(C.altRow) },
  }});
  reqs.push({ updateBorders: {
    range: gridRange(SID,9,1010,0,12),
    innerHorizontal: { style: 'DOTTED', color: hex(C.border) },
    innerVertical:   { style: 'DOTTED', color: hex(C.border) },
    bottom: thin, left: thin, right: thin,
  }});

  // ── Data Validation ─────────────────────────────────────────────────────────
  const dv = (range, values) => ({
    setDataValidation: {
      range,
      rule: {
        condition: { type: 'ONE_OF_LIST', values: values.map(v => ({ userEnteredValue: v })) },
        showCustomUi: true, strict: true,
      },
    },
  });

  // B4: Family Branch
  reqs.push(dv(gridRange(SID,3,4,1,2), ['All','Hartwell',"O'Brien"]));
  // E4: Living Status
  reqs.push(dv(gridRange(SID,3,4,4,5), ['All','Living','Deceased','Unknown']));
  // B5: Generation Label
  reqs.push(dv(gridRange(SID,4,5,1,2), ['All','Self / Root','Parent','Grandparent','Great-Grandparent','2× Great-Grandparent','3× Great-Grandparent','Collateral']));
  // E5: Research Status
  reqs.push(dv(gridRange(SID,4,5,4,5), ['All','Resolved','Researching','Brick Wall','Conflicting Evidence','DNA Pending','Not Started']));
  // B6: Evidence Confidence
  reqs.push(dv(gridRange(SID,5,6,1,2), ['All','Confirmed','Strong','Reasonable','Probable','Estimated','Speculative']));

  await batchUpdate(id, reqs, 'search-fmt');

  // ── Values ───────────────────────────────────────────────────────────────────
  // The FILTER formula: Master People A:AE (31 cols), result columns via CHOOSECOLS:
  //   1=PersonID, 2=DisplayName, 10=Sex/Gender, 11=LivingStatus, 12=BirthDate,
  //   14=BirthPlace, 15=DeathDate, 17=DeathPlace, 25=EvidenceConf, 26=Generation,
  //   27=FamilyBranch, 30=ResearchStatus

  const MP = `'Master People'!`;
  const BY = `IFERROR(VALUE(LEFT(${MP}L8:L5107,4)),0)`;

  const filterExpr =
    `(($B$3="")+ISNUMBER(SEARCH($B$3,${MP}B8:B5107)))*` +
    `(($B$4="All")+(${MP}AA8:AA5107=$B$4))*` +
    `(($E$4="All")+(${MP}K8:K5107=$E$4))*` +
    `(($B$5="All")+(${MP}Z8:Z5107=$B$5))*` +
    `(($E$5="All")+(${MP}AD8:AD5107=$E$5))*` +
    `(($B$6="All")+(${MP}Y8:Y5107=$B$6))*` +
    `(($E$6="")+ISNUMBER(SEARCH($E$6,${MP}N8:N5107)))*` +
    `(($B$7="")+($B$7<>"")*((${BY})>=$B$7))*` +
    `(($E$7="")+($E$7<>"")*((${BY})<=$E$7))`;

  const resultCols = '1,2,10,11,12,14,15,17,25,26,27,30';
  const noResultMsg = '{"No results — try broadening your search criteria","","","","","","","","","","",""}';

  const filterFormula =
    `=IFERROR(CHOOSECOLS(FILTER(${MP}A8:AE5107,${filterExpr}),${resultCols}),${noResultMsg})`;

  const vals = [
    { range: `'${S}'!A1`, values: [['SEARCH & FILTER — Master People']] },
    { range: `'${S}'!A2`, values: [['Results update automatically as you enter criteria. Leave a field blank or set to "All" to include all values.']] },
    // Criteria labels and inputs
    { range: `'${S}'!A3:E7`, values: [
      ['Name contains:',  '',    '', '',                ''],
      ['Family Branch:',  'All', '', 'Living Status:',  'All'],
      ['Generation:',     'All', '', 'Research Status:','All'],
      ['Evidence Conf.:', 'All', '', 'Birth Place:',    ''],
      ['Birth Year From:','',    '', 'Birth Year To:',  ''],
    ]},
    // Results section header
    { range: `'${S}'!A9`, values: [['SEARCH RESULTS']] },
    // Column headers
    { range: `'${S}'!A10:L10`, values: [['Person ID','Preferred Name','Sex / Gender','Living Status','Birth Date','Birth Place','Death Date','Death Place','Evidence Conf.','Generation','Family Branch','Research Status']] },
    // FILTER formula
    { range: `'${S}'!A11`, values: [[filterFormula]] },
  ];

  await valuesBatchUpdate(id, vals, 'search-vals');
  console.log(`✓ Search & Filter — multi-criteria live search across ${103} people`);
}

main().catch(e => { console.error(e); process.exit(1); });
