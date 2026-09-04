'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, colL, C } = require('./lib');
const { id, sheetMap } = JSON.parse(require('fs').readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Family Tree'];
const S   = 'Family Tree';

// ── Layout constants ─────────────────────────────────────────────────────────
// 32 columns (A-AF) × 55px = 1760px total
// Gen4 = 16 boxes × 2 cols each
// Gen3 = 8 boxes × 4 cols each
// Gen2 = 4 boxes × 8 cols each
// Gen1 = 2 boxes × 16 cols each
// Gen0 = 1 box × 32 cols
const NCOLS = 32;
const CW    = 55;

// Row layout (0-indexed):
// 0        Title (48px)
// 1        Gen4 generation label (18px)
// 2-4      Gen4 person boxes (3×20px)
// 5        Connector row (8px)
// 6        Gen3 label (18px)
// 7-11     Gen3 person boxes (5×22px)
// 12       Connector row (8px)
// 13       Gen2 label (18px)
// 14-20    Gen2 person boxes (7×22px)
// 21       Connector row (8px)
// 22       Gen1 label (18px)
// 23-33    Gen1 person boxes (11×24px)
// 34       Connector row (8px)
// 35       Gen0 label (18px)
// 36-44    Gen0 root box (9×40px)

const TITLE_ROW   = 0;
const G4_LBL_ROW  = 1;
const G4_BOX_R1   = 2;  const G4_BOX_R2  = 5;   // rows 2-4 (exclusive end 5)
const G4_CONN_ROW = 5;
const G3_LBL_ROW  = 6;
const G3_BOX_R1   = 7;  const G3_BOX_R2  = 12;  // rows 7-11
const G3_CONN_ROW = 12;
const G2_LBL_ROW  = 13;
const G2_BOX_R1   = 14; const G2_BOX_R2  = 21;  // rows 14-20
const G2_CONN_ROW = 21;
const G1_LBL_ROW  = 22;
const G1_BOX_R1   = 23; const G1_BOX_R2  = 34;  // rows 23-33
const G1_CONN_ROW = 34;
const G0_LBL_ROW  = 35;
const G0_BOX_R1   = 36; const G0_BOX_R2  = 45;  // rows 36-44

// ── Colors ───────────────────────────────────────────────────────────────────
const BOX_BG = [
  C.primary,       // Gen4 — light steel blue
  C.secondary,     // Gen3 — light sage
  C.blush,         // Gen2 — light rose
  C.wheat,         // Gen1 — light wheat
  C.primaryDeep,   // Gen0 — Emma (prominent dark blue)
];
const BOX_FG = [C.white, C.text, C.text, C.text, C.white];
const LBL_BG = [C.primaryDeep, '#5C7F9F', C.primary, C.secondary, C.primaryDeep];

const CONN_COLOR = '#6A8FAF'; // C.primaryDeep — tree branch lines

// ── People data ──────────────────────────────────────────────────────────────
// Ordered left-to-right, matching pedigree position (paternal side left)

// Each Gen4 box: 2 cols wide. 16 people across = 32 cols.
// Pairs: [0-1],[2-3]→Gen3[0]; [4-5],[6-7]→Gen3[1]; etc.
const GEN4 = [
  // William Hartwell's parents (→ Robert → James → Emma, left branch)
  { pid:'P-00016', name:'George A.\nHartwell',          born:'b. c.1860',  died:'d. 1938\nWorcester, MA' },
  { pid:'P-00017', name:'Martha Holt\nHartwell',         born:'b. c.1862',  died:'d. 1940\nSpringfield, MA' },
  // Agnes Fletcher's parents
  { pid:'P-00018', name:'Edward J.\nFletcher',           born:'b. c.1865',  died:'d. c.1920\nCoventry, England' },
  { pid:'P-00019', name:'Sarah Pickles\nFletcher',       born:'b. c.1868',  died:'d. c.1930\nBoston, MA' },
  // Charles Chapman's parents (unknown)
  { pid:'P-00070', name:'Bartholomew R.\nChapman',       born:'b. c.1858',  died:'d. c.1930\nHartford, CT' },
  { pid:null,      name:'[Unknown]\nChapman',            born:'c. 1860',    died:'Unknown' },
  // Harriet Davies' parents (unknown)
  { pid:null,      name:'[Unknown]\nDavies',             born:'c. 1862',    died:'Unknown' },
  { pid:null,      name:'[Unknown]\nDavies',             born:'c. 1865',    died:'Unknown' },
  // Michael O'Brien's parents
  { pid:'P-00020', name:"Cornelius P.\nO'Brien",         born:'b. c.1863',  died:'d. 1935\nCo. Cork, Ireland' },
  { pid:'P-00021', name:"Honora Crowley\nO'Brien",       born:'b. c.1865',  died:'d. 1940\nCo. Cork, Ireland' },
  // Brigid Shaughnessy's parents
  { pid:'P-00022', name:'Daniel\nShaughnessy',           born:'b. c.1870',  died:'d. 1945\nCo. Cork, Ireland' },
  { pid:'P-00023', name:'Ellen Sullivan\nShaughnessy',   born:'b. c.1872',  died:'d. 1948\nCo. Kerry, Ireland' },
  // Thomas Murphy's parents
  { pid:'P-00024', name:'Patrick M.\nMurphy',            born:'b. c.1870',  died:'d. 1945\nCo. Mayo, Ireland' },
  { pid:'P-00025', name:'Bridget Ryan\nMurphy',          born:'b. c.1873',  died:'d. 1950\nCo. Roscommon' },
  // Mary Gallagher's parents
  { pid:'P-00041', name:'John M.\nGallagher',            born:'b. c.1872',  died:'d. 1948\nCo. Galway, Ireland' },
  { pid:'P-00042', name:'Anne Connelly\nGallagher',      born:'b. c.1875',  died:'d. 1955\nCo. Galway, Ireland' },
];

const GEN3 = [
  { pid:'P-00008', name:"William H.\nHartwell",         born:'b. 1888\nCoventry, England', died:'d. 1952\nSpringfield, MA' },
  { pid:'P-00009', name:"Agnes Fletcher\nHartwell",     born:'b. 1892\nCoventry, England', died:'d. 1960\nSpringfield, MA' },
  { pid:'P-00010', name:"Charles F.\nChapman",          born:'b. c.1895\nHartford, CT',    died:'d. 1958\nSpringfield, MA' },
  { pid:'P-00011', name:"Harriet Davies\nChapman",      born:'b. 1898\nHartford, CT',      died:'d. 1971\nSpringfield, MA' },
  { pid:'P-00012', name:"Michael P.\nO'Brien",          born:"b. c.1882\nCo. Cork, Ireland",died:'d. 1944\nProvidence, RI' },
  { pid:'P-00013', name:"Brigid Shaughnessy\nO'Brien",  born:"b. c.1886\nCo. Cork, Ireland",died:'d. 1949\nProvidence, RI' },
  { pid:'P-00014', name:"Thomas F.\nMurphy",            born:'b. c.1886\nCo. Clare, Ireland',died:'d. 1938\nProvidence, RI' },
  { pid:'P-00015', name:"Mary Gallagher\nMurphy",       born:'b. c.1889\nCo. Clare, Ireland',died:'d. 1955\nProvidence, RI' },
];

const GEN2 = [
  { pid:'P-00004', name:"Robert Charles Hartwell Sr.",   born:'b. 1922\nSpringfield, MA',   died:'d. 1995\nWorcester, MA' },
  { pid:'P-00005', name:"Eleanor Chapman Hartwell",      born:'b. 1926\nSpringfield, MA',   died:'d. 2003\nWorcester, MA' },
  { pid:'P-00006', name:"Patrick J. O'Brien",            born:'b. 1920\nProvidence, RI',    died:'d. 1998\nProvidence, RI' },
  { pid:'P-00007', name:"Catherine Murphy O'Brien",      born:'b. 1924\nProvidence, RI',    died:'d. 2007\nProvidence, RI' },
];

const GEN1 = [
  { pid:'P-00002', name:"James Edward Hartwell",
    born:'b. June 22, 1955  •  Worcester, MA', marr:'m. 1981  •  Worcester, MA',
    died:'d. Aug 5, 2019  •  Boston, MA',      occ:'Accountant  •  P-00002' },
  { pid:'P-00003', name:"Margaret Anne O'Brien Hartwell",
    born:'b. Nov 3, 1958  •  Providence, RI',  marr:'m. 1981  •  Worcester, MA',
    died:'Living',                              occ:'School Teacher  •  P-00003' },
];

const GEN0 = { pid: 'P-00001' };

// ── Formula helpers: look up data live from Master People ────────────────────
const _MP = `'Master People'`;
const _R  = `$8:$5107`;
function _col(c)        { return `${_MP}!$${c}${_R}`; }
function _m(pid)        { return `MATCH("${pid}",${_col('A')},0)`; }
function _get(pid, c)   { return `IFERROR(INDEX(${_col(c)},${_m(pid)}),"")`;  }
// Format a date cell to "mmm d, yyyy"; fall back to raw text for estimates
function _dtFmt(pid, c) {
  const raw = _get(pid, c);
  return `IFERROR(TEXT(DATEVALUE(${raw}),"mmm d, yyyy"),${raw})`;
}
// Year only for tiny boxes; fall back to raw
function _yrFmt(pid, c) {
  const raw = _get(pid, c);
  return `IFERROR(TEXT(DATEVALUE(${raw}),"yyyy"),${raw})`;
}
// "b. Date • Place"  or ""
function _bLine(pid) {
  const dt = _dtFmt(pid,'L'), pl = _get(pid,'N');
  return `IF(${dt}<>"","b. "&${dt}&IF(${pl}<>""," • "&${pl},""),"")`;
}
// "d. Date • Place"  or ""
function _dLine(pid) {
  const dt = _dtFmt(pid,'O'), pl = _get(pid,'Q');
  return `IF(${dt}<>"","d. "&${dt}&IF(${pl}<>""," • "&${pl},""),"")`;
}
// "Living" or "d. Date • Place"
function _ldLine(pid) {
  const status = _get(pid,'K'), dt = _dtFmt(pid,'O'), pl = _get(pid,'Q');
  return `IF(${status}="Living","Living",IF(${dt}<>"","d. "&${dt}&IF(${pl}<>""," • "&${pl},""),""))`;
}
// Concatenate formula parts with CHAR(10) newlines (wrap-strategy renders them)
function _fml(...parts) { return '=' + parts.filter(p => p).join('&CHAR(10)&'); }

// ── Per-generation content builders ──────────────────────────────────────────
// Gen4: tiny box — name, birth year, death year, PID
function buildG4(p) {
  if (!p.pid) return `${p.name}\n${p.born}\n${p.died}`;
  const { pid } = p;
  const by = _yrFmt(pid,'L'), dy = _yrFmt(pid,'O');
  return _fml(
    _get(pid,'B'),
    `IF(${by}<>"","b. "&${by},"")`,
    `IF(${dy}<>"","d. "&${dy},"")`,
    `"${pid}"`
  );
}
// Gen3: name, birth+place, death+place, PID
function buildG3(p) {
  const { pid } = p;
  return _fml(_get(pid,'B'), _bLine(pid), _dLine(pid), `"${pid}"`);
}
// Gen2: name, birth+place, death+place, occupation, PID
function buildG2(p) {
  const { pid } = p;
  return _fml(_get(pid,'B'), _bLine(pid), _dLine(pid), _get(pid,'S'), `"${pid}"`);
}
// Gen1: name, birth+place, living/death, occupation, PID
function buildG1(p) {
  const { pid } = p;
  return _fml(_get(pid,'B'), _bLine(pid), _ldLine(pid), _get(pid,'S'), `"${pid}"`);
}
// Gen0 root: name, tagline, birth+place, status, PID
function buildG0(p) {
  const { pid } = p;
  return _fml(
    _get(pid,'B'),
    `"Root Person  •  Hartwell & O'Brien Family Tree  •  Five Generations Researched"`,
    _bLine(pid),
    _get(pid,'K'),
    `"${pid}"`
  );
}

// ── Helper: connector cells ───────────────────────────────────────────────────
// For a connector row at rowIdx, color specific cells to draw tree branch lines
// Each generation transition has a specific connector pattern:
// Gen4→Gen3: 8 groups of 4 cols; center at [4k+1, 4k+2]
// Gen3→Gen2: 4 groups of 8 cols; center at [8k+3, 8k+4]
// Gen2→Gen1: 2 groups of 16 cols; center at [16k+7, 16k+8]
// Gen1→Gen0: 1 group; center at [15, 16]
function connectorCells(rowIdx, centers) {
  const reqs = [];
  // Background of the whole connector row = C.bg
  reqs.push({ repeatCell: {
    range: gridRange(SID, rowIdx, rowIdx+1, 0, NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat.backgroundColor',
  }});
  // Color the connector cells
  for (const [c1, c2] of centers) {
    reqs.push({ repeatCell: {
      range: gridRange(SID, rowIdx, rowIdx+1, c1, c2+1),
      cell: { userEnteredFormat: { backgroundColor: hex(CONN_COLOR) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }
  return reqs;
}

async function main() {
  const reqs = [];
  const vals = [];

  // ── Sheet properties ────────────────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: {
      sheetId: SID,
      tabColor: hex(C.secondaryDeep),
      tabColorStyle: { rgbColor: hex(C.secondaryDeep) },
      gridProperties: { frozenRowCount: 1, columnCount: NCOLS },
    },
    fields: 'tabColor,tabColorStyle,gridProperties.frozenRowCount,gridProperties.columnCount',
  }});

  // Unmerge all from previous layout
  reqs.push({ unmergeCells: { range: gridRange(SID, 0, 500, 0, 33) } });

  // Column widths
  for (let i = 0; i < NCOLS; i++) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: CW }, fields: 'pixelSize',
    }});
  }

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: TITLE_ROW, endIndex: TITLE_ROW+1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: G4_LBL_ROW, endIndex: G4_LBL_ROW+1 }, properties: { pixelSize: 18 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: G4_BOX_R1, endIndex: G4_BOX_R2 }, properties: { pixelSize: 20 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: G4_CONN_ROW, endIndex: G4_CONN_ROW+1 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: G3_LBL_ROW, endIndex: G3_LBL_ROW+1 }, properties: { pixelSize: 18 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: G3_BOX_R1, endIndex: G3_BOX_R2 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: G3_CONN_ROW, endIndex: G3_CONN_ROW+1 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: G2_LBL_ROW, endIndex: G2_LBL_ROW+1 }, properties: { pixelSize: 18 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: G2_BOX_R1, endIndex: G2_BOX_R2 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: G2_CONN_ROW, endIndex: G2_CONN_ROW+1 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: G1_LBL_ROW, endIndex: G1_LBL_ROW+1 }, properties: { pixelSize: 18 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: G1_BOX_R1, endIndex: G1_BOX_R2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: G1_CONN_ROW, endIndex: G1_CONN_ROW+1 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: G0_LBL_ROW, endIndex: G0_LBL_ROW+1 }, properties: { pixelSize: 18 }, fields: 'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: G0_BOX_R1, endIndex: G0_BOX_R2 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // ── Title ────────────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID,TITLE_ROW,TITLE_ROW+1,0,NCOLS), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(SID,TITLE_ROW,TITLE_ROW+1,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryDeep),
      textFormat: { foregroundColor: hex(C.white), fontSize: 16, bold: true },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat',
  }});
  vals.push({ range:`'${S}'!A1`, values:[["HARTWELL – O'BRIEN FAMILY PEDIGREE CHART  •  Five Generations  •  Root Person: Emma Rose Hartwell"]] });

  // ── Whole sheet background ────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID,1,G0_BOX_R2,0,NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat.backgroundColor',
  }});

  // ── Generation label helper ───────────────────────────────────────────────────
  function genLabel(rowIdx, text, bg, fg) {
    reqs.push({ mergeCells: { range: gridRange(SID,rowIdx,rowIdx+1,0,NCOLS), mergeType:'MERGE_ALL' }});
    reqs.push({ repeatCell: {
      range: gridRange(SID,rowIdx,rowIdx+1,0,NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(fg), fontSize: 9, bold: true },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat',
    }});
    vals.push({ range:`'${S}'!A${rowIdx+1}`, values:[[text]] });
  }

  genLabel(G4_LBL_ROW, '2× GREAT-GRANDPARENTS  (Generation IV)', LBL_BG[0], C.white);
  genLabel(G3_LBL_ROW, 'GREAT-GRANDPARENTS  (Generation III)',   LBL_BG[1], C.white);
  genLabel(G2_LBL_ROW, 'GRANDPARENTS  (Generation II)',          LBL_BG[2], C.white);
  genLabel(G1_LBL_ROW, 'PARENTS  (Generation I)',                LBL_BG[3], C.text);
  genLabel(G0_LBL_ROW, 'ROOT PERSON  (Generation 0)',            LBL_BG[4], C.white);

  // ── Box helper ────────────────────────────────────────────────────────────────
  const BOX_BORDER = (bg) => ({ style: 'SOLID', color: hex(C.bg) });

  function personBox(r1, r2, c1, c2, bg, fg, fontSize, bold, content) {
    reqs.push({ mergeCells: { range: gridRange(SID,r1,r2,c1,c2), mergeType:'MERGE_ALL' }});
    reqs.push({ repeatCell: {
      range: gridRange(SID,r1,r2,c1,c2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(fg), fontSize, bold },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        wrapStrategy: 'WRAP',
      }},
      fields: 'userEnteredFormat',
    }});
    const sep = { style: 'SOLID_MEDIUM', color: hex(C.bg) };
    reqs.push({ updateBorders: {
      range: gridRange(SID,r1,r2,c1,c2),
      top: sep, bottom: sep, left: sep, right: sep,
    }});
    vals.push({ range:`'${S}'!${colL(c1)}${r1+1}`, values:[[content]] });
  }

  // ── GEN4: 16 boxes × 2 cols each ─────────────────────────────────────────────
  GEN4.forEach((p, i) => {
    const c1 = i * 2;
    const isUnk = !p.pid;
    const bg = isUnk ? C.neutral : BOX_BG[0];
    personBox(G4_BOX_R1, G4_BOX_R2, c1, c1+2, bg, isUnk ? C.secText : BOX_FG[0], 7, !isUnk, buildG4(p));
  });

  // ── GEN3: 8 boxes × 4 cols each ──────────────────────────────────────────────
  GEN3.forEach((p, i) => {
    personBox(G3_BOX_R1, G3_BOX_R2, i*4, i*4+4, BOX_BG[1], BOX_FG[1], 8, true, buildG3(p));
  });

  // ── GEN2: 4 boxes × 8 cols each ──────────────────────────────────────────────
  GEN2.forEach((p, i) => {
    personBox(G2_BOX_R1, G2_BOX_R2, i*8, i*8+8, BOX_BG[2], BOX_FG[2], 9, true, buildG2(p));
  });

  // ── GEN1: 2 boxes × 16 cols each ─────────────────────────────────────────────
  GEN1.forEach((p, i) => {
    personBox(G1_BOX_R1, G1_BOX_R2, i*16, i*16+16, BOX_BG[3], BOX_FG[3], 10, true, buildG1(p));
  });

  // ── GEN0: Emma — full width ───────────────────────────────────────────────────
  personBox(G0_BOX_R1, G0_BOX_R2, 0, NCOLS, BOX_BG[4], BOX_FG[4], 13, true, buildG0(GEN0));

  // ── Connector rows ────────────────────────────────────────────────────────────
  // Gen4→Gen3: 8 groups of 4 cols, center cells = [4k+1, 4k+2]
  const g4g3centers = Array.from({length:8}, (_,k) => [4*k+1, 4*k+2]);
  reqs.push(...connectorCells(G4_CONN_ROW, g4g3centers));

  // Gen3→Gen2: 4 groups of 8 cols, center cells = [8k+3, 8k+4]
  const g3g2centers = Array.from({length:4}, (_,k) => [8*k+3, 8*k+4]);
  reqs.push(...connectorCells(G3_CONN_ROW, g3g2centers));

  // Gen2→Gen1: 2 groups of 16 cols, center cells = [16k+7, 16k+8]
  const g2g1centers = Array.from({length:2}, (_,k) => [16*k+7, 16*k+8]);
  reqs.push(...connectorCells(G2_CONN_ROW, g2g1centers));

  // Gen1→Gen0: center cells = [15, 16]
  reqs.push(...connectorCells(G1_CONN_ROW, [[15,16]]));

  await batchUpdate(id, reqs, 'tree-fmt');
  await valuesBatchUpdate(id, vals, 'tree-vals');
  console.log(`✓ Family Tree — top-to-bottom pedigree, 5 generations, 31 ancestor positions`);
}

main().catch(e => { console.error(e); process.exit(1); });
