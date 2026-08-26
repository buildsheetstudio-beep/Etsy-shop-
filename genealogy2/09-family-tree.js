'use strict';
const { sheets, batchUpdate, valuesBatchUpdate, gridRange, hex, colL, C } = require('./lib');
const { id, sheetMap } = JSON.parse(require('fs').readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Family Tree'];
const S   = 'Family Tree';

// 5-generation ancestor pedigree (root → gen4 = 2× great-grandparents)
// Layout: 19 cols (A-S), 50 rows (title + gen-labels + 48 chart rows)
//   Gen0: cols A-C  | Gen1: E-G  | Gen2: I-K  | Gen3: M-O  | Gen4: Q-S
//   Connectors: D, H, L, P (narrow grey columns)
// Rows per gen4 slot: 3  →  16 slots × 3 = 48 chart rows

const GEN0 = { pid:'P-00001', name:'EMMA ROSE HARTWELL',
  extra:"b. 14 March 1988 · Boston, MA\nLiving  (Root Person)\n────────────────\nPaternal branch: HARTWELL\nMaternal branch: O'BRIEN\n────────────────\n5 generations researched\nP-00001" };

const GEN1 = [
  { pid:'P-00002', name:'James E. Hartwell',           birth:'b. 1955 Worcester, MA', death:'d. 2019 Boston, MA'   },
  { pid:'P-00003', name:"Margaret A. O'Brien Hartwell",birth:'b. 1958 Providence, RI',death:'Living'               },
];

const GEN2 = [
  { pid:'P-00004', name:'Robert C. Hartwell Sr.',           birth:'b. 1922 Springfield, MA', death:'d. 1995 Worcester, MA' },
  { pid:'P-00005', name:'Eleanor M. Chapman Hartwell',      birth:'b. 1924 Hartford, CT',    death:'d. 2003'               },
  { pid:'P-00006', name:"Patrick J. O'Brien",              birth:'b. 1925 Co. Cork, Ireland',death:'d. 1998 Providence, RI'},
  { pid:'P-00007', name:"Catherine M. Murphy O'Brien",     birth:'b. 1928 Co. Mayo, Ireland',death:'d. 2010 Providence, RI'},
];

const GEN3 = [
  { pid:'P-00008', name:'William T. Hartwell',         birth:'b. 1892 Worcester, MA',      death:'d. 1958' },
  { pid:'P-00009', name:'Agnes M. Fletcher Hartwell',  birth:'b. 1895 Springfield, MA',    death:'d. 1968' },
  { pid:'P-00010', name:'Charles H. Chapman',          birth:'b. 1890 Hartford, CT',        death:'d. 1952' },
  { pid:'P-00011', name:'Harriet L. Davies Chapman',   birth:'b. 1892 New Haven, CT',       death:'d. 1960' },
  { pid:'P-00012', name:"Michael S. O'Brien",          birth:'b. 1895 Co. Cork, Ireland',   death:'d. 1955' },
  { pid:'P-00013', name:"Brigid Shaughnessy O'Brien",  birth:'b. 1898 Co. Galway, Ireland', death:'d. 1962' },
  { pid:'P-00014', name:'Thomas P. Murphy',            birth:'b. 1898 Co. Mayo, Ireland',   death:'d. 1970' },
  { pid:'P-00015', name:'Mary C. Gallagher Murphy',    birth:'b. 1900 Co. Galway, Ireland', death:'d. 1978' },
];

const GEN4 = [
  { pid:'P-00016', name:'George A. Hartwell',          birth:'b. 1860 · d. 1938', place:'Worcester, MA'        },
  { pid:'P-00017', name:'Martha E. Holt Hartwell',     birth:'b. 1862 · d. 1940', place:'Springfield, MA'      },
  { pid:'P-00018', name:'Edward J. Fletcher',          birth:'b. 1865 · d. 1920', place:'Springfield, MA'      },
  { pid:'P-00019', name:'Sarah A. Pickles Fletcher',   birth:'b. 1868 · d. 1930', place:'Boston, MA'           },
  { pid:'P-00070', name:'Bartholomew R. Chapman',      birth:'b. 1858 · d. 1930', place:'Hartford, CT'         },
  {                name:'[Unknown] Chapman',           birth:'c. 1860',            place:'Unknown'              },
  {                name:'[Unknown] Davies',            birth:'c. 1862',            place:'Unknown'              },
  {                name:'[Unknown] Davies',            birth:'c. 1865',            place:'Unknown'              },
  { pid:'P-00020', name:"Cornelius P. O'Brien",        birth:'b. 1863 · d. 1935', place:'Co. Cork, Ireland'    },
  { pid:'P-00021', name:"Honora Crowley O'Brien",      birth:'b. 1865 · d. 1940', place:'Co. Cork, Ireland'    },
  { pid:'P-00022', name:'Daniel Shaughnessy',          birth:'b. 1870 · d. 1945', place:'Co. Galway, Ireland'  },
  { pid:'P-00023', name:'Ellen Sullivan Shaughnessy',  birth:'b. 1872 · d. 1948', place:'Co. Kerry, Ireland'   },
  { pid:'P-00024', name:'Patrick M. Murphy',           birth:'b. 1870 · d. 1945', place:'Co. Mayo, Ireland'    },
  { pid:'P-00025', name:'Bridget Ryan Murphy',         birth:'b. 1873 · d. 1950', place:'Co. Roscommon, Ireland'},
  { pid:'P-00041', name:'John M. Gallagher',           birth:'b. 1872 · d. 1948', place:'Co. Galway, Ireland'  },
  { pid:'P-00042', name:'Anne Connelly Gallagher',     birth:'b. 1875 · d. 1955', place:'Co. Galway, Ireland'  },
];

// Generation palette: darker as you go further back
const GEN_BG   = [C.primaryDeep, '#5C7F9F', C.primary, C.secondary, C.blush];
const GEN_TEXT = [C.white,       C.white,   C.white,   C.text,      C.text ];
const GEN_FS   = [13,            11,        10,        10,          9      ];

async function main() {
  const reqs = [];

  reqs.push({
    updateSheetProperties: {
      properties: {
        sheetId: SID,
        tabColor: hex(C.secondaryDeep),
        tabColorStyle: { rgbColor: hex(C.secondaryDeep) },
        gridProperties: { frozenRowCount: 2 },
      },
      fields: 'tabColor,tabColorStyle,gridProperties.frozenRowCount',
    },
  });

  // Column widths: A-S (19 cols)
  // Gen0:A-C=72×3, ConnD=10, Gen1:E-G=78×3, ConnH=10, Gen2:I-K=85×3, ConnL=10,
  // Gen3:M-O=90×3, ConnP=10, Gen4:Q-S=90×3
  const cw = [72,72,72, 10, 78,78,78, 10, 85,85,85, 10, 90,90,90, 10, 90,90,90];
  cw.forEach((px,i) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px },
      fields: 'pixelSize',
    },
  }));

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 50 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });

  // ── Merges ──────────────────────────────────────────────────────────────────
  // Title A1:S1
  reqs.push({ mergeCells: { range: gridRange(SID,0,1,0,19), mergeType:'MERGE_ALL' } });
  // Gen label cells in row 2 (only the gen columns, not connectors)
  for (const [c1,c2] of [[0,3],[4,7],[8,11],[12,15],[16,19]]) {
    reqs.push({ mergeCells: { range: gridRange(SID,1,2,c1,c2), mergeType:'MERGE_ALL' } });
  }
  // Gen0 box: A3:C50 (idx 2-49, cols 0-2)
  reqs.push({ mergeCells: { range: gridRange(SID,2,50,0,3), mergeType:'MERGE_ALL' } });
  // Gen1 boxes (2, each 24 rows)
  GEN1.forEach((_,l) => reqs.push({ mergeCells: { range: gridRange(SID, 2+l*24, 26+l*24, 4, 7), mergeType:'MERGE_ALL' } }));
  // Gen2 boxes (4, each 12 rows)
  GEN2.forEach((_,k) => reqs.push({ mergeCells: { range: gridRange(SID, 2+k*12, 14+k*12, 8, 11), mergeType:'MERGE_ALL' } }));
  // Gen3 boxes (8, each 6 rows)
  GEN3.forEach((_,j) => reqs.push({ mergeCells: { range: gridRange(SID, 2+j*6, 8+j*6, 12, 15), mergeType:'MERGE_ALL' } }));
  // Gen4 boxes (16, each 3 rows)
  GEN4.forEach((_,i) => reqs.push({ mergeCells: { range: gridRange(SID, 2+i*3, 5+i*3, 16, 19), mergeType:'MERGE_ALL' } }));

  // ── Formatting ──────────────────────────────────────────────────────────────
  // Title
  reqs.push({ repeatCell: {
    range: gridRange(SID,0,1,0,19),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryDeep),
      textFormat: { foregroundColor: hex(C.white), fontSize: 15, bold: true },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat',
  }});

  // Chart area background
  reqs.push({ repeatCell: {
    range: gridRange(SID,1,50,0,19),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat.backgroundColor',
  }});

  // Connector columns: D,H,L,P (indices 3,7,11,15) — narrow grey dividers
  for (const col of [3,7,11,15]) {
    reqs.push({ repeatCell: {
      range: gridRange(SID,1,50,col,col+1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.border) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }

  // Gen label row formatting (per generation column group)
  const genLabelRanges = [[0,3],[4,7],[8,11],[12,15],[16,19]];
  genLabelRanges.forEach(([c1,c2],g) => reqs.push({ repeatCell: {
    range: gridRange(SID,1,2,c1,c2),
    cell: { userEnteredFormat: {
      backgroundColor: hex(GEN_BG[g]),
      textFormat: { foregroundColor: hex(GEN_TEXT[g]), fontSize: 9, bold: true },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat',
  }}));

  // Gen0 box
  reqs.push({ repeatCell: {
    range: gridRange(SID,2,50,0,3),
    cell: { userEnteredFormat: {
      backgroundColor: hex(GEN_BG[0]),
      textFormat: { foregroundColor: hex(GEN_TEXT[0]), fontSize: GEN_FS[0], bold: true },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
    }},
    fields: 'userEnteredFormat',
  }});

  // Gen1 boxes
  GEN1.forEach((_,l) => reqs.push({ repeatCell: {
    range: gridRange(SID, 2+l*24, 26+l*24, 4, 7),
    cell: { userEnteredFormat: {
      backgroundColor: hex(GEN_BG[1]),
      textFormat: { foregroundColor: hex(GEN_TEXT[1]), fontSize: GEN_FS[1], bold: true },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
    }},
    fields: 'userEnteredFormat',
  }}));

  // Gen2 boxes
  GEN2.forEach((_,k) => reqs.push({ repeatCell: {
    range: gridRange(SID, 2+k*12, 14+k*12, 8, 11),
    cell: { userEnteredFormat: {
      backgroundColor: hex(GEN_BG[2]),
      textFormat: { foregroundColor: hex(GEN_TEXT[2]), fontSize: GEN_FS[2], bold: true },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
    }},
    fields: 'userEnteredFormat',
  }}));

  // Gen3 boxes
  GEN3.forEach((_,j) => reqs.push({ repeatCell: {
    range: gridRange(SID, 2+j*6, 8+j*6, 12, 15),
    cell: { userEnteredFormat: {
      backgroundColor: hex(GEN_BG[3]),
      textFormat: { foregroundColor: hex(GEN_TEXT[3]), fontSize: GEN_FS[3], bold: true },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
    }},
    fields: 'userEnteredFormat',
  }}));

  // Gen4 boxes — unknown ancestors get neutral color + italic
  GEN4.forEach((p,i) => {
    const isUnknown = !p.pid;
    reqs.push({ repeatCell: {
      range: gridRange(SID, 2+i*3, 5+i*3, 16, 19),
      cell: { userEnteredFormat: {
        backgroundColor: hex(isUnknown ? C.neutral : GEN_BG[4]),
        textFormat: {
          foregroundColor: hex(GEN_TEXT[4]),
          fontSize: GEN_FS[4],
          bold: !isUnknown,
          italic: isUnknown,
        },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
      }},
      fields: 'userEnteredFormat',
    }});
  });

  // Box borders (bg-color outline to separate boxes visually)
  const sep = { style: 'SOLID_MEDIUM', color: hex(C.bg) };
  GEN2.forEach((_,k) => reqs.push({ updateBorders: { range: gridRange(SID, 2+k*12, 14+k*12, 8, 11), top:sep,bottom:sep,left:sep,right:sep } }));
  GEN3.forEach((_,j) => reqs.push({ updateBorders: { range: gridRange(SID, 2+j*6, 8+j*6, 12, 15), top:sep,bottom:sep,left:sep,right:sep } }));
  GEN4.forEach((_,i) => reqs.push({ updateBorders: { range: gridRange(SID, 2+i*3, 5+i*3, 16, 19), top:sep,bottom:sep,left:sep,right:sep } }));

  await batchUpdate(id, reqs, 'tree-fmt');

  // ── Values ───────────────────────────────────────────────────────────────────
  const boxText = (p) => `${p.name}\n${p.birth}\n${p.death}\n${p.pid}`;
  const box4txt = (p) => p.pid
    ? `${p.name}\n${p.birth}\n${p.place}\n${p.pid}`
    : `${p.name}\n${p.birth}\n${p.place}`;

  const vals = [
    { range: `'${S}'!A1`, values: [["HARTWELL – O'BRIEN FAMILY PEDIGREE CHART  •  Five Generations"]] },
    { range: `'${S}'!A2`,  values: [['ROOT (YOU)']] },
    { range: `'${S}'!E2`,  values: [['PARENTS — GENERATION 1']] },
    { range: `'${S}'!I2`,  values: [['GRANDPARENTS — GENERATION 2']] },
    { range: `'${S}'!M2`,  values: [['GREAT-GRANDPARENTS — GEN. 3']] },
    { range: `'${S}'!Q2`,  values: [['2× GREAT-GRANDPARENTS — GEN. 4']] },
    // Gen0
    { range: `'${S}'!A3`,  values: [[GEN0.extra]] },
  ];

  // Gen1 — each box starts at rows 3 and 27 (l*24+3)
  GEN1.forEach((p,l) => {
    const r = 3 + l*24;
    vals.push({ range: `'${S}'!${colL(4)}${r}`, values: [[boxText(p)]] });
  });

  // Gen2 — boxes start at rows 3,15,27,39 (k*12+3)
  GEN2.forEach((p,k) => {
    const r = 3 + k*12;
    vals.push({ range: `'${S}'!${colL(8)}${r}`, values: [[boxText(p)]] });
  });

  // Gen3 — boxes start at rows 3,9,15,21,27,33,39,45 (j*6+3)
  GEN3.forEach((p,j) => {
    const r = 3 + j*6;
    vals.push({ range: `'${S}'!${colL(12)}${r}`, values: [[boxText(p)]] });
  });

  // Gen4 — boxes start at rows 3,6,9,...,48 (i*3+3)
  GEN4.forEach((p,i) => {
    const r = 3 + i*3;
    vals.push({ range: `'${S}'!${colL(16)}${r}`, values: [[box4txt(p)]] });
  });

  await valuesBatchUpdate(id, vals, 'tree-vals');
  console.log(`✓ Family Tree — 5-generation pedigree, ${1+2+4+8+16} ancestor positions`);
}

main().catch(e => { console.error(e); process.exit(1); });
