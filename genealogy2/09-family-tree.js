'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, colL, C } = require('./lib');
const { id, sheetMap } = JSON.parse(require('fs').readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Family Tree'];
const S   = 'Family Tree';

const BPR  = 4;          // boxes per row
const BCOL = 3;          // columns per box
const BROW = 3;          // rows per box
const TCOL = BPR * BCOL; // 12

// [pid, displayName, bornLine, diedLine, branch]
const SECTIONS = [
  {
    title: '4× GREAT-GRANDPARENTS',
    sub: 'c. 1775–1815  •  Earliest confirmed ancestors',
    hBg: C.primaryDeep, hFg: C.white,
    people: [
      ['P-00061', 'Thaddeus O\'Brien',           'b. c.1775 • Co. Cork, Ireland',     'd. c.1840',                     "O'Brien"],
      ['P-00062', 'Catherine McCarthy O\'Brien', 'b. c.1778 • Co. Cork, Ireland',     'd. c.1842',                     "O'Brien"],
      ['P-00063', 'Thomas Hartwell',             'b. c.1780 • Warwickshire, England', 'd. c.1850',                     'Hartwell'],
      ['P-00064', 'Jane Booth Hartwell',         'b. c.1785 • Warwickshire, England', 'd. c.1855',                     'Hartwell'],
      ['P-00084', 'Hannah Turner',               'b. c.1790 • Warwickshire, England', 'd. c.1860',                     'Hartwell'],
      ['P-00085', 'John Turner',                 'b. c.1785 • Warwickshire, England', 'd. c.1855',                     'Hartwell'],
      ['P-00091', 'Alice Booth',                 'b. c.1810 • Warwickshire, England', 'd. c.1878',                     'Hartwell'],
    ],
  },
  {
    title: '3× GREAT-GRANDPARENTS',
    sub: 'c. 1808–1860',
    hBg: '#5C7F9F', hFg: C.white,
    people: [
      ['P-00047', 'Frederick G. Hartwell',       'b. c.1815 • Warwickshire, England', 'd. c.1880 • Coventry, England', 'Hartwell'],
      ['P-00048', 'Ann Turner Hartwell',         'b. c.1820 • Warwickshire, England', 'd. c.1885 • Coventry, England', 'Hartwell'],
      ['P-00065', 'Francis M. Holt',             'b. c.1820 • Coventry, England',     'd. c.1890',                     'Hartwell'],
      ['P-00067', 'Alice V. Holt Fletcher',      'b. c.1825 • Birmingham, England',   'd. c.1895',                     'Hartwell'],
      ['P-00090', 'Albert F. Holt',              'b. c.1855 • Coventry, England',     'd. c.1920',                     'Hartwell'],
      ['P-00102', 'Cecil E. Hartwell',           'b. c.1854 • Coventry, England',     'd. c.1932',                     'Hartwell'],
      ['P-00049', 'Jeremiah O\'Brien',           'b. c.1808 • Co. Cork, Ireland',     'd. 1855 • Co. Cork, Ireland',   "O'Brien"],
      ['P-00050', 'Mary Driscoll O\'Brien',      'b. c.1812 • Co. Cork, Ireland',     'd. 1848 • Co. Cork, Ireland',   "O'Brien"],
      ['P-00059', '[Unknown] Crowley',           'b. Unknown • Co. Cork, Ireland',    'Deceased — records lost',       "O'Brien"],
      ['P-00066', 'Nora O\'Connell Crowley',     'b. c.1825 • Co. Cork, Ireland',     'd. c.1890',                     "O'Brien"],
    ],
  },
  {
    title: '2× GREAT-GRANDPARENTS',
    sub: 'c. 1845–1875',
    hBg: C.primary, hFg: C.white,
    people: [
      ['P-00016', 'George T. Hartwell',          'b. c.1850 • Coventry, England',     'd. c.1920 • Coventry, England', 'Hartwell'],
      ['P-00017', 'Martha Holt Hartwell',        'b. c.1854 • Coventry, England',     'd. c.1925 • Coventry, England', 'Hartwell'],
      ['P-00018', 'Edward J. Fletcher',          'b. c.1858 • Birmingham, England',   'd. c.1930 • Coventry, England', 'Hartwell'],
      ['P-00019', 'Sarah Pickles Fletcher',      'b. c.1862 • Birmingham, England',   'd. c.1935 • Coventry, England', 'Hartwell'],
      ['P-00043', 'Henry W. Fletcher',           'b. c.1860 • Birmingham, England',   'd. c.1928 • Coventry, England', 'Hartwell'],
      ['P-00020', 'Cornelius O\'Brien',          'b. c.1845 • Co. Cork, Ireland',     'd. 1895 • Co. Cork, Ireland',   "O'Brien"],
      ['P-00021', 'Honora Crowley O\'Brien',     'b. c.1848 • Co. Cork, Ireland',     'd. 1900 • Co. Cork, Ireland',   "O'Brien"],
      ['P-00022', 'Daniel Shaughnessy',          'b. c.1850 • Co. Cork, Ireland',     'd. 1908 • Co. Cork, Ireland',   "O'Brien"],
      ['P-00023', 'Ellen Sullivan Shaughnessy',  'b. c.1852 • Co. Cork, Ireland',     'd. 1910 • Co. Cork, Ireland',   "O'Brien"],
      ['P-00024', 'Patrick Murphy',              'b. c.1852 • Co. Clare, Ireland',    'd. 1910 • Co. Clare, Ireland',  "O'Brien"],
      ['P-00025', 'Bridget Ryan Murphy',         'b. c.1855 • Co. Clare, Ireland',    'd. 1912 • Co. Clare, Ireland',  "O'Brien"],
      ['P-00041', 'John M. Gallagher',           'b. c.1855 • Co. Clare, Ireland',    'd. c.1920 • Co. Clare, Ireland',"O'Brien"],
      ['P-00042', 'Anne Connelly Gallagher',     'b. c.1858 • Co. Clare, Ireland',    'd. c.1925 • Co. Clare, Ireland',"O'Brien"],
    ],
  },
  {
    title: 'GREAT-GRANDPARENTS',
    sub: 'c. 1880–1900',
    hBg: C.secondary, hFg: C.text,
    people: [
      ['P-00008', 'William H. Hartwell',         'b. 1888 • Coventry, England',       'd. 1952 • Springfield, MA',     'Hartwell'],
      ['P-00009', 'Agnes Fletcher Hartwell',     'b. 1892 • Coventry, England',       'd. 1960 • Springfield, MA',     'Hartwell'],
      ['P-00010', 'Charles F. Chapman',          'b. c.1895 • Hartford, CT',          'd. 1958 • Springfield, MA',     'Hartwell'],
      ['P-00011', 'Harriet Davies Chapman',      'b. 1898 • Hartford, CT',            'd. 1971 • Springfield, MA',     'Hartwell'],
      ['P-00039', 'Arthur R. Chapman',           'b. c.1898 • Hartford, CT',          'd. 1962 • Springfield, MA',     'Hartwell'],
      ['P-00012', 'Michael P. O\'Brien',         'b. c.1882 • Co. Cork, Ireland',     'd. 1944 • Providence, RI',      "O'Brien"],
      ['P-00013', 'Brigid Shaughnessy O\'Brien', 'b. c.1886 • Co. Cork, Ireland',     'd. 1949 • Providence, RI',      "O'Brien"],
      ['P-00014', 'Thomas F. Murphy',            'b. c.1886 • Co. Clare, Ireland',    'd. 1938 • Providence, RI',      "O'Brien"],
      ['P-00015', 'Mary Gallagher Murphy',       'b. c.1889 • Co. Clare, Ireland',    'd. 1955 • Providence, RI',      "O'Brien"],
    ],
  },
  {
    title: 'GRANDPARENTS',
    sub: 'c. 1920–1928',
    hBg: C.aqua, hFg: C.text,
    people: [
      ['P-00004', 'Robert C. Hartwell Sr.',      'b. 1922 • Springfield, MA',         'd. 1995 • Worcester, MA',       'Hartwell'],
      ['P-00005', 'Eleanor Chapman Hartwell',    'b. 1926 • Springfield, MA',         'd. 2003 • Worcester, MA',       'Hartwell'],
      ['P-00006', 'Patrick J. O\'Brien',         'b. 1920 • Providence, RI',          'd. 1998 • Providence, RI',      "O'Brien"],
      ['P-00007', 'Catherine Murphy O\'Brien',   'b. 1924 • Providence, RI',          'd. 2007 • Providence, RI',      "O'Brien"],
    ],
  },
  {
    title: 'PARENTS',
    sub: 'c. 1955–1958',
    hBg: C.lavender, hFg: C.text,
    people: [
      ['P-00002', 'James Edward Hartwell',       'b. 1955 • Worcester, MA',           'd. 2019 • Boston, MA',          'Hartwell'],
      ['P-00003', 'Margaret O\'Brien Hartwell',  'b. 1958 • Providence, RI',          'Living',                        "O'Brien"],
    ],
  },
  {
    title: 'ROOT GENERATION — Emma & Siblings',
    sub: 'c. 1975–1993',
    hBg: C.secondaryDeep, hFg: C.white,
    people: [
      ['P-00001', '★ EMMA ROSE HARTWELL',        'b. 1988-03-14 • Boston, MA',        'Living  (Root Person)',          'root'],
      ['P-00026', 'Thomas James Hartwell',       'b. 1983 • Worcester, MA',           'Living',                        'Hartwell'],
      ['P-00027', 'Sarah Hartwell Novak',        'b. 1990 • Worcester, MA',           'Living',                        'Hartwell'],
      ['P-00038', 'Helen Hartwell (adopted)',    'b. c.1993',                         'Living',                        'Hartwell'],
      ['P-00045', 'Daniel Paul Hartwell',        'b. 1975 • Hartford, CT',            'Living  (half-brother)',         'Hartwell'],
    ],
  },
  {
    title: 'DESCENDANTS',
    sub: 'c. 1978–2015',
    hBg: C.wheat, hFg: C.text,
    people: [
      ['P-00034', 'Kevin Richard Hartwell',      'b. 1978 • Worcester, MA',           'Living',                        'Hartwell'],
      ['P-00035', 'Jennifer Hartwell Walsh',     'b. 1980 • Worcester, MA',           'Living',                        'Hartwell'],
      ['P-00057', 'Sean Michael O\'Brien Jr.',   'b. 1980 • Providence, RI',          'Living',                        "O'Brien"],
      ['P-00092', 'Declan Patrick O\'Brien',     'b. 1982 • Providence, RI',          'Living',                        "O'Brien"],
      ['P-00055', 'Fiona O\'Brien Donovan',      'b. 1983 • Providence, RI',          'Living',                        "O'Brien"],
      ['P-00056', 'Connor Patrick Donovan',      'b. 1985 • Providence, RI',          'Living',                        "O'Brien"],
      ['P-00029', 'Oliver James Hartwell',       'b. 2012 • Chicago, IL',             'Living',                        'Hartwell'],
      ['P-00030', 'Sophia Grace Hartwell',       'b. 2015 • Chicago, IL',             'Living',                        'Hartwell'],
      ['P-00094', 'Liam Connor O\'Brien',        'b. 2010 • Providence, RI',          'Living',                        "O'Brien"],
      ['P-00095', 'Siobhan Grace O\'Brien',      'b. 2013 • Providence, RI',          'Living',                        "O'Brien"],
      ['P-00098', 'Patrick Martin O\'Brien',     'b. 1912 • Providence, RI',          'd. 1985 • Providence, RI',      "O'Brien"],
      ['P-00099', 'Evelyn McCarthy O\'Brien',    'b. 1915 • Providence, RI',          'd. 1992 • Providence, RI',      "O'Brien"],
    ],
  },
  {
    title: 'SPOUSES & PARTNERS',
    sub: 'Married-in and partnered family members',
    hBg: C.blush, hFg: C.text,
    people: [
      ['P-00028', 'Linda Patel Hartwell',        'b. 1985 • Chicago, IL',             'Living  (Thomas\'s spouse)',     'Hartwell'],
      ['P-00031', 'Peter Novak',                 'b. 1988 • Boston, MA',              'Living  (Sarah\'s husband)',     'Hartwell'],
      ['P-00058', 'Nicole Rossi O\'Brien',       'b. 1982 • Providence, RI',          'Living  (Sean Jr.\'s wife)',     "O'Brien"],
      ['P-00093', 'Aisling Murphy O\'Brien',     'b. 1984 • Providence, RI',          'Living  (Declan\'s wife)',       "O'Brien"],
    ],
  },
  {
    title: 'EXTENDED FAMILY & OTHER RELATIVES',
    sub: 'Collateral branches, in-laws, and associated individuals',
    hBg: C.neutral, hFg: C.text,
    people: [
      ['P-00032', 'Richard A. Hartwell',         'b. 1952 • Worcester, MA',           'd. 2020 • Worcester, MA',       'Hartwell'],
      ['P-00033', 'Dorothy Simmons Hartwell',    'b. 1954 • Worcester, MA',           'Living',                        'Hartwell'],
      ['P-00036', 'Sean Patrick O\'Brien',       'b. 1956 • Providence, RI',          'Living',                        "O'Brien"],
      ['P-00037', 'Colleen O\'Brien Donovan',    'b. 1960 • Providence, RI',          'd. 2015 • Providence, RI',      "O'Brien"],
      ['P-00040', 'Rose Kelly Chapman',          'b. c.1900 • Hartford, CT',          'd. 1978 • Springfield, MA',     'Hartwell'],
      ['P-00044', 'Elizabeth A. Hartwell',       'b. 1918 • Springfield, MA',         'd. 1919  (infant)',             'Hartwell'],
      ['P-00046', 'Carol Morrison',              'b. 1953 • Hartford, CT',            'Living  (James\'s 1st wife)',    'Hartwell'],
      ['P-00051', 'Grace Walsh Hartwell',        'b. 1955 • Worcester, MA',           'Living',                        'Hartwell'],
      ['P-00052', 'Brian Francis Walsh',         'b. 1950 • Worcester, MA',           'd. 2018 • Worcester, MA',       'Hartwell'],
      ['P-00053', 'Maureen Donovan',             'b. 1958 • Providence, RI',          'Living',                        "O'Brien"],
      ['P-00054', 'Patrick Sean Donovan',        'b. 1957 • Providence, RI',          'Living',                        "O'Brien"],
      ['P-00060', 'James A. Flynn',              'b. c.1884 • Co. Cork, Ireland',     'd. c.1950 • Providence, RI',    "O'Brien"],
      ['P-00068', 'Domenico Rossi',              'b. 1910 • Palermo, Sicily',         'd. 1978 • Providence, RI',      "O'Brien"],
      ['P-00069', 'Rosa Marino Rossi',           'b. 1915 • Palermo, Sicily',         'd. 1990 • Providence, RI',      "O'Brien"],
      ['P-00070', 'Bartholomew Chapman',         'b. c.1865 • Hartford, CT',          'd. c.1940 • Hartford, CT',      'Hartwell'],
      ['P-00071', 'Joseph B. O\'Brien',          'b. 1905 • Providence, RI',          'd. 1945 • Normandy (KIA)',      "O'Brien"],
      ['P-00072', 'Margaret Alice O\'Brien',     'b. 1910 • Providence, RI',          'Unknown — Brick Wall',          "O'Brien"],
      ['P-00073', 'William J. O\'Brien',         'b. 1908 • Providence, RI',          'd. 1980 • Providence, RI',      "O'Brien"],
      ['P-00074', 'Kathleen O\'Brien Carey',     'b. 1912 • Providence, RI',          'd. 2005 • Providence, RI',      "O'Brien"],
      ['P-00075', 'Robert F. Hartwell',          'b. 1924 • Springfield, MA',         'd. 1943 • Tarawa (KIA)',         'Hartwell'],
      ['P-00076', 'Mildred Hartwell Perry',      'b. 1928 • Springfield, MA',         'd. 2010 • Springfield, MA',     'Hartwell'],
      ['P-00077', 'Howard E. Perry',             'b. 1925 • Springfield, MA',         'd. 1998 • Springfield, MA',     'Hartwell'],
      ['P-00078', 'Barbara Perry Morrison',      'b. 1952 • Springfield, MA',         'Living',                        'Hartwell'],
      ['P-00079', 'Harold E. Perry',             'b. 1955 • Springfield, MA',         'Living',                        'Hartwell'],
      ['P-00080', 'Timothy J. Murphy',           'b. 1918 • Providence, RI',          'd. 1988 • Providence, RI',      "O'Brien"],
      ['P-00081', 'Frances Murphy Kelly',        'b. 1920 • Providence, RI',          'd. 2001 • Providence, RI',      "O'Brien"],
      ['P-00082', 'Gerard T. Kelly',             'b. 1918 • Providence, RI',          'd. 1975 • Providence, RI',      "O'Brien"],
      ['P-00083', 'Christopher West',            'b. 1960 • Boston, MA',              'Living  (Margaret\'s partner)', 'Hartwell'],
      ['P-00086', 'Salvatore Marino',            'b. c.1885 • Palermo, Sicily',       'd. 1960 • Providence, RI',      "O'Brien"],
      ['P-00087', 'Giuseppa Ferrara Marino',     'b. c.1890 • Palermo, Sicily',       'd. 1965 • Providence, RI',      "O'Brien"],
      ['P-00088', 'Andrew J. Gallagher',         'b. c.1879 • Co. Clare, Ireland',    'Status Unknown',                "O'Brien"],
      ['P-00089', 'Sarah [Unknown] Hartwell',    'b. Unknown',                        'Deceased — Brick Wall',         'Hartwell'],
      ['P-00096', 'Martin J. O\'Brien',          'b. c.1880 • Co. Cork, Ireland',     'd. 1960 • Providence, RI',      "O'Brien"],
      ['P-00097', 'Anastasia Shea O\'Brien',     'b. c.1883 • Providence, RI',        'd. 1968 • Providence, RI',      "O'Brien"],
      ['P-00100', 'James P. O\'Brien',           'b. 1940 • Providence, RI',          'Living',                        "O'Brien"],
      ['P-00101', 'Dorothy Walsh O\'Brien',      'b. 1943 • Providence, RI',          'Living',                        "O'Brien"],
      ['P-00103', 'Eleanor Parker',              'b. 1965 • Worcester, MA',           'Living',                        'Hartwell'],
    ],
  },
];

const BRANCH_BG = { 'Hartwell': C.info, "O'Brien": C.secondary, root: C.primaryDeep };
const BRANCH_FG = { 'Hartwell': C.text, "O'Brien": C.text, root: C.white };

async function main() {
  const reqs = [];
  const vals = [];

  // Sheet properties
  reqs.push({ updateSheetProperties: {
    properties: {
      sheetId: SID,
      tabColor: hex(C.secondaryDeep),
      tabColorStyle: { rgbColor: hex(C.secondaryDeep) },
      gridProperties: { frozenRowCount: 2 },
    },
    fields: 'tabColor,tabColorStyle,gridProperties.frozenRowCount',
  }});

  // Unmerge all from previous layout
  reqs.push({ unmergeCells: { range: gridRange(SID, 0, 500, 0, 20) } });

  // Column widths: 12 cols at 75px each
  for (let i = 0; i < TCOL; i++) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: 75 }, fields: 'pixelSize',
    }});
  }

  // Row 1: title
  reqs.push({ mergeCells: { range: gridRange(SID,0,1,0,TCOL), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(SID,0,1,0,TCOL),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryDeep),
      textFormat: { foregroundColor: hex(C.white), fontSize: 14, bold: true },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat',
  }});
  vals.push({ range: `'${S}'!A1`, values: [["HARTWELL – O'BRIEN COMPLETE FAMILY TREE  •  All 103 Members Across 10 Generations"]] });

  // Row 2: legend
  reqs.push({ mergeCells: { range: gridRange(SID,1,2,0,TCOL), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(SID,1,2,0,TCOL),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.bg),
      textFormat: { foregroundColor: hex(C.secText), fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat',
  }});
  vals.push({ range: `'${S}'!A2`, values: [['■ HARTWELL BRANCH (blue)     ■ O\'BRIEN BRANCH (green)     ★ = Root Person     Oldest ancestors at top → Youngest at bottom']] });

  let curRow = 2;

  for (const sec of SECTIONS) {
    const { title, sub, hBg, hFg, people } = sec;

    // Section header
    reqs.push({ mergeCells: { range: gridRange(SID, curRow, curRow+1, 0, TCOL), mergeType:'MERGE_ALL' }});
    reqs.push({ repeatCell: {
      range: gridRange(SID, curRow, curRow+1, 0, TCOL),
      cell: { userEnteredFormat: {
        backgroundColor: hex(hBg),
        textFormat: { foregroundColor: hex(hFg), fontSize: 10, bold: true },
        horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat',
    }});
    vals.push({ range: `'${S}'!A${curRow+1}`, values: [[`  ${title}  —  ${sub}  (${people.length} people)`]] });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'ROWS', startIndex: curRow, endIndex: curRow+1 },
      properties: { pixelSize: 24 }, fields: 'pixelSize',
    }});
    curRow++;

    // Person boxes
    let pos = 0;
    let boxStartRow = curRow;

    for (const [pid, name, born, died, branch] of people) {
      const col = pos * BCOL;
      const r1 = boxStartRow + Math.floor(pos / BPR) * BROW;
      const r2 = r1 + BROW;
      const c1 = (pos % BPR) * BCOL;
      const c2 = c1 + BCOL;

      reqs.push({ mergeCells: { range: gridRange(SID,r1,r2,c1,c2), mergeType:'MERGE_ALL' }});

      const bg = BRANCH_BG[branch] || C.neutral;
      const fg = BRANCH_FG[branch] || C.text;
      const isRoot = pid === 'P-00001';

      reqs.push({ repeatCell: {
        range: gridRange(SID,r1,r2,c1,c2),
        cell: { userEnteredFormat: {
          backgroundColor: hex(bg),
          textFormat: { foregroundColor: hex(fg), fontSize: 9, bold: isRoot },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
        }},
        fields: 'userEnteredFormat',
      }});

      const border = { style: 'SOLID_MEDIUM', color: hex(C.bg) };
      reqs.push({ updateBorders: {
        range: gridRange(SID,r1,r2,c1,c2),
        top: border, bottom: border, left: border, right: border,
      }});

      vals.push({ range: `'${S}'!${colL(c1)}${r1+1}`, values: [[`${name}\n${born}\n${died}  •  ${pid}`]] });
      pos++;
    }

    const rowsUsed = Math.ceil(people.length / BPR) * BROW;
    // Set row heights for box rows
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'ROWS', startIndex: boxStartRow, endIndex: boxStartRow + rowsUsed },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});

    curRow = boxStartRow + rowsUsed;
  }

  // Title row height
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 18 }, fields: 'pixelSize',
  }});

  await batchUpdate(id, reqs, 'tree-fmt');
  await valuesBatchUpdate(id, vals, 'tree-vals');

  const total = SECTIONS.reduce((s, sec) => s + sec.people.length, 0);
  console.log(`✓ Family Tree — complete register, ${total} people across ${SECTIONS.length} sections`);
}

main().catch(e => { console.error(e); process.exit(1); });
