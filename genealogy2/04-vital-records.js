'use strict';
const { sheets, batchUpdate, valuesBatchUpdate, gridRange, hex, colL, C } = require('./lib');
const { id, sheetMap } = JSON.parse(require('fs').readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Vital Records'];
const S = 'Vital Records';

const HEADERS = [
  'Event ID','Person ID','Person Name','Event Type','Event Date','Date Precision',
  'Event Place','Event Place Detail','Witnesses / Officiants','Source ID',
  'Evidence Confidence','Citation Status','Notes',
]; // A-M (13 cols)

// [personID, eventType, eventDate, datePrecision, place, placeDetail, witnesses, srcID, confidence, citStatus, notes]
const VITALS = [
  // === BIRTHS ===
  ['P-00001','Birth','1988-03-14','Exact','Boston, MA, USA','Massachusetts General Hospital','','SRC-00001','Confirmed','Complete Citation','Emma Rose Hartwell born.'],
  ['P-00002','Birth','1955-06-22','Exact','Worcester, MA, USA','St. Vincent Hospital','','SRC-00002','Confirmed','Complete Citation','James Edward Hartwell born.'],
  ['P-00003','Birth','1958-11-03','Exact','Providence, RI, USA','Women & Infants Hospital','','SRC-00003','Confirmed','Complete Citation','Margaret Anne O\'Brien born.'],
  ['P-00004','Birth','1922-04-10','Exact','Springfield, MA, USA','Home birth','','SRC-00004','Strong','Complete Citation','Robert Charles Hartwell born.'],
  ['P-00005','Birth','1926-09-28','Exact','Springfield, MA, USA','Home birth','','SRC-00005','Confirmed','Complete Citation','Eleanor Mae Chapman born.'],
  ['P-00006','Birth','1920-08-15','Exact','Providence, RI, USA','Home birth','','SRC-00006','Confirmed','Complete Citation','Patrick Joseph O\'Brien born.'],
  ['P-00007','Birth','1924-02-04','Exact','Providence, RI, USA','Home birth','','SRC-00007','Confirmed','Complete Citation','Catherine Mary Murphy born.'],
  ['P-00008','Birth','1888-07-03','Exact','Coventry, England','Parish of Holy Trinity','','SRC-00008','Confirmed','Complete Citation','William Henry Hartwell born. Baptism record.'],
  ['P-00009','Birth','1892-03-17','Exact','Coventry, England','Parish of Holy Trinity','','SRC-00009','Confirmed','Complete Citation','Agnes Louise Fletcher born.'],
  ['P-00010','Birth','1895-01-22','Estimated','Hartford, CT, USA','','','SRC-00010','Strong','Partial Citation','Charles Frederick Chapman born. Date estimated from census.'],
  ['P-00011','Birth','1898-06-11','Exact','Hartford, CT, USA','','','SRC-00011','Confirmed','Complete Citation','Harriet Jane Davies born.'],
  ['P-00012','Birth','1882-11-05','Year Only','County Cork, Ireland','Parish of Ballycotton','','SRC-00012','Confirmed','Complete Citation','Michael Patrick O\'Brien born. Civil registration.'],
  ['P-00013','Birth','1886-03-20','Year Only','County Cork, Ireland','Parish of Cloyne','','SRC-00013','Confirmed','Complete Citation','Brigid Nora Shaughnessy born.'],
  ['P-00014','Birth','1886-05-28','Year Only','County Clare, Ireland','Parish of Kilmihil','','SRC-00014','Confirmed','Complete Citation','Thomas Francis Murphy born.'],
  ['P-00015','Birth','1889-08-15','Year Only','County Clare, Ireland','Parish of Kilmihil','','SRC-00015','Confirmed','Complete Citation','Mary Josephine Gallagher born.'],
  ['P-00016','Birth','1850-04-12','About','Coventry, England','','','SRC-00016','Moderate','Partial Citation','George Thomas Hartwell born. Date approximate.'],
  ['P-00017','Birth','1854-09-20','About','Coventry, England','','','SRC-00017','Moderate','Partial Citation','Martha Ellen Holt born.'],
  ['P-00020','Birth','1845-10-08','Year Only','County Cork, Ireland','','','SRC-00020','Tentative','Partial Citation','Cornelius O\'Brien born. Pre-civil registration. Baptismal estimate.'],
  ['P-00021','Birth','1848-05-25','Year Only','County Cork, Ireland','','','SRC-00021','Tentative','Partial Citation','Honora Crowley born.'],
  ['P-00022','Birth','1850-03-15','Year Only','County Cork, Ireland','','','SRC-00022','Moderate','Partial Citation','Daniel Shaughnessy born.'],
  ['P-00023','Birth','1852-09-22','Year Only','County Cork, Ireland','','','SRC-00023','Moderate','Partial Citation','Ellen Sullivan born.'],
  ['P-00024','Birth','1852-07-04','Year Only','County Clare, Ireland','','','SRC-00023','Moderate','Partial Citation','Patrick Murphy born.'],
  ['P-00025','Birth','1855-02-19','Year Only','County Clare, Ireland','','','SRC-00023','Moderate','Partial Citation','Bridget Ryan born.'],
  ['P-00026','Birth','1983-09-11','Exact','Worcester, MA, USA','St. Vincent Hospital','','SRC-00024','Confirmed','Complete Citation','Thomas James Hartwell born.'],
  ['P-00027','Birth','1990-05-07','Exact','Worcester, MA, USA','UMass Memorial Medical Center','','SRC-00025','Confirmed','Complete Citation','Sarah Elizabeth Hartwell born.'],
  ['P-00028','Birth','1985-02-24','Exact','Chicago, IL, USA','Northwestern Memorial Hospital','','SRC-00026','Confirmed','Complete Citation','Linda Grace Patel born.'],
  ['P-00036','Birth','1956-03-08','Exact','Providence, RI, USA','Miriam Hospital','','SRC-00029','Confirmed','Complete Citation','Sean Patrick O\'Brien born.'],
  ['P-00037','Birth','1960-07-14','Exact','Providence, RI, USA','Women & Infants Hospital','','SRC-00030','Confirmed','Complete Citation','Colleen Frances O\'Brien born.'],
  ['P-00038','Birth','1993-08-20','Estimated','Unknown','Unknown','','','Unknown','Not Cited','Helen Rose Hartwell (adopted). Birth location unknown.'],
  ['P-00045','Birth','1975-11-20','Exact','Hartford, CT, USA','Hartford Hospital','','SRC-00002','Moderate','Partial Citation','Daniel Paul Hartwell born. Son of James and Carol.'],
  ['P-00047','Birth','1815-06-15','About','Warwickshire, England','','','SRC-00037','Moderate','Partial Citation','Frederick George Hartwell born. Date approximate.'],
  ['P-00049','Birth','1808-04-20','Year Only','County Cork, Ireland','','','SRC-00039','Tentative','Not Cited','Jeremiah O\'Brien born. Pre-civil registration. Famine era.'],
  ['P-00061','Birth','1775-01-01','Estimated','County Cork, Ireland','','','SRC-00045','Tentative','Not Cited','Thaddeus O\'Brien born. Estimated. Pre-civil registration.'],
  ['P-00063','Birth','1780-01-01','Estimated','Warwickshire, England','','','SRC-00046','Tentative','Not Cited','Thomas Hartwell born. Estimated.'],
  ['P-00068','Birth','1910-05-14','Exact','Palermo, Sicily, Italy','','','SRC-00049','Strong','Complete Citation','Domenico Rossi born.'],
  ['P-00069','Birth','1915-09-01','Exact','Palermo, Sicily, Italy','','','SRC-00050','Strong','Complete Citation','Rosa Carmela Marino born.'],
  ['P-00071','Birth','1905-04-25','Exact','Providence, RI, USA','Home birth','','SRC-00051','Confirmed','Complete Citation','Joseph Bernard O\'Brien born.'],
  ['P-00073','Birth','1908-07-19','Exact','Providence, RI, USA','Home birth','','SRC-00053','Confirmed','Complete Citation','William James O\'Brien born.'],
  ['P-00074','Birth','1912-03-17','Exact','Providence, RI, USA','Home birth','','SRC-00054','Confirmed','Complete Citation','Kathleen Agnes O\'Brien born.'],
  ['P-00075','Birth','1924-08-15','Exact','Springfield, MA, USA','Home birth','','SRC-00055','Confirmed','Complete Citation','Robert Francis Hartwell born.'],
  ['P-00076','Birth','1928-06-02','Exact','Springfield, MA, USA','Home birth','','SRC-00056','Confirmed','Complete Citation','Mildred Jean Hartwell born.'],
  ['P-00096','Birth','1880-03-28','Year Only','County Cork, Ireland','','','SRC-00065','Confirmed','Complete Citation','Martin Joseph O\'Brien born.'],
  ['P-00098','Birth','1912-08-10','Exact','Providence, RI, USA','Home birth','','SRC-00067','Confirmed','Complete Citation','Patrick Martin O\'Brien born.'],
  ['P-00100','Birth','1940-02-22','Exact','Providence, RI, USA','Rhode Island Hospital','','SRC-00069','Strong','Complete Citation','James Patrick O\'Brien born.'],
  ['P-00092','Birth','1982-07-30','Exact','Providence, RI, USA','Women & Infants Hospital','','','Strong','Not Cited','Declan Patrick O\'Brien born.'],
  // === BAPTISMS / CHRISTENINGS ===
  ['P-00008','Baptism / Christening','1888-07-10','Exact','Coventry, England','Holy Trinity Church','Rev. Thomas Ward','SRC-00008','Confirmed','Complete Citation','William Hartwell baptised. Church of England.'],
  ['P-00009','Baptism / Christening','1892-03-24','Exact','Coventry, England','Holy Trinity Church','Rev. Thomas Ward','SRC-00009','Confirmed','Complete Citation','Agnes Fletcher baptised.'],
  ['P-00012','Baptism / Christening','1882-11-12','Year Only','County Cork, Ireland','St. Colman\'s Church, Ballycotton','Fr. Brennan','SRC-00012','Confirmed','Complete Citation','Michael O\'Brien baptised. Catholic.'],
  ['P-00013','Baptism / Christening','1886-03-27','Year Only','County Cork, Ireland','Parish of Cloyne','Fr. McCarthy','SRC-00013','Confirmed','Complete Citation','Brigid Shaughnessy baptised.'],
  ['P-00014','Baptism / Christening','1886-06-04','Year Only','County Clare, Ireland','St. Michael\'s Church, Kilmihil','Fr. O\'Grady','SRC-00014','Confirmed','Complete Citation','Thomas Murphy baptised.'],
  ['P-00020','Baptism / Christening','1845-01-01','Year Only','County Cork, Ireland','Ballycotton Parish','','SRC-00020','Tentative','Partial Citation','Cornelius O\'Brien baptism estimated from parish book damage.'],
  ['P-00049','Baptism / Christening','1808-01-01','Year Only','County Cork, Ireland','','','SRC-00039','Tentative','Not Cited','Jeremiah O\'Brien baptism. Record inferred.'],
  // === MARRIAGES ===
  ['P-00002','Marriage','1981-06-15','Exact','Boston, MA, USA','Cathedral of the Holy Cross','Fr. O\'Sullivan','SRC-00002','Confirmed','Complete Citation','James Hartwell and Margaret O\'Brien married.'],
  ['P-00002','Marriage','1974-09-20','Exact','Hartford, CT, USA','First Congregational Church','Rev. Smith','','Moderate','Partial Citation','James Hartwell and Carol Morrison — first marriage.'],
  ['P-00004','Marriage','1950-08-12','Exact','Springfield, MA, USA','First Baptist Church','Rev. Brown','SRC-00004','Confirmed','Complete Citation','Robert and Eleanor Chapman Hartwell married.'],
  ['P-00006','Marriage','1948-06-10','Exact','Providence, RI, USA','St. Michael\'s Church','Fr. Doyle','SRC-00006','Confirmed','Complete Citation','Patrick O\'Brien and Catherine Murphy married.'],
  ['P-00008','Marriage','1916-07-04','Exact','Coventry, England','Holy Trinity Church','Rev. Hargreaves','SRC-00008','Strong','Complete Citation','William Hartwell and Agnes Fletcher married. England.'],
  ['P-00010','Marriage','1920-05-01','Exact','Hartford, CT, USA','Second Congregational Church','Rev. Porter','SRC-00010','Strong','Complete Citation','Charles Chapman and Harriet Davies married.'],
  ['P-00012','Marriage','1905-09-15','Exact','Providence, RI, USA','St. Michael\'s Church','Fr. Dolan','SRC-00012','Confirmed','Complete Citation','Michael O\'Brien and Brigid Shaughnessy married.'],
  ['P-00014','Marriage','1912-04-22','Exact','Providence, RI, USA','St. Patrick\'s Church','Fr. Walsh','SRC-00014','Confirmed','Complete Citation','Thomas Murphy and Mary Gallagher married.'],
  ['P-00016','Marriage','1875-01-01','Year Only','Coventry, England','St. Michael\'s Church','','SRC-00016','Moderate','Partial Citation','George Hartwell and Martha Holt married. England.'],
  ['P-00020','Marriage','1870-01-01','Year Only','County Cork, Ireland','Ballycotton Parish','','SRC-00020','Tentative','Partial Citation','Cornelius and Honora Crowley O\'Brien married. Ireland.'],
  ['P-00026','Marriage','2009-08-15','Exact','Chicago, IL, USA','Fourth Presbyterian Church','Rev. Carlson','SRC-00026','Confirmed','Complete Citation','Thomas Hartwell and Linda Patel married. Chicago.'],
  ['P-00027','Marriage','2013-09-28','Exact','Boston, MA, USA','Faneuil Hall, civil ceremony','Justice P. Chen','SRC-00025','Confirmed','Complete Citation','Sarah Hartwell and Peter Novak married.'],
  ['P-00032','Marriage','1977-11-20','Exact','Worcester, MA, USA','St. John\'s Church','Fr. Kelly','SRC-00028','Confirmed','Complete Citation','Richard Hartwell and Dorothy Simmons married.'],
  ['P-00037','Marriage','1982-03-17','Exact','Providence, RI, USA','St. Patrick\'s Church','Fr. Murphy','SRC-00030','Confirmed','Complete Citation','Colleen O\'Brien and Patrick Donovan married. St. Patrick\'s Day.'],
  ['P-00047','Marriage','1845-01-01','Year Only','Warwickshire, England','','','SRC-00037','Moderate','Partial Citation','Frederick Hartwell and Ann Turner married. England.'],
  ['P-00057','Marriage','2007-06-01','Exact','Newport, RI, USA','Ochre Court','','','Strong','Partial Citation','Sean Michael O\'Brien and Nicole Rossi married.'],
  ['P-00068','Marriage','1938-04-15','Exact','Providence, RI, USA','Holy Ghost Church','Fr. DiNapoli','SRC-00049','Strong','Complete Citation','Domenico Rossi and Rosa Marino married.'],
  ['P-00082','Marriage','1940-10-12','Exact','Providence, RI, USA','St. Patrick\'s Church','Fr. Collins','SRC-00060','Confirmed','Complete Citation','Gerard Kelly and Frances Murphy married.'],
  ['P-00092','Marriage','2008-07-19','Exact','Providence, RI, USA','Cathedral of SS. Peter & Paul','Fr. O\'Brien','','Confirmed','Not Cited','Declan O\'Brien and Aisling Murphy married.'],
  // === DIVORCES ===
  ['P-00002','Divorce','1979-05-01','Exact','Hartford, CT, USA','Hartford Superior Court','','SRC-00002','Moderate','Partial Citation','James Hartwell and Carol Morrison divorced.'],
  // === DEATHS ===
  ['P-00002','Death','2019-08-05','Exact','Boston, MA, USA','Massachusetts General Hospital','','SRC-00002','Confirmed','Complete Citation','James Edward Hartwell died. Age 64.'],
  ['P-00004','Death','1995-12-18','Exact','Worcester, MA, USA','UMass Memorial Medical Center','','SRC-00004','Confirmed','Complete Citation','Robert Charles Hartwell Sr. died. Age 73.'],
  ['P-00005','Death','2003-03-15','Exact','Worcester, MA, USA','St. Vincent Hospital','','SRC-00005','Confirmed','Complete Citation','Eleanor Mae Hartwell died. Age 76.'],
  ['P-00006','Death','1998-05-22','Exact','Providence, RI, USA','Rhode Island Hospital','','SRC-00006','Confirmed','Complete Citation','Patrick Joseph O\'Brien died. Age 77.'],
  ['P-00007','Death','2007-10-11','Exact','Providence, RI, USA','Miriam Hospital','','SRC-00007','Confirmed','Complete Citation','Catherine Mary O\'Brien died. Age 83.'],
  ['P-00008','Death','1952-11-30','Exact','Springfield, MA, USA','Home','','SRC-00008','Confirmed','Complete Citation','William Henry Hartwell died. Age 64.'],
  ['P-00009','Death','1960-04-09','Exact','Springfield, MA, USA','Home','','SRC-00009','Confirmed','Complete Citation','Agnes Louise Hartwell died. Age 68.'],
  ['P-00010','Death','1958-08-14','Exact','Springfield, MA, USA','Springfield Hospital','','SRC-00010','Strong','Complete Citation','Charles Frederick Chapman died. Age 63.'],
  ['P-00011','Death','1971-01-28','Exact','Springfield, MA, USA','Home','','SRC-00011','Confirmed','Complete Citation','Harriet Jane Chapman died. Age 72.'],
  ['P-00012','Death','1944-07-19','Exact','Providence, RI, USA','Home','','SRC-00012','Confirmed','Complete Citation','Michael Patrick O\'Brien died. Age 61.'],
  ['P-00013','Death','1949-09-12','Exact','Providence, RI, USA','Home','','SRC-00013','Confirmed','Complete Citation','Brigid Nora O\'Brien died. Age 63.'],
  ['P-00014','Death','1938-02-14','Exact','Providence, RI, USA','Rhode Island Hospital','','SRC-00014','Confirmed','Complete Citation','Thomas Francis Murphy died. Age 51.'],
  ['P-00015','Death','1955-11-22','Exact','Providence, RI, USA','Home','','SRC-00015','Confirmed','Complete Citation','Mary Josephine Murphy died. Age 66.'],
  ['P-00016','Death','1920-06-30','Year Only','Coventry, England','','','SRC-00016','Moderate','Partial Citation','George Thomas Hartwell died in England. Never emigrated.'],
  ['P-00017','Death','1925-03-15','Year Only','Coventry, England','','','SRC-00017','Moderate','Partial Citation','Martha Ellen Hartwell died. Coventry.'],
  ['P-00020','Death','1895-08-17','Year Only','County Cork, Ireland','','','SRC-00020','Tentative','Not Cited','Cornelius O\'Brien died. County Cork. Record uncertain.'],
  ['P-00021','Death','1900-12-03','Year Only','County Cork, Ireland','','','SRC-00021','Tentative','Partial Citation','Honora O\'Brien died. County Cork.'],
  ['P-00037','Death','2015-04-02','Exact','Providence, RI, USA','Rhode Island Hospital','','SRC-00030','Confirmed','Complete Citation','Colleen Frances Donovan died. Age 54.'],
  ['P-00044','Death','1919-01-15','Exact','Springfield, MA, USA','Home','','SRC-00036','Confirmed','Complete Citation','Elizabeth Anne Hartwell died age 6 weeks. Infant mortality.'],
  ['P-00047','Death','1880-03-22','About','Coventry, England','','','SRC-00037','Moderate','Partial Citation','Frederick George Hartwell died. Coventry.'],
  ['P-00049','Death','1855-09-18','Year Only','County Cork, Ireland','','','SRC-00039','Tentative','Not Cited','Jeremiah O\'Brien died. Famine era.'],
  ['P-00071','Death','1945-06-06','Exact','Normandy, France','Omaha Beach area','','SRC-00051','Confirmed','Complete Citation','Joseph Bernard O\'Brien KIA D-Day. US Army.'],
  ['P-00073','Death','1980-04-10','Exact','Providence, RI, USA','Miriam Hospital','','SRC-00053','Confirmed','Complete Citation','William James O\'Brien died. Age 71.'],
  ['P-00074','Death','2005-11-22','Exact','Providence, RI, USA','Rhode Island Hospital','','SRC-00054','Confirmed','Complete Citation','Kathleen Agnes O\'Brien Carey died. Age 93.'],
  ['P-00075','Death','1943-11-14','Exact','Tarawa, Pacific','Battle of Tarawa','','SRC-00055','Confirmed','Complete Citation','Robert Francis Hartwell KIA WWII. US Marine Corps.'],
  ['P-00076','Death','2010-08-19','Exact','Springfield, MA, USA','Baystate Medical Center','','SRC-00056','Confirmed','Complete Citation','Mildred Jean Perry died. Age 82.'],
  ['P-00077','Death','1998-02-28','Exact','Springfield, MA, USA','Baystate Medical Center','','SRC-00057','Confirmed','Complete Citation','Howard Earl Perry died. Age 72.'],
  ['P-00082','Death','1975-04-18','Exact','Providence, RI, USA','Rhode Island Hospital','','SRC-00060','Confirmed','Complete Citation','Gerard Thomas Kelly died. Age 56.'],
  ['P-00096','Death','1960-11-04','Exact','Providence, RI, USA','Rhode Island Hospital','','SRC-00065','Confirmed','Complete Citation','Martin Joseph O\'Brien died. Age 80.'],
  ['P-00098','Death','1985-01-14','Exact','Providence, RI, USA','Rhode Island Hospital','','SRC-00067','Confirmed','Complete Citation','Patrick Martin O\'Brien died. Age 72.'],
  // === BURIALS ===
  ['P-00002','Burial','2019-08-10','Exact','Worcester, MA, USA','St. John\'s Cemetery','','SRC-00002','Confirmed','Complete Citation','James Edward Hartwell buried at St. John\'s.'],
  ['P-00004','Burial','1996-01-05','Exact','Worcester, MA, USA','Pine Grove Cemetery','','SRC-00004','Confirmed','Complete Citation','Robert Charles Hartwell Sr. buried.'],
  ['P-00005','Burial','2003-03-20','Exact','Worcester, MA, USA','Pine Grove Cemetery','','SRC-00005','Confirmed','Complete Citation','Eleanor Mae Hartwell buried.'],
  ['P-00008','Burial','1952-12-03','Exact','Springfield, MA, USA','Greenwood Cemetery','','SRC-00008','Confirmed','Complete Citation','William Hartwell buried. Greenwood.'],
  ['P-00009','Burial','1960-04-12','Exact','Springfield, MA, USA','Greenwood Cemetery','','SRC-00009','Confirmed','Complete Citation','Agnes Hartwell buried.'],
  ['P-00012','Burial','1944-07-22','Exact','Providence, RI, USA','Sacred Heart Cemetery','','SRC-00012','Confirmed','Complete Citation','Michael O\'Brien buried.'],
  ['P-00071','Burial','1945-01-01','Year Only','Normandy, France','Normandy American Cemetery, Colleville-sur-Mer','','SRC-00051','Confirmed','Complete Citation','Joseph Bernard O\'Brien buried at Normandy American Cemetery.'],
  ['P-00075','Burial','1944-01-01','Year Only','Hawaii, USA','National Memorial Cemetery of the Pacific','','SRC-00055','Confirmed','Complete Citation','Robert Francis Hartwell buried. Punchbowl Cemetery.'],
];

(async () => {
  const reqs = [];
  const vals = [];

  // Tab setup
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, tabColor: hex(C.lavender), gridProperties: { frozenRowCount: 7, frozenColumnCount: 3 } },
    fields: 'tabColor,gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
  }});

  // Row 1 — Title
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 13), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 0, 1, 0, 13),
    cell: { userEnteredFormat: {
      backgroundColor: hex('#7B6D9E'),
      textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  vals.push({ range: `'${S}'!A1`, values: [['VITAL RECORDS — BIRTHS, MARRIAGES, DEATHS & MORE']] });

  // Row 2 — Subtitle
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 13), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 1, 2, 0, 13),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lavender),
      textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  vals.push({ range: `'${S}'!A2`, values: [['Births, baptisms, marriages, divorces, deaths, burials, and other life events with source citations.']] });

  // Row 3-4 — Stats
  const statLabels = ['Total Events','Births','Marriages','Deaths','Burials','Cited Events'];
  const statCols = [0,2,4,6,8,10];
  const deepColor = '#7B6D9E';
  statLabels.forEach((lbl, i) => {
    const c = statCols[i];
    reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, c, c+2), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, c, c+2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, 2, 3, c, c+2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(deepColor),
        textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(SID, 3, 4, c, c+2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.input),
        textFormat: { bold: true, fontSize: 11, foregroundColor: hex(deepColor), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});
    vals.push({ range: `'${S}'!${colL(c)}3`, values: [[lbl]] });
  });
  const statFormulas = [
    `=COUNTA($B$8:$B$6007)`,
    `=COUNTIF($D$8:$D$6007,"Birth")`,
    `=COUNTIF($D$8:$D$6007,"Marriage")`,
    `=COUNTIF($D$8:$D$6007,"Death")`,
    `=COUNTIF($D$8:$D$6007,"Burial")`,
    `=COUNTIF($L$8:$L$6007,"Complete Citation")`,
  ];
  statFormulas.forEach((f, i) => {
    vals.push({ range: `'${S}'!${colL(statCols[i])}4`, values: [[f]] });
  });

  // Row 5-6 — spacer
  reqs.push({ repeatCell: {
    range: gridRange(SID, 4, 6, 0, 13),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat(backgroundColor)',
  }});

  // Row 7 — Headers
  reqs.push({ repeatCell: {
    range: gridRange(SID, 6, 7, 0, 13),
    cell: { userEnteredFormat: {
      backgroundColor: hex(deepColor),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  vals.push({ range: `'${S}'!A7`, values: [HEADERS] });

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  // Column widths: A=90, B=90, C=180, D=140, E=100, F=100, G=200, H=200, I=180, J=100, K=120, L=110, M=260
  const colWidths = [90,90,180,140,100,100,200,200,180,100,120,110,260];
  colWidths.forEach((px, ci) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  // Border on header row
  reqs.push({ updateBorders: {
    range: gridRange(SID, 6, 7, 0, 13),
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.border) },
  }});

  // Conditional formatting — Event Type (col D = index 3) color-coded
  const CF_EVENT = [
    { val: 'Birth',             bg: '#C8E6C9' },
    { val: 'Baptism / Christening', bg: '#DCEDC8' },
    { val: 'Marriage',          bg: '#FCE4EC' },
    { val: 'Civil Union',       bg: '#FCE4EC' },
    { val: 'Divorce',           bg: '#FFCCBC' },
    { val: 'Death',             bg: '#CFD8DC' },
    { val: 'Burial',            bg: '#ECEFF1' },
    { val: 'Cremation',         bg: '#ECEFF1' },
  ];
  CF_EVENT.forEach(({ val, bg }) => {
    reqs.push({ addConditionalFormatRule: { rule: {
      ranges: [gridRange(SID, 7, 6007, 3, 4)],
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: val }] },
        format: { backgroundColor: hex(bg) },
      },
    }, index: 0 }});
  });

  // Conditional formatting — Evidence Confidence (col K = index 10)
  const CF_CONF = [
    { val: 'Confirmed',  bg: C.confirmed },
    { val: 'Strong',     bg: C.secondary },
    { val: 'Moderate',   bg: C.info },
    { val: 'Tentative',  bg: C.review },
    { val: 'Conflicting',bg: C.conflict },
    { val: 'Unknown',    bg: C.neutral },
  ];
  CF_CONF.forEach(({ val, bg }) => {
    reqs.push({ addConditionalFormatRule: { rule: {
      ranges: [gridRange(SID, 7, 6007, 10, 11)],
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: val }] },
        format: { backgroundColor: hex(bg) },
      },
    }, index: 0 }});
  });

  // Data validation
  // D = Vital Event Types (col 3) — Reference Data!$E$2:$E$12
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 6007, 3, 4),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$E$2:$E$12` }] },
      strict: false, showCustomUi: true,
    },
  }});
  // F = Date Precision (col 5) — Reference Data!$F$2:$F$9
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 6007, 5, 6),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$F$2:$F$9` }] },
      strict: false, showCustomUi: true,
    },
  }});
  // K = Evidence Confidence (col 10) — Reference Data!$J$2:$J$7
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 6007, 10, 11),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$J$2:$J$7` }] },
      strict: false, showCustomUi: true,
    },
  }});
  // L = Citation Status (col 11) — Reference Data!$L$2:$L$5
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 6007, 11, 12),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$L$2:$L$5` }] },
      strict: false, showCustomUi: true,
    },
  }});

  // Alternating row fill
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, 7, 6007, 0, 13)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND(MOD(ROW(),2)=0,LEN($B8)>0)' }] },
      format: { backgroundColor: hex(C.altRow) },
    },
  }, index: 0 }});

  // Build data rows
  const dataRows = VITALS.map(([pid, evType, evDate, datPrec, place, placeDetail, witnesses, srcID, conf, citStatus, notes]) => {
    return [
      `=IF(B{R}="","","EVT-"&TEXT(ROW()-7,"00000"))`,
      pid,
      `=IFERROR(VLOOKUP(B{R},'Master People'!$A$8:$B$5007,2,FALSE),"")`,
      evType,
      evDate,
      datPrec,
      place,
      placeDetail,
      witnesses,
      srcID,
      conf,
      citStatus,
      notes,
    ];
  });

  dataRows.forEach((row, i) => {
    const r = i + 8;
    for (let c = 0; c < row.length; c++) {
      if (typeof row[c] === 'string') {
        row[c] = row[c].replace(/\{R\}/g, String(r));
      }
    }
  });

  await batchUpdate(id, reqs, 'vit-fmt');
  await valuesBatchUpdate(id, vals, 'vit-vals');
  await valuesBatchUpdate(id, [{ range: `'${S}'!A8`, values: dataRows }], 'vit-data');
  console.log(`✓ Vital Records — ${VITALS.length} rows written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
