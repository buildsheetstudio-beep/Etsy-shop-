'use strict';
const { sheets, batchUpdate, valuesBatchUpdate, gridRange, hex, colL, C } = require('./lib');
const { id, sheetMap } = JSON.parse(require('fs').readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Locations & Migration'];
const S = 'Locations & Migration';

const HEADERS = [
  'Record ID','Person ID','Person Name','Location Event Type','Date','Date Precision',
  'Country','State / Province / County','City / Town / Parish','Specific Address / Detail',
  'Latitude / Coordinates','Source ID','Evidence Confidence','Notes',
]; // A-N (14 cols)

// [pid, locType, date, datePrec, country, stateRegion, city, detail, coords, srcID, conf, notes]
const LOCS = [
  // === EMMA (P-00001) ===
  ['P-00001','Birthplace','1988-03-14','Exact','USA','Massachusetts','Boston','Massachusetts General Hospital','42.3601,-71.0589','SRC-00001','Confirmed','Emma Hartwell born Boston.'],
  ['P-00001','Childhood Residence','1988-01-01','Year Only','USA','Massachusetts','Worcester','Grew up in Worcester family home','','SRC-00001','Confirmed','Emma grew up in Worcester.'],
  ['P-00001','Residence','2010-01-01','Year Only','USA','Massachusetts','Cambridge','Graduate student housing, MIT area','','','Strong','Emma at MIT/Cambridge for graduate studies.'],
  ['P-00001','Residence','2015-01-01','Year Only','USA','Massachusetts','Boston','Current residence, Boston','','','Strong','Emma currently lives in Boston.'],
  // === JAMES (P-00002) ===
  ['P-00002','Birthplace','1955-06-22','Exact','USA','Massachusetts','Worcester','St. Vincent Hospital','42.2626,-71.8023','SRC-00002','Confirmed','James Hartwell born Worcester.'],
  ['P-00002','Childhood Residence','1955-01-01','Year Only','USA','Massachusetts','Worcester','Family home on Elm Street','','SRC-00002','Confirmed','James grew up in Worcester.'],
  ['P-00002','Residence','1974-01-01','Year Only','USA','Connecticut','Hartford','Apartment — first marriage period','','','Moderate','James lived in Hartford during first marriage.'],
  ['P-00002','Residence','1981-01-01','Year Only','USA','Massachusetts','Boston','Family home, Boston area','','SRC-00002','Confirmed','James and Margaret\'s family home.'],
  ['P-00002','Death Location','2019-08-05','Exact','USA','Massachusetts','Boston','Massachusetts General Hospital','','SRC-00002','Confirmed','James died Boston.'],
  // === MARGARET (P-00003) ===
  ['P-00003','Birthplace','1958-11-03','Exact','USA','Rhode Island','Providence','Women & Infants Hospital','41.8240,-71.4128','SRC-00003','Confirmed','Margaret O\'Brien born Providence.'],
  ['P-00003','Childhood Residence','1958-01-01','Year Only','USA','Rhode Island','Providence','O\'Brien family home, Elmwood Ave','','SRC-00003','Confirmed','Margaret grew up in Providence.'],
  ['P-00003','Residence','1981-01-01','Year Only','USA','Massachusetts','Boston','Hartwell family home, Boston area','','SRC-00003','Confirmed','Margaret moved to Boston after marriage.'],
  // === ROBERT SR. (P-00004) ===
  ['P-00004','Birthplace','1922-04-10','Exact','USA','Massachusetts','Springfield','Home birth','42.1015,-72.5898','SRC-00004','Strong','Robert Hartwell Sr. born Springfield.'],
  ['P-00004','Childhood Residence','1922-01-01','Year Only','USA','Massachusetts','Springfield','Hartwell family home, Springfield','','SRC-00004','Strong','Robert grew up in Springfield.'],
  ['P-00004','Residence','1950-01-01','Year Only','USA','Massachusetts','Worcester','Moved to Worcester after marriage','','SRC-00004','Confirmed','Robert and Eleanor settled in Worcester.'],
  ['P-00004','Death Location','1995-12-18','Exact','USA','Massachusetts','Worcester','UMass Memorial Medical Center','','SRC-00004','Confirmed','Robert Sr. died Worcester.'],
  ['P-00004','Burial Location','1996-01-05','Exact','USA','Massachusetts','Worcester','Pine Grove Cemetery','','SRC-00004','Confirmed','Robert Sr. buried Pine Grove.'],
  // === ELEANOR (P-00005) ===
  ['P-00005','Birthplace','1926-09-28','Exact','USA','Massachusetts','Springfield','Home birth','','SRC-00005','Confirmed','Eleanor Chapman born Springfield.'],
  ['P-00005','Childhood Residence','1926-01-01','Year Only','USA','Massachusetts','Springfield','Chapman family home, Springfield','','SRC-00005','Confirmed','Eleanor grew up in Springfield.'],
  ['P-00005','Residence','1950-01-01','Year Only','USA','Massachusetts','Worcester','Moved to Worcester after marriage','','SRC-00005','Confirmed','Eleanor and Robert settled in Worcester.'],
  ['P-00005','Death Location','2003-03-15','Exact','USA','Massachusetts','Worcester','St. Vincent Hospital','','SRC-00005','Confirmed','Eleanor died Worcester.'],
  // === PATRICK O\'BRIEN (P-00006) ===
  ['P-00006','Birthplace','1920-08-15','Exact','USA','Rhode Island','Providence','Home birth, Elmwood Ave area','','SRC-00006','Confirmed','Patrick O\'Brien born Providence.'],
  ['P-00006','Residence','1920-01-01','Year Only','USA','Rhode Island','Providence','O\'Brien family neighborhood, Elmwood','','SRC-00006','Confirmed','Patrick lived all his life in Providence.'],
  ['P-00006','Death Location','1998-05-22','Exact','USA','Rhode Island','Providence','Rhode Island Hospital','','SRC-00006','Confirmed','Patrick died Providence.'],
  ['P-00006','Burial Location','1998-05-27','Exact','USA','Rhode Island','Providence','St. Michael\'s Cemetery','','SRC-00006','Confirmed','Patrick buried St. Michael\'s.'],
  // === MICHAEL O\'BRIEN (P-00012) ===
  ['P-00012','Birthplace','1882-11-05','Year Only','Ireland','County Cork','Ballycotton','Parish of Ballycotton','51.8283,-8.0086','SRC-00012','Confirmed','Michael O\'Brien born County Cork, Ireland.'],
  ['P-00012','Childhood Residence','1882-01-01','Year Only','Ireland','County Cork','Ballycotton','O\'Brien family farm','','SRC-00012','Confirmed','Michael grew up in Ballycotton.'],
  ['P-00012','Immigration Origin','1903-01-01','Year Only','Ireland','County Cork','Cobh (Queenstown)','Departure port for America','','SRC-00012','Confirmed','Michael emigrated from Cobh/Queenstown 1903.'],
  ['P-00012','Immigration Destination','1903-01-01','Year Only','USA','Rhode Island','Providence','Arrival Ellis Island then to Providence','40.6892,-74.0445','SRC-00012','Confirmed','Michael arrived USA c.1903. Settled Providence.'],
  ['P-00012','Residence','1903-01-01','Year Only','USA','Rhode Island','Providence','Elmwood Ave neighborhood, Providence','','SRC-00012','Confirmed','Michael settled in the Irish neighborhood of Providence.'],
  ['P-00012','Death Location','1944-07-19','Exact','USA','Rhode Island','Providence','Home, Elmwood Ave','','SRC-00012','Confirmed','Michael O\'Brien died at home, Providence.'],
  // === BRIGID O\'BRIEN (P-00013) ===
  ['P-00013','Birthplace','1886-03-20','Year Only','Ireland','County Cork','Cloyne area','Parish of Cloyne','51.8594,-8.1177','SRC-00013','Confirmed','Brigid Shaughnessy born County Cork.'],
  ['P-00013','Immigration Origin','1904-01-01','Year Only','Ireland','County Cork','Cobh (Queenstown)','Emigrated shortly after Michael','','SRC-00013','Confirmed','Brigid emigrated from Cork 1904.'],
  ['P-00013','Immigration Destination','1904-01-01','Year Only','USA','Rhode Island','Providence','Joined Irish community Providence','','SRC-00013','Confirmed','Brigid arrived Providence 1904.'],
  // === WILLIAM HARTWELL (P-00008) ===
  ['P-00008','Birthplace','1888-07-03','Exact','England','West Midlands','Coventry','Parish of Holy Trinity','52.4081,-1.5106','SRC-00008','Confirmed','William Hartwell born Coventry, England.'],
  ['P-00008','Childhood Residence','1888-01-01','Year Only','England','West Midlands','Coventry','Hartwell family home, Coventry','','SRC-00008','Confirmed','William grew up in Coventry.'],
  ['P-00008','Immigration Origin','1910-01-01','Year Only','England','West Midlands','Coventry','Departed England for America c.1910','','SRC-00008','Strong','William emigrated from Coventry c.1910.'],
  ['P-00008','Immigration Destination','1910-01-01','Year Only','USA','Massachusetts','Springfield','Settled in Springfield MA','','SRC-00008','Strong','William arrived Springfield c.1910.'],
  ['P-00008','Residence','1910-01-01','Year Only','USA','Massachusetts','Springfield','Springfield, MA — mill worker neighborhood','','SRC-00008','Confirmed','William lived in Springfield until death.'],
  ['P-00008','Death Location','1952-11-30','Exact','USA','Massachusetts','Springfield','Home, Springfield','','SRC-00008','Confirmed','William died at home, Springfield.'],
  ['P-00008','Burial Location','1952-12-03','Exact','USA','Massachusetts','Springfield','Greenwood Cemetery','','SRC-00008','Confirmed','William buried Greenwood, Springfield.'],
  // === THOMAS MURPHY (P-00014) ===
  ['P-00014','Birthplace','1886-05-28','Year Only','Ireland','County Clare','Kilmihil','Parish of Kilmihil','52.7564,-9.2272','SRC-00014','Confirmed','Thomas Murphy born County Clare, Ireland.'],
  ['P-00014','Immigration Origin','1907-01-01','Year Only','Ireland','County Clare','Cobh (Queenstown)','Emigrated 1907','','SRC-00014','Confirmed','Thomas emigrated 1907.'],
  ['P-00014','Immigration Destination','1907-01-01','Year Only','USA','Rhode Island','Providence','Ellis Island, then Providence','','SRC-00014','Confirmed','Thomas arrived Providence 1907.'],
  ['P-00014','Residence','1907-01-01','Year Only','USA','Rhode Island','Providence','South Providence neighborhood','','SRC-00014','Confirmed','Thomas Murphy settled Providence.'],
  ['P-00014','Death Location','1938-02-14','Exact','USA','Rhode Island','Providence','Rhode Island Hospital','','SRC-00014','Confirmed','Thomas Murphy died Providence.'],
  // === CORNELIUS O\'BRIEN (P-00020) ===
  ['P-00020','Birthplace','1845-10-08','Year Only','Ireland','County Cork','Ballycotton','O\'Brien family farm','','SRC-00020','Tentative','Cornelius O\'Brien born Ballycotton area.'],
  ['P-00020','Childhood Residence','1845-01-01','Year Only','Ireland','County Cork','Ballycotton','Cork family farmland','','SRC-00020','Tentative','Cornelius remained in Cork all his life.'],
  ['P-00020','Death Location','1895-08-17','Year Only','Ireland','County Cork','Ballycotton','Died in Ireland','','SRC-00020','Tentative','Cornelius died County Cork. Never emigrated.'],
  // === GEORGE HARTWELL (P-00016) ===
  ['P-00016','Birthplace','1850-04-12','About','England','West Midlands','Coventry','','','SRC-00016','Moderate','George Hartwell born Coventry. Approximate date.'],
  ['P-00016','Childhood Residence','1850-01-01','Year Only','England','West Midlands','Coventry','Coventry weaving community','','SRC-00016','Moderate','George grew up in Coventry weaving trade.'],
  ['P-00016','Residence','1850-01-01','Year Only','England','West Midlands','Coventry','Coventry — never emigrated','','SRC-00016','Moderate','George stayed in England.'],
  ['P-00016','Death Location','1920-06-30','Year Only','England','West Midlands','Coventry','','','SRC-00016','Moderate','George Hartwell died Coventry.'],
  // === MARTIN O\'BRIEN (P-00096) ===
  ['P-00096','Birthplace','1880-03-28','Year Only','Ireland','County Cork','County Cork','','','SRC-00065','Confirmed','Martin O\'Brien born County Cork.'],
  ['P-00096','Immigration Origin','1903-01-01','Year Only','Ireland','County Cork','Cobh (Queenstown)','Emigrated with brother Michael','','SRC-00065','Confirmed','Martin emigrated with Michael 1903.'],
  ['P-00096','Immigration Destination','1903-01-01','Year Only','USA','Rhode Island','Providence','Settled Providence with brother','','SRC-00065','Confirmed','Martin settled Providence 1903.'],
  ['P-00096','Residence','1903-01-01','Year Only','USA','Rhode Island','Providence','South Providence, near Michael','','SRC-00065','Confirmed','Martin O\'Brien lived in Providence.'],
  // === JOSEPH BERNARD O\'BRIEN (P-00071) ===
  ['P-00071','Birthplace','1905-04-25','Exact','USA','Rhode Island','Providence','Home birth, Providence','','SRC-00051','Confirmed','Joseph Bernard O\'Brien born Providence.'],
  ['P-00071','Military Location','1944-01-01','Year Only','USA','Virginia','Camp A.P. Hill','Army training camp','','SRC-00051','Confirmed','Joseph trained at Camp A.P. Hill before deployment.'],
  ['P-00071','Military Location','1945-06-06','Exact','France','Normandy','Omaha Beach','Normandy landings, D-Day','49.3715,0.8503','SRC-00051','Confirmed','Joseph Bernard O\'Brien KIA at Normandy, D-Day.'],
  ['P-00071','Burial Location','1945-01-01','Year Only','France','Normandy','Colleville-sur-Mer','Normandy American Cemetery','','SRC-00051','Confirmed','Buried Normandy American Cemetery.'],
  // === ROBERT FRANCIS HARTWELL (P-00075) ===
  ['P-00075','Birthplace','1924-08-15','Exact','USA','Massachusetts','Springfield','Home birth, Springfield','','SRC-00055','Confirmed','Robert Francis Hartwell born Springfield.'],
  ['P-00075','Military Location','1943-01-01','Year Only','USA','California','San Diego','USMC training, Marine Corps Recruit Depot','','SRC-00055','Confirmed','Robert Francis trained at MCRD San Diego.'],
  ['P-00075','Military Location','1943-11-14','Exact','Pacific','','Tarawa Atoll','Battle of Tarawa, Kiribati','1.3667,172.9833','SRC-00055','Confirmed','Robert Francis Hartwell KIA at Battle of Tarawa.'],
  ['P-00075','Burial Location','1944-01-01','Year Only','USA','Hawaii','Honolulu','National Memorial Cemetery of the Pacific, Punchbowl','','SRC-00055','Confirmed','Robert Francis buried Punchbowl, Hawaii.'],
  // === DOMENICO ROSSI (P-00068) ===
  ['P-00068','Birthplace','1910-05-14','Exact','Italy','Sicily','Palermo','','38.1157,13.3615','SRC-00049','Strong','Domenico Rossi born Palermo, Sicily.'],
  ['P-00068','Immigration Origin','1935-01-01','Year Only','Italy','Sicily','Palermo','Departed Palermo for America','','SRC-00049','Strong','Domenico emigrated 1935.'],
  ['P-00068','Immigration Destination','1935-01-01','Year Only','USA','Rhode Island','Providence','Italian community, Federal Hill, Providence','','SRC-00049','Strong','Domenico settled Federal Hill, Providence.'],
  ['P-00068','Residence','1935-01-01','Year Only','USA','Rhode Island','Providence','Federal Hill neighborhood (Little Italy)','','SRC-00049','Strong','Domenico lived Federal Hill, Providence.'],
  ['P-00068','Death Location','1978-03-25','Exact','USA','Rhode Island','Providence','Rhode Island Hospital','','SRC-00049','Strong','Domenico Rossi died Providence.'],
  // === SALVATORE MARINO (P-00086) ===
  ['P-00086','Birthplace','1885-11-20','Year Only','Italy','Sicily','Palermo','','','SRC-00062','Moderate','Salvatore Marino born Palermo, Sicily.'],
  ['P-00086','Immigration Origin','1912-01-01','Year Only','Italy','Sicily','Palermo','Emigrated 1912','','SRC-00062','Moderate','Salvatore emigrated 1912.'],
  ['P-00086','Immigration Destination','1912-01-01','Year Only','USA','Rhode Island','Providence','Italian community, Providence','','SRC-00062','Moderate','Salvatore Marino arrived Providence 1912.'],
  // === THOMAS JAMES HARTWELL (P-00026) ===
  ['P-00026','Birthplace','1983-09-11','Exact','USA','Massachusetts','Worcester','St. Vincent Hospital','','SRC-00024','Confirmed','Thomas J. Hartwell born Worcester.'],
  ['P-00026','Childhood Residence','1983-01-01','Year Only','USA','Massachusetts','Worcester','Grew up in Worcester with family','','SRC-00024','Confirmed','Thomas grew up in Worcester.'],
  ['P-00026','Residence','2006-01-01','Year Only','USA','Illinois','Chicago','Moved to Chicago for work','','','Strong','Thomas relocated to Chicago.'],
  ['P-00026','Residence','2009-01-01','Year Only','USA','Illinois','Chicago','Family home, Chicago','','','Strong','Thomas and Linda living in Chicago.'],
  // === SARAH HARTWELL (P-00027) ===
  ['P-00027','Birthplace','1990-05-07','Exact','USA','Massachusetts','Worcester','UMass Memorial Medical Center','','SRC-00025','Confirmed','Sarah Hartwell born Worcester.'],
  ['P-00027','Residence','2012-01-01','Year Only','USA','Massachusetts','Boston','Moved to Boston, nursing career','','','Strong','Sarah lives in Boston.'],
  // === HELEN HARTWELL (P-00038) ===
  ['P-00038','Birthplace','1993-08-20','Estimated','Unknown','Unknown','Unknown','Birth location unknown. Adopted.','','','Unknown','Helen\'s birth location is unknown. Brick wall.'],
  ['P-00038','Childhood Residence','1994-01-01','Year Only','USA','Massachusetts','Worcester','Hartwell family home, Worcester','','','Unknown','Helen raised in Worcester by Hartwell family from age 1.'],
  // === PATRICK MURPHY (P-00024) ===
  ['P-00024','Birthplace','1852-07-04','Year Only','Ireland','County Clare','County Clare','Rural farming area','','SRC-00023','Moderate','Patrick Murphy born County Clare.'],
  ['P-00024','Childhood Residence','1852-01-01','Year Only','Ireland','County Clare','County Clare','Murphy family farm','','SRC-00023','Moderate','Patrick grew up in Clare farming community.'],
  ['P-00024','Residence','1878-01-01','Year Only','Ireland','County Clare','County Clare','Remained in Clare before children emigrated','','SRC-00023','Moderate','Patrick Murphy stayed in Ireland.'],
  // === JEREMIAH O\'BRIEN (P-00049) ===
  ['P-00049','Birthplace','1808-04-20','Year Only','Ireland','County Cork','County Cork','Pre-civil registration','','SRC-00039','Tentative','Jeremiah O\'Brien born County Cork.'],
  ['P-00049','Childhood Residence','1808-01-01','Year Only','Ireland','County Cork','County Cork','Famine-era Cork countryside','','SRC-00039','Tentative','Jeremiah lived in Cork. Famine era.'],
  ['P-00049','Death Location','1855-09-18','Year Only','Ireland','County Cork','County Cork','Possibly famine-related death','','SRC-00039','Tentative','Jeremiah O\'Brien died County Cork during or after famine.'],
  // === DANIEL HARTWELL (P-00045) — HALF BROTHER ===
  ['P-00045','Birthplace','1975-11-20','Exact','USA','Connecticut','Hartford','Hartford Hospital','','SRC-00002','Moderate','Daniel Hartwell born Hartford, CT.'],
  ['P-00045','Childhood Residence','1975-01-01','Year Only','USA','Connecticut','Hartford','Carol Morrison\'s home, Hartford','','','Moderate','Daniel raised by Carol Morrison in Hartford.'],
  ['P-00045','Residence','2000-01-01','Year Only','USA','Connecticut','Hartford','Lives in Hartford area','','','Moderate','Daniel Hartwell currently Hartford area.'],
  // === COLLEEN DONOVAN (P-00037) ===
  ['P-00037','Birthplace','1960-07-14','Exact','USA','Rhode Island','Providence','Women & Infants Hospital','','SRC-00030','Confirmed','Colleen O\'Brien born Providence.'],
  ['P-00037','Residence','1982-01-01','Year Only','USA','Rhode Island','Providence','Donovan family home, Providence','','SRC-00030','Confirmed','Colleen lived Providence with Patrick Donovan.'],
  ['P-00037','Death Location','2015-04-02','Exact','USA','Rhode Island','Providence','Rhode Island Hospital','','SRC-00030','Confirmed','Colleen Donovan died Providence.'],
  // === FIONA DONOVAN (P-00055) ===
  ['P-00055','Birthplace','1983-04-17','Exact','USA','Rhode Island','Providence','Women & Infants Hospital','','','Strong','Fiona Donovan born Providence.'],
  ['P-00055','Residence','2005-01-01','Year Only','USA','Rhode Island','Providence','Lives in Providence area','','','Strong','Fiona lives in Providence.'],
  // === SEAN MICHAEL O\'BRIEN (P-00057) ===
  ['P-00057','Birthplace','1980-05-14','Exact','USA','Rhode Island','Providence','Women & Infants Hospital','','','Strong','Sean Michael O\'Brien born Providence.'],
  ['P-00057','Residence','2003-01-01','Year Only','USA','Rhode Island','Providence','Providence — law practice','','','Strong','Sean Michael O\'Brien practices law in Providence.'],
  // === DECLAN O\'BRIEN (P-00092) ===
  ['P-00092','Birthplace','1982-07-30','Exact','USA','Rhode Island','Providence','Women & Infants Hospital','','','Strong','Declan O\'Brien born Providence.'],
  ['P-00092','Residence','2005-01-01','Year Only','USA','Rhode Island','Providence','Providence Police Department area','','','Strong','Declan is a Providence police officer.'],
  // === COUNTY CORK ANCESTORS — MIGRATION TRAIL ===
  ['P-00022','Birthplace','1850-03-15','Year Only','Ireland','County Cork','County Cork','Shaughnessy family area','','SRC-00022','Moderate','Daniel Shaughnessy born Cork.'],
  ['P-00022','Residence','1876-01-01','Year Only','Ireland','County Cork','County Cork','Remained in County Cork','','SRC-00022','Moderate','Daniel Shaughnessy stayed in Cork, Ireland.'],
  ['P-00061','Birthplace','1775-01-01','Estimated','Ireland','County Cork','County Cork','Pre-civil registration','','SRC-00045','Tentative','Thaddeus O\'Brien born Cork. Highly approximate.'],
  // === JOHN GALLAGHER (P-00041) ===
  ['P-00041','Birthplace','1855-01-17','Year Only','Ireland','County Clare','County Clare','Rural Clare','','SRC-00033','Tentative','John Gallagher born County Clare.'],
  ['P-00041','Residence','1875-01-01','Year Only','Ireland','County Clare','County Clare','Gallagher family farm','','SRC-00033','Tentative','John Gallagher remained in County Clare.'],
  ['P-00041','Death Location','1920-11-08','Year Only','Ireland','County Clare','County Clare','Died in Clare','','SRC-00033','Tentative','John Gallagher died County Clare.'],
  // === JAMES PATRICK O\'BRIEN (P-00100) ===
  ['P-00100','Birthplace','1940-02-22','Exact','USA','Rhode Island','Providence','Rhode Island Hospital','','SRC-00069','Strong','James Patrick O\'Brien born Providence.'],
  ['P-00100','Childhood Residence','1940-01-01','Year Only','USA','Rhode Island','Providence','Family home Providence','','SRC-00069','Strong','James grew up in Providence.'],
  ['P-00100','Employment Location','1960-01-01','Year Only','USA','Rhode Island','Providence','Providence Fire Department','','SRC-00069','Strong','James Patrick O\'Brien — fire chief career, Providence.'],
  ['P-00100','Residence','1965-01-01','Year Only','USA','Rhode Island','Providence','Family home, South Providence','','SRC-00069','Strong','James Patrick O\'Brien lives in Providence (retired).'],
  // === FREDERICK HARTWELL (P-00047) ===
  ['P-00047','Birthplace','1815-06-15','About','England','Warwickshire','Warwickshire','Coventry/Warwickshire area','','SRC-00037','Moderate','Frederick Hartwell born Warwickshire, England.'],
  ['P-00047','Residence','1840-01-01','Year Only','England','West Midlands','Coventry','Moved to Coventry weaving trade','','SRC-00037','Moderate','Frederick Hartwell in Coventry.'],
  ['P-00047','Death Location','1880-03-22','About','England','West Midlands','Coventry','Died Coventry','','SRC-00037','Moderate','Frederick Hartwell died Coventry.'],
];

(async () => {
  const reqs = [];
  const vals = [];

  // Tab setup
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, tabColor: hex(C.aqua), gridProperties: { frozenRowCount: 7 } },
    fields: 'tabColor,gridProperties.frozenRowCount',
  }});

  const deepAqua = '#4A8A88';

  // Row 1 — Title
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 0, 1, 0, 14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(deepAqua),
      textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  vals.push({ range: `'${S}'!A1`, values: [['LOCATIONS & MIGRATION HISTORY']] });

  // Row 2 — Subtitle
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 1, 2, 0, 14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.aqua),
      textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  vals.push({ range: `'${S}'!A2`, values: [['Track birthplaces, residences, immigration journeys, military postings, and migration paths across generations.']] });

  // Row 3-4 — Stats
  const statLabels = ['Total Records','Birthplaces','Residences','Immigration','Military','Countries'];
  const statCols = [0,2,4,6,8,10];
  statLabels.forEach((lbl, i) => {
    const c = statCols[i];
    reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, c, c+2), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, c, c+2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, 2, 3, c, c+2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(deepAqua),
        textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(SID, 3, 4, c, c+2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.input),
        textFormat: { bold: true, fontSize: 11, foregroundColor: hex(deepAqua), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});
    vals.push({ range: `'${S}'!${colL(c)}3`, values: [[lbl]] });
  });
  const statFormulas = [
    `=COUNTA($B$8:$B$8007)`,
    `=COUNTIF($D$8:$D$8007,"Birthplace")`,
    `=COUNTIF($D$8:$D$8007,"Residence")`,
    `=COUNTIFS($D$8:$D$8007,"Immigration Origin")+COUNTIFS($D$8:$D$8007,"Immigration Destination")`,
    `=COUNTIFS($D$8:$D$8007,"Military Location")`,
    `=SUMPRODUCT(IFERROR((LEN($G$8:$G$8007)>0)*(1/COUNTIF($G$8:$G$8007,$G$8:$G$8007)),0))`,
  ];
  statFormulas.forEach((f, i) => {
    vals.push({ range: `'${S}'!${colL(statCols[i])}4`, values: [[f]] });
  });

  // Row 5-6 — spacer
  reqs.push({ repeatCell: {
    range: gridRange(SID, 4, 6, 0, 14),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat(backgroundColor)',
  }});

  // Row 7 — Headers
  reqs.push({ repeatCell: {
    range: gridRange(SID, 6, 7, 0, 14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(deepAqua),
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

  // Column widths: A=90, B=90, C=180, D=160, E=100, F=100, G=100, H=160, I=160, J=160, K=130, L=100, M=120, N=240
  const colWidths = [90,90,180,160,100,100,100,160,160,160,130,100,120,240];
  colWidths.forEach((px, ci) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  // Border header
  reqs.push({ updateBorders: {
    range: gridRange(SID, 6, 7, 0, 14),
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.border) },
  }});

  // Conditional formatting — Location Event Type (col D = index 3)
  const CF_LOC = [
    { val: 'Birthplace',              bg: '#C8E6C9' },
    { val: 'Immigration Origin',      bg: '#FFF9C4' },
    { val: 'Immigration Destination', bg: '#FFF176' },
    { val: 'Military Location',       bg: '#FFCCBC' },
    { val: 'Death Location',          bg: '#CFD8DC' },
    { val: 'Burial Location',         bg: '#ECEFF1' },
    { val: 'Residence',               bg: C.formula },
    { val: 'Census Residence',        bg: C.formula },
  ];
  CF_LOC.forEach(({ val, bg }) => {
    reqs.push({ addConditionalFormatRule: { rule: {
      ranges: [gridRange(SID, 7, 8007, 3, 4)],
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: val }] },
        format: { backgroundColor: hex(bg) },
      },
    }, index: 0 }});
  });

  // Conditional formatting — Evidence Confidence (col M = index 12)
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
      ranges: [gridRange(SID, 7, 8007, 12, 13)],
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: val }] },
        format: { backgroundColor: hex(bg) },
      },
    }, index: 0 }});
  });

  // Data validation
  // D = Location Event Types (col 3) — Reference Data!$G$2:$G$14
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 8007, 3, 4),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$G$2:$G$14` }] },
      strict: false, showCustomUi: true,
    },
  }});
  // F = Date Precision (col 5) — Reference Data!$F$2:$F$9
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 8007, 5, 6),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$F$2:$F$9` }] },
      strict: false, showCustomUi: true,
    },
  }});
  // M = Evidence Confidence (col 12) — Reference Data!$J$2:$J$7
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 8007, 12, 13),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$J$2:$J$7` }] },
      strict: false, showCustomUi: true,
    },
  }});

  // Alternating row fill
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, 7, 8007, 0, 14)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND(MOD(ROW(),2)=0,LEN($B8)>0)' }] },
      format: { backgroundColor: hex(C.altRow) },
    },
  }, index: 0 }});

  // Build data rows
  const dataRows = LOCS.map(([pid, locType, date, datePrec, country, stateRegion, city, detail, coords, srcID, conf, notes]) => {
    return [
      `=IF(B{R}="","","LOC-"&TEXT(ROW()-7,"00000"))`,
      pid,
      `=IFERROR(VLOOKUP(B{R},'Master People'!$A$8:$B$5007,2,FALSE),"")`,
      locType,
      date,
      datePrec,
      country,
      stateRegion,
      city,
      detail,
      coords,
      srcID,
      conf,
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

  await batchUpdate(id, reqs, 'loc-fmt');
  await valuesBatchUpdate(id, vals, 'loc-vals');
  await valuesBatchUpdate(id, [{ range: `'${S}'!A8`, values: dataRows }], 'loc-data');
  console.log(`✓ Locations & Migration — ${LOCS.length} rows written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
