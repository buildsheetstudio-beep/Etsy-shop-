'use strict';
const { sheets, batchUpdate, valuesBatchUpdate, gridRange, hex, colL, C } = require('./lib');
const { id, sheetMap } = JSON.parse(require('fs').readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Sources & Citations'];
const S = 'Sources & Citations';

const HEADERS = [
  'Source ID','Source Title','Source Type','Repository','Repository Type',
  'Date Accessed','Publication Date','Author / Creator','URL / Call Number',
  'Persons Cited','Evidence Confidence','Source Quality','Citation Status','Notes',
]; // A-N (14 cols)

// [srcID, title, srcType, repository, repoType, dateAccessed, pubDate, author, urlCallNum, personsCited, conf, quality, citStatus, notes]
const SOURCES = [
  ['SRC-00001','Emma Hartwell Personal Records','Birth Certificate','Personal Collection','Personal Collection','2024-01-10','1988','Hartwell, James (father)','Filed with MA Registry of Vital Records','P-00001','Confirmed','Original Record','Complete Citation','Emma\'s original birth certificate, SS card, and personal documents.'],
  ['SRC-00002','James Edward Hartwell Personal File','Death Certificate','Personal Collection','Personal Collection','2024-01-10','2019','','Death cert. filed MA Vital Records','P-00002, P-00001','Confirmed','Original Record','Complete Citation','James\'s death certificate, marriage certificates, divorce decree, and personal papers.'],
  ['SRC-00003','Margaret O\'Brien Hartwell Personal Records','Birth Certificate','Personal Collection','Personal Collection','2024-01-10','1958','','RI Vital Records','P-00003','Confirmed','Original Record','Complete Citation','Margaret\'s birth certificate and personal documents.'],
  ['SRC-00004','Robert Charles Hartwell Sr. File','Death Certificate','Personal Collection','Personal Collection','2024-01-15','1995','','MA Vital Records','P-00004, P-00005','Strong','Original Record','Complete Citation','Death cert, burial record, 1940 & 1950 census entries for Robert.'],
  ['SRC-00005','Eleanor Mae Chapman Hartwell File','Death Certificate','Personal Collection','Personal Collection','2024-01-15','2003','','MA Vital Records','P-00005','Confirmed','Original Record','Complete Citation','Eleanor\'s death certificate, 1930 & 1940 census entries.'],
  ['SRC-00006','Patrick Joseph O\'Brien File','Death Certificate','St. Michael\'s Cemetery Office, Providence RI','Cemetery','2024-01-15','1998','','Cemetery records, RI Vital Records','P-00006, P-00007','Confirmed','Original Record','Complete Citation','Patrick\'s death cert, burial record, 1930 census Providence.'],
  ['SRC-00007','Catherine Mary Murphy O\'Brien File','Death Certificate','Personal Collection','Personal Collection','2024-01-15','2007','','RI Vital Records','P-00007','Confirmed','Original Record','Complete Citation','Catherine\'s death cert and 1930 census entry.'],
  ['SRC-00008','William Henry Hartwell Immigration and Vital Records','Passenger List','National Archives, Washington DC','National Archive','2024-02-01','1910','','NARA, RG 85, Ellis Island Ship Manifest 1910','P-00008, P-00009','Confirmed','Original Record','Complete Citation','William and Agnes\'s 1910 ship manifest. Baptism records from Holy Trinity, Coventry (copy).'],
  ['SRC-00009','Agnes Louise Fletcher Hartwell File','Church Record','Holy Trinity Church, Coventry, England','Church','2024-02-01','1892','','Coventry Diocese Archives, ref. HTR/1892/F','P-00009','Confirmed','Original Record','Complete Citation','Agnes\'s baptism record. Obtained from Coventry Diocese.'],
  ['SRC-00010','Charles Frederick Chapman File','Census','National Archives, Washington DC','National Archive','2024-02-10','1900,1910,1920','','NARA, 1900 CT Census, New Haven Co., Hartford','P-00010, P-00011','Strong','Original Record','Partial Citation','1900, 1910, and 1920 census entries for Charles Chapman family.'],
  ['SRC-00011','Harriet Jane Davies Chapman File','Death Certificate','Springfield City Clerk','Local Archive','2024-02-10','1971','','Springfield MA Death Certificate No. 1971-0088','P-00011','Confirmed','Original Record','Complete Citation','Harriet\'s death certificate and 1920 census.'],
  ['SRC-00012','Michael Patrick O\'Brien Immigration and Family File','Passenger List','National Archives, Washington DC','National Archive','2024-02-15','1903','','NARA, RG 85, Manifest List No. 7, SS Oceanic 1903','P-00012, P-00013','Confirmed','Original Record','Complete Citation','Michael\'s 1903 ship manifest from Queenstown. Irish parish records copy. 1920 & 1930 census.'],
  ['SRC-00013','Brigid Nora Shaughnessy O\'Brien Baptism Record','Church Record','Roman Catholic Diocese of Cloyne','Church','2024-02-15','1886','','Cloyne Diocesan Archives, Baptism Register Vol. 14','P-00013','Confirmed','Original Record','Complete Citation','Brigid\'s baptism record from Cloyne Parish, County Cork.'],
  ['SRC-00014','Thomas Francis Murphy File','Church Record','Roman Catholic Diocese of Killaloe','Church','2024-02-20','1886','','Killaloe Archives, Kilmihil Parish Register 1880-1895','P-00014, P-00015','Confirmed','Original Record','Complete Citation','Thomas Murphy baptism record and 1912 marriage record, Providence RI.'],
  ['SRC-00015','Mary Josephine Gallagher Murphy File','Death Certificate','Providence City Clerk','Local Archive','2024-02-20','1955','','Providence RI Death Certificate No. 1955-1122','P-00015','Confirmed','Original Record','Complete Citation','Mary Murphy\'s death certificate. 1920 & 1930 census entries.'],
  ['SRC-00016','George Thomas Hartwell England Records','Obituary','Coventry Evening Telegraph Archives','Newspaper Archive','2024-03-01','1920','','Coventry Evening Telegraph, 1920 Death Notice collection','P-00016, P-00017','Moderate','Derivative Record','Partial Citation','Death notice for George Hartwell. Marriage index from England & Wales GRO.'],
  ['SRC-00017','Martha Ellen Holt Hartwell England Records','Church Record','Coventry Diocese Archives','Church','2024-03-01','1925','','Coventry Diocese ref. STM/1854/H','P-00017','Moderate','Original Record','Partial Citation','Martha Holt\'s baptism register entry. Death inferred from 1925 England index.'],
  ['SRC-00018','Edward James Fletcher Conflicting Records','Probate Record','Coventry Probate Registry','Courthouse','2024-03-05','1930','','Coventry Probate, Case No. 1930/F-211','P-00018, P-00019','Tentative','Derivative Record','Needs Review','Edward Fletcher probate record conflicts with baptism date. Two possible matches in index.'],
  ['SRC-00019','Sarah Ann Pickles Fletcher Record','Birth Certificate','General Register Office, England','National Archive','2024-03-05','1862','','GRO Index, Warwickshire, 1862 Q3, Vol. 6d p.44','P-00019','Tentative','Original Record','Partial Citation','Sarah Pickles birth index entry. Marriage to Edward Fletcher inferred only.'],
  ['SRC-00020','Cornelius O\'Brien Ballycotton Parish Records','Church Record','Ballycotton Parish, Cork','Church','2024-03-10','1845','','Ballycotton RC Parish Register 1840-1870, digitized NLI','P-00020, P-00021','Tentative','Original Record','Partial Citation','Cornelius O\'Brien baptism estimated from damaged parish register. NLI ref B-BC-1845.'],
  ['SRC-00021','Honora Crowley O\'Brien and Crowley Family Records','Land / Property Record','Griffith\'s Valuation Archive, Dublin','National Archive','2024-03-10','1851','','Griffith\'s Valuation 1848-1864, County Cork, Ballycotton area','P-00021, P-00059, P-00066','Moderate','Original Record','Partial Citation','Nora O\'Connell Crowley and husband found in 1851 Griffith\'s Valuation. Key discovery.'],
  ['SRC-00022','Shaughnessy Family Cork Records','Church Record','Cloyne Diocese Archive','Church','2024-03-15','1876','','Cloyne Marriage Registers Vol. 8, 1870-1880','P-00022, P-00023','Moderate','Original Record','Partial Citation','Shaughnessy-Sullivan marriage record found in Cloyne registers.'],
  ['SRC-00023','Murphy and Ryan Family Clare Records','Church Record','Roman Catholic Diocese of Killaloe','Church','2024-03-20','1878','','Killaloe Marriage Registers, Clare 1875-1885','P-00024, P-00025','Moderate','Original Record','Partial Citation','Patrick Murphy and Bridget Ryan marriage. Clare baptism registers for Thomas and siblings.'],
  ['SRC-00024','Thomas James Hartwell Personal File','Birth Certificate','Personal Collection','Personal Collection','2024-04-01','1983','','MA Vital Records Ref. 1983-WOR-0911','P-00026','Confirmed','Original Record','Complete Citation','Thomas Hartwell birth certificate.'],
  ['SRC-00025','Sarah Elizabeth Hartwell Personal File','Birth Certificate','Personal Collection','Personal Collection','2024-04-01','1990','','MA Vital Records Ref. 1990-WOR-0507','P-00027','Confirmed','Original Record','Complete Citation','Sarah Hartwell birth certificate.'],
  ['SRC-00026','Linda Grace Patel Hartwell Personal File','Birth Certificate','Personal Collection','Personal Collection','2024-04-05','1985','','Cook County IL Vital Records','P-00028','Confirmed','Original Record','Complete Citation','Linda Patel\'s birth certificate. Marriage cert Thomas & Linda 2009.'],
  ['SRC-00027','Richard Allen Hartwell Death Certificate','Death Certificate','Worcester City Clerk','Local Archive','2024-04-10','2020','','Worcester MA Death Cert No. 2020-0107','P-00032, P-00033','Confirmed','Original Record','Complete Citation','Richard Hartwell\'s death certificate and obituary.'],
  ['SRC-00028','Dorothy Mae Simmons Hartwell File','Birth Certificate','Worcester City Clerk','Local Archive','2024-04-10','1954','','Worcester MA Vital Records 1954','P-00033','Confirmed','Original Record','Complete Citation','Dorothy Simmons birth cert and marriage record.'],
  ['SRC-00029','Sean Patrick O\'Brien Personal File','Birth Certificate','Personal Collection','Personal Collection','2024-05-01','1956','','RI Vital Records 1956','P-00036','Confirmed','Original Record','Complete Citation','Sean O\'Brien birth certificate.'],
  ['SRC-00030','Colleen Frances O\'Brien Donovan File','Death Certificate','Providence City Clerk','Local Archive','2024-05-01','2015','','Providence RI Death Cert No. 2015-0402','P-00037, P-00054, P-00055, P-00056','Confirmed','Original Record','Complete Citation','Colleen Donovan death cert, marriage cert 1982, and family records.'],
  ['SRC-00031','Arthur Raymond Chapman File','Death Certificate','Springfield City Clerk','Local Archive','2024-05-15','1962','','Springfield MA Death Cert No. 1962-0605','P-00039, P-00040','Moderate','Original Record','Partial Citation','Arthur Chapman death cert and 1930 census entry.'],
  ['SRC-00032','Rose Margaret Kelly Chapman File','Death Certificate','Springfield City Clerk','Local Archive','2024-05-15','1978','','Springfield MA Death Cert No. 1978-0930','P-00040','Moderate','Original Record','Partial Citation','Rose Kelly Chapman death cert.'],
  ['SRC-00033','Gallagher Family Clare Records','Church Record','Diocese of Killaloe Archives','Church','2024-05-20','1855','','Killaloe Baptism Register, Clare 1850-1860','P-00041, P-00042, P-00088','Tentative','Original Record','Partial Citation','John Gallagher baptism record. Limited Clare records from this period.'],
  ['SRC-00034','Anne Connelly Gallagher Record','Church Record','Diocese of Killaloe Archives','Church','2024-05-20','1858','','Killaloe Baptism Register, Clare 1855-1865','P-00042','Tentative','Original Record','Partial Citation','Anne Connelly baptism record from Killaloe Diocese.'],
  ['SRC-00035','Henry Fletcher England Records','Church Record','Birmingham Diocese Archives','Church','2024-06-01','1860','','Birmingham RC Baptism Register 1858-1865','P-00043','Tentative','Original Record','Needs Review','Henry Fletcher baptism record. May be Edward\'s brother or unrelated.'],
  ['SRC-00036','Elizabeth Anne Hartwell Infant Death Record','Death Certificate','Springfield City Clerk','Local Archive','2024-06-05','1919','','Springfield MA Death Cert No. 1919-0115','P-00044','Confirmed','Original Record','Complete Citation','Infant Elizabeth Hartwell died 6 weeks old. Death certificate.'],
  ['SRC-00037','Frederick George Hartwell England Records','Obituary','Coventry Evening Telegraph Archives','Newspaper Archive','2024-06-15','1880','','Coventry Evening Telegraph 1880 Death Notices','P-00047, P-00048','Moderate','Derivative Record','Partial Citation','Frederick Hartwell death notice. England & Wales GRO marriage index.'],
  ['SRC-00038','Ann Elizabeth Turner Hartwell England Records','Census','The National Archives, Kew, England','National Archive','2024-06-15','1851','','TNA HO107/2086, Warwickshire 1851 Census','P-00048','Moderate','Original Record','Partial Citation','Ann Turner found in 1851 England census. Birth inferred.'],
  ['SRC-00039','Jeremiah O\'Brien Pre-Famine Records','Church Record','Ballycotton Parish Archives','Church','2024-06-20','1808','','Ballycotton RC Register 1800-1830, damaged','P-00049, P-00050','Tentative','Original Record','Not Cited','Jeremiah O\'Brien and Mary Driscoll — inferred from damaged famine-era register.'],
  ['SRC-00040','Mary Driscoll O\'Brien Famine Record','Church Record','Ballycotton Parish Archives','Church','2024-06-20','1812','','Ballycotton RC Register 1810-1830','P-00050','Tentative','Original Record','Not Cited','Mary Driscoll — baptism record barely legible. Possibly died in famine 1848.'],
  ['SRC-00041','Brian Francis Walsh Obituary','Obituary','Worcester Telegram & Gazette','Newspaper Archive','2024-07-01','2018','','Worcester Telegram & Gazette, June 9, 2018','P-00052','Confirmed','Original Record','Complete Citation','Brian Walsh obituary. Father of Jennifer Walsh.'],
  ['SRC-00042','Maureen Donovan Personal Records','Birth Certificate','Personal Collection','Personal Collection','2024-07-05','1958','','RI Vital Records 1958','P-00053','Confirmed','Original Record','Partial Citation','Maureen Donovan personal records.'],
  ['SRC-00043','Patrick Sean Donovan Personal File','Birth Certificate','Personal Collection','Personal Collection','2024-07-05','1957','','RI Vital Records 1957','P-00054','Confirmed','Original Record','Partial Citation','Patrick Donovan birth cert.'],
  ['SRC-00044','James Aloysius Flynn Conflicting Parish Records','Church Record','Roman Catholic Diocese of Cork and Ross','Church','2024-07-20','1884','','Cork & Ross Diocese Archives — two conflicting registers, refs CF/1884/A and CF/1884/B','P-00060','Conflicting','Original Record','Needs Review','Two parish records disagree on whether Flynn is an O\'Brien half-sibling. First record suggests shared parentage; second contradicts. Awaiting DNA evidence.'],
  ['SRC-00045','Thaddeus O\'Brien Estimated Records','Family Tree','Genealogy Website','Genealogy Website','2024-07-25','','','Ancestry.com tree, uploaded by user OBrienDescendants_RI, 2019','P-00061, P-00062','Tentative','User-Contributed Tree','Not Cited','Online family tree suggests Thaddeus and Catherine McCarthy as Jeremiah\'s parents. Not verified against primary sources.'],
  ['SRC-00046','Thomas and Jane Hartwell Estimated Records','Family Tree','Genealogy Website','Genealogy Website','2024-08-01','','','FindMyPast family tree, user HartwellResearcher, 2020','P-00063, P-00064','Tentative','User-Contributed Tree','Not Cited','Unverified family tree entry for Thomas Hartwell and Jane Booth as Frederick\'s parents.'],
  ['SRC-00047','Nora O\'Connell Crowley Discovery — Griffith\'s Valuation','Land / Property Record','Griffith\'s Valuation Archive, Dublin','National Archive','2024-08-10','1851','Griffith, Richard (Commissioner)','Griffith\'s Primary Valuation of Tenements 1848-1864','P-00066','Moderate','Original Record','Complete Citation','Key discovery: Nora O\'Connell as wife of a Crowley in Cork parish 1851.'],
  ['SRC-00048','Alice Victoria Holt Fletcher Multiple Sources','Birth Certificate','General Register Office, England','National Archive','2024-08-15','1825','','GRO Index England & Wales 1825 Q1, Coventry Vol. 6a, p.12','P-00067','Strong','Original Record','Complete Citation','Three sources confirm Alice Holt: GRO birth index, 1851 census, and church marriage record.'],
  ['SRC-00049','Domenico Rossi Immigration and Personal Records','Passenger List','National Archives, Washington DC','National Archive','2024-08-20','1935','','NARA, RG 85, Manifest, SS Conte di Savoia 1935','P-00068, P-00069','Strong','Original Record','Complete Citation','Domenico Rossi ship manifest 1935. Italian birth record (comune di Palermo copy).'],
  ['SRC-00050','Rosa Carmela Marino Rossi Personal File','Birth Certificate','Personal Collection','Personal Collection','2024-08-20','1915','','Palermo Civil Registry — comune copy','P-00069','Strong','Original Record','Complete Citation','Rosa Marino birth certificate from Palermo civil registry.'],
  ['SRC-00051','Joseph Bernard O\'Brien WWII Military Records','Military Record','National Archives, Washington DC','National Archive','2024-09-05','1945','Department of the Army','NARA, IDPF (Individual Deceased Personnel File), O\'Brien, Joseph Bernard, Army Serial No. 31-xxx','P-00071','Confirmed','Original Record','Complete Citation','Joseph O\'Brien KIA D-Day 1945. IDPF, burial record, Normandy American Cemetery.'],
  ['SRC-00052','Margaret Alice O\'Brien 1930 Census Entry','Census','National Archives, Washington DC','National Archive','2024-09-10','1930','','NARA, 1930 Federal Census, Providence RI, ED 30-123','P-00072','Moderate','Original Record','Partial Citation','Last known record of Margaret Alice O\'Brien, age 19, in 1930 census.'],
  ['SRC-00053','William James O\'Brien Personal File','Death Certificate','Providence City Clerk','Local Archive','2024-09-15','1980','','Providence RI Death Cert No. 1980-0410','P-00073','Confirmed','Original Record','Complete Citation','William James O\'Brien death certificate and 1930-1940 census entries.'],
  ['SRC-00054','Kathleen Agnes O\'Brien Carey File','Death Certificate','Providence City Clerk','Local Archive','2024-09-15','2005','','Providence RI Death Cert No. 2005-1122','P-00074','Confirmed','Original Record','Complete Citation','Kathleen O\'Brien Carey death certificate.'],
  ['SRC-00055','Robert Francis Hartwell WWII Military Records','Military Record','National Archives, Washington DC','National Archive','2024-09-20','1943','US Marine Corps','NARA, IDPF Robert Francis Hartwell, USMC Serial No. 378xxx','P-00075','Confirmed','Original Record','Complete Citation','Robert Francis Hartwell KIA Tarawa 1943. IDPF and Punchbowl burial record.'],
  ['SRC-00056','Mildred Jean Hartwell Perry File','Death Certificate','Hampden County Clerk','Local Archive','2024-09-20','2010','','Hampden County MA Death Cert No. 2010-0819','P-00076','Confirmed','Original Record','Complete Citation','Mildred Perry death certificate. Birth cert and 1930-1940 census.'],
  ['SRC-00057','Howard Earl Perry Personal File','Death Certificate','Hampden County Clerk','Local Archive','2024-09-20','1998','','Hampden County MA Death Cert No. 1998-0228','P-00077','Confirmed','Original Record','Complete Citation','Howard Perry death certificate.'],
  ['SRC-00058','Timothy Joseph Murphy Personal File','Death Certificate','Providence City Clerk','Local Archive','2024-10-01','1988','','Providence RI Death Cert No. 1988-0704','P-00080','Confirmed','Original Record','Complete Citation','Timothy Murphy death certificate.'],
  ['SRC-00059','Frances Rose Murphy Kelly Personal File','Death Certificate','Providence City Clerk','Local Archive','2024-10-01','2001','','Providence RI Death Cert No. 2001-1030','P-00081','Confirmed','Original Record','Complete Citation','Frances Kelly death cert and 1940 census entry.'],
  ['SRC-00060','Gerard Thomas Kelly Personal File','Death Certificate','Providence City Clerk','Local Archive','2024-10-01','1975','','Providence RI Death Cert No. 1975-0418','P-00082','Confirmed','Original Record','Complete Citation','Gerard Kelly death certificate.'],
  ['SRC-00061','Turner Family Warwickshire Census Records','Census','The National Archives, Kew, England','National Archive','2024-10-10','1851,1861','','TNA HO107/2086 & RG9/2193, Warwickshire','P-00084, P-00085','Tentative','Original Record','Partial Citation','John and Hannah Turner in 1851 and 1861 England census.'],
  ['SRC-00062','Salvatore Marino Immigration Records','Passenger List','National Archives, Washington DC','National Archive','2024-10-15','1912','','NARA, RG 85, Manifest, SS Konig Albert 1912','P-00086, P-00087','Moderate','Original Record','Complete Citation','Salvatore Marino ship manifest 1912. Arrived Providence via Ellis Island.'],
  ['SRC-00063','Giuseppa Ferrara Marino Personal Records','Birth Certificate','Comune di Palermo Archives','Local Archive','2024-10-15','1890','','Palermo civil birth register 1890','P-00087','Moderate','Original Record','Partial Citation','Giuseppa Ferrara birth record from Palermo.'],
  ['SRC-00064','Alice Booth Warwickshire Records','Census','The National Archives, Kew, England','National Archive','2024-10-23','1851','','TNA HO107/2086, Warwickshire 1851 Census — Booth household','P-00091','Moderate','Original Record','Partial Citation','Alice Booth found in 1851 England census with daughter Jane.'],
  ['SRC-00065','Martin Joseph O\'Brien Family File','Death Certificate','Providence City Clerk','Local Archive','2024-10-25','1960','','Providence RI Death Cert No. 1960-1104','P-00096, P-00097','Confirmed','Original Record','Complete Citation','Martin O\'Brien death cert and 1920-1930 census entries.'],
  ['SRC-00066','Anastasia Bridget Shea O\'Brien File','Death Certificate','Providence City Clerk','Local Archive','2024-10-25','1968','','Providence RI Death Cert No. 1968-0823','P-00097','Confirmed','Original Record','Complete Citation','Anastasia O\'Brien death certificate.'],
  ['SRC-00067','Patrick Martin O\'Brien File','Death Certificate','Providence City Clerk','Local Archive','2024-10-25','1985','','Providence RI Death Cert No. 1985-0114','P-00098, P-00099','Confirmed','Original Record','Complete Citation','Patrick Martin O\'Brien death cert and marriage record 1938.'],
  ['SRC-00068','Evelyn Grace McCarthy O\'Brien File','Death Certificate','Providence City Clerk','Local Archive','2024-10-25','1992','','Providence RI Death Cert No. 1992-0618','P-00099','Confirmed','Original Record','Complete Citation','Evelyn McCarthy O\'Brien death certificate.'],
  ['SRC-00069','James Patrick O\'Brien Personal File','Birth Certificate','Personal Collection','Personal Collection','2024-10-26','1940','','RI Vital Records 1940','P-00100, P-00101','Strong','Original Record','Partial Citation','James Patrick O\'Brien birth cert and personal records.'],
  ['SRC-00070','Dorothy Ann Walsh O\'Brien Personal File','Birth Certificate','Personal Collection','Personal Collection','2024-10-26','1943','','RI Vital Records 1943','P-00101','Strong','Original Record','Partial Citation','Dorothy Walsh O\'Brien birth cert.'],
  ['SRC-00071','Cecil Edwin Hartwell England Death Notice','Obituary','Coventry Evening Telegraph Archives','Newspaper Archive','2024-10-26','1932','','Coventry Evening Telegraph, January 1932','P-00102','Tentative','Derivative Record','Partial Citation','Cecil Hartwell death notice. Suggests sibling relationship to George Hartwell.'],
];

(async () => {
  const reqs = [];
  const vals = [];

  const deepWheat = '#9A7D3A';

  // Tab setup
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, tabColor: hex(C.wheat), gridProperties: { frozenRowCount: 7 } },
    fields: 'tabColor,gridProperties.frozenRowCount',
  }});

  // Row 1 — Title
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 0, 1, 0, 14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(deepWheat),
      textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  vals.push({ range: `'${S}'!A1`, values: [['SOURCES & CITATIONS LIBRARY']] });

  // Row 2 — Subtitle
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 1, 2, 0, 14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.wheat),
      textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  vals.push({ range: `'${S}'!A2`, values: [['All primary and secondary sources used in this family tree. Every source ID referenced in other sheets links back here.']] });

  // Row 3-4 — Stats
  const statLabels = ['Total Sources','Original Records','Derivative Records','Cited Completely','Needs Review','User Trees'];
  const statCols = [0,2,4,6,8,10];
  statLabels.forEach((lbl, i) => {
    const c = statCols[i];
    reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, c, c+2), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, c, c+2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, 2, 3, c, c+2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(deepWheat),
        textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(SID, 3, 4, c, c+2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.input),
        textFormat: { bold: true, fontSize: 11, foregroundColor: hex(deepWheat), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});
    vals.push({ range: `'${S}'!${colL(c)}3`, values: [[lbl]] });
  });
  const statFormulas = [
    `=COUNTA($A$8:$A$6007)`,
    `=COUNTIF($L$8:$L$6007,"Original Record")`,
    `=COUNTIF($L$8:$L$6007,"Derivative Record")`,
    `=COUNTIF($M$8:$M$6007,"Complete Citation")`,
    `=COUNTIF($M$8:$M$6007,"Needs Review")`,
    `=COUNTIF($L$8:$L$6007,"User-Contributed Tree")`,
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
      backgroundColor: hex(deepWheat),
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
  // Taller data rows to handle long title text
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 6007 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Column widths: A=90, B=260, C=130, D=200, E=130, F=100, G=100, H=160, I=220, J=180, K=120, L=130, M=110, N=260
  const colWidths = [90,260,130,200,130,100,100,160,220,180,120,130,110,260];
  colWidths.forEach((px, ci) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  // Wrap text for title and notes columns
  reqs.push({ repeatCell: {
    range: gridRange(SID, 7, 6007, 1, 2),
    cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat(wrapStrategy)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(SID, 7, 6007, 13, 14),
    cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat(wrapStrategy)',
  }});

  // Border header
  reqs.push({ updateBorders: {
    range: gridRange(SID, 6, 7, 0, 14),
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.border) },
  }});

  // Conditional formatting — Source Quality (col L = index 11)
  const CF_QUALITY = [
    { val: 'Original Record',       bg: C.confirmed },
    { val: 'Derivative Record',     bg: C.secondary },
    { val: 'Authored Narrative',    bg: C.info },
    { val: 'Oral History',          bg: C.review },
    { val: 'User-Contributed Tree', bg: C.conflict },
    { val: 'Unknown',               bg: C.neutral },
  ];
  CF_QUALITY.forEach(({ val, bg }) => {
    reqs.push({ addConditionalFormatRule: { rule: {
      ranges: [gridRange(SID, 7, 6007, 11, 12)],
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: val }] },
        format: { backgroundColor: hex(bg) },
      },
    }, index: 0 }});
  });

  // Conditional formatting — Citation Status (col M = index 12)
  const CF_CIT = [
    { val: 'Complete Citation',  bg: C.confirmed },
    { val: 'Partial Citation',   bg: C.info },
    { val: 'Needs Review',       bg: C.review },
    { val: 'Not Cited',          bg: C.conflict },
  ];
  CF_CIT.forEach(({ val, bg }) => {
    reqs.push({ addConditionalFormatRule: { rule: {
      ranges: [gridRange(SID, 7, 6007, 12, 13)],
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: val }] },
        format: { backgroundColor: hex(bg) },
      },
    }, index: 0 }});
  });

  // Data validation
  // C = Source Types (col 2) — Reference Data!$H$2:$H$29
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 6007, 2, 3),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$H$2:$H$29` }] },
      strict: false, showCustomUi: true,
    },
  }});
  // E = Repository Types (col 4) — Reference Data!$I$2:$I$15
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 6007, 4, 5),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$I$2:$I$15` }] },
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
  // L = Source Quality (col 11) — Reference Data!$K$2:$K$7
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 6007, 11, 12),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$K$2:$K$7` }] },
      strict: false, showCustomUi: true,
    },
  }});
  // M = Citation Status (col 12) — Reference Data!$L$2:$L$5
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 6007, 12, 13),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$L$2:$L$5` }] },
      strict: false, showCustomUi: true,
    },
  }});

  // Alternating row fill
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, 7, 6007, 0, 14)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND(MOD(ROW(),2)=0,LEN($A8)>0)' }] },
      format: { backgroundColor: hex(C.altRow) },
    },
  }, index: 0 }});

  // Data rows — source ID is already static (A col), no formula needed
  const dataRows = SOURCES.map(([srcID, title, srcType, repo, repoType, dateAccessed, pubDate, author, urlCallNum, personsCited, conf, quality, citStatus, notes]) => [
    srcID, title, srcType, repo, repoType, dateAccessed, pubDate, author, urlCallNum, personsCited, conf, quality, citStatus, notes,
  ]);

  await batchUpdate(id, reqs, 'src-fmt');
  await valuesBatchUpdate(id, vals, 'src-vals');
  await valuesBatchUpdate(id, [{ range: `'${S}'!A8`, values: dataRows }], 'src-data');
  console.log(`✓ Sources & Citations — ${SOURCES.length} rows written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
