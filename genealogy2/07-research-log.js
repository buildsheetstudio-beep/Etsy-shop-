'use strict';
const { sheets, batchUpdate, valuesBatchUpdate, gridRange, hex, colL, C } = require('./lib');
const { id, sheetMap } = JSON.parse(require('fs').readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Research Log'];
const S = 'Research Log';

const HEADERS = [
  'Log ID','Date','Person ID','Person Name','Research Goal','Repository / Resource',
  'Research Status','Research Outcome','Source Found','Evidence Confidence',
  'Follow-Up Required','Next Steps / Notes',
]; // A-L (12 cols)

// [date, personID, goal, repository, status, outcome, srcFound, confidence, followUp, nextSteps]
const LOG = [
  ['2024-01-10','P-00001','Verify Emma\'s birth certificate and establish root person','Massachusetts Vital Records Registry','Resolved','Existing Fact Confirmed','SRC-00001','Confirmed','No','Root person established. Birth cert confirmed.'],
  ['2024-01-10','P-00002','Locate James Hartwell\'s death certificate and obituary','Massachusetts Vital Records Registry','Resolved','Source Found','SRC-00002','Confirmed','No','Death cert obtained. Obituary found in Boston Globe archives.'],
  ['2024-01-15','P-00004','Verify Robert Hartwell Sr. birth and death records','MA State Archives, UMass Memorial records','Resolved','Existing Fact Confirmed','SRC-00004','Confirmed','No','Robert\'s records verified: birth cert, death cert, and Pine Grove burial.'],
  ['2024-01-15','P-00006','Confirm Patrick O\'Brien family in Providence records','RI State Archives, St. Michael\'s Cemetery','Resolved','Existing Fact Confirmed','SRC-00006','Confirmed','No','Patrick and Catherine O\'Brien confirmed in Providence records.'],
  ['2024-02-01','P-00008','Find William Hartwell\'s immigration ship manifest','NARA, Ellis Island Foundation database','Resolved','Source Found','SRC-00008','Confirmed','No','Found 1910 ship manifest for William and Agnes Hartwell on SS Carmania.'],
  ['2024-02-01','P-00009','Obtain Agnes Fletcher baptism from Coventry diocese','Coventry Diocese Archives','Resolved','Source Found','SRC-00009','Confirmed','No','Agnes Fletcher 1892 baptism record obtained from Holy Trinity, Coventry.'],
  ['2024-02-10','P-00010','Research Charles Chapman\'s origins in Hartford CT','CT State Archives, 1900/1910 Census','Resolved','Source Found','SRC-00010','Strong','No','Charles found in 1900, 1910, 1920 census records. Birth approx 1895.'],
  ['2024-02-15','P-00012','Trace Michael O\'Brien\'s immigration from County Cork','NARA Passenger Lists, Ellis Island','Resolved','Source Found','SRC-00012','Confirmed','No','1903 ship manifest for Michael O\'Brien found. SS Oceanic, Queenstown to New York.'],
  ['2024-02-15','P-00013','Locate Brigid Shaughnessy baptism in Cork diocese','Roman Catholic Diocese of Cloyne Archives','Resolved','Source Found','SRC-00013','Confirmed','No','Brigid\'s 1886 baptism found at Cloyne Parish. Record clear and complete.'],
  ['2024-02-20','P-00014','Research Thomas Murphy family in County Clare','RC Diocese of Killaloe Archives','Resolved','Source Found','SRC-00014','Confirmed','No','Thomas Murphy 1886 baptism and 1907 emigration confirmed.'],
  ['2024-03-01','P-00016','Investigate George Hartwell in Coventry England records','GRO England & Wales Index, Coventry Archives','Researching','New Person Found','SRC-00016','Moderate','Yes','George found in Coventry death notices 1920. Marriage index located. Need to confirm birth year — two possible matches in GRO index.'],
  ['2024-03-05','P-00018','Clarify conflicting records for Edward Fletcher','Coventry Probate Registry, GRO Index','Conflicting Evidence','Existing Fact Challenged','SRC-00018','Tentative','Yes','Probate record (1930) conflicts with estimated birth date. Two men named Edward Fletcher in Coventry 1890 census. DNA test recommended.'],
  ['2024-03-10','P-00020','Search for Cornelius O\'Brien in Ballycotton parish records','NLI Digitized Parish Registers, Catholic Parish Registers','Researching','Lead Found','SRC-00020','Tentative','Yes','Damaged parish register has partial baptism entry for a Cornelius O\'Brien c.1845. Needs expert document restoration review.'],
  ['2024-03-10','P-00021','Identify Honora Crowley\'s parents in Cork','Griffith\'s Valuation, NLI, Cork Archives','Evidence Found','New Person Found','SRC-00021','Moderate','Yes','Found family in 1851 Griffith\'s Valuation. Nora O\'Connell identified as probable mother. Father surname Crowley but given name unknown.'],
  ['2024-03-15','P-00022','Trace Daniel Shaughnessy family Cork records','Cloyne Diocese Archives','Researching','Source Found','SRC-00022','Moderate','No','Shaughnessy-Sullivan marriage 1876 found in Cloyne registers. Parents not yet identified.'],
  ['2024-04-01','P-00026','Verify Thomas Hartwell birth and current location','MA Vital Records, personal knowledge','Not Started','No Result','','Strong','No','Thomas\'s birth cert on file. Residence Chicago from family knowledge.'],
  ['2024-04-05','P-00028','Confirm Linda Patel\'s background','Cook County IL Records, personal knowledge','Not Started','No Result','SRC-00026','Confirmed','No','Linda\'s birth cert obtained. Marriage cert 2009 obtained.'],
  ['2024-04-10','P-00032','Research Richard Hartwell\'s death and estate','Worcester Telegram & Gazette, MA Vital Records','Resolved','Source Found','SRC-00027','Confirmed','No','Richard\'s death cert and obituary located. Estate probate file open at Worcester Probate Court.'],
  ['2024-05-10','P-00038','Research Helen Hartwell\'s birth family','Adoption records, RI DCYF sealed records','Brick Wall','No Result','','Unknown','Yes','Birth records sealed. Birth mother unknown. DNA testing initiated — no close matches yet. Biological family research stalled.'],
  ['2024-05-20','P-00041','Search for John Gallagher\'s parents in County Clare','Diocese of Killaloe, Griffith\'s Valuation','Brick Wall','No Result','SRC-00033','Tentative','Yes','John Gallagher\'s parents not identified. Limited Clare records from 1855 era. Griffith\'s Valuation shows a Gallagher family but cannot confirm connection.'],
  ['2024-06-05','P-00044','Verify infant Elizabeth Hartwell death record','Springfield City Clerk, Greenwood Cemetery','Resolved','Existing Fact Confirmed','SRC-00036','Confirmed','No','Infant Elizabeth Hartwell death cert confirmed. Burial record at Greenwood Cemetery located.'],
  ['2024-06-10','P-00045','Find Daniel Hartwell birth record and contact info','CT Vital Records, Hartford Hospital','Researching','Lead Found','SRC-00002','Moderate','Yes','Daniel\'s birth found in CT records. Half-sibling status confirmed via James\'s divorce records. No current contact established.'],
  ['2024-06-15','P-00047','Research Frederick Hartwell in Warwickshire England','GRO Index, England & Wales, Coventry archives','Researching','Source Found','SRC-00037','Moderate','Yes','Death notice found 1880. Marriage index suggests wife Ann Turner. Birth year approximate. Need to search 1841-1861 censuses.'],
  ['2024-06-20','P-00049','Search famine-era Jeremiah O\'Brien Cork records','NLI Catholic Parish Registers, Cork Archives','Brick Wall','No Result','SRC-00039','Tentative','Yes','Famine era records extremely sparse. One possible entry in damaged Ballycotton register. Cannot confirm parents or exact dates.'],
  ['2024-07-01','P-00051','Verify Grace Walsh — Jennifer Walsh\'s family connection','Worcester personal contacts, family interview','Not Started','No Result','','Confirmed','No','Grace Walsh identified through family interview. No documentary evidence obtained yet.'],
  ['2024-07-15','P-00059','Research Unknown Crowley — Honora\'s father','Cork Archives, NLI, Church of LDS microfilms','Brick Wall','No Result','SRC-00021','Tentative','Yes','Crowley surname identified from Griffith\'s Valuation but given name not found. Multiple Crowley families in Ballycotton area. Cannot narrow down without DNA.'],
  ['2024-07-20','P-00060','Investigate James Flynn possible half-sibling relationship','RC Diocese of Cork and Ross Archives','Conflicting Evidence','Existing Fact Challenged','SRC-00044','Conflicting','Yes','Two Cork parish records give conflicting parentage for James Flynn. Record A suggests shared O\'Brien parentage; Record B shows different father. DNA evidence needed urgently.'],
  ['2024-07-25','P-00061','Verify Thaddeus O\'Brien as Jeremiah\'s father','Ancestry.com family trees, NLI Cork records','Researching','Lead Found','SRC-00045','Tentative','Yes','Online family tree suggests Thaddeus O\'Brien (b.1775). Cannot verify against primary sources. NLI search in progress.'],
  ['2024-08-05','P-00065','Research Francis Holt — Martha Holt\'s father','Coventry/Birmingham records, GRO Index','Not Started','No Result','','Unknown','Yes','Francis Holt hypothesized as Martha\'s father based on surname match. No records found. Priority: low.'],
  ['2024-08-10','P-00066','Verify Nora O\'Connell Crowley discovery from Griffith\'s','NLI online, Cork Archives','Evidence Found','New Person Found','SRC-00047','Moderate','Yes','Nora O\'Connell confirmed in 1851 Griffith\'s Valuation. Now need to find her maiden name baptism record to establish O\'Connell family.'],
  ['2024-08-15','P-00067','Research Alice Victoria Holt Fletcher — Agnes\'s grandmother','GRO England, Birmingham Diocese, 1851 Census','Resolved','Existing Fact Confirmed','SRC-00048','Strong','No','Three sources confirm Alice Holt: GRO birth 1825, 1851 census, church marriage record. Strong evidence she is Agnes Fletcher\'s grandmother.'],
  ['2024-08-20','P-00068','Research Domenico Rossi immigration from Palermo','NARA Passenger Lists, Ellis Island database','Resolved','Source Found','SRC-00049','Strong','No','Domenico Rossi found on SS Conte di Savoia manifest 1935. Italian birth record obtained from comune di Palermo.'],
  ['2024-09-05','P-00071','Verify Joseph O\'Brien KIA D-Day and burial records','NARA IDPF, Normandy American Cemetery records','Resolved','Existing Fact Confirmed','SRC-00051','Confirmed','No','Joseph Bernard O\'Brien KIA D-Day confirmed. IDPF located. Burial at Normandy American Cemetery, plot reference obtained.'],
  ['2024-09-10','P-00072','Locate Margaret Alice O\'Brien post-1930','Census records 1940, RI Vital Records, SS Death Index','Brick Wall','No Result','SRC-00052','Moderate','Yes','Last record: 1930 census age 19. No 1940 census entry found. No SS death record. No marriage record found. Brick wall — may have changed name or died without record.'],
  ['2024-09-20','P-00075','Verify Robert Francis Hartwell WWII records','NARA IDPF, USMC Archives, Punchbowl records','Resolved','Existing Fact Confirmed','SRC-00055','Confirmed','No','Robert Francis KIA Tarawa 1943 confirmed. Punchbowl burial record located. IDPF retrieved from NARA.'],
  ['2024-10-01','P-00080','Research Timothy Murphy family','Providence RI Vital Records, death cert 1988','Resolved','Existing Fact Confirmed','SRC-00058','Confirmed','No','Timothy Murphy death cert obtained. Confirmed as Thomas and Mary Murphy\'s son.'],
  ['2024-10-10','P-00084','Search for Hannah Turner in Warwickshire records','TNA Kew, 1851/1861 England Census','Researching','Lead Found','SRC-00061','Tentative','Yes','Hannah Turner found in 1851 England census. Birth approx 1790. Need to confirm parentage. 1841 census search needed.'],
  ['2024-10-15','P-00086','Research Salvatore Marino immigration from Palermo','NARA Passenger Lists','Resolved','Source Found','SRC-00062','Moderate','No','Salvatore Marino found on SS Konig Albert manifest 1912. Palermo origin confirmed.'],
  ['2024-10-18','P-00088','Search for Andrew Gallagher after 1900 emigration','Census records USA, SS Death Index, RI Vital Records','Brick Wall','No Result','SRC-00033','Tentative','Yes','Andrew Gallagher emigrated c.1900 from Clare. No matching arrival manifest or US census entry found. Status unknown — may have emigrated to Canada or Australia.'],
  ['2024-10-20','P-00089','Research Sarah Unknown — William Hartwell\'s possible sister','Coventry Archives, family letters','Brick Wall','No Result','','Unknown','Yes','Referenced in a family letter only. No documentary evidence. Coventry parish records search found no matching Sarah Hartwell. Research on hold.'],
  ['2024-10-23','P-00091','Verify Alice Booth in Warwickshire records','TNA Kew, 1851 England Census','Evidence Found','New Person Found','SRC-00064','Moderate','Yes','Alice Booth found in 1851 Warwickshire census with daughter Jane. Now need to find Alice\'s parents to extend the line.'],
  ['2024-10-25','P-00096','Verify Martin O\'Brien emigration and family','Providence RI Vital Records, 1920 census','Resolved','Existing Fact Confirmed','SRC-00065','Confirmed','No','Martin O\'Brien confirmed: emigrated 1903, death cert 1960, 1920 census Providence.'],
  ['2024-10-26','P-00100','Research James Patrick O\'Brien career and current status','Providence Fire Department records, personal knowledge','Not Started','No Result','SRC-00069','Strong','No','James Patrick O\'Brien confirmed retired fire chief. Living Providence. No further research needed at this time.'],
  ['2024-10-26','P-00102','Research Cecil Hartwell in Coventry England records','Coventry Evening Telegraph Archives, GRO Index','Researching','Lead Found','SRC-00071','Tentative','Yes','Cecil Hartwell death notice 1932 found. Probable sibling of George Hartwell. Birth record search in progress — need to confirm shared parentage with Frederick and Ann Turner.'],
  ['2024-10-27','P-00060','Initiate DNA testing for Flynn-O\'Brien connection','AncestryDNA, 23andMe','Researching','Follow-Up Needed','','Conflicting','Yes','DNA kit ordered for living O\'Brien descendant. Awaiting results. This could resolve the Flynn-O\'Brien disputed relationship definitively.'],
  ['2024-10-28','P-00038','DNA testing for Helen Hartwell adoptee search','AncestryDNA adoption registry','Researching','Follow-Up Needed','','Unknown','Yes','Helen\'s DNA uploaded to Ancestry and GEDMatch. One distant match (3rd-4th cousin range) identified. Need to build match\'s tree to identify common ancestor.'],
  // Additional entries to reach 80+
  ['2024-10-29','P-00003','Research Margaret\'s childhood Providence neighborhood','Providence City Archives, RI Historical Society','Resolved','Existing Fact Confirmed','SRC-00003','Confirmed','No','Providence neighborhood confirmed. Elmwood Ave O\'Brien family home identified in 1930 census and city directory.'],
  ['2024-10-29','P-00037','Verify Colleen Donovan marriage and death records','Providence City Clerk, St. Patrick\'s Church','Resolved','Existing Fact Confirmed','SRC-00030','Confirmed','No','Colleen\'s marriage cert (1982) and death cert (2015) obtained. Marriage at St. Patrick\'s on St. Patrick\'s Day confirmed.'],
  ['2024-10-30','P-00047','Search 1841/1851 Warwickshire census for Frederick Hartwell','TNA Kew England Census Microfilm','Researching','Lead Found','SRC-00037','Moderate','Yes','Frederick Hartwell possibly in 1851 Warwickshire census — age 35, weaver. Match plausible but not confirmed. Need to verify against marriage record for Ann Turner.'],
  ['2024-10-30','P-00014','Confirm Thomas Murphy stonemason in Providence directories','Providence City Directories 1907-1938','Resolved','Existing Fact Confirmed','SRC-00014','Confirmed','No','Thomas Murphy listed as stonemason in 1910, 1920, and 1930 Providence City Directories.'],
  ['2024-10-31','P-00012','Search for Michael O\'Brien siblings in Cork','NLI Catholic Parish Registers','Evidence Found','New Person Found','SRC-00065','Confirmed','Yes','Martin O\'Brien confirmed as Michael\'s brother via immigration record (both on same 1903 manifest). Research to find other siblings ongoing.'],
  ['2024-10-31','P-00069','Search for Rosa Marino Rossi\'s Italian birth record','Palermo Civil Registry (comune)','Resolved','Source Found','SRC-00050','Strong','No','Rosa Carmela Marino 1915 birth record obtained from Palermo civil registry. Parents: Salvatore Marino and Giuseppa Ferrara.'],
  ['2024-11-01','P-00058','Research Nicole Rossi O\'Brien family background','Personal knowledge, Rossi family papers','Researching','Lead Found','','Strong','Yes','Nicole\'s grandparents Domenico and Rosa Rossi confirmed. Parents not yet fully documented in this database. Next: obtain Nicole\'s parents\' vital records.'],
  ['2024-11-01','P-00057','Research Sean Michael O\'Brien law career','Rhode Island Bar Association, personal knowledge','Not Started','No Result','','Strong','No','Sean Michael confirmed as licensed Rhode Island attorney. No further genealogical research needed.'],
  ['2024-11-02','P-00092','Verify Declan O\'Brien police career records','Providence Police Department, personal knowledge','Not Started','No Result','','Confirmed','No','Declan confirmed as Providence police officer. No further research needed at this time.'],
  ['2024-11-02','P-00016','Search for George Hartwell\'s parents in Coventry records','Coventry Diocese Archives, England & Wales GRO Index','Researching','Lead Found','SRC-00037','Moderate','Yes','Frederick Hartwell death notice mentions son George. GRO marriage index for Frederick-Ann Turner found. Birth records search for Frederick\'s parents ongoing — Frederick\'s parents tentatively Thomas and Jane Hartwell.'],
  ['2024-11-03','P-00029','Verify Oliver Hartwell birth record','Cook County IL Vital Records','Not Started','No Result','','Strong','No','Oliver born 2012 Chicago. Birth cert available but not yet obtained for this file.'],
  ['2024-11-03','P-00030','Verify Sophia Hartwell birth record','Cook County IL Vital Records','Not Started','No Result','','Strong','No','Sophia born 2015 Chicago. Birth cert available but not obtained.'],
  ['2024-11-04','P-00013','Research Brigid Shaughnessy O\'Brien siblings in Cork','Cloyne Diocese Archives','Researching','Lead Found','SRC-00022','Moderate','Yes','Cloyne register shows other Shaughnessy children in same family. Brigid had at least two siblings — names partially legible. Research continues.'],
  ['2024-11-04','P-00020','Commission expert review of damaged Ballycotton register','NLI Conservation Dept., Cork Archives','Researching','Follow-Up Needed','SRC-00020','Tentative','Yes','NLI conservation team contacted regarding damaged register. Awaiting assessment. This could resolve Cornelius O\'Brien birth date and parentage.'],
  ['2024-11-05','P-00096','Search for Martin O\'Brien children beyond Patrick Martin','Providence RI Vital Records, 1930 Census','Researching','Follow-Up Needed','SRC-00065','Confirmed','Yes','1930 census shows Martin O\'Brien household with at least 3 children. Patrick Martin confirmed. Other children need further research.'],
  ['2024-11-05','P-00098','Research Patrick Martin O\'Brien Korean War era service','NARA Military Records, RI Adjutant General','Researching','Lead Found','SRC-00067','Confirmed','Yes','Patrick Martin O\'Brien b.1912 — possible WWII or early Cold War service record. Age suggests eligible. Search in progress at NARA.'],
  ['2024-11-06','P-00008','Locate William Hartwell naturalization papers','NARA, INS Records, MA District Court records','Researching','Lead Found','SRC-00008','Strong','Yes','William Hartwell applied for naturalization c.1915. Index card found at NARA. Full naturalization petition to be ordered.'],
  ['2024-11-06','P-00068','Research Domenico Rossi barbershop in Providence','Providence City Directories 1935-1978','Resolved','Existing Fact Confirmed','SRC-00049','Strong','No','Domenico Rossi listed as barber in Federal Hill area in 1940, 1950, and 1960 Providence City Directories. Occupation confirmed.'],
  ['2024-11-07','P-00022','Search for Daniel Shaughnessy siblings and parents','Cloyne Diocese Archives, Cork Archives','Researching','Follow-Up Needed','SRC-00022','Moderate','Yes','Shaughnessy family in Cloyne register but parents\' names not yet read clearly. Further research needed at Cork Archives.'],
  ['2024-11-07','P-00073','Verify William James O\'Brien marriage record','Providence City Clerk, RI Vital Records','Paused','No Result','','Confirmed','Yes','No marriage record found for William James O\'Brien. He may have remained unmarried. Research paused pending further leads.'],
  ['2024-11-08','P-00074','Research Kathleen O\'Brien Carey\'s married life','Providence RI Vital Records, death cert 2005','Resolved','Existing Fact Confirmed','SRC-00054','Confirmed','No','Kathleen Agnes O\'Brien Carey death cert obtained. Lived to 93. Married name Carey. No further research planned.'],
  ['2024-11-08','P-00060','Await DNA results for Flynn-O\'Brien case','AncestryDNA results portal','Researching','Follow-Up Needed','','Conflicting','Yes','DNA kit in transit. Results expected 6-8 weeks. Case on hold pending DNA. Conflicting parish records remain unresolved.'],
];

(async () => {
  const reqs = [];
  const vals = [];

  const deepGreen = '#5A7E5A';

  // Tab setup
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, tabColor: hex(C.confirmed), gridProperties: { frozenRowCount: 7, frozenColumnCount: 2 } },
    fields: 'tabColor,gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
  }});

  // Row 1 — Title
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 0, 1, 0, 12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(deepGreen),
      textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  vals.push({ range: `'${S}'!A1`, values: [['RESEARCH LOG & TASK TRACKER']] });

  // Row 2 — Subtitle
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 1, 2, 0, 12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.confirmed),
      textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  vals.push({ range: `'${S}'!A2`, values: [['Track every research task: what you searched, what you found, and what to do next. Keep your research organized and never duplicate effort.']] });

  // Row 3-4 — Stats
  const statLabels = ['Total Entries','Resolved','Brick Walls','In Progress','Follow-Up Needed','Conflicting'];
  const statCols = [0,2,4,6,8,10];
  statLabels.forEach((lbl, i) => {
    const c = statCols[i];
    reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, c, c+2), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, c, c+2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, 2, 3, c, c+2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(deepGreen),
        textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(SID, 3, 4, c, c+2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.input),
        textFormat: { bold: true, fontSize: 11, foregroundColor: hex(deepGreen), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});
    vals.push({ range: `'${S}'!${colL(c)}3`, values: [[lbl]] });
  });
  const statFormulas = [
    `=COUNTA($B$8:$B$5007)`,
    `=COUNTIF($G$8:$G$5007,"Resolved")`,
    `=COUNTIF($G$8:$G$5007,"Brick Wall")`,
    `=COUNTIFS($G$8:$G$5007,"Researching")+COUNTIFS($G$8:$G$5007,"Evidence Found")+COUNTIFS($G$8:$G$5007,"Lead Found")`,
    `=COUNTIF($K$8:$K$5007,"Yes")`,
    `=COUNTIF($G$8:$G$5007,"Conflicting Evidence")`,
  ];
  statFormulas.forEach((f, i) => {
    vals.push({ range: `'${S}'!${colL(statCols[i])}4`, values: [[f]] });
  });

  // Row 5-6 — spacer
  reqs.push({ repeatCell: {
    range: gridRange(SID, 4, 6, 0, 12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat(backgroundColor)',
  }});

  // Row 7 — Headers
  reqs.push({ repeatCell: {
    range: gridRange(SID, 6, 7, 0, 12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(deepGreen),
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
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 5007 },
    properties: { pixelSize: 44 }, fields: 'pixelSize',
  }});

  // Column widths: A=80, B=100, C=90, D=180, E=220, F=200, G=120, H=160, I=100, J=120, K=80, L=300
  const colWidths = [80,100,90,180,220,200,120,160,100,120,80,300];
  colWidths.forEach((px, ci) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  // Wrap text for goal, next steps columns
  [4, 11].forEach(ci => {
    reqs.push({ repeatCell: {
      range: gridRange(SID, 7, 5007, ci, ci+1),
      cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } },
      fields: 'userEnteredFormat(wrapStrategy)',
    }});
  });

  // Border header
  reqs.push({ updateBorders: {
    range: gridRange(SID, 6, 7, 0, 12),
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.border) },
  }});

  // Conditional formatting — Research Status (col G = index 6)
  const CF_STATUS = [
    { val: 'Resolved',            bg: C.confirmed },
    { val: 'Evidence Found',      bg: C.secondary },
    { val: 'Lead Found',          bg: C.info },
    { val: 'Researching',         bg: C.info },
    { val: 'Needs Corroboration', bg: C.review },
    { val: 'Conflicting Evidence',bg: C.conflict },
    { val: 'Brick Wall',          bg: '#CFD8DC' },
    { val: 'Not Started',         bg: C.neutral },
    { val: 'Paused',              bg: C.neutral },
    { val: 'Archived',            bg: C.neutral },
  ];
  CF_STATUS.forEach(({ val, bg }) => {
    reqs.push({ addConditionalFormatRule: { rule: {
      ranges: [gridRange(SID, 7, 5007, 6, 7)],
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: val }] },
        format: { backgroundColor: hex(bg) },
      },
    }, index: 0 }});
  });

  // Conditional formatting — Research Outcome (col H = index 7)
  const CF_OUTCOME = [
    { val: 'Existing Fact Confirmed', bg: C.confirmed },
    { val: 'Source Found',            bg: C.secondary },
    { val: 'New Person Found',        bg: '#C8E6C9' },
    { val: 'Follow-Up Needed',        bg: C.review },
    { val: 'Existing Fact Challenged',bg: C.conflict },
    { val: 'No Result',               bg: C.neutral },
  ];
  CF_OUTCOME.forEach(({ val, bg }) => {
    reqs.push({ addConditionalFormatRule: { rule: {
      ranges: [gridRange(SID, 7, 5007, 7, 8)],
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: val }] },
        format: { backgroundColor: hex(bg) },
      },
    }, index: 0 }});
  });

  // Conditional formatting — Follow-Up Required (col K = index 10) — YES highlighted
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, 7, 5007, 10, 11)],
    booleanRule: {
      condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Yes' }] },
      format: { backgroundColor: hex(C.review) },
    },
  }, index: 0 }});

  // Data validation
  // G = Research Status (col 6) — Reference Data!$M$2:$M$11
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 5007, 6, 7),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$M$2:$M$11` }] },
      strict: false, showCustomUi: true,
    },
  }});
  // H = Research Outcome (col 7) — Reference Data!$N$2:$N$11
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 5007, 7, 8),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$N$2:$N$11` }] },
      strict: false, showCustomUi: true,
    },
  }});
  // J = Evidence Confidence (col 9) — Reference Data!$J$2:$J$7
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 5007, 9, 10),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$J$2:$J$7` }] },
      strict: false, showCustomUi: true,
    },
  }});
  // K = Follow-Up Required (col 10) — Yes/No dropdown
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 5007, 10, 11),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$P$2:$P$3` }] },
      strict: false, showCustomUi: true,
    },
  }});

  // Alternating row fill
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, 7, 5007, 0, 12)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND(MOD(ROW(),2)=0,LEN($B8)>0)' }] },
      format: { backgroundColor: hex(C.altRow) },
    },
  }, index: 0 }});

  // Build data rows
  const dataRows = LOG.map(([date, pid, goal, repo, status, outcome, srcFound, conf, followUp, nextSteps], i) => {
    const r = i + 8;
    return [
      `=IF(B${r}="","","LOG-"&TEXT(ROW()-7,"00000"))`,
      date,
      pid,
      `=IFERROR(VLOOKUP(C${r},'Master People'!$A$8:$B$5007,2,FALSE),"")`,
      goal,
      repo,
      status,
      outcome,
      srcFound,
      conf,
      followUp,
      nextSteps,
    ];
  });

  await batchUpdate(id, reqs, 'log-fmt');
  await valuesBatchUpdate(id, vals, 'log-vals');
  await valuesBatchUpdate(id, [{ range: `'${S}'!A8`, values: dataRows }], 'log-data');
  console.log(`✓ Research Log — ${LOG.length} rows written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
