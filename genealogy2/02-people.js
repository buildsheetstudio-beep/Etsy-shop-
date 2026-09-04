'use strict';
const { sheets, batchUpdate, valuesBatchUpdate, gridRange, hex, colL, C } = require('./lib');
const { id, sheetMap } = JSON.parse(require('fs').readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Master People'];
const S = 'Master People';

const HEADERS = [
  'Person ID','Preferred Display Name','Given Name','Middle Name','Maiden / Birth Surname',
  'Current / Married Surname','Alternate Name','Prefix','Suffix','Sex / Gender Record',
  'Living Status','Birth Date','Birth Date Confidence','Birth Place','Death Date',
  'Death Date Confidence','Death Place','Burial Place','Primary Occupation',
  'Primary Language','Religion / Affiliation','Nationality / Historical Identity Note',
  'Photo / Document URL','Primary Source ID','Overall Evidence Confidence',
  'Generation Label','Family Branch','Date Added','Last Updated','Research Status',
  'Biography / Notes'
]; // A-AE (31 cols)

// Fictional people — 4+ generations, multiple branches
// Col order: [given, middle, birthSurname, marriedSurname, altName, prefix, suffix, gender, living,
//             birthDate, birthConf, birthPlace, deathDate, deathConf, deathPlace, burialPlace,
//             occupation, language, religion, nationality, photoURL, srcID, confidence, generation, branch, dateAdded, lastUpdated, resStatus, notes]
const PEOPLE = [
  // === HARTWELL BRANCH — 4 generations ===
  // Gen 0 — Root
  ['Emma','Rose','Hartwell','Hartwell','','','','Female','Living','1988-03-14','Exact','Boston, MA, USA','','','','','Software Engineer','English','','','','SRC-00001','Confirmed','Self / Root','Hartwell','2024-01-10','2024-11-15','Researching','Root person of this family tree.'],
  // Gen 1 — Parents
  ['James','Edward','Hartwell','Hartwell','','Mr.','','','Deceased','1955-06-22','Exact','Worcester, MA, USA','2019-08-05','Exact','Boston, MA, USA','Forest Hills Cemetery, MA','Accountant','English','','','','SRC-00002','Confirmed','Parent','Hartwell','2024-01-10','2024-09-20','Resolved','Emma\'s father.'],
  ['Margaret','Anne','O\'Brien','Hartwell','Maggie','Mrs.','','Female','Living','1958-11-03','Exact','Providence, RI, USA','','','','','School Teacher','English','Catholic','Irish-American','','SRC-00003','Confirmed','Parent','O\'Brien','2024-01-10','2024-11-01','Researching','Emma\'s mother. Maiden name O\'Brien.'],
  // Gen 2 — Hartwell grandparents
  ['Robert','Charles','Hartwell','Hartwell','','Mr.','','','Deceased','1922-04-10','Exact','Springfield, MA, USA','1995-12-18','Exact','Worcester, MA, USA','Pine Grove Cemetery, MA','Factory Foreman','English','','','','SRC-00004','Strong','Grandparent','Hartwell','2024-01-15','2024-08-12','Resolved','James\'s father.'],
  ['Eleanor','Mae','Chapman','Hartwell','Ellie','Mrs.','','Female','Deceased','1926-09-28','Exact','Springfield, MA, USA','2003-03-15','Exact','Worcester, MA, USA','Pine Grove Cemetery, MA','Homemaker','English','','','','SRC-00005','Confirmed','Grandparent','Hartwell','2024-01-15','2024-08-12','Resolved','James\'s mother. Maiden name Chapman.'],
  // Gen 2 — O\'Brien grandparents
  ['Patrick','Joseph','O\'Brien','O\'Brien','Paddy','Mr.','','','Deceased','1920-08-15','Exact','Providence, RI, USA','1998-05-22','Exact','Providence, RI, USA','St. Michael\'s Cemetery, RI','Dockworker','English','Catholic','Irish-American','','SRC-00006','Confirmed','Grandparent','O\'Brien','2024-01-15','2024-07-30','Resolved','Margaret\'s father.'],
  ['Catherine','Mary','Murphy','O\'Brien','Kitty','Mrs.','','Female','Deceased','1924-02-04','Exact','Providence, RI, USA','2007-10-11','Exact','Providence, RI, USA','St. Michael\'s Cemetery, RI','Seamstress','English','Catholic','Irish-American','','SRC-00007','Confirmed','Grandparent','O\'Brien','2024-01-15','2024-07-30','Resolved','Margaret\'s mother. Maiden name Murphy.'],
  // Gen 3 — Hartwell great-grandparents
  ['William','Henry','Hartwell','Hartwell','Bill','Mr.','','','Deceased','1888-07-03','Exact','Coventry, England','1952-11-30','Exact','Springfield, MA, USA','Greenwood Cemetery, MA','Mill Worker','English','Protestant','British','','SRC-00008','Confirmed','Great-Grandparent','Hartwell','2024-02-01','2024-06-15','Resolved','Robert\'s father. Immigrated from England c.1910.'],
  ['Agnes','Louise','Fletcher','Hartwell','','Mrs.','','Female','Deceased','1892-03-17','Exact','Coventry, England','1960-04-09','Exact','Springfield, MA, USA','Greenwood Cemetery, MA','Homemaker','English','Protestant','British','','SRC-00009','Confirmed','Great-Grandparent','Hartwell','2024-02-01','2024-06-15','Resolved','Robert\'s mother. Maiden name Fletcher. Immigrated with husband.'],
  ['Charles','Frederick','Chapman','Chapman','','Mr.','','','Deceased','1895-01-22','Estimated','Hartford, CT, USA','1958-08-14','Exact','Springfield, MA, USA','Greenwood Cemetery, MA','Carpenter','English','','','','SRC-00010','Strong','Great-Grandparent','Hartwell','2024-02-10','2024-05-20','Resolved','Eleanor\'s father.'],
  ['Harriet','Jane','Davies','Chapman','Hattie','Mrs.','','Female','Deceased','1898-06-11','Exact','Hartford, CT, USA','1971-01-28','Exact','Springfield, MA, USA','Greenwood Cemetery, MA','Homemaker','English','','','','SRC-00011','Confirmed','Great-Grandparent','Hartwell','2024-02-10','2024-05-20','Resolved','Eleanor\'s mother. Maiden name Davies.'],
  // Gen 3 — O\'Brien great-grandparents
  ['Michael','Patrick','O\'Brien','O\'Brien','','Mr.','','','Deceased','1882-11-05','Year Only','County Cork, Ireland','1944-07-19','Exact','Providence, RI, USA','Sacred Heart Cemetery, RI','Laborer','English','Catholic','Irish','','SRC-00012','Confirmed','Great-Grandparent','O\'Brien','2024-02-15','2024-09-05','Resolved','Patrick\'s father. Immigrated 1903.'],
  ['Brigid','Nora','Shaughnessy','O\'Brien','Biddy','Mrs.','','Female','Deceased','1886-03-20','Year Only','County Cork, Ireland','1949-09-12','Exact','Providence, RI, USA','Sacred Heart Cemetery, RI','Homemaker','English','Catholic','Irish','','SRC-00013','Confirmed','Great-Grandparent','O\'Brien','2024-02-15','2024-09-05','Resolved','Patrick\'s mother. Maiden name Shaughnessy.'],
  ['Thomas','Francis','Murphy','Murphy','Tom','Mr.','','','Deceased','1886-05-28','Year Only','County Clare, Ireland','1938-02-14','Exact','Providence, RI, USA','St. Ann\'s Cemetery, RI','Stonemason','English','Catholic','Irish','','SRC-00014','Confirmed','Great-Grandparent','O\'Brien','2024-02-20','2024-09-05','Resolved','Catherine\'s father. Immigrated 1907.'],
  ['Mary','Josephine','Gallagher','Murphy','','Mrs.','','Female','Deceased','1889-08-15','Year Only','County Clare, Ireland','1955-11-22','Exact','Providence, RI, USA','St. Ann\'s Cemetery, RI','Homemaker','English','Catholic','Irish','','SRC-00015','Confirmed','Great-Grandparent','O\'Brien','2024-02-20','2024-09-05','Resolved','Catherine\'s mother. Maiden name Gallagher.'],
  // Gen 4 — Hartwell 2x great-grandparents (partial)
  ['George','Thomas','Hartwell','Hartwell','','Mr.','','','Deceased','1850-04-12','About','Coventry, England','1920-06-30','Year Only','Coventry, England','St. Michael\'s Churchyard, Coventry','Weaver','English','Protestant','British','','SRC-00016','Moderate','2x Great-Grandparent','Hartwell','2024-03-01','2024-07-10','Researching','William\'s father. Stayed in England.'],
  ['Martha','Ellen','Holt','Hartwell','','Mrs.','','Female','Deceased','1854-09-20','About','Coventry, England','1925-03-15','Year Only','Coventry, England','St. Michael\'s Churchyard, Coventry','Weaver','English','Protestant','British','','SRC-00017','Moderate','2x Great-Grandparent','Hartwell','2024-03-01','2024-07-10','Researching','William\'s mother. Maiden name Holt.'],
  ['Edward','James','Fletcher','Fletcher','','Mr.','','','Deceased','1858-02-07','Year Only','Birmingham, England','1930-11-05','Year Only','Coventry, England','','Metalworker','English','','British','','SRC-00018','Tentative','2x Great-Grandparent','Hartwell','2024-03-05','2024-06-01','Needs Corroboration','Agnes\'s father. Record uncertain.'],
  ['Sarah','Ann','Pickles','Fletcher','','Mrs.','','Female','Deceased','1862-07-14','Year Only','Birmingham, England','1935-04-20','Year Only','Coventry, England','','Homemaker','English','','British','','SRC-00019','Tentative','2x Great-Grandparent','Hartwell','2024-03-05','2024-06-01','Needs Corroboration','Agnes\'s mother. Maiden name Pickles.'],
  // O\'Brien 2x great-grandparents (Ireland)
  ['Cornelius','','O\'Brien','O\'Brien','Con','Mr.','','','Deceased','1845-10-08','Year Only','County Cork, Ireland','1895-08-17','Year Only','County Cork, Ireland','Ballycotton Churchyard, Cork','Farmer','Irish','Catholic','Irish','','SRC-00020','Tentative','2x Great-Grandparent','O\'Brien','2024-03-10','2024-09-01','Researching','Michael\'s father.'],
  ['Honora','','Crowley','O\'Brien','Nora','Mrs.','','Female','Deceased','1848-05-25','Year Only','County Cork, Ireland','1900-12-03','Year Only','County Cork, Ireland','Ballycotton Churchyard, Cork','Homemaker','Irish','Catholic','Irish','','SRC-00021','Tentative','2x Great-Grandparent','O\'Brien','2024-03-10','2024-09-01','Researching','Michael\'s mother. Maiden name Crowley.'],
  ['Daniel','','Shaughnessy','Shaughnessy','','Mr.','','','Deceased','1850-03-15','Year Only','County Cork, Ireland','1908-07-09','Year Only','County Cork, Ireland','','Farmer','Irish','Catholic','Irish','','SRC-00022','Moderate','2x Great-Grandparent','O\'Brien','2024-03-15','2024-08-05','Researching','Brigid\'s father.'],
  ['Ellen','','Sullivan','Shaughnessy','','Mrs.','','Female','Deceased','1852-09-22','Year Only','County Cork, Ireland','1910-05-30','Year Only','County Cork, Ireland','','Homemaker','Irish','Catholic','Irish','','SRC-00023','Moderate','2x Great-Grandparent','O\'Brien','2024-03-15','2024-08-05','Researching','Brigid\'s mother. Maiden name Sullivan.'],
  // Murphy 2x great-grandparents
  ['Patrick','','Murphy','Murphy','','Mr.','','','Deceased','1852-07-04','Year Only','County Clare, Ireland','1910-03-18','Year Only','County Clare, Ireland','','Farmer','Irish','Catholic','Irish','','SRC-00023','Moderate','2x Great-Grandparent','O\'Brien','2024-03-20','2024-08-10','Researching','Thomas\'s father.'],
  ['Bridget','','Ryan','Murphy','','Mrs.','','Female','Deceased','1855-02-19','Year Only','County Clare, Ireland','1912-08-27','Year Only','County Clare, Ireland','','Homemaker','Irish','Catholic','Irish','','SRC-00023','Moderate','2x Great-Grandparent','O\'Brien','2024-03-20','2024-08-10','Researching','Thomas\'s mother. Maiden name Ryan.'],
  // === EXTENDED FAMILY — Siblings, spouses, and descendants ===
  // Emma\'s siblings
  ['Thomas','James','Hartwell','Hartwell','Tom','','','','Living','1983-09-11','Exact','Worcester, MA, USA','','','','','Financial Analyst','English','','','','SRC-00024','Confirmed','Sibling','Hartwell','2024-04-01','2024-10-15','Not Started','Emma\'s older brother.'],
  ['Sarah','Elizabeth','Hartwell','Novak','Sadie','Mrs.','','Female','Living','1990-05-07','Exact','Worcester, MA, USA','','','','','Nurse','English','','','','SRC-00025','Confirmed','Sibling','Hartwell','2024-04-01','2024-10-15','Not Started','Emma\'s younger sister. Married surname Novak.'],
  // Thomas\'s spouse and children
  ['Linda','Grace','Patel','Hartwell','','Mrs.','','Female','Living','1985-02-24','Exact','Chicago, IL, USA','','','','','Physician','English','','','','SRC-00026','Confirmed','Spouse','Hartwell','2024-04-05','2024-10-01','Not Started','Thomas Hartwell\'s spouse.'],
  ['Oliver','James','Hartwell','Hartwell','','','','','Living','2012-06-30','Exact','Chicago, IL, USA','','','','','','English','','','','','Strong','Descendant','Hartwell','2024-04-05','2024-10-01','Not Started','Thomas\'s son.'],
  ['Sophia','Grace','Hartwell','Hartwell','','','','Female','Living','2015-11-18','Exact','Chicago, IL, USA','','','','','','English','','','','','Strong','Descendant','Hartwell','2024-04-05','2024-10-01','Not Started','Thomas\'s daughter.'],
  // Sarah\'s spouse
  ['Peter','','Novak','Novak','','Mr.','','','Living','1988-08-04','Exact','Boston, MA, USA','','','','','Teacher','English','','Czech-American','','','Moderate','Spouse','Hartwell','2024-04-05','2024-10-01','Not Started','Sarah Hartwell\'s husband.'],
  // Emma\'s cousin (James\'s brother\'s family)
  ['Richard','Allen','Hartwell','Hartwell','Rich','Mr.','','','Deceased','1952-04-18','Exact','Worcester, MA, USA','2020-01-07','Exact','Worcester, MA, USA','','Hardware Store Owner','English','','','','SRC-00027','Confirmed','Other','Hartwell','2024-04-10','2024-09-15','Resolved','James Hartwell\'s older brother.'],
  ['Dorothy','Mae','Simmons','Hartwell','Dot','Mrs.','','Female','Living','1954-07-22','Exact','Worcester, MA, USA','','','','','Retired Secretary','English','','','','SRC-00028','Confirmed','Other','Hartwell','2024-04-10','2024-09-15','Not Started','Richard Hartwell\'s wife.'],
  ['Kevin','Richard','Hartwell','Hartwell','','','','','Living','1978-03-09','Exact','Worcester, MA, USA','','','','','Contractor','English','','','','','Strong','Descendant','Hartwell','2024-04-10','2024-10-20','Not Started','Richard\'s son. Emma\'s first cousin.'],
  ['Jennifer','Lynn','Hartwell','Walsh','Jen','Mrs.','','Female','Living','1980-10-25','Exact','Worcester, MA, USA','','','','','Social Worker','English','','','','','Strong','Descendant','Hartwell','2024-04-10','2024-10-20','Not Started','Richard\'s daughter. Emma\'s first cousin.'],
  // Margaret\'s siblings
  ['Sean','Patrick','O\'Brien','O\'Brien','','Mr.','','','Living','1956-03-08','Exact','Providence, RI, USA','','','','','Police Officer (Retired)','English','Catholic','Irish-American','','SRC-00029','Confirmed','Sibling','O\'Brien','2024-05-01','2024-09-10','Not Started','Margaret\'s older brother.'],
  ['Colleen','Frances','O\'Brien','Donovan','','Mrs.','','Female','Deceased','1960-07-14','Exact','Providence, RI, USA','2015-04-02','Exact','Providence, RI, USA','','Bookkeeper','English','Catholic','Irish-American','','SRC-00030','Confirmed','Sibling','O\'Brien','2024-05-01','2024-09-10','Resolved','Margaret\'s sister. Married name Donovan.'],
  // Adoptive relationship
  ['Helen','Rose','Unknown','Hartwell','','','Adopted','Female','Living','1993-08-20','Estimated','Unknown','','','','','Student','English','','','','','Unknown','Sibling','Hartwell','2024-05-10','2024-10-01','Researching','Adopted by James and Margaret Hartwell 1994. Birth family unknown.'],
  // === CHAPMAN BRANCH extended ===
  ['Arthur','Raymond','Chapman','Chapman','Art','Mr.','','','Deceased','1898-02-14','Year Only','Hartford, CT, USA','1962-06-05','Exact','Springfield, MA, USA','','Insurance Agent','English','','','','SRC-00031','Moderate','Great-Grandparent','Hartwell','2024-05-15','2024-07-20','Resolved','Charles Chapman\'s brother. Eleanor\'s uncle.'],
  ['Rose','Margaret','Kelly','Chapman','','Mrs.','','Female','Deceased','1900-05-12','Year Only','Hartford, CT, USA','1978-09-30','Exact','Springfield, MA, USA','','Homemaker','English','','Irish-American','','SRC-00032','Moderate','Other','Hartwell','2024-05-15','2024-07-20','Not Started','Arthur Chapman\'s wife.'],
  // === GALLAGHER BRANCH (Mary Murphy\'s family) ===
  ['John','Michael','Gallagher','Gallagher','','Mr.','','','Deceased','1855-01-17','Year Only','County Clare, Ireland','1920-11-08','Year Only','County Clare, Ireland','','Farmer','Irish','Catholic','Irish','','SRC-00033','Tentative','2x Great-Grandparent','O\'Brien','2024-05-20','2024-08-25','Brick Wall','Mary Gallagher\'s father. Limited records.'],
  ['Anne','','Connelly','Gallagher','Nan','Mrs.','','Female','Deceased','1858-04-09','Year Only','County Clare, Ireland','1925-02-17','Year Only','County Clare, Ireland','','Homemaker','Irish','Catholic','Irish','','SRC-00034','Tentative','2x Great-Grandparent','O\'Brien','2024-05-20','2024-08-25','Brick Wall','Mary Gallagher\'s mother. Maiden name Connelly.'],
  // === FLETCHER BRANCH extended ===
  ['Henry','William','Fletcher','Fletcher','','Mr.','','','Deceased','1860-08-21','Year Only','Birmingham, England','1928-04-14','Year Only','Coventry, England','','Blacksmith','English','Protestant','British','','SRC-00035','Tentative','2x Great-Grandparent','Hartwell','2024-06-01','2024-07-01','Needs Corroboration','Possibly Edward Fletcher\'s brother. Conflicting evidence.'],
  // === Deceased with conflicting records ===
  ['Elizabeth','Anne','Hartwell','Hartwell','Lizzie','','','Female','Deceased','1918-12-01','Exact','Springfield, MA, USA','1919-01-15','Exact','Springfield, MA, USA','Greenwood Cemetery, MA','','English','','','','SRC-00036','Confirmed','Other','Hartwell','2024-06-05','2024-08-01','Resolved','Robert Hartwell\'s infant sister. Died age 6 weeks.'],
  // === Half-sibling ===
  ['Daniel','Paul','Hartwell','Hartwell','Danny','','','','Living','1975-11-20','Exact','Hartford, CT, USA','','','','','Mechanic','English','','','','','Moderate','Sibling','Hartwell','2024-06-10','2024-10-05','Researching','James Hartwell\'s son from first marriage. Emma\'s half-brother.'],
  // James\'s first marriage
  ['Carol','Lynn','Morrison','Hartwell','','Mrs.','','Female','Living','1953-04-11','Exact','Hartford, CT, USA','','','','','Retired Nurse','English','','','','','Moderate','Other','Hartwell','2024-06-10','2024-10-05','Not Started','James Hartwell\'s first wife (divorced 1979).'],
  // === Additional Hartwell ancestors ===
  ['Frederick','George','Hartwell','Hartwell','Fred','Mr.','','','Deceased','1815-06-15','About','Warwickshire, England','1880-03-22','About','Coventry, England','','Weaver','English','Protestant','British','','SRC-00037','Moderate','3x Great-Grandparent','Hartwell','2024-06-15','2024-07-20','Researching','George Hartwell\'s father.'],
  ['Ann','Elizabeth','Turner','Hartwell','','Mrs.','','Female','Deceased','1820-09-10','About','Warwickshire, England','1885-07-14','About','Coventry, England','','Homemaker','English','Protestant','British','','SRC-00038','Moderate','3x Great-Grandparent','Hartwell','2024-06-15','2024-07-20','Researching','George Hartwell\'s mother. Maiden name Turner.'],
  // O\'Brien 3x great-grandparents
  ['Jeremiah','','O\'Brien','O\'Brien','Jerry','Mr.','','','Deceased','1808-04-20','Year Only','County Cork, Ireland','1855-09-18','Year Only','County Cork, Ireland','','Fisherman','Irish','Catholic','Irish','','SRC-00039','Tentative','3x Great-Grandparent','O\'Brien','2024-06-20','2024-09-15','Brick Wall','Cornelius O\'Brien\'s father. Famine era. Limited records.'],
  ['Mary','','Driscoll','O\'Brien','','Mrs.','','Female','Deceased','1812-02-15','Year Only','County Cork, Ireland','1848-11-30','Year Only','County Cork, Ireland','','Homemaker','Irish','Catholic','Irish','','SRC-00040','Tentative','3x Great-Grandparent','O\'Brien','2024-06-20','2024-09-15','Brick Wall','Cornelius O\'Brien\'s mother. Maiden name Driscoll. Possibly died in famine.'],
  // === Additional extended family to reach 120+ ===
  ['Grace','Ellen','Walsh','Hartwell','Gracie','Mrs.','','Female','Living','1955-03-22','Exact','Worcester, MA, USA','','','','','Retired Teacher','English','','','','','Confirmed','Other','Hartwell','2024-07-01','2024-10-10','Not Started','Jennifer Walsh\'s mother-in-law (Kevin\'s aunt by marriage).'],
  ['Brian','Francis','Walsh','Walsh','','Mr.','','','Deceased','1950-09-14','Exact','Worcester, MA, USA','2018-06-07','Exact','Worcester, MA, USA','','Plumber','English','','','','SRC-00041','Confirmed','Other','Hartwell','2024-07-01','2024-09-01','Resolved','Jennifer Walsh\'s father.'],
  ['Maureen','Claire','Donovan','Donovan','Mo','Mrs.','','Female','Living','1958-08-30','Exact','Providence, RI, USA','','','','','Florist','English','Catholic','Irish-American','','SRC-00042','Confirmed','Other','O\'Brien','2024-07-05','2024-10-05','Not Started','Colleen O\'Brien\'s daughter-in-law.'],
  ['Patrick','Sean','Donovan','Donovan','Pat','Mr.','','','Living','1957-12-01','Exact','Providence, RI, USA','','','','','Electrician','English','Catholic','Irish-American','','SRC-00043','Confirmed','Other','O\'Brien','2024-07-05','2024-10-05','Not Started','Colleen O\'Brien\'s husband (died 2015 — NOTE: Colleen died 2015 as well; this Donovan is her husband and survives).'],
  ['Fiona','Mary','O\'Brien','Donovan','','Mrs.','','Female','Living','1983-04-17','Exact','Providence, RI, USA','','','','','Graphic Designer','English','Catholic','Irish-American','','','Strong','Descendant','O\'Brien','2024-07-05','2024-10-05','Not Started','Colleen Donovan\'s daughter.'],
  ['Connor','Patrick','Donovan','Donovan','','','','','Living','1985-07-22','Exact','Providence, RI, USA','','','','','Firefighter','English','Catholic','Irish-American','','','Strong','Descendant','O\'Brien','2024-07-05','2024-10-05','Not Started','Colleen Donovan\'s son.'],
  ['Sean','Michael','O\'Brien','O\'Brien','','Mr.','','','Living','1980-05-14','Exact','Providence, RI, USA','','','','','Lawyer','English','Catholic','Irish-American','','','Strong','Descendant','O\'Brien','2024-07-10','2024-10-10','Not Started','Sean O\'Brien Jr. son of Sean Patrick O\'Brien.'],
  ['Nicole','Ann','Rossi','O\'Brien','Nicky','Mrs.','','Female','Living','1982-09-03','Exact','Providence, RI, USA','','','','','Marketing Manager','English','Catholic','Italian-American','','','Strong','Spouse','O\'Brien','2024-07-10','2024-10-10','Not Started','Sean O\'Brien Jr.\'s wife.'],
  // Additional ancestry with unknown dates
  ['Unknown','','Crowley','Crowley','','','','','Deceased','','Unknown','County Cork, Ireland','','Unknown','County Cork, Ireland','','','Irish','Catholic','Irish','','','Unknown','3x Great-Grandparent','O\'Brien','2024-07-15','2024-08-01','Brick Wall','Honora Crowley\'s father. Name and dates unknown. Brick wall.'],
  // Disputed relationship
  ['James','Aloysius','Flynn','Flynn','','Mr.','','','Deceased','1884-03-11','Year Only','County Cork, Ireland','1950-08-20','Year Only','Providence, RI, USA','','Laborer','English','Catholic','Irish','','SRC-00044','Conflicting','Other','O\'Brien','2024-07-20','2024-10-01','Conflicting Evidence','Possibly a half-sibling of Michael O\'Brien. Evidence disputed — two conflicting parish records.'],
  // O\'Brien 4x great-grandparent attempt
  ['Thaddeus','','O\'Brien','O\'Brien','','Mr.','','','Deceased','1775-01-01','Estimated','County Cork, Ireland','1840-01-01','Estimated','County Cork, Ireland','','Farmer','Irish','Catholic','Irish','','SRC-00045','Tentative','4x Great-Grandparent','O\'Brien','2024-07-25','2024-09-20','Brick Wall','Jeremiah O\'Brien\'s probable father. Dates estimated. Pre-civil registration.'],
  ['Catherine','','McCarthy','O\'Brien','','Mrs.','','Female','Deceased','1778-01-01','Estimated','County Cork, Ireland','1842-01-01','Estimated','County Cork, Ireland','','Homemaker','Irish','Catholic','Irish','','SRC-00045','Tentative','4x Great-Grandparent','O\'Brien','2024-07-25','2024-09-20','Brick Wall','Jeremiah O\'Brien\'s probable mother. Maiden name McCarthy.'],
  // Hartwell 4x great-grandparent attempt
  ['Thomas','','Hartwell','Hartwell','','Mr.','','','Deceased','1780-01-01','Estimated','Warwickshire, England','1850-01-01','Estimated','Warwickshire, England','','Weaver','English','Protestant','British','','SRC-00046','Tentative','4x Great-Grandparent','Hartwell','2024-08-01','2024-09-10','Researching','Frederick Hartwell\'s probable father.'],
  ['Jane','','Booth','Hartwell','','Mrs.','','Female','Deceased','1785-01-01','Estimated','Warwickshire, England','1855-01-01','Estimated','Warwickshire, England','','Homemaker','English','Protestant','British','','SRC-00046','Tentative','4x Great-Grandparent','Hartwell','2024-08-01','2024-09-10','Researching','Frederick Hartwell\'s probable mother. Maiden name Booth.'],
  // No-source person
  ['Francis','Michael','Holt','Holt','Frank','Mr.','','','Deceased','1820-01-01','Estimated','Coventry, England','1890-01-01','Estimated','Coventry, England','','Weaver','English','Protestant','British','','','Unknown','3x Great-Grandparent','Hartwell','2024-08-05','2024-08-05','Not Started','Martha Holt\'s father. No source records found.'],
  // Newly discovered person
  ['Nora','','O\'Connell','Crowley','','Mrs.','','Female','Deceased','1825-06-01','Year Only','County Cork, Ireland','1890-03-18','Year Only','County Cork, Ireland','','Homemaker','Irish','Catholic','Irish','','SRC-00047','Moderate','3x Great-Grandparent','O\'Brien','2024-08-10','2024-10-20','Evidence Found','Honora Crowley\'s mother. Maiden name O\'Connell. Found in 1851 Griffith\'s Valuation.'],
  // Milestone person with multiple sources
  ['Alice','Victoria','Holt','Fletcher','','Mrs.','','Female','Deceased','1825-03-08','Year Only','Birmingham, England','1895-10-21','Year Only','Coventry, England','','Lacemaker','English','Protestant','British','','SRC-00048','Strong','3x Great-Grandparent','Hartwell','2024-08-15','2024-10-15','Resolved','Agnes Fletcher\'s grandmother. Multiple sources confirmed.'],
  // Person with alternate name
  ['Domenico','','Rossi','Rossi','','Mr.','','','Deceased','1910-05-14','Exact','Palermo, Sicily, Italy','1978-03-25','Exact','Providence, RI, USA','','Barber','Italian','Catholic','Italian-American','','SRC-00049','Strong','Other','O\'Brien','2024-08-20','2024-10-10','Resolved','Nicole Rossi\'s grandfather. Immigrated 1935.'],
  ['Rosa','Carmela','Marino','Rossi','','Mrs.','','Female','Deceased','1915-09-01','Exact','Palermo, Sicily, Italy','1990-12-07','Exact','Providence, RI, USA','','Homemaker','Italian','Catholic','Italian-American','','SRC-00050','Strong','Other','O\'Brien','2024-08-20','2024-10-10','Resolved','Nicole Rossi\'s grandmother. Maiden name Marino.'],
  // Overdue follow-up person
  ['Bartholomew','','Chapman','Chapman','','Mr.','','','Deceased','1865-09-01','Year Only','Hartford, CT, USA','1940-07-15','Year Only','Hartford, CT, USA','','Merchant','English','','','','','Tentative','Other','Hartwell','2024-09-01','2024-09-01','Lead Found','Possible brother of Charles Chapman. Follow-up research needed.'],
  // Civil registration era
  ['Joseph','Bernard','O\'Brien','O\'Brien','','Mr.','','','Deceased','1905-04-25','Exact','Providence, RI, USA','1945-06-06','Exact','Normandy, France','Normandy American Cemetery, France','US Army Soldier','English','Catholic','American','','SRC-00051','Confirmed','Other','O\'Brien','2024-09-05','2024-10-25','Resolved','Michael O\'Brien\'s second son. KIA WWII D-Day.'],
  // Person with no death record
  ['Margaret','Alice','O\'Brien','Unknown','Meg','','','Female','Unknown','1910-12-30','Exact','Providence, RI, USA','','Unknown','','','Unknown','English','Catholic','Irish-American','','SRC-00052','Moderate','Other','O\'Brien','2024-09-10','2024-10-15','Brick Wall','Michael O\'Brien\'s daughter. No record after 1930 census. Brick wall.'],
  // Additional people for count
  ['William','James','O\'Brien','O\'Brien','Willie','Mr.','','','Deceased','1908-07-19','Exact','Providence, RI, USA','1980-04-10','Exact','Providence, RI, USA','','Textile Worker','English','Catholic','Irish-American','','SRC-00053','Confirmed','Other','O\'Brien','2024-09-15','2024-09-15','Resolved','Michael O\'Brien\'s son. Patrick\'s older brother.'],
  ['Kathleen','Agnes','O\'Brien','Carey','Kay','Mrs.','','Female','Deceased','1912-03-17','Exact','Providence, RI, USA','2005-11-22','Exact','Providence, RI, USA','','Homemaker','English','Catholic','Irish-American','','SRC-00054','Confirmed','Other','O\'Brien','2024-09-15','2024-09-15','Resolved','Michael O\'Brien\'s daughter. Married name Carey.'],
  ['Robert','Francis','Hartwell','Hartwell','Bobby','Mr.','','','Deceased','1924-08-15','Exact','Springfield, MA, USA','1943-11-14','Exact','Tarawa, Pacific','','US Marine','English','','','','SRC-00055','Confirmed','Other','Hartwell','2024-09-20','2024-10-01','Resolved','Robert Hartwell Sr.\'s younger brother. KIA WWII.'],
  ['Mildred','Jean','Hartwell','Perry','Millie','Mrs.','','Female','Deceased','1928-06-02','Exact','Springfield, MA, USA','2010-08-19','Exact','Springfield, MA, USA','','Secretary','English','','','','SRC-00056','Confirmed','Other','Hartwell','2024-09-20','2024-10-01','Resolved','Robert Hartwell Sr.\'s younger sister. Married name Perry.'],
  ['Howard','Earl','Perry','Perry','','Mr.','','','Deceased','1925-04-14','Exact','Springfield, MA, USA','1998-02-28','Exact','Springfield, MA, USA','','Factory Worker','English','','','','SRC-00057','Confirmed','Other','Hartwell','2024-09-20','2024-10-01','Resolved','Mildred Hartwell\'s husband.'],
  ['Barbara','Lee','Perry','Morrison','Barb','Mrs.','','Female','Living','1952-10-07','Exact','Springfield, MA, USA','','','','','Retired Librarian','English','','','','','Strong','Other','Hartwell','2024-09-25','2024-10-10','Not Started','Mildred Perry\'s daughter. Married name Morrison.'],
  ['Harold','Edward','Perry','Perry','Hal','Mr.','','','Living','1955-03-29','Exact','Springfield, MA, USA','','','','','Retired Engineer','English','','','','','Strong','Other','Hartwell','2024-09-25','2024-10-10','Not Started','Mildred Perry\'s son.'],
  // More extended O\'Brien family
  ['Timothy','Joseph','Murphy','Murphy','Tim','Mr.','','','Deceased','1918-09-12','Exact','Providence, RI, USA','1988-07-04','Exact','Providence, RI, USA','','Postal Worker','English','Catholic','Irish-American','','SRC-00058','Confirmed','Other','O\'Brien','2024-10-01','2024-10-15','Resolved','Thomas Murphy\'s son. Catherine Murphy\'s brother.'],
  ['Frances','Rose','Murphy','Kelly','Fran','Mrs.','','Female','Deceased','1920-05-27','Exact','Providence, RI, USA','2001-10-30','Exact','Providence, RI, USA','','Homemaker','English','Catholic','Irish-American','','SRC-00059','Confirmed','Other','O\'Brien','2024-10-01','2024-10-15','Resolved','Thomas Murphy\'s daughter. Catherine Murphy\'s sister. Married name Kelly.'],
  ['Gerard','Thomas','Kelly','Kelly','Gerry','Mr.','','','Deceased','1918-11-05','Exact','Providence, RI, USA','1975-04-18','Exact','Providence, RI, USA','','Dockworker','English','Catholic','Irish-American','','SRC-00060','Confirmed','Other','O\'Brien','2024-10-01','2024-10-15','Resolved','Frances Murphy\'s husband.'],
  // Step-family
  ['Christopher','Allen','West','West','Chris','Mr.','','','Living','1960-01-30','Exact','Boston, MA, USA','','','','','IT Manager','English','','','','','Moderate','Other','Hartwell','2024-10-05','2024-10-20','Not Started','Margaret Hartwell\'s partner since 2021. Not legally married.'],
  // Additional gen-4 ancestors
  ['Hannah','','Turner','Hartwell','','Mrs.','','Female','Deceased','1790-03-01','Estimated','Warwickshire, England','1860-01-01','Estimated','Warwickshire, England','','Homemaker','English','Protestant','British','','SRC-00061','Tentative','4x Great-Grandparent','Hartwell','2024-10-10','2024-10-20','Researching','Ann Turner\'s mother (Thomas Hartwell\'s mother-in-law).'],
  ['John','','Turner','Turner','','Mr.','','','Deceased','1785-01-01','Estimated','Warwickshire, England','1855-01-01','Estimated','Warwickshire, England','','Farmer','English','Protestant','British','','SRC-00061','Tentative','4x Great-Grandparent','Hartwell','2024-10-10','2024-10-20','Researching','Ann Turner\'s father.'],
  // Recently discovered immigrants
  ['Salvatore','','Marino','Marino','Sal','Mr.','','','Deceased','1885-11-20','Year Only','Palermo, Sicily, Italy','1960-06-12','Exact','Providence, RI, USA','','Fisherman','Italian','Catholic','Italian-American','','SRC-00062','Moderate','Other','O\'Brien','2024-10-15','2024-10-15','Evidence Found','Rosa Marino\'s father. Immigrated 1912.'],
  ['Giuseppa','','Ferrara','Marino','','Mrs.','','Female','Deceased','1890-08-14','Year Only','Palermo, Sicily, Italy','1965-03-28','Exact','Providence, RI, USA','','Homemaker','Italian','Catholic','Italian-American','','SRC-00063','Moderate','Other','O\'Brien','2024-10-15','2024-10-15','Evidence Found','Rosa Marino\'s mother. Maiden name Ferrara.'],
  // Person with living status unknown
  ['Andrew','James','Gallagher','Gallagher','Andy','Mr.','','','Unknown','1879-10-18','Year Only','County Clare, Ireland','','Unknown','','','','Irish','Catholic','Irish','','','Tentative','Other','O\'Brien','2024-10-18','2024-10-18','Researching','John Gallagher\'s son. Emigrated c.1900. No further record. Status unknown.'],
  // Brick wall example
  ['Sarah','','Unknown','Unknown','','','','Female','Deceased','','Unknown','Unknown','','Unknown','','','','','','','','','Unknown','Other','Hartwell','2024-10-20','2024-10-20','Brick Wall','William Hartwell\'s possible sister. Referenced in a letter. No other records.'],
  // Person with no source
  ['Albert','Frederick','Holt','Holt','','Mr.','','','Deceased','1855-01-01','Estimated','Coventry, England','1920-01-01','Estimated','Coventry, England','','','English','','British','','','Unknown','3x Great-Grandparent','Hartwell','2024-10-22','2024-10-22','Not Started','Francis Holt\'s son. No source records.'],
  // Recent addition with strong evidence
  ['Alice','','Booth','Booth','','Mrs.','','Female','Deceased','1810-07-14','Year Only','Warwickshire, England','1878-09-22','Year Only','Warwickshire, England','','Lacemaker','English','Protestant','British','','SRC-00064','Moderate','4x Great-Grandparent','Hartwell','2024-10-23','2024-10-23','Evidence Found','Jane Booth\'s mother. Thomas Hartwell\'s mother-in-law.'],
  // TOTALS: we need to reach 120+ — adding more to be sure
  ['Declan','Patrick','O\'Brien','O\'Brien','Dec','Mr.','','','Living','1982-07-30','Exact','Providence, RI, USA','','','','','Police Officer','English','Catholic','Irish-American','','','Strong','Descendant','O\'Brien','2024-10-24','2024-10-24','Not Started','Sean O\'Brien\'s son.'],
  ['Aisling','Marie','Murphy','O\'Brien','Ash','Mrs.','','Female','Living','1984-02-14','Exact','Providence, RI, USA','','','','','Nurse','English','Catholic','Irish-American','','','Strong','Spouse','O\'Brien','2024-10-24','2024-10-24','Not Started','Declan O\'Brien\'s wife.'],
  ['Liam','Connor','O\'Brien','O\'Brien','','','','','Living','2010-04-22','Exact','Providence, RI, USA','','','','','','English','','','','','Confirmed','Descendant','O\'Brien','2024-10-24','2024-10-24','Not Started','Declan O\'Brien\'s son.'],
  ['Siobhan','Grace','O\'Brien','O\'Brien','','','','Female','Living','2013-09-15','Exact','Providence, RI, USA','','','','','','English','','','','','Confirmed','Descendant','O\'Brien','2024-10-24','2024-10-24','Not Started','Declan O\'Brien\'s daughter.'],
  ['Martin','Joseph','O\'Brien','O\'Brien','Marty','Mr.','','','Deceased','1880-03-28','Year Only','County Cork, Ireland','1960-11-04','Exact','Providence, RI, USA','','Ironworker','English','Catholic','Irish-American','','SRC-00065','Confirmed','Other','O\'Brien','2024-10-25','2024-10-25','Resolved','Michael O\'Brien\'s brother. Came with him to Providence.'],
  ['Anastasia','Bridget','Shea','O\'Brien','Stacy','Mrs.','','Female','Deceased','1883-05-16','Year Only','Providence, RI, USA','1968-08-23','Exact','Providence, RI, USA','','Homemaker','English','Catholic','Irish-American','','SRC-00066','Confirmed','Other','O\'Brien','2024-10-25','2024-10-25','Resolved','Martin O\'Brien\'s wife.'],
  ['Patrick','Martin','O\'Brien','O\'Brien','Marty Jr.','Mr.','','','Deceased','1912-08-10','Exact','Providence, RI, USA','1985-01-14','Exact','Providence, RI, USA','','Steel Worker','English','Catholic','Irish-American','','SRC-00067','Confirmed','Descendant','O\'Brien','2024-10-25','2024-10-25','Resolved','Martin O\'Brien\'s son.'],
  ['Evelyn','Grace','McCarthy','O\'Brien','Eve','Mrs.','','Female','Deceased','1915-04-07','Exact','Providence, RI, USA','1992-06-18','Exact','Providence, RI, USA','','Bookkeeper','English','Catholic','Irish-American','','SRC-00068','Confirmed','Descendant','O\'Brien','2024-10-25','2024-10-25','Resolved','Patrick Martin O\'Brien\'s wife.'],
  ['James','Patrick','O\'Brien','O\'Brien','Jimmy','Mr.','','','Living','1940-02-22','Exact','Providence, RI, USA','','','','','Retired Fire Chief','English','Catholic','Irish-American','','SRC-00069','Strong','Other','O\'Brien','2024-10-26','2024-10-26','Not Started','Patrick Martin O\'Brien\'s son.'],
  ['Dorothy','Ann','Walsh','O\'Brien','Dot','Mrs.','','Female','Living','1943-07-18','Exact','Providence, RI, USA','','','','','Retired Teacher','English','Catholic','Irish-American','','SRC-00070','Strong','Other','O\'Brien','2024-10-26','2024-10-26','Not Started','James Patrick O\'Brien\'s wife.'],
  // Additional distant Hartwell relatives
  ['Cecil','Edwin','Hartwell','Hartwell','','Mr.','','','Deceased','1854-05-05','Year Only','Coventry, England','1932-01-20','Year Only','Coventry, England','','Factory Worker','English','Protestant','British','','SRC-00071','Tentative','3x Great-Grandparent','Hartwell','2024-10-26','2024-10-26','Researching','George Hartwell\'s brother. No emigration records found.'],
  // Person with alternate/maiden name and remarriage
  ['Eleanor','Faye','Parker','Hartwell-Brooks','Ellie','Mrs.','','Female','Living','1965-12-09','Exact','Worcester, MA, USA','','','','','Accountant','English','','','','','Moderate','Other','Hartwell','2024-10-27','2024-10-27','Not Started','Kevin Hartwell\'s former girlfriend. Mother of a child with unclear paternity — excluded from tree pending research.'],
];

(async () => {
  const reqs = [];
  const vals = [];

  // Tab setup
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, tabColor: hex(C.primary), gridProperties: { frozenRowCount: 7, columnCount: 31 } },
    fields: 'tabColor,gridProperties.frozenRowCount,gridProperties.columnCount',
  }});

  // Row 1 — Title
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 31), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 0, 1, 0, 31),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryDeep),
      textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  vals.push({ range: `'${S}'!A1`, values: [['MASTER PEOPLE / ANCESTOR DATABASE']] });

  // Row 2 — Subtitle
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 31), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 1, 2, 0, 31),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  vals.push({ range: `'${S}'!A2`, values: [['Single source of truth for all persons. Enter each person once. All other tabs link by Person ID.']] });

  // Rows 3-6 — Summary stats (2-column pairs)
  const STATS = [
    ['Total People', `=COUNTA('${S}'!$A$8:$A$5007)`],
    ['Living People', `=COUNTIF('${S}'!$K$8:$K$5007,"Living")`],
    ['Deceased People', `=COUNTIF('${S}'!$K$8:$K$5007,"Deceased")`],
    ['Confirmed People', `=COUNTIF('${S}'!$Y$8:$Y$5007,"Confirmed")`],
    ['Needs Research', `=SUMPRODUCT(('${S}'!$AD$8:$AD$5007<>"")*(('${S}'!$AD$8:$AD$5007="Not Started")+('${S}'!$AD$8:$AD$5007="Researching")+('${S}'!$AD$8:$AD$5007="Brick Wall")>0))`],
    ['Earliest Known Birth', `=IFERROR(TEXT(MIN(IF('${S}'!$L$8:$L$5007<>"",DATEVALUE('${S}'!$L$8:$L$5007))),"yyyy"),"—")`],
    ['Latest Person Added', `=IFERROR(INDEX('${S}'!$B$8:$B$5007,MATCH(MAX('${S}'!$AB$8:$AB$5007),'${S}'!$AB$8:$AB$5007,0)),"—")`],
    ['Generations Tracked', `=IFERROR(SUMPRODUCT((1/COUNTIF(IF('${S}'!$Z$8:$Z$5007<>"",'${S}'!$Z$8:$Z$5007,"~"),$IF('${S}'!$Z$8:$Z$5007<>"",'${S}'!$Z$8:$Z$5007,"~")))*('${S}'!$Z$8:$Z$5007<>"")),"—")`],
  ];

  const statCols = [0,2,4,6]; // A,C,E,G
  STATS.forEach(([lbl, frm], i) => {
    const row = Math.floor(i/4) + 2; // rows 3-4 (0-indexed: 2-3)
    const colPair = i % 4;
    const labelCol = statCols[colPair];
    const valCol = statCols[colPair]+1;
    vals.push({ range: `'${S}'!${colL(labelCol)}${row+1}`, values: [[lbl]] });
    vals.push({ range: `'${S}'!${colL(valCol)}${row+1}`, values: [[frm]] });

    reqs.push({ repeatCell: {
      range: gridRange(SID, row, row+1, labelCol, labelCol+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.bg),
        textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryDeep), fontFamily: 'Arial' },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(SID, row, row+1, valCol, valCol+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.formula),
        textFormat: { fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    }});
  });

  // Row 7 — Column headers
  reqs.push({ repeatCell: {
    range: gridRange(SID, 6, 7, 0, HEADERS.length),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryDeep),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
  }});
  vals.push({ range: `'${S}'!A7`, values: [HEADERS] });

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 7 },
    properties: { pixelSize: 22 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 5007 },
    properties: { pixelSize: 22 }, fields: 'pixelSize',
  }});

  // Column widths
  const colWidths = [90,160,120,100,140,140,120,60,60,90,80,120,100,160,120,100,160,160,120,80,100,140,160,90,100,130,120,100,100,120,200];
  colWidths.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Alt row shading + base fill for data rows
  reqs.push({ repeatCell: {
    range: gridRange(SID, 7, 5007, 0, HEADERS.length),
    cell: { userEnteredFormat: { backgroundColor: hex(C.white) } },
    fields: 'userEnteredFormat(backgroundColor)',
  }});

  // Date format for date columns (L, O, AB, AC = indices 11,14,27,28)
  [11,14,27,28].forEach(ci => {
    reqs.push({ repeatCell: {
      range: gridRange(SID, 7, 5007, ci, ci+1),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } },
      fields: 'userEnteredFormat(numberFormat)',
    }});
  });

  // Person ID column — formula column styling
  reqs.push({ repeatCell: {
    range: gridRange(SID, 7, 5007, 0, 1),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula),
      textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.primaryDeep) },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});

  // Conditional formatting — Evidence Confidence
  const cfConf = [
    ['Confirmed', C.confirmed], ['Strong', C.primary], ['Moderate', C.wheat],
    ['Tentative', C.lavender], ['Conflicting', C.conflict], ['Unknown', C.neutral],
  ];
  cfConf.forEach(([val, color]) => {
    reqs.push({ addConditionalFormatRule: { rule: {
      ranges: [gridRange(SID, 7, 5007, 24, 25)], // Y column
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: val }] },
        format: { backgroundColor: hex(color) },
      },
    }, index: 0 }});
  });

  // Research Status CF
  const cfRS = [
    ['Resolved', C.confirmed], ['Evidence Found', C.secondary],
    ['Researching', C.info], ['Lead Found', C.aqua],
    ['Needs Corroboration', C.wheat], ['Conflicting Evidence', C.blush],
    ['Brick Wall', C.lavender], ['Paused', C.neutral],
  ];
  cfRS.forEach(([val, color]) => {
    reqs.push({ addConditionalFormatRule: { rule: {
      ranges: [gridRange(SID, 7, 5007, 29, 30)], // AD column
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: val }] },
        format: { backgroundColor: hex(color) },
      },
    }, index: 0 }});
  });

  // Dropdowns
  const dvMap = [
    [9,  'Sex / Gender Record'], [10, 'Living Status'], [12, 'Date Precision'],
    [15, 'Date Precision'], [24, 'Evidence Confidence'], [25, 'Generation Labels'],
    [29, 'Research Status'],
  ];
  dvMap.forEach(([ci, listName]) => {
    const opts = {
      'Sex / Gender Record': ['Female','Male','Nonbinary / Other','Unknown','Not Recorded'],
      'Living Status': ['Living','Deceased','Unknown'],
      'Date Precision': ['Exact','Month / Year','Year Only','Before','After','About','Estimated','Unknown'],
      'Evidence Confidence': ['Confirmed','Strong','Moderate','Tentative','Conflicting','Unknown'],
      'Generation Labels': ['Self / Root','Parent','Grandparent','Great-Grandparent','2x Great-Grandparent','3x Great-Grandparent','4x Great-Grandparent','Descendant','Other'],
      'Research Status': ['Not Started','Researching','Lead Found','Evidence Found','Needs Corroboration','Conflicting Evidence','Resolved','Brick Wall','Paused','Archived'],
    }[listName];
    reqs.push({ setDataValidation: {
      range: gridRange(SID, 7, 5007, ci, ci+1),
      rule: { condition: { type: 'ONE_OF_LIST', values: opts.map(v => ({ userEnteredValue: v })) }, strict: false, showCustomUi: true },
    }});
  });

  // Build data rows — Person ID is =IF(B8="","","P-"&TEXT(ROW()-7,"00000"))
  const dataRows = PEOPLE.map((p, i) => {
    const row = i + 8;
    const [given, middle, birthSurname, marriedSurname, altName, prefix, suffix, gender, living,
           birthDate, birthConf, birthPlace, deathDate, deathConf, deathPlace, burialPlace,
           occupation, language, religion, nationality, photoURL, srcID, confidence, generation, branch,
           dateAdded, lastUpdated, resStatus, notes] = p;
    const idFormula = `=IF(B${row}="","","P-"&TEXT(ROW()-7,"00000"))`;
    const displayNameFormula = `=IF(B${row}<>"",B${row},TRIM(C${row}&" "&D${row}&" "&IF(F${row}<>"",F${row},E${row})))`;
    return [
      idFormula, '', given, middle, birthSurname, marriedSurname, altName, prefix, suffix,
      gender, living, birthDate, birthConf, birthPlace, deathDate, deathConf, deathPlace,
      burialPlace, occupation, language, religion, nationality, photoURL, srcID, confidence,
      generation, branch, dateAdded, lastUpdated, resStatus, notes,
    ];
  });

  // Set preferred display names (col B) for each person
  PEOPLE.forEach((p, i) => {
    const row = i + 8;
    const [given, , birthSurname, marriedSurname] = p;
    const displayName = marriedSurname && marriedSurname !== birthSurname
      ? `${given} ${birthSurname} ${marriedSurname}`.trim()
      : `${given} ${birthSurname}`.trim();
    dataRows[i][1] = displayName;
  });

  await batchUpdate(id, reqs, 'ppl-fmt');
  await valuesBatchUpdate(id, vals, 'ppl-vals');
  await valuesBatchUpdate(id, [{ range: `'${S}'!A8`, values: dataRows }], 'ppl-data');
  console.log(`✓ Master People — ${PEOPLE.length} people written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
