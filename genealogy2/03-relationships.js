'use strict';
const { sheets, batchUpdate, valuesBatchUpdate, gridRange, hex, colL, C } = require('./lib');
const { id, sheetMap } = JSON.parse(require('fs').readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Relationships'];
const S = 'Relationships';

const HEADERS = [
  'Relationship ID','Person 1 ID','Person 1 Name','Relationship Type',
  'Person 2 ID','Person 2 Name','Relationship Status','Start Date','End Date',
  'Source ID','Evidence Confidence','Expected Reciprocal','Notes',
]; // A-M (13 cols)

// [p1ID, relType, p2ID, relStatus, startDate, endDate, srcID, confidence, notes]
const RELS = [
  // === BIOLOGICAL PARENT-CHILD ===
  ['P-00002','Biological Parent','P-00001','Confirmed','1988-03-14','','SRC-00001','Confirmed','James Hartwell is Emma\'s father.'],
  ['P-00003','Biological Parent','P-00001','Confirmed','1988-03-14','','SRC-00001','Confirmed','Margaret Hartwell is Emma\'s mother.'],
  ['P-00002','Biological Parent','P-00026','Confirmed','1983-09-11','','SRC-00024','Confirmed','James is Thomas\'s father.'],
  ['P-00003','Biological Parent','P-00026','Confirmed','1983-09-11','','SRC-00024','Confirmed','Margaret is Thomas\'s mother.'],
  ['P-00002','Biological Parent','P-00027','Confirmed','1990-05-07','','SRC-00025','Confirmed','James is Sarah\'s father.'],
  ['P-00003','Biological Parent','P-00027','Confirmed','1990-05-07','','SRC-00025','Confirmed','Margaret is Sarah\'s mother.'],
  ['P-00002','Biological Parent','P-00045','Confirmed','1975-11-20','','SRC-00002','Moderate','James is Daniel\'s father (first marriage).'],
  ['P-00046','Biological Parent','P-00045','Confirmed','1975-11-20','','','Moderate','Carol Morrison is Daniel\'s mother.'],
  ['P-00004','Biological Parent','P-00002','Confirmed','1955-06-22','','SRC-00002','Confirmed','Robert Hartwell Sr. is James\'s father.'],
  ['P-00005','Biological Parent','P-00002','Confirmed','1955-06-22','','SRC-00002','Confirmed','Eleanor Hartwell is James\'s mother.'],
  ['P-00004','Biological Parent','P-00032','Confirmed','1952-04-18','','SRC-00027','Confirmed','Robert is Richard\'s father.'],
  ['P-00005','Biological Parent','P-00032','Confirmed','1952-04-18','','SRC-00027','Confirmed','Eleanor is Richard\'s mother.'],
  ['P-00008','Biological Parent','P-00004','Confirmed','1922-04-10','','SRC-00004','Strong','William Hartwell is Robert\'s father.'],
  ['P-00009','Biological Parent','P-00004','Confirmed','1922-04-10','','SRC-00004','Strong','Agnes Hartwell is Robert\'s mother.'],
  ['P-00008','Biological Parent','P-00075','Confirmed','1924-08-15','','SRC-00055','Confirmed','William is Robert Francis\'s father (KIA WWII Tarawa).'],
  ['P-00009','Biological Parent','P-00075','Confirmed','1924-08-15','','SRC-00055','Confirmed','Agnes is Robert Francis\'s mother.'],
  ['P-00008','Biological Parent','P-00076','Confirmed','1928-06-02','','SRC-00056','Confirmed','William is Mildred\'s father.'],
  ['P-00009','Biological Parent','P-00076','Confirmed','1928-06-02','','SRC-00056','Confirmed','Agnes is Mildred\'s mother.'],
  ['P-00008','Biological Parent','P-00044','Confirmed','1918-12-01','','SRC-00036','Confirmed','William is infant Elizabeth\'s father (died 6 weeks).'],
  ['P-00009','Biological Parent','P-00044','Confirmed','1918-12-01','','SRC-00036','Confirmed','Agnes is infant Elizabeth\'s mother.'],
  ['P-00010','Biological Parent','P-00005','Confirmed','1926-09-28','','SRC-00005','Confirmed','Charles Chapman is Eleanor\'s father.'],
  ['P-00011','Biological Parent','P-00005','Confirmed','1926-09-28','','SRC-00005','Confirmed','Harriet Chapman is Eleanor\'s mother.'],
  ['P-00016','Biological Parent','P-00008','Confirmed','1888-07-03','','SRC-00008','Confirmed','George Hartwell is William\'s father.'],
  ['P-00017','Biological Parent','P-00008','Confirmed','1888-07-03','','SRC-00008','Confirmed','Martha Holt Hartwell is William\'s mother.'],
  ['P-00018','Biological Parent','P-00009','Tentative','1892-03-17','','SRC-00018','Tentative','Edward Fletcher is Agnes\'s father. Record uncertain.'],
  ['P-00019','Biological Parent','P-00009','Tentative','1892-03-17','','SRC-00019','Tentative','Sarah Pickles Fletcher is Agnes\'s mother.'],
  ['P-00047','Biological Parent','P-00016','Moderate','1850-04-12','','SRC-00016','Moderate','Frederick Hartwell is George\'s father.'],
  ['P-00048','Biological Parent','P-00016','Moderate','1850-04-12','','SRC-00017','Moderate','Ann Turner Hartwell is George\'s mother.'],
  ['P-00063','Biological Parent','P-00047','Tentative','1815-06-15','','SRC-00046','Tentative','Thomas Hartwell (4x great-grandparent) is Frederick\'s probable father.'],
  ['P-00064','Biological Parent','P-00047','Tentative','1815-06-15','','SRC-00046','Tentative','Jane Booth Hartwell is Frederick\'s probable mother.'],
  ['P-00006','Biological Parent','P-00003','Confirmed','1958-11-03','','SRC-00003','Confirmed','Patrick O\'Brien is Margaret\'s father.'],
  ['P-00007','Biological Parent','P-00003','Confirmed','1958-11-03','','SRC-00003','Confirmed','Catherine O\'Brien is Margaret\'s mother.'],
  ['P-00006','Biological Parent','P-00036','Confirmed','1956-03-08','','SRC-00029','Confirmed','Patrick is Sean\'s father.'],
  ['P-00007','Biological Parent','P-00036','Confirmed','1956-03-08','','SRC-00029','Confirmed','Catherine is Sean\'s mother.'],
  ['P-00006','Biological Parent','P-00037','Confirmed','1960-07-14','','SRC-00030','Confirmed','Patrick is Colleen\'s father.'],
  ['P-00007','Biological Parent','P-00037','Confirmed','1960-07-14','','SRC-00030','Confirmed','Catherine is Colleen\'s mother.'],
  ['P-00012','Biological Parent','P-00006','Confirmed','1920-08-15','','SRC-00006','Confirmed','Michael O\'Brien is Patrick\'s father.'],
  ['P-00013','Biological Parent','P-00006','Confirmed','1920-08-15','','SRC-00006','Confirmed','Brigid O\'Brien is Patrick\'s mother.'],
  ['P-00012','Biological Parent','P-00073','Confirmed','1908-07-19','','SRC-00053','Confirmed','Michael is William James\'s father.'],
  ['P-00013','Biological Parent','P-00073','Confirmed','1908-07-19','','SRC-00053','Confirmed','Brigid is William James\'s mother.'],
  ['P-00012','Biological Parent','P-00071','Confirmed','1905-04-25','','SRC-00051','Confirmed','Michael is Joseph Bernard\'s father (KIA D-Day 1945).'],
  ['P-00013','Biological Parent','P-00071','Confirmed','1905-04-25','','SRC-00051','Confirmed','Brigid is Joseph Bernard\'s mother.'],
  ['P-00012','Biological Parent','P-00072','Moderate','1910-12-30','','SRC-00052','Moderate','Michael is Margaret Alice\'s father. No record after 1930.'],
  ['P-00013','Biological Parent','P-00072','Moderate','1910-12-30','','SRC-00052','Moderate','Brigid is Margaret Alice\'s mother.'],
  ['P-00012','Biological Parent','P-00074','Confirmed','1912-03-17','','SRC-00054','Confirmed','Michael is Kathleen\'s father.'],
  ['P-00013','Biological Parent','P-00074','Confirmed','1912-03-17','','SRC-00054','Confirmed','Brigid is Kathleen\'s mother.'],
  ['P-00020','Biological Parent','P-00012','Confirmed','1882-11-05','','SRC-00012','Confirmed','Cornelius O\'Brien is Michael\'s father.'],
  ['P-00021','Biological Parent','P-00012','Confirmed','1882-11-05','','SRC-00012','Confirmed','Honora Crowley O\'Brien is Michael\'s mother.'],
  ['P-00020','Biological Parent','P-00096','Confirmed','1880-03-28','','SRC-00065','Confirmed','Cornelius is Martin O\'Brien\'s father.'],
  ['P-00021','Biological Parent','P-00096','Confirmed','1880-03-28','','SRC-00065','Confirmed','Honora is Martin\'s mother.'],
  ['P-00022','Biological Parent','P-00013','Confirmed','1886-03-20','','SRC-00013','Confirmed','Daniel Shaughnessy is Brigid\'s father.'],
  ['P-00023','Biological Parent','P-00013','Confirmed','1886-03-20','','SRC-00013','Confirmed','Ellen Sullivan Shaughnessy is Brigid\'s mother.'],
  ['P-00049','Biological Parent','P-00020','Tentative','1845-10-08','','SRC-00020','Tentative','Jeremiah O\'Brien is Cornelius\'s father.'],
  ['P-00050','Biological Parent','P-00020','Tentative','1845-10-08','','SRC-00021','Tentative','Mary Driscoll O\'Brien is Cornelius\'s mother.'],
  ['P-00061','Biological Parent','P-00049','Tentative','1808-04-20','','SRC-00039','Tentative','Thaddeus O\'Brien is Jeremiah\'s probable father. Pre-famine era.'],
  ['P-00062','Biological Parent','P-00049','Tentative','1808-04-20','','SRC-00039','Tentative','Catherine McCarthy is Jeremiah\'s probable mother.'],
  ['P-00014','Biological Parent','P-00007','Confirmed','1924-02-04','','SRC-00007','Confirmed','Thomas Murphy is Catherine\'s father.'],
  ['P-00015','Biological Parent','P-00007','Confirmed','1924-02-04','','SRC-00007','Confirmed','Mary Gallagher Murphy is Catherine\'s mother.'],
  ['P-00014','Biological Parent','P-00080','Confirmed','1918-09-12','','SRC-00058','Confirmed','Thomas Murphy is Timothy\'s father.'],
  ['P-00015','Biological Parent','P-00080','Confirmed','1918-09-12','','SRC-00058','Confirmed','Mary Gallagher is Timothy\'s mother.'],
  ['P-00014','Biological Parent','P-00081','Confirmed','1920-05-27','','SRC-00059','Confirmed','Thomas Murphy is Frances\'s father.'],
  ['P-00015','Biological Parent','P-00081','Confirmed','1920-05-27','','SRC-00059','Confirmed','Mary Gallagher is Frances\'s mother.'],
  ['P-00024','Biological Parent','P-00014','Confirmed','1886-05-28','','SRC-00014','Confirmed','Patrick Murphy is Thomas\'s father.'],
  ['P-00025','Biological Parent','P-00014','Confirmed','1886-05-28','','SRC-00014','Confirmed','Bridget Ryan Murphy is Thomas\'s mother.'],
  ['P-00041','Biological Parent','P-00015','Tentative','1889-08-15','','SRC-00033','Tentative','John Gallagher is Mary\'s father. Limited Irish records.'],
  ['P-00042','Biological Parent','P-00015','Tentative','1889-08-15','','SRC-00034','Tentative','Anne Connelly Gallagher is Mary\'s mother.'],
  ['P-00026','Biological Parent','P-00029','Confirmed','2012-06-30','','','Strong','Thomas Hartwell is Oliver\'s father.'],
  ['P-00028','Biological Parent','P-00029','Confirmed','2012-06-30','','','Strong','Linda Patel Hartwell is Oliver\'s mother.'],
  ['P-00026','Biological Parent','P-00030','Confirmed','2015-11-18','','','Strong','Thomas is Sophia\'s father.'],
  ['P-00028','Biological Parent','P-00030','Confirmed','2015-11-18','','','Strong','Linda is Sophia\'s mother.'],
  ['P-00032','Biological Parent','P-00034','Confirmed','1978-03-09','','','Strong','Richard Hartwell is Kevin\'s father.'],
  ['P-00033','Biological Parent','P-00034','Confirmed','1978-03-09','','','Strong','Dorothy Simmons Hartwell is Kevin\'s mother.'],
  ['P-00032','Biological Parent','P-00035','Confirmed','1980-10-25','','','Strong','Richard is Jennifer\'s father.'],
  ['P-00033','Biological Parent','P-00035','Confirmed','1980-10-25','','','Strong','Dorothy is Jennifer\'s mother.'],
  ['P-00037','Biological Parent','P-00055','Confirmed','1983-04-17','','SRC-00043','Strong','Colleen Donovan is Fiona\'s mother.'],
  ['P-00054','Biological Parent','P-00055','Confirmed','1983-04-17','','SRC-00043','Strong','Patrick Donovan is Fiona\'s father.'],
  ['P-00037','Biological Parent','P-00056','Confirmed','1985-07-22','','SRC-00043','Strong','Colleen is Connor\'s mother.'],
  ['P-00054','Biological Parent','P-00056','Confirmed','1985-07-22','','SRC-00043','Strong','Patrick Donovan is Connor\'s father.'],
  ['P-00036','Biological Parent','P-00057','Strong','1980-05-14','','','Strong','Sean Patrick O\'Brien is Sean Michael\'s father.'],
  ['P-00036','Biological Parent','P-00092','Strong','1982-07-30','','','Strong','Sean Patrick is Declan\'s father.'],
  ['P-00096','Biological Parent','P-00098','Confirmed','1912-08-10','','SRC-00067','Confirmed','Martin O\'Brien is Patrick Martin\'s father.'],
  ['P-00097','Biological Parent','P-00098','Confirmed','1912-08-10','','SRC-00067','Confirmed','Anastasia Shea O\'Brien is Patrick Martin\'s mother.'],
  ['P-00098','Biological Parent','P-00100','Confirmed','1940-02-22','','SRC-00069','Confirmed','Patrick Martin O\'Brien is James Patrick\'s father.'],
  ['P-00099','Biological Parent','P-00100','Confirmed','1940-02-22','','SRC-00069','Confirmed','Evelyn McCarthy O\'Brien is James Patrick\'s mother.'],
  ['P-00092','Biological Parent','P-00094','Confirmed','2010-04-22','','','Confirmed','Declan O\'Brien is Liam\'s father.'],
  ['P-00093','Biological Parent','P-00094','Confirmed','2010-04-22','','','Confirmed','Aisling Murphy O\'Brien is Liam\'s mother.'],
  ['P-00092','Biological Parent','P-00095','Confirmed','2013-09-15','','','Confirmed','Declan is Siobhan\'s father.'],
  ['P-00093','Biological Parent','P-00095','Confirmed','2013-09-15','','','Confirmed','Aisling is Siobhan\'s mother.'],
  ['P-00077','Biological Parent','P-00078','Confirmed','1952-10-07','','SRC-00057','Strong','Howard Perry is Barbara\'s father.'],
  ['P-00076','Biological Parent','P-00078','Confirmed','1952-10-07','','SRC-00056','Strong','Mildred Hartwell Perry is Barbara\'s mother.'],
  ['P-00077','Biological Parent','P-00079','Confirmed','1955-03-29','','SRC-00057','Strong','Howard Perry is Harold\'s father.'],
  ['P-00076','Biological Parent','P-00079','Confirmed','1955-03-29','','SRC-00056','Strong','Mildred is Harold\'s mother.'],
  ['P-00086','Biological Parent','P-00069','Moderate','1915-09-01','','SRC-00062','Moderate','Salvatore Marino is Rosa\'s father.'],
  ['P-00087','Biological Parent','P-00069','Moderate','1915-09-01','','SRC-00063','Moderate','Giuseppa Ferrara Marino is Rosa\'s mother.'],
  ['P-00059','Biological Parent','P-00021','Tentative','1848-05-25','','SRC-00021','Tentative','Unknown Crowley is Honora\'s father. Brick wall.'],
  ['P-00066','Biological Parent','P-00021','Moderate','1848-05-25','','SRC-00047','Moderate','Nora O\'Connell Crowley is Honora\'s mother. Found in Griffith\'s Valuation.'],
  ['P-00065','Biological Parent','P-00017','Unknown','1854-09-20','','','Unknown','Francis Holt is Martha\'s father. No source records found.'],
  ['P-00065','Biological Parent','P-00090','Unknown','1855-01-01','','','Unknown','Francis Holt is Albert Holt\'s father. No source records.'],
  ['P-00085','Biological Parent','P-00048','Tentative','1820-09-10','','SRC-00061','Tentative','John Turner is Ann Turner\'s father.'],
  ['P-00084','Biological Parent','P-00048','Tentative','1820-09-10','','SRC-00061','Tentative','Hannah Turner is Ann Turner\'s mother.'],
  ['P-00041','Biological Parent','P-00088','Tentative','1879-10-18','','SRC-00033','Tentative','John Gallagher is Andrew\'s father. Andrew emigrated c.1900, no further record.'],
  ['P-00042','Biological Parent','P-00088','Tentative','1879-10-18','','SRC-00034','Tentative','Anne Connelly Gallagher is Andrew\'s mother.'],
  ['P-00091','Biological Parent','P-00064','Moderate','1785-01-01','','SRC-00064','Moderate','Alice Booth is Jane Booth\'s mother.'],
  ['P-00067','Biological Parent','P-00018','Possible','1858-02-07','','SRC-00035','Tentative','Alice Holt Fletcher possibly Edward Fletcher\'s mother. Research needed.'],
  ['P-00047','Biological Parent','P-00102','Tentative','','','SRC-00071','Tentative','Frederick Hartwell probably also Cecil\'s father (George\'s sibling).'],
  ['P-00048','Biological Parent','P-00102','Tentative','','','SRC-00071','Tentative','Ann Turner probably Cecil\'s mother. Unconfirmed.'],
  // === SPOUSES / PARTNERS ===
  ['P-00002','Spouse','P-00003','Confirmed','1981-06-15','','SRC-00002','Confirmed','James and Margaret Hartwell married c.1981.'],
  ['P-00002','Former Spouse','P-00046','Confirmed','1974-09-20','1979-05-01','SRC-00002','Moderate','James and Carol Morrison married 1974, divorced 1979.'],
  ['P-00004','Spouse','P-00005','Confirmed','1950-08-12','','SRC-00004','Confirmed','Robert and Eleanor Hartwell married c.1950.'],
  ['P-00006','Spouse','P-00007','Confirmed','1948-06-10','','SRC-00006','Confirmed','Patrick and Catherine O\'Brien married c.1948.'],
  ['P-00008','Spouse','P-00009','Confirmed','1916-07-04','','SRC-00008','Strong','William and Agnes Hartwell married c.1916.'],
  ['P-00010','Spouse','P-00011','Confirmed','1920-05-01','','SRC-00010','Strong','Charles and Harriet Chapman married c.1920.'],
  ['P-00012','Spouse','P-00013','Confirmed','1905-09-15','','SRC-00012','Confirmed','Michael and Brigid O\'Brien married c.1905.'],
  ['P-00014','Spouse','P-00015','Confirmed','1912-04-22','','SRC-00014','Confirmed','Thomas and Mary Murphy married c.1912. Providence.'],
  ['P-00016','Spouse','P-00017','Moderate','1875-01-01','','SRC-00016','Moderate','George and Martha Hartwell married c.1875. England.'],
  ['P-00018','Spouse','P-00019','Tentative','1890-01-01','','SRC-00018','Tentative','Edward Fletcher and Sarah Pickles married c.1890.'],
  ['P-00020','Spouse','P-00021','Tentative','1870-01-01','','SRC-00020','Tentative','Cornelius and Honora O\'Brien married c.1870. County Cork.'],
  ['P-00022','Spouse','P-00023','Moderate','1876-01-01','','SRC-00022','Moderate','Daniel Shaughnessy and Ellen Sullivan married c.1876.'],
  ['P-00024','Spouse','P-00025','Moderate','1878-01-01','','SRC-00023','Moderate','Patrick and Bridget Ryan Murphy married c.1878.'],
  ['P-00026','Spouse','P-00028','Confirmed','2009-08-15','','SRC-00026','Confirmed','Thomas Hartwell and Linda Patel married 2009. Chicago.'],
  ['P-00027','Spouse','P-00031','Confirmed','2013-09-28','','SRC-00025','Confirmed','Sarah Hartwell and Peter Novak married 2013.'],
  ['P-00032','Spouse','P-00033','Confirmed','1977-11-20','','SRC-00028','Confirmed','Richard and Dorothy Simmons Hartwell married 1977.'],
  ['P-00037','Spouse','P-00054','Confirmed','1982-03-17','2015-04-02','SRC-00030','Confirmed','Colleen O\'Brien and Patrick Donovan married 1982. Colleen died 2015.'],
  ['P-00047','Spouse','P-00048','Moderate','1845-01-01','','SRC-00037','Moderate','Frederick Hartwell and Ann Turner married c.1845. Warwickshire.'],
  ['P-00049','Spouse','P-00050','Tentative','1843-01-01','','SRC-00039','Tentative','Jeremiah and Mary Driscoll O\'Brien married c.1843. Pre-famine.'],
  ['P-00061','Spouse','P-00062','Tentative','1805-01-01','','SRC-00045','Tentative','Thaddeus and Catherine McCarthy O\'Brien married c.1805. Estimated.'],
  ['P-00063','Spouse','P-00064','Tentative','1810-01-01','','SRC-00046','Tentative','Thomas Hartwell and Jane Booth married c.1810. Estimated.'],
  ['P-00057','Spouse','P-00058','Confirmed','2007-06-01','','','Strong','Sean Michael O\'Brien and Nicole Rossi married 2007.'],
  ['P-00082','Spouse','P-00081','Confirmed','1940-10-12','','SRC-00060','Confirmed','Gerard Kelly and Frances Murphy Kelly married 1940.'],
  ['P-00077','Spouse','P-00076','Confirmed','1949-08-01','','SRC-00057','Confirmed','Howard Perry and Mildred Hartwell Perry married 1949.'],
  ['P-00092','Spouse','P-00093','Confirmed','2008-07-19','','','Confirmed','Declan O\'Brien and Aisling Murphy married 2008.'],
  ['P-00096','Spouse','P-00097','Confirmed','1908-11-01','','SRC-00065','Confirmed','Martin O\'Brien and Anastasia Shea married 1908.'],
  ['P-00098','Spouse','P-00099','Confirmed','1938-05-20','','SRC-00068','Confirmed','Patrick Martin and Evelyn McCarthy O\'Brien married 1938.'],
  ['P-00100','Spouse','P-00101','Confirmed','1965-09-04','','SRC-00069','Strong','James Patrick and Dorothy Walsh O\'Brien married 1965.'],
  ['P-00039','Spouse','P-00040','Moderate','1920-01-01','','SRC-00031','Moderate','Arthur Chapman and Rose Kelly Chapman married c.1920.'],
  ['P-00059','Spouse','P-00066','Tentative','1820-01-01','','SRC-00021','Tentative','Unknown Crowley and Nora O\'Connell married c.1820. County Cork.'],
  ['P-00068','Spouse','P-00069','Confirmed','1938-04-15','','SRC-00049','Strong','Domenico Rossi and Rosa Marino Rossi married 1938. Providence.'],
  ['P-00086','Spouse','P-00087','Moderate','1912-01-01','','SRC-00062','Moderate','Salvatore Marino and Giuseppa Ferrara Marino married c.1912.'],
  ['P-00083','Partner','P-00003','Confirmed','2021-01-01','','','Moderate','Christopher West has been Margaret Hartwell\'s partner since 2021. Not legally married.'],
  // === SIBLINGS ===
  ['P-00001','Sibling','P-00026','Confirmed','','','SRC-00001','Confirmed','Emma and Thomas Hartwell are siblings.'],
  ['P-00001','Sibling','P-00027','Confirmed','','','SRC-00001','Confirmed','Emma and Sarah Hartwell are siblings.'],
  ['P-00001','Half Sibling','P-00045','Confirmed','','','SRC-00002','Moderate','Emma and Daniel Hartwell are half-siblings. Same father James, different mothers.'],
  ['P-00001','Sibling','P-00038','Confirmed','','','','Unknown','Emma and Helen Hartwell (adopted) grew up as sisters.'],
  ['P-00026','Sibling','P-00027','Confirmed','','','SRC-00024','Confirmed','Thomas and Sarah Hartwell are siblings.'],
  ['P-00002','Sibling','P-00032','Confirmed','','','SRC-00027','Confirmed','James and Richard Hartwell are brothers.'],
  ['P-00003','Sibling','P-00036','Confirmed','','','SRC-00029','Confirmed','Margaret and Sean O\'Brien are siblings.'],
  ['P-00003','Sibling','P-00037','Confirmed','','','SRC-00030','Confirmed','Margaret and Colleen O\'Brien are siblings.'],
  ['P-00036','Sibling','P-00037','Confirmed','','','SRC-00029','Confirmed','Sean and Colleen O\'Brien are siblings.'],
  ['P-00004','Sibling','P-00075','Confirmed','','','SRC-00055','Confirmed','Robert Sr. and Robert Francis Hartwell are brothers. Robert Francis KIA Tarawa 1943.'],
  ['P-00004','Sibling','P-00076','Confirmed','','','SRC-00056','Confirmed','Robert Sr. and Mildred Hartwell Perry are siblings.'],
  ['P-00004','Sibling','P-00044','Confirmed','','','SRC-00036','Confirmed','Robert Sr. and infant Elizabeth Hartwell were siblings. Elizabeth died 6 weeks old 1919.'],
  ['P-00075','Sibling','P-00076','Confirmed','','','SRC-00055','Confirmed','Robert Francis and Mildred Hartwell are siblings.'],
  ['P-00006','Sibling','P-00073','Confirmed','','','SRC-00053','Confirmed','Patrick and William James O\'Brien are brothers.'],
  ['P-00006','Sibling','P-00071','Confirmed','','','SRC-00051','Confirmed','Patrick and Joseph Bernard O\'Brien are brothers. Joseph KIA D-Day 1945.'],
  ['P-00006','Sibling','P-00072','Moderate','','','SRC-00052','Moderate','Patrick and Margaret Alice O\'Brien are siblings. Margaret Alice\'s fate unknown after 1930.'],
  ['P-00006','Sibling','P-00074','Confirmed','','','SRC-00054','Confirmed','Patrick and Kathleen O\'Brien are siblings.'],
  ['P-00073','Sibling','P-00071','Confirmed','','','SRC-00051','Confirmed','William James and Joseph Bernard O\'Brien are brothers.'],
  ['P-00073','Sibling','P-00074','Confirmed','','','SRC-00053','Confirmed','William James and Kathleen O\'Brien are siblings.'],
  ['P-00007','Sibling','P-00080','Confirmed','','','SRC-00058','Confirmed','Catherine O\'Brien Murphy and Timothy Murphy are siblings.'],
  ['P-00007','Sibling','P-00081','Confirmed','','','SRC-00059','Confirmed','Catherine and Frances Murphy are sisters.'],
  ['P-00080','Sibling','P-00081','Confirmed','','','SRC-00058','Confirmed','Timothy and Frances Murphy are siblings.'],
  ['P-00057','Sibling','P-00092','Strong','','','','Strong','Sean Michael and Declan O\'Brien are brothers (same father Sean Patrick).'],
  ['P-00078','Sibling','P-00079','Confirmed','','','SRC-00057','Confirmed','Barbara and Harold Perry are siblings.'],
  ['P-00055','Sibling','P-00056','Confirmed','','','SRC-00043','Confirmed','Fiona and Connor Donovan are siblings.'],
  ['P-00094','Sibling','P-00095','Confirmed','','','','Confirmed','Liam and Siobhan O\'Brien are siblings.'],
  ['P-00016','Sibling','P-00102','Tentative','','','SRC-00071','Tentative','George Hartwell and Cecil Hartwell are probable brothers. Cecil stayed in England.'],
  ['P-00034','Sibling','P-00035','Confirmed','','','','Strong','Kevin and Jennifer Hartwell (Richard\'s children) are siblings.'],
  ['P-00029','Sibling','P-00030','Confirmed','','','','Strong','Oliver and Sophia Hartwell (Thomas\'s children) are siblings.'],
  ['P-00012','Sibling','P-00096','Confirmed','','','SRC-00065','Confirmed','Michael O\'Brien and Martin O\'Brien are brothers. Both immigrated to Providence.'],
  ['P-00039','Sibling','P-00010','Moderate','','','SRC-00031','Moderate','Arthur Chapman and Charles Chapman are brothers.'],
  // === ADOPTIVE ===
  ['P-00002','Adoptive Parent','P-00038','Confirmed','1994-01-01','','','Unknown','James Hartwell adopted Helen 1994. Birth family unknown.'],
  ['P-00003','Adoptive Parent','P-00038','Confirmed','1994-01-01','','','Unknown','Margaret Hartwell adopted Helen 1994.'],
  // === GRANDPARENT (explicit) ===
  ['P-00004','Grandparent','P-00001','Confirmed','','','SRC-00004','Confirmed','Robert Hartwell Sr. is Emma\'s paternal grandfather.'],
  ['P-00005','Grandparent','P-00001','Confirmed','','','SRC-00005','Confirmed','Eleanor Chapman Hartwell is Emma\'s paternal grandmother.'],
  ['P-00006','Grandparent','P-00001','Confirmed','','','SRC-00006','Confirmed','Patrick O\'Brien is Emma\'s maternal grandfather.'],
  ['P-00007','Grandparent','P-00001','Confirmed','','','SRC-00007','Confirmed','Catherine Murphy O\'Brien is Emma\'s maternal grandmother.'],
  ['P-00008','Grandparent','P-00002','Confirmed','','','SRC-00008','Strong','William Hartwell is James\'s paternal grandfather.'],
  ['P-00009','Grandparent','P-00002','Confirmed','','','SRC-00009','Strong','Agnes Fletcher Hartwell is James\'s paternal grandmother.'],
  ['P-00010','Grandparent','P-00002','Confirmed','','','SRC-00010','Strong','Charles Chapman is James\'s maternal grandfather (Eleanor\'s father).'],
  ['P-00011','Grandparent','P-00002','Confirmed','','','SRC-00011','Strong','Harriet Davies Chapman is James\'s maternal grandmother.'],
  ['P-00012','Grandparent','P-00003','Confirmed','','','SRC-00012','Confirmed','Michael O\'Brien is Margaret\'s paternal grandfather.'],
  ['P-00013','Grandparent','P-00003','Confirmed','','','SRC-00013','Confirmed','Brigid Shaughnessy O\'Brien is Margaret\'s paternal grandmother.'],
  ['P-00014','Grandparent','P-00003','Confirmed','','','SRC-00014','Confirmed','Thomas Murphy is Margaret\'s maternal grandfather.'],
  ['P-00015','Grandparent','P-00003','Confirmed','','','SRC-00015','Confirmed','Mary Gallagher Murphy is Margaret\'s maternal grandmother.'],
  // === DISPUTED / OTHER ===
  ['P-00060','Sibling','P-00012','Disputed','','','SRC-00044','Conflicting','James Flynn possibly a half-sibling of Michael O\'Brien. Two conflicting parish records.'],
];

(async () => {
  const reqs = [];
  const vals = [];

  // Tab setup
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, tabColor: hex(C.blush), gridProperties: { frozenRowCount: 7 } },
    fields: 'tabColor,gridProperties.frozenRowCount',
  }});

  // Row 1 — Title (A1:M1)
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 13), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 0, 1, 0, 13),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.blushDeep),
      textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  vals.push({ range: `'${S}'!A1`, values: [['RELATIONSHIPS & FAMILY CONNECTIONS']] });

  // Row 2 — Subtitle
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 13), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 1, 2, 0, 13),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.blush),
      textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  vals.push({ range: `'${S}'!A2`, values: [['All family relationships: parent-child, spouses, siblings, adoptions, and more. Column L shows the expected reciprocal relationship.']] });

  // Row 3-4 — Stats labels and values
  const statLabels = ['Total Relationships','Confirmed','Probable / Possible','Disputed','Active Spouses','Biological Parent'];
  const statCols = [0,2,4,6,8,10];
  statLabels.forEach((lbl, i) => {
    const c = statCols[i];
    reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, c, c+2), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, c, c+2), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(SID, 2, 3, c, c+2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.blushDeep),
        textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(SID, 3, 4, c, c+2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.input),
        textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.primaryDeep), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    }});
    vals.push({ range: `'${S}'!${colL(c)}3`, values: [[lbl]] });
  });

  const statFormulas = [
    `=COUNTA($B$8:$B$8007)`,
    `=COUNTIF($G$8:$G$8007,"Confirmed")`,
    `=COUNTIFS($G$8:$G$8007,"Probable")+COUNTIFS($G$8:$G$8007,"Possible")`,
    `=COUNTIF($G$8:$G$8007,"Disputed")`,
    `=COUNTIF($D$8:$D$8007,"Spouse")`,
    `=COUNTIF($D$8:$D$8007,"Biological Parent")`,
  ];
  statFormulas.forEach((f, i) => {
    vals.push({ range: `'${S}'!${colL(statCols[i])}4`, values: [[f]] });
  });

  // Row 5-6 — empty spacer with bg
  reqs.push({ repeatCell: {
    range: gridRange(SID, 4, 6, 0, 13),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat(backgroundColor)',
  }});

  // Row 7 — Column headers
  reqs.push({ repeatCell: {
    range: gridRange(SID, 6, 7, 0, 13),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.blushDeep),
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
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 4 },
    properties: { pixelSize: 22 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  // Column widths: A=110, B=90, C=180, D=150, E=90, F=180, G=110, H=100, I=100, J=100, K=120, L=150, M=250
  const colWidths = [110,90,180,150,90,180,110,100,100,100,120,150,250];
  colWidths.forEach((px, ci) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  // Freeze rows/cols already set via gridProperties above

  // Border on header row
  reqs.push({ updateBorders: {
    range: gridRange(SID, 6, 7, 0, 13),
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.border) },
  }});

  // Conditional formatting — Relationship Status (col G = col index 6)
  const CF_STATUS = [
    { val: 'Confirmed',  bg: C.confirmed },
    { val: 'Probable',   bg: C.info },
    { val: 'Possible',   bg: C.neutral },
    { val: 'Disputed',   bg: C.conflict },
    { val: 'Unknown',    bg: C.neutral },
  ];
  CF_STATUS.forEach(({ val, bg }) => {
    reqs.push({ addConditionalFormatRule: { rule: {
      ranges: [gridRange(SID, 7, 8007, 6, 7)],
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: val }] },
        format: { backgroundColor: hex(bg) },
      },
    }, index: 0 }});
  });

  // Conditional formatting — Evidence Confidence (col K = col index 10)
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
      ranges: [gridRange(SID, 7, 8007, 10, 11)],
      booleanRule: {
        condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: val }] },
        format: { backgroundColor: hex(bg) },
      },
    }, index: 0 }});
  });

  // Data validation
  // D = Relationship Types (col 3) — Reference Data!$C$2:$C$19
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 8007, 3, 4),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$C$2:$C$19` }] },
      strict: false, showCustomUi: true,
    },
  }});
  // G = Relationship Status (col 6) — Reference Data!$D$2:$D$6
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 8007, 6, 7),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$D$2:$D$6` }] },
      strict: false, showCustomUi: true,
    },
  }});
  // K = Evidence Confidence (col 10) — Reference Data!$J$2:$J$7
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 8007, 10, 11),
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$J$2:$J$7` }] },
      strict: false, showCustomUi: true,
    },
  }});

  // Alternating row fill (data rows)
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, 7, 8007, 0, 13)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND(MOD(ROW(),2)=0,LEN($B8)>0)' }] },
      format: { backgroundColor: hex(C.altRow) },
    },
  }, index: 0 }});

  // Build data rows
  const dataRows = RELS.map(([p1, rel, p2, status, startDt, endDt, src, conf, notes]) => [
    `=IF(OR(B${'{R}'}="",E${'{R}'}=""),"","REL-"&TEXT(ROW()-7,"00000"))`,
    p1,
    `=IFERROR(VLOOKUP(B${'{R}'},'Master People'!$A$8:$B$5007,2,FALSE),"")`,
    rel,
    p2,
    `=IFERROR(VLOOKUP(E${'{R}'},'Master People'!$A$8:$B$5007,2,FALSE),"")`,
    status,
    startDt,
    endDt,
    src,
    conf,
    `=IFERROR(INDEX('Reference Data'!$R$2:$R$20,MATCH(D${'{R}'},'Reference Data'!$Q$2:$Q$20,0)),"")`,
    notes,
  ]);

  // Replace {R} row placeholders
  dataRows.forEach((row, i) => {
    const r = i + 8;
    for (let c = 0; c < row.length; c++) {
      if (typeof row[c] === 'string') {
        row[c] = row[c].replace(/\{R\}/g, String(r));
      }
    }
  });

  await batchUpdate(id, reqs, 'rel-fmt');
  await valuesBatchUpdate(id, vals, 'rel-vals');
  await valuesBatchUpdate(id, [{ range: `'${S}'!A8`, values: dataRows }], 'rel-data');
  console.log(`✓ Relationships — ${RELS.length} rows written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
