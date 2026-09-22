'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const FM = sheetMap['👨‍👩‍👧‍👦 Family Members'];

// Columns A-O (15):
// ID, Full Name, Nickname, Birth Date, Birth Place, Death Date, Death Place,
// Cause of Death, Nationality, Languages Spoken, Highest Education,
// Parent 1 ID, Parent 2 ID, Spouse ID, Notes

const HEADERS = [
  'ID','Full Name','Nickname','Birth Date','Birth Place','Death Date','Death Place',
  'Cause of Death','Nationality','Languages Spoken','Highest Education',
  'Parent 1 ID','Parent 2 ID','Spouse ID','Notes'
];

// GEN 1 (IDs 1-8): Great-grandparents — all deceased, full fields
// GEN 2 (IDs 9-18): Grandparents — mostly deceased
// GEN 3 (IDs 19-30): Parents & siblings — mix living/dead
// GEN 4 (IDs 31-40): User's generation — all living (Death Date/Place/Cause = '' '' '')
const ROWS = [
  // GEN 1
  [1,'Patrick Joseph Sullivan','Pat','1878-03-15','County Cork, Ireland','1952-11-02','Boston, MA, USA','Heart failure','Irish-American','English, Irish Gaelic','Primary school','','','2','Emigrated from Ireland 1901; worked as dockworker in Boston.'],
  [2,'Brigid Mary Flanagan','Biddy','1882-07-04','County Kerry, Ireland','1960-05-20','Boston, MA, USA','Pneumonia','Irish-American','English, Irish Gaelic','Primary school','','','1','Arrived Ellis Island 1904. Known for soda bread recipe.'],
  [3,'Józef Kowalski','Józek','1875-09-22','Kraków, Poland','1943-04-01','Warsaw, Poland','Wartime','Polish','Polish, Russian','Trade apprenticeship','','','4','Cabinet maker. Died during WWII occupation.'],
  [4,'Maria Wisniewska','Marysia','1880-02-11','Łódź, Poland','1958-08-14','Chicago, IL, USA','Old age','Polish-American','Polish, English','Primary school','','','3','Emigrated 1921 after the war. Settled in Chicago Polish neighborhood.'],
  [5,'Taro Nakamura','','1885-05-30','Hiroshima, Japan','1945-08-06','Hiroshima, Japan','Atomic bombing','Japanese','Japanese','Secondary school','','','6','Agricultural farmer. Victim of Hiroshima bombing.'],
  [6,'Yuki Tanaka','','1890-11-03','Osaka, Japan','1968-01-17','Honolulu, HI, USA','Cancer','Japanese-American','Japanese, English','Primary school','','','5','Emigrated to Hawaii 1920. Worked in sugar cane fields.'],
  [7,'Wilhelm Fischer','Willi','1872-12-08','Bavaria, Germany','1935-03-22','Hamburg, Germany','Tuberculosis','German','German, French','University','','','8','Schoolteacher. Letters survive in family archive.'],
  [8,'Hilde Braun','','1877-06-19','Stuttgart, Germany','1949-09-30','New York, NY, USA','Heart disease','German-American','German, English','Secondary school','','','7','Widowed 1935; emigrated to USA 1938 fleeing political persecution.'],
  // GEN 2 — deceased
  [9,'Thomas Patrick Sullivan','Tommy','1908-06-10','Boston, MA, USA','1979-12-03','Boston, MA, USA','Stroke','American','English','High school','1','2','10','WWII veteran, Pacific theater. Loved baseball.'],
  [10,'Eleanor Rose Kowalski','Ellie','1912-04-25','Chicago, IL, USA','1995-07-18','Boston, MA, USA','Alzheimers','American','English, Polish','High school','3','4','9','Taught herself piano. Made pierogies every Christmas.'],
  [11,'Kenji Nakamura','Ken','1915-08-12','Honolulu, HI, USA','1998-02-28','Los Angeles, CA, USA','Respiratory failure','Japanese-American','Japanese, English','College','5','6','12','Interned at Manzanar 1942-44. Became electrical engineer post-war.'],
  [12,'Ruth Miriam Fischer','Ruthie','1920-03-07','New York, NY, USA','2003-10-10','Los Angeles, CA, USA','Cancer','German-American','English, German','College','7','8','11','Concert violinist in local orchestra for 30 years.'],
  [13,'John Joseph Sullivan','Jack','1910-11-30','Boston, MA, USA','1968-07-04','Boston, MA, USA','Heart attack','American','English','High school','1','2','14','Older brother of Tommy. Korean War veteran.'],
  [14,'Agnes Vera Murphy','','1913-09-01','Quincy, MA, USA','1999-03-14','Boston, MA, USA','Old age','American','English','High school','','','13','School librarian for 35 years.'],
  [15,'Stefan Kowalski','Steve','1918-01-14','Chicago, IL, USA','2001-05-22','Milwaukee, WI, USA','Parkinsons','American','English, Polish','Trade school','3','4','16','Younger brother of Eleanor. Auto mechanic.'],
  [16,'Dorothy Ann Wallace','Dot','1922-08-03','Milwaukee, WI, USA','2010-11-01','Milwaukee, WI, USA','Natural causes','American','English','High school','','','15','Homemaker and community volunteer.'],
  // GEN 2 — living
  [17,'Hiroshi Nakamura','','1940-05-20','Manzanar, CA, USA','','','','Japanese-American','Japanese, English','College','11','12','18','Born in internment camp. Built successful import business.'],
  [18,'Linda Sato','','1942-09-15','Seattle, WA, USA','','','','Japanese-American','English, Japanese','College','','','17','Retired teacher. Still lives in Los Angeles.'],
  // GEN 3 — mix
  [19,'Michael Thomas Sullivan','Mike','1940-08-22','Boston, MA, USA','2018-03-10','Boston, MA, USA','Prostate cancer','American','English','College','9','10','20','Retired attorney. Coached Little League for 20 years.'],
  [20,'Carol Jean Nakamura','Carol','1943-11-04','Los Angeles, CA, USA','','','','American','English, Japanese','College','11','12','19','Retired nurse. Keeps the family photo albums.'],
  [21,'Patricia Ann Sullivan','Patty','1942-05-17','Boston, MA, USA','','','','American','English','College','9','10','22','Elder sister of Michael. Retired school principal.'],
  [22,'Robert Allen Greene','Bob','1939-02-28','Providence, RI, USA','2005-08-15','Providence, RI, USA','Accident','American','English','High school','','','21','Worked in construction.'],
  [23,'James Kevin Sullivan','Jimmy','1947-12-01','Boston, MA, USA','','','','American','English','High school','9','10','24','Youngest Sullivan sibling. Runs a hardware store.'],
  [24,'Barbara Ann Kelly','Barb','1950-07-19','Cambridge, MA, USA','','','','American','English','Some college','','','23','Bookkeeper. Avid gardener.'],
  [25,'Daniel John Sullivan','Danny','1945-09-03','Boston, MA, USA','1972-06-20','Vietnam','Combat wounds','American','English','Some college','13','14','','Vietnam War KIA. Photo on family mantle.'],
  [26,'Susan Marie Sullivan','Sue','1948-03-11','Boston, MA, USA','','','','American','English','College','13','14','27','Social worker. Moved to Arizona 1985.'],
  [27,'Peter Andrew Ramirez','Pete','1945-10-28','Tucson, AZ, USA','2019-01-03','Tucson, AZ, USA','Heart failure','American','English, Spanish','High school','','','26','Retired postal worker.'],
  [28,'Steven Paul Kowalski','Steve Jr.','1950-04-07','Milwaukee, WI, USA','','','','American','English, Polish','Trade school','15','16','29','Plumber. Keeps grandfather\'s cabinet tools.'],
  [29,'Mary Lou Henderson','','1952-11-22','Milwaukee, WI, USA','','','','American','English','Some college','','','28','Dental hygienist.'],
  [30,'Alan James Nakamura','Al','1968-07-14','Los Angeles, CA, USA','','','','Japanese-American','English, Japanese','Graduate','17','18','','Business consultant. Unmarried.'],
  // GEN 4 — all living
  [31,'Sarah Elizabeth Sullivan','Sarah','1970-05-14','Boston, MA, USA','','','','American','English, French','Graduate','19','20','32','THE USER — genealogy researcher.'],
  [32,'David Alan Park','Dave','1968-09-21','Portland, OR, USA','','','','American','English','Graduate','','','31','Married Sarah 1996.'],
  [33,'Christopher Michael Sullivan','Chris','1972-11-02','Boston, MA, USA','','','','American','English','College','19','20','34','Sarah\'s brother. Works in finance.'],
  [34,'Jennifer Lynn Torres','Jen','1974-03-17','Miami, FL, USA','','','','American','English, Spanish','College','','','33','Marketing director.'],
  [35,'Kevin Robert Greene','Kev','1968-07-25','Providence, RI, USA','','','','American','English','Trade school','22','21','36','Electrician.'],
  [36,'Amy Lynn Greene','','1970-09-03','Providence, RI, USA','','','','American','English','College','22','21','35','Graphic designer.'],
  [37,'Thomas James Sullivan','Tommy Jr.','1975-04-08','Boston, MA, USA','','','','American','English','High school','23','24','','Hardware store co-owner. Unmarried.'],
  [38,'Emily Grace Sullivan','Em','1978-12-19','Boston, MA, USA','','','','American','English','College','23','24','39','Therapist in private practice.'],
  [39,'Marcus William Chen','Marc','1977-02-14','San Francisco, CA, USA','','','','Chinese-American','English, Mandarin','Graduate','','','38','Software engineer.'],
  [40,'Mia Yuki Nakamura','Mia','1995-06-30','Los Angeles, CA, USA','','','','Japanese-American','English, Japanese','College','30','','','Alan\'s daughter. Graphic designer.'],
];

(async () => {
  // Verify column counts
  ROWS.forEach((r, i) => {
    if (r.length !== 15) throw new Error(`Row ${i+1} (ID ${r[0]}) has ${r.length} cols, expected 15`);
  });

  const data = [
    { range: "'👨‍👩‍👧‍👦 Family Members'!A1:O1", values: [HEADERS] },
    { range: "'👨‍👩‍👧‍👦 Family Members'!A2:O41", values: ROWS },
  ];
  await valuesBatchUpdate(id, data, 'fm-values');

  const reqs = [];

  reqs.push({
    repeatCell: {
      range: gridRange(FM, 0, 1, 0, 15),
      cell: { userEnteredFormat: { backgroundColor: hex(C.sepia), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  for (let r = 1; r <= 40; r++) {
    const bg = r % 2 === 1 ? C.altRow : C.white;
    reqs.push({
      repeatCell: {
        range: gridRange(FM, r, r+1, 0, 15),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
  }

  reqs.push({
    repeatCell: {
      range: gridRange(FM, 1, 41, 0, 1),
      cell: { userEnteredFormat: { textFormat: { bold: true }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(textFormat,horizontalAlignment)',
    },
  });

  [[0,40],[1,180],[2,90],[3,100],[4,160],[5,100],[6,160],[7,130],[8,110],[9,140],[10,110],[11,70],[12,70],[13,70],[14,220]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: FM, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  reqs.push({ updateDimensionProperties: { range: { sheetId: FM, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'fm-format');
  console.log('Family Members complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
