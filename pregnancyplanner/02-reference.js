'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

(async () => {
  const reqs = [];

  // Tab header
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: REF, gridProperties: { frozenRowCount: 1 } },
    fields: 'gridProperties.frozenRowCount',
  }});

  // Column headers background (row 0)
  reqs.push({ repeatCell: {
    range: gridRange(REF, 0, 1, 0, 22),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});

  // Alt row shading rows 1-54
  for (let r = 1; r < 54; r += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(REF, r, r+1, 0, 22),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderMist) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }

  // Column widths
  const colW = [70,120,200,160,90,90,160,140,130,90,120,100,120,120,110,110,120,120,120,100,100,100];
  colW.forEach((w, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: REF, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: w }, fields: 'pixelSize',
  }}));

  await batchUpdate(id, reqs, 'ref-format');

  // ── Values ──────────────────────────────────────────────────────────────────
  const data = [];

  // Row 0: headers
  data.push({ range: `'📋 Reference Data'!A1:V1`, values: [[
    'Week','Baby Size','Milestone',
    'Symptoms','Severity','Mood',
    'Appointment Types','Budget Categories','Nursery Categories','Priority',
    'Bag Sections','Name Vibes','Meal Categories','Dietary Tags','Meal Status',
    'Wellbeing','Feeding','Contact Roles','Note Topics','Note Status',
    'RSVP','Trimester',
  ]]});

  // Baby size/milestone table A2:C40 (weeks 4-42)
  const weeks = [
    [4,'Poppy seed','Neural tube forming'],
    [5,'Apple seed','Heart starts beating'],
    [6,'Sweet pea','Brain & spinal cord forming'],
    [7,'Blueberry','Facial features forming'],
    [8,'Raspberry','Fingers & toes developing'],
    [9,'Grape','Muscles forming'],
    [10,'Kumquat','Critical organs complete'],
    [11,'Fig','Tooth buds forming'],
    [12,'Lime','Reflexes developing'],
    [13,'Lemon','Fingerprints forming'],
    [14,'Peach','Baby makes facial expressions'],
    [15,'Apple','Baby senses light'],
    [16,'Avocado','Baby can hear sounds'],
    [17,'Pear','Baby practicing swallowing'],
    [18,'Sweet potato','Ears moving to final position'],
    [19,'Mango','Vernix caseosa forming'],
    [20,'Banana','Halfway! Kicks now felt'],
    [21,'Carrot','Eyebrows & eyelids visible'],
    [22,'Spaghetti squash','Lanugo covers body'],
    [23,'Grapefruit','Sense of movement developing'],
    [24,'Corn','Lung surfactant forming (viability!)'],
    [25,'Cauliflower','Nostrils beginning to open'],
    [26,'Scallion','Eyes beginning to open'],
    [27,'Cabbage','Brain very active, dreaming possible'],
    [28,'Eggplant','3rd trimester! Fat layers forming'],
    [29,'Butternut squash','Muscles & lungs maturing'],
    [30,'Cabbage','Brain developing rapidly'],
    [31,'Coconut','Baby can track light'],
    [32,'Squash','Fingernails & toenails complete'],
    [33,'Pineapple','Skull bones hardening'],
    [34,'Cantaloupe','CNS & lungs maturing'],
    [35,'Honeydew','Kidneys fully developed'],
    [36,'Head of romaine','Gaining ~1oz per day'],
    [37,'Winter melon','Considered early term'],
    [38,'Leek','Vocal cords developed'],
    [39,'Mini watermelon','Full term! Ready to meet you'],
    [40,'Pumpkin','Due date week!'],
    [41,'Jackfruit','Post-term, monitored closely'],
    [42,'Watermelon','Induction likely recommended'],
  ];
  data.push({ range: `'📋 Reference Data'!A2:C40`, values: weeks });

  // D: Symptoms (col 3)
  const symptoms = [
    'Morning sickness','Fatigue','Heartburn','Back pain','Round ligament pain',
    'Swelling (edema)','Braxton Hicks','Shortness of breath','Headache','Insomnia',
    'Leg cramps','Frequent urination','Food cravings','Food aversions','Mood swings',
    'Constipation','Bloating','Tender breasts','Dizziness','Nausea',
  ];
  data.push({ range: `'📋 Reference Data'!D2:D21`, values: symptoms.map(s => [s]) });

  // E: Severity
  data.push({ range: `'📋 Reference Data'!E2:E6`, values: [['Mild'],['Moderate'],['Severe'],['Very Severe'],['Manageable']] });

  // F: Mood
  data.push({ range: `'📋 Reference Data'!F2:F9`, values: [['😊 Happy'],['😌 Content'],['😰 Anxious'],['😢 Emotional'],['😤 Frustrated'],['😴 Exhausted'],['🥰 Excited'],['😐 Neutral']] });

  // G: Appointment Types
  const apptTypes = ['OB Checkup','Midwife Visit','Ultrasound','Blood Draw','GD Screening','Anatomy Scan','GBS Test','NST','BPP','Specialist Consult','Dental','Chiropractor'];
  data.push({ range: `'📋 Reference Data'!G2:G13`, values: apptTypes.map(a => [a]) });

  // H: Budget Categories
  const budgetCats = ['Nursery & Furniture','Clothing & Gear','Medical & Healthcare','Stroller & Car Seat','Feeding Supplies','Postpartum Care','Education & Books','Entertainment & Toys','Shower & Celebrations','Travel & Transport','Tech & Gadgets','Miscellaneous'];
  data.push({ range: `'📋 Reference Data'!H2:H13`, values: budgetCats.map(c => [c]) });

  // I: Nursery Categories
  const nurseryCats = ['Furniture','Bedding & Textiles','Storage & Organisation','Lighting','Décor & Wall Art','Safety','Electronics','Feeding Station','Changing Area','Clothing Storage'];
  data.push({ range: `'📋 Reference Data'!I2:I11`, values: nurseryCats.map(c => [c]) });

  // J: Priority
  data.push({ range: `'📋 Reference Data'!J2:J4`, values: [['High'],['Medium'],['Low']] });

  // K: Bag Sections
  data.push({ range: `'📋 Reference Data'!K2:K7`, values: [['Labour & Delivery'],['Postpartum (Mum)'],['Postpartum (Baby)'],['Partner Bag'],['Documents'],['Going Home']] });

  // L: Name Vibes
  data.push({ range: `'📋 Reference Data'!L2:L9`, values: [['Classic'],['Modern'],['Nature'],['Vintage'],['Unique'],['Cultural'],['Spiritual'],['Whimsical']] });

  // M: Meal Categories
  data.push({ range: `'📋 Reference Data'!M2:M8`, values: [['Casserole'],['Soup & Stew'],['Pasta'],['Breakfast'],['Snack & Bites'],['Grains & Rice'],['Baked Goods']] });

  // N: Dietary Tags
  data.push({ range: `'📋 Reference Data'!N2:N7`, values: [['Vegetarian'],['Vegan'],['Gluten-Free'],['Dairy-Free'],['Nut-Free'],['None']] });

  // O: Meal Status
  data.push({ range: `'📋 Reference Data'!O2:O5`, values: [['Planned'],['Prepped'],['Frozen'],['Used']] });

  // P: Wellbeing (postpartum)
  data.push({ range: `'📋 Reference Data'!P2:P7`, values: [['Great'],['Good'],['Fair'],['Struggling'],['Needs Support'],['Crisis — Get Help']] });

  // Q: Feeding
  data.push({ range: `'📋 Reference Data'!Q2:Q5`, values: [['Breastfeeding'],['Formula'],['Mixed'],['Pumping']] });

  // R: Contact Roles
  data.push({ range: `'📋 Reference Data'!R2:R11`, values: [['OB/Midwife'],['Paediatrician'],['Lactation Consultant'],['Doula'],['Partner'],['Family'],['Friend'],['Specialist'],['Therapist'],['Other']] });

  // S: Note Topics
  data.push({ range: `'📋 Reference Data'!S2:S9`, values: [['Birth Plan'],['Feeding Notes'],['Sleep Notes'],['Milestone Notes'],['Medical Notes'],['Emotional Check-in'],['Gratitude'],['General']] });

  // T: Note Status
  data.push({ range: `'📋 Reference Data'!T2:T4`, values: [['Draft'],['Final'],['Archived']] });

  // U: RSVP (shower)
  data.push({ range: `'📋 Reference Data'!U2:U5`, values: [['Yes'],['No'],['Maybe'],['Pending']] });

  // V: Trimester
  data.push({ range: `'📋 Reference Data'!V2:V4`, values: [['First (Weeks 1-13)'],['Second (Weeks 14-27)'],['Third (Weeks 28-40)']] });

  await valuesBatchUpdate(id, data, 'ref-values');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
