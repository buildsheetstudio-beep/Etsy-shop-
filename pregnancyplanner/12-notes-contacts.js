'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const NC = sheetMap['📝 Notes & Contacts'];

// 8 cols, 70 rows
// Row 0: Title, Row 1: Contacts section header, Row 2: col headers, Row 3-18: contacts
// Row 19: Birth Plan section header, Row 20: col headers, Row 21-35: birth plan items
// Row 36: General Notes section header, Row 37: col headers, Row 38-69: notes

(async () => {
  const reqs = [];

  // Title
  reqs.push({ mergeCells: { range: gridRange(NC, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(NC, 0, 1, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: NC, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });

  // Section headers
  const sections = [1, 19, 36];
  sections.forEach(r => {
    reqs.push({ mergeCells: { range: gridRange(NC, r, r+1, 0, 8), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(NC, r, r+1, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.softLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'LEFT' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
    }});
    reqs.push({ updateDimensionProperties: { range: { sheetId: NC, dimension: 'ROWS', startIndex: r, endIndex: r+1 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  });

  // Col header rows 2, 20, 37
  [2, 20, 37].forEach(r => {
    reqs.push({ repeatCell: {
      range: gridRange(NC, r, r+1, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.taupeGold), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
    reqs.push({ updateDimensionProperties: { range: { sheetId: NC, dimension: 'ROWS', startIndex: r, endIndex: r+1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  });

  // Alt rows contacts 3-18
  for (let r = 3; r < 19; r += 2) {
    reqs.push({ repeatCell: { range: gridRange(NC, r, r+1, 0, 8), cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderMist) } }, fields: 'userEnteredFormat.backgroundColor' }});
  }
  // Alt rows birth plan 21-35
  for (let r = 21; r < 36; r += 2) {
    reqs.push({ repeatCell: { range: gridRange(NC, r, r+1, 0, 8), cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderMist) } }, fields: 'userEnteredFormat.backgroundColor' }});
  }
  // Alt rows notes 38-69
  for (let r = 38; r < 70; r += 2) {
    reqs.push({ repeatCell: { range: gridRange(NC, r, r+1, 0, 8), cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderMist) } }, fields: 'userEnteredFormat.backgroundColor' }});
  }

  // Wrap
  reqs.push({ repeatCell: { range: gridRange(NC, 3, 70, 0, 8), cell: { userEnteredFormat: { wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(wrapStrategy,verticalAlignment)' }});

  // Data row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: NC, dimension: 'ROWS', startIndex: 3, endIndex: 70 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  // Col widths
  const colW = [30, 160, 130, 130, 160, 160, 130, 220];
  colW.forEach((w, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: NC, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: w }, fields: 'pixelSize',
  }}));

  const border = { style: 'SOLID', color: hex(C.lavenderBorder) };
  [[2,19],[20,36],[37,70]].forEach(([s,e]) => reqs.push({ updateBorders: { range: gridRange(NC, s, e, 0, 8), innerHorizontal: border, innerVertical: border, top: border, bottom: border, left: border, right: border }}));

  await batchUpdate(id, reqs, 'nc-format');

  const vdata = [];
  vdata.push({ range: `'📝 Notes & Contacts'!A1`, values: [['📝 Notes & Contacts']] });

  vdata.push({ range: `'📝 Notes & Contacts'!A2`, values: [['  👥 Key Contacts']] });
  vdata.push({ range: `'📝 Notes & Contacts'!A3:H3`, values: [['#','Name','Role','Phone','Email','Practice / Hospital','Notes','Emergency Contact?']] });

  const contacts = [
    [1,'Dr Sarah Collins','OB/Midwife','(555) 234-5678','scollins@cityob.com','City OB Clinic','Primary OB. Appts every 4 weeks, then 2 weeks, then weekly.','Yes'],
    [2,'Maria Lopez','Midwife','(555) 345-6789','mlopez@cityob.com','City OB Clinic','Back-up midwife when Dr Collins unavailable.','Yes'],
    [3,'St Mary\'s Labour & Delivery','Hospital','(555) 100-0000','labordel@stmarys.com','St Mary\'s Medical Centre','Pre-registered. Know the parking situation!','Yes'],
    [4,'Dr James Park','Paediatrician','(555) 456-7890','jpark@kidsfirst.com','Kids First Paediatrics','Chosen paed. Will do newborn check at hospital.','Yes'],
    [5,'Lisa Chen','Lactation Consultant','(555) 567-8901','lchen@breastfeed.com','Breastfeeding Support Centre','Can do home visits. Book in advance for week 1.','No'],
    [6,'Emma Rowe','Doula','(555) 678-9012','emma@doulasupport.com','Independent','Hired for labour support. On call from week 37.','Yes'],
    [7,'My Mum','Family','(555) 789-0123','mum@email.com','','Coming to stay week 1. Pick up from airport Dec 10.','Yes'],
    [8,'Partner — Jake','Partner','(555) 890-1234','jake@email.com','','Primary support person. Works from home Tue-Thu.','Yes'],
    [9,'Insurance — BlueCross','Healthcare','1-800-555-2345','','bcbs.com','Policy #BC-123456789. Pre-auth for epidural done.','No'],
    [10,'Postpartum Support Int\'l','Specialist','1-800-944-4773','','postpartum.net','Helpline — 24/7 support for PPD & anxiety.','Yes'],
  ];
  vdata.push({ range: `'📝 Notes & Contacts'!A4:H13`, values: contacts });

  vdata.push({ range: `'📝 Notes & Contacts'!A20`, values: [['  🌿 Birth Plan Summary']] });
  vdata.push({ range: `'📝 Notes & Contacts'!A21:H21`, values: [['#','Topic','My Preference','Alternatives OK?','Notes','','','']] });

  const birthPlan = [
    [1,'Pain management','Epidural preferred, open to natural','Yes','Try breathing & movement first, epidural if needed','','',''],
    [2,'Labour environment','Calm, dim lighting, soft music','Yes','Bring Bluetooth speaker with playlist','','',''],
    [3,'Monitoring','Intermittent if low risk','Yes','Continuous if needed for baby safety','','',''],
    [4,'Labour mobility','Move freely, use birth ball','Yes','Wireless monitor preferred if possible','','',''],
    [5,'Pushing','Breathing techniques, avoid coached','Yes','Follow midwife guidance','','',''],
    [6,'Cord clamping','Delayed cord clamping requested','Yes','At least 1-2 minutes','','',''],
    [7,'Skin to skin','Immediate, uninterrupted if possible','Yes','At least 1 hour for bonding','','',''],
    [8,'Feeding','Breastfeeding — no formula unless medical','Yes','Lactation consult within first hour','','',''],
    [9,'Photography','Jake to take photos/video','No','First moments private','','',''],
    [10,'Visitors','No visitors during labour','No','Immediate family only after birth, 2-hour delay','','',''],
    [11,'Vaccinations','Yes to Hep B and Vit K for baby','No','Standard newborn care','','',''],
    [12,'Circumcision','N/A','N/A','','','',''],
  ];
  vdata.push({ range: `'📝 Notes & Contacts'!A22:H33`, values: birthPlan });

  vdata.push({ range: `'📝 Notes & Contacts'!A37`, values: [['  📋 General Notes']] });
  vdata.push({ range: `'📝 Notes & Contacts'!A38:H38`, values: [['#','Date','Topic','Status','Notes','Tags','','Priority']] });

  const notes = [
    [1,'2025-06-01','Feeding Notes','Final','Starting research on breastfeeding. Watching videos, reading Ina May.','Feeding, Preparation','','High'],
    [2,'2025-07-15','Birth Plan','Draft','Drafted preferences with Jake. Need to discuss epidural decision.','Birth Plan, Medical','','High'],
    [3,'2025-08-01','Nursery Notes','Final','Finalized nursery theme: botanical garden. Ordered crib and glider.','Nursery, Shopping','','Medium'],
    [4,'2025-08-20','Medical Notes','Final','GD screening passed. No gestational diabetes. Huge relief!','Medical','','High'],
    [5,'2025-09-01','Sleep Notes','Draft','Reading Precious Little Sleep. Planning to try wake windows method.','Sleep','','Medium'],
    [6,'2025-09-15','Emotional Check-in','Final','Feeling anxious about birth but also so ready. Therapy helping a lot.','Emotional, Mental Health','','High'],
    [7,'2025-10-01','Gratitude','Final','Grateful for Jake, my body, this journey. Counting every kick.','Gratitude','','Low'],
    [8,'2025-10-15','General','Draft','Hospital bag 80% packed. Need to add toiletries and baby coming-home outfit.','Hospital Bag','','High'],
  ];
  vdata.push({ range: `'📝 Notes & Contacts'!A39:H46`, values: notes });

  await valuesBatchUpdate(id, vdata, 'nc-values');
  console.log('Notes & Contacts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
