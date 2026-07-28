'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CRM = sheetMap['Client CRM'];
const S   = "'Client CRM'";

const headers = [
  'Client ID','Client / Company Name','Contact Name','Email','Phone',
  'Billing Address 1','Billing Address 2','City','State / Province','Postal Code',
  'Country','Tax ID','Default Payment Terms','Default Tax Rate',
  'Currency','Preferred Delivery Method','Active?','Notes'
];

const clients = [
  ['','Apex Digital Agency','Priya Nair','priya@apexdigital.com','(415) 555-0241','200 Market St','Ste 800','San Francisco','CA','94103','US','','Net 30',0.0875,'USD','Email',true,'Ongoing branding retainer'],
  ['','BlueSky Architecture','Tom Hargrove','tom@blueskyarch.com','(503) 555-0173','88 NW Davis Ave','','Portland','OR','97209','US','','Net 14',0.0,'USD','Email',true,'Project-based only'],
  ['','Cascade Wellness Spa','Leah Ortega','leah@cascadewellness.com','(206) 555-0318','1450 Pine St','','Seattle','WA','98101','US','','Due on Receipt',0.1,'USD','In Person',true,'Monthly newsletter design'],
  ['','Driftwood Publishing','Carlos Mendez','carlos@driftwoodpub.com','(512) 555-0294','4200 Congress Ave','Unit 2','Austin','TX','78701','US','82-4471029','Net 30',0.0825,'USD','Email',true,'Book cover & editorial work'],
  ['','Emerald Coast Realty','Sandra Kim','sandra@emeraldrealty.com','(850) 555-0162','300 Gulf Shore Blvd','','Destin','FL','32541','US','','Net 15',0.07,'USD','Email',true,'Quarterly marketing package'],
  ['','Flint & Steel Brewing Co.','Jake Tillman','jake@flintsteel.com','(720) 555-0487','2900 Walnut St','','Denver','CO','80205','US','','Net 30',0.029,'USD','Printed',true,'Tap room event materials'],
  ['','Gilded Thread Boutique','Nina Castellano','nina@gildedthread.com','(212) 555-0311','55 Spring St','Floor 3','New York','NY','10012','US','','Due on Receipt',0.08875,'USD','Email',true,'Seasonal lookbook photography'],
  ['','Harbor View Hotels','Daniel Reese','dreese@harborview.com','(617) 555-0229','One Harborside Dr','','Boston','MA','02128','US','HV-550012','Net 45',0.0625,'USD','Email',true,'Annual brand refresh'],
  ['','Ironclad Fitness','Marcus Webb','marcus@ironcladfit.com','(480) 555-0376','7100 E Camelback Rd','','Scottsdale','AZ','85251','US','','Net 7',0.056,'USD','Email',true,'Social media graphic package'],
  ['','Juniper Lane Interiors','Fiona Grant','fiona@juniperlane.com','(615) 555-0198','312 Broadway','Suite 400','Nashville','TN','37201','US','','Net 30',0.09,'USD','Email',true,'Residential staging visuals'],
  ['','Keystone Law Group','Rachel Odum','rodum@keystonelaw.com','(312) 555-0147','225 W Wacker Dr','','Chicago','IL','60606','US','KLG-778821','Net 15',0.01025,'USD','Email',true,'Confidential — invoice only'],
  ['','Luminary Learning Co.','Brian Cho','brian@luminarylearn.com','(408) 555-0263','3000 El Camino Real','','Palo Alto','CA','94306','US','','Net 30',0.0725,'USD','Email',true,'Course content & slide design'],
  ['','Maple & Birch Bakery','Claire Fontaine','claire@maplebakery.com','(651) 555-0192','88 Grand Ave','','St. Paul','MN','55102','US','','Due on Receipt',0.0688,'USD','In Person',true,'Logo refresh + menu design'],
  ['','Northgate Consulting','James Whitfield','james@northgatecon.com','(202) 555-0381','1800 K St NW','Suite 500','Washington','DC','20006','US','NC-334512','Net 30',0.06,'USD','Email',true,'Quarterly report design'],
  ['','Opal Events Co.','Mei-Lin Chu','meilin@opalevents.com','(323) 555-0274','6500 Sunset Blvd','','Los Angeles','CA','90028','US','','Net 14',0.1025,'USD','Email',true,'Wedding stationery + event collateral'],
  ['','Prairie Wind Photography','Tyler Nguyen','tyler@prairiewind.com','(402) 555-0119','4400 Dodge St','','Omaha','NE','68131','US','','Due on Receipt',0.055,'USD','In Person',true,'Editing subscription'],
  ['','Quartz Ridge Analytics','Sasha Volkov','sasha@quartzridge.com','(206) 555-0354','700 5th Ave','Floor 22','Seattle','WA','98104','US','QRA-112983','Net 60',0.0,'USD','Email',true,'Data dashboard consulting'],
  ['','River Oak Counseling','Dr. Amara Singh','amara@riveroakcnslg.com','(817) 555-0288','2100 Park Row','','Fort Worth','TX','76110','US','','Due on Receipt',0.0825,'USD','Email',true,'Brochure + website copy'],
  ['','Sunrise Solar Solutions','Paul Everett','paul@sunrisesolar.com','(602) 555-0341','9100 N Central Ave','','Phoenix','AZ','85020','US','SSS-998823','Net 30',0.056,'USD','Email',true,'Annual report + pitch deck'],
  ['','Tidewater Creative','Gwen Harper','gwen@tidewatercreative.com','(757) 555-0163','220 Granby St','Suite 110','Norfolk','VA','23510','US','','Net 30',0.053,'USD','Email',false,'On pause — follow up Q1 2027'],
];

(async () => {
  const data = [];

  data.push({ range: `${S}!A1`, values: [['CLIENT CRM']] });
  data.push({ range: `${S}!A2`, values: [['Track all client contacts and billing defaults here. Client ID is auto-generated. Active clients appear in invoice dropdowns.']] });
  data.push({ range: `${S}!A5`, values: [headers] });

  // Client ID formulas + data rows
  for (let i = 0; i < clients.length; i++) {
    const r = i + 6;
    data.push({ range: `${S}!A${r}`, values: [[`=IF(B${r}="","","CL-"&TEXT(ROW()-5,"000"))`]] });
    const row = clients[i];
    // Write C through R (indices 1-17 in clients array, cols B-R = cols 1-17)
    data.push({ range: `${S}!B${r}`, values: [[row[1]]] }); // Company Name
    data.push({ range: `${S}!C${r}`, values: [[row[2]]] }); // Contact Name
    data.push({ range: `${S}!D${r}`, values: [[row[3]]] }); // Email
    data.push({ range: `${S}!E${r}`, values: [[row[4]]] }); // Phone
    data.push({ range: `${S}!F${r}`, values: [[row[5]]] }); // Address 1
    data.push({ range: `${S}!G${r}`, values: [[row[6]]] }); // Address 2
    data.push({ range: `${S}!H${r}`, values: [[row[7]]] }); // City
    data.push({ range: `${S}!I${r}`, values: [[row[8]]] }); // State
    data.push({ range: `${S}!J${r}`, values: [[row[9]]] }); // Postal
    data.push({ range: `${S}!K${r}`, values: [[row[10]]] }); // Country
    data.push({ range: `${S}!L${r}`, values: [[row[11]]] }); // Tax ID
    data.push({ range: `${S}!M${r}`, values: [[row[12]]] }); // Payment Terms
    data.push({ range: `${S}!N${r}`, values: [[row[13]]] }); // Tax Rate
    data.push({ range: `${S}!O${r}`, values: [[row[14]]] }); // Currency
    data.push({ range: `${S}!P${r}`, values: [[row[15]]] }); // Delivery Method
    data.push({ range: `${S}!Q${r}`, values: [[row[16]]] }); // Active? (boolean)
    data.push({ range: `${S}!R${r}`, values: [[row[17]]] }); // Notes
  }
  // Formulas for blank rows 26-305
  for (let r = 26; r <= 305; r++) {
    data.push({ range: `${S}!A${r}`, values: [[`=IF(B${r}="","","CL-"&TEXT(ROW()-5,"000"))`]] });
  }

  await valuesBatchUpdate(id, data, 'clientcrm-values');

  const reqs = [];

  // Sheet bg
  reqs.push({ repeatCell: { range: gridRange(CRM, 0, 310, 0, 20), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Title
  reqs.push({ mergeCells: { range: gridRange(CRM, 0, 1, 0, 18), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(CRM, 0, 1, 0, 18), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 16 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: CRM, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  // Subtitle
  reqs.push({ mergeCells: { range: gridRange(CRM, 1, 2, 0, 18), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(CRM, 1, 2, 0, 18), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { italic: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // Column headers row 5
  reqs.push({ repeatCell: { range: gridRange(CRM, 4, 5, 0, 18), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white) },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  // Data rows alternating
  reqs.push({ repeatCell: { range: gridRange(CRM, 5, 305, 0, 18), cell: { userEnteredFormat: { backgroundColor: hex(C.panel) } }, fields: 'userEnteredFormat.backgroundColor' } });
  for (let r = 1; r < 300; r += 2) {
    reqs.push({ repeatCell: { range: gridRange(CRM, 5 + r, 6 + r, 0, 18), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }

  // Formula cells (col A)
  reqs.push({ repeatCell: { range: gridRange(CRM, 5, 305, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Tax Rate column N — percent format
  reqs.push({ repeatCell: { range: gridRange(CRM, 5, 305, 13, 14), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Column widths
  const colWidths = [90, 200, 160, 200, 130, 200, 130, 120, 100, 80, 80, 110, 130, 80, 60, 140, 60, 200];
  colWidths.forEach((px, ci) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: CRM, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  });

  // Freeze rows 1:5 only (merged title row spans A:R, so column freeze is not supported)
  reqs.push({ updateSheetProperties: { properties: { sheetId: CRM, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'clientcrm-format');
  console.log('✓ Client CRM');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
