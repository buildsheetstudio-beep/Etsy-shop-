'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const PS = sheetMap['Products & Services'];
const S  = "'Products & Services'";

const headers = [
  'Item ID','Item Name','Item Type','Description','Unit',
  'Default Rate','Tax Treatment','Default Tax Rate','Discount Eligible?','Active?','Notes'
];

// [Name, Type, Description, Unit, Rate, TaxTreatment, TaxRate, DiscountEligible, Active, Notes]
const items = [
  ['Brand Strategy Session','Hourly Work','Discovery and brand positioning workshop','Hour',195,'Taxable',0.0875,true,true,'Rate per billable hour'],
  ['Logo Design — Standard','Service','Primary logo + 2 alternates, vector files','Project',850,'Taxable',0.0875,true,true,'Includes 2 rounds of revisions'],
  ['Logo Design — Premium','Service','Full identity: logo, colors, typography guide','Project',1800,'Taxable',0.0875,true,true,'Includes style guide PDF'],
  ['Business Card Design','Service','Print-ready front/back business card','Each',150,'Taxable',0.0875,true,true,'Print files + digital versions'],
  ['Brochure Design — Trifold','Service','Three-panel print brochure, print-ready files','Each',380,'Taxable',0.0875,true,true,'Up to 2 revisions'],
  ['Social Media Post Pack','Package','20 branded social post templates','Package',420,'Taxable',0.0875,true,true,'Instagram/Facebook/LinkedIn sizes'],
  ['Photography — Half Day','Hourly Work','Commercial photography on location','Hour',225,'Taxable',0.0875,false,true,'4-hour minimum'],
  ['Photography — Full Day','Service','Full-day commercial photo shoot','Day',980,'Taxable',0.0875,false,true,'Up to 8 hours on location'],
  ['Photo Editing','Hourly Work','Retouching and color correction','Hour',85,'Taxable',0.0875,true,true,'Per image rate available on request'],
  ['Website Copywriting','Hourly Work','SEO-optimized website page copy','Hour',120,'Taxable',0.0875,true,true,'Per-page packages available'],
  ['Slide Deck Design','Service','Custom branded presentation design','Project',650,'Taxable',0.0875,true,true,'Up to 30 slides'],
  ['Monthly Retainer — Basic','Retainer','5 hours design/month, priority scheduling','Month',750,'Taxable',0.0875,false,true,'Auto-renews monthly'],
  ['Monthly Retainer — Growth','Retainer','12 hours design + strategy calls','Month',1650,'Taxable',0.0875,false,true,'Includes monthly analytics review'],
  ['Annual Retainer — Premium','Retainer','Full-year partnership: 20hr/mo + priority','Month',2800,'Taxable',0.0875,false,true,'Paid monthly or quarterly'],
  ['Print Management','Service','Vendor coordination for print projects','Project',275,'Non-Taxable',0.0,true,true,'Vendor invoices billed separately'],
  ['Rush Surcharge','Service','Priority turnaround within 48 hours','Each',150,'Taxable',0.0875,false,true,'Added to base project rate'],
  ['Copyright Transfer','Service','Full ownership transfer of deliverables','Each',350,'Non-Taxable',0.0,false,true,'Per project basis'],
  ['Stock Image License','Reimbursement','Licensed stock photography for project','Each',45,'Non-Taxable',0.0,false,true,'Cost passthrough, receipt provided'],
  ['Domain Registration','Reimbursement','Annual domain name registration fee','Each',18,'Non-Taxable',0.0,false,true,'Billed at cost'],
  ['Hosting Setup Fee','Service','Web hosting environment configuration','Each',220,'Taxable',0.0875,true,true,'One-time fee per site'],
  ['SEO Audit','Service','Technical and on-page SEO analysis report','Project',480,'Taxable',0.0875,true,true,'Includes action plan'],
  ['Social Media Management','Retainer','Monthly social posting + engagement','Month',890,'Taxable',0.0875,false,true,'3 platforms, 4 posts/week'],
  ['Email Newsletter Design','Service','Monthly branded email template design','Each',240,'Taxable',0.0875,true,true,'HTML + PDF versions'],
  ['Packaging Design','Service','Product label or box design, print-ready','Each',520,'Taxable',0.0875,true,true,'Physical product brands'],
  ['Training / Coaching Session','Hourly Work','1-on-1 business or design skill coaching','Hour',165,'Taxable',0.0875,true,true,'Zoom or in-person'],
  ['Infographic Design','Service','Custom data visualization or explainer','Each',310,'Taxable',0.0875,true,true,'Up to 2 revisions'],
  ['Consultation Call','Hourly Work','Strategy, feedback, or planning consultation','Hour',95,'Taxable',0.0875,true,true,'Billed in 30-min increments'],
  ['Travel Reimbursement','Reimbursement','Mileage and travel expenses','Mile',0.67,'Non-Taxable',0.0,false,true,'IRS standard mileage rate'],
  ['Shipping & Handling','Shipping','Delivery of physical print materials','Each',18,'Taxable',0.0875,false,true,'Actual carrier cost + handling'],
  ['Early Payment Discount','Discount','Discount for payment within 5 business days','Each',-50,'Non-Taxable',0.0,false,true,'Fixed or % — adjust per invoice'],
];

(async () => {
  const data = [];

  data.push({ range: `${S}!A1`, values: [['PRODUCTS & SERVICES']] });
  data.push({ range: `${S}!A2`, values: [['Maintain your service and product catalog here. Item IDs auto-generate. Use Item ID in Invoice Generator to pre-fill rates and descriptions.']] });
  data.push({ range: `${S}!A5`, values: [headers] });

  for (let i = 0; i < items.length; i++) {
    const r = i + 6;
    data.push({ range: `${S}!A${r}`, values: [[`=IF(B${r}="","","ITEM-"&TEXT(ROW()-5,"000"))`]] });
    const [name, type, desc, unit, rate, taxTreat, taxRate, discElig, active, notes] = items[i];
    data.push({ range: `${S}!B${r}:K${r}`, values: [[name, type, desc, unit, rate, taxTreat, taxRate, discElig, active, notes]] });
  }
  // Blank formula rows
  for (let r = 36; r <= 305; r++) {
    data.push({ range: `${S}!A${r}`, values: [[`=IF(B${r}="","","ITEM-"&TEXT(ROW()-5,"000"))`]] });
  }

  await valuesBatchUpdate(id, data, 'products-values');

  const reqs = [];

  reqs.push({ repeatCell: { range: gridRange(PS, 0, 310, 0, 12), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Title
  reqs.push({ mergeCells: { range: gridRange(PS, 0, 1, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(PS, 0, 1, 0, 11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 16 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: PS, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  reqs.push({ mergeCells: { range: gridRange(PS, 1, 2, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(PS, 1, 2, 0, 11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { italic: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  reqs.push({ repeatCell: { range: gridRange(PS, 4, 5, 0, 11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  reqs.push({ repeatCell: { range: gridRange(PS, 5, 305, 0, 11), cell: { userEnteredFormat: { backgroundColor: hex(C.panel) } }, fields: 'userEnteredFormat.backgroundColor' } });
  for (let r = 1; r < 300; r += 2) {
    reqs.push({ repeatCell: { range: gridRange(PS, 5 + r, 6 + r, 0, 11), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }

  reqs.push({ repeatCell: { range: gridRange(PS, 5, 305, 0, 1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Currency format col F (rate)
  reqs.push({ repeatCell: { range: gridRange(PS, 5, 305, 5, 6), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00;[Red]-$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  // Tax rate col H
  reqs.push({ repeatCell: { range: gridRange(PS, 5, 305, 7, 8), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Column widths: A(90) B(200) C(130) D(240) E(80) F(100) G(110) H(90) I(80) J(60) K(200)
  [90,200,130,240,80,100,110,90,80,60,200].forEach((px, ci) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: PS, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: px }, fields: 'pixelSize' } });
  });

  reqs.push({ updateSheetProperties: { properties: { sheetId: PS, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'products-format');
  console.log('✓ Products & Services');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
