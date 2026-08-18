'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Milestones & Progress'];
const S = "'Milestones & Progress'";
const PS = "'Project Setup'";
const MTL = "'Master Task Log'";
const NC = 12; // A-L

// Column layout (must match 02-project-setup.js references):
// A(0)=Milestone ID(fml) B(1)=Project ID C(2)=Project Name(fml) D(3)=Milestone Name
// E(4)=Description F(5)=Target Date G(6)=Actual Date H(7)=Status
// I(8)=Pct Complete(fml) J(9)=Priority K(10)=Owner L(11)=Notes
// Data at row 8 (1-indexed) = row 7 (0-indexed)

const M = (projId, name, desc, targetDate, actualDate, status, pri, owner, notes) => {
  const row = Array(NC).fill('');
  row[1]  = projId;
  row[3]  = name;
  row[4]  = desc;
  row[5]  = targetDate;
  row[6]  = actualDate;
  row[7]  = status;
  row[9]  = pri;
  row[10] = owner;
  row[11] = notes;
  return row;
};

const MILESTONES = [
  // PRJ-001 Riverfront Apartment App
  M('PRJ-001','UI Wireframes Approved','All core screens wireframed and approved by client','3/7/2026','3/5/2026','Achieved','High','Marcus Webb','Delivered 2 days early; client signed off 3/5'),
  M('PRJ-001','Backend API Development Complete','All REST endpoints for listings, auth, and search live and documented','7/31/2026','','In Progress','High','Tyler Banks','On track; 60% of endpoints complete as of 8/18'),
  M('PRJ-001','App Launch — iOS and Android','App live on App Store and Google Play, client UAT signed off','9/30/2026','','Not Started','High','Sarah Chen',''),

  // PRJ-002 Annual Marketing Rebrand
  M('PRJ-002','Brand Positioning Framework Approved','Messaging pillars, audience personas, and brand voice signed off','9/26/2026','','Not Started','High','Olivia Park','Requires stakeholder interviews to be complete first'),
  M('PRJ-002','Logo Design Shortlisted','3 logo concepts presented; steering committee selects 1 direction','10/17/2026','','Not Started','High','Marcus Webb',''),
  M('PRJ-002','Brand Guidelines Delivered','Final brand bible document distributed to all teams','12/5/2026','','Not Started','High','Olivia Park','End-of-project deliverable'),

  // PRJ-003 Nova X200 Product Launch
  M('PRJ-003','First Production Run QC Passed','500-unit initial manufacturing run quality-cleared and warehoused','5/31/2026','5/28/2026','Achieved','Critical','Marcus Webb','QC passed 5/28; 3 days ahead of schedule'),
  M('PRJ-003','Website Landing Page Live','Product landing page published with countdown timer and pre-order CTA','5/15/2026','5/15/2026','Achieved','High','Marcus Webb','Launched on schedule; 1,200 email sign-ups in first week'),
  M('PRJ-003','Product Launch Event Held','Hybrid launch event conducted; product available for purchase','9/30/2026','','In Progress','Critical','Olivia Park','Event planning 25% complete; target 500 attendees'),

  // PRJ-004 HR Policy Update 2026
  M('PRJ-004','HR Policy Executive Approval','All 8 updated policies signed off by CEO and COO','4/30/2026','4/29/2026','Achieved','High','Zara Osei','Approved 1 day early; no changes from executive review'),
  M('PRJ-004','Company-Wide Policy Rollout Complete','All 85 employees trained on updated policies','5/30/2026','5/30/2026','Achieved','High','Zara Osei','100% attendance at department training sessions'),

  // PRJ-005 E-commerce Platform Migration
  M('PRJ-005','Product Catalog Migration 75% Complete','9,000 of 12,000 products migrated and QA-approved in Shopify','7/31/2026','','Delayed','High','Tyler Banks','Now 8,900 products — slightly behind; delayed by spec issues'),
  M('PRJ-005','Customer Data Migration Complete','All 8,000 customer accounts and order history migrated','8/15/2026','','Not Started','Critical','Tyler Banks','BLOCKED — data format mismatch unresolved; at risk of delay'),
  M('PRJ-005','E-commerce Platform Go-Live','Zero-downtime cutover complete; store live on Shopify Plus','9/13/2026','','Not Started','Critical','Tyler Banks','Dependent on all migration and QA milestones'),

  // PRJ-006 Q3 Financial Audit
  M('PRJ-006','Revenue Reconciliation Complete','July and August revenue reconciled and variance-free','7/31/2026','7/30/2026','Achieved','High','Sarah Chen','Completed ahead of deadline; no material variances found'),
  M('PRJ-006','Q3 Audit Report Delivered to Board','Final management report and auditor sign-off delivered','9/15/2026','','Not Started','High','Sarah Chen',''),

  // PRJ-007 Employee Wellness Day
  M('PRJ-007','Venue Confirmed and Booked','Lakeside Conference Center secured for Wellness Day event','6/20/2026','6/18/2026','Achieved','Medium','Olivia Park','Secured 2 days early; deposit paid'),

  // PRJ-008 Mobile App Redesign
  M('PRJ-008','Journey Mapping Workshops Complete','3 cross-functional workshops completed; insights synthesized','5/31/2026','5/29/2026','Achieved','High','Marcus Webb','All 3 workshops complete; 47 actionable insights documented'),
  M('PRJ-008','Hi-Fi Wireframes Phase 1 Delivered','22 of 30 core screens in hi-fi wireframes, developer-ready','8/31/2026','','In Progress','High','Marcus Webb','On track; 22 screens done; 8 remaining (nav, search, settings)'),
  M('PRJ-008','Design System Complete','Full 80-component library and token set delivered to engineering','9/15/2026','','In Progress','High','Marcus Webb','50% complete; component documentation in progress'),
  M('PRJ-008','Beta Testing and Iteration Complete','25-user beta panel testing done; all critical issues resolved','11/22/2026','','Not Started','High','Marcus Webb',''),

  // PRJ-009 Client Onboarding System
  M('PRJ-009','Client Portal Live (Beta)','Self-service client portal available for pilot clients','7/31/2026','','In Progress','High','Tyler Banks','75% complete; document upload feature in QA'),
  M('PRJ-009','Document Template Library Complete','All 12 document templates created and available in portal','8/15/2026','','In Progress','Medium','Tyler Banks','3 of 12 templates complete; 9 in draft'),
  M('PRJ-009','Onboarding System Go-Live','All new enterprise clients onboarded through automated system','9/30/2026','','Not Started','High','Tyler Banks',''),

  // PRJ-010 Website SEO Overhaul
  M('PRJ-010','New Site Architecture Approved','Revised silo structure and URL plan approved by marketing and engineering','9/26/2026','','Not Started','High','Zara Osei',''),
  M('PRJ-010','Technical SEO Fixes Complete','All crawl errors, page speed, and Core Web Vital issues resolved','10/31/2026','','Not Started','High','Zara Osei',''),

  // PRJ-012 Online Course Development
  M('PRJ-012','Course Platform Ready','LMS configured; payment, enrollment, and content delivery tested','9/26/2026','','Not Started','High','Sarah Chen',''),
  M('PRJ-012','Module 1 Delivered','First module recorded, edited, and published on LMS','10/31/2026','','Not Started','High','Sarah Chen',''),

  // PRJ-013 Photography Portfolio Site
  M('PRJ-013','Gallery Published','All 50 curated photos uploaded and organized in site gallery','9/20/2026','','Not Started','Medium','Self',''),
  M('PRJ-013','Portfolio Site Launched','Site live; SEO submitted; shared to LinkedIn and photography community','10/5/2026','','Not Started','Medium','Self',''),

  // PRJ-014 Tax Filing 2025
  M('PRJ-014','Tax Returns Filed and Confirmed','Federal and state returns e-filed; IRS acceptance received','4/12/2026','4/12/2026','Achieved','High','Self','Filed 4/10; IRS acceptance confirmation received 4/12'),

  // PRJ-015 Team Training Workshop
  M('PRJ-015','Training Materials Complete','50-page workbook and all facilitation materials finalized','8/29/2026','','In Progress','High','Olivia Park','50% complete; workbook writing in progress'),
  M('PRJ-015','Workshop Delivered','2-day PM skills workshop facilitated for all 12 team leads','9/11/2026','','Not Started','High','Olivia Park',''),

  // PRJ-016 Product Catalog Update
  M('PRJ-016','All Product Photos Complete','89 updated product photos shot, edited, and web-optimized','8/31/2026','','In Progress','High','Tyler Banks','40 of 89 photos complete; studio shoot 8/25-8/27'),
  M('PRJ-016','Full Catalog Live in Store','All 417 products updated with new specs, photos, and pricing','9/30/2026','','Not Started','High','Tyler Banks',''),

  // PRJ-017 Social Media Strategy Q3
  M('PRJ-017','Q3 Paid Campaign Launched','Meta awareness campaign live; 3 ad creatives A/B testing','8/25/2026','','Not Started','High','Zara Osei',''),

  // PRJ-019 Brand Video Production
  M('PRJ-019','Production Complete','All filming days complete; raw footage approved for editing','8/13/2026','','Delayed','Critical','Sarah Chen','DELAYED — location permit issue blocking production days 1 and 2'),

  // PRJ-020 Supplier Contract Renewal
  M('PRJ-020','All Supplier Contracts Signed','6 renewed supplier agreements fully executed and filed','10/27/2026','','Not Started','High','Tyler Banks',''),
];

(async () => {
  const fmt = [];
  const vals = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['MILESTONES & PROGRESS']] });

  // ── Subtitle ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.teal), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [['Track key project milestones. Pct Complete reflects tasks tagged with the Milestone ID in the Master Task Log. Project Setup reads milestone data for portfolio health.']] });

  // ── Instructions ──────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.terra), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A3`, values: [['FORMULA COLUMNS (blue): Milestone ID, Project Name, Pct Complete. Tag tasks in the Master Task Log Milestone ID column to auto-calculate completion %.']] });

  // ── Stats row ─────────────────────────────────────────────────────────────
  const statW = 3;
  const stats = [
    ['Total Milestones', `=COUNTA(D8:D1007)`],
    ['Achieved',         `=COUNTIF(H8:H1007,"Achieved")`],
    ['In Progress',      `=COUNTIF(H8:H1007,"In Progress")`],
    ['Delayed',          `=COUNTIF(H8:H1007,"Delayed")`],
  ];
  stats.forEach(([label, fml], i) => {
    const c = i * statW;
    fmt.push({ mergeCells: { range: gridRange(SID, 3, 4, c, c+statW), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, c, c+statW), cell: { userEnteredFormat: { backgroundColor: hex(C.teal), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ mergeCells: { range: gridRange(SID, 4, 5, c, c+statW), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, c, c+statW), cell: { userEnteredFormat: { backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}4`, values: [[label]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}5`, values: [[fml]] });
  });

  fmt.push({ repeatCell: { range: gridRange(SID, 5, 6, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // ── Column Headers (row 6) ────────────────────────────────────────────────
  const COLS = ['Milestone ID','Project ID','Project Name','Milestone Name','Description',
    'Target Date','Actual Date','Status','Pct Complete','Priority','Owner','Notes'];
  fmt.push({ repeatCell: { range: gridRange(SID, 6, 7, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)' } });
  vals.push({ range: `${S}!A7`, values: [COLS] });

  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Data Validation ───────────────────────────────────────────────────────
  const dvList = (col, items) => ({
    setDataValidation: {
      range: gridRange(SID, 7, 1007, col, col+1),
      rule: { condition: { type: 'ONE_OF_LIST', values: items.map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: false },
    },
  });
  fmt.push(dvList(7, ['Not Started','In Progress','Achieved','Delayed','Cancelled']));
  fmt.push(dvList(9, ['Low','Medium','High','Critical']));

  // ── Formula columns ───────────────────────────────────────────────────────
  const N_ROWS = 1000;
  const idFmls  = Array.from({ length: N_ROWS }, (_, i) => [`=IF(D${8+i}="","","MST-"&TEXT(ROW()-7,"000"))`]);
  const nameFmls = Array.from({ length: N_ROWS }, (_, i) =>
    [`=IFERROR(INDEX(${PS}!$B$8:$B$507,MATCH(B${8+i},${PS}!$A$8:$A$507,0)),"")`]);
  const pctFmls = Array.from({ length: N_ROWS }, (_, i) =>
    [`=IFERROR(COUNTIFS(${MTL}!$W$8:$W$3007,A${8+i},${MTL}!$N$8:$N$3007,"Complete")/COUNTIFS(${MTL}!$W$8:$W$3007,A${8+i}),"")`]);

  vals.push({ range: `${S}!A8:A${7+N_ROWS}`, values: idFmls });
  vals.push({ range: `${S}!C8:C${7+N_ROWS}`, values: nameFmls });
  vals.push({ range: `${S}!I8:I${7+N_ROWS}`, values: pctFmls });

  // ── Number formats ────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 1007, 5, 7), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 1007, 8, 9), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

  [0, 2, 8].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID, 7, 1007, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  for (let r = 0; r < 200; r++) {
    if (r % 2 !== 0) {
      fmt.push({ repeatCell: { range: gridRange(SID, 7+r, 8+r, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
    }
  }

  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  const colWidths = [75,75,150,220,200,90,90,90,70,70,90,200];
  colWidths.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, fmt, 'ms-fmt');
  vals.push({ range: `${S}!A8`, values: MILESTONES });
  await valuesBatchUpdate(id, vals, 'ms-vals');
  console.log(`✓ Milestones & Progress complete — ${MILESTONES.length} milestones written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
