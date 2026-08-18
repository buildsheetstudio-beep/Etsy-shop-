'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Master Task Log'];
const S = "'Master Task Log'";
const PS = "'Project Setup'";
const NC = 26;

// Column layout (0-indexed):
// A(0)=Task ID(fml) B(1)=Project ID C(2)=Project Name(fml) D(3)=Task Type E(4)=Assignee
// F(5)=Task Name G(6)=Priority H(7)=Pct Complete I(8)=Important J(9)=Urgent
// K(10)=Start Date L(11)=Phase M(12)=Due Date N(13)=Status O(14)=Dep Task ID
// P(15)=Est Hours Q(16)=Actual Hours R(17)=Dep Type S(18)=Recurring T(19)=Blocked
// U(20)=Rec Template ID V(21)=Overdue(fml) W(22)=Milestone ID X(23)=Days Remaining(fml) Y(24)=Notes Z(25)=spare
// Data starts at row 8 (1-indexed) = row 7 (0-indexed)

const T = (b,d,e,f,g,h,i,j,k,l,m,n,o='',p=0,q=0,r='None',s=false,t=false,u='',w='',y='') => {
  const row = Array(26).fill('');
  row[1]=b; row[3]=d; row[4]=e; row[5]=f; row[6]=g; row[7]=h;
  row[8]=i; row[9]=j; row[10]=k; row[11]=l; row[12]=m; row[13]=n;
  row[14]=o; row[15]=p; row[16]=q; row[17]=r; row[18]=s; row[19]=t;
  row[20]=u; row[22]=w; row[24]=y;
  return row;
};

const TASKS = [
  // PRJ-001 Riverfront Apartment App
  T('PRJ-001','Meeting','Sarah Chen','Discovery kickoff with Apex Realty team','High','100%',true,true,'1/15/2026','1/16/2026','Phase 1','Complete','',2,2,'None',false,false,'','','Align on scope, timeline, and stakeholder expectations'),
  T('PRJ-001','Deliverable','Sarah Chen','Requirements and feature specification document','High','100%',true,false,'1/17/2026','2/7/2026','Phase 1','Complete','',12,14,'None',false,false,'','','Full functional and non-functional requirements sign-off'),
  T('PRJ-001','Deliverable','Marcus Webb','UI wireframes — all core screens','High','100%',true,false,'2/8/2026','3/7/2026','Phase 1','Complete','',20,18,'None',false,false,'','MST-001','Low-fidelity wireframes reviewed and approved by client'),
  T('PRJ-001','Task','Marcus Webb','Database schema and API design','High','75%',true,false,'3/8/2026','4/15/2026','Phase 2','In Progress','',24,18,'None',false,false,'','','ERD and API contract doc in progress'),
  T('PRJ-001','Task','Tyler Banks','Backend API development — property listings','High','50%',true,true,'4/16/2026','7/31/2026','Phase 2','In Progress','TSK-0004','Finish-to-Start',40,20,'None',false,false,'','','RESTful endpoints for listing CRUD and search'),
  T('PRJ-001','Task','Tyler Banks','Backend API development — user auth','High','50%',false,false,'4/16/2026','7/15/2026','Phase 2','In Progress','',16,8,'None',false,false,'','','JWT-based authentication and role management'),
  T('PRJ-001','Task','Marcus Webb','Frontend integration and UI build','High','25%',true,false,'7/1/2026','8/31/2026','Phase 2','In Progress','TSK-0005','Finish-to-Start',40,10,'None',false,false,'','MST-002','React Native screens wired to API'),
  T('PRJ-001','Task','Sarah Chen','QA testing and bug triage','High','0%',true,true,'9/1/2026','9/20/2026','Phase 3','Not Started','TSK-0007','Finish-to-Start',20,0,'None',false,false,'','','Full regression testing across iOS and Android'),
  T('PRJ-001','Approval','Sarah Chen','Client UAT sign-off and launch approval','High','0%',true,true,'9/21/2026','9/30/2026','Phase 3','Not Started','TSK-0008','Finish-to-Start',8,0,'None',false,false,'','MST-003','Final client sign-off before app store submission'),
  T('PRJ-001','Task','Tyler Banks','App store submission and deployment','High','0%',false,false,'9/28/2026','9/30/2026','Phase 3','Not Started','TSK-0009','Start-to-Start',4,0,'None',false,false,'','','iOS App Store + Google Play submission'),
  T('PRJ-001','Follow-Up','Sarah Chen','Post-launch review meeting with client','Medium','0%',false,false,'10/7/2026','10/7/2026','Phase 3','Not Started','TSK-0009','Finish-to-Start',2,0,'None',false,false,'','','30-day retrospective and roadmap discussion'),
  T('PRJ-001','Meeting','Sarah Chen','Sprint planning — Phase 2 kickoff','Medium','100%',false,false,'4/1/2026','4/1/2026','Phase 2','Complete','',1,1,'None',false,false,'','','Internal sprint planning for backend and frontend phases'),

  // PRJ-002 Annual Marketing Rebrand
  T('PRJ-002','Task','Olivia Park','Brand audit — current assets and perception','Medium','0%',true,false,'8/1/2026','8/22/2026','Discovery','Not Started','',8,0,'None',false,false,'','','Audit logo, colors, typography, tone across all channels'),
  T('PRJ-002','Task','Olivia Park','Competitive landscape analysis','Medium','0%',false,false,'8/1/2026','8/22/2026','Discovery','Not Started','',6,0,'None',false,false,'','','Document competitor brand positioning and design trends'),
  T('PRJ-002','Meeting','Olivia Park','Stakeholder brand interviews — 5 sessions','High','0%',true,false,'8/25/2026','9/5/2026','Discovery','Not Started','TSK-0013','Finish-to-Start',10,0,'None',false,false,'','','1-hour interviews with CEO, Marketing, Sales, Customer Support, Product leads'),
  T('PRJ-002','Deliverable','Olivia Park','Brand positioning and messaging framework','High','0%',true,false,'9/8/2026','9/26/2026','Strategy','Not Started','TSK-0015','Finish-to-Start',12,0,'None',false,false,'','MST-004','Core messaging pillars and audience personas document'),
  T('PRJ-002','Deliverable','Marcus Webb','New logo design — 3 concept directions','High','0%',true,false,'9/29/2026','10/17/2026','Design','Not Started','TSK-0016','Finish-to-Start',20,0,'None',false,false,'','','Presented to steering committee for shortlist selection'),
  T('PRJ-002','Deliverable','Marcus Webb','Color system and typography guidelines','Medium','0%',false,false,'10/20/2026','11/7/2026','Design','Not Started','TSK-0017','Finish-to-Start',10,0,'None',false,false,'','','Primary + secondary palette, font families, usage rules'),
  T('PRJ-002','Deliverable','Olivia Park','Brand guidelines master document','High','0%',true,false,'11/10/2026','12/5/2026','Production','Not Started','TSK-0018','Finish-to-Start',16,0,'None',false,false,'','MST-005','PDF brand bible: logo usage, colors, type, photography, voice'),
  T('PRJ-002','Task','Marcus Webb','Brand asset library creation','Medium','0%',false,false,'12/8/2026','12/26/2026','Production','Backlog','TSK-0019','Finish-to-Start',12,0,'None',false,false,'','','All logo variants, templates, icon sets delivered to cloud drive'),

  // PRJ-003 Nova X200 Product Launch
  T('PRJ-003','Task','Marcus Webb','Product specification finalization','Critical','100%',true,true,'3/1/2026','3/31/2026','Pre-Launch','Complete','',8,8,'None',false,false,'','','Final tech spec approved by engineering and product teams'),
  T('PRJ-003','Task','Marcus Webb','Manufacturing coordination — first production run','Critical','100%',true,false,'3/15/2026','5/31/2026','Pre-Launch','Complete','',16,18,'None',false,false,'','MST-006','500-unit initial run QC passed'),
  T('PRJ-003','Deliverable','Marcus Webb','Product website landing page — design','High','100%',true,false,'4/1/2026','5/15/2026','Pre-Launch','Complete','',24,22,'None',false,false,'','','High-converting landing page with product video and pre-order CTA'),
  T('PRJ-003','Task','Zara Osei','Pre-launch PR campaign — press outreach','High','75%',true,true,'6/1/2026','8/31/2026','Launch','In Progress','',30,22,'None',false,false,'','','TechCrunch, The Verge, Wired — 15 outlets targeted'),
  T('PRJ-003','Task','Zara Osei','Social media content creation — launch series','High','50%',false,false,'7/1/2026','9/15/2026','Launch','In Progress','',20,10,'None',false,false,'','','12-post countdown series across Instagram, Twitter, LinkedIn'),
  T('PRJ-003','Task','Tyler Banks','Distribution partner onboarding — 3 retailers','High','50%',true,false,'6/15/2026','8/31/2026','Launch','In Progress','',24,12,'None',false,false,'','','Amazon, Best Buy, B&H Photo — portal setup and inventory allocation'),
  T('PRJ-003','Task','Olivia Park','Launch event planning — virtual livestream','High','25%',false,false,'7/15/2026','9/30/2026','Launch','In Progress','',20,5,'None',false,false,'','MST-007','Hybrid launch event targeting 500 registrations'),
  T('PRJ-003','Task','Tyler Banks','Inventory fulfillment and warehouse setup','High','0%',true,true,'8/15/2026','9/20/2026','Launch','Not Started','TSK-0026','Finish-to-Start',12,0,'None',false,false,'','','3PL warehouse configured for launch day dispatch'),
  T('PRJ-003','Task','Marcus Webb','Post-launch analytics and reporting setup','Medium','0%',false,false,'9/1/2026','10/1/2026','Post-Launch','Not Started','',6,0,'None',false,false,'','','GA4, Mixpanel, and ad attribution configured'),
  T('PRJ-003','Deliverable','Zara Osei','Customer support documentation and FAQs','Medium','0%',false,false,'8/15/2026','9/30/2026','Post-Launch','Not Started','',8,0,'None',false,false,'','','Knowledge base articles and support ticket templates'),
  T('PRJ-003','Follow-Up','Marcus Webb','30-day post-launch retrospective','Medium','0%',false,false,'10/15/2026','10/15/2026','Post-Launch','Not Started','TSK-0029','Finish-to-Start',3,0,'None',false,false,'','','Sales, returns, support volume review vs targets'),
  T('PRJ-003','Task','Tyler Banks','Retailer co-op marketing submissions','Medium','0%',false,false,'8/1/2026','9/15/2026','Launch','Not Started','',6,0,'None',false,false,'','','Submit co-op materials to all 3 retail partners'),

  // PRJ-004 HR Policy Update 2026
  T('PRJ-004','Task','Zara Osei','Policy gap analysis against 2025 legislation','High','100%',true,true,'1/10/2026','1/31/2026','Review','Complete','',12,11,'None',false,false,'','','Identified 8 policies requiring substantive updates'),
  T('PRJ-004','Task','Zara Osei','Legal compliance review with external counsel','High','100%',true,false,'2/1/2026','2/28/2026','Review','Complete','',8,9,'None',false,false,'','','Counsel confirmed compliance alignment for all 8 policies'),
  T('PRJ-004','Deliverable','Zara Osei','Updated PTO and leave policy','High','100%',true,false,'3/1/2026','3/22/2026','Drafting','Complete','',6,6,'None',false,false,'','','Expanded sick leave, bereavement, and parental leave'),
  T('PRJ-004','Deliverable','Zara Osei','Updated remote and hybrid work policy','High','100%',false,false,'3/1/2026','3/22/2026','Drafting','Complete','',6,5,'None',false,false,'','','Clear guidelines on WFH eligibility, home office stipend'),
  T('PRJ-004','Deliverable','Zara Osei','Updated DEI and anti-harassment policy','High','100%',true,false,'3/1/2026','3/22/2026','Drafting','Complete','',8,8,'None',false,false,'','','Expanded reporting procedures and training requirements'),
  T('PRJ-004','Approval','Zara Osei','Internal stakeholder review — HR and managers','High','100%',true,false,'3/25/2026','4/15/2026','Review','Complete','TSK-0039','Finish-to-Start',6,7,'None',false,false,'','','Distributed draft policies to 14 department heads for feedback'),
  T('PRJ-004','Approval','Zara Osei','Executive leadership sign-off','High','100%',true,true,'4/16/2026','4/30/2026','Approval','Complete','TSK-0041','Finish-to-Start',4,4,'None',false,false,'','MST-008','CEO and COO approved all 8 updated policies'),
  T('PRJ-004','Task','Zara Osei','Company-wide policy rollout and training','High','100%',true,true,'5/1/2026','5/30/2026','Rollout','Complete','TSK-0042','Finish-to-Start',12,12,'None',false,false,'','','All-hands presentation + department-specific training sessions'),

  // PRJ-005 E-commerce Platform Migration
  T('PRJ-005','Task','Tyler Banks','Legacy platform data export and mapping','High','100%',true,true,'5/1/2026','5/31/2026','Discovery','Complete','',16,18,'None',false,false,'','','12,000 products and 8,000 customer records exported'),
  T('PRJ-005','Task','Tyler Banks','Shopify Plus environment setup and theme config','High','100%',true,false,'5/15/2026','6/15/2026','Build','Complete','',24,26,'None',false,false,'','','Theme selected, brand colors applied, navigation configured'),
  T('PRJ-005','Task','Tyler Banks','Product catalog migration — all 12,000 SKUs','High','75%',true,true,'6/1/2026','7/31/2026','Build','In Progress','TSK-0046','Finish-to-Start',40,30,'None',false,false,'','MST-009','8,900 of 12,000 products migrated and QA-checked'),
  T('PRJ-005','Task','Tyler Banks','Customer data and order history migration','Critical','25%',true,true,'7/1/2026','8/15/2026','Build','Blocked','TSK-0047','Finish-to-Start',32,8,'None',false,true,'','','BLOCKED: data format mismatch with Shopify customer import API'),
  T('PRJ-005','Task','Tyler Banks','Payment gateway integration — Stripe and PayPal','Critical','0%',true,true,'8/1/2026','8/25/2026','Build','Not Started','TSK-0048','Finish-to-Start',12,0,'None',false,false,'','','Multi-currency and regional tax configuration required'),
  T('PRJ-005','Task','Tyler Banks','Custom checkout and discount engine configuration','High','0%',false,false,'8/20/2026','9/5/2026','Build','Not Started','TSK-0049','Finish-to-Start',16,0,'None',false,false,'','','Loyalty program rules and bulk-order discounts'),
  T('PRJ-005','Task','Tyler Banks','Staging environment QA — full regression','High','0%',true,true,'9/1/2026','9/10/2026','QA','Not Started','TSK-0050','Finish-to-Start',20,0,'None',false,false,'','','100-case test plan across product, cart, checkout, account flows'),
  T('PRJ-005','Task','Olivia Park','Team training on Shopify Plus admin','Medium','0%',false,false,'9/5/2026','9/12/2026','Launch','Not Started','TSK-0050','Start-to-Start',8,0,'None',false,false,'','','Operations, marketing, and support teams trained'),
  T('PRJ-005','Milestone Task','Tyler Banks','Go-live cutover — zero-downtime migration','Critical','0%',true,true,'9/13/2026','9/13/2026','Launch','Not Started','TSK-0052','Finish-to-Start',6,0,'None',false,false,'','MST-010','Planned for Sunday 2am — 2-hour maintenance window'),

  // PRJ-006 Q3 Financial Audit
  T('PRJ-006','Task','Sarah Chen','Revenue reconciliation — July and August','High','100%',true,true,'7/1/2026','7/31/2026','Review','Complete','',12,11,'None',false,false,'','','Gross and net revenue reconciled against bank statements'),
  T('PRJ-006','Task','Sarah Chen','Accounts payable detailed review','High','75%',true,false,'7/15/2026','8/22/2026','Review','In Progress','',10,7,'None',false,false,'','','All vendor invoices cross-checked against POs and receipts'),
  T('PRJ-006','Task','Sarah Chen','Accounts receivable aging analysis','Medium','50%',false,false,'7/20/2026','8/31/2026','Review','In Progress','',8,4,'None',false,false,'','','Identify invoices 30, 60, 90+ days outstanding'),
  T('PRJ-006','Task','Sarah Chen','Expense categorization and policy compliance audit','Medium','25%',false,false,'8/1/2026','8/31/2026','Review','In Progress','',8,2,'None',false,false,'','','Verify all expenses meet policy limits and categories'),
  T('PRJ-006','Task','Sarah Chen','Variance analysis — budget vs actuals Q3','High','0%',true,false,'9/1/2026','9/5/2026','Analysis','Not Started','TSK-0059','Finish-to-Start',10,0,'None',false,false,'','','Line-by-line budget-to-actual comparison with commentary'),
  T('PRJ-006','Meeting','Sarah Chen','External auditor review session','High','0%',true,true,'9/5/2026','9/10/2026','Review','Not Started','TSK-0063','Start-to-Start',4,0,'None',false,false,'','','Present audit workpapers to external auditors'),
  T('PRJ-006','Deliverable','Sarah Chen','Q3 management financial report draft','High','0%',true,false,'9/8/2026','9/12/2026','Reporting','Not Started','TSK-0063','Finish-to-Start',8,0,'None',false,false,'','MST-011','Report for board and executive team distribution'),
  T('PRJ-006','Approval','Sarah Chen','Final audit sign-off and filing','High','0%',true,true,'9/13/2026','9/15/2026','Reporting','Not Started','TSK-0065','Finish-to-Start',2,0,'None',false,false,'','','CFO and external auditor joint sign-off'),

  // PRJ-007 Employee Wellness Day
  T('PRJ-007','Task','Olivia Park','Venue research and selection','Medium','100%',false,false,'6/1/2026','6/20/2026','Planning','Complete','',6,6,'None',false,false,'','','Evaluated 5 venues; Lakeside Conference Center selected'),
  T('PRJ-007','Task','Olivia Park','Activity and workshop planning','Medium','50%',false,false,'6/21/2026','7/31/2026','Planning','In Progress','',8,4,'None',false,false,'','','Yoga, mindfulness workshop, nutrition talk, team cooking activity'),
  T('PRJ-007','Task','Olivia Park','Catering and nutrition session vendor booking','Medium','0%',false,false,'8/1/2026','8/22/2026','Planning','Waiting','TSK-0069','Finish-to-Start',4,0,'None',false,false,'','','ON HOLD — awaiting budget approval to proceed'),
  T('PRJ-007','Task','Olivia Park','External wellness facilitator booking','Medium','0%',false,false,'8/1/2026','8/22/2026','Planning','Waiting','TSK-0069','Start-to-Start',4,0,'None',false,false,'','','Same hold — need confirmed headcount first'),
  T('PRJ-007','Task','Olivia Park','Staff communication and registration','Medium','0%',false,false,'8/25/2026','9/5/2026','Launch','Not Started','TSK-0071','Finish-to-Start',4,0,'None',false,false,'','','Email invite + RSVP form to all 85 employees'),
  T('PRJ-007','Task','Olivia Park','Day-of coordination plan and run sheet','Low','0%',false,false,'9/10/2026','9/15/2026','Launch','Not Started','TSK-0072','Finish-to-Start',3,0,'None',false,false,'','','Minute-by-minute schedule and vendor arrival coordination'),
  T('PRJ-007','Task','Olivia Park','Post-event survey and impact report','Low','0%',false,false,'9/21/2026','9/28/2026','Post-Event','Backlog','',2,0,'None',false,false,'','','Measure satisfaction, NPS, and wellness engagement score'),

  // PRJ-008 Mobile App Redesign
  T('PRJ-008','Task','Marcus Webb','User research — interviews and surveys','High','100%',true,false,'4/1/2026','4/30/2026','Discovery','Complete','',24,26,'None',false,false,'','','15 user interviews + 200-response survey; report delivered'),
  T('PRJ-008','Task','Marcus Webb','UX and usability audit of current app','High','100%',true,false,'4/15/2026','5/15/2026','Discovery','Complete','',16,15,'None',false,false,'','','Identified 47 usability issues; severity ranked'),
  T('PRJ-008','Task','Marcus Webb','Journey mapping workshops — 3 sessions','High','100%',false,false,'5/1/2026','5/31/2026','Discovery','Complete','',12,12,'None',false,false,'','MST-012','3 workshops with product, engineering, and CS teams'),
  T('PRJ-008','Deliverable','Marcus Webb','New information architecture and site map','High','100%',true,false,'6/1/2026','6/28/2026','Design','Complete','TSK-0078','Finish-to-Start',20,19,'None',false,false,'','','Card-sort validated IA approved by product team'),
  T('PRJ-008','Deliverable','Marcus Webb','Hi-fi wireframes — core screens (30 screens)','High','75%',true,true,'7/1/2026','8/31/2026','Design','In Progress','TSK-0081','Finish-to-Start',40,30,'None',false,false,'','MST-013','22 of 30 screens complete; home, search, product detail done'),
  T('PRJ-008','Deliverable','Marcus Webb','Design system — component library','High','50%',false,false,'7/15/2026','9/15/2026','Design','In Progress','TSK-0081','Start-to-Start',30,15,'None',false,false,'','','80 components in Figma; tokens defined; dev handoff pending'),
  T('PRJ-008','Deliverable','Marcus Webb','Developer handoff documentation','High','25%',false,false,'8/1/2026','9/30/2026','Design','In Progress','TSK-0083','Start-to-Start',12,3,'None',false,false,'','','Zeplin annotations and spec sheets for all delivered screens'),
  T('PRJ-008','Task','Marcus Webb','Onboarding flow redesign','High','0%',false,false,'9/1/2026','10/15/2026','Design','Not Started','TSK-0082','Finish-to-Start',20,0,'None',false,false,'','','5-step onboarding with contextual walkthroughs'),
  T('PRJ-008','Task','Marcus Webb','Navigation and search redesign','High','0%',false,false,'9/15/2026','10/31/2026','Design','Not Started','TSK-0082','Start-to-Start',16,0,'None',false,false,'','','Bottom tab bar and global search UX overhaul'),
  T('PRJ-008','Task','Marcus Webb','Accessibility audit and remediation','Medium','0%',false,false,'10/1/2026','10/31/2026','QA','Not Started','TSK-0085','Finish-to-Start',12,0,'None',false,false,'','','WCAG 2.1 AA compliance across all redesigned screens'),
  T('PRJ-008','Task','Marcus Webb','Beta testing with 25-user panel','High','0%',true,false,'11/1/2026','11/22/2026','QA','Backlog','TSK-0088','Finish-to-Start',20,0,'None',false,false,'','MST-014','Moderated usability sessions + unmoderated remote testing'),
  T('PRJ-008','Task','Marcus Webb','Design iteration and final handoff','High','0%',true,false,'11/23/2026','11/30/2026','Delivery','Backlog','TSK-0089','Finish-to-Start',16,0,'None',false,false,'','','Incorporate beta feedback; deliver final assets to engineering'),

  // PRJ-009 Client Onboarding System
  T('PRJ-009','Task','Tyler Banks','As-is process mapping and gap analysis','High','100%',true,false,'2/15/2026','3/7/2026','Discovery','Complete','',16,16,'None',false,false,'','','Current 14-step manual process documented; 8 automation gaps found'),
  T('PRJ-009','Task','Tyler Banks','Requirements gathering with account management team','High','100%',true,false,'3/8/2026','3/28/2026','Discovery','Complete','',10,10,'None',false,false,'','','5 AM stakeholders aligned on priority requirements'),
  T('PRJ-009','Task','Tyler Banks','Workflow automation tool configuration','High','100%',true,false,'4/1/2026','5/15/2026','Build','Complete','TSK-0093','Finish-to-Start',30,32,'None',false,false,'','','HubSpot workflows + Zapier automations for 12 trigger events'),
  T('PRJ-009','Task','Tyler Banks','Client portal setup and branding','High','75%',true,false,'5/1/2026','7/31/2026','Build','In Progress','',20,15,'None',false,false,'','MST-015','Self-service portal with document upload, status tracker, chat'),
  T('PRJ-009','Task','Tyler Banks','Onboarding email sequence — 8-touch series','High','50%',false,false,'6/1/2026','7/31/2026','Build','In Progress','',12,6,'None',false,false,'','','Welcome through 30-day check-in automated sequence'),
  T('PRJ-009','Deliverable','Tyler Banks','Document template library — 12 templates','Medium','25%',false,false,'6/15/2026','8/15/2026','Build','In Progress','',10,2,'None',false,false,'','','NDAs, SOWs, project charters, kick-off agendas'),
  T('PRJ-009','Task','Tyler Banks','CRM integration — HubSpot sync','High','0%',true,false,'8/1/2026','8/31/2026','Build','Not Started','TSK-0097','Finish-to-Start',16,0,'None',false,false,'','','Bi-directional sync of deal stages and onboarding status'),
  T('PRJ-009','Task','Tyler Banks','Pilot with 3 new enterprise clients','High','0%',true,true,'9/1/2026','9/15/2026','Testing','Not Started','TSK-0099','Finish-to-Start',12,0,'None',false,false,'','','End-to-end pilot with Vertex Solutions and 2 TBD clients'),
  T('PRJ-009','Task','Tyler Banks','Account manager training and playbook','Medium','0%',false,false,'9/5/2026','9/20/2026','Launch','Not Started','TSK-0099','Start-to-Start',8,0,'None',false,false,'','','Training session + 30-page playbook with screenshots'),
  T('PRJ-009','Milestone Task','Tyler Banks','System go-live — all new clients on new flow','High','0%',true,true,'9/25/2026','9/30/2026','Launch','Not Started','TSK-0101','Finish-to-Start',4,0,'None',false,false,'','MST-016','All new enterprise clients onboarded through new system'),
  T('PRJ-009','Follow-Up','Tyler Banks','45-day post-launch review with AM team','Medium','0%',false,false,'11/14/2026','11/14/2026','Post-Launch','Backlog','TSK-0102','Finish-to-Start',2,0,'None',false,false,'','','Review time-to-onboard KPI and client satisfaction scores'),

  // PRJ-010 Website SEO Overhaul
  T('PRJ-010','Task','Zara Osei','Technical SEO audit — crawl and index issues','High','0%',true,false,'8/15/2026','9/5/2026','Discovery','Not Started','',16,0,'None',false,false,'','','Screaming Frog audit; 404s, redirect chains, canonical issues'),
  T('PRJ-010','Task','Zara Osei','Keyword research — core and long-tail','High','0%',true,false,'8/15/2026','9/5/2026','Discovery','Not Started','',12,0,'None',false,false,'','','SEMrush: map 200 target keywords to site structure'),
  T('PRJ-010','Task','Zara Osei','Competitor backlink and authority analysis','Medium','0%',false,false,'8/20/2026','9/10/2026','Discovery','Not Started','',8,0,'None',false,false,'','','Ahrefs: 3 competitor gap analyses; identify 50 link opportunities'),
  T('PRJ-010','Deliverable','Zara Osei','Revised site architecture and URL structure plan','High','0%',true,false,'9/8/2026','9/26/2026','Strategy','Not Started','TSK-0105','Finish-to-Start',12,0,'None',false,false,'','MST-017','New silo structure to improve topical authority'),
  T('PRJ-010','Deliverable','Zara Osei','On-page optimization guide — all 80 pages','High','0%',false,false,'9/29/2026','10/17/2026','Optimization','Not Started','TSK-0108','Finish-to-Start',20,0,'None',false,false,'','','Title tags, meta descriptions, H1/H2 rewrite specs'),
  T('PRJ-010','Task','Zara Osei','Technical SEO fixes implementation','High','0%',true,false,'10/1/2026','10/31/2026','Optimization','Not Started','TSK-0106','Finish-to-Start',16,0,'None',false,false,'','','Dev team to implement: schema, page speed, Core Web Vitals'),
  T('PRJ-010','Task','Zara Osei','Content refresh — top 20 underperforming pages','Medium','0%',false,false,'10/20/2026','11/15/2026','Optimization','Not Started','TSK-0109','Start-to-Start',20,0,'None',false,false,'','','Expand thin content; add FAQ sections; improve internal linking'),
  T('PRJ-010','Task','Zara Osei','Link building outreach — 50 prospects','Medium','0%',false,false,'11/1/2026','12/5/2026','Off-Page','Backlog','TSK-0110','Start-to-Start',16,0,'None',false,false,'','','Guest post pitches + resource page link acquisition'),

  // PRJ-011 Home Office Renovation
  T('PRJ-011','Task','Self','Measure space and plan desk/shelf layout','Low','100%',false,false,'7/20/2026','7/22/2026','Planning','Complete','',2,2,'None',false,false,'','','Floor plan drawn; electrical outlet positions mapped'),
  T('PRJ-011','Task','Self','Order motorized standing desk — Uplift V2','Medium','100%',false,false,'7/23/2026','7/25/2026','Procurement','Complete','',1,1,'None',false,false,'','','Order placed; delivery 7/31'),
  T('PRJ-011','Task','Self','Order monitor arms, webcam, and accessories','Medium','50%',false,false,'7/26/2026','8/10/2026','Procurement','In Progress','TSK-0114','Finish-to-Start',2,1,'None',false,false,'','','Ergotron LX arms x2; Logitech Brio 4K; cable management kit'),
  T('PRJ-011','Task','Self','Desk assembly and initial setup','Medium','0%',false,false,'8/5/2026','8/9/2026','Build','Not Started','TSK-0114','Finish-to-Start',3,0,'None',false,false,'','','Desk delivered 7/31; need a helper for assembly'),
  T('PRJ-011','Task','Self','Cable management installation','Low','0%',false,false,'8/10/2026','8/20/2026','Build','Not Started','TSK-0116','Finish-to-Start',4,0,'None',false,false,'','','Under-desk trays, wall raceways, and power strip mounting'),
  T('PRJ-011','Task','Self','New LED lighting installation','Low','0%',false,false,'8/15/2026','8/30/2026','Build','Not Started','TSK-0116','Start-to-Start',3,0,'None',false,false,'','','Bias lighting behind monitor + Elgato Key Light for video calls'),
  T('PRJ-011','Task','Self','Floating shelving installation and organization','Low','0%',false,false,'8/25/2026','9/10/2026','Build','Not Started','TSK-0118','Start-to-Start',4,0,'None',false,false,'','','3 IKEA Lack shelves; books, plants, reference materials'),

  // PRJ-012 Online Course Development
  T('PRJ-012','Task','Sarah Chen','Course curriculum outline — all 12 modules','High','0%',true,false,'9/1/2026','9/20/2026','Planning','Not Started','',16,0,'None',false,false,'','','Module topics, duration, learning outcomes, and assessments'),
  T('PRJ-012','Task','Sarah Chen','Learning objectives — 3 per module','High','0%',false,false,'9/1/2026','9/20/2026','Planning','Not Started','',8,0,'None',false,false,'','',"Bloom's taxonomy aligned; measurable outcomes defined"),
  T('PRJ-012','Task','Sarah Chen','LMS platform selection and setup','Medium','0%',false,false,'9/8/2026','9/26/2026','Planning','Not Started','',6,0,'None',false,false,'','','Evaluate Teachable, Thinkific, Kajabi — select and configure'),
  T('PRJ-012','Deliverable','Sarah Chen','Module 1 script and workbook','High','0%',true,false,'9/22/2026','10/10/2026','Production','Not Started','TSK-0121','Finish-to-Start',20,0,'None',false,false,'','MST-018','Project management fundamentals — 45-min video + 12-page workbook'),
  T('PRJ-012','Task','Sarah Chen','Slide deck design template creation','Medium','0%',false,false,'9/22/2026','10/3/2026','Production','Not Started','TSK-0121','Start-to-Start',6,0,'None',false,false,'','','Google Slides master template with brand styling'),
  T('PRJ-012','Task','Sarah Chen','Recording studio and equipment setup','Medium','0%',false,false,'9/15/2026','9/30/2026','Production','Not Started','',4,0,'None',false,false,'','','Microphone, lighting, teleprompter, green screen calibration'),
  T('PRJ-012','Task','Sarah Chen','Module 1 recording and video editing','High','0%',true,false,'10/13/2026','10/31/2026','Production','Backlog','TSK-0125','Finish-to-Start',12,0,'None',false,false,'','','Record 5 video segments; edit in DaVinci Resolve'),
  T('PRJ-012','Task','Sarah Chen','Beta cohort recruitment — 20 students','Medium','0%',false,false,'11/1/2026','11/30/2026','Launch','Backlog','',6,0,'None',false,false,'','','LinkedIn and email outreach to PM community connections'),

  // PRJ-013 Photography Portfolio Site
  T('PRJ-013','Task','Self','Select and cull best 50 portfolio images','Medium','100%',false,false,'7/1/2026','7/15/2026','Planning','Complete','',6,7,'None',false,false,'','','Culled from 2,000+ images; 50 final selects across 6 categories'),
  T('PRJ-013','Task','Self','Domain registration and web hosting setup','Medium','100%',false,false,'7/1/2026','7/3/2026','Planning','Complete','',1,1,'None',false,false,'','','sarahmcphotography.com registered; Squarespace Business plan'),
  T('PRJ-013','Task','Self','Site template selection and customization','Medium','50%',false,false,'7/5/2026','8/5/2026','Build','In Progress','TSK-0131','Finish-to-Start',6,3,'None',false,false,'','','Brine template selected; fonts and colors customized to brand'),
  T('PRJ-013','Deliverable','Self','Bio and about page copy','Low','25%',false,false,'7/20/2026','8/15/2026','Build','In Progress','',4,1,'None',false,false,'','','First-person narrative; services list; client quote placeholder'),
  T('PRJ-013','Task','Self','Photo editing — color grade and export 50 images','Medium','0%',false,false,'8/10/2026','8/31/2026','Build','Not Started','TSK-0132','Finish-to-Start',10,0,'None',false,false,'','','Lightroom edit; JPEG export at 2048px; web-optimized'),
  T('PRJ-013','Task','Self','Gallery upload and page layout','Medium','0%',false,false,'9/1/2026','9/20/2026','Build','Not Started','TSK-0135','Finish-to-Start',6,0,'None',false,false,'','','6 gallery categories; titles, captions, and ordering'),
  T('PRJ-013','Task','Self','SEO meta, speed optimization, and site launch','Medium','0%',false,false,'9/21/2026','10/5/2026','Launch','Backlog','TSK-0136','Finish-to-Start',4,0,'None',false,false,'','MST-019','Submit to search; image compression; share to LinkedIn'),

  // PRJ-014 Tax Filing 2025
  T('PRJ-014','Task','Self','Collect all W-2 and 1099 income documents','High','100%',true,true,'1/3/2026','1/31/2026','Preparation','Complete','',3,3,'None',false,false,'','','Employer W-2, freelance 1099-NEC, dividend 1099-DIV collected'),
  T('PRJ-014','Task','Self','Bank and investment statement collection','High','100%',true,false,'1/3/2026','1/31/2026','Preparation','Complete','',2,2,'None',false,false,'','','All 2025 statements downloaded; brokerage 1099-B included'),
  T('PRJ-014','Task','Self','Organize and categorize business expense receipts','High','100%',true,true,'1/15/2026','2/28/2026','Preparation','Complete','',6,7,'None',false,false,'','','12 expense categories; QuickBooks export and reconciliation'),
  T('PRJ-014','Meeting','Self','CPA consultation and tax strategy review','High','100%',true,false,'3/1/2026','3/1/2026','Filing','Complete','',2,2,'None',false,false,'','','Reviewed estimated vs actual liability; confirmed filing strategy'),
  T('PRJ-014','Task','Self','Federal income tax return preparation','High','100%',true,true,'3/5/2026','3/31/2026','Filing','Complete','TSK-0143','Finish-to-Start',4,4,'None',false,false,'','','Reviewed by CPA; refund of $2,140 expected'),
  T('PRJ-014','Task','Self','State income tax return preparation','High','100%',true,false,'3/5/2026','3/31/2026','Filing','Complete','TSK-0143','Start-to-Start',2,2,'None',false,false,'','','State return matched federal AGI; small refund'),
  T('PRJ-014','Task','Self','File returns and confirm IRS acceptance','High','100%',true,true,'4/1/2026','4/12/2026','Filing','Complete','TSK-0145','Finish-to-Start',1,1,'None',false,false,'','','E-filed 4/10; IRS acceptance received 4/12 ✓'),

  // PRJ-015 Team Training Workshop
  T('PRJ-015','Task','Olivia Park','Define learning objectives and success metrics','High','100%',true,false,'8/1/2026','8/8/2026','Design','Complete','',6,6,'None',false,false,'','','5 skill outcomes defined; pre/post self-assessment designed'),
  T('PRJ-015','Task','Olivia Park','Design facilitation agenda — 2-day schedule','High','100%',true,false,'8/5/2026','8/15/2026','Design','Complete','TSK-0149','Start-to-Start',8,8,'None',false,false,'','','Minute-by-minute agenda with activities and break structure'),
  T('PRJ-015','Deliverable','Olivia Park','Workbook and participant materials','High','50%',true,false,'8/10/2026','8/29/2026','Design','In Progress','',12,6,'None',false,false,'','','50-page illustrated workbook with exercises and templates'),
  T('PRJ-015','Task','Olivia Park','Venue and AV confirmation — hotel conference room','Medium','75%',false,false,'8/1/2026','8/20/2026','Logistics','In Progress','',3,2,'None',false,false,'','','Marriott downtown; catering, projector, whiteboard confirmed'),
  T('PRJ-015','Task','Olivia Park','Pre-workshop participant survey','Medium','0%',false,false,'9/1/2026','9/4/2026','Pre-Work','Not Started','TSK-0151','Finish-to-Start',2,0,'None',false,false,'','','Identify skill gaps and expectations from 12 attendees'),
  T('PRJ-015','Task','Olivia Park','Workshop facilitation — Day 1 and Day 2','High','0%',true,true,'9/10/2026','9/11/2026','Delivery','Not Started','TSK-0153','Finish-to-Start',16,0,'None',false,false,'','MST-020','Full 2-day PM skills workshop for 12 team leads'),
  T('PRJ-015','Task','Olivia Park','Post-workshop action plan follow-up','Medium','0%',false,false,'9/18/2026','9/25/2026','Follow-Up','Not Started','TSK-0155','Finish-to-Start',4,0,'None',false,false,'','','Each participant submits 30-day personal development action plan'),
  T('PRJ-015','Deliverable','Olivia Park','Workshop impact assessment report','Medium','0%',false,false,'9/26/2026','10/3/2026','Follow-Up','Backlog','TSK-0156','Finish-to-Start',4,0,'None',false,false,'','','Pre/post skill scores; NPS; manager 30-day feedback collected'),

  // PRJ-016 Product Catalog Update
  T('PRJ-016','Task','Tyler Banks','Catalog gap audit — specs, images, pricing','Medium','100%',false,false,'7/15/2026','7/31/2026','Discovery','Complete','',8,8,'None',false,false,'','','417 products audited; 203 need spec update; 89 need new photos'),
  T('PRJ-016','Task','Tyler Banks','Vendor spec sheet collection — all 23 vendors','High','75%',true,false,'8/1/2026','8/22/2026','Production','In Progress','',10,7,'None',false,false,'','','17 of 23 vendors responded; 6 follow-ups pending'),
  T('PRJ-016','Task','Tyler Banks','Product photography coordination — 89 items','High','50%',false,false,'8/1/2026','8/31/2026','Production','In Progress','',20,10,'None',false,false,'','','Studio shoot scheduled 8/25-8/27; 40 products already shot'),
  T('PRJ-016','Task','Tyler Banks','Updated product description writing — 203 items','High','25%',false,false,'8/5/2026','9/10/2026','Production','In Progress','',30,7,'None',false,false,'','','SEO-optimized; feature bullets; compatibility notes'),
  T('PRJ-016','Task','Tyler Banks','Pricing review and competitive benchmarking','Medium','0%',false,false,'8/20/2026','9/5/2026','Production','Not Started','',6,0,'None',false,false,'','','Review pricing on 417 products vs 3 main competitors'),
  T('PRJ-016','Task','Tyler Banks','Bulk upload to e-commerce platform','High','0%',true,false,'9/8/2026','9/22/2026','Launch','Not Started','TSK-0163','Finish-to-Start',8,0,'None',false,false,'','','CSV import into Shopify; validate all fields post-upload'),
  T('PRJ-016','Task','Tyler Banks','QA review of all updated live listings','High','0%',true,false,'9/22/2026','9/26/2026','Launch','Not Started','TSK-0165','Finish-to-Start',12,0,'None',false,false,'','','Sample-check 10% of listings; verify images, specs, pricing'),
  T('PRJ-016','Task','Tyler Banks','Category and navigation taxonomy update','Medium','0%',false,false,'9/15/2026','9/25/2026','Launch','Not Started','',6,0,'None',false,false,'','','Restructure 12 top-level categories to match new brand taxonomy'),
  T('PRJ-016','Task','Tyler Banks','Customer announcement and email campaign','Low','0%',false,false,'9/27/2026','9/30/2026','Launch','Backlog','TSK-0167','Finish-to-Start',3,0,'None',false,false,'','','Announce catalog refresh to 22k email subscribers'),
  T('PRJ-016','Follow-Up','Tyler Banks','Vendor feedback on updated listings','Low','0%',false,false,'10/10/2026','10/10/2026','Post-Launch','Backlog','TSK-0168','Finish-to-Start',1,0,'None',false,false,'','','Share updated listing links with each vendor for verification'),

  // PRJ-017 Social Media Strategy Q3
  T('PRJ-017','Task','Zara Osei','Q3 content calendar creation — all channels','High','100%',true,false,'7/1/2026','7/15/2026','Strategy','Complete','',8,8,'None',true,false,'REC-003','','Instagram, LinkedIn, Twitter/X, TikTok content mapped'),
  T('PRJ-017','Task','Zara Osei','Brand voice and tone guidelines for social','High','100%',true,false,'7/1/2026','7/15/2026','Strategy','Complete','',6,6,'None',false,false,'','','2-page quick reference guide distributed to content team'),
  T('PRJ-017','Task','Zara Osei','Instagram content batch — August (20 posts)','High','75%',true,true,'8/1/2026','8/22/2026','Execution','In Progress','',12,9,'None',false,false,'','','Carousel posts, Reels, and Stories scheduled in Buffer'),
  T('PRJ-017','Task','Zara Osei','LinkedIn thought leadership article series — 4 posts','Medium','50%',false,false,'8/1/2026','8/31/2026','Execution','In Progress','',8,4,'None',false,false,'','','Long-form POV posts from CEO and VP Marketing perspectives'),
  T('PRJ-017','Task','Zara Osei','Paid social campaign — product awareness (Meta)','High','0%',true,false,'8/25/2026','9/20/2026','Execution','Not Started','TSK-0171','Finish-to-Start',12,0,'None',false,false,'','','$15K budget; targeting lookalike audiences; A/B test 3 creatives'),
  T('PRJ-017','Task','Zara Osei','Influencer outreach and content coordination','Medium','0%',false,false,'8/15/2026','9/10/2026','Execution','Not Started','',8,0,'None',false,false,'','','5 micro-influencers (10K-50K); product send + brief'),
  T('PRJ-017','Task','Zara Osei','Social analytics dashboard setup','Medium','0%',false,false,'8/20/2026','9/1/2026','Reporting','Not Started','',4,0,'None',false,false,'','','Sprout Social dashboard: reach, engagement, click-through KPIs'),
  T('PRJ-017','Deliverable','Zara Osei','Q3 social media performance report','Medium','0%',false,false,'9/25/2026','9/30/2026','Reporting','Backlog','TSK-0176','Finish-to-Start',4,0,'None',false,false,'','','Monthly KPI summary + content highlights deck for leadership'),
  T('PRJ-017','Recurring Task','Zara Osei','Weekly content scheduling and posting','Low','50%',false,false,'7/1/2026','9/30/2026','Execution','In Progress','',3,1,'None',true,false,'REC-003','',"Recurring every Monday — schedule week's content in Buffer"),

  // PRJ-018 Inventory Audit Q3
  T('PRJ-018','Task','Marcus Webb','Pre-audit preparation — count sheet setup','Medium','100%',false,false,'8/1/2026','8/7/2026','Preparation','Complete','',4,4,'None',false,false,'','','Barcode scan sheets and tablet apps configured for 3 zones'),
  T('PRJ-018','Task','Marcus Webb','Warehouse zone assignment and team briefing','Medium','100%',false,false,'8/7/2026','8/8/2026','Preparation','Complete','',2,2,'None',false,false,'','','Zones A (electronics), B (accessories), C (returns)'),
  T('PRJ-018','Task','Marcus Webb','Physical inventory count — Zone A','High','0%',true,false,'8/12/2026','8/15/2026','Count','Waiting','TSK-0181','Finish-to-Start',8,0,'None',false,false,'','','ON HOLD — ERP system update paused all audit activities'),
  T('PRJ-018','Task','Marcus Webb','Physical inventory count — Zone B','High','0%',false,false,'8/14/2026','8/17/2026','Count','Waiting','TSK-0183','Start-to-Start',8,0,'None',false,false,'','','Awaiting ERP update before count can proceed'),
  T('PRJ-018','Task','Marcus Webb','Physical inventory count — Zone C (returns)','Medium','0%',false,false,'8/15/2026','8/18/2026','Count','Waiting','TSK-0183','Start-to-Start',6,0,'None',false,false,'','','Damaged goods triage included in Zone C count'),
  T('PRJ-018','Task','Marcus Webb','Variance reconciliation and discrepancy report','High','0%',true,false,'8/19/2026','8/24/2026','Analysis','Not Started','TSK-0185','Finish-to-Start',8,0,'None',false,false,'','','Identify shrinkage, misplacements, and data entry errors'),
  T('PRJ-018','Task','Marcus Webb','ERP system inventory update','High','0%',true,false,'8/25/2026','8/31/2026','Completion','Not Started','TSK-0186','Finish-to-Start',4,0,'None',false,false,'','','Update system counts; write-offs approved by finance'),

  // PRJ-019 Brand Video Production
  T('PRJ-019','Task','Sarah Chen','Creative brief and storyboard approval','High','100%',true,true,'5/1/2026','5/20/2026','Pre-Production','Complete','',8,8,'None',false,false,'','','2-minute video; 4 scenes; Apex Realty approved brief 5/18'),
  T('PRJ-019','Task','Sarah Chen','Location scouting — 3 candidate sites','High','100%',true,false,'5/10/2026','5/31/2026','Pre-Production','Complete','',6,7,'None',false,false,'','','City skyline rooftop selected (location 2 of 3)'),
  T('PRJ-019','Task','Sarah Chen','Talent and crew booking','High','100%',false,false,'5/15/2026','6/15/2026','Pre-Production','Complete','',8,8,'None',false,false,'','','Director, DP, 2 actors, sound engineer; all contracted'),
  T('PRJ-019','Deliverable','Sarah Chen','Production day 1 — interior scenes','Critical','50%',true,true,'7/15/2026','7/15/2026','Production','Blocked','',8,4,'None',false,true,'','','BLOCKED: permit for rooftop location denied — seeking alternative'),
  T('PRJ-019','Deliverable','Sarah Chen','Production day 2 — exterior scenes','Critical','0%',true,true,'7/16/2026','7/16/2026','Production','Blocked','TSK-0191','Finish-to-Start',8,0,'None',false,true,'','','BLOCKED: depends on resolving Day 1 location issue'),
  T('PRJ-019','Task','Sarah Chen','Raw footage review and selects','High','0%',true,false,'8/1/2026','8/5/2026','Post-Production','Not Started','TSK-0192','Finish-to-Start',4,0,'None',false,false,'','','Director selects; editorial assembly begins'),
  T('PRJ-019','Task','Sarah Chen','Video editing and color grading','High','0%',true,false,'8/6/2026','8/10/2026','Post-Production','Not Started','TSK-0193','Finish-to-Start',20,0,'None',false,false,'','','Full edit + DaVinci Resolve color grade + music licensing'),
  T('PRJ-019','Approval','Sarah Chen','Client review and revision rounds','High','0%',true,true,'8/11/2026','8/13/2026','Post-Production','Not Started','TSK-0194','Finish-to-Start',8,0,'None',false,false,'','','2 revision rounds with Apex Realty stakeholders'),

  // PRJ-020 Supplier Contract Renewal
  T('PRJ-020','Task','Tyler Banks','Supplier performance review — scorecard for 6 suppliers','Medium','0%',false,false,'9/1/2026','9/12/2026','Review','Not Started','',8,0,'None',false,false,'','','On-time delivery, quality defect rate, responsiveness, pricing'),
  T('PRJ-020','Task','Tyler Banks','Draft renewal pricing and terms matrix','High','0%',true,false,'9/8/2026','9/19/2026','Negotiation','Not Started','TSK-0197','Finish-to-Start',6,0,'None',false,false,'','','Target: 5-8% cost reduction; 30-day payment terms; 2-year terms'),
  T('PRJ-020','Task','Tyler Banks','Legal review of contract templates','High','0%',true,false,'9/1/2026','9/19/2026','Legal','Not Started','',4,0,'None',false,false,'','','In-house counsel review of indemnity, IP, and termination clauses'),
  T('PRJ-020','Meeting','Tyler Banks','Negotiation call — Supplier 1 (Apex Parts)','High','0%',true,false,'9/22/2026','9/22/2026','Negotiation','Not Started','TSK-0198','Finish-to-Start',2,0,'None',false,false,'','','Largest supplier by spend — priority negotiation'),
  T('PRJ-020','Meeting','Tyler Banks','Negotiation calls — Suppliers 2-6','High','0%',false,false,'9/23/2026','10/10/2026','Negotiation','Not Started','TSK-0200','Finish-to-Start',8,0,'None',false,false,'','','Schedule individual 90-min calls with each remaining supplier'),
  T('PRJ-020','Deliverable','Tyler Banks','Final contract preparation — all 6 suppliers','High','0%',true,false,'10/13/2026','10/24/2026','Contracting','Not Started','TSK-0201','Finish-to-Start',12,0,'None',false,false,'','','Redlined contracts incorporating agreed terms and pricing'),
  T('PRJ-020','Approval','Tyler Banks','Executive sign-off on all supplier contracts','High','0%',true,true,'10/25/2026','10/27/2026','Contracting','Not Started','TSK-0202','Finish-to-Start',2,0,'None',false,false,'','MST-021','CFO and COO approval required for contracts over $250K'),
  T('PRJ-020','Task','Tyler Banks','Signed contract filing and ERP update','Medium','0%',false,false,'10/28/2026','10/31/2026','Contracting','Backlog','TSK-0203','Finish-to-Start',2,0,'None',false,false,'','','DocuSign filing + update vendor master in ERP with new terms'),
];

(async () => {
  const fmt = [];
  const vals = [];

  // ── Title ────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['MASTER TASK LOG']] });

  // ── Subtitle ─────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.teal), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [['Single source of truth — Kanban, Gantt, Weekly, Calendar, and Eisenhower views all read from this log. Add tasks here only.']] });

  // ── Instructions ─────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.terra), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A3`, values: [['FORMULA COLUMNS (blue background): Task ID, Project Name, Overdue, Days Remaining — do not edit. Fill in Project ID first to auto-populate Project Name.']] });

  // ── Quick stats row (rows 3-4, 0-indexed) ────────────────────────────────
  const stats = [
    ['Total Tasks',    `=COUNTA(F8:F3007)`],
    ['Active',         `=COUNTIFS(N8:N3007,"In Progress")`],
    ['Complete',       `=COUNTIFS(N8:N3007,"Complete")`],
    ['Overdue',        `=COUNTIFS(V8:V3007,TRUE)`],
    ['Blocked',        `=COUNTIFS(T8:T3007,TRUE)`],
    ['Due This Week',  `=COUNTIFS(L8:L3007,">="&TODAY(),L8:L3007,"<="&(TODAY()+7),N8:N3007,"<>Complete",N8:N3007,"<>Cancelled")`],
  ];
  const statW = Math.floor(NC / stats.length);
  stats.forEach(([label, fml], i) => {
    const c = i * statW;
    fmt.push({ mergeCells: { range: gridRange(SID, 3, 4, c, c+statW), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, c, c+statW), cell: { userEnteredFormat: { backgroundColor: hex(C.teal), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ mergeCells: { range: gridRange(SID, 4, 5, c, c+statW), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, c, c+statW), cell: { userEnteredFormat: { backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}4`, values: [[label]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}5`, values: [[fml]] });
  });

  // ── Blank spacer row 5 (0-indexed) ───────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 5, 6, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // ── Column Headers row 6 (0-indexed = row 7) ────────────────────────────
  const COLS = [
    'Task ID','Project ID','Project Name','Task Type','Assignee','Task Name',
    'Priority','Pct Complete','Important','Urgent','Start Date','Due Date','Phase','Status',
    'Dep Task ID','Est Hrs','Actual Hrs','Dep Type','Recurring','Blocked',
    'Rec Template','Overdue','Milestone ID','Days Rem','Notes','',
  ];
  fmt.push({ repeatCell: { range: gridRange(SID, 6, 7, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)' } });
  vals.push({ range: `${S}!A7`, values: [COLS] });

  // ── Freeze rows 1-7 (no frozenColumnCount — title merges full width) ─────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Data Validation ───────────────────────────────────────────────────────
  const dvList = (col, items, strict=false) => ({
    setDataValidation: {
      range: gridRange(SID, 7, 3007, col, col+1),
      rule: {
        condition: { type: 'ONE_OF_LIST', values: items.map(v => ({ userEnteredValue: v })) },
        showCustomUi: true, strict,
      },
    },
  });
  const dvBool = (col) => ({
    setDataValidation: {
      range: gridRange(SID, 7, 3007, col, col+1),
      rule: { condition: { type: 'BOOLEAN' }, strict: true },
    },
  });
  fmt.push(dvList(3, ['Task','Subtask','Milestone Task','Meeting','Approval','Deliverable','Follow-Up','Recurring Task','Other']));
  fmt.push(dvList(4, ['Sarah Chen','Marcus Webb','Olivia Park','Tyler Banks','Zara Osei','Self','TBD']));
  fmt.push(dvList(6, ['Low','Medium','High','Urgent']));
  fmt.push(dvList(7, ['0%','25%','50%','75%','100%'], false));
  fmt.push(dvBool(8));   // Important
  fmt.push(dvBool(9));   // Urgent
  fmt.push(dvList(13, ['Backlog','Not Started','In Progress','Waiting','Blocked','Review','Complete','Cancelled']));
  fmt.push(dvList(17, ['None','Finish-to-Start','Start-to-Start','Finish-to-Finish','Start-to-Finish']));
  fmt.push(dvBool(18));  // Recurring
  fmt.push(dvBool(19));  // Blocked

  // ── Formula columns (A, C, V, X) ─────────────────────────────────────────
  const fmlCols = [
    [0,  r => `=IF(F${r}="","","TSK-"&TEXT(ROW()-7,"0000"))`],
    [2,  r => `=IFERROR(INDEX(${PS}!$B$8:$B$507,MATCH(B${r},${PS}!$A$8:$A$507,0)),"")`],
    [21, r => `=IF(AND(L${r}<>"",L${r}<TODAY(),N${r}<>"Complete",N${r}<>"Cancelled"),TRUE,FALSE)`],
    [23, r => `=IF(L${r}="","",IF(N${r}="Complete","Done",L${r}-TODAY()))`],
  ];
  const fmlVals = [];
  fmlCols.forEach(([ci, gen]) => {
    const col = String.fromCharCode(65+ci);
    const formulas = Array.from({ length: 3000 }, (_, i) => [gen(8+i)]);
    fmlVals.push({ range: `${S}!${col}8:${col}3007`, values: formulas });
  });

  // ── Number formats ────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 3007, 10, 11), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } }); // Start Date
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 3007, 11, 12), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } }); // Due Date
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 3007,  7,  8), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } }); // Pct Complete
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 3007, 15, 17), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0' } } }, fields: 'userEnteredFormat.numberFormat' } }); // Hours

  // ── Formula column backgrounds ────────────────────────────────────────────
  [0, 2, 21, 23].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID, 7, 3007, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // ── Row styling ───────────────────────────────────────────────────────────
  for (let r = 0; r < 400; r++) {
    if (r % 2 !== 0) {
      fmt.push({ repeatCell: { range: gridRange(SID, 7+r, 8+r, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
    }
  }

  // ── Row and column dimensions ─────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1  }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7  }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  const colWidths = [75,70,150,80,85,240,70,70,60,55,90,80,90,90,75,55,60,80,55,55,75,55,80,60,220,40];
  colWidths.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, fmt, 'tl-fmt');

  // ── Sample task data ──────────────────────────────────────────────────────
  vals.push({ range: `${S}!A8`, values: TASKS });
  fmlVals.forEach(fv => vals.push(fv));

  await valuesBatchUpdate(id, vals, 'tl-vals');
  console.log(`✓ Master Task Log complete — ${TASKS.length} tasks written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
