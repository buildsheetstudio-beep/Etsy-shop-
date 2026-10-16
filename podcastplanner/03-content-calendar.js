'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CAL = sheetMap['🎙 Content Calendar'];

// A=Episode# | B=Content Pillar | C=Episode Type | D=Target Audience | E=Episode Title
// F=Date | G=Status | H=Guest | I=Episode Description | J=Keywords | K=CTA | L=Script Link | M=Script URL (hidden)

const HEADERS = ['Episode #','Content Pillar','Episode Type','Target Audience','Episode Title','Date','Status','Guest','Episode Description','Keywords','CTA','Script Link'];
const WIDTHS  = [80, 140, 120, 140, 220, 100, 110, 120, 260, 180, 180, 100];

// 30 sample rows — all pillars, types, audiences, statuses, solo/guest mix
const sampleData = [
  [1,'Personal Growth','Solo Talk','New Freelancers','From 9-to-5 to Freelance: My First 90 Days','2024-01-08','Published','N/A','The honest story of leaving a corporate job with no clients lined up and how I survived.','freelance transition, first clients, side hustle','Join my free community at the link below','https://docs.google.com/document/d/abc1'],
  [2,'Content Creation','Guest Interview','Content Creators','How She Built 50K Followers Creating Content Part-Time','2024-01-15','Published','Sarah Chen','Sarah shares her exact content system for consistent posting while working full-time.','content creation, part-time, social media growth','Download the content calendar template','https://docs.google.com/document/d/abc2'],
  [3,'Productivity','How-To','Remote Workers','My 5-Block Daily Schedule That Tripled My Output','2024-01-22','Published','N/A','The time-blocking method I use to get deep work done without burning out.','time blocking, productivity, schedule, deep work','Grab my free daily planner template','https://docs.google.com/document/d/abc3'],
  [4,'Freelancing Tips','Guest Interview','New Freelancers','Landing Your First $5K Client: What Actually Works','2024-01-29','Published','Marcus Webb','Marcus went from zero to consistent $10K months in under a year. Here is how.','client acquisition, freelancing, pricing','Book a strategy call','https://docs.google.com/document/d/abc4'],
  [5,'Online Business','Solo Talk','Digital Product Sellers','Why I Stopped Trading Time for Money (And What I Do Instead)','2024-02-05','Published','N/A','Passive income is a lie — but productized services and digital products are real.','passive income, digital products, business model','Shop my Etsy digital products','https://docs.google.com/document/d/abc5'],
  [6,'Mindset & Motivation','Storytime','Overwhelmed Beginners','The Panic Attack That Changed How I Run My Business','2024-02-12','Published','N/A','Burnout hit me hard. Here is what I changed and what I wish I had known sooner.','burnout, mental health, freelance boundaries','DM me "burnout" on Instagram',''],
  [7,'Tools & Workflows','How-To','Solo Creators','My Notion Setup for Running a 6-Figure Solo Business','2024-02-19','Published','N/A','The exact templates and dashboards I use to track projects, clients, and revenue.','Notion, productivity tools, solo business, templates','Download my Notion template','https://docs.google.com/document/d/abc7'],
  [8,'Audience Building','Guest Interview','Aspiring Entrepreneurs','Email List of 10K in 6 Months: The Pinterest Strategy Nobody Talks About','2024-02-26','Published','Priya Nair','Priya grew her email list faster than anyone I know — using a platform most ignore.','email list, Pinterest, audience growth','Join her free workshop','https://docs.google.com/document/d/abc8'],
  [9,'Work-Life Balance','Solo Talk','Self-employed Millennials','How I Take Real Weekends Off Without Losing Clients','2024-03-04','Published','N/A','Boundaries, automation, and communication frameworks that actually work.','work-life balance, boundaries, automation','Join the waitlist for my course',''],
  [10,'Monetization','How-To','Digital Product Sellers','Pricing Your Digital Products: Stop Undercharging','2024-03-11','Published','N/A','The psychology of pricing, perceived value, and why your $9 product should be $27.','pricing strategy, digital products, value','Download my pricing calculator','https://docs.google.com/document/d/abc10'],
  [11,'Personal Growth','Guest Interview','Content Creators','Identity Shift: Thinking Like a CEO When You Are Still a Freelancer','2024-03-18','Published','Dr. Aisha Brown','A therapist and business coach on the mindset work that separates six-figure solopreneurs.','mindset, CEO thinking, personal development','Book a session with Dr. Brown','https://docs.google.com/document/d/abc11'],
  [12,'Content Creation','Q&A','Solo Creators','Your Top 20 Content Questions Answered (Q&A Episode)','2024-03-25','Published','N/A','Answered live from Instagram stories — batching, repurposing, writer block solutions.','content creation Q&A, repurposing, batching','Submit your question for next Q&A',''],
  [13,'Productivity','Guest Interview','Remote Workers','ADHD and Entrepreneurship: Building Systems That Work for Your Brain','2024-04-01','Published','Jordan Lee','Jordan built a multiple six-figure business with ADHD. Here is every system they use.','ADHD, productivity, systems, neurodivergent business','Grab Jordan\'s free resource pack','https://docs.google.com/document/d/abc13'],
  [14,'Freelancing Tips','Mini Episode','New Freelancers','5 Things to Do Before You Quit Your Job','2024-04-08','Published','N/A','A 10-minute rapid-fire checklist for anyone planning to go full-time freelance this year.','quit job, freelance, checklist, side hustle','Download the pre-quit checklist','https://docs.google.com/document/d/abc14'],
  [15,'Online Business','Behind-the-Scenes','Aspiring Entrepreneurs','Behind My $8K Month: Revenue Breakdown and Lessons Learned','2024-04-15','Published','N/A','Transparent income report covering every revenue stream, what worked, what did not.','income report, revenue breakdown, transparency','Join my monthly income report newsletter',''],
  [16,'Mindset & Motivation','Guest Interview','Overwhelmed Beginners','From 0 to $100K: The Unsexy Truth About Building a Business','2024-04-22','Published','Tomás Rivera','Tomás debunks overnight success myths and shares the actual timeline of his journey.','success, realism, entrepreneurship, motivation','Follow Tomás on LinkedIn','https://docs.google.com/document/d/abc16'],
  [17,'Tools & Workflows','How-To','Content Creators','How I Batch 30 Days of Content in One Afternoon','2024-04-29','Published','N/A','My exact batching workflow — from outline to final post — with AI and repurposing tools.','content batching, repurposing, AI tools, workflow','Grab my batching workshop replay','https://docs.google.com/document/d/abc17'],
  [18,'Audience Building','Solo Talk','Digital Product Sellers','Why Your Audience Size Does Not Matter (And What Does)','2024-05-06','Published','N/A','Micro audiences, high conversion, and how I make more with 3K followers than some do with 30K.','small audience, conversion, email list, niche','Download my small-audience sales guide',''],
  [19,'Work-Life Balance','Lessons Learned','Self-employed Millennials','5 Lessons from My First 5 Years in Business','2024-05-13','Published','N/A','Honest reflection on pricing mistakes, client nightmares, and things I would do differently.','lessons learned, business anniversary, reflection','Join my anniversary giveaway',''],
  [20,'Monetization','Guest Interview','Aspiring Entrepreneurs','Selling on Etsy: From Hobby to $15K/Month with Digital Products','2024-05-20','Published','Emily Park','Emily left teaching to sell Etsy templates and now earns more than she ever did in a classroom.','Etsy, digital products, passive income, side hustle','Shop Emily\'s Etsy store','https://docs.google.com/document/d/abc20'],
  [21,'Personal Growth','Monthly Recap','Solo Creators','May Recap: Launches, Lessons, and What\'s Coming in June','2024-05-27','Published','N/A','A transparent monthly wrap-up covering wins, struggles, and what I have planned next.','monthly recap, transparency, business update','Subscribe so you don\'t miss June',''],
  [22,'Content Creation','How-To','Content Creators','The 3-Part Content Framework That Gets More Saves Than Likes','2024-06-03','Scheduled','N/A','Why saves beat likes for algorithm growth and the framework to write save-worthy content.','content framework, saves, algorithm, engagement','Download the framework one-pager','https://docs.google.com/document/d/abc22'],
  [23,'Productivity','Guest Interview','Remote Workers','Deep Work in the Age of Notifications: Cal Newport Would Approve','2024-06-10','Scheduled','Alex Kim','Alex runs a remote team and has cracked the code on focused productivity at scale.','deep work, notifications, remote work, focus','Join Alex\'s focus challenge','https://docs.google.com/document/d/abc23'],
  [24,'Freelancing Tips','Solo Talk','New Freelancers','How to Raise Your Rates Without Losing Clients','2024-06-17','Recording','N/A','Scripts, timing, and framing for delivering a rate increase your clients will actually accept.','rate increase, pricing, client communication','Download my rate-raise email template',''],
  [25,'Online Business','Launch Day','Digital Product Sellers','The Ultimate Content Creator Toolkit — LAUNCH DAY Episode','2024-06-24','Recording','N/A','Everything inside the toolkit, who it is for, and the launch week bonuses.','product launch, digital toolkit, creator economy','Buy the toolkit — link in bio','https://docs.google.com/document/d/abc25'],
  [26,'Mindset & Motivation','Guest Interview','Overwhelmed Beginners','Imposter Syndrome at $500K: It Doesn\'t Go Away, But This Helps','2024-07-01','Editing','Dr. Maya Singh','Dr. Singh explains why imposter syndrome scales with success and practical techniques to manage it.','imposter syndrome, confidence, success mindset','Follow Dr. Singh on Instagram','https://docs.google.com/document/d/abc26'],
  [27,'Tools & Workflows','Behind-the-Scenes','Solo Creators','My Entire Tech Stack for Running a Podcast + Online Business','2024-07-08','Editing','N/A','Every tool I use, what it costs, what I would cut first, and budget alternatives.','tech stack, podcast tools, software, overhead','Download the tech stack spreadsheet',''],
  [28,'Audience Building','Q&A','Aspiring Entrepreneurs','Audience Building Q&A: Your Questions on List Building and Visibility','2024-07-15','Ready to Record','N/A','25 listener questions answered live on email, SEO, Pinterest, and social growth.','audience growth Q&A, email list, SEO, visibility','Submit your question via the form',''],
  [29,'Work-Life Balance','Mini Episode','Self-employed Millennials','3-Minute Monday: One Rule to Protect Your Energy Every Week','2024-07-22','On Hold','N/A','A quick solo riff on the single weekly habit that keeps me from burning out.','energy management, quick tip, Monday motivation','DM me your energy-protecting rule',''],
  [30,'Monetization','Guest Interview','Digital Product Sellers','Affiliate Marketing for Digital Product Creators: What Converts','2024-07-29','Scheduled','Lisa Tran','Lisa earns more from affiliates than her own products. She breaks down her exact strategy.','affiliate marketing, digital products, passive income','Join Lisa\'s affiliate program','https://docs.google.com/document/d/abc30'],
];

(async () => {
  const reqs = [];

  // Row 0: Title (col A standalone for col-freeze compat)
  reqs.push({ mergeCells: { range: gridRange(CAL, 0, 1, 1, 13), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(CAL, 0, 1, 0, 1),
      cell: {
        userEnteredValue: { stringValue: '🎙' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepCharcoal),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(CAL, 0, 1, 1, 13),
      cell: {
        userEnteredValue: { stringValue: 'Content Calendar & Episode Planner' },
        userEnteredFormat: {
          backgroundColor: hex(C.deepCharcoal),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CAL, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Row 1: Headers
  reqs.push({
    repeatCell: {
      range: gridRange(CAL, 1, 2, 0, 13),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.burntOrange),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CAL, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});

  // Data rows 2-31
  for (let r = 2; r < 32; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(CAL, r, r+1, 0, 13),
        cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.inputBg : C.altRow), verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: CAL, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    }});
  }
  // Empty rows
  reqs.push({
    repeatCell: {
      range: gridRange(CAL, 32, 500, 0, 13),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat',
    },
  });

  // Col L (index 11): formula bg, Col M (index 12): hidden URL col
  reqs.push({ repeatCell: {
    range: gridRange(CAL, 2, 500, 11, 12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg) } },
    fields: 'userEnteredFormat',
  }});

  // Hide col M (script URL)
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CAL, dimension: 'COLUMNS', startIndex: 12, endIndex: 13 },
    properties: { hiddenByUser: true }, fields: 'hiddenByUser',
  }});

  // Col widths (A-L)
  for (const [i, px] of WIDTHS.entries()) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: CAL, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  }

  // Number formats
  reqs.push({ repeatCell: {
    range: gridRange(CAL, 2, 500, 5, 6),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat.numberFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(CAL, 2, 500, 0, 1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true } } },
    fields: 'userEnteredFormat',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(CAL, 2, 500, 6, 7),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat',
  }});

  await batchUpdate(id, reqs, 'cal-format');

  // Values
  const data = [];
  data.push({ range: `'🎙 Content Calendar'!A2:L2`, values: [HEADERS] });

  for (let i = 0; i < sampleData.length; i++) {
    const row = i + 3;
    const rowData = [...sampleData[i]];
    const scriptUrl = rowData.pop(); // col M (hidden, index 12)
    // Col L: Script Link formula uses col M (same row)
    rowData.push(`=IF(M${row}="","",HYPERLINK(M${row},"View Script"))`);
    rowData.push(scriptUrl); // col M
    data.push({ range: `'🎙 Content Calendar'!A${row}:M${row}`, values: [rowData] });
  }

  await valuesBatchUpdate(id, data, 'cal-values');
  console.log('Content Calendar complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
