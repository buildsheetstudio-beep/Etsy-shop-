'use strict';
// Fix script for Ultimate Life Planner
// 1. Dashboard A21: wrong bills date column (I:I → J:J for Due Date)
// 2. Calendar rows 4-9: same wrong bills date column in all 42 cell formulas
// 3. Update sample data dates from Jan-Mar 2025 to May-Jul 2026 so dashboard
//    shows real numbers for the current month instead of 0s
const { batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const BUD_NAME = "'💰 Budget & Expense Tracker'";
const TSK_NAME = "'✅ Task Manager'";

(async () => {
  const data = [];

  // ── FIX 1: Dashboard A21 — Bills Due This Week (I:I → J:J) ──────────────
  // Root cause: bills section H:K has H=Bill Name, I=Amount, J=Due Date, K=Paid.
  // The original formula used I:I (Amount) as the date range, so the date
  // comparison always failed and no bills ever appeared.
  data.push({ range: "'✨ Dashboard'!A21", values: [[
    `=IFERROR(IF(COUNTIFS(${BUD_NAME}!J:J,">="&TODAY(),${BUD_NAME}!J:J,"<="&TODAY()+7,${BUD_NAME}!K:K,FALSE)=0,"✅ No unpaid bills due this week",TEXTJOIN(CHAR(10),TRUE,FILTER(${BUD_NAME}!H:H,(${BUD_NAME}!J:J>=TODAY())*(${BUD_NAME}!J:J<=TODAY()+7)*(${BUD_NAME}!K:K=FALSE)))),"")`
  ]] });

  // ── FIX 2: Calendar rows 4–9 — all 42 cell formulas (I:I → J:J) ────────
  for (let week = 0; week < 6; week++) {
    const rowIdx = 4 + week;
    const rowData = [];
    for (let dow = 0; dow < 7; dow++) {
      const offset = week * 7 + dow;
      const dayRef = `DATE($D$2,$B$2,1)+${offset}-WEEKDAY(DATE($D$2,$B$2,1),2)+1`;
      const inMonth = `IF(MONTH(${dayRef})=$B$2,${dayRef},"")`;
      const formula = `=IFERROR(IF(IFERROR(${inMonth},"")="","",TEXT(${inMonth},"D")&CHAR(10)&IFERROR(TEXTJOIN(", ",TRUE,FILTER(${TSK_NAME}!A:A,${TSK_NAME}!E:E=${dayRef})),"")&IFERROR(IF(COUNTIF(${BUD_NAME}!J:J,${dayRef})>0,CHAR(10)&"Bill: "&TEXTJOIN(", ",TRUE,FILTER(${BUD_NAME}!H:H,${BUD_NAME}!J:J=${dayRef})),""),"")),"")`;
      rowData.push(formula);
    }
    data.push({ range: `'📅 Monthly Calendar View'!A${rowIdx}:G${rowIdx}`, values: [rowData] });
  }

  // ── FIX 3: Budget transactions — shift Jan–Mar 2025 → May–Jul 2026 ─────
  // Dashboard SUMIFS filters by current month (July 2026). Old sample data
  // (Jan–Mar 2025) produced $0 for every finance stat. New dates give real
  // income and expense totals for the current month.
  const TXN_ROWS = [
    ['2026-05-03','Housing','May Rent','Expense',1850,'Bank Transfer',''],
    ['2026-05-05','Food','Weekly groceries - Trader Joe\'s','Expense',112.47,'Debit Card',''],
    ['2026-05-08','Transportation','Gas fill-up','Expense',58.20,'Credit Card',''],
    ['2026-05-10','Personal','Coffee & lunch out','Expense',24.80,'Credit Card',''],
    ['2026-05-15','Income','Paycheck','Income',3200,'Direct Deposit',''],
    ['2026-05-15','Utilities','Electric bill','Expense',94.50,'Auto-Pay',''],
    ['2026-05-16','Entertainment','Netflix subscription','Expense',15.99,'Credit Card',''],
    ['2026-05-18','Health','Gym membership','Expense',45.00,'Credit Card',''],
    ['2026-05-20','Food','Groceries','Expense',89.30,'Debit Card',''],
    ['2026-05-22','Savings','Transfer to savings','Expense',300,'Bank Transfer','Monthly savings goal'],
    ['2026-05-25','Personal','Haircut','Expense',45,'Cash',''],
    ['2026-05-28','Food','Takeout - sushi night','Expense',62.40,'Credit Card',''],
    ['2026-05-30','Income','Freelance payment','Income',850,'Bank Transfer','Web project'],
    ['2026-06-01','Housing','Renter\'s insurance','Expense',22.50,'Auto-Pay',''],
    ['2026-06-03','Transportation','Monthly bus pass','Expense',75,'Cash',''],
    ['2026-06-05','Food','Groceries','Expense',104.85,'Debit Card',''],
    ['2026-06-08','Debt Payments','Student loan','Expense',250,'Auto-Pay',''],
    ['2026-06-10','Personal','Amazon order','Expense',38.99,'Credit Card','Household supplies'],
    ['2026-06-14','Entertainment','Dinner out','Expense',95,'Credit Card',''],
    ['2026-06-15','Income','Paycheck','Income',3200,'Direct Deposit',''],
    ['2026-06-18','Health','Dentist co-pay','Expense',50,'Credit Card',''],
    ['2026-06-20','Food','Groceries','Expense',98.60,'Debit Card',''],
    ['2026-06-22','Savings','Transfer to savings','Expense',300,'Bank Transfer',''],
    ['2026-06-25','Utilities','Phone bill','Expense',65,'Auto-Pay',''],
    ['2026-06-27','Other','Car registration renewal','Expense',120,'Online','Annual fee'],
    ['2026-07-01','Housing','July Rent','Expense',1850,'Bank Transfer',''],
    ['2026-07-03','Food','Grocery run','Expense',91.20,'Debit Card',''],
    ['2026-07-07','Transportation','Uber rides this week','Expense',28.50,'Credit Card',''],
    ['2026-07-12','Entertainment','Concert tickets','Expense',110,'Credit Card','Local show'],
    ['2026-07-15','Income','Paycheck','Income',3200,'Direct Deposit',''],
  ];
  data.push({ range: `${BUD_NAME}!A2:G31`, values: TXN_ROWS });

  // ── FIX 4: Budget bills — shift Apr 2025 due dates → Jul–Aug 2026 ───────
  // Several bills now fall within the next 7 days so "Bills Due This Week"
  // shows real data (requires Fix 1 to be live first).
  const BILL_ROWS = [
    ['Rent',          1850,  '2026-08-01', false],
    ['Electric',      94.50, '2026-08-05', false],
    ['Phone',         65,    '2026-07-28', true ],   // already paid
    ['Internet',      59.99, '2026-07-29', false],   // due in 3 days
    ['Renter\'s Insurance', 22.50, '2026-08-01', false],
    ['Gym Membership',45,    '2026-08-15', false],
    ['Student Loan',  250,   '2026-07-30', false],   // due in 4 days
    ['Netflix',       15.99, '2026-07-31', false],   // due in 5 days
  ];
  data.push({ range: `${BUD_NAME}!H2:K9`, values: BILL_ROWS });

  // ── FIX 5: Habit tracker — shift Jan 2025 dates → Jul 19–26, 2026 ───────
  // Dashboard COUNTIFS looks for habits logged THIS WEEK. Old dates (Jan 2025)
  // produced "—" for habit completion %. This week's dates give real results.
  const habits = ['Morning Meditation','Exercise / Move','Read 20 Minutes','Drink 8 Glasses of Water','Gratitude Journal'];
  const dates  = ['2026-07-19','2026-07-20','2026-07-21','2026-07-22','2026-07-23','2026-07-24','2026-07-25','2026-07-26'];
  const completions = {
    'Morning Meditation':       [true, true, false,true, true, false,true, true ],
    'Exercise / Move':          [true, false,true, true, false,true, true, false],
    'Read 20 Minutes':          [true, true, true, false,true, true, true, true ],
    'Drink 8 Glasses of Water': [false,true, true, true, true, false,true, true ],
    'Gratitude Journal':        [true, true, false,false,true, true, false,true ],
  };
  const HAB_ROWS = [];
  dates.forEach((date, di) => { habits.forEach(habit => { HAB_ROWS.push([date, habit, completions[habit][di]]); }); });
  data.push({ range: "'🔁 Habit Tracker'!A2:C41", values: HAB_ROWS });

  // ── FIX 6: Journal — shift Jan 2025 dates → Jul 13–26, 2026 ─────────────
  // Dashboard FILTER pulls mood entries from the last 7 days. Old dates (Jan
  // 2025) produced "No mood entries this week."
  const JRN_ROWS = [
    ['2026-07-13','My warm apartment on a summer morning','A job I actually don\'t hate','My good health','Feeling motivated going into this week. Set up the life planner. Excited to track everything.','Great'],
    ['2026-07-14','Good coffee before work','A friend texting to check in','Running water','Long work day but manageable. Texted Marcus — haven\'t talked in months. He\'s doing well.','Good'],
    ['2026-07-15','Rest after a tiring day','Groceries already prepped','My cozy bed','Tired tonight. Skipped gym. Didn\'t beat myself up — just went to sleep earlier.','Okay'],
    ['2026-07-16','A sunny summer day','Making progress on Python','Evening walk outside','Python lesson 14 done. The sunset was genuinely beautiful tonight. Good reset.','Good'],
    ['2026-07-17','Finishing the work week','My playlist on the commute','Plans with friends this weekend','End-of-week tiredness but in a good way. Looking forward to the weekend.','Good'],
    ['2026-07-18','A long Saturday sleep-in','Cooking a real meal','Calling home','Made pasta from scratch. Called parents. Two hours of reading. This is what weekends are for.','Great'],
    ['2026-07-19','A reset Sunday','Gym session that felt strong','Prepping for the week','Meal prepped for the week. Planned my goals. Feel genuinely ready for Monday for once.','Great'],
    ['2026-07-20','My morning routine holding steady','Progress feeling real','A good conversation today','Hard day at work — unexpected conflict. Handled it okay. Journaling helps more than I expected.','Okay'],
    ['2026-07-21','Warm tea on a gray day','Small wins still count','My support system','Feeling a little flat today. Not sad, just… muted. Still did all my habits. That counts.','Low'],
    ['2026-07-22','Payday','Feeling caught up for once','Room cleaned and calm','Paid rent, transferred to savings, still have money left. Small win. Budget is working.','Good'],
    ['2026-07-23','Friends who check in without prompting','My couch and a good book','Music I forgot I loved','Nothing dramatic happened today. Just quiet and nice. Grateful for ordinary good days.','Good'],
    ['2026-07-24','A healthy lunch I actually enjoyed','Fresh air on a walk','Momentum building','Work felt heavy. Walk at lunch helped reset. Remembering to give myself breaks.','Okay'],
    ['2026-07-25','Progress on the emergency fund','Running farther than last month','Clarity about priorities','Ran 2.8 miles — almost 3! Proud of myself. Finance section really helps me see progress.','Great'],
    ['2026-07-26','A good Saturday morning','Time to reflect','Plans ahead','A week well spent. Grateful for the habits that held and for the reset today brings.','Good'],
  ];
  data.push({ range: "'📓 Journal & Gratitude Log'!A2:F15", values: JRN_ROWS });

  await valuesBatchUpdate(id, data, 'fix-lifeplanner');

  console.log('Life Planner fixes applied:');
  console.log('  ✓ Dashboard A21: bills date column corrected (I→J)');
  console.log('  ✓ Calendar rows 4-9: all 42 cell formulas corrected (I→J)');
  console.log('  ✓ Budget transactions: dates updated to May-Jul 2026');
  console.log('    → Income this month (Jul 2026): $3,200');
  console.log('    → Expenses this month (Jul 2026): ~$2,080 (rent + food + uber + concert)');
  console.log('  ✓ Bills: due dates updated to Jul 29-Aug 2026');
  console.log('    → Bills due this week: Internet, Student Loan, Netflix, Renter\'s Insurance, Rent');
  console.log('  ✓ Habits: dates updated to Jul 19-26, 2026 (~70% completion this week)');
  console.log('  ✓ Journal: dates updated to Jul 13-26, 2026 (7 mood entries this week)');
  console.log('Spreadsheet:', `https://docs.google.com/spreadsheets/d/${id}/edit`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
