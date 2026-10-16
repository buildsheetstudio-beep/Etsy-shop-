'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Reference Data'];
const S = "'Reference Data'";

(async () => {
  const vals = [];
  const fmt = [];

  const col = (header, items, colLetter) => {
    vals.push({ range: `${S}!${colLetter}1`, values: [[header]] });
    vals.push({ range: `${S}!${colLetter}2:${colLetter}${1+items.length}`, values: items.map(v => [v]) });
  };

  // A: Workout Types (10)
  col('Workout Types', ['Strength','Cardio','Mobility','Flexibility','Conditioning','Circuit','HIIT','Recovery','Mixed','Other'], 'A');

  // B: Muscle Groups (20)
  col('Muscle Groups', ['Chest','Back','Shoulders','Arms','Biceps','Triceps','Forearms','Core','Abs','Glutes','Quads','Hamstrings','Calves','Legs','Full Body','Upper Body','Lower Body','Push','Pull','Other'], 'B');

  // C: Exercise Categories (7)
  col('Exercise Categories', ['Strength','Bodyweight','Cardio','Mobility','Flexibility','Conditioning','Other'], 'C');

  // D: Tracking Modes (8)
  col('Tracking Modes', ['Weight + Reps','Reps Only','Time','Distance','Speed','Pace','Assisted Weight','Bodyweight'], 'D');

  // E: Workout Status (7)
  col('Workout Status', ['Planned','Completed','Partially Completed','Skipped','Rest Day','Rescheduled','Cancelled'], 'E');

  // F: Session Intensity (5)
  col('Session Intensity', ['Very Light','Light','Moderate','Hard','Very Hard'], 'F');

  // G: RPE (10)
  col('RPE', [1,2,3,4,5,6,7,8,9,10], 'G');

  // H: Workout Location (6)
  col('Workout Location', ['Gym','Home','Outdoors','Studio / Class','Travel','Other'], 'H');

  // I: Equipment (14)
  col('Equipment', ['None / Bodyweight','Dumbbells','Barbell','Kettlebell','Cable','Machine','Resistance Band','Bench','Pull-Up Bar','Cardio Machine','Bike','Treadmill','Rowing Machine','Other'], 'I');

  // J: Goal Types (8)
  col('Goal Types', ['Workouts per Week','Workouts per Month','Strength Goal','Cardio Goal','Consistency Goal','Streak Goal','Exercise PR Goal','Custom'], 'J');

  // K: Goal Status (7)
  col('Goal Status', ['Not Started','In Progress','On Track','Behind Plan','Achieved','Paused','Archived'], 'K');

  // L: Day of Week (7)
  col('Day of Week', ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'], 'L');

  // M: Yes / No
  col('Yes / No', ['Yes','No'], 'M');

  // N: Weight Units
  col('Weight Units', ['lb','kg'], 'N');

  // O: Distance Units
  col('Distance Units', ['miles','km','meters'], 'O');

  // P: Week Start Days
  col('Week Start', ['Monday','Sunday'], 'P');

  // Background + header
  fmt.push({ repeatCell: { range: gridRange(SID,0,200,0,20), cell: { userEnteredFormat: { backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial' } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,20), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Hide the sheet
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, hidden: true }, fields: 'hidden' } });

  await valuesBatchUpdate(id, vals, '02-reference values');
  await batchUpdate(id, fmt, '02-reference format');
  console.log('✅  Reference Data done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
