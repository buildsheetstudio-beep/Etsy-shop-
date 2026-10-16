'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C,
  STATUSES, PRIORITIES, DEPARTMENTS, PROJECTS, RECURRENCE, TASK_TYPES, EFFORTS,
  YESNO, WEEKDAYS, PROJ_HEALTH, TEAM_MEMBERS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const RD = sheetMap['Reference Data'];
const S = "'Reference Data'";

(async () => {
  const data = [];

  // Column headers
  data.push({ range:`${S}!A1`, values:[['Task Statuses']] });
  data.push({ range:`${S}!B1`, values:[['Priorities']] });
  data.push({ range:`${S}!C1`, values:[['Departments']] });
  data.push({ range:`${S}!D1`, values:[['Projects']] });
  data.push({ range:`${S}!E1`, values:[['Recurrence Types']] });
  data.push({ range:`${S}!F1`, values:[['Task Types']] });
  data.push({ range:`${S}!G1`, values:[['Effort Levels']] });
  data.push({ range:`${S}!H1`, values:[['Yes / No']] });
  data.push({ range:`${S}!I1`, values:[['Weekdays']] });
  data.push({ range:`${S}!J1`, values:[['Project Health']] });
  data.push({ range:`${S}!K1`, values:[['Team Members']] });
  // Filters column (with All options)
  data.push({ range:`${S}!M1`, values:[['Member Filter']] });
  data.push({ range:`${S}!N1`, values:[['Project Filter']] });
  data.push({ range:`${S}!O1`, values:[['Dept Filter']] });
  data.push({ range:`${S}!P1`, values:[['Month Names']] });

  // A: Task Statuses
  STATUSES.forEach((v,i) => data.push({ range:`${S}!A${2+i}`, values:[[v]] }));
  // B: Priorities
  PRIORITIES.forEach((v,i) => data.push({ range:`${S}!B${2+i}`, values:[[v]] }));
  // C: Departments
  DEPARTMENTS.forEach((v,i) => data.push({ range:`${S}!C${2+i}`, values:[[v]] }));
  // D: Projects
  PROJECTS.forEach((v,i) => data.push({ range:`${S}!D${2+i}`, values:[[v]] }));
  // E: Recurrence
  RECURRENCE.forEach((v,i) => data.push({ range:`${S}!E${2+i}`, values:[[v]] }));
  // F: Task Types
  TASK_TYPES.forEach((v,i) => data.push({ range:`${S}!F${2+i}`, values:[[v]] }));
  // G: Efforts
  EFFORTS.forEach((v,i) => data.push({ range:`${S}!G${2+i}`, values:[[v]] }));
  // H: Yes/No
  YESNO.forEach((v,i) => data.push({ range:`${S}!H${2+i}`, values:[[v]] }));
  // I: Weekdays
  WEEKDAYS.forEach((v,i) => data.push({ range:`${S}!I${2+i}`, values:[[v]] }));
  // J: Project Health
  PROJ_HEALTH.forEach((v,i) => data.push({ range:`${S}!J${2+i}`, values:[[v]] }));
  // K: Team Members
  TEAM_MEMBERS.forEach((v,i) => data.push({ range:`${S}!K${2+i}`, values:[[v]] }));

  // M: Member filter (All Team Members + members)
  data.push({ range:`${S}!M2`, values:[['All Team Members']] });
  TEAM_MEMBERS.forEach((v,i) => data.push({ range:`${S}!M${3+i}`, values:[[v]] }));

  // N: Project filter
  data.push({ range:`${S}!N2`, values:[['All Projects']] });
  PROJECTS.forEach((v,i) => data.push({ range:`${S}!N${3+i}`, values:[[v]] }));

  // O: Dept filter
  data.push({ range:`${S}!O2`, values:[['All Departments']] });
  DEPARTMENTS.forEach((v,i) => data.push({ range:`${S}!O${3+i}`, values:[[v]] }));

  // P: Month names
  ['January','February','March','April','May','June','July','August','September','October','November','December']
    .forEach((v,i) => data.push({ range:`${S}!P${2+i}`, values:[[v]] }));

  await valuesBatchUpdate(id, data, 'ref-values');

  // Formatting
  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(RD,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });

  fmt(0,1,0,16,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(1,20,0,16,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ fontSize:9, fontFamily:'Arial' } } });

  reqs.push({ updateDimensionProperties:{ range:{ sheetId:RD, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:24 }, fields:'pixelSize' } });
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:RD, gridProperties:{ frozenRowCount:1 } }, fields:'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'ref-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
