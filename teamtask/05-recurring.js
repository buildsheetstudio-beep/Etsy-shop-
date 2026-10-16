'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const { google } = require('googleapis');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const RP = sheetMap['Recurring Task Planner'];
const S = "'Recurring Task Planner'";

// [name, project, dept, assignedTo, priority, effort, recurrenceType, startDate, lastCompleted, estHrs, active, notes]
const RECURRING = [
  ['Weekly team meeting agenda','Team Operations','Operations','Jamie Lee','Medium','Quick','Weekly','2026-01-05','2026-07-21',1,true,'Prepare 3 agenda items minimum'],
  ['Monthly payroll review','Monthly Reporting','Finance','Morgan Chen','High','Small','Monthly','2026-01-01','2026-06-28',2,true,'Run by EOD on last business day'],
  ['Daily inbox triage','Team Operations','Operations','Jamie Lee','Medium','Quick','Weekdays','2026-01-05','2026-07-23',0.5,true,'Flag urgent items for team lead'],
  ['Monthly financial close','Monthly Reporting','Finance','Morgan Chen','High','Medium','Monthly','2026-01-01','2026-06-30',6,true,'Coordinate with accounting by 3rd'],
  ['Quarterly process audit','Process Improvement','Operations','Jamie Lee','High','Large','Quarterly','2026-01-01','2026-03-31',12,true,'Review 5 core workflows per quarter'],
  ['Monthly social media planning','Social Media Campaign','Marketing','Taylor Brooks','Medium','Medium','Monthly','2026-01-01','2026-06-30',8,true,'Schedule 20 posts per month'],
  ['Weekly client status update','Client Onboarding','Customer Service','Sam Thompson','High','Small','Weekly','2026-01-05','2026-07-21',1,true,'Email all active clients by Monday'],
  ['Monthly budget review','Monthly Reporting','Finance','Morgan Chen','High','Medium','Monthly','2026-01-01','2026-06-30',4,true,'Compare actuals vs budget'],
  ['Weekly website backup check','Team Operations','Product','Jordan Patel','Medium','Quick','Weekly','2026-01-05','2026-07-21',0.5,true,'Verify backup completed successfully'],
  ['Yearly policy review','Team Operations','Human Resources','Avery Walker','Medium','Large','Yearly','2026-01-01','','8',true,'Review all 12 HR policies'],
  ['Biweekly sprint planning','Product Launch','Product','Jordan Patel','High','Medium','Every 2 Weeks','2026-01-12','2026-07-14',2,true,'Prioritize backlog with team'],
  ['Monthly newsletter creation','Social Media Campaign','Marketing','Riley Davis','Medium','Medium','Monthly','2026-01-01','2026-06-30',5,true,'Design + copy + send'],
  ['Weekly sales pipeline review','Monthly Reporting','Sales','Casey Rivera','High','Small','Weekly','2026-01-05','2026-07-21',1,true,'Update CRM forecast'],
  ['Quarterly team goal check-in','Team Operations','Leadership','Alex Morgan','High','Medium','Quarterly','2026-01-01','2026-03-31',3,true,'OKR progress review'],
  ['Monthly training content update','Training Program','Human Resources','Avery Walker','Low','Medium','Monthly','2026-02-01','2026-06-30',4,true,'Update one training module per month'],
];

// Next due date formula by recurrence type
function nextDueFormula(r) {
  // Uses EDATE, WORKDAY, and conditional logic; all Excel-compatible
  return `=IFERROR(IF(M${r}=FALSE,"",IF(J${r}="",I${r},IF(H${r}="Daily",J${r}+1,IF(H${r}="Weekdays",WORKDAY(J${r},1),IF(H${r}="Weekly",J${r}+7,IF(H${r}="Every 2 Weeks",J${r}+14,IF(H${r}="Monthly",EDATE(J${r},1),IF(H${r}="Quarterly",EDATE(J${r},3),IF(H${r}="Yearly",EDATE(J${r},12),J${r}+7))))))))),"")`;
}

function readyToAddFormula(r) {
  // Yes when: active=TRUE, next due <= today, and no incomplete task with same REC ID in Master Task Log
  return `=IFERROR(IF(AND(M${r}=TRUE,K${r}<>"",K${r}<=TODAY(),COUNTIFS('Master Task Log'!$Q$6:$Q$505,A${r},'Master Task Log'!$K$6:$K$505,"<>Completed",'Master Task Log'!$K$6:$K$505,"<>Cancelled")=0),"Yes","No"),"No")`;
}

function instructionsFormula(r) {
  return `=IF(N${r}="Yes","Copy row to Master Task Log — task is due now.","Next due: "&IFERROR(TEXT(K${r},"MMM D, YYYY"),"—"))`;
}

(async () => {
  // Dynamic pre-pass: unmerge any existing merges in the first 6 rows so re-runs are safe
  const auth2 = (() => {
    const s = JSON.parse(fs.readFileSync(__dirname + '/client_secret.json'));
    const { client_id, client_secret, redirect_uris } = s.installed;
    const a = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
    a.setCredentials(JSON.parse(fs.readFileSync(__dirname + '/tokens.json')));
    return a;
  })();
  const sheets2 = google.sheets({ version: 'v4', auth: auth2 });
  const spInfo = await sheets2.spreadsheets.get({ spreadsheetId:id, fields:'sheets(properties.sheetId,merges)' });
  const rpSheet = (spInfo.data.sheets || []).find(sh => sh.properties && sh.properties.sheetId === RP);
  const existingMerges = (rpSheet && rpSheet.merges) ? rpSheet.merges.filter(m => m.startRowIndex < 6) : [];
  if (existingMerges.length > 0) {
    await sheets2.spreadsheets.batchUpdate({ spreadsheetId:id, requestBody:{ requests: existingMerges.map(range => ({ unmergeCells:{ range } })) } });
  }

  const data = [];

  data.push({ range:`${S}!A1`, values:[['RECURRING TASK PLANNER']] });
  data.push({ range:`${S}!A2`, values:[['Define repeating responsibilities here. This planner calculates the next due date and provides a ready-to-copy task row for the Master Task Log.']] });
  data.push({ range:`${S}!A3`, values:[['Note: Copy ready tasks manually into the Master Task Log. Without macros, this planner cannot append rows automatically.']] });

  // Headers row 5
  data.push({ range:`${S}!A5:Q5`, values:[[
    'Recurrence ID','Task Name','Project','Department','Assigned To','Priority','Effort',
    'Recurrence Type','Start Date','Last Completed','Next Due Date','Est. Hours',
    'Active?','Ready to Add?','Generated Task ID','Instructions','Notes'
  ]] });

  // Formula rows 6-20
  for (let r = 6; r <= 20; r++) {
    data.push({ range:`${S}!A${r}`, values:[[`=IF(B${r}="","","REC-"&TEXT(ROW()-5,"0000"))`]] });
    data.push({ range:`${S}!K${r}`, values:[[nextDueFormula(r)]] });
    data.push({ range:`${S}!N${r}`, values:[[readyToAddFormula(r)]] });
    data.push({ range:`${S}!O${r}`, values:[[`=IF(B${r}="","",IF(N${r}="Yes","TSK-"&TEXT(COUNTIF('Master Task Log'!$A$6:$A$505,"<>")+1,"0000"),"—"))`]] });
    data.push({ range:`${S}!P${r}`, values:[[instructionsFormula(r)]] });
  }

  // Sample data
  RECURRING.forEach(([name,proj,dept,assigned,pri,effort,recType,start,lastComp,est,active,notes], i) => {
    const r = 6 + i;
    data.push({ range:`${S}!B${r}:M${r}`, values:[[name,proj,dept,assigned,pri,effort,recType,start,lastComp,est,active,'']] });
    if (notes) data.push({ range:`${S}!Q${r}`, values:[[notes]] });
  });

  await valuesBatchUpdate(id, data, 'recurring-values');

  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(RP,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID', color:hex(C.border) };
  const dateF  = { type:'DATE', pattern:'MMM D, YYYY' };

  // Title (row 0 = A1 only)
  reqs.push({ mergeCells:{ range:gridRange(RP,0,1,0,17), mergeType:'MERGE_ALL' } });
  fmt(0,1,0,17,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:18, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  // Description (row 1 = A2)
  reqs.push({ mergeCells:{ range:gridRange(RP,1,2,0,17), mergeType:'MERGE_ALL' } });
  fmt(1,2,0,17,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  // Warning note (row 2 = A3)
  reqs.push({ mergeCells:{ range:gridRange(RP,2,3,0,17), mergeType:'MERGE_ALL' } });
  fmt(2,3,0,17,{ userEnteredFormat:{ backgroundColor:hex(C.warning), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:9, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP' } });
  // Spacer row (row 3 = A4)
  fmt(3,4,0,17,{ userEnteredFormat:{ backgroundColor:hex(C.bg) } });

  // Headers row 5
  fmt(4,5,0,17,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });

  // Data rows
  for (let r = 5; r < 20; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,17,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  // Formula cols A(0), K(10), N(13), O(14), P(15) — pale blue
  for (const ci of [0,10,13,14,15]) {
    reqs.push({ repeatCell:{ range:gridRange(RP,5,20,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 } } }, fields:'userEnteredFormat' } });
  }
  // Input cols
  for (const ci of [1,2,3,4,5,6,7,11,16]) {
    reqs.push({ repeatCell:{ range:gridRange(RP,5,20,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' } });
  }
  // Date cols I(8), J(9), K(10)
  for (const ci of [8,9,10]) {
    reqs.push({ repeatCell:{ range:gridRange(RP,5,20,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:dateF, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  }
  // Checkboxes for Active col M (12)
  for (let r = 5; r < 20; r++) {
    reqs.push({ setDataValidation:{ range:gridRange(RP,r,r+1,12,13), rule:{ condition:{ type:'BOOLEAN' }, showCustomUi:true } } });
  }
  // Ready to Add col N (13) — colored
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(RP,5,20,13,14)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:'Yes' }] }, format:{ backgroundColor:hex(C.success), textFormat:{ foregroundColor:hex(C.white), bold:true } } } }, index:0 } });
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(RP,5,20,13,14)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:'No' }] }, format:{ backgroundColor:hex(C.bg) } } }, index:0 } });

  reqs.push({ updateBorders:{ range:gridRange(RP,4,20,0,17), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  [[0,90],[1,200],[2,140],[3,120],[4,120],[5,80],[6,80],[7,120],[8,100],[9,110],[10,100],[11,80],[12,70],[13,100],[14,100],[15,260],[16,180]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:RP, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:RP, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:36 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:RP, dimension:'ROWS', startIndex:1, endIndex:5 }, properties:{ pixelSize:24 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:RP, dimension:'ROWS', startIndex:5, endIndex:30 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  reqs.push({ updateSheetProperties:{ properties:{ sheetId:RP, gridProperties:{ frozenRowCount:5 } }, fields:'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'recurring-format');
  console.log('Recurring Task Planner complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
