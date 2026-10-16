'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TS = sheetMap['Team Setup'];
const S = "'Team Setup'";

const MEMBER_DATA = [
  ['Alex Morgan','Leadership','Team Lead','alex@brightpath.com',25,true,''],
  ['Jamie Lee','Operations','Operations Manager','jamie@brightpath.com',20,true,''],
  ['Taylor Brooks','Marketing','Marketing Lead','taylor@brightpath.com',18,true,''],
  ['Jordan Patel','Product','Product Manager','jordan@brightpath.com',22,true,''],
  ['Casey Rivera','Sales','Sales Associate','casey@brightpath.com',20,true,''],
  ['Morgan Chen','Finance','Finance Analyst','morgan@brightpath.com',15,true,''],
  ['Riley Davis','Creative','Designer','riley@brightpath.com',18,true,''],
  ['Sam Thompson','Customer Service','Support Lead','sam@brightpath.com',20,true,''],
  ['Avery Walker','Human Resources','HR Coordinator','avery@brightpath.com',15,true,''],
  ['Quinn Martinez','Administration','Admin Assistant','quinn@brightpath.com',18,true,''],
  ['Member 11','General Team','','',20,false,''],
  ['Member 12','General Team','','',20,false,''],
  ['Member 13','General Team','','',20,false,''],
  ['Member 14','General Team','','',20,false,''],
  ['Member 15','General Team','','',20,false,''],
];

const PROJECT_DATA = [
  ['Website Redesign','Taylor Brooks','Marketing','2026-01-15','2026-06-30','In Progress','High',true,'Full site overhaul'],
  ['Product Launch','Jordan Patel','Product','2026-02-01','2026-07-31','In Progress','Urgent',true,'Q3 product release'],
  ['Client Onboarding','Sam Thompson','Customer Service','2026-01-01','2026-12-31','In Progress','Medium',true,'Ongoing onboarding'],
  ['Monthly Reporting','Morgan Chen','Finance','2026-01-01','2026-12-31','In Progress','Medium',true,'Monthly financials'],
  ['Social Media Campaign','Taylor Brooks','Marketing','2026-03-01','2026-09-30','In Progress','High',true,'Q2-Q3 campaign'],
  ['Team Operations','Jamie Lee','Operations','2026-01-01','2026-12-31','In Progress','Medium',true,'Ongoing ops'],
  ['Event Planning','Avery Walker','Human Resources','2026-04-01','2026-08-15','Not Started','Medium',true,'Annual team event'],
  ['Process Improvement','Jamie Lee','Operations','2026-02-01','2026-05-31','In Progress','High',true,'Workflow audit'],
  ['Training Program','Avery Walker','Human Resources','2026-03-01','2026-10-31','Not Started','Low',true,'Staff development'],
  ['General Tasks','Alex Morgan','Leadership','2026-01-01','2026-12-31','In Progress','Low',true,'Miscellaneous team tasks'],
  ['','','','','','','',false,''],
  ['','','','','','','',false,''],
  ['','','','','','','',false,''],
  ['','','','','','','',false,''],
  ['','','','','','','',false,''],
];

(async () => {
  const data = [];

  data.push({ range:`${S}!A1`, values:[['TEAM SETUP & WORKSPACE SETTINGS']] });
  data.push({ range:`${S}!A3`, values:[['Update team members and projects here before using the tracker. Keep names consistent so task assignments and project reports remain connected.']] });

  data.push({ range:`${S}!A5:B5`, values:[['TEAM INFORMATION','']] });
  data.push({ range:`${S}!A6`, values:[['Team Name:']] });
  data.push({ range:`${S}!B6`, values:[['BrightPath Operations']] });
  data.push({ range:`${S}!A7`, values:[['Organization:']] });
  data.push({ range:`${S}!B7`, values:[['BrightPath Studio']] });
  data.push({ range:`${S}!A8`, values:[['Team Lead:']] });
  data.push({ range:`${S}!B8`, values:[['Alex Morgan']] });
  data.push({ range:`${S}!A9`, values:[['Workweek Start:']] });
  data.push({ range:`${S}!B9`, values:[['Monday']] });
  data.push({ range:`${S}!A10`, values:[['Reporting Year:']] });
  data.push({ range:`${S}!B10`, values:[[2026]] });
  data.push({ range:`${S}!A11`, values:[['Default Weekly Capacity (tasks):']] });
  data.push({ range:`${S}!B11`, values:[[20]] });
  data.push({ range:`${S}!A12`, values:[['Active Project Count:']] });
  data.push({ range:`${S}!B12`, values:[['=IFERROR(COUNTIF(B30:B44,"<>"),0)']] });
  data.push({ range:`${S}!A13`, values:[['Active Team Member Count:']] });
  data.push({ range:`${S}!B13`, values:[['=IFERROR(COUNTIF(J19:J33,TRUE),0)']] });

  // Team member table
  data.push({ range:`${S}!A18`, values:[['TEAM MEMBER TABLE']] });
  data.push({ range:`${S}!A19:H19`, values:[['Member ID','Team Member','Department','Role','Email / Contact','Weekly Task Capacity','Active?','Notes']] });

  MEMBER_DATA.forEach(([name,dept,role,email,cap,active,notes], i) => {
    const r = 20 + i;
    data.push({ range:`${S}!A${r}`, values:[[`=IF(B${r}="","","MBR-"&TEXT(ROW()-19,"00"))`]] });
    data.push({ range:`${S}!B${r}:H${r}`, values:[[name,dept,role,email,cap,active,notes]] });
  });

  // Project table
  data.push({ range:`${S}!A28`, values:[['PROJECT SETUP TABLE']] });
  data.push({ range:`${S}!A29:J29`, values:[['Project ID','Project Name','Project Owner','Department','Start Date','Target End Date','Project Status','Project Priority','Active?','Notes']] });

  PROJECT_DATA.forEach(([name,owner,dept,start,end,status,priority,active,notes], i) => {
    const r = 30 + i;
    data.push({ range:`${S}!A${r}`, values:[[`=IF(B${r}="","","PRJ-"&TEXT(ROW()-29,"00"))`]] });
    data.push({ range:`${S}!B${r}:J${r}`, values:[[name,owner,dept,start,end,status,priority,active,notes]] });
  });

  await valuesBatchUpdate(id, data, 'teamsetup-values');

  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(TS,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID', color:hex(C.border) };

  // Title
  reqs.push({ mergeCells:{ range:gridRange(TS,0,2,0,10), mergeType:'MERGE_ALL' } });
  fmt(0,2,0,10,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:20, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Guidance
  reqs.push({ mergeCells:{ range:gridRange(TS,2,4,0,10), mergeType:'MERGE_ALL' } });
  fmt(2,4,0,10,{ userEnteredFormat:{ backgroundColor:hex(C.warning), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:9, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:TS, dimension:'ROWS', startIndex:2, endIndex:4 }, properties:{ pixelSize:28 }, fields:'pixelSize' } });

  // Team info section
  reqs.push({ mergeCells:{ range:gridRange(TS,4,5,0,1), mergeType:'MERGE_ALL' } });
  fmt(4,5,0,2,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(5,13,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial' } } });
  fmt(5,13,1,2,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:9, fontFamily:'Arial' } } });
  reqs.push({ updateBorders:{ range:gridRange(TS,4,13,0,2), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Member table section header
  reqs.push({ mergeCells:{ range:gridRange(TS,17,18,0,7), mergeType:'MERGE_ALL' } });
  fmt(17,18,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(18,19,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 19; r < 34; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,8,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  fmt(19,34,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 }, horizontalAlignment:'CENTER' } });
  reqs.push({ updateBorders:{ range:gridRange(TS,18,34,0,8), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Checkboxes Active? col G (index 6)
  for (let r = 19; r < 34; r++) {
    reqs.push({ setDataValidation:{ range:gridRange(TS,r,r+1,6,7), rule:{ condition:{ type:'BOOLEAN' }, showCustomUi:true } } });
  }

  // Project table section header
  reqs.push({ mergeCells:{ range:gridRange(TS,27,28,0,9), mergeType:'MERGE_ALL' } });
  fmt(27,28,0,10,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(28,29,0,10,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 29; r < 44; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,10,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  fmt(29,44,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 }, horizontalAlignment:'CENTER' } });
  const dateF = { type:'DATE', pattern:'MMM D, YYYY' };
  for (const ci of [4,5]) {
    reqs.push({ repeatCell:{ range:gridRange(TS,29,44,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:dateF, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  }
  // Active checkboxes col I (index 8)
  for (let r = 29; r < 44; r++) {
    reqs.push({ setDataValidation:{ range:gridRange(TS,r,r+1,8,9), rule:{ condition:{ type:'BOOLEAN' }, showCustomUi:true } } });
  }
  reqs.push({ updateBorders:{ range:gridRange(TS,28,44,0,10), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  [[0,80],[1,140],[2,120],[3,130],[4,160],[5,90],[6,70],[7,70],[8,70],[9,160]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:TS, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:TS, dimension:'ROWS', startIndex:0, endIndex:2 }, properties:{ pixelSize:40 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:TS, dimension:'ROWS', startIndex:2, endIndex:80 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  // Freeze 4 rows to include both title merge (rows 0-1) and guidance merge (rows 2-3)
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:TS, gridProperties:{ frozenRowCount:4 } }, fields:'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'teamsetup-format');
  console.log('Team Setup complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
