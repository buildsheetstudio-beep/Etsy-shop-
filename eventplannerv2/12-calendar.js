'use strict';
const { sheets, C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const path = require('path');
const { id } = JSON.parse(fs.readFileSync(path.join(__dirname, 'spreadsheet.json')));

const SID = 10;

const CAT_COLORS = {
  Business:'#C9DCEF', Conference:'#D8CDE8', Workshop:'#F2D2B6', Networking:'#C5E0DD',
  Community:'#CBDCC5', Nonprofit:'#E9DBA7', School:'#D5E4F2',  Fundraiser:'#E8C0B8',
  Social:'#E9CCD8',   Birthday:'#F3D0BB',   Wedding:'#DDD0DF',  Family:'#EEE2C8',
  Holiday:'#D0E5D8',  Sports:'#C6D9E8',     Virtual:'#DDD5ED',  Other:'#E3E1DE',
};

function cl(i) { if (i<26) return String.fromCharCode(65+i); return 'A'+String.fromCharCode(65+i-26); }

// Sample events per month for April through December 2026 (hardcoded for calendar display)
// Each event: [date(1-31), eventName, category, duration_label]
const CALENDAR_EVENTS_BY_MONTH = {
  4: [ // April 2026
    [15,'EVT-001 Leadership Summit (Day 1/3)','Conference','Multi-Day'],
    [16,'EVT-001 Leadership Summit (Day 2/3)','Conference','Multi-Day'],
    [17,'EVT-001 Leadership Summit (Day 3/3)','Conference','Multi-Day'],
  ],
  5: [ // May 2026
    [2,'EVT-002 Spring Gala','Fundraiser','Single-Day'],
    [15,'EVT-003 Community Health Fair','Community','Single-Day'],
    [20,'EVT-004 Digital Marketing Workshop','Workshop','Single-Day'],
    [27,'EVT-005 Board Meeting','Business','Single-Day'],
  ],
  6: [ // June 2026
    [5,'EVT-006 Youth Soccer (Day 1/2)','Sports','Multi-Day'],
    [6,'EVT-006 Youth Soccer (Day 2/2)','Sports','Multi-Day'],
    [10,'EVT-007 Virtual Career Fair','Virtual','Single-Day'],
    [17,'EVT-008 Networking Mixer','Networking','Single-Day'],
    [19,'EVT-009 School Science Fair','School','Single-Day'],
    [25,'EVT-010 Birthday Bash','Birthday','Single-Day'],
  ],
  7: [ // July 2026
    [4,'EVT-011 Holiday Block Party','Holiday','Single-Day'],
    [8,'EVT-012 Grant Writing Seminar','Nonprofit','Single-Day'],
    [15,'EVT-013 Hybrid Tech Conf (Day 1/3)','Conference','Multi-Day'],
    [16,'EVT-013 Hybrid Tech Conf (Day 2/3)','Conference','Multi-Day'],
    [17,'EVT-013 Hybrid Tech Conf (Day 3/3)','Conference','Multi-Day'],
    [18,'EVT-014 Family Reunion (Day 1/2)','Family','Multi-Day'],
    [19,'EVT-014 Family Reunion (Day 2/2)','Family','Multi-Day'],
    [22,'EVT-015 Volunteer Appreciation Dinner','Nonprofit','Single-Day'],
  ],
  8: [ // August 2026
    [1,'EVT-016 Back-to-School Fair','Community','Single-Day'],
    [7,'EVT-017 Wellness Retreat (Day 1/3)','Social','Multi-Day'],
    [8,'EVT-017 Wellness Retreat (Day 2/3)','Social','Multi-Day'],
    [9,'EVT-017 Wellness Retreat (Day 3/3)','Social','Multi-Day'],
    [12,'EVT-018 Business Plan Competition','Business','Single-Day'],
    [15,'EVT-019 Heritage Festival (Day 1/2)','Community','Multi-Day'],
    [16,'EVT-019 Heritage Festival (Day 2/2)','Community','Multi-Day'],
    [19,'EVT-020 Team Leadership Workshop','Workshop','Single-Day'],
    [25,'EVT-021 Fundraising Webinar','Virtual','Single-Day'],
  ],
  9: [ // September 2026
    [9,'EVT-022 Fall Golf Tournament','Sports','Single-Day'],
  ],
  10: [ // October 2026
    [17,'EVT-023 Kids Halloween Party','Birthday','Single-Day'],
    [21,'EVT-024 Board Retreat (Day 1/3)','Business','Multi-Day'],
    [22,'EVT-024 Board Retreat (Day 2/3)','Business','Multi-Day'],
    [23,'EVT-024 Board Retreat (Day 3/3)','Business','Multi-Day'],
  ],
  11: [ // November 2026
    [21,'EVT-025 Thanksgiving Feast','Community','Single-Day'],
  ],
  12: [ // December 2026
    [1,'EVT-027 Toy Drive Begins','Fundraiser','Multi-Day'],
    [11,'EVT-026 Year-End Celebration','Social','Single-Day'],
    [18,'EVT-027 Toy Drive Ends','Fundraiser','Multi-Day'],
    [30,'EVT-028 New Year Networking','Networking','Single-Day'],
  ],
};

// Month info [year, month(1-12), firstDayOfWeek(0=Sun,1=Mon,...), numDays]
const MONTHS = [
  [2026,4,3,30],  // April 2026 starts Wednesday
  [2026,5,5,31],  // May 2026 starts Friday
  [2026,6,1,30],  // June 2026 starts Monday
  [2026,7,3,31],  // July 2026 starts Wednesday
  [2026,8,6,31],  // August 2026 starts Saturday
  [2026,9,2,30],  // September 2026 starts Tuesday
  [2026,10,4,31], // October 2026 starts Thursday
  [2026,11,0,30], // November 2026 starts Sunday
  [2026,12,2,31], // December 2026 starts Wednesday
];

const MONTH_NAMES = ['','January','February','March','April','May','June','July','August','September','October','November','December'];
const DAY_HDRS = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

(async () => {
  const values = [];
  const fmt = [];

  // Title
  values.push({ range:"'Monthly Calendar'!A1", values:[['MONTHLY CALENDAR']] });
  values.push({ range:"'Monthly Calendar'!A2", values:[['Visual calendar showing all events with multi-day spans and pastel category colors. Navigate to any month to see your events.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,7), mergeType:'MERGE_ALL' }});
  fmt.push({ mergeCells:{ range: gridRange(SID,1,2,0,7), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,7), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3336'), textFormat:{bold:true,fontSize:16,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,7), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:9,foregroundColor:hex('#D5D8DB'),italic:true},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1},
    properties:{pixelSize:44}, fields:'pixelSize' }});

  // KPI summary bar row 3
  values.push({ range:"'Monthly Calendar'!A3", values:[[
    `=IFERROR("Total Events: "&COUNTA('Master Event Log'!B6:B500)&" | Confirmed: "&COUNTIF('Master Event Log'!F6:F500,"Confirmed")&" | Planning: "&COUNTIF('Master Event Log'!F6:F500,"Planning")&" | Completed: "&COUNTIF('Master Event Log'!F6:F500,"Completed"),"Loading...")`
  ]]});
  fmt.push({ mergeCells:{ range: gridRange(SID,2,3,0,7), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,2,3,0,7), cell:{userEnteredFormat:{
    backgroundColor:hex('#3D6374'), textFormat:{bold:true,fontSize:10,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:2,endIndex:3},
    properties:{pixelSize:28}, fields:'pixelSize' }});

  // Set column widths (7 day columns)
  for (let c=0; c<7; c++) {
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:c,endIndex:c+1},
      properties:{pixelSize:145}, fields:'pixelSize' }});
  }

  let currentRow = 3; // 0-indexed, starts at row 4

  MONTHS.forEach((monthInfo,mi)=>{
    const [year,month,firstDow,numDays] = monthInfo;
    const monthEvents = CALENDAR_EVENTS_BY_MONTH[month] || [];

    // Month header
    values.push({ range:`'Monthly Calendar'!A${currentRow+1}`, values:[[`${MONTH_NAMES[month]} ${year}`]] });
    fmt.push({ mergeCells:{ range: gridRange(SID,currentRow,currentRow+1,0,7), mergeType:'MERGE_ALL' }});
    fmt.push({ repeatCell:{ range: gridRange(SID,currentRow,currentRow+1,0,7), cell:{userEnteredFormat:{
      backgroundColor:hex('#3D6374'), textFormat:{bold:true,fontSize:13,foregroundColor:hex('#FFFFFF')},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:currentRow,endIndex:currentRow+1},
      properties:{pixelSize:36}, fields:'pixelSize' }});
    currentRow++;

    // Day headers
    values.push({ range:`'Monthly Calendar'!A${currentRow+1}`, values:[DAY_HDRS] });
    fmt.push({ repeatCell:{ range: gridRange(SID,currentRow,currentRow+1,0,7), cell:{userEnteredFormat:{
      backgroundColor:hex('#4A5056'), textFormat:{bold:true,fontSize:10,foregroundColor:hex('#FFFFFF')},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:currentRow,endIndex:currentRow+1},
      properties:{pixelSize:28}, fields:'pixelSize' }});
    currentRow++;

    // Calendar grid: 6 rows max
    const startRow = currentRow;
    let col = firstDow; // starting column (0=Sun)
    let day = 1;

    // 6 rows of calendar
    for (let week = 0; week < 6 && day <= numDays; week++) {
      const weekRow = startRow + week;

      // Fill date numbers and event names
      for (let d = (week === 0 ? firstDow : 0); d < 7 && day <= numDays; d++) {
        const cellRef = `${cl(d)}${weekRow+1}`;
        // Find events on this day
        const dayEvents = monthEvents.filter(e => e[0] === day);
        let cellContent = `${day}`;
        if (dayEvents.length > 0) {
          cellContent = `${day}\n` + dayEvents.map(e=>`• ${e[1]}`).join('\n');
        }
        values.push({ range:`'Monthly Calendar'!${cellRef}`, values:[[cellContent]] });

        const hasEvent = dayEvents.length > 0;
        const evCategory = hasEvent ? dayEvents[0][2] : null;
        const isMultiDay = hasEvent && dayEvents.some(e=>e[3]==='Multi-Day');
        const cellBg = hasEvent ? (CAT_COLORS[evCategory]||'#E3E1DE') : (d===0||d===6 ? '#F2F1EE' : '#FDFCF9');

        fmt.push({ repeatCell:{ range: gridRange(SID,weekRow,weekRow+1,d,d+1), cell:{userEnteredFormat:{
          backgroundColor:hex(cellBg),
          textFormat:{
            bold: hasEvent,
            fontSize: hasEvent ? 8 : 10,
            foregroundColor: hex(hasEvent ? '#2F3336' : (d===0||d===6 ? '#9A9896' : '#2F3336')),
          },
          verticalAlignment:'TOP',
          horizontalAlignment: hasEvent ? 'LEFT' : 'LEFT',
          wrapStrategy:'WRAP',
          padding:{left:4,top:4},
          borders:{
            top:{ style:'SOLID', color:hex('#C9C8C4'), width:1 },
            left:{ style:'SOLID', color:hex('#C9C8C4'), width:1 },
            bottom:{ style:'SOLID', color:hex('#C9C8C4'), width:1 },
            right:{ style:'SOLID', color:hex('#C9C8C4'), width:1 },
          },
        }}, fields:'userEnteredFormat' }});

        day++;
      }

      fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:weekRow,endIndex:weekRow+1},
        properties:{pixelSize: monthEvents.some(e=> {
          // check if any event falls in this week
          for(let dd=0;dd<7;dd++){
            const wday = week*7+dd-(week===0?firstDow:0)+1;
            if (monthEvents.some(ev=>ev[0]===wday)) return true;
          }
          return false;
        }) ? 80 : 60}, fields:'pixelSize' }});
    }

    currentRow += 6; // always 6 rows per month grid

    // Gap between months
    currentRow += 1;
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:currentRow-1,endIndex:currentRow},
      properties:{pixelSize:16}, fields:'pixelSize' }});
  });

  // Category legend at bottom
  values.push({ range:`'Monthly Calendar'!A${currentRow+1}`, values:[['EVENT CATEGORY COLOR LEGEND']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,currentRow,currentRow+1,0,7), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,currentRow,currentRow+1,0,7), cell:{userEnteredFormat:{
    backgroundColor:hex('#3D6374'), textFormat:{bold:true,fontSize:10,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:12},
  }}, fields:'userEnteredFormat' }});
  currentRow++;

  const cats = Object.keys(CAT_COLORS);
  // 2 rows of 8
  for (let i=0; i<cats.length; i++) {
    const col = i%7;
    const row = currentRow + Math.floor(i/7);
    values.push({ range:`'Monthly Calendar'!${cl(col)}${row+1}`, values:[[cats[i]]] });
    fmt.push({ repeatCell:{ range: gridRange(SID,row,row+1,col,col+1), cell:{userEnteredFormat:{
      backgroundColor:hex(CAT_COLORS[cats[i]]),
      textFormat:{bold:true,fontSize:9,foregroundColor:hex('#2F3336')},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
  }

  // Freeze title rows
  fmt.push({ updateSheetProperties:{ properties:{sheetId:SID,gridProperties:{frozenRowCount:3}},
    fields:'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, values, '12-calendar values');
  await batchUpdate(id, fmt, '12-calendar format');
  console.log('✅ Monthly Calendar done.');
})().catch(e => { console.error(e.errors||e.message||e); process.exit(1); });
