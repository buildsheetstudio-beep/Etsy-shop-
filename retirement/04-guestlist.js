'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GL = sheetMap['Guest List & RSVP'];
const RD = sheetMap['Reference Data'];
const S = "'Guest List & RSVP'";
const R = "'Reference Data'";

// 18 columns A-R
const HEADERS = ['Guest ID','Household / Organization','Guest Name','Guest Type','Relationship to Retiree','Phone','Email','Invitation Status','RSVP Status','# Attending','Dietary Restriction','Accessibility Notes','Gift / Contribution?','Gift Description','Thank-You Needed?','Seating / Group','Follow-Up Date','Notes'];

// 37 sample guests
const GUESTS = [
  // [Household, Name, Type, Relationship, Phone, Email, InvStatus, RSVP, #Att, Dietary, Access, Gift, GiftDesc, Seating]
  ['Collins Family','Patricia Collins','Family','Retiree (Guest of Honor)','(951) 555-0101','patricia.collins@email.com','Confirmed','Confirmed',1,'None','','TRUE','Gift basket from family','Head Table'],
  ['Collins Family','James Collins','Family','Spouse','(951) 555-0102','james.collins@email.com','Confirmed','Confirmed',1,'None','','FALSE','','Head Table'],
  ['Collins Family','Emily Collins','Family','Daughter','(951) 555-0103','emily.c@email.com','Confirmed','Confirmed',1,'Vegetarian','','FALSE','','Table 1'],
  ['Collins Family','Daniel Collins','Family','Son','(951) 555-0104','dan.collins@email.com','Confirmed','Confirmed',1,'None','','FALSE','','Table 1'],
  ['Northfield Community Services','Robert Marsh','Manager','Executive Director','(951) 555-0110','r.marsh@northfield.org','Invitation Sent','Confirmed',1,'None','','FALSE','','Table 2'],
  ['Northfield Community Services','Sandra Okafor','Coworker','Deputy Director','(951) 555-0111','s.okafor@northfield.org','Invitation Sent','Confirmed',1,'Gluten-Free','','TRUE','Office gift card pool','Table 2'],
  ['Northfield Community Services','Kevin Liu','Coworker','Program Manager','(951) 555-0112','k.liu@northfield.org','Invitation Sent','Confirmed',1,'None','','FALSE','','Table 2'],
  ['Northfield Community Services','Maria Vasquez','Coworker','Operations Coordinator','(951) 555-0113','m.vasquez@northfield.org','Invitation Sent','Confirmed',1,'Vegetarian','','FALSE','','Table 3'],
  ['Northfield Community Services','Thomas Brennan','Coworker','Senior Analyst','(951) 555-0114','t.brennan@northfield.org','Invitation Sent','Confirmed',1,'None','','FALSE','','Table 3'],
  ['Northfield Community Services','Priya Sharma','Coworker','Budget Analyst','(951) 555-0115','p.sharma@northfield.org','Invitation Sent','Confirmed',1,'Vegan','','FALSE','','Table 3'],
  ['Northfield Community Services','David Wu','Direct Report','Operations Specialist','(951) 555-0116','d.wu@northfield.org','Invitation Sent','Confirmed',1,'None','','TRUE','Custom plaque','Table 4'],
  ['Northfield Community Services','Angela Torres','Direct Report','Admin Assistant','(951) 555-0117','a.torres@northfield.org','Invitation Sent','Confirmed',1,'None','','FALSE','','Table 4'],
  ['Northfield Community Services','Marcus Johnson','Direct Report','Field Coordinator','(951) 555-0118','m.johnson@northfield.org','Invitation Sent','Maybe',1,'None','','FALSE','','Table 4'],
  ['Northfield Community Services','Fatima Al-Rashid','Coworker','HR Director','(951) 555-0119','f.alrashid@northfield.org','Invitation Sent','Confirmed',1,'Halal','','FALSE','','Table 5'],
  ['Northfield Community Services','Liam O\'Brien','Coworker','IT Manager','(951) 555-0120','l.obrien@northfield.org','Invitation Sent','Confirmed',1,'None','','FALSE','','Table 5'],
  ['Northfield Community Services','Yuki Tanaka','Coworker','Grant Writer','(951) 555-0121','y.tanaka@northfield.org','Invitation Sent','No Response',1,'None','','FALSE','',''],
  ['Northfield Community Services','Chris Patel','Coworker','Events Coordinator','(951) 555-0122','c.patel@northfield.org','Follow-Up Needed','No Response',1,'Nut-Free','','FALSE','',''],
  ['Northfield Community Services','Rachel Kim','Coworker','Communications Manager','(951) 555-0123','r.kim@northfield.org','Invitation Sent','Confirmed',1,'None','','FALSE','','Table 5'],
  ['Northfield Community Services','James Adeyemi','Coworker','Volunteer Coordinator','(951) 555-0124','j.adeyemi@northfield.org','Invitation Sent','Declined',0,'None','','FALSE','',''],
  ['Northfield Community Services','Helen Kowalski','Manager','Board Chair','(951) 555-0125','h.kowalski@northfield.org','Invitation Sent','Confirmed',1,'None','','TRUE','Framed certificate','Head Table'],
  ['Riverside City Council','Mayor Tom Fischer','Community Member','City Mayor','(951) 555-0130','tfischer@riverside.gov','Invitation Sent','Confirmed',1,'None','','TRUE','City commendation certificate','VIP Table'],
  ['Riverside City Council','Councilwoman Diana Parks','Community Member','City Councilmember','(951) 555-0131','dparks@riverside.gov','Invitation Sent','Maybe',1,'Vegetarian','','FALSE','',''],
  ['Collins Family Friends','Nancy & Bill Morgan','Friend','Family Friends','(951) 555-0140','nmorgan@email.com','Confirmed','Confirmed',2,'None','','TRUE','Wine set','Table 6'],
  ['Collins Family Friends','George Hendricks','Friend','Childhood Friend','(951) 555-0141','ghendricks@email.com','Invitation Sent','Confirmed',1,'None','','FALSE','','Table 6'],
  ['Collins Family Friends','Theresa Walsh','Friend','Neighbor','(951) 555-0142','twalsh@email.com','Invitation Sent','Confirmed',1,'Dairy-Free','Wheelchair accessible seating','FALSE','','Table 6'],
  ['Collins Family Friends','Sam & Joyce Rivera','Friend','Family Friends','(951) 555-0143','srivera@email.com','Invitation Sent','Confirmed',2,'None','','TRUE','Photo album','Table 7'],
  ['Collins Family Friends','Karen Olson','Neighbor','Next-door neighbor','(951) 555-0144','kolson@email.com','Invitation Sent','No Response',1,'None','','FALSE','',''],
  ['Collins Family Friends','Richard Lamb','Friend','Golf partner','(951) 555-0145','rlamb@email.com','Follow-Up Needed','No Response',1,'None','','FALSE','',''],
  ['Northfield Alumni','Dr. Elaine Sousa','Friend','Former Colleague','(951) 555-0150','esousa@university.edu','Invitation Sent','Confirmed',1,'Vegan','','TRUE','Book collection','Table 7'],
  ['Northfield Alumni','Fred & Linda Cho','Friend','Former Colleagues','(951) 555-0151','lcho@email.com','Invitation Sent','Confirmed',2,'None','','FALSE','','Table 8'],
  ['Northfield Alumni','Pastor James Weston','Community Member','Community Leader','(951) 555-0152','jweston@church.org','Invitation Sent','Confirmed',1,'None','','FALSE','','Table 8'],
  ['Northfield Alumni','Barbara Eaton','Friend','Retired Coworker','(951) 555-0153','beaton@email.com','Invitation Sent','Confirmed',1,'Gluten-Free','','TRUE','Garden set','Table 8'],
  ['Collins Family','Nora Collins','Family','Mother','(951) 555-0160','nora.collins@email.com','Confirmed','Confirmed',1,'None','Seated near exit','FALSE','','Table 1'],
  ['Collins Family','Charlie Collins','Family','Brother','(951) 555-0161','charlie.c@email.com','Confirmed','Confirmed',1,'None','','FALSE','','Table 1'],
  ['Collins Family','Grace & Mark Abbott','Family','Sister & Brother-in-law','(951) 555-0162','gabbott@email.com','Confirmed','Confirmed',2,'Kosher','','TRUE','Keepsake box','Table 1'],
  ['Northfield Community Services','Patricia Chang','Coworker','Finance Director','(951) 555-0170','p.chang@northfield.org','Invitation Sent','Confirmed',1,'None','','FALSE','','Table 9'],
  ['Northfield Community Services','Allan Ferreira','Coworker','Maintenance Supervisor','(951) 555-0171','a.ferreira@northfield.org','Invitation Sent','Confirmed',1,'None','','FALSE','','Table 9'],
];

(async () => {
  const reqs = [];
  const data = [];

  // ── Background ─────────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(GL,0,400,0,18),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) }},
    fields:'userEnteredFormat(backgroundColor)'
  }});

  // ── Freeze FIRST (before any merges that cross col boundary) ─────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId:GL, gridProperties: { frozenRowCount:5, frozenColumnCount:3 } },
    fields:'gridProperties.frozenRowCount,gridProperties.frozenColumnCount'
  }});

  // ── Title / subtitle / section rows ───────────────────────────────────────
  data.push({ range:`${S}!A1`, values:[['GUEST LIST & RSVP']] });
  data.push({ range:`${S}!A2`, values:[['Track every guest, RSVP status, dietary needs, and gift contributions']] });
  data.push({ range:`${S}!A3`, values:[['Auto-calculated: Guest ID and Thank-You Needed columns are formula-driven from your entries']] });
  data.push({ range:`${S}!A4`, values:[['HOW TO USE: Enter guest name (col C), then fill remaining columns. Dropdowns available for Guest Type, Invitation Status, RSVP Status, and Dietary Restriction.']] });

  // Title row (row 0): style cols 0-2 separately, merge cols 3-17
  reqs.push({ repeatCell: {
    range: gridRange(GL,0,1,0,3),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:16, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ mergeCells: { range: gridRange(GL,0,1,3,18), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(GL,0,1,3,18),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:16, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:GL, dimension:'ROWS', startIndex:0, endIndex:1 },
    properties: { pixelSize:44 }, fields:'pixelSize'
  }});

  // Subtitle (row 1): cols 0-2 + merge cols 3-17
  reqs.push({ repeatCell: {
    range: gridRange(GL,1,2,0,3),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { italic:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ mergeCells: { range: gridRange(GL,1,2,3,18), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(GL,1,2,3,18),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { italic:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});

  // Info rows 2-3 (indices 2,3): cols 0-2 + merge cols 3-17
  [2,3].forEach(ri => {
    reqs.push({ repeatCell: {
      range: gridRange(GL,ri,ri+1,0,3),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.bg),
        textFormat: { italic:true, fontSize:9, foregroundColor:hex(C.secText), fontFamily:'Arial' },
        horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE',
        wrapStrategy:'WRAP'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)'
    }});
    reqs.push({ mergeCells: { range: gridRange(GL,ri,ri+1,3,18), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(GL,ri,ri+1,3,18),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.bg),
        textFormat: { italic:true, fontSize:9, foregroundColor:hex(C.secText), fontFamily:'Arial' },
        horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE',
        wrapStrategy:'WRAP'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)'
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:GL, dimension:'ROWS', startIndex:ri, endIndex:ri+1 },
      properties: { pixelSize:24 }, fields:'pixelSize'
    }});
  });

  // ── Header row (row 4, index 4) ────────────────────────────────────────────
  data.push({ range:`${S}!A5`, values:[HEADERS] });
  reqs.push({ repeatCell: {
    range: gridRange(GL,4,5,0,18),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
      wrapStrategy:'WRAP'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:GL, dimension:'ROWS', startIndex:4, endIndex:5 },
    properties: { pixelSize:40 }, fields:'pixelSize'
  }});

  // ── Sample data rows 5-41 (sheet rows 6-42) ───────────────────────────────
  GUESTS.forEach(([household, name, type, rel, phone, email, invStatus, rsvp, att, diet, access, gift, giftDesc, seating], i) => {
    const row = 6 + i; // sheet row
    const ri = 5 + i;  // 0-indexed
    // A: Guest ID formula
    data.push({ range:`${S}!A${row}`, values:[[`=IF(C${row}="","","GST-"&TEXT(ROW()-5,"000"))`]] });
    // B-N data
    data.push({ range:`${S}!B${row}`, values:[[household]] });
    data.push({ range:`${S}!C${row}`, values:[[name]] });
    data.push({ range:`${S}!D${row}`, values:[[type]] });
    data.push({ range:`${S}!E${row}`, values:[[rel]] });
    data.push({ range:`${S}!F${row}`, values:[[phone]] });
    data.push({ range:`${S}!G${row}`, values:[[email]] });
    data.push({ range:`${S}!H${row}`, values:[[invStatus]] });
    data.push({ range:`${S}!I${row}`, values:[[rsvp]] });
    data.push({ range:`${S}!J${row}`, values:[[att]] });
    data.push({ range:`${S}!K${row}`, values:[[diet]] });
    data.push({ range:`${S}!L${row}`, values:[[access]] });
    data.push({ range:`${S}!M${row}`, values:[[gift === 'TRUE']] }); // will set as checkbox
    data.push({ range:`${S}!N${row}`, values:[[giftDesc]] });
    // O: Thank-You Needed formula
    data.push({ range:`${S}!O${row}`, values:[[`=IF(C${row}="","",IF(M${row}=TRUE,"Yes","No"))`]] });
    data.push({ range:`${S}!P${row}`, values:[[seating]] });
  });

  // ── Formula-only rows for remaining data rows 43-305 ──────────────────────
  for (let row = 43; row <= 305; row++) {
    data.push({ range:`${S}!A${row}`, values:[[`=IF(C${row}="","","GST-"&TEXT(ROW()-5,"000"))`]] });
    data.push({ range:`${S}!O${row}`, values:[[`=IF(C${row}="","",IF(M${row}=TRUE,"Yes","No"))`]] });
  }

  // ── Dropdowns ──────────────────────────────────────────────────────────────
  const dvReqs = [];

  // Guest Type col D (index 3)
  dvReqs.push({ setDataValidation: {
    range: gridRange(GL,5,305,3,4),
    rule: { condition: { type:'ONE_OF_RANGE', values:[{ userEnteredValue:`=${R}!$C$2:$C$10` }] }, showCustomUi:true, strict:false }
  }});

  // Invitation Status col H (index 7)
  dvReqs.push({ setDataValidation: {
    range: gridRange(GL,5,305,7,8),
    rule: { condition: { type:'ONE_OF_RANGE', values:[{ userEnteredValue:`=${R}!$A$2:$A$9` }] }, showCustomUi:true, strict:false }
  }});

  // RSVP Status col I (index 8)
  dvReqs.push({ setDataValidation: {
    range: gridRange(GL,5,305,8,9),
    rule: { condition: { type:'ONE_OF_RANGE', values:[{ userEnteredValue:`=${R}!$A$2:$A$9` }] }, showCustomUi:true, strict:false }
  }});

  // Dietary Restriction col K (index 10)
  dvReqs.push({ setDataValidation: {
    range: gridRange(GL,5,305,10,11),
    rule: { condition: { type:'ONE_OF_RANGE', values:[{ userEnteredValue:`=${R}!$W$2:$W$11` }] }, showCustomUi:true, strict:false }
  }});

  // Gift / Contribution checkbox col M (index 12)
  dvReqs.push({ setDataValidation: {
    range: gridRange(GL,5,305,12,13),
    rule: { condition: { type:'BOOLEAN' }, showCustomUi:true }
  }});

  // ── Alternating row background ─────────────────────────────────────────────
  for (let ri = 5; ri < 305; ri += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(GL,ri,ri+1,0,18),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) }},
      fields:'userEnteredFormat(backgroundColor)'
    }});
  }

  // ── Formula cell (col A, O) styling ───────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(GL,5,305,0,1),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula),
      textFormat: { fontSize:9, fontFamily:'Arial' },
      horizontalAlignment:'CENTER'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
  }});
  reqs.push({ repeatCell: {
    range: gridRange(GL,5,305,14,15),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula),
      textFormat: { fontSize:9, fontFamily:'Arial' },
      horizontalAlignment:'CENTER'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
  }});

  // Input cells styling (cols B-N excl formula cols A, O)
  reqs.push({ repeatCell: {
    range: gridRange(GL,5,305,1,14),
    cell: { userEnteredFormat: {
      textFormat: { fontSize:9, fontFamily:'Arial' },
      verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(textFormat,verticalAlignment)'
  }});

  // ── Conditional formatting: RSVP Status col I (index 8) ───────────────────
  const rsvpCF = [
    { val:'Confirmed',        bg: C.success,   text: C.mainText },
    { val:'Declined',         bg: C.border,    text: C.secText },
    { val:'Maybe',            bg: C.warning,   text: C.mainText },
    { val:'No Response',      bg: C.mutedMauve,text: C.mainText },
    { val:'Follow-Up Needed', bg: C.attention, text: C.white },
    { val:'Invitation Sent',  bg: C.mutedBlue, text: C.mainText },
  ];
  rsvpCF.forEach((rule, idx) => {
    reqs.push({ addConditionalFormatRule: {
      index: idx,
      rule: {
        ranges: [gridRange(GL,5,305,8,9)],
        booleanRule: {
          condition: { type:'TEXT_EQ', values:[{ userEnteredValue:rule.val }] },
          format: { backgroundColor: hex(rule.bg), textFormat: { foregroundColor: hex(rule.text) } }
        }
      }
    }});
  });

  // CF: full-row sage when RSVP=Confirmed
  reqs.push({ addConditionalFormatRule: {
    index: 6,
    rule: {
      ranges: [gridRange(GL,5,305,0,18)],
      booleanRule: {
        condition: { type:'CUSTOM_FORMULA', values:[{ userEnteredValue:`=$I6="Confirmed"` }] },
        format: { backgroundColor: hex('#E8F1E9') }
      }
    }
  }});

  // Tab color
  reqs.push({ updateSheetProperties: {
    properties: { sheetId:GL, tabColorStyle: { rgbColor: hex(C.primary) } },
    fields:'tabColorStyle'
  }});

  // Column widths
  const colW = [80,130,130,100,160,110,180,110,110,65,110,130,65,140,80,90,90,160];
  colW.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:GL, dimension:'COLUMNS', startIndex:i, endIndex:i+1 },
      properties: { pixelSize:w }, fields:'pixelSize'
    }});
  });

  await batchUpdate(id, reqs);
  await batchUpdate(id, dvReqs);
  await valuesBatchUpdate(id, data);
  console.log('Guest List complete');
})();
