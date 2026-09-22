'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const REF = sheetMap['📋 Reference Data'];

(async () => {
  const values = [];
  const reqs = [];

  // Col A: Task Categories
  values.push({ range: '\'📋 Reference Data\'!A1', values: [['Task Category']] });
  values.push({ range: '\'📋 Reference Data\'!A2', values: [
    ['8 Weeks Out'],['6 Weeks Out'],['4 Weeks Out'],['2 Weeks Out'],
    ['Moving Week'],['Moving Day'],['Post-Move'],['Admin'],
  ]});

  // Col B: Priority
  values.push({ range: '\'📋 Reference Data\'!B1', values: [['Priority']] });
  values.push({ range: '\'📋 Reference Data\'!B2', values: [['High'],['Medium'],['Low']] });

  // Col C: Task Status
  values.push({ range: '\'📋 Reference Data\'!C1', values: [['Task Status']] });
  values.push({ range: '\'📋 Reference Data\'!C2', values: [
    ['Not Started'],['In Progress'],['Done'],['Blocked'],
  ]});

  // Col D: Room Names
  values.push({ range: '\'📋 Reference Data\'!D1', values: [['Room']] });
  values.push({ range: '\'📋 Reference Data\'!D2', values: [
    ['Living Room'],['Kitchen'],['Master Bedroom'],['Bedroom 2'],['Bedroom 3'],
    ['Bathroom'],['Home Office'],['Garage'],['Storage'],['Laundry'],['Outdoor'],['Other'],
  ]});

  // Col E: Packing Priority
  values.push({ range: '\'📋 Reference Data\'!E1', values: [['Pack Priority']] });
  values.push({ range: '\'📋 Reference Data\'!E2', values: [
    ['Pack First'],['Standard'],['Pack Last'],
  ]});

  // Col F: Packing Status
  values.push({ range: '\'📋 Reference Data\'!F1', values: [['Pack Status']] });
  values.push({ range: '\'📋 Reference Data\'!F2', values: [
    ['To Pack'],['Packed'],['Loaded'],['Arrived'],['Unpacked'],
  ]});

  // Col G: Move Type
  values.push({ range: '\'📋 Reference Data\'!G1', values: [['Move Type']] });
  values.push({ range: '\'📋 Reference Data\'!G2', values: [
    ['Local'],['Interstate'],['International'],
  ]});

  // Col H: Budget Category
  values.push({ range: '\'📋 Reference Data\'!H1', values: [['Budget Category']] });
  values.push({ range: '\'📋 Reference Data\'!H2', values: [
    ['Moving Company'],['Packing Supplies'],['Storage Unit'],['Cleaning Services'],
    ['Repairs & Maintenance'],['New Furniture'],['Utility Connections'],
    ['Travel & Fuel'],['Temporary Accommodation'],['Miscellaneous'],
  ]});

  // Col I: Notification Method
  values.push({ range: '\'📋 Reference Data\'!I1', values: [['Notify Method']] });
  values.push({ range: '\'📋 Reference Data\'!I2', values: [
    ['Online'],['Phone'],['Letter'],['In Person'],['Email'],
  ]});

  // Col J: Address Category
  values.push({ range: '\'📋 Reference Data\'!J1', values: [['Address Category']] });
  values.push({ range: '\'📋 Reference Data\'!J2', values: [
    ['Government & Official'],['Financial'],['Subscriptions & Delivery'],
    ['Healthcare'],['Insurance'],['Work & Professional'],['Personal Contacts'],
  ]});

  // Col K: Utility Service Type
  values.push({ range: '\'📋 Reference Data\'!K1', values: [['Utility Service']] });
  values.push({ range: '\'📋 Reference Data\'!K2', values: [
    ['Electricity'],['Gas'],['Water'],['Internet'],['Mobile'],['TV Licence'],
    ['Home Insurance'],['Contents Insurance'],['Council Tax'],['Waste Collection'],
    ['Landline'],['Solar'],['Other'],
  ]});

  // Col L: Old Utility Status
  values.push({ range: '\'📋 Reference Data\'!L1', values: [['Old Utility Status']] });
  values.push({ range: '\'📋 Reference Data\'!L2', values: [
    ['Active'],['Cancellation Requested'],['Disconnected'],
  ]});

  // Col M: New Utility Status
  values.push({ range: '\'📋 Reference Data\'!M1', values: [['New Utility Status']] });
  values.push({ range: '\'📋 Reference Data\'!M2', values: [
    ['Not Started'],['Applied'],['Active'],['Connected'],
  ]});

  // Col N: School Type
  values.push({ range: '\'📋 Reference Data\'!N1', values: [['School Type']] });
  values.push({ range: '\'📋 Reference Data\'!N2', values: [
    ['Public'],['Private'],['Charter'],['Montessori'],['Religious'],['International'],
  ]});

  // Col O: School Status
  values.push({ range: '\'📋 Reference Data\'!O1', values: [['School Status']] });
  values.push({ range: '\'📋 Reference Data\'!O2', values: [
    ['Research'],['Contacted'],['Visited'],['Applied'],['Enrolled'],
  ]});

  // Col P: Childcare Type
  values.push({ range: '\'📋 Reference Data\'!P1', values: [['Childcare Type']] });
  values.push({ range: '\'📋 Reference Data\'!P2', values: [
    ['Long Day Care'],['Family Day Care'],['Preschool'],
    ['Occasional Care'],['After School Care'],
  ]});

  // Col Q: Childcare Status
  values.push({ range: '\'📋 Reference Data\'!Q1', values: [['Childcare Status']] });
  values.push({ range: '\'📋 Reference Data\'!Q2', values: [
    ['Research'],['Contacted'],['Visited'],['Applied'],['Enrolled'],['Waitlisted'],
  ]});

  // Col R: Timeline Status
  values.push({ range: '\'📋 Reference Data\'!R1', values: [['Timeline Status']] });
  values.push({ range: '\'📋 Reference Data\'!R2', values: [
    ['Not Started'],['In Progress'],['Done'],
  ]});

  // Col S: Condition
  values.push({ range: '\'📋 Reference Data\'!S1', values: [['Condition']] });
  values.push({ range: '\'📋 Reference Data\'!S2', values: [
    ['Excellent'],['Good'],['Fair'],['Poor'],
  ]});

  await valuesBatchUpdate(id, values, 'reference-values');

  // Format headers
  for (let c = 0; c < 19; c++) {
    reqs.push({
      repeatCell: {
        range: gridRange(REF, 0, 1, c, c + 1),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.deepForest),
            textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
      },
    });
  }

  reqs.push({
    autoResizeDimensions: {
      dimensions: { sheetId: REF, dimension: 'COLUMNS', startIndex: 0, endIndex: 19 },
    },
  });

  await batchUpdate(id, reqs, 'reference-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
