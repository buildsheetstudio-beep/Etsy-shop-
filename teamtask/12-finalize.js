'use strict';
const { batchUpdate, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

// Tab order: Dashboard first, Reference Data last (hidden)
// Desired order: Team Dashboard, Master Task Log, Weekly Calendar, Monthly Calendar,
//                Project Progress, Recurring Task Planner, Team Setup, Reference Data

(async () => {
  const reqs = [];

  const tabColors = {
    'Team Dashboard':         C.primary,      // Deep Aubergine
    'Master Task Log':        C.secondary,    // Electric Teal
    'Weekly Calendar':        C.mutedBlue,    // Muted Blue
    'Monthly Calendar':       C.lavender,     // Lavender
    'Project Progress':       C.success,      // Sage Green
    'Recurring Task Planner': C.warning,      // Amber
    'Team Setup':             C.secText,      // Medium grey
    'Reference Data':         C.border,       // Light grey (hidden)
  };

  // Apply tab colors
  for (const [name, color] of Object.entries(tabColors)) {
    const sheetId = sheetMap[name];
    if (sheetId === undefined) continue;
    reqs.push({
      updateSheetProperties:{
        properties:{ sheetId, tabColorStyle:{ rgbColor:{ red: parseInt(color.slice(1,3),16)/255, green: parseInt(color.slice(3,5),16)/255, blue: parseInt(color.slice(5,7),16)/255 } } },
        fields:'tabColorStyle'
      }
    });
  }

  // Hide Reference Data
  const RD = sheetMap['Reference Data'];
  reqs.push({
    updateSheetProperties:{
      properties:{ sheetId:RD, hidden:true },
      fields:'hidden'
    }
  });

  // Reorder tabs
  const desiredOrder = [
    'Team Dashboard',
    'Master Task Log',
    'Weekly Calendar',
    'Monthly Calendar',
    'Project Progress',
    'Recurring Task Planner',
    'Team Setup',
    'Reference Data',
  ];
  desiredOrder.forEach((name, index) => {
    const sheetId = sheetMap[name];
    if (sheetId === undefined) return;
    reqs.push({
      updateSheetProperties:{
        properties:{ sheetId, index },
        fields:'index'
      }
    });
  });

  await batchUpdate(id, reqs, 'finalize');
  console.log('Finalize complete — tabs colored, reordered, Reference Data hidden');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
