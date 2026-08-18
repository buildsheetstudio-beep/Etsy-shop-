'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

// Shared credentials from projecttracker directory
const KEY   = JSON.parse(fs.readFileSync(path.resolve(__dirname, '../projecttracker/client_secret.json')));
const TOKEN = JSON.parse(fs.readFileSync(path.resolve(__dirname, '../projecttracker/tokens.json')));

const auth = new google.auth.OAuth2(KEY.installed.client_id, KEY.installed.client_secret);
auth.setCredentials(TOKEN);
const sheets = google.sheets({ version: 'v4', auth });

// ── Warm Natural Home palette ─────────────────────────────────────────────────
const C = {
  primary:    '#6F7860', // Weathered Olive Stone
  secondary:  '#B87558', // Burnt Clay
  bg:         '#F3F0E8', // Soft Limestone
  panel:      '#FFFFFF',
  olive:      '#E1E5D8', // Light Olive Tint
  clay:       '#EAD5CA', // Light Clay Tint
  wheat:      '#D7C69C', // Accent Wheat
  success:    '#8FA48A',
  warning:    '#D0AE67',
  attention:  '#B57269',
  info:       '#A8BCC6',
  text:       '#353632',
  secText:    '#74736C',
  border:     '#C9C4B9',
  white:      '#FFFFFF',
  gray:       '#B0AEA8',
  input:      '#FFFDF5', // pale warm yellow for input cells
  formula:    '#EBF2F5', // pale blue-gray for formula cells
  altRow:     '#F0EDE6',
};

function hex(colorStr) {
  const c = colorStr.replace('#', '');
  return {
    red:   parseInt(c.slice(0,2),16)/255,
    green: parseInt(c.slice(2,4),16)/255,
    blue:  parseInt(c.slice(4,6),16)/255,
  };
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

const CHUNK_FMT  = 400;
const CHUNK_VALS = 200;

async function batchUpdate(spreadsheetId, requests, label='') {
  for (let i = 0; i < requests.length; i += CHUNK_FMT) {
    const chunk = requests.slice(i, i + CHUNK_FMT);
    const batch = Math.floor(i/CHUNK_FMT)+1;
    const total = Math.ceil(requests.length/CHUNK_FMT);
    if (total > 1) console.log(`  [${label}] batch ${batch}/${total}`);
    else console.log(`  [${label}] batch 1/1`);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: chunk } });
  }
}

async function valuesBatchUpdate(spreadsheetId, data, label='') {
  for (let i = 0; i < data.length; i += CHUNK_VALS) {
    const chunk = data.slice(i, i + CHUNK_VALS);
    const batch = Math.floor(i/CHUNK_VALS)+1;
    const total = Math.ceil(data.length/CHUNK_VALS);
    if (total > 1) console.log(`  [${label}] values batch ${batch}/${total}`);
    else console.log(`  [${label}] values batch 1/1`);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: chunk },
    });
  }
}

module.exports = { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C };
