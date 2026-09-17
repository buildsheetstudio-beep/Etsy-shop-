'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const KEY   = JSON.parse(fs.readFileSync(path.resolve(__dirname, '../projecttracker/client_secret.json')));
const TOKEN = JSON.parse(fs.readFileSync(path.resolve(__dirname, '../projecttracker/tokens.json')));

const auth = new google.auth.OAuth2(KEY.installed.client_id, KEY.installed.client_secret);
auth.setCredentials(TOKEN);
const sheets = google.sheets({ version: 'v4', auth });

// Light Genealogy Pastels palette
const C = {
  primary:      '#9FB7CF', // Faded Heritage Blue
  secondary:    '#B9CCB7', // Soft Sage
  blush:        '#D9BCC2', // Branch Accent 1 - Dusty Blush
  lavender:     '#C8C0DB', // Branch Accent 2 - Pale Lavender
  wheat:        '#E1D0A3', // Branch Accent 3 - Soft Wheat
  aqua:         '#B8D2D0', // Branch Accent 4 - Misty Aqua
  bg:           '#F7F4EE', // Background Warm Milk
  panel:        '#FFFFFF',
  text:         '#343536',
  secText:      '#777779',
  border:       '#CBC7BF',
  info:         '#A9BDD0', // Informational
  review:       '#D1B36D', // Review/Warning
  conflict:     '#B98482', // Conflict
  confirmed:    '#91AA93', // Confirmed
  neutral:      '#B9B8B3', // Neutral
  input:        '#FDFAF4', // pale warm for input cells
  formula:      '#EEF2F7', // pale blue-gray for formula cells
  altRow:       '#FAF8F3', // alternating row fill
  white:        '#FFFFFF',
  // Deeper header shades
  primaryDeep:  '#6A8FAF', // deeper heritage blue for headers
  secondaryDeep:'#7A9E7A', // deeper sage
  blushDeep:    '#B89098',
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

function colL(i) {
  if (i < 26) return String.fromCharCode(65 + i);
  return String.fromCharCode(64 + Math.floor(i / 26)) + String.fromCharCode(65 + (i % 26));
}

const CHUNK_FMT  = 400;
const CHUNK_VALS = 200;

async function batchUpdate(spreadsheetId, requests, label='') {
  for (let i = 0; i < requests.length; i += CHUNK_FMT) {
    const chunk = requests.slice(i, i + CHUNK_FMT);
    const batch = Math.floor(i/CHUNK_FMT)+1;
    const total = Math.ceil(requests.length/CHUNK_FMT);
    console.log(`  [${label}] batch ${batch}/${total}`);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: chunk } });
  }
}

async function valuesBatchUpdate(spreadsheetId, data, label='') {
  for (let i = 0; i < data.length; i += CHUNK_VALS) {
    const chunk = data.slice(i, i + CHUNK_VALS);
    const batch = Math.floor(i/CHUNK_VALS)+1;
    const total = Math.ceil(data.length/CHUNK_VALS);
    console.log(`  [${label}] values batch ${batch}/${total}`);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: chunk },
    });
  }
}

module.exports = { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, colL, C };
