'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const KEY   = JSON.parse(fs.readFileSync(path.resolve(__dirname, '../projecttracker/client_secret.json')));
const TOKEN = JSON.parse(fs.readFileSync(path.resolve(__dirname, '../projecttracker/tokens.json')));

const auth = new google.auth.OAuth2(KEY.installed.client_id, KEY.installed.client_secret);
auth.setCredentials(TOKEN);
const sheets = google.sheets({ version: 'v4', auth });

// Soft Culinary Pastels palette
const C = {
  primary:      '#9FB4D1', // Dusty Cornflower
  secondary:    '#E2B59E', // Soft Apricot
  sage:         '#B8C9B5', // Accent Sage
  butter:       '#E8D8A7', // Accent Butter
  lavender:     '#C9BDD8', // Accent Lavender
  blush:        '#DDBFC7', // Accent Blush
  bg:           '#F8F5EE', // Background Vanilla
  panel:        '#FFFFFF',
  text:         '#333436',
  secText:      '#747477',
  border:       '#CCC8C1',
  success:      '#91AB93',
  warning:      '#D0AF6B',
  attention:    '#B87972',
  info:         '#ACC1D4',
  input:        '#FEFCF0', // pale warm yellow for input cells
  formula:      '#EEF3F8', // pale blue-gray for formula cells
  altRow:       '#FDFAF2', // alternating row fill
  white:        '#FFFFFF',
  // Deeper header shades
  primaryDeep:  '#6B8FB8', // deeper cornflower for headers
  secondaryDeep:'#C47858', // deeper apricot
  sageDeep:     '#7A9B78',
  // Category color map
  breakfastCol: '#E8D8A7', // butter
  lunchCol:     '#B8C9B5', // sage
  dinnerCol:    '#9FB4D1', // cornflower
  dessertCol:   '#C9BDD8', // lavender
  snackCol:     '#E2B59E', // apricot
  favoritesCol: '#DDBFC7', // blush
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
