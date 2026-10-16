'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const KEY   = JSON.parse(fs.readFileSync(path.resolve(__dirname, '../projecttracker/client_secret.json')));
const TOKEN = JSON.parse(fs.readFileSync(path.resolve(__dirname, '../projecttracker/tokens.json')));

const auth = new google.auth.OAuth2(KEY.installed.client_id, KEY.installed.client_secret);
auth.setCredentials(TOKEN);
const sheets = google.sheets({ version: 'v4', auth });

// Multi-pastel sinking funds palette
const C = {
  // Pastel fund card colors (10 rotating)
  seafoam:    '#AFCFC7',
  dustyBlue:  '#B7C9DF',
  softLilac:  '#CFC2DE',
  mutedRose:  '#D9B9BB',
  paleOlive:  '#C8CAA7',
  softPeach:  '#E6C3AD',
  powderAqua: '#BDD8D9',
  mistyPeri:  '#C3C8E3',
  warmBlush:  '#DEC4CC',
  softSage:   '#C3D4BC',
  // Neutral UI
  bg:         '#F7F4EE',
  panel:      '#FFFFFF',
  text:       '#303234',
  secText:    '#727477',
  border:     '#CBC7BF',
  success:    '#90AA91',
  warning:    '#D1AF69',
  attention:  '#B67670',
  info:       '#AFC3D2',
  // Derived darks for headers
  primary:    '#5A8880',  // deep seafoam
  secondary:  '#4A6E8A',  // deep dusty blue
  // Input / formula cells
  input:      '#FFFDF5',
  formula:    '#EBF2F5',
  altRow:     '#F5F2EC',
  white:      '#FFFFFF',
  gray:       '#BBBBBB',
  darkText:   '#1A1C1E',
  // Aliases used in tab scripts
  stripeBg:   '#F5F2EC',   // = altRow
  lightGray:  '#F0EFED',
  inputBg:    '#FFFDF5',   // = input
  borderLight:'#E5E1DC',
  textMid:    '#727477',   // = secText
  textDark:   '#1A1C1E',   // = darkText
};

// 10 pastel card colors in cycle order
C.cardPalette = [
  C.seafoam, C.dustyBlue, C.softLilac, C.mutedRose, C.paleOlive,
  C.softPeach, C.powderAqua, C.mistyPeri, C.warmBlush, C.softSage,
];

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
