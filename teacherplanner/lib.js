'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SECRET_PATH = path.join(__dirname, 'client_secret.json');
const TOKEN_PATH  = path.join(__dirname, 'tokens.json');

function getAuth() {
  const { client_secret, client_id, redirect_uris } =
    JSON.parse(fs.readFileSync(SECRET_PATH)).installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
  return oAuth2Client;
}

const sheets = google.sheets({ version: 'v4', auth: getAuth() });

// Soft Classroom Pastel theme
const C = {
  // Interface
  text:      '#2F3437',
  secText:   '#6E7478',
  bg:        '#F8F7F4',
  panel:     '#FFFFFF',
  border:    '#CBC9C4',
  success:   '#93AF97',
  warning:   '#D4B46B',
  attention: '#BC7A74',
  info:      '#AEC7D8',
  white:     '#FFFFFF',
  input:     '#FEFCE8',
  formula:   '#EEF4F8',
  altRow:    '#F3F2EF',
  // Subject pastels
  ELA:       '#C9DDF2',
  Writing:   '#DED2EC',
  Math:      '#CDE5D5',
  Science:   '#C5E1DD',
  SocialStudies: '#E9D7B8',
  Art:       '#EBCFD9',
  Music:     '#D8D2EA',
  PE:        '#F2D0BA',
  Technology:'#D2E2F0',
  WorldLang: '#EADDAA',
  SpEd:      '#D4E0C8',
  Advisory:  '#E4E2DF',
};

function hex(h) {
  const r = parseInt(h.slice(1,3),16)/255;
  const g = parseInt(h.slice(3,5),16)/255;
  const b = parseInt(h.slice(5,7),16)/255;
  return { red: r, green: g, blue: b };
}

async function batchUpdate(id, requests, label) {
  const CHUNK = 400;
  for (let i = 0; i < requests.length; i += CHUNK) {
    const chunk = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: id, requestBody: { requests: chunk } });
    if (label) console.log(`  [${label}] batch ${Math.floor(i/CHUNK)+1}/${Math.ceil(requests.length/CHUNK)}`);
    if (i + CHUNK < requests.length) await new Promise(r => setTimeout(r, 300));
  }
}

async function valuesBatchUpdate(id, data, label) {
  const CHUNK = 200;
  for (let i = 0; i < data.length; i += CHUNK) {
    const chunk = data.slice(i, i + CHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: id,
      requestBody: { valueInputOption: 'USER_ENTERED', data: chunk },
    });
    if (label) console.log(`  [${label}] values batch ${Math.floor(i/CHUNK)+1}/${Math.ceil(data.length/CHUNK)}`);
    if (i + CHUNK < data.length) await new Promise(r => setTimeout(r, 300));
  }
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

module.exports = { getAuth, sheets, C, hex, batchUpdate, valuesBatchUpdate, gridRange };
