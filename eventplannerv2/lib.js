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

// Interface colors
const C = {
  // UI
  text:      '#2F3336',
  secText:   '#6E7478',
  bg:        '#F7F6F3',
  panel:     '#FFFFFF',
  border:    '#C9C8C4',
  success:   '#92AE95',
  warning:   '#D5B56C',
  attention: '#BC7770',
  info:      '#AFC7D8',
  white:     '#FFFFFF',
  input:     '#FDFCF9',
  formula:   '#EEF4F8',
  altRow:    '#F2F1EE',
  // Event category pastels
  Business:    '#C9DCEF',
  Conference:  '#D8CDE8',
  Workshop:    '#F2D2B6',
  Networking:  '#C5E0DD',
  Community:   '#CBDCC5',
  Nonprofit:   '#E9DBA7',
  School:      '#D5E4F2',
  Fundraiser:  '#E8C0B8',
  Social:      '#E9CCD8',
  Birthday:    '#F3D0BB',
  Wedding:     '#DDD0DF',
  Family:      '#EEE2C8',
  Holiday:     '#D0E5D8',
  Sports:      '#C6D9E8',
  Virtual:     '#DDD5ED',
  Other:       '#E3E1DE',
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
  }
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

module.exports = { getAuth, sheets, C, hex, batchUpdate, valuesBatchUpdate, gridRange };
