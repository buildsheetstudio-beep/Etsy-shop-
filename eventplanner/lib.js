'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const secret = JSON.parse(fs.readFileSync(__dirname + '/client_secret.json'));
const { client_id, client_secret, redirect_uris } = secret.installed;
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
const tokens = JSON.parse(fs.readFileSync(__dirname + '/tokens.json'));
oAuth2Client.setCredentials(tokens);
const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });

const C = {
  deepBlush:    '#8B4A6B',
  rose:         '#A5607A',
  antiqueGold:  '#B8860B',
  champagne:    '#FDF5F0',
  ivory:        '#FFFDF7',
  paleRose:     '#F8F0F4',
  sageGreen:    '#5A7A4A',
  deepCrimson:  '#9B2335',
  darkPlum:     '#2A1A20',
  white:        '#FFFFFF',
  dustyRose:    '#E8D5DC',
  warmWhite:    '#FFF8F0',
  softBlush:    '#FFE0EC',
  lightGrey:    '#F5F5F5',
  gold:         '#B8860B',
  plum:         '#7A4A6B',
  softPurple:   '#6B4A8B',
  warmBrown:    '#8B6A3A',
  slateBlue:    '#3A6B8B',
  tealSage:     '#4A6B5A',
  dustyBrown:   '#8B4A4A',
  nearBlack:    '#1C1C1E',
  amber:        '#B5830A',
  errorBg:      '#FDECEA',
  errorFg:      '#9B2335',
  doneBg:       '#E8F5E9',
  doneFg:       '#2E7D32',
  warnBg:       '#FFF8E1',
  warnFg:       '#F57F17',
};

function hex(h) {
  const c = h.replace('#', '');
  return {
    red:   parseInt(c.slice(0, 2), 16) / 255,
    green: parseInt(c.slice(2, 4), 16) / 255,
    blue:  parseInt(c.slice(4, 6), 16) / 255,
  };
}

function gridRange(sheetId, r0, r1, c0, c1) {
  return { sheetId, startRowIndex: r0, endRowIndex: r1, startColumnIndex: c0, endColumnIndex: c1 };
}

function colLetter(n) {
  let s = ''; n++;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
  return s;
}

const CHUNK = 400;
async function batchUpdate(spreadsheetId, requests, label) {
  for (let i = 0; i < requests.length; i += CHUNK) {
    const chunk = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: chunk } });
    if (requests.length > CHUNK) console.log(`  ${label} chunk ${Math.floor(i/CHUNK)+1}/${Math.ceil(requests.length/CHUNK)}`);
  }
}

const VCHUNK = 200;
async function valuesBatchUpdate(spreadsheetId, data, label) {
  for (let i = 0; i < data.length; i += VCHUNK) {
    const chunk = data.slice(i, i + VCHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: chunk },
    });
    if (data.length > VCHUNK) console.log(`  ${label} chunk ${Math.floor(i/VCHUNK)+1}/${Math.ceil(data.length/VCHUNK)}`);
  }
}

module.exports = { sheets, batchUpdate, valuesBatchUpdate, hex, C, gridRange, colLetter };
