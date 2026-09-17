'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

// Palette: Calm Focus Spectrum
const C = {
  slate:       '#344047',  // Deep Slate — text, headers
  coral:       '#D98278',  // Muted Coral — NOW
  blue:        '#8EAFCA',  // Powder Blue — NEXT
  lavender:    '#B3A3C7',  // Soft Lavender — LATER
  teal:        '#6F9E96',  // Dusty Teal — EVENTS
  sand:        '#D5B88F',  // Warm Sand — REVISIT / RECOVERY
  bg:          '#F7F3EC',  // Warm Ivory — background
  white:       '#FFFFFF',
  softGray:    '#E8E5E0',  // Soft Gray
  medGray:     '#B7B3AD',  // Medium Gray
  text:        '#344047',
  // Tints (section backgrounds)
  coralTint:   '#F5E5E3',
  blueTint:    '#E4EEF5',
  lavTint:     '#EDE9F4',
  tealTint:    '#E4EEEC',
  sandTint:    '#F5EBE0',
  formulaCell: '#EEEAE4',
  altRow:      '#F0EDE8',
  // Status colors
  completedBg: '#EBEBEB',
  archivedBg:  '#E0DDD9',
  waitingBg:   '#EEF0EE',
  rescheduledBg:'#F5EBE0',
  // Header variants
  headerDark:  '#2D3840',
  headerMid:   '#4A5A62',
};

function hex(h) {
  const r = parseInt(h.slice(1,3),16)/255;
  const g = parseInt(h.slice(3,5),16)/255;
  const b = parseInt(h.slice(5,7),16)/255;
  return { red: r, green: g, blue: b };
}

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}

async function getSheets() {
  const auth = await getAuth();
  return google.sheets({ version: 'v4', auth });
}

const CHUNK = 400;
async function batchUpdate(spreadsheetId, requests, label) {
  const sheets = await getSheets();
  for (let i = 0; i < requests.length; i += CHUNK) {
    const slice = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: slice } });
    if (requests.length > CHUNK) console.log(`  ${label} chunk ${Math.floor(i/CHUNK)+1}`);
  }
}

const VCHUNK = 200;
async function valuesBatchUpdate(spreadsheetId, data, label) {
  const sheets = await getSheets();
  for (let i = 0; i < data.length; i += VCHUNK) {
    const slice = data.slice(i, i + VCHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: slice },
    });
    if (data.length > VCHUNK) console.log(`  ${label} value chunk ${Math.floor(i/VCHUNK)+1}`);
  }
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

module.exports = { C, hex, getAuth, getSheets, batchUpdate, valuesBatchUpdate, gridRange };
