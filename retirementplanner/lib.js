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

// Deep Aubergine Navy / Muted Celadon / Soft Limestone theme
const C = {
  // Core UI
  text:        '#2B2D3E',   // near-black text on light bg
  secText:     '#6B6E84',   // secondary / muted labels
  bg:          '#F4F2EC',   // Soft Limestone — page background
  panel:       '#FFFFFF',   // card / cell background
  border:      '#D5D2C8',   // subtle border
  altRow:      '#EAE8E2',   // alternate row tint
  input:       '#FAFAF7',   // user-input cell
  formula:     '#EFF4F2',   // formula cell tint (celadon-tinted)
  white:       '#FFFFFF',

  // Accent
  primary:     '#3F435D',   // Deep Aubergine Navy — headers, KPI bands
  primaryText: '#FFFFFF',
  secondary:   '#8FB6A0',   // Muted Celadon — highlights, chart bars
  secondaryDk: '#6A967E',   // darker celadon for borders/icons

  // Status
  success:     '#8FB6A0',   // same as secondary — positive / on-track
  warning:     '#D4A84B',   // amber — approaching / caution
  attention:   '#B35C54',   // muted red — off-track / shortfall
  info:        '#7B9FC2',   // muted blue — informational

  // Scenario palette (SCN-01 through SCN-06)
  SCN01:       '#7B9FC2',   // Baseline — muted blue
  SCN02:       '#8FB6A0',   // Earlier Retirement — celadon green
  SCN03:       '#A89BC2',   // Later Retirement — soft lavender
  SCN04:       '#D4A84B',   // Higher Inflation — amber
  SCN05:       '#B35C54',   // Lower Returns — muted red
  SCN06:       '#7EB8B5',   // Reduced Spending — teal

  // Account type colors
  preTax:      '#C9D8EC',   // Pre-Tax (401k, IRA) — soft blue
  roth:        '#CDE5D5',   // Roth — soft green
  taxable:     '#E9D7B8',   // Taxable brokerage — soft tan
  taxDeferred: '#DED2EC',   // Tax-Deferred annuity — soft lavender

  // Section header fills
  hdrA:        '#3F435D',   // primary dark
  hdrB:        '#525770',   // primary mid
  hdrC:        '#6B6E84',   // primary light
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
