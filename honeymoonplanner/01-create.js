'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const { getAuth } = (() => {
  const { google: g2 } = require('googleapis');
  const SECRET = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const { client_id, client_secret, redirect_uris } = SECRET.installed || SECRET.web;
  const oAuth2 = new g2.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  oAuth2.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return { getAuth: () => oAuth2 };
})();

const TABS = [
  { title: '📋 Reference Data',                    color: '#9E9E9E' },
  { title: '🌅 Honeymoon Overview',                color: '#E0A47F' },
  { title: '🗓️ Multi-City Day-by-Day Itinerary',   color: '#3E5C58' },
  { title: '💰 Budget & Expense Tracker',           color: '#D9A548' },
  { title: '🔐 Travel Document & Confirmation Vault', color: '#B3453A' },
  { title: '✨ Activity & Excursion Wishlist',       color: '#7A9B7A' },
  { title: '✅ Research & Reservation Checklist',   color: '#2C4440' },
  { title: '🧳 Packing List',                       color: '#3A2E28' },
];

(async () => {
  const s = google.sheets({ version: 'v4', auth: getAuth() });
  const res = await s.spreadsheets.create({
    requestBody: {
      properties: { title: '🌅 Ultimate Honeymoon Planner' },
      sheets: TABS.map((t, i) => ({
        properties: {
          sheetId: i,
          title: t.title,
          index: i,
          tabColorStyle: { rgbColor: { red: parseInt(t.color.slice(1,3),16)/255, green: parseInt(t.color.slice(3,5),16)/255, blue: parseInt(t.color.slice(5,7),16)/255 } },
        },
      })),
    },
  });
  const id = res.data.spreadsheetId;
  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });
  fs.writeFileSync(path.join(__dirname, 'spreadsheet.json'), JSON.stringify({ id, sheetMap }, null, 2));
  console.log('Created spreadsheet:', id);
  console.log('URL: https://docs.google.com/spreadsheets/d/' + id);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
