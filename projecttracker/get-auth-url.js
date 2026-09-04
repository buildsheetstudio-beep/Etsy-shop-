'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const KEY = JSON.parse(fs.readFileSync(path.resolve(__dirname, 'client_secret.json')));
const auth = new google.auth.OAuth2(
  KEY.installed.client_id,
  KEY.installed.client_secret,
  'urn:ietf:wg:oauth:2.0:oob'
);

const url = auth.generateAuthUrl({
  access_type: 'offline',
  scope: ['https://www.googleapis.com/auth/spreadsheets'],
  prompt: 'consent',
});

console.log(url);
