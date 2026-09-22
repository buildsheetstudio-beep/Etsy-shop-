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

const code = process.argv[2];
if (!code) { console.error('Usage: node save-tokens.js <auth-code>'); process.exit(1); }

auth.getToken(code).then(({ tokens }) => {
  fs.writeFileSync(path.resolve(__dirname, 'tokens.json'), JSON.stringify(tokens, null, 2));
  console.log('✓ tokens.json saved');
  console.log('  access_token:', tokens.access_token ? tokens.access_token.slice(0,20)+'...' : 'none');
  console.log('  refresh_token:', tokens.refresh_token ? 'present' : 'missing');
}).catch(e => { console.error('Token exchange failed:', e.message); process.exit(1); });
