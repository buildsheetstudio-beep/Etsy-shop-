'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const LOG_SID   = sheetMap['Daily Workout Log'];
const SETUP_SID = sheetMap['Workout Setup'];
const REF = 'Reference Data';

// ONE_OF_RANGE format: ='Reference Data'!A2:A10
// (= prefix required, quoted sheet name for spaces, NO $ signs in the range address)
const oneOfRange = (range, src) => ({
  setDataValidation: {
    range,
    rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='${REF}'!${src}` }] },
      strict: false,
      showCustomUi: true,
    },
  },
});

const boolCheck = (range) => ({
  setDataValidation: {
    range,
    rule: { condition: { type: 'BOOLEAN' }, strict: true },
  },
});

(async () => {
  const fmt = [];

  // ── Daily Workout Log — data rows ri=7..4999 (rows 8–5000) ─────────────────
  // Headers (row 7): Log ID | Session ID | Date | Year | Month | Day |
  //   Status(G=6) | Rest Day?(H=7) | Workout Type(I=8) | Workout Name(J=9) |
  //   Exercise ID(K=10) | ... | Set#(P=15) | Weight(Q=16) | Reps(R=17) |
  //   Duration(S=18) | Distance(T=19) | Speed(U=20) | Rest(V=21) |
  //   RPE(W=22) | Intensity(X=23) | Location(Y=24) | Planned?(Z=25) |
  //   Completed?(AA=26) | Est.Volume(AB=27) | PR Value(AC=28) | Notes(AD=29)

  // Col G (6): Status — Reference Data E2:E8 (7 values)
  fmt.push(oneOfRange(gridRange(LOG_SID, 7, 5000, 6, 7), 'E2:E8'));

  // Col H (7): Rest Day? — checkbox
  fmt.push(boolCheck(gridRange(LOG_SID, 7, 5000, 7, 8)));

  // Col I (8): Workout Type — Reference Data A2:A11 (10 types)
  fmt.push(oneOfRange(gridRange(LOG_SID, 7, 5000, 8, 9), 'A2:A11'));

  // Col W (22): RPE 1-10 — Reference Data G2:G11
  fmt.push(oneOfRange(gridRange(LOG_SID, 7, 5000, 22, 23), 'G2:G11'));

  // Col X (23): Intensity — Reference Data F2:F6
  fmt.push(oneOfRange(gridRange(LOG_SID, 7, 5000, 23, 24), 'F2:F6'));

  // Col Y (24): Location — Reference Data H2:H7
  fmt.push(oneOfRange(gridRange(LOG_SID, 7, 5000, 24, 25), 'H2:H7'));

  // Col Z (25): Planned? — checkbox
  fmt.push(boolCheck(gridRange(LOG_SID, 7, 5000, 25, 26)));

  // Col AA (26): Completed? — checkbox
  fmt.push(boolCheck(gridRange(LOG_SID, 7, 5000, 26, 27)));

  // ── Workout Setup — exercise rows ri=16..98 (rows 17–99, 83 exercises) ──────
  // Exercise table headers (row 16): Exercise ID | Name | Category(C=2) |
  //   Primary Muscle(D=3) | Secondary Muscle(E=4) | Tracking Mode(F=5) |
  //   Equipment(G=6) | Default Sets(H=7) | Default Reps(I=8) |
  //   Default Weight(J=9) | Default Duration(K=10) | Default Distance(L=11) |
  //   Active(M=12) | Notes(N=13)

  // Col C (2): Category — Reference Data C2:C8
  fmt.push(oneOfRange(gridRange(SETUP_SID, 16, 99, 2, 3), 'C2:C8'));

  // Col D (3): Primary Muscle — Reference Data B2:B21
  fmt.push(oneOfRange(gridRange(SETUP_SID, 16, 99, 3, 4), 'B2:B21'));

  // Col E (4): Secondary Muscle — Reference Data B2:B21
  fmt.push(oneOfRange(gridRange(SETUP_SID, 16, 99, 4, 5), 'B2:B21'));

  // Col F (5): Tracking Mode — Reference Data D2:D9
  fmt.push(oneOfRange(gridRange(SETUP_SID, 16, 99, 5, 6), 'D2:D9'));

  // Col G (6): Equipment — Reference Data I2:I15
  fmt.push(oneOfRange(gridRange(SETUP_SID, 16, 99, 6, 7), 'I2:I15'));

  // Col M (12): Active — checkbox
  fmt.push(boolCheck(gridRange(SETUP_SID, 16, 99, 12, 13)));

  await batchUpdate(id, fmt, '12-validation');
  console.log('✅  Data Validation done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
