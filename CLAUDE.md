# Etsy Shop — Project Instructions

## Quality Gate: Formula Validation Before Delivery

**Every spreadsheet build must pass a formula error check before being considered finished.**

After all scripts have run, scan every tab for `#VALUE`, `#REF`, and `#ERROR` cells and fix them before reporting the build as complete. Do not deliver a spreadsheet that still contains these errors.

### Known error patterns to watch for in Google Sheets scripts

1. **COUNTIFS with computed criteria ranges** — `COUNTIFS(YEAR(range), value, ...)` and `COUNTIFS(MONTH(range), value, ...)` always produce `#VALUE`. Replace with `SUMPRODUCT` boolean multiplication:
   ```js
   // Wrong
   `=COUNTIFS(YEAR('Sheet'!$B$6:$B$500), 2026, 'Sheet'!$P$6:$P$500, "Received")`
   // Correct
   `=SUMPRODUCT((YEAR('Sheet'!$B$6:$B$500)=2026)*('Sheet'!$P$6:$P$500="Received"))`
   ```

2. **Formula string embedding** — variables that start with `=` must have the leading `=` stripped before being embedded inside another formula string:
   ```js
   // Wrong — produces =IFERROR(=SUMPRODUCT(...)/=SUMPRODUCT(...),0)
   const grossRev = `=SUMPRODUCT(...)`;
   `=IFERROR(${grossRev}/${grossRev},0)`

   // Correct — define bare expressions without leading =
   const grossRevExpr = `SUMPRODUCT(...)`;
   `=IFERROR(${grossRevExpr}/${grossRevExpr},0)`
   ```

3. **`COUNTPRODUCT` is not a function** — use `SUMPRODUCT` instead.

4. **Chart API field names** — use `basicChart` with `chartType: 'COMBO'` or `chartType: 'BAR'`; never `comboChart` or `barChart`. Series type for COMBO must be `'COLUMN'`, not `'COLUMNS'`.

5. **Conditional format condition types** — use `NUMBER_GREATER_THAN_EQ`, not `NUMBER_GREATER_EQ`.

6. **Freeze column + merged title conflict** — do not set `frozenColumnCount` when a merged title row spans all columns across that freeze boundary.
