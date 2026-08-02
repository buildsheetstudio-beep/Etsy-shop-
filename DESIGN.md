# Multi-Agent Stock Research System — Design Doc

**Purpose**: A self-sufficient, automated weekly research tool that monitors owned stocks and
a watchlist, surfaces favorable trends, and discovers new watchlist candidates — for
research and awareness only, not automated trading.

**Status**: Design locked. Build in progress (see `README.md` for current implementation status).

---

## 1. Scope

- Broad market scan with an overlay of the user's owned stocks and watchlist
- No automated buying/selling — output is informational only
- Ticker list (owned + watchlist) is manually maintained by the user for now; brokerage
  sync is a possible future upgrade

## 2. Cadence & Runtime Environment

- Runs fully automated, **weekly**, via a scheduled cron job
- Lives in a **Claude Code project** (not inside the chat interface) — needed because:
  - Cron/scheduling isn't available natively in chat
  - Technical indicator computation across 25–50+ tickers is compute-heavier than a
    chat workflow supports well
- Chat remains available afterward for Q&A about a given week's report

## 3. Data Sources

| Provider | Role | Free Tier |
|---|---|---|
| Finnhub | Primary — quotes, volume, news, sentiment, analyst estimates/targets | 300 calls/day, 60 calls/min |
| Twelve Data | Backup — server-side technical indicators (MA, RSI, etc.) if Finnhub's depth is insufficient | 800 calls/day, 8 calls/sec |
| Web search | Supplementary — screener-style "top movers / oversold / breakout" lists for Discovery Agent candidate generation | N/A |

> Note: free tiers are rate-limited and not guaranteed long-term stable — agents should fail
> gracefully (skip/retry) rather than assume 100% uptime.

## 4. Ticker List

- Manually maintained file: `tickers.json` (or CSV)
- Fields: `symbol`, `type` (`owned` / `watchlist`)
- User edits this file directly when buying/selling or adding to watchlist
- Read fresh by the Data Agent at the start of each weekly run

## 5. Signal Criteria

Flagged signals fall into three categories, all surfaced and labeled (not filtered to one type):

- **Technical**: moving average crossovers (e.g., 20/50-day), RSI moving into
  overbought/oversold territory, volume spikes
- **News/Catalyst**: earnings releases, M&A activity, guidance changes, leadership changes
- **Analyst**: rating changes, price target revisions

## 6. Agent Architecture

### 6.1 Discovery Agent

- Generates new watchlist *candidates* the user doesn't currently track
- Pulls shortlists from screener-style sources across three setup types:
  - Momentum/breakout (price breakout, volume surge)
  - Oversold bounce (low RSI, pulled back from highs)
  - Undervalued fundamentals (low P/E, analyst target above price)
- Covers full price range ($0.01–$1,000), split into two risk tiers:
  - Standard (~$5+)
  - Speculative (<$5) — explicitly flagged high-risk (liquidity/spread/manipulation caveats)
- Target output: **10+ ranked candidates/week**, each tagged with triggering signal type(s)
  and risk tier
- Candidates are passed through the Data Agent for real verification before appearing on
  the dashboard — not just a raw screener echo

### 6.2 Data Agent

- For every ticker (owned + watchlist + Discovery candidates): pulls current quote,
  volume, 52-week range, ~30 days OHLCV
- Computes technical indicators locally (20/50-day MA, RSI) from raw OHLCV
- Pulls analyst recommendation trends + price targets
- Outputs one structured record per ticker

### 6.3 News/Catalyst Agent

- Pulls each ticker's news feed for the past 7 days
- Flags known catalyst categories (earnings, M&A, guidance, leadership, major rating changes)
- Outputs a "notable events" list per ticker (empty if nothing happened — no
  manufactured content)

### 6.4 Trend Agent

- Reads last week's saved snapshot (JSON: ticker → price, MA, RSI, analyst target)
- Diffs this week vs. last week: % price change, MA crossovers, RSI zone changes, analyst
  target revisions
- This is the filtering layer — tickers with no meaningful change and no news get dropped
  rather than padding the report
- Saves this week's snapshot for next week's comparison

### 6.5 Synthesis Agent

- Combines Discovery, Data, News, and Trend outputs
- Builds the final dashboard with three sections:
  1. **Your Holdings** — hold / watch / reconsider framing
  2. **Your Watchlist** — entry-signal framing
  3. **New Suggestions** — Discovery Agent candidates, tagged by signal type and risk tier
     (standard vs. speculative)
- Ranks within each section by signal strength (e.g., MA crossover + analyst downgrade
  ranks above a minor RSI dip)
- Renders as an HTML dashboard artifact

## 7. Pipeline Sequence (weekly run)

```
Discovery Agent ─┐
                  ├─→ Data Agent ─→ News Agent ─→ Trend Agent ─→ Synthesis Agent ─→ Dashboard
   (candidates) ──┘   (verifies)
```

## 8. Memory / Storage

- Single JSON (or lightweight SQLite) snapshot file: `weekly_snapshot.json`
- Keyed by ticker → last week's price, MA, RSI, analyst target
- Overwritten each run after the Trend Agent's diff step

## 9. Output

- HTML dashboard, color-coded by signal direction
- Three sections as described in 6.5
- Saved locally each week; delivery mechanism (file only vs. email) TBD at build time

## 10. Open Items for Build Phase

- Finalize exact Finnhub/Twelve Data endpoints per agent
- Decide dashboard delivery: local file vs. email vs. both
- Define exact MA/RSI thresholds that count as "crossover" / "overbought" / "oversold"
- Decide whether this becomes a formal Claude Skill once stable (ticker format,
  thresholds, and dashboard structure are strong candidates for this)
- Future: optional read-only brokerage sync to replace manual ticker list maintenance
