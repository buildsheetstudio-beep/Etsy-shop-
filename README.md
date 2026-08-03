# Multi-Agent Stock Research System

Automated weekly research tool that monitors owned stocks and a watchlist,
surfaces favorable trends, and discovers new watchlist candidates. Research
and awareness only — no automated trading. Full design: [`DESIGN.md`](./DESIGN.md).

## Status

All five agents (DESIGN.md section 6) are implemented and tested: Data,
Discovery, News/Catalyst, Trend, and Synthesis. `python -m stock_agent.main`
runs the real end-to-end pipeline — including live web search for
Discovery — and produces an actual HTML dashboard file.

Dashboard delivery (DESIGN.md section 10) is **file-only for now** —
the Synthesis Agent renders to `data/run/dashboard.html`; email delivery
was left for a later pass since it's additive to the existing file output,
not a redesign of it. It also resolves a DESIGN.md open item on its own
account: the doc calls for "hold / watch / reconsider" (Holdings) and
"entry-signal" (Watchlist) framing without defining how signals map to
those labels — see `config.py`'s `SYNTHESIS_*` constants for the
resolution (a numeric signal-strength score plus bullish/bearish/neutral
direction).

The Trend Agent resolves DESIGN.md section 10's threshold open item:
RSI oversold/overbought at 30/70, MA crossover as a sign flip of
(MA20 − MA50) with no magnitude buffer, volume spike at 2x the trailing
20-day average, price-target revision at ±2%, and a ±3% weekly price
move as the bar for "meaningful" in the report filter. See `config.py`
for the exact constants and rationale.

The Discovery Agent's web search backend (DESIGN.md section 10's other
open item — no vendor named, free tier "N/A") is resolved with **Tavily**
(`clients/tavily_client.py`), chosen for its agent-oriented API (clean
title/url/content results, no HTML scraping) and usable free tier (1,000
searches/month). `DiscoveryAgent()` and `pipeline.run_discovery_stage()`
default to it; both still accept an explicit `WebSearchProvider` if you
want to swap backends or inject a fake for testing.

## Project layout

```
DESIGN.md                     Full design doc
config/
  tickers.example.json        Template — copy to tickers.json and edit
src/stock_agent/
  config.py                   API keys, file paths, indicator windows, all resolved thresholds
  models.py                   Ticker / OHLCVBar / AnalystData / TickerRecord / DiscoveryCandidate / NotableEvent / TrendResult
  indicators.py                Local SMA + RSI + volume-spike computation
  storage.py                  Ticker list + JSON snapshot I/O
  clients/
    finnhub_client.py         Primary data source
    twelvedata_client.py      Backup OHLCV / indicators
    tavily_client.py          Discovery Agent's web search backend
    websearch_client.py       WebSearchProvider protocol Tavily (and any alternative) implements
  agents/
    data_agent.py             Implemented — DESIGN.md 6.2
    discovery_agent.py        Implemented — DESIGN.md 6.1
    news_agent.py              Implemented — DESIGN.md 6.3
    trend_agent.py              Implemented — DESIGN.md 6.4
    synthesis_agent.py          Implemented — DESIGN.md 6.5
  pipeline.py                 Wires all five agent stages together end to end
  main.py                     CLI entrypoint (runs the Data stage)
tests/                        pytest suite
data/
  run/                        Per-run output (*_agent_output.json, dashboard.html, gitignored)
  snapshots/                  weekly_snapshot.json for the Trend Agent's week-over-week diff (gitignored)
```

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements-dev.txt

cp .env.example .env            # fill in FINNHUB_API_KEY / TWELVEDATA_API_KEY / TAVILY_API_KEY
cp config/tickers.example.json config/tickers.json   # edit with your real tickers
```

`TAVILY_API_KEY` is only needed if you want Discovery's New Suggestions
section populated with real candidates; every other part of the pipeline
runs (and degrades gracefully) without it.

## Running the full pipeline

```bash
python -m stock_agent.main
```

Runs Discovery → Data → News → Trend → Synthesis end to end and writes
`data/run/dashboard.html` — this is what the weekly cron job (DESIGN.md
section 2) should invoke. Every stage fails soft: a missing API key or a
dead endpoint shows up as an empty section or a per-ticker error, never
a crash.

## Running the Data Agent

```python
from stock_agent.pipeline import run_data_stage
records = run_data_stage()
```

Reads `config/tickers.json`, fetches quote / 52-week range / OHLCV / analyst
data per ticker from Finnhub (falling back to Twelve Data for OHLCV if
needed), computes 20/50-day MA and RSI(14) locally, and writes one
`TickerRecord` per ticker to `data/run/data_agent_output.json`. Failures on
individual data points are recorded on that ticker's `errors` list rather
than aborting the run.

## Running the Discovery Agent

```python
from stock_agent.pipeline import run_discovery_stage
candidates = run_discovery_stage()  # uses Tavily by default
```

Runs three setup-type screeners (momentum/breakout, oversold bounce,
undervalued fundamentals) as web search queries, extracts ticker symbols
and a rough price from the result snippets, excludes anything already in
`config/tickers.json`, and ranks candidates by how many setup types
flagged them. Output is provisional (`DiscoveryCandidate`, not
`TickerRecord`) — DESIGN.md 6.1 requires these to go through the Data
Agent for real verification before they'd appear on a dashboard;
`run_synthesis_stage()` handles that hand-off automatically.

## Running the News Agent

```python
from stock_agent.pipeline import run_news_stage
events_by_symbol = run_news_stage()  # defaults to config/tickers.json
```

Pulls each ticker's 7-day news feed from Finnhub, classifies headlines and
summaries into catalyst categories — earnings, M&A, guidance changes,
leadership changes, analyst rating changes — via keyword matching, and
returns a `NotableEvent` list per symbol. A ticker with no matching news
gets an empty list rather than manufactured content, per DESIGN.md 6.3.
Output is written to `data/run/news_agent_output.json`.

## Running the Trend Agent

```python
from stock_agent.pipeline import run_trend_stage
results = run_trend_stage()  # chains Data -> News -> Trend by default
```

Loads last week's snapshot (`data/snapshots/weekly_snapshot.json`), diffs
this week's `TickerRecord`s against it (% price change, MA crossover,
RSI zone change, rating change, price-target revision), carries through
the Data Agent's volume-spike flag and the News Agent's notable events,
and drops any ticker with no signal and no news — DESIGN.md 6.4's
filtering layer. A ticker with no prior snapshot (first run, or a brand
new Discovery candidate) is always kept, tagged `is_new=True`, since
there's nothing to diff against yet. Overwrites the snapshot for next
week's comparison and writes results to
`data/run/trend_agent_output.json`.

Note: DESIGN.md section 8 describes the snapshot as price/MA/RSI/analyst
target; this implementation also stores each ticker's recommendation
counts so rating changes are diffable week over week rather than only
reflecting current sentiment — documented in `trend_agent.py`.

## Running the Synthesis Agent

```python
from stock_agent.pipeline import run_synthesis_stage
html = run_synthesis_stage()  # chains Discovery -> Data -> News -> Trend -> Synthesis
```

Combines Trend Agent results and verified Discovery candidates into a
single self-contained HTML dashboard with three ranked, color-coded
sections — Your Holdings, Your Watchlist, New Suggestions — matching
DESIGN.md 6.5/9. Pass `include_suggestions=False` to skip Discovery
entirely (no web search calls, New Suggestions renders empty), or
`search_provider=` to use a different backend than Tavily:

```python
html = run_synthesis_stage(include_suggestions=False)          # Holdings/Watchlist only
html = run_synthesis_stage(search_provider=my_search_provider)  # different backend
```

Written to `data/run/dashboard.html`. Each row is ranked within its
section by a signal-strength score and color-coded bullish/green,
bearish/red, or neutral/gray; notable news headlines are linked inline.

## Weekly automation (GitHub Actions)

`.github/workflows/weekly-research.yml` runs the full pipeline every
Monday (DESIGN.md section 2's "scheduled cron job") and uploads
`dashboard.html` as a downloadable workflow artifact. It also supports
`workflow_dispatch`, so you can trigger a run manually from the Actions
tab any time.

**This repo must be private before you use this.** Actions artifacts,
logs, and caches on a *public* repo are visible to anyone who can view
the repo — that would mean your real ticker holdings and dashboard
(prices, portfolio composition) are effectively public. Make the repo
private first: repo → Settings → Danger Zone → Change visibility.

Required repo secrets (Settings → Secrets and variables → Actions):

| Secret | Purpose |
|---|---|
| `FINNHUB_API_KEY` | Data Agent's primary source |
| `TWELVEDATA_API_KEY` | Data Agent's OHLCV fallback |
| `TAVILY_API_KEY` | Discovery Agent's web search backend |
| `TICKERS_JSON` | The full contents of your real `config/tickers.json` — it's gitignored (personal holdings), so the workflow can't read it from the repo checkout. Without this secret, the workflow falls back to `config/tickers.example.json` and warns in the run log. |

Every stage fails soft (DESIGN.md section 3), so a missing secret shows
up as an empty dashboard section or a per-ticker error in that step's
log, not a failed run.

The weekly snapshot (`data/snapshots/weekly_snapshot.json`, needed for
the Trend Agent's week-over-week diff) persists across runs via
`actions/cache` rather than being committed — a cache miss just makes
that week's tickers all read as "new" rather than breaking anything.

The cron is set to 11:00 UTC Mondays (7am ET at setup time). GitHub
Actions cron is always UTC and doesn't shift for daylight saving —
adjust the hour by one twice a year if you want it pinned to a specific
local time, or leave it and accept an hour of drift each side of DST.

## Tests

```bash
pytest
```

## Next steps

All five agents are built and wired end to end with a real search
backend, and the weekly run is automated via GitHub Actions. What's left
is on your side: make the repo private, add the four secrets above, and
decide if local-artifact delivery is enough or you want email too.
