# Multi-Agent Stock Research System

Automated weekly research tool that monitors owned stocks and a watchlist,
surfaces favorable trends, and discovers new watchlist candidates. Research
and awareness only — no automated trading. Full design: [`DESIGN.md`](./DESIGN.md).

## Status

Scaffolding in progress. The **Data Agent** (section 6.2), **Discovery
Agent** (section 6.1), **News/Catalyst Agent** (section 6.3), and **Trend
Agent** (section 6.4) are implemented and tested. Only the Synthesis
Agent remains stubbed (`NotImplementedError`) — it needs a decision on
dashboard delivery (DESIGN.md section 10) before it can be built.

The Trend Agent resolves DESIGN.md section 10's threshold open item:
RSI oversold/overbought at 30/70, MA crossover as a sign flip of
(MA20 − MA50) with no magnitude buffer, volume spike at 2x the trailing
20-day average, price-target revision at ±2%, and a ±3% weekly price
move as the bar for "meaningful" in the report filter. See `config.py`
for the exact constants and rationale.

The Discovery Agent depends on a `WebSearchProvider` interface
(`clients/websearch_client.py`) rather than a hardcoded vendor, since
DESIGN.md doesn't name a web search backend (free tier: "N/A" — itself an
open item). No concrete provider ships yet; `pipeline.run_discovery_stage`
takes one as a required argument, and tests use a fake. Wiring a real
backend (e.g. Claude Code's own web search, a search API) is the next
step before this stage can run for real.

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
    websearch_client.py       WebSearchProvider protocol — no concrete backend wired yet
  agents/
    data_agent.py             Implemented — DESIGN.md 6.2
    discovery_agent.py        Implemented — DESIGN.md 6.1
    news_agent.py              Implemented — DESIGN.md 6.3
    trend_agent.py              Implemented — DESIGN.md 6.4
    synthesis_agent.py          Stub — DESIGN.md 6.5
  pipeline.py                 Wires agent stages together (Data + Discovery + News + Trend, so far)
  main.py                     CLI entrypoint (runs the Data stage)
tests/                        pytest suite
data/
  run/                        Per-run output (*_agent_output.json, gitignored)
  snapshots/                  weekly_snapshot.json for the Trend Agent's week-over-week diff (gitignored)
```

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements-dev.txt

cp .env.example .env            # fill in FINNHUB_API_KEY / TWELVEDATA_API_KEY
cp config/tickers.example.json config/tickers.json   # edit with your real tickers
```

## Running the Data Agent

```bash
python -m stock_agent.main
```

Reads `config/tickers.json`, fetches quote / 52-week range / OHLCV / analyst
data per ticker from Finnhub (falling back to Twelve Data for OHLCV if
needed), computes 20/50-day MA and RSI(14) locally, and writes one
`TickerRecord` per ticker to `data/run/data_agent_output.json`. Failures on
individual data points are recorded on that ticker's `errors` list rather
than aborting the run.

## Running the Discovery Agent

Not runnable from the CLI yet — it needs a `WebSearchProvider` (see
Status above). Once one exists:

```python
from stock_agent.pipeline import run_discovery_stage
candidates = run_discovery_stage(my_search_provider)
```

Runs three setup-type screeners (momentum/breakout, oversold bounce,
undervalued fundamentals) as web search queries, extracts ticker symbols
and a rough price from the result snippets, excludes anything already in
`config/tickers.json`, and ranks candidates by how many setup types
flagged them. Output is provisional (`DiscoveryCandidate`, not
`TickerRecord`) — DESIGN.md 6.1 requires these to go through the Data
Agent for real verification before they'd appear on a dashboard; that
hand-off isn't wired up yet.

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

## Tests

```bash
pytest
```

## Next steps

Only the **Synthesis Agent** (DESIGN.md 6.5) remains. It needs a
decision on dashboard delivery (local file vs. email vs. both — DESIGN.md
section 10) before the output format can be finalized, though the HTML
dashboard itself can likely be built file-only first and delivery added
later without much rework.
