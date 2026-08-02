# Multi-Agent Stock Research System

Automated weekly research tool that monitors owned stocks and a watchlist,
surfaces favorable trends, and discovers new watchlist candidates. Research
and awareness only — no automated trading. Full design: [`DESIGN.md`](./DESIGN.md).

## Status

Scaffolding in progress. The **Data Agent** (section 6.2), **Discovery
Agent** (section 6.1), and **News/Catalyst Agent** (section 6.3) are
implemented and tested. Trend and Synthesis agents are stubbed
(`NotImplementedError`) pending the open items in DESIGN.md section 10.

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
  config.py                   API keys, file paths, indicator windows, risk-tier threshold
  models.py                   Ticker / OHLCVBar / AnalystData / TickerRecord / DiscoveryCandidate
  indicators.py                Local SMA + RSI computation
  storage.py                  Ticker list + JSON snapshot I/O
  clients/
    finnhub_client.py         Primary data source
    twelvedata_client.py      Backup OHLCV / indicators
    websearch_client.py       WebSearchProvider protocol — no concrete backend wired yet
  agents/
    data_agent.py             Implemented — DESIGN.md 6.2
    discovery_agent.py        Implemented — DESIGN.md 6.1
    news_agent.py              Implemented — DESIGN.md 6.3
    trend_agent.py              Stub — DESIGN.md 6.4
    synthesis_agent.py          Stub — DESIGN.md 6.5
  pipeline.py                 Wires agent stages together (Data + Discovery + News stages, so far)
  main.py                     CLI entrypoint (runs the Data stage)
tests/                        pytest suite
data/
  run/                        Per-run output (data_agent_output.json, discovery_agent_output.json, news_agent_output.json, gitignored)
  snapshots/                  weekly_snapshot.json for the Trend Agent (gitignored)
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

## Tests

```bash
pytest
```

## Next steps

See DESIGN.md section 10 for the open build-phase decisions (exact
crossover/overbought/oversold thresholds, dashboard delivery mechanism,
web search backend, etc.) — those need answers before the Trend and
Synthesis agents can be built.
