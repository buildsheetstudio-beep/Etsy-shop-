# Multi-Agent Stock Research System

Automated weekly research tool that monitors owned stocks and a watchlist,
surfaces favorable trends, and discovers new watchlist candidates. Research
and awareness only — no automated trading. Full design: [`DESIGN.md`](./DESIGN.md).

## Status

All five agents (DESIGN.md section 6) are implemented and tested: Data,
Discovery, News/Catalyst, Trend, and Synthesis. The end-to-end pipeline
runs and produces a real HTML dashboard file.

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

## Running the Synthesis Agent (full pipeline)

```python
from stock_agent.pipeline import run_synthesis_stage
html = run_synthesis_stage()  # chains Data -> News -> Trend -> Synthesis
```

Combines Trend Agent results and (if a `WebSearchProvider` is passed)
verified Discovery candidates into a single self-contained HTML
dashboard with three ranked, color-coded sections — Your Holdings, Your
Watchlist, New Suggestions — matching DESIGN.md 6.5/9. Without a search
provider, New Suggestions renders empty (candidates need Discovery to
run first); Holdings and Watchlist work either way:

```python
from stock_agent.pipeline import run_synthesis_stage
html = run_synthesis_stage(search_provider=my_search_provider)  # includes New Suggestions
```

Written to `data/run/dashboard.html`. Each row is ranked within its
section by a signal-strength score and color-coded bullish/green,
bearish/red, or neutral/gray; notable news headlines are linked inline.

## Tests

```bash
pytest
```

## Next steps

All five agents are built. What's left is operational, not architectural:
wire a real `WebSearchProvider` (DESIGN.md section 10) so Discovery can
run for real, decide on email delivery for the dashboard if wanted, and
set up the actual weekly cron job (DESIGN.md section 2) to call
`run_synthesis_stage`.
