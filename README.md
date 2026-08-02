# Multi-Agent Stock Research System

Automated weekly research tool that monitors owned stocks and a watchlist,
surfaces favorable trends, and discovers new watchlist candidates. Research
and awareness only — no automated trading. Full design: [`DESIGN.md`](./DESIGN.md).

## Status

Scaffolding in progress. The **Data Agent** (design section 6.2) is
implemented and tested. Discovery, News, Trend, and Synthesis agents are
stubbed (`NotImplementedError`) pending the open items in DESIGN.md section 10.

## Project layout

```
DESIGN.md                     Full design doc
config/
  tickers.example.json        Template — copy to tickers.json and edit
src/stock_agent/
  config.py                   API keys, file paths, indicator windows
  models.py                   Ticker / OHLCVBar / AnalystData / TickerRecord
  indicators.py                Local SMA + RSI computation
  storage.py                  Ticker list + JSON snapshot I/O
  clients/
    finnhub_client.py         Primary data source
    twelvedata_client.py      Backup OHLCV / indicators
  agents/
    data_agent.py             Implemented — DESIGN.md 6.2
    discovery_agent.py        Stub — DESIGN.md 6.1
    news_agent.py              Stub — DESIGN.md 6.3
    trend_agent.py              Stub — DESIGN.md 6.4
    synthesis_agent.py          Stub — DESIGN.md 6.5
  pipeline.py                 Wires agent stages together (Data stage only, so far)
  main.py                     CLI entrypoint
tests/                        pytest suite
data/
  run/                        Per-run output (data_agent_output.json, gitignored)
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

## Tests

```bash
pytest
```

## Next steps

See DESIGN.md section 10 for the open build-phase decisions (exact
crossover/overbought/oversold thresholds, dashboard delivery mechanism,
etc.) — those need answers before the Trend and Synthesis agents can be
built. Discovery and News agents can be built independently of those
decisions.
