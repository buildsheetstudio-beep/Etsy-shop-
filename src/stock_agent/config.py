"""Central configuration: API keys, file paths, and indicator parameters.

Exact MA/RSI crossover thresholds are an explicit open item in DESIGN.md
section 10 ("Define exact MA/RSI thresholds that count as 'crossover' /
'overbought' / 'oversold'") — the Trend Agent owns that decision. This
module only holds the window sizes the Data Agent needs to compute the
raw indicator values.
"""

from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

ROOT_DIR = Path(__file__).resolve().parents[2]

FINNHUB_API_KEY = os.environ.get("FINNHUB_API_KEY", "")
TWELVEDATA_API_KEY = os.environ.get("TWELVEDATA_API_KEY", "")

FINNHUB_BASE_URL = "https://finnhub.io/api/v1"
TWELVEDATA_BASE_URL = "https://api.twelvedata.com"

TICKERS_FILE = ROOT_DIR / "config" / "tickers.json"
SNAPSHOT_FILE = ROOT_DIR / "data" / "snapshots" / "weekly_snapshot.json"
RUN_OUTPUT_DIR = ROOT_DIR / "data" / "run"

# DESIGN.md 6.2 asks for "~30 days OHLCV", but a 50-day MA needs 50+
# trading days of history to be defined at all. We pull more calendar
# days than the doc's headline number so the 50-day MA isn't silently
# null for every ticker; ~35 trading weeks of calendar days comfortably
# covers 50+ trading days including weekends/holidays.
OHLCV_LOOKBACK_DAYS = 100

MA_WINDOWS = (20, 50)
RSI_PERIOD = 14

# Finnhub free tier: 300 calls/day, 60 calls/min.
FINNHUB_MIN_SECONDS_BETWEEN_CALLS = 60 / 60

# Twelve Data free tier: 800 calls/day, 8 calls/sec.
TWELVEDATA_MIN_SECONDS_BETWEEN_CALLS = 1 / 8

# DESIGN.md 6.1: "Speculative (<$5)" / "Standard (~$5+)".
DISCOVERY_SPECULATIVE_PRICE_CEILING = 5.0

# How many web search results the Discovery Agent asks for per query.
DISCOVERY_RESULTS_PER_QUERY = 10

# DESIGN.md 6.3: "Pulls each ticker's news feed for the past 7 days".
NEWS_LOOKBACK_DAYS = 7
