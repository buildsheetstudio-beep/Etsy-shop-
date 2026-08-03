"""Central configuration: API keys, file paths, and indicator parameters.

DESIGN.md section 10 leaves exact MA/RSI/volume thresholds as an open
item. Resolved values below:
- RSI oversold/overbought: <=30 / >=70 (standard Wilder convention)
- MA crossover: sign flip of (MA20 - MA50) vs. last week, no magnitude
  buffer — a crossover is a crossover regardless of how close the values
  are. If choppy back-to-back cross/re-cross noise shows up in real
  output, revisit with a buffer then rather than guessing one now.
- Volume spike: most recent day's volume >= 2x the trailing 20-day
  average (excluding today)
- Price target revision: |change| >= 2% of last week's mean target
- "Meaningful" price move for the Trend Agent's report filter: >= 3%
  week-over-week
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
DASHBOARD_FILE = RUN_OUTPUT_DIR / "dashboard.html"

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

# DESIGN.md 10 open item — see module docstring for rationale.
RSI_OVERSOLD = 30.0
RSI_OVERBOUGHT = 70.0
VOLUME_SPIKE_WINDOW = 20
VOLUME_SPIKE_MULTIPLIER = 2.0
PRICE_TARGET_REVISION_PCT = 0.02
MEANINGFUL_PRICE_CHANGE_PCT = 0.03

# DESIGN.md 6.5 asks for "hold / watch / reconsider" (Holdings) and
# "entry-signal" (Watchlist) framing but doesn't define the mapping from
# signals to those labels. Resolved here the same way as the Trend Agent
# thresholds: a numeric "strength" score (see synthesis_agent.py) plus a
# bullish/bearish/neutral direction. Holdings: bearish direction at or
# above RECONSIDER strength -> "reconsider"; bearish (any strength) or
# strength at/above WATCH -> "watch"; otherwise "hold". Watchlist:
# bullish direction at or above ENTRY strength -> "entry signal"; bullish
# at/above POSSIBLE_ENTRY -> "possible entry"; otherwise "wait".
SYNTHESIS_RECONSIDER_STRENGTH = 4.0
SYNTHESIS_WATCH_STRENGTH = 2.0
SYNTHESIS_ENTRY_STRENGTH = 4.0
SYNTHESIS_POSSIBLE_ENTRY_STRENGTH = 2.0
