"""Finnhub client — primary data source (DESIGN.md section 3).

Covers what the Data Agent needs: quote, 52-week range, daily OHLCV
candles, analyst recommendation trends, and price targets. Every method
returns None on failure instead of raising, so a single bad ticker or a
rate-limit hit doesn't take down the whole weekly run.
"""

from __future__ import annotations

import time
from typing import Optional

from .. import config
from ..models import OHLCVBar
from .base import RateLimitedClient


class FinnhubClient(RateLimitedClient):
    def __init__(self, api_key: str = config.FINNHUB_API_KEY):
        super().__init__(min_seconds_between_calls=config.FINNHUB_MIN_SECONDS_BETWEEN_CALLS)
        self._api_key = api_key

    def _get(self, path: str, params: dict) -> Optional[dict | list]:
        return self.get_json(
            f"{config.FINNHUB_BASE_URL}{path}",
            {**params, "token": self._api_key},
        )

    def get_quote(self, symbol: str) -> Optional[dict]:
        """Current price, change, % change, day high/low/open, prev close."""
        data = self._get("/quote", {"symbol": symbol})
        if not data or data.get("c") in (None, 0):
            return None
        return {
            "price": data["c"],
            "change": data.get("d"),
            "percent_change": data.get("dp"),
            "day_high": data.get("h"),
            "day_low": data.get("l"),
            "open": data.get("o"),
            "prev_close": data.get("pc"),
        }

    def get_52_week_range(self, symbol: str) -> Optional[dict]:
        data = self._get("/stock/metric", {"symbol": symbol, "metric": "all"})
        if not data:
            return None
        metric = data.get("metric", {})
        high = metric.get("52WeekHigh")
        low = metric.get("52WeekLow")
        if high is None and low is None:
            return None
        return {"week_52_high": high, "week_52_low": low}

    def get_candles(
        self, symbol: str, lookback_days: int = config.OHLCV_LOOKBACK_DAYS
    ) -> Optional[list[OHLCVBar]]:
        """Daily OHLCV bars for the last `lookback_days` calendar days."""
        now = int(time.time())
        start = now - lookback_days * 86400
        data = self._get(
            "/stock/candle",
            {"symbol": symbol, "resolution": "D", "from": start, "to": now},
        )
        if not data or data.get("s") != "ok":
            return None

        bars = []
        for i, ts in enumerate(data["t"]):
            bars.append(
                OHLCVBar(
                    date=time.strftime("%Y-%m-%d", time.gmtime(ts)),
                    open=data["o"][i],
                    high=data["h"][i],
                    low=data["l"][i],
                    close=data["c"][i],
                    volume=data["v"][i],
                )
            )
        return bars

    def get_recommendation_trends(self, symbol: str) -> Optional[dict]:
        """Most recent analyst buy/hold/sell counts."""
        data = self._get("/stock/recommendation", {"symbol": symbol})
        if not data:
            return None
        latest = data[0]
        return {
            "strong_buy": latest.get("strongBuy", 0),
            "buy": latest.get("buy", 0),
            "hold": latest.get("hold", 0),
            "sell": latest.get("sell", 0),
            "strong_sell": latest.get("strongSell", 0),
        }

    def get_price_target(self, symbol: str) -> Optional[dict]:
        data = self._get("/stock/price-target", {"symbol": symbol})
        if not data:
            return None
        return {
            "price_target_high": data.get("targetHigh"),
            "price_target_low": data.get("targetLow"),
            "price_target_mean": data.get("targetMean"),
            "price_target_median": data.get("targetMedian"),
        }

    def get_news(self, symbol: str, from_date: str, to_date: str) -> Optional[list[dict]]:
        """Raw news feed for a symbol between two ISO dates (used by the News Agent)."""
        return self._get(
            "/company-news", {"symbol": symbol, "from": from_date, "to": to_date}
        )
