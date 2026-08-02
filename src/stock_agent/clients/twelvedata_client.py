"""Twelve Data client — backup source (DESIGN.md section 3).

Used only when Finnhub can't supply OHLCV history or the Data Agent wants
a server-side cross-check for MA/RSI. Same fail-soft contract as
FinnhubClient: every method returns None rather than raising.
"""

from __future__ import annotations

from typing import Optional

from .. import config
from ..models import OHLCVBar
from .base import RateLimitedClient


class TwelveDataClient(RateLimitedClient):
    def __init__(self, api_key: str = config.TWELVEDATA_API_KEY):
        super().__init__(min_seconds_between_calls=config.TWELVEDATA_MIN_SECONDS_BETWEEN_CALLS)
        self._api_key = api_key

    def _get(self, path: str, params: dict) -> Optional[dict]:
        return self.get_json(
            f"{config.TWELVEDATA_BASE_URL}{path}",
            {**params, "apikey": self._api_key},
        )

    def get_time_series(
        self, symbol: str, outputsize: int = config.OHLCV_LOOKBACK_DAYS
    ) -> Optional[list[OHLCVBar]]:
        data = self._get(
            "/time_series",
            {"symbol": symbol, "interval": "1day", "outputsize": outputsize},
        )
        if not data or data.get("status") == "error" or "values" not in data:
            return None

        bars = [
            OHLCVBar(
                date=v["datetime"],
                open=float(v["open"]),
                high=float(v["high"]),
                low=float(v["low"]),
                close=float(v["close"]),
                volume=int(v["volume"]),
            )
            for v in data["values"]
        ]
        return list(reversed(bars))  # Twelve Data returns newest-first

    def get_sma(self, symbol: str, time_period: int) -> Optional[float]:
        data = self._get(
            "/sma",
            {"symbol": symbol, "interval": "1day", "time_period": time_period},
        )
        if not data or data.get("status") == "error" or not data.get("values"):
            return None
        return float(data["values"][0]["sma"])

    def get_rsi(self, symbol: str, time_period: int = config.RSI_PERIOD) -> Optional[float]:
        data = self._get(
            "/rsi",
            {"symbol": symbol, "interval": "1day", "time_period": time_period},
        )
        if not data or data.get("status") == "error" or not data.get("values"):
            return None
        return float(data["values"][0]["rsi"])
