"""Data Agent — DESIGN.md section 6.2.

For every ticker (owned + watchlist + Discovery candidates): pulls current
quote, volume, 52-week range, and OHLCV history; computes 20/50-day MA and
RSI locally from that OHLCV; pulls analyst recommendation trends and price
targets; and outputs one structured TickerRecord per ticker.

Finnhub is the primary source for everything. Twelve Data is only
consulted as an OHLCV backup when Finnhub's candles are unavailable
(DESIGN.md section 3) — indicators are still computed locally from
whichever OHLCV series was obtained, keeping one code path for both.
Per-ticker failures are recorded on TickerRecord.errors rather than
raising, so one bad symbol never aborts the run.
"""

from __future__ import annotations

import logging

from .. import config
from ..clients.finnhub_client import FinnhubClient
from ..clients.twelvedata_client import TwelveDataClient
from ..indicators import rsi, simple_moving_average
from ..models import AnalystData, Ticker, TickerRecord

logger = logging.getLogger(__name__)


class DataAgent:
    def __init__(
        self,
        finnhub_client: FinnhubClient | None = None,
        twelvedata_client: TwelveDataClient | None = None,
    ):
        self._finnhub = finnhub_client or FinnhubClient()
        self._twelvedata = twelvedata_client or TwelveDataClient()

    def run(self, tickers: list[Ticker]) -> list[TickerRecord]:
        return [self.process_ticker(ticker) for ticker in tickers]

    def process_ticker(self, ticker: Ticker) -> TickerRecord:
        record = TickerRecord(symbol=ticker.symbol, type=ticker.type)

        quote = self._finnhub.get_quote(ticker.symbol)
        if quote:
            record.price = quote["price"]
            record.change = quote["change"]
            record.percent_change = quote["percent_change"]
        else:
            record.errors.append("quote unavailable")

        week_52 = self._finnhub.get_52_week_range(ticker.symbol)
        if week_52:
            record.week_52_high = week_52["week_52_high"]
            record.week_52_low = week_52["week_52_low"]
        else:
            record.errors.append("52-week range unavailable")

        ohlcv = self._finnhub.get_candles(ticker.symbol)
        if not ohlcv:
            record.errors.append("Finnhub OHLCV unavailable, falling back to Twelve Data")
            ohlcv = self._twelvedata.get_time_series(ticker.symbol)

        if ohlcv:
            record.ohlcv = ohlcv
            if ohlcv[-1].volume is not None:
                record.volume = ohlcv[-1].volume
            closes = [bar.close for bar in ohlcv]
            record.ma_20 = simple_moving_average(closes, 20)
            record.ma_50 = simple_moving_average(closes, 50)
            record.rsi_14 = rsi(closes, config.RSI_PERIOD)
        else:
            record.errors.append("OHLCV unavailable from both Finnhub and Twelve Data")

        analyst = AnalystData()
        recs = self._finnhub.get_recommendation_trends(ticker.symbol)
        if recs:
            analyst.strong_buy = recs["strong_buy"]
            analyst.buy = recs["buy"]
            analyst.hold = recs["hold"]
            analyst.sell = recs["sell"]
            analyst.strong_sell = recs["strong_sell"]
        else:
            record.errors.append("analyst recommendation trends unavailable")

        targets = self._finnhub.get_price_target(ticker.symbol)
        if targets:
            analyst.price_target_high = targets["price_target_high"]
            analyst.price_target_low = targets["price_target_low"]
            analyst.price_target_mean = targets["price_target_mean"]
            analyst.price_target_median = targets["price_target_median"]
        else:
            record.errors.append("price target unavailable")

        record.analyst = analyst

        if record.errors:
            logger.warning("%s: %s", ticker.symbol, "; ".join(record.errors))

        return record
