from stock_agent.agents.data_agent import DataAgent
from stock_agent.models import OHLCVBar, Ticker


class FakeFinnhubClient:
    """Stands in for FinnhubClient without making network calls."""

    def __init__(self, quote=None, week_52=None, candles=None, recs=None, targets=None):
        self._quote = quote
        self._week_52 = week_52
        self._candles = candles
        self._recs = recs
        self._targets = targets

    def get_quote(self, symbol):
        return self._quote

    def get_52_week_range(self, symbol):
        return self._week_52

    def get_candles(self, symbol, lookback_days=None):
        return self._candles

    def get_recommendation_trends(self, symbol):
        return self._recs

    def get_price_target(self, symbol):
        return self._targets


class FakeTwelveDataClient:
    def __init__(self, time_series=None):
        self._time_series = time_series

    def get_time_series(self, symbol, outputsize=None):
        return self._time_series


def make_bars(closes):
    return [
        OHLCVBar(date=f"2026-01-{i+1:02d}", open=c, high=c, low=c, close=c, volume=1000)
        for i, c in enumerate(closes)
    ]


def test_process_ticker_happy_path():
    closes = list(range(1, 61))  # 60 increasing closes -> enough for MA50 + RSI14
    finnhub = FakeFinnhubClient(
        quote={"price": 100.0, "change": 1.0, "percent_change": 1.0},
        week_52={"week_52_high": 120.0, "week_52_low": 80.0},
        candles=make_bars(closes),
        recs={"strong_buy": 1, "buy": 2, "hold": 3, "sell": 0, "strong_sell": 0},
        targets={
            "price_target_high": 130.0,
            "price_target_low": 90.0,
            "price_target_mean": 110.0,
            "price_target_median": 108.0,
        },
    )
    agent = DataAgent(finnhub_client=finnhub, twelvedata_client=FakeTwelveDataClient())

    record = agent.process_ticker(Ticker(symbol="AAPL", type="owned"))

    assert record.symbol == "AAPL"
    assert record.price == 100.0
    assert record.week_52_high == 120.0
    assert record.ma_20 is not None
    assert record.ma_50 is not None
    assert record.rsi_14 == 100.0  # strictly increasing closes
    assert record.analyst.buy == 2
    assert record.analyst.price_target_mean == 110.0
    assert record.errors == []


def test_process_ticker_falls_back_to_twelvedata_when_finnhub_candles_missing():
    closes = list(range(1, 61))
    finnhub = FakeFinnhubClient(
        quote={"price": 50.0, "change": 0.0, "percent_change": 0.0},
        week_52={"week_52_high": 60.0, "week_52_low": 40.0},
        candles=None,
        recs={"strong_buy": 0, "buy": 0, "hold": 0, "sell": 0, "strong_sell": 0},
        targets=None,
    )
    twelvedata = FakeTwelveDataClient(time_series=make_bars(closes))
    agent = DataAgent(finnhub_client=finnhub, twelvedata_client=twelvedata)

    record = agent.process_ticker(Ticker(symbol="MSFT", type="watchlist"))

    assert record.ma_50 is not None
    assert any("Finnhub OHLCV unavailable" in e for e in record.errors)


def test_process_ticker_never_raises_when_everything_fails():
    finnhub = FakeFinnhubClient(quote=None, week_52=None, candles=None, recs=None, targets=None)
    twelvedata = FakeTwelveDataClient(time_series=None)
    agent = DataAgent(finnhub_client=finnhub, twelvedata_client=twelvedata)

    record = agent.process_ticker(Ticker(symbol="ZZZZ", type="candidate"))

    assert record.price is None
    assert record.ohlcv == []
    assert len(record.errors) > 0
