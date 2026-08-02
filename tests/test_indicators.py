from stock_agent.indicators import rsi, simple_moving_average


def test_sma_insufficient_data_returns_none():
    assert simple_moving_average([1, 2, 3], window=5) is None


def test_sma_basic():
    closes = [1, 2, 3, 4, 5]
    assert simple_moving_average(closes, window=5) == 3.0
    assert simple_moving_average(closes, window=2) == 4.5


def test_rsi_insufficient_data_returns_none():
    assert rsi([1.0] * 10, period=14) is None


def test_rsi_all_gains_is_100():
    closes = [float(i) for i in range(1, 20)]  # strictly increasing
    assert rsi(closes, period=14) == 100.0


def test_rsi_all_losses_is_0():
    closes = [float(i) for i in range(20, 1, -1)]  # strictly decreasing
    assert rsi(closes, period=14) == 0.0


def test_rsi_mixed_is_between_0_and_100():
    closes = [10, 11, 10, 12, 11, 13, 12, 14, 13, 15, 14, 16, 15, 17, 16]
    value = rsi(closes, period=14)
    assert value is not None
    assert 0.0 < value < 100.0
