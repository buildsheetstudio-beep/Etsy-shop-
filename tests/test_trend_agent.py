from stock_agent import config
from stock_agent.agents.trend_agent import TrendAgent
from stock_agent.models import AnalystData, NotableEvent, TickerRecord


def make_record(symbol="AAPL", type="owned", **overrides) -> TickerRecord:
    defaults = dict(
        symbol=symbol,
        type=type,
        price=100.0,
        ma_20=50.0,
        ma_50=40.0,  # ma_20 > ma_50
        rsi_14=50.0,  # neutral
        volume_spike=False,
        analyst=AnalystData(
            strong_buy=1, buy=2, hold=3, sell=0, strong_sell=0, price_target_mean=110.0
        ),
    )
    defaults.update(overrides)
    return TickerRecord(**defaults)


def test_first_run_is_new_and_always_meaningful():
    result = TrendAgent.diff_ticker(make_record(), previous=None, notable_events=[])

    assert result.is_new is True
    assert result.meaningful is True
    assert result.price_change_pct is None


def test_dropped_when_nothing_changed():
    record = make_record()
    previous = {
        "price": 100.0,
        "ma_20": 50.0,
        "ma_50": 40.0,
        "rsi_14": 50.0,
        "price_target_mean": 110.0,
        "recommendation_counts": {
            "strong_buy": 1,
            "buy": 2,
            "hold": 3,
            "sell": 0,
            "strong_sell": 0,
        },
    }

    result = TrendAgent.diff_ticker(record, previous, notable_events=[])

    assert result.meaningful is False


def test_price_change_below_threshold_not_meaningful():
    record = make_record(price=101.0)  # +1%, below the 3% bar
    previous = {"price": 100.0, "ma_20": 50.0, "ma_50": 40.0, "rsi_14": 50.0}

    result = TrendAgent.diff_ticker(record, previous, notable_events=[])

    assert result.price_change_pct == 1.0
    assert result.meaningful is False


def test_price_change_above_threshold_is_meaningful():
    record = make_record(price=105.0)  # +5%, above the 3% bar
    previous = {"price": 100.0, "ma_20": 50.0, "ma_50": 40.0, "rsi_14": 50.0}

    result = TrendAgent.diff_ticker(record, previous, notable_events=[])

    assert result.price_change_pct == 5.0
    assert result.meaningful is True


def test_rsi_zone_change_to_oversold_is_meaningful():
    record = make_record(rsi_14=25.0)  # oversold
    previous = {"price": 100.0, "ma_20": 50.0, "ma_50": 40.0, "rsi_14": 50.0}  # neutral

    result = TrendAgent.diff_ticker(record, previous, notable_events=[])

    assert result.rsi_zone == "oversold"
    assert result.rsi_zone_changed is True
    assert result.meaningful is True


def test_rsi_same_zone_not_flagged():
    record = make_record(rsi_14=55.0)  # still neutral
    previous = {"price": 100.0, "ma_20": 50.0, "ma_50": 40.0, "rsi_14": 50.0}  # neutral

    result = TrendAgent.diff_ticker(record, previous, notable_events=[])

    assert result.rsi_zone_changed is False


def test_golden_cross_detected():
    record = make_record(ma_20=55.0, ma_50=50.0)  # ma_20 > ma_50 now
    previous = {"price": 100.0, "ma_20": 40.0, "ma_50": 50.0, "rsi_14": 50.0}  # ma_20 < ma_50 last week

    result = TrendAgent.diff_ticker(record, previous, notable_events=[])

    assert result.ma_crossover == "golden_cross"
    assert result.meaningful is True


def test_death_cross_detected():
    record = make_record(ma_20=45.0, ma_50=50.0)  # ma_20 < ma_50 now
    previous = {"price": 100.0, "ma_20": 55.0, "ma_50": 50.0, "rsi_14": 50.0}  # ma_20 > ma_50 last week

    result = TrendAgent.diff_ticker(record, previous, notable_events=[])

    assert result.ma_crossover == "death_cross"


def test_no_crossover_when_relative_position_unchanged():
    record = make_record(ma_20=52.0, ma_50=40.0)  # still ma_20 > ma_50
    previous = {"price": 100.0, "ma_20": 50.0, "ma_50": 40.0, "rsi_14": 50.0}

    result = TrendAgent.diff_ticker(record, previous, notable_events=[])

    assert result.ma_crossover is None


def test_rating_change_detected():
    record = make_record()  # buy=2 in defaults
    previous = {
        "price": 100.0,
        "ma_20": 50.0,
        "ma_50": 40.0,
        "rsi_14": 50.0,
        "recommendation_counts": {
            "strong_buy": 1,
            "buy": 1,  # was 1, now 2
            "hold": 3,
            "sell": 0,
            "strong_sell": 0,
        },
    }

    result = TrendAgent.diff_ticker(record, previous, notable_events=[])

    assert result.rating_change is True
    assert result.meaningful is True


def test_price_target_revision_above_threshold():
    # 110 -> 100 is a ~9.1% move, above the 2% bar
    record = make_record()
    previous = {
        "price": 100.0,
        "ma_20": 50.0,
        "ma_50": 40.0,
        "rsi_14": 50.0,
        "price_target_mean": 100.0,
    }

    result = TrendAgent.diff_ticker(record, previous, notable_events=[])

    assert result.price_target_revision_pct is not None
    assert result.price_target_revision_pct > config.PRICE_TARGET_REVISION_PCT * 100
    assert result.meaningful is True


def test_price_target_revision_below_threshold_not_flagged():
    # 110 -> 109 is under 1%, below the 2% bar
    record = make_record()
    previous = {
        "price": 100.0,
        "ma_20": 50.0,
        "ma_50": 40.0,
        "rsi_14": 50.0,
        "price_target_mean": 109.0,
    }

    result = TrendAgent.diff_ticker(record, previous, notable_events=[])

    assert result.price_target_revision_pct is None


def test_volume_spike_carried_through_and_meaningful():
    record = make_record(volume_spike=True)
    previous = {"price": 100.0, "ma_20": 50.0, "ma_50": 40.0, "rsi_14": 50.0}

    result = TrendAgent.diff_ticker(record, previous, notable_events=[])

    assert result.volume_spike is True
    assert result.meaningful is True


def test_notable_event_alone_makes_ticker_meaningful():
    record = make_record()
    previous = {"price": 100.0, "ma_20": 50.0, "ma_50": 40.0, "rsi_14": 50.0}
    events = [NotableEvent(categories=["earnings"], headline="Beats estimates")]

    result = TrendAgent.diff_ticker(record, previous, notable_events=events)

    assert result.notable_events == events
    assert result.meaningful is True


def test_run_filters_and_writes_snapshot(tmp_path, monkeypatch):
    snapshot_path = tmp_path / "weekly_snapshot.json"
    monkeypatch.setattr(config, "SNAPSHOT_FILE", snapshot_path)

    unchanged = make_record(symbol="FLAT", price=100.0)
    moved = make_record(symbol="MOVR", price=100.0)

    # First run: no prior snapshot, so both are "new" and kept.
    first_results = TrendAgent().run([unchanged, moved], news_by_symbol={})
    assert {r.symbol for r in first_results} == {"FLAT", "MOVR"}
    assert snapshot_path.exists()

    # Second run: FLAT hasn't moved (dropped), MOVR jumped 10% (kept).
    unchanged_week2 = make_record(symbol="FLAT", price=100.0)
    moved_week2 = make_record(symbol="MOVR", price=110.0)

    second_results = TrendAgent().run([unchanged_week2, moved_week2], news_by_symbol={})

    assert {r.symbol for r in second_results} == {"MOVR"}
    assert second_results[0].price_change_pct == 10.0
