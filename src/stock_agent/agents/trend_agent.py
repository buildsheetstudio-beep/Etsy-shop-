"""Trend Agent — DESIGN.md section 6.4.

Diffs this week's Data Agent output against last week's saved snapshot —
% price change, MA crossovers, RSI zone changes, analyst rating changes
and price-target revisions — and carries through the Data Agent's
volume-spike flag and the News Agent's notable events. This is the
filtering layer: run() returns only tickers judged "meaningful"; a
ticker with no signal and no news is dropped rather than padding the
report (DESIGN.md 6.4). The snapshot is overwritten for next week's
comparison once every ticker has been diffed.

Threshold values are DESIGN.md section 10's open item, resolved in
config.py (RSI_OVERSOLD/OVERBOUGHT, VOLUME_SPIKE_*, PRICE_TARGET_
REVISION_PCT, MEANINGFUL_PRICE_CHANGE_PCT) — see that module's docstring
for rationale.

DESIGN.md section 8 describes the snapshot as price/MA/RSI/analyst
target. This implementation also stores each ticker's recommendation
counts so rating changes (DESIGN.md section 5's third signal category)
are diffable week over week instead of only reflecting current
sentiment — the same kind of minor, documented extension used for the
Data Agent's OHLCV lookback window.
"""

from __future__ import annotations

import logging
from typing import Optional

from .. import config, storage
from ..models import MACrossoverType, NotableEvent, RSIZone, TickerRecord, TrendResult

logger = logging.getLogger(__name__)


def _rsi_zone(rsi: Optional[float]) -> Optional[RSIZone]:
    if rsi is None:
        return None
    if rsi <= config.RSI_OVERSOLD:
        return "oversold"
    if rsi >= config.RSI_OVERBOUGHT:
        return "overbought"
    return "neutral"


def _ma_sign(ma_20: Optional[float], ma_50: Optional[float]) -> Optional[int]:
    if ma_20 is None or ma_50 is None:
        return None
    if ma_20 == ma_50:
        return 0
    return 1 if ma_20 > ma_50 else -1


def _ma_crossover(prev_sign: Optional[int], current_sign: Optional[int]) -> Optional[MACrossoverType]:
    if prev_sign is None or current_sign is None or prev_sign == current_sign:
        return None
    if current_sign > 0:
        return "golden_cross"
    if current_sign < 0:
        return "death_cross"
    return None


def _recommendation_counts(record: TickerRecord) -> Optional[dict]:
    if record.analyst is None:
        return None
    return {
        "strong_buy": record.analyst.strong_buy,
        "buy": record.analyst.buy,
        "hold": record.analyst.hold,
        "sell": record.analyst.sell,
        "strong_sell": record.analyst.strong_sell,
    }


def _snapshot_entry(record: TickerRecord) -> dict:
    return {
        "price": record.price,
        "ma_20": record.ma_20,
        "ma_50": record.ma_50,
        "rsi_14": record.rsi_14,
        "price_target_mean": record.analyst.price_target_mean if record.analyst else None,
        "recommendation_counts": _recommendation_counts(record),
    }


class TrendAgent:
    def run(
        self,
        records: list[TickerRecord],
        news_by_symbol: Optional[dict[str, list[NotableEvent]]] = None,
    ) -> list[TrendResult]:
        news_by_symbol = news_by_symbol or {}
        previous_snapshot = storage.load_json(config.SNAPSHOT_FILE)

        all_results = [
            self.diff_ticker(
                record, previous_snapshot.get(record.symbol), news_by_symbol.get(record.symbol, [])
            )
            for record in records
        ]

        new_snapshot = {record.symbol: _snapshot_entry(record) for record in records}
        storage.save_json(config.SNAPSHOT_FILE, new_snapshot)

        meaningful_results = [r for r in all_results if r.meaningful]
        logger.info(
            "Kept %d of %d tickers after trend filtering", len(meaningful_results), len(all_results)
        )
        return meaningful_results

    @staticmethod
    def diff_ticker(
        record: TickerRecord,
        previous: Optional[dict],
        notable_events: list[NotableEvent],
    ) -> TrendResult:
        is_new = previous is None

        price_change_pct = None
        if not is_new and previous.get("price") not in (None, 0) and record.price is not None:
            price_change_pct = (record.price - previous["price"]) / previous["price"] * 100

        rsi_zone = _rsi_zone(record.rsi_14)
        rsi_zone_changed = False
        if not is_new:
            prev_zone = _rsi_zone(previous.get("rsi_14"))
            rsi_zone_changed = prev_zone is not None and rsi_zone is not None and prev_zone != rsi_zone

        ma_crossover = None
        if not is_new:
            prev_sign = _ma_sign(previous.get("ma_20"), previous.get("ma_50"))
            current_sign = _ma_sign(record.ma_20, record.ma_50)
            ma_crossover = _ma_crossover(prev_sign, current_sign)

        rating_change = False
        if not is_new:
            prev_counts = previous.get("recommendation_counts")
            current_counts = _recommendation_counts(record)
            rating_change = (
                prev_counts is not None
                and current_counts is not None
                and prev_counts != current_counts
            )

        price_target_revision_pct = None
        if not is_new and record.analyst is not None:
            prev_target = previous.get("price_target_mean")
            current_target = record.analyst.price_target_mean
            if prev_target not in (None, 0) and current_target is not None:
                pct = (current_target - prev_target) / prev_target
                if abs(pct) >= config.PRICE_TARGET_REVISION_PCT:
                    price_target_revision_pct = pct * 100

        volume_spike = bool(record.volume_spike)

        meaningful = (
            is_new
            or bool(notable_events)
            or ma_crossover is not None
            or rsi_zone_changed
            or rating_change
            or price_target_revision_pct is not None
            or volume_spike
            or (
                price_change_pct is not None
                and abs(price_change_pct) >= config.MEANINGFUL_PRICE_CHANGE_PCT * 100
            )
        )

        return TrendResult(
            symbol=record.symbol,
            type=record.type,
            is_new=is_new,
            price=record.price,
            price_change_pct=price_change_pct,
            rsi_14=record.rsi_14,
            rsi_zone=rsi_zone,
            rsi_zone_changed=rsi_zone_changed,
            ma_crossover=ma_crossover,
            volume_spike=volume_spike,
            rating_change=rating_change,
            price_target_revision_pct=price_target_revision_pct,
            notable_events=notable_events,
            meaningful=meaningful,
        )
