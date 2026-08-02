"""Local technical indicator computation from raw OHLCV closes.

DESIGN.md 6.2: "Computes technical indicators locally (20/50-day MA, RSI)
from raw OHLCV" — Twelve Data is only a backup source for these (DESIGN.md
section 3), so the primary path never depends on it.
"""

from __future__ import annotations

from typing import Optional, Sequence


def simple_moving_average(closes: Sequence[float], window: int) -> Optional[float]:
    """Average of the most recent `window` closes, or None if not enough data."""
    if len(closes) < window:
        return None
    return sum(closes[-window:]) / window


def rsi(closes: Sequence[float], period: int = 14) -> Optional[float]:
    """Wilder's RSI over the most recent `period` changes, or None if not enough data."""
    if len(closes) < period + 1:
        return None

    deltas = [closes[i] - closes[i - 1] for i in range(1, len(closes))]

    gains = [max(d, 0.0) for d in deltas]
    losses = [max(-d, 0.0) for d in deltas]

    avg_gain = sum(gains[:period]) / period
    avg_loss = sum(losses[:period]) / period

    for gain, loss in zip(gains[period:], losses[period:]):
        avg_gain = (avg_gain * (period - 1) + gain) / period
        avg_loss = (avg_loss * (period - 1) + loss) / period

    if avg_loss == 0:
        return 100.0
    rs = avg_gain / avg_loss
    return 100 - (100 / (1 + rs))
