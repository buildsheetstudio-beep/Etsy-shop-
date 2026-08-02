"""Structured records shared across agents."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Literal, Optional

TickerType = Literal["owned", "watchlist", "candidate"]
SetupType = Literal["momentum_breakout", "oversold_bounce", "undervalued_fundamentals"]
RiskTier = Literal["standard", "speculative", "unknown"]


@dataclass
class Ticker:
    symbol: str
    type: TickerType


@dataclass
class DiscoveryCandidate:
    """Raw output of the Discovery Agent — DESIGN.md section 6.1.

    Not yet verified: price_hint comes from whatever a search snippet
    mentioned, not a real quote. The Data Agent re-fetches the real price
    for these symbols before they reach the dashboard.
    """

    symbol: str
    setup_types: list[SetupType] = field(default_factory=list)
    risk_tier: RiskTier = "unknown"
    price_hint: Optional[float] = None
    rationale: list[str] = field(default_factory=list)


@dataclass
class OHLCVBar:
    date: str  # ISO date, e.g. "2026-07-30"
    open: float
    high: float
    low: float
    close: float
    volume: int


@dataclass
class AnalystData:
    strong_buy: int = 0
    buy: int = 0
    hold: int = 0
    sell: int = 0
    strong_sell: int = 0
    price_target_high: Optional[float] = None
    price_target_low: Optional[float] = None
    price_target_mean: Optional[float] = None
    price_target_median: Optional[float] = None


@dataclass
class TickerRecord:
    """One structured record per ticker — the Data Agent's unit of output."""

    symbol: str
    type: TickerType

    price: Optional[float] = None
    change: Optional[float] = None
    percent_change: Optional[float] = None
    volume: Optional[int] = None

    week_52_high: Optional[float] = None
    week_52_low: Optional[float] = None

    ohlcv: list[OHLCVBar] = field(default_factory=list)

    ma_20: Optional[float] = None
    ma_50: Optional[float] = None
    rsi_14: Optional[float] = None

    analyst: Optional[AnalystData] = None

    fetched_at: str = field(
        default_factory=lambda: datetime.now(timezone.utc).isoformat()
    )
    # Populated when a data source fails so downstream agents (and the
    # dashboard) can see partial data rather than a silent gap. Per
    # DESIGN.md section 3: "agents should fail gracefully (skip/retry)
    # rather than assume 100% uptime."
    errors: list[str] = field(default_factory=list)
