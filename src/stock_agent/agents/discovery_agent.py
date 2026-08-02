"""Discovery Agent — DESIGN.md section 6.1. Not yet implemented.

Will generate untracked watchlist candidates from screener-style web
searches across three setup types (momentum/breakout, oversold bounce,
undervalued fundamentals), split into standard (~$5+) and speculative
(<$5) risk tiers. Candidates are meant to be handed to DataAgent for
verification before reaching the dashboard — they are not built or
scored here.
"""

from __future__ import annotations

from ..models import Ticker


class DiscoveryAgent:
    def run(self) -> list[Ticker]:
        raise NotImplementedError(
            "Discovery Agent not yet built — see DESIGN.md section 6.1"
        )
