"""Synthesis Agent — DESIGN.md section 6.5. Not yet implemented.

Will combine Discovery, Data, News, and Trend output into the three-section
HTML dashboard (Your Holdings / Your Watchlist / New Suggestions), ranked
within each section by signal strength.
"""

from __future__ import annotations

from ..models import TickerRecord


class SynthesisAgent:
    def run(
        self,
        records: list[TickerRecord],
        news: dict[str, list[dict]],
    ) -> str:
        """Returns rendered dashboard HTML."""
        raise NotImplementedError(
            "Synthesis Agent not yet built — see DESIGN.md section 6.5"
        )
