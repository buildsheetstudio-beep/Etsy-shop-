"""News/Catalyst Agent — DESIGN.md section 6.3. Not yet implemented.

Will pull each ticker's 7-day news feed (FinnhubClient.get_news already
supports this) and flag known catalyst categories: earnings, M&A,
guidance changes, leadership changes, major rating changes. Outputs an
empty "notable events" list when nothing happened rather than
manufacturing content.
"""

from __future__ import annotations

from ..models import TickerRecord


class NewsAgent:
    def run(self, records: list[TickerRecord]) -> dict[str, list[dict]]:
        raise NotImplementedError(
            "News/Catalyst Agent not yet built — see DESIGN.md section 6.3"
        )
