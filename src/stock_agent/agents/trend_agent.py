"""Trend Agent — DESIGN.md section 6.4. Not yet implemented.

Will read last week's snapshot (storage.load_json(config.SNAPSHOT_FILE)),
diff this week's TickerRecords against it (% price change, MA crossovers,
RSI zone changes, analyst target revisions), drop tickers with no
meaningful change and no news, and overwrite the snapshot for next week
via storage.save_json. This is also where DESIGN.md section 10's open
item — exact crossover/overbought/oversold thresholds — needs to be
decided and encoded.
"""

from __future__ import annotations

from ..models import TickerRecord


class TrendAgent:
    def run(self, records: list[TickerRecord]) -> list[TickerRecord]:
        raise NotImplementedError(
            "Trend Agent not yet built — see DESIGN.md section 6.4"
        )
