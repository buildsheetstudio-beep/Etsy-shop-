"""Weekly pipeline orchestration — DESIGN.md section 7.

    Discovery Agent ─┐
                      ├─→ Data Agent ─→ News Agent ─→ Trend Agent ─→ Synthesis Agent ─→ Dashboard
       (candidates) ──┘   (verifies)

Only the Data Agent stage is implemented so far. Discovery/News/Trend/
Synthesis are wired in as stubs that raise NotImplementedError; this
function runs the Data Agent stage against the maintained ticker list
(owned + watchlist) and writes its output to data/run/ so the shape of
the handoff is established before the later stages land.
"""

from __future__ import annotations

import dataclasses
import logging

from . import config, storage
from .agents.data_agent import DataAgent
from .models import TickerRecord

logger = logging.getLogger(__name__)


def run_data_stage() -> list[TickerRecord]:
    tickers = storage.load_tickers(config.TICKERS_FILE)
    logger.info("Loaded %d tickers from %s", len(tickers), config.TICKERS_FILE)

    records = DataAgent().run(tickers)

    output_path = config.RUN_OUTPUT_DIR / "data_agent_output.json"
    storage.save_json(
        output_path,
        {"records": [dataclasses.asdict(r) for r in records]},
    )
    logger.info("Wrote %d records to %s", len(records), output_path)

    return records
