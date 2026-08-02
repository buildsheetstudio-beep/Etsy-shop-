"""Weekly pipeline orchestration — DESIGN.md section 7.

    Discovery Agent ─┐
                      ├─→ Data Agent ─→ News Agent ─→ Trend Agent ─→ Synthesis Agent ─→ Dashboard
       (candidates) ──┘   (verifies)

Data Agent and Discovery Agent are implemented; News/Trend/Synthesis are
still stubs that raise NotImplementedError. run_discovery_stage() takes a
WebSearchProvider argument rather than defaulting to one, since which
search backend to use is itself an open item (DESIGN.md section 10) — no
default exists yet to wire in.
"""

from __future__ import annotations

import dataclasses
import logging

from . import config, storage
from .agents.data_agent import DataAgent
from .agents.discovery_agent import DiscoveryAgent
from .clients.websearch_client import WebSearchProvider
from .models import DiscoveryCandidate, Ticker, TickerRecord

logger = logging.getLogger(__name__)


def run_data_stage(tickers: list[Ticker] | None = None) -> list[TickerRecord]:
    if tickers is None:
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


def run_discovery_stage(search_provider: WebSearchProvider) -> list[DiscoveryCandidate]:
    tickers = storage.load_tickers(config.TICKERS_FILE)
    excluded_symbols = {t.symbol for t in tickers}

    candidates = DiscoveryAgent(search_provider).run(excluded_symbols)

    output_path = config.RUN_OUTPUT_DIR / "discovery_agent_output.json"
    storage.save_json(
        output_path,
        {"candidates": [dataclasses.asdict(c) for c in candidates]},
    )
    logger.info("Wrote %d candidates to %s", len(candidates), output_path)

    return candidates
