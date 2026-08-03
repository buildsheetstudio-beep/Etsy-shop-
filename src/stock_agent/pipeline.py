"""Weekly pipeline orchestration — DESIGN.md section 7.

    Discovery Agent ─┐
                      ├─→ Data Agent ─→ News Agent ─→ Trend Agent ─→ Synthesis Agent ─→ Dashboard
       (candidates) ──┘   (verifies)

All five agents are implemented. DESIGN.md section 10's web-search-backend
open item is resolved with Tavily (clients/tavily_client.py) — Discovery
Agent and run_discovery_stage() default to it, so run_synthesis_stage()
with no arguments now runs the real end-to-end pipeline, including New
Suggestions. Pass include_suggestions=False to skip Discovery entirely
(e.g. to avoid web search calls in a constrained environment), or pass a
different search_provider to swap the backend.
"""

from __future__ import annotations

import dataclasses
import logging

from . import config, storage
from .agents.data_agent import DataAgent
from .agents.discovery_agent import DiscoveryAgent
from .agents.news_agent import NewsAgent
from .agents.synthesis_agent import SynthesisAgent
from .agents.trend_agent import TrendAgent
from .clients.websearch_client import WebSearchProvider
from .models import DiscoveryCandidate, NotableEvent, Ticker, TickerRecord, TrendResult

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


def run_discovery_stage(
    search_provider: WebSearchProvider | None = None,
) -> list[DiscoveryCandidate]:
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


def run_news_stage(tickers: list[Ticker] | None = None) -> dict[str, list[NotableEvent]]:
    if tickers is None:
        tickers = storage.load_tickers(config.TICKERS_FILE)
        logger.info("Loaded %d tickers from %s", len(tickers), config.TICKERS_FILE)

    events_by_symbol = NewsAgent().run(tickers)

    output_path = config.RUN_OUTPUT_DIR / "news_agent_output.json"
    storage.save_json(
        output_path,
        {
            "events": {
                symbol: [dataclasses.asdict(e) for e in events]
                for symbol, events in events_by_symbol.items()
            }
        },
    )
    total_events = sum(len(events) for events in events_by_symbol.values())
    logger.info(
        "Wrote %d notable events across %d tickers to %s",
        total_events,
        len(events_by_symbol),
        output_path,
    )

    return events_by_symbol


def run_trend_stage(
    records: list[TickerRecord] | None = None,
    news_by_symbol: dict[str, list[NotableEvent]] | None = None,
) -> list[TrendResult]:
    if records is None:
        records = run_data_stage()
    if news_by_symbol is None:
        tickers = [Ticker(symbol=r.symbol, type=r.type) for r in records]
        news_by_symbol = run_news_stage(tickers)

    results = TrendAgent().run(records, news_by_symbol)

    output_path = config.RUN_OUTPUT_DIR / "trend_agent_output.json"
    storage.save_json(
        output_path,
        {"results": [dataclasses.asdict(r) for r in results]},
    )
    logger.info(
        "Wrote %d meaningful trend results (of %d tickers) to %s",
        len(results),
        len(records),
        output_path,
    )

    return results


def run_synthesis_stage(
    search_provider: WebSearchProvider | None = None,
    include_suggestions: bool = True,
    trend_results: list[TrendResult] | None = None,
    candidates: list[DiscoveryCandidate] | None = None,
) -> str:
    """Renders the dashboard. If trend_results isn't supplied, chains the
    earlier stages: when include_suggestions is True (the default),
    Discovery candidates are verified through the Data Agent alongside
    owned/watchlist tickers so "New Suggestions" is populated; set it to
    False to skip Discovery (and its web search calls) entirely.
    """
    if trend_results is None:
        tickers = storage.load_tickers(config.TICKERS_FILE)
        if include_suggestions:
            candidates = run_discovery_stage(search_provider)
            tickers = tickers + [Ticker(symbol=c.symbol, type="candidate") for c in candidates]
        else:
            candidates = candidates or []
        trend_results = run_trend_stage(records=run_data_stage(tickers))
    else:
        candidates = candidates or []

    dashboard_html = SynthesisAgent().run(trend_results, candidates)

    config.DASHBOARD_FILE.parent.mkdir(parents=True, exist_ok=True)
    config.DASHBOARD_FILE.write_text(dashboard_html)
    logger.info("Wrote dashboard to %s", config.DASHBOARD_FILE)

    return dashboard_html
