"""Discovery Agent — DESIGN.md section 6.1.

Generates untracked watchlist candidates by running screener-style web
search queries across three setup types (momentum/breakout, oversold
bounce, undervalued fundamentals), extracts ticker symbols and a rough
price from the search snippets, and ranks candidates by how many setup
types flagged them (a symbol surfaced by multiple screeners is a
stronger signal than one hit from a single query).

This agent does NOT verify anything — DESIGN.md 6.1 is explicit that
candidates are "passed through the Data Agent for real verification
before appearing on the dashboard — not just a raw screener echo," so
price_hint/risk_tier here are provisional, derived only from whatever a
search snippet happened to mention.

Which web search backend to use is an open item (DESIGN.md sections 3
and 10: no provider is named, free tier is "N/A"). This agent depends on
the WebSearchProvider protocol so any backend — Claude Code's own web
search, a search API, etc. — can be wired in at call time.
"""

from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass
from typing import Optional

from .. import config
from ..clients.websearch_client import WebSearchProvider
from ..models import DiscoveryCandidate, RiskTier, SetupType

# Cashtag ($AAPL) or exchange-qualified parenthetical (NASDAQ: AAPL) —
# deliberately narrower than "any uppercase word" to avoid false
# positives on ordinary capitalized words in search snippets.
_CASHTAG_RE = re.compile(r"\$([A-Z]{1,5})\b")
_EXCHANGE_RE = re.compile(r"\((?:NASDAQ|NYSE|AMEX)[:\s]+([A-Z]{1,5})\)")
_PRICE_RE = re.compile(r"\$(\d{1,4}(?:\.\d{1,2})?)\b")

_QUERIES: dict[SetupType, list[str]] = {
    "momentum_breakout": [
        "stocks breaking out this week high volume momentum",
        "stock price breakout volume surge today",
    ],
    "oversold_bounce": [
        "oversold stocks RSI below 30 pulled back from highs",
        "stocks bouncing off oversold levels this week",
    ],
    "undervalued_fundamentals": [
        "undervalued stocks low PE analyst price target above current price",
        "cheap stocks analyst upgrade price target this week",
    ],
}


@dataclass
class _Hit:
    symbol: str
    setup_type: SetupType
    price_hint: Optional[float]
    snippet: str


def _extract_symbols(text: str) -> list[str]:
    return list(dict.fromkeys(_CASHTAG_RE.findall(text) + _EXCHANGE_RE.findall(text)))


def _extract_price_hint(text: str) -> Optional[float]:
    match = _PRICE_RE.search(text)
    return float(match.group(1)) if match else None


def _risk_tier(price_hint: Optional[float]) -> RiskTier:
    if price_hint is None:
        return "unknown"
    return (
        "speculative"
        if price_hint < config.DISCOVERY_SPECULATIVE_PRICE_CEILING
        else "standard"
    )


class DiscoveryAgent:
    def __init__(self, search_provider: WebSearchProvider):
        self._search = search_provider

    def run(self, excluded_symbols: set[str]) -> list[DiscoveryCandidate]:
        """excluded_symbols: tickers already owned/watchlisted — DESIGN.md
        6.1 scopes this agent to "candidates the user doesn't currently
        track"."""
        hits_by_symbol: dict[str, list[_Hit]] = defaultdict(list)
        for hit in self._collect_hits():
            if hit.symbol in excluded_symbols:
                continue
            hits_by_symbol[hit.symbol].append(hit)

        candidates = [
            self._build_candidate(symbol, hits) for symbol, hits in hits_by_symbol.items()
        ]
        candidates.sort(key=lambda c: (len(c.setup_types), len(c.rationale)), reverse=True)
        return candidates

    def _collect_hits(self) -> list[_Hit]:
        hits: list[_Hit] = []
        for setup_type, queries in _QUERIES.items():
            for query in queries:
                results = self._search.search(
                    query, max_results=config.DISCOVERY_RESULTS_PER_QUERY
                )
                for result in results:
                    text = f"{result.title} {result.snippet}"
                    price_hint = _extract_price_hint(text)
                    for symbol in _extract_symbols(text):
                        hits.append(
                            _Hit(
                                symbol=symbol,
                                setup_type=setup_type,
                                price_hint=price_hint,
                                snippet=result.snippet,
                            )
                        )
        return hits

    @staticmethod
    def _build_candidate(symbol: str, hits: list[_Hit]) -> DiscoveryCandidate:
        setup_types = list(dict.fromkeys(hit.setup_type for hit in hits))
        price_hints = [hit.price_hint for hit in hits if hit.price_hint is not None]
        price_hint = price_hints[0] if price_hints else None
        rationale = list(dict.fromkeys(hit.snippet for hit in hits if hit.snippet))
        return DiscoveryCandidate(
            symbol=symbol,
            setup_types=setup_types,
            risk_tier=_risk_tier(price_hint),
            price_hint=price_hint,
            rationale=rationale,
        )
