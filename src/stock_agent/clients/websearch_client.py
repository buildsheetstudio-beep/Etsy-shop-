"""Web search abstraction for the Discovery Agent — DESIGN.md section 3.

DESIGN.md lists "Web search" as a supplementary source for screener-style
candidate lists but names no concrete provider (free tier: "N/A") — which
backend to use (Claude Code's own web search when this runs as an agent
script, a search API, etc.) is an open item in section 10. DiscoveryAgent
depends on this Protocol instead of a specific vendor so a real backend
can be wired in later without touching screener logic.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol


@dataclass
class SearchResult:
    title: str
    snippet: str
    url: str = ""


class WebSearchProvider(Protocol):
    def search(self, query: str, max_results: int = 10) -> list[SearchResult]:
        ...
