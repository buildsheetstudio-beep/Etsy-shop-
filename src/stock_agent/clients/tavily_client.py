"""Tavily web search client — concrete WebSearchProvider for the Discovery
Agent (DESIGN.md section 6.1).

DESIGN.md section 3 lists "Web search" as a supplementary source without
naming a backend (free tier: "N/A" — an open item in section 10). Tavily
was chosen because it's built for agent/LLM use cases: results come back
as clean title/url/content rather than raw HTML to scrape, and its free
tier (1,000 searches/month as of this writing) comfortably covers the
Discovery Agent's 6 queries/week.
"""

from __future__ import annotations

from .. import config
from .base import RateLimitedClient
from .websearch_client import SearchResult


class TavilySearchProvider(RateLimitedClient):
    def __init__(self, api_key: str = config.TAVILY_API_KEY):
        super().__init__(min_seconds_between_calls=config.TAVILY_MIN_SECONDS_BETWEEN_CALLS)
        self._api_key = api_key

    def search(self, query: str, max_results: int = 10) -> list[SearchResult]:
        data = self.post_json(
            f"{config.TAVILY_BASE_URL}/search",
            {
                "api_key": self._api_key,
                "query": query,
                "max_results": max_results,
                "search_depth": "basic",
            },
        )
        if not data or "results" not in data:
            return []
        return [
            SearchResult(
                title=item.get("title", ""),
                snippet=item.get("content", ""),
                url=item.get("url", ""),
            )
            for item in data["results"]
        ]
