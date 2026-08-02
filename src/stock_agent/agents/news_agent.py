"""News/Catalyst Agent — DESIGN.md section 6.3.

Pulls each ticker's news feed for the past 7 days (FinnhubClient.get_news),
classifies headlines/summaries into known catalyst categories — earnings,
M&A, guidance changes, leadership changes, major analyst rating changes —
via keyword matching, and returns a "notable events" list per ticker.
Finnhub's own `category` field on news items is too coarse (e.g. "company
news", "top news") to distinguish these, so classification runs on the
headline + summary text instead.

Articles that don't match any known category are dropped rather than
included as an uncategorized catch-all — DESIGN.md 6.3: output is "empty
if nothing happened — no manufactured content", so a ticker with no
matching news gets an empty list, not padding.
"""

from __future__ import annotations

import logging
from datetime import datetime, timedelta, timezone

from .. import config
from ..clients.finnhub_client import FinnhubClient
from ..models import CatalystCategory, NotableEvent, Ticker

logger = logging.getLogger(__name__)

_CATEGORY_KEYWORDS: dict[CatalystCategory, list[str]] = {
    "earnings": [
        "earnings",
        "eps",
        "quarterly results",
        "quarterly report",
        "beats estimates",
        "misses estimates",
        "revenue beat",
        "revenue miss",
    ],
    "m_and_a": [
        "acquisition",
        "acquire",
        "acquires",
        "acquired",
        "merger",
        "merge with",
        "to be acquired",
        "buyout",
        "takeover",
    ],
    "guidance": [
        "guidance",
        "raises forecast",
        "cuts forecast",
        "raises outlook",
        "lowers outlook",
        "raises full-year",
        "lowers full-year",
    ],
    "leadership": [
        "ceo",
        "chief executive",
        "cfo",
        "chief financial officer",
        "steps down",
        "resigns",
        "resignation",
        "appoints new",
        "names new",
        "names interim",
    ],
    "rating_change": [
        "upgrade",
        "upgrades",
        "downgrade",
        "downgrades",
        "price target raised",
        "price target lowered",
        "initiates coverage",
        "reiterates rating",
    ],
}


def _classify(text: str) -> list[CatalystCategory]:
    lowered = text.lower()
    return [
        category
        for category, keywords in _CATEGORY_KEYWORDS.items()
        if any(keyword in lowered for keyword in keywords)
    ]


def _to_iso(unix_ts: float | None) -> str:
    if not unix_ts:
        return ""
    return datetime.fromtimestamp(unix_ts, tz=timezone.utc).isoformat()


class NewsAgent:
    def __init__(self, finnhub_client: FinnhubClient | None = None):
        self._finnhub = finnhub_client or FinnhubClient()

    def run(self, tickers: list[Ticker]) -> dict[str, list[NotableEvent]]:
        return {ticker.symbol: self.process_ticker(ticker.symbol) for ticker in tickers}

    def process_ticker(self, symbol: str) -> list[NotableEvent]:
        to_date = datetime.now(timezone.utc).date()
        from_date = to_date - timedelta(days=config.NEWS_LOOKBACK_DAYS)

        raw_news = self._finnhub.get_news(symbol, from_date.isoformat(), to_date.isoformat())
        if not raw_news:
            if raw_news is None:
                logger.warning("%s: news feed unavailable", symbol)
            return []

        events = []
        for item in raw_news:
            text = f"{item.get('headline', '')} {item.get('summary', '')}"
            categories = _classify(text)
            if not categories:
                continue
            events.append(
                NotableEvent(
                    categories=categories,
                    headline=item.get("headline", ""),
                    summary=item.get("summary", ""),
                    source=item.get("source", ""),
                    url=item.get("url", ""),
                    published_at=_to_iso(item.get("datetime")),
                )
            )
        return events
