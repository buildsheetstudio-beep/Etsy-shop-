from datetime import datetime, timezone

from stock_agent.agents.news_agent import NewsAgent
from stock_agent.models import Ticker


class FakeFinnhubClient:
    """Stands in for FinnhubClient.get_news without making network calls."""

    def __init__(self, news_by_symbol: dict):
        self._news_by_symbol = news_by_symbol

    def get_news(self, symbol, from_date, to_date):
        return self._news_by_symbol.get(symbol)


def test_classifies_earnings_headline():
    finnhub = FakeFinnhubClient(
        {
            "AAPL": [
                {
                    "headline": "Apple reports record quarterly earnings, beats estimates",
                    "summary": "Apple posted strong EPS growth this quarter.",
                    "source": "Reuters",
                    "url": "https://example.com/1",
                    "datetime": 1700000000,
                }
            ]
        }
    )
    agent = NewsAgent(finnhub_client=finnhub)

    events = agent.process_ticker("AAPL")

    assert len(events) == 1
    assert "earnings" in events[0].categories
    assert events[0].headline.startswith("Apple reports")
    assert events[0].published_at == datetime.fromtimestamp(
        1700000000, tz=timezone.utc
    ).isoformat()


def test_headline_can_match_multiple_categories():
    finnhub = FakeFinnhubClient(
        {
            "MSFT": [
                {
                    "headline": "Microsoft reports Q3 earnings and raises full-year guidance",
                    "summary": "",
                    "source": "Bloomberg",
                    "url": "https://example.com/2",
                    "datetime": 1700000100,
                }
            ]
        }
    )
    agent = NewsAgent(finnhub_client=finnhub)

    events = agent.process_ticker("MSFT")

    assert len(events) == 1
    assert set(events[0].categories) == {"earnings", "guidance"}


def test_unmatched_article_is_dropped_not_manufactured():
    finnhub = FakeFinnhubClient(
        {
            "NVDA": [
                {
                    "headline": "Analysts discuss chip industry trends at conference",
                    "summary": "A general industry roundup with no company-specific news.",
                    "source": "Wire",
                    "url": "https://example.com/3",
                    "datetime": 1700000200,
                }
            ]
        }
    )
    agent = NewsAgent(finnhub_client=finnhub)

    events = agent.process_ticker("NVDA")

    assert events == []


def test_no_news_returns_empty_list_not_error():
    finnhub = FakeFinnhubClient({"ZZZZ": []})
    agent = NewsAgent(finnhub_client=finnhub)

    assert agent.process_ticker("ZZZZ") == []


def test_news_feed_unavailable_returns_empty_list_not_error():
    finnhub = FakeFinnhubClient({})  # symbol missing entirely -> get_news returns None
    agent = NewsAgent(finnhub_client=finnhub)

    assert agent.process_ticker("UNKNOWN") == []


def test_run_keys_output_by_symbol():
    finnhub = FakeFinnhubClient(
        {
            "AAPL": [
                {
                    "headline": "Apple CEO steps down, names interim replacement",
                    "summary": "",
                    "source": "Reuters",
                    "url": "https://example.com/4",
                    "datetime": 1700000300,
                }
            ],
            "MSFT": [],
        }
    )
    agent = NewsAgent(finnhub_client=finnhub)

    result = agent.run(
        [Ticker(symbol="AAPL", type="owned"), Ticker(symbol="MSFT", type="watchlist")]
    )

    assert set(result.keys()) == {"AAPL", "MSFT"}
    assert "leadership" in result["AAPL"][0].categories
    assert result["MSFT"] == []


def test_rating_change_keywords():
    finnhub = FakeFinnhubClient(
        {
            "TSLA": [
                {
                    "headline": "Analyst downgrades stock, cuts price target lowered",
                    "summary": "",
                    "source": "Wire",
                    "url": "https://example.com/5",
                    "datetime": 1700000400,
                }
            ]
        }
    )
    agent = NewsAgent(finnhub_client=finnhub)

    events = agent.process_ticker("TSLA")

    assert "rating_change" in events[0].categories


def test_mna_keywords():
    finnhub = FakeFinnhubClient(
        {
            "XYZ": [
                {
                    "headline": "XYZ Corp to be acquired in all-cash merger deal",
                    "summary": "",
                    "source": "Wire",
                    "url": "https://example.com/6",
                    "datetime": 1700000500,
                }
            ]
        }
    )
    agent = NewsAgent(finnhub_client=finnhub)

    events = agent.process_ticker("XYZ")

    assert "m_and_a" in events[0].categories
