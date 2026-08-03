from stock_agent.clients.tavily_client import TavilySearchProvider


def test_search_maps_tavily_response_to_search_results(monkeypatch):
    provider = TavilySearchProvider(api_key="test-key")
    monkeypatch.setattr(
        provider,
        "post_json",
        lambda url, json_body: {
            "results": [
                {"title": "A", "content": "snippet a", "url": "https://a.example"},
                {"title": "B", "content": "snippet b", "url": "https://b.example"},
            ]
        },
    )

    results = provider.search("test query", max_results=5)

    assert len(results) == 2
    assert results[0].title == "A"
    assert results[0].snippet == "snippet a"
    assert results[0].url == "https://a.example"


def test_search_sends_api_key_and_query(monkeypatch):
    provider = TavilySearchProvider(api_key="test-key")
    captured = {}

    def fake_post_json(url, json_body):
        captured["url"] = url
        captured["json_body"] = json_body
        return {"results": []}

    monkeypatch.setattr(provider, "post_json", fake_post_json)
    provider.search("breakout stocks", max_results=7)

    assert captured["json_body"]["api_key"] == "test-key"
    assert captured["json_body"]["query"] == "breakout stocks"
    assert captured["json_body"]["max_results"] == 7
    assert captured["url"].endswith("/search")


def test_search_returns_empty_list_when_request_fails(monkeypatch):
    provider = TavilySearchProvider(api_key="test-key")
    monkeypatch.setattr(provider, "post_json", lambda url, json_body: None)

    assert provider.search("anything") == []


def test_search_returns_empty_list_when_results_key_missing(monkeypatch):
    provider = TavilySearchProvider(api_key="test-key")
    monkeypatch.setattr(provider, "post_json", lambda url, json_body: {})

    assert provider.search("anything") == []
