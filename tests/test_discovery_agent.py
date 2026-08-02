from stock_agent.agents.discovery_agent import DiscoveryAgent
from stock_agent.clients.websearch_client import SearchResult


class FakeWebSearchProvider:
    """Maps each query string to a canned list of SearchResults."""

    def __init__(self, results_by_query: dict[str, list[SearchResult]]):
        self._results_by_query = results_by_query

    def search(self, query, max_results=10):
        return self._results_by_query.get(query, [])


RESULTS_BY_QUERY = {
    # momentum_breakout
    "stocks breaking out this week high volume momentum": [
        SearchResult(
            title="Breakout watch",
            snippet="$ABCD breaking out on volume, now trading near $12.50",
        )
    ],
    "stock price breakout volume surge today": [
        SearchResult(title="Momentum play", snippet="$WXYZ momentum play at $3.20")
    ],
    # oversold_bounce
    "oversold stocks RSI below 30 pulled back from highs": [
        SearchResult(
            title="Oversold bounce",
            snippet="$ABCD oversold bounce candidate around $12.75",
        )
    ],
    "stocks bouncing off oversold levels this week": [
        SearchResult(title="RSI watch", snippet="$OVSD RSI below 30 near $45")
    ],
    # undervalued_fundamentals
    "undervalued stocks low PE analyst price target above current price": [
        SearchResult(title="Value pick", snippet="$ABCD undervalued with target above $15")
    ],
    "cheap stocks analyst upgrade price target this week": [
        SearchResult(title="Cheap stock", snippet="$UVFD trading cheap at $60 vs target")
    ],
}


def test_extracts_cashtag_and_price_hint():
    provider = FakeWebSearchProvider(RESULTS_BY_QUERY)
    candidates = DiscoveryAgent(provider).run(excluded_symbols=set())

    by_symbol = {c.symbol: c for c in candidates}
    assert "ABCD" in by_symbol
    assert by_symbol["ABCD"].price_hint == 12.50


def test_multi_setup_type_symbol_ranks_first():
    provider = FakeWebSearchProvider(RESULTS_BY_QUERY)
    candidates = DiscoveryAgent(provider).run(excluded_symbols=set())

    # ABCD is hit by all three setup types; everything else by at most one.
    assert candidates[0].symbol == "ABCD"
    assert set(candidates[0].setup_types) == {
        "momentum_breakout",
        "oversold_bounce",
        "undervalued_fundamentals",
    }


def test_risk_tier_assigned_from_price_hint():
    provider = FakeWebSearchProvider(RESULTS_BY_QUERY)
    candidates = DiscoveryAgent(provider).run(excluded_symbols=set())
    by_symbol = {c.symbol: c for c in candidates}

    assert by_symbol["ABCD"].risk_tier == "standard"  # ~$12.50
    assert by_symbol["WXYZ"].risk_tier == "speculative"  # $3.20 < $5


def test_excludes_already_tracked_symbols():
    provider = FakeWebSearchProvider(RESULTS_BY_QUERY)
    candidates = DiscoveryAgent(provider).run(excluded_symbols={"OVSD", "UVFD"})

    symbols = {c.symbol for c in candidates}
    assert "OVSD" not in symbols
    assert "UVFD" not in symbols
    assert "ABCD" in symbols


def test_no_search_results_returns_empty_list():
    provider = FakeWebSearchProvider({})
    candidates = DiscoveryAgent(provider).run(excluded_symbols=set())
    assert candidates == []


def test_rationale_deduplicated_across_repeated_snippets():
    results = {
        "stocks breaking out this week high volume momentum": [
            SearchResult(title="a", snippet="$DUPE breaking out at $9.00"),
            SearchResult(title="b", snippet="$DUPE breaking out at $9.00"),
        ]
    }
    provider = FakeWebSearchProvider(results)
    candidates = DiscoveryAgent(provider).run(excluded_symbols=set())

    assert len(candidates) == 1
    assert candidates[0].rationale == ["$DUPE breaking out at $9.00"]
