from stock_agent.agents.synthesis_agent import SynthesisAgent
from stock_agent.models import DiscoveryCandidate, NotableEvent, TrendResult


def make_result(symbol="AAPL", type="owned", **overrides) -> TrendResult:
    defaults = dict(symbol=symbol, type=type, price=100.0, price_change_pct=0.5)
    defaults.update(overrides)
    return TrendResult(**defaults)


def section(html: str, heading: str) -> str:
    """Slice out one <h2>...</h2> section's HTML for section-scoped assertions."""
    start = html.index(f"<h2>{heading}</h2>")
    headings = ["Your Holdings", "Your Watchlist", "New Suggestions"]
    later_headings = [h for h in headings if h != heading]
    end = len(html)
    for h in later_headings:
        marker = f"<h2>{h}</h2>"
        if marker in html and html.index(marker) > start:
            end = min(end, html.index(marker))
    return html[start:end]


def test_holdings_watchlist_candidate_separation():
    results = [
        make_result(symbol="OWN1", type="owned"),
        make_result(symbol="WATCH1", type="watchlist"),
        make_result(symbol="CAND1", type="candidate"),
    ]
    html = SynthesisAgent().run(results, candidates=[])

    holdings_html = section(html, "Your Holdings")
    watchlist_html = section(html, "Your Watchlist")

    assert "OWN1" in holdings_html
    assert "WATCH1" not in holdings_html
    assert "CAND1" not in holdings_html

    assert "WATCH1" in watchlist_html
    assert "OWN1" not in watchlist_html


def test_ranking_by_strength_within_section():
    strong = make_result(symbol="STRONG", type="owned", ma_crossover="golden_cross", rating_change=True)
    weak = make_result(symbol="WEAK", type="owned", price_change_pct=0.2)

    html = SynthesisAgent().run([weak, strong], candidates=[])
    holdings_html = section(html, "Your Holdings")

    assert holdings_html.index("STRONG") < holdings_html.index("WEAK")


def test_holdings_stance_reconsider_for_strong_bearish():
    result = make_result(
        type="owned",
        price_change_pct=-6.0,
        ma_crossover="death_cross",
        rating_change=True,
    )
    html = SynthesisAgent().run([result], candidates=[])

    assert "reconsider" in section(html, "Your Holdings")


def test_holdings_stance_hold_for_quiet_bullish():
    result = make_result(type="owned", price_change_pct=0.5)
    html = SynthesisAgent().run([result], candidates=[])

    holdings_html = section(html, "Your Holdings")
    assert "hold" in holdings_html
    assert "reconsider" not in holdings_html


def test_watchlist_entry_signal_for_strong_bullish():
    result = make_result(
        type="watchlist",
        price_change_pct=6.0,
        ma_crossover="golden_cross",
        rating_change=True,
    )
    html = SynthesisAgent().run([result], candidates=[])

    assert "entry signal" in section(html, "Your Watchlist")


def test_watchlist_wait_for_bearish():
    result = make_result(type="watchlist", price_change_pct=-5.0, ma_crossover="death_cross")
    html = SynthesisAgent().run([result], candidates=[])

    watchlist_html = section(html, "Your Watchlist")
    assert "wait" in watchlist_html
    assert "entry signal" not in watchlist_html


def test_suggestions_joins_candidate_with_verified_trend_result():
    candidate = DiscoveryCandidate(
        symbol="CAND1", setup_types=["momentum_breakout"], risk_tier="speculative"
    )
    verified = make_result(symbol="CAND1", type="candidate", price=3.5)

    html = SynthesisAgent().run([verified], candidates=[candidate])
    suggestions_html = section(html, "New Suggestions")

    assert "CAND1" in suggestions_html
    assert "speculative" in suggestions_html
    assert "momentum breakout" in suggestions_html


def test_unverified_candidate_excluded_from_suggestions():
    candidate = DiscoveryCandidate(symbol="NOTVERIFIED", setup_types=["oversold_bounce"])

    html = SynthesisAgent().run([], candidates=[candidate])
    suggestions_html = section(html, "New Suggestions")

    assert "NOTVERIFIED" not in suggestions_html
    assert "No verified candidates" in suggestions_html


def test_empty_sections_render_placeholder_not_crash():
    html = SynthesisAgent().run([], candidates=[])

    assert "No meaningful changes this week." in section(html, "Your Holdings")
    assert "No meaningful changes this week." in section(html, "Your Watchlist")
    assert "No verified candidates this week." in section(html, "New Suggestions")


def test_notable_event_headline_is_html_escaped():
    result = make_result(
        type="owned",
        notable_events=[
            NotableEvent(categories=["earnings"], headline="<script>alert(1)</script>")
        ],
    )
    html = SynthesisAgent().run([result], candidates=[])

    assert "<script>alert(1)</script>" not in html
    assert "&lt;script&gt;" in html
