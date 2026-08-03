"""Synthesis Agent — DESIGN.md section 6.5.

Combines Trend Agent output (owned/watchlist/verified-candidate signals)
with Discovery Agent candidates (setup types, risk tier) into a single
self-contained HTML dashboard with three sections: Your Holdings, Your
Watchlist, New Suggestions. Ranks within each section by signal strength
and color-codes by signal direction (DESIGN.md section 9).

DESIGN.md doesn't define how "hold / watch / reconsider" (Holdings) or
"entry-signal" (Watchlist) map onto the underlying signals, or what
"signal strength" means numerically beyond one example ("MA crossover +
analyst downgrade ranks above a minor RSI dip"). Resolved here with the
config.py SYNTHESIS_* constants — see that module for the exact mapping.

Dashboard delivery is file-only for now (DESIGN.md section 10 leaves
email TBD); this agent only renders HTML, pipeline.run_synthesis_stage
writes it to disk.
"""

from __future__ import annotations

import html
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Optional

from .. import config
from ..models import DiscoveryCandidate, NotableEvent, TrendResult

Direction = str  # "bullish" | "bearish" | "neutral"


@dataclass
class _Scored:
    result: TrendResult
    strength: float
    direction: Direction


def _direction_score(result: TrendResult) -> float:
    score = 0.0
    if result.price_change_pct:
        score += 1 if result.price_change_pct > 0 else -1
    if result.ma_crossover == "golden_cross":
        score += 2
    elif result.ma_crossover == "death_cross":
        score -= 2
    if result.rsi_zone == "oversold":
        score += 1
    elif result.rsi_zone == "overbought":
        score -= 1
    if result.price_target_revision_pct:
        score += 1 if result.price_target_revision_pct > 0 else -1
    return score


def _direction(score: float) -> Direction:
    if score > 0:
        return "bullish"
    if score < 0:
        return "bearish"
    return "neutral"


def _strength(result: TrendResult) -> float:
    strength = 0.0
    if result.ma_crossover is not None:
        strength += 3
    if result.rsi_zone_changed:
        strength += 2
    if result.rating_change:
        strength += 2
    if result.price_target_revision_pct is not None:
        strength += 2
    if result.volume_spike:
        strength += 2
    strength += len(result.notable_events)
    if result.price_change_pct is not None:
        strength += min(abs(result.price_change_pct) / 10, 1.0)
    return strength


def _score(result: TrendResult) -> _Scored:
    return _Scored(
        result=result,
        strength=_strength(result),
        direction=_direction(_direction_score(result)),
    )


def _holdings_stance(scored: _Scored) -> str:
    if scored.direction == "bearish" and scored.strength >= config.SYNTHESIS_RECONSIDER_STRENGTH:
        return "reconsider"
    if scored.direction == "bearish" or scored.strength >= config.SYNTHESIS_WATCH_STRENGTH:
        return "watch"
    return "hold"


def _watchlist_entry_signal(scored: _Scored) -> str:
    if scored.direction == "bullish" and scored.strength >= config.SYNTHESIS_ENTRY_STRENGTH:
        return "entry signal"
    if scored.direction == "bullish" and scored.strength >= config.SYNTHESIS_POSSIBLE_ENTRY_STRENGTH:
        return "possible entry"
    return "wait"


def _esc(text: str) -> str:
    return html.escape(text)


def _badge(text: str, css_class: str) -> str:
    return f'<span class="badge {css_class}">{_esc(text)}</span>'


def _signal_badges(result: TrendResult) -> str:
    badges = []
    if result.ma_crossover:
        cls = "bullish" if result.ma_crossover == "golden_cross" else "bearish"
        badges.append(_badge(result.ma_crossover.replace("_", " "), cls))
    if result.rsi_zone_changed and result.rsi_zone:
        cls = {"oversold": "bullish", "overbought": "bearish"}.get(result.rsi_zone, "neutral")
        badges.append(_badge(f"RSI → {result.rsi_zone}", cls))
    if result.volume_spike:
        badges.append(_badge("volume spike", "neutral"))
    if result.rating_change:
        badges.append(_badge("rating change", "neutral"))
    if result.price_target_revision_pct is not None:
        cls = "bullish" if result.price_target_revision_pct > 0 else "bearish"
        badges.append(_badge(f"target {result.price_target_revision_pct:+.1f}%", cls))

    seen_categories: list[str] = []
    for event in result.notable_events:
        for category in event.categories:
            if category not in seen_categories:
                seen_categories.append(category)
    for category in seen_categories:
        badges.append(_badge(category.replace("_", " "), "neutral"))

    return "".join(badges) if badges else '<span class="muted">no active signals</span>'


def _event_list(events: list[NotableEvent]) -> str:
    if not events:
        return ""
    items = []
    for event in events:
        headline = _esc(event.headline)
        if event.url:
            headline = f'<a href="{_esc(event.url)}" target="_blank" rel="noopener">{headline}</a>'
        items.append(f"<li>{headline}</li>")
    return f'<ul class="events">{"".join(items)}</ul>'


def _signals_cell(result: TrendResult) -> str:
    return _signal_badges(result) + _event_list(result.notable_events)


def _price_cell(price: Optional[float]) -> str:
    return f"${price:,.2f}" if price is not None else "&mdash;"


def _change_cell(price_change_pct: Optional[float]) -> str:
    return f"{price_change_pct:+.1f}%" if price_change_pct is not None else "&mdash;"


def _holdings_row(scored: _Scored) -> str:
    r = scored.result
    stance = _holdings_stance(scored)
    return f"""
      <tr class="{scored.direction}">
        <td class="symbol">{_esc(r.symbol)}</td>
        <td>{_price_cell(r.price)}</td>
        <td>{_change_cell(r.price_change_pct)}</td>
        <td><span class="tag {scored.direction}">{_esc(stance)}</span></td>
        <td>{_signals_cell(r)}</td>
      </tr>"""


def _watchlist_row(scored: _Scored) -> str:
    r = scored.result
    entry = _watchlist_entry_signal(scored)
    return f"""
      <tr class="{scored.direction}">
        <td class="symbol">{_esc(r.symbol)}</td>
        <td>{_price_cell(r.price)}</td>
        <td>{_change_cell(r.price_change_pct)}</td>
        <td><span class="tag {scored.direction}">{_esc(entry)}</span></td>
        <td>{_signals_cell(r)}</td>
      </tr>"""


def _suggestion_row(candidate: DiscoveryCandidate, scored: _Scored) -> str:
    r = scored.result
    setup_badges = "".join(
        _badge(setup.replace("_", " "), "neutral") for setup in candidate.setup_types
    )
    return f"""
      <tr class="{scored.direction}">
        <td class="symbol">{_esc(r.symbol)}</td>
        <td>{_price_cell(r.price)}</td>
        <td><span class="tag {candidate.risk_tier}">{_esc(candidate.risk_tier)}</span></td>
        <td>{setup_badges}</td>
        <td>{_signals_cell(r)}</td>
      </tr>"""


_TABLE_HEAD = {
    "holdings": ("Symbol", "Price", "Change", "Stance", "Signals"),
    "watchlist": ("Symbol", "Price", "Change", "Entry Signal", "Signals"),
    "suggestions": ("Symbol", "Price", "Risk Tier", "Setup", "Signals"),
}


def _table(section_id: str, rows: list[str], empty_message: str) -> str:
    headers = "".join(f"<th>{h}</th>" for h in _TABLE_HEAD[section_id])
    body = "".join(rows) if rows else f'<tr><td colspan="5" class="muted">{_esc(empty_message)}</td></tr>'
    return f"""
    <table>
      <thead><tr>{headers}</tr></thead>
      <tbody>{body}</tbody>
    </table>"""


class SynthesisAgent:
    def run(
        self,
        trend_results: list[TrendResult],
        candidates: Optional[list[DiscoveryCandidate]] = None,
    ) -> str:
        candidates = candidates or []
        scored = [_score(r) for r in trend_results]

        holdings = sorted(
            (s for s in scored if s.result.type == "owned"),
            key=lambda s: s.strength,
            reverse=True,
        )
        watchlist = sorted(
            (s for s in scored if s.result.type == "watchlist"),
            key=lambda s: s.strength,
            reverse=True,
        )
        verified_candidates = {s.result.symbol: s for s in scored if s.result.type == "candidate"}
        suggestions = sorted(
            (
                (candidate, verified_candidates[candidate.symbol])
                for candidate in candidates
                if candidate.symbol in verified_candidates
            ),
            key=lambda pair: pair[1].strength,
            reverse=True,
        )

        return _render(holdings, watchlist, suggestions)


def _render(
    holdings: list[_Scored],
    watchlist: list[_Scored],
    suggestions: list[tuple[DiscoveryCandidate, _Scored]],
) -> str:
    generated_at = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    holdings_table = _table(
        "holdings", [_holdings_row(s) for s in holdings], "No meaningful changes this week."
    )
    watchlist_table = _table(
        "watchlist", [_watchlist_row(s) for s in watchlist], "No meaningful changes this week."
    )
    suggestions_table = _table(
        "suggestions",
        [_suggestion_row(c, s) for c, s in suggestions],
        "No verified candidates this week.",
    )

    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Stock Research Dashboard — {generated_at}</title>
<style>
  :root {{
    --bg: #ffffff; --fg: #1a1a1a; --muted: #666; --border: #eee; --border-strong: #ddd;
    --pill-neutral-bg: #eee; --pill-neutral-fg: #444;
  }}
  @media (prefers-color-scheme: dark) {{
    :root {{
      --bg: #14161a; --fg: #e6e6e6; --muted: #999; --border: #2a2a2a; --border-strong: #333;
      --pill-neutral-bg: #333; --pill-neutral-fg: #ccc;
    }}
  }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    max-width: 960px; margin: 2rem auto; padding: 0 1rem;
    background: var(--bg); color: var(--fg);
  }}
  h1 {{ font-size: 1.4rem; }}
  h2 {{ font-size: 1.1rem; margin-top: 2.5rem; border-bottom: 2px solid var(--border-strong); padding-bottom: 0.3rem; }}
  .generated-at {{ color: var(--muted); font-size: 0.85rem; }}
  table {{ width: 100%; border-collapse: collapse; margin-top: 0.75rem; }}
  th, td {{ text-align: left; padding: 0.5rem 0.6rem; border-bottom: 1px solid var(--border); font-size: 0.9rem; }}
  th {{ color: var(--muted); font-weight: 600; text-transform: uppercase; font-size: 0.75rem; }}
  td.symbol {{ font-weight: 700; }}
  tr.bullish td.symbol {{ color: #1a7f37; }}
  tr.bearish td.symbol {{ color: #cf222e; }}
  tr.neutral td.symbol {{ color: var(--pill-neutral-fg); }}
  .tag {{ display: inline-block; padding: 0.15rem 0.5rem; border-radius: 999px; font-size: 0.75rem; font-weight: 600; }}
  .tag.bullish {{ background: #dafbe1; color: #1a7f37; }}
  .tag.bearish {{ background: #ffebe9; color: #cf222e; }}
  .tag.neutral {{ background: var(--pill-neutral-bg); color: var(--pill-neutral-fg); }}
  .tag.standard {{ background: #ddf4ff; color: #0969da; }}
  .tag.speculative {{ background: #fff8c5; color: #9a6700; }}
  .tag.unknown {{ background: var(--pill-neutral-bg); color: var(--pill-neutral-fg); }}
  .badge {{ display: inline-block; padding: 0.1rem 0.45rem; border-radius: 4px; font-size: 0.72rem; margin: 0 0.2rem 0.2rem 0; }}
  .badge.bullish {{ background: #dafbe1; color: #1a7f37; }}
  .badge.bearish {{ background: #ffebe9; color: #cf222e; }}
  .badge.neutral {{ background: var(--pill-neutral-bg); color: var(--pill-neutral-fg); }}
  .muted {{ color: var(--muted); font-style: italic; }}
  ul.events {{ margin: 0.3rem 0 0; padding-left: 1.1rem; font-size: 0.8rem; }}
  ul.events li {{ margin-bottom: 0.15rem; }}
  ul.events a {{ color: inherit; }}
</style>
</head>
<body>
  <h1>Stock Research Dashboard</h1>
  <div class="generated-at">Generated {generated_at}</div>

  <h2>Your Holdings</h2>
  {holdings_table}

  <h2>Your Watchlist</h2>
  {watchlist_table}

  <h2>New Suggestions</h2>
  {suggestions_table}
</body>
</html>
"""
