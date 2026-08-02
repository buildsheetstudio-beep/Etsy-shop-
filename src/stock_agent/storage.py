"""Reading the ticker list and generic JSON snapshot I/O.

Snapshot read/diff is the Trend Agent's job (DESIGN.md 6.4); these helpers
are shared plumbing so every agent writes/reads run artifacts the same way.
"""

from __future__ import annotations

import json
from pathlib import Path

from .models import Ticker


def load_tickers(path: Path) -> list[Ticker]:
    if not path.exists():
        raise FileNotFoundError(
            f"Ticker list not found at {path}. Copy config/tickers.example.json "
            "to config/tickers.json and edit it (see DESIGN.md section 4)."
        )
    raw = json.loads(path.read_text())
    return [Ticker(symbol=item["symbol"].upper(), type=item["type"]) for item in raw]


def load_json(path: Path) -> dict:
    if not path.exists():
        return {}
    return json.loads(path.read_text())


def save_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, default=str))
