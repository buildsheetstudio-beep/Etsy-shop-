"""CLI entrypoint — runs the full weekly pipeline (DESIGN.md section 7)
end to end and writes the dashboard. This is what the weekly cron job
(DESIGN.md section 2) should invoke.

Usage:
    python -m stock_agent.main
"""

from __future__ import annotations

import logging

from . import config
from .pipeline import run_synthesis_stage


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
    run_synthesis_stage()
    print(f"Dashboard written to {config.DASHBOARD_FILE}")


if __name__ == "__main__":
    main()
