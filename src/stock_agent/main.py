"""CLI entrypoint. Currently runs only the Data Agent stage of the pipeline.

Usage:
    python -m stock_agent.main
"""

from __future__ import annotations

import logging

from .pipeline import run_data_stage


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
    records = run_data_stage()
    for record in records:
        status = "ok" if not record.errors else f"errors: {record.errors}"
        print(f"{record.symbol:6s} [{record.type:10s}] price={record.price} {status}")


if __name__ == "__main__":
    main()
