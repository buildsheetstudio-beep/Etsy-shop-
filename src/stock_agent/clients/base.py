"""Shared HTTP plumbing: rate-limited, fail-soft GET requests.

DESIGN.md section 3: free tiers are rate-limited and not guaranteed
long-term stable, so "agents should fail gracefully (skip/retry) rather
than assume 100% uptime." get_json() never raises for network/HTTP
errors — it returns None and lets the caller record the failure on the
TickerRecord instead of aborting the whole run.
"""

from __future__ import annotations

import logging
import time
from typing import Optional

import requests

logger = logging.getLogger(__name__)


class RateLimitedClient:
    def __init__(self, min_seconds_between_calls: float, timeout: float = 10.0):
        self._min_interval = min_seconds_between_calls
        self._timeout = timeout
        self._last_call_at: Optional[float] = None
        self._session = requests.Session()

    def _throttle(self) -> None:
        if self._last_call_at is None:
            return
        elapsed = time.monotonic() - self._last_call_at
        remaining = self._min_interval - elapsed
        if remaining > 0:
            time.sleep(remaining)

    def get_json(self, url: str, params: dict) -> Optional[dict | list]:
        self._throttle()
        try:
            response = self._session.get(url, params=params, timeout=self._timeout)
            self._last_call_at = time.monotonic()
            response.raise_for_status()
            return response.json()
        except requests.RequestException as exc:
            self._last_call_at = time.monotonic()
            logger.warning("Request failed for %s: %s", url, exc)
            return None
