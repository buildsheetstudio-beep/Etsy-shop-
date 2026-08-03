"""Shared HTTP plumbing: rate-limited, fail-soft GET/POST requests.

DESIGN.md section 3: free tiers are rate-limited and not guaranteed
long-term stable, so "agents should fail gracefully (skip/retry) rather
than assume 100% uptime." get_json()/post_json() never raise for
network/HTTP errors — they return None and let the caller record the
failure rather than aborting the whole run.
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
        return self._request("GET", url, params=params)

    def post_json(self, url: str, json_body: dict) -> Optional[dict | list]:
        return self._request("POST", url, json=json_body)

    def _request(self, method: str, url: str, **kwargs) -> Optional[dict | list]:
        self._throttle()
        try:
            response = self._session.request(method, url, timeout=self._timeout, **kwargs)
            self._last_call_at = time.monotonic()
            response.raise_for_status()
            return response.json()
        except requests.RequestException as exc:
            self._last_call_at = time.monotonic()
            logger.warning("Request failed for %s: %s", url, exc)
            return None
