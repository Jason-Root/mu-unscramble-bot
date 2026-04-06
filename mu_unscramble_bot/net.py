from __future__ import annotations

from functools import lru_cache
import ssl
from typing import Any
import urllib.request


@lru_cache(maxsize=1)
def _https_context() -> ssl.SSLContext:
    try:
        import certifi

        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


def urlopen(request: Any, *, timeout: float):
    url = getattr(request, "full_url", None) or getattr(request, "get_full_url", lambda: "")()
    if isinstance(url, str) and url.lower().startswith("https://"):
        return urllib.request.urlopen(request, timeout=timeout, context=_https_context())
    return urllib.request.urlopen(request, timeout=timeout)
