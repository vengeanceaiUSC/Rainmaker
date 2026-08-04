"""Deterministic SEC EDGAR companyfacts (XBRL JSON) client.

No HTML scraping. No LLM extraction.
"""

from __future__ import annotations

import json
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any, Dict, Optional

DEFAULT_UA = "FICOFinancialModel research@example.com"
COMPANY_TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
COMPANYFACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"


class EdgarError(RuntimeError):
    pass


def _get_json(url: str, user_agent: str, timeout: int = 60) -> Dict[str, Any]:
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": user_agent,
            "Accept": "application/json",
            "Accept-Encoding": "gzip, deflate",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read()
            # handle gzip if urllib didn't
            if raw[:2] == b"\x1f\x8b":
                import gzip

                raw = gzip.decompress(raw)
            return json.loads(raw.decode("utf-8"))
    except urllib.error.HTTPError as e:
        raise EdgarError(f"HTTP {e.code} fetching {url}") from e
    except Exception as e:
        raise EdgarError(f"Failed fetching {url}: {e}") from e


def resolve_cik(ticker: str, user_agent: str = DEFAULT_UA, cache_path: Optional[Path] = None) -> str:
    ticker = ticker.upper().strip()
    if cache_path and cache_path.exists():
        data = json.loads(cache_path.read_text())
    else:
        data = _get_json(COMPANY_TICKERS_URL, user_agent)
        if cache_path:
            cache_path.parent.mkdir(parents=True, exist_ok=True)
            cache_path.write_text(json.dumps(data))
        time.sleep(0.2)  # be polite to SEC
    for row in data.values() if isinstance(data, dict) else data:
        if str(row.get("ticker", "")).upper() == ticker:
            return f"{int(row['cik_str']):010d}"
    raise EdgarError(f"Ticker not found in SEC company_tickers.json: {ticker}")


def fetch_companyfacts(
    ticker: str,
    *,
    cik: Optional[str] = None,
    user_agent: str = DEFAULT_UA,
    cache_dir: Optional[Path] = None,
    force_refresh: bool = False,
) -> Dict[str, Any]:
    """Fetch XBRL companyfacts JSON for ticker. Cache to disk when possible."""
    cache_dir = cache_dir or Path(__file__).resolve().parents[1] / "data"
    cache_dir.mkdir(parents=True, exist_ok=True)
    tickers_cache = cache_dir / "company_tickers.json"
    if not cik:
        cik = resolve_cik(ticker, user_agent=user_agent, cache_path=tickers_cache)
    out_path = cache_dir / f"companyfacts_CIK{cik}.json"
    if out_path.exists() and not force_refresh and out_path.stat().st_size > 1000:
        return json.loads(out_path.read_text())
    url = COMPANYFACTS_URL.format(cik=cik)
    data = _get_json(url, user_agent)
    out_path.write_text(json.dumps(data))
    return data
