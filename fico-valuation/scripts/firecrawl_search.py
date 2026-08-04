#!/usr/bin/env python3
"""Firecrawl search helper (replaces Brave Search for this project).

Usage:
  export FIRECRAWL_API_KEY=fc-...   # or load from repo-root .env.local
  python3 fico-valuation/scripts/firecrawl_search.py "free DCF Excel template"
"""

from __future__ import annotations

import json
import os
import sys
import urllib.request
from pathlib import Path


def load_key() -> str:
    key = os.environ.get("FIRECRAWL_API_KEY")
    if key:
        return key
    env_path = Path(__file__).resolve().parents[2] / ".env.local"
    if env_path.exists():
        for line in env_path.read_text().splitlines():
            if line.startswith("FIRECRAWL_API_KEY="):
                return line.split("=", 1)[1].strip().strip('"').strip("'")
    raise SystemExit("FIRECRAWL_API_KEY not set (export it or add .env.local)")


def search(query: str, limit: int = 8) -> dict:
    payload = json.dumps({"query": query, "limit": limit}).encode()
    req = urllib.request.Request(
        "https://api.firecrawl.dev/v1/search",
        data=payload,
        headers={
            "Authorization": f"Bearer {load_key()}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read().decode())


def scrape(url: str) -> dict:
    payload = json.dumps({"url": url, "formats": ["markdown"]}).encode()
    req = urllib.request.Request(
        "https://api.firecrawl.dev/v1/scrape",
        data=payload,
        headers={
            "Authorization": f"Bearer {load_key()}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=90) as resp:
        return json.loads(resp.read().decode())


def main(argv: list[str]) -> None:
    if len(argv) < 2:
        raise SystemExit(__doc__)
    if argv[1] == "--scrape":
        if len(argv) < 3:
            raise SystemExit("Usage: firecrawl_search.py --scrape URL")
        print(json.dumps(scrape(argv[2]), indent=2)[:4000])
        return
    query = " ".join(argv[1:])
    res = search(query)
    data = res.get("data") or []
    if isinstance(data, dict):
        data = data.get("web") or []
    for item in data:
        print(f"{item.get('title')}\n  {item.get('url')}\n")


if __name__ == "__main__":
    main(sys.argv)
