"""
scraper.py — URL collection and content extraction pipeline

Provides:
  web_search(query, n, max_scrape, log_fn)        → uses DDG (primary)
  serper_web_search(query, n, max_scrape, log_fn) → uses Serper.dev (swap-in)

Both return [{"url": str, "content": str}, ...].

To switch the agentic pipeline from DDG to Serper.dev, change the one import line
in pipeline.py from `web_search` to `serper_web_search` — nothing else changes.

Raises DDGRateLimitError if DuckDuckGo rate-limits the session so the caller can
abort immediately rather than continuing with empty evidence.
"""

import asyncio
import os
import threading
import time


# ── Custom exception ──────────────────────────────────────────────────────────

class DDGRateLimitError(Exception):
    """Raised when DuckDuckGo actively rate-limits the search session."""
    pass


# ── Concurrency guard ─────────────────────────────────────────────────────────

# Limits concurrent DDG calls across all ThreadPoolExecutor workers
_DDG_SEMAPHORE = threading.Semaphore(2)

_JUNK_DOMAINS = {
    # Social / video
    "youtube.com", "facebook.com", "twitter.com", "x.com",
    "instagram.com", "tiktok.com", "pinterest.com",
    "reddit.com", "quora.com",
    # Support / help / account pages (not useful content)
    "support.google.com", "accounts.google.com", "policies.google.com",
    "play.google.com", "help.twitter.com", "help.instagram.com",
    "support.apple.com", "support.microsoft.com",
}


def _is_junk_domain(url: str) -> bool:
    from urllib.parse import urlparse
    try:
        host = urlparse(url).netloc.lower().removeprefix("www.")
        return any(host == d or host.endswith("." + d) for d in _JUNK_DOMAINS)
    except Exception:
        return False


def _make_logger(log_fn):
    def _log(msg, is_error=False):
        print(msg)
        if log_fn:
            log_fn(msg, is_error)
    return _log


def _is_rate_limit_error(e: Exception) -> bool:
    """Detect DDG rate limit from exception type name or message."""
    name = type(e).__name__.lower()
    msg  = str(e).lower()
    return "ratelimit" in name or "ratelimit" in msg or "202" in msg or "rate limit" in msg


# ── URL collection ────────────────────────────────────────────────────────────

def ddg_search(query: str, n: int = 10, log_fn=None) -> list[str]:
    """
    Return up to n direct destination URLs from DuckDuckGo.

    Raises DDGRateLimitError if DDG rate-limits the session.
    Returns [] on other (non-rate-limit) errors.
    Sleeps 1s after every successful call to reduce rate-limit risk.
    """
    from ddgs import DDGS
    _log = _make_logger(log_fn)
    with _DDG_SEMAPHORE:
        try:
            with DDGS() as ddgs:
                results = list(ddgs.text(query, max_results=n))
            urls = [r["href"] for r in results if "href" in r]
            _log(f"[ddg] '{query}' → {len(urls)} URL(s)")
            time.sleep(1)  # courtesy pause to reduce rate-limit risk
            return urls
        except Exception as e:
            if _is_rate_limit_error(e):
                _log(f"[ddg] Rate limit hit for '{query}': {e}", True)
                raise DDGRateLimitError(f"DDG rate limit: {e}") from e
            _log(f"[ddg] Search failed for '{query}': {e}", True)
            return []


def serper_search(query: str, n: int = 10, log_fn=None) -> list[str]:
    """
    Return up to n direct URLs from Serper.dev Google Search.
    Requires SERPER_API_KEY in .env.
    """
    import requests
    _log = _make_logger(log_fn)
    try:
        resp = requests.post(
            "https://google.serper.dev/search",
            headers={
                "X-API-KEY": os.getenv("SERPER_API_KEY", ""),
                "Content-Type": "application/json",
            },
            json={"q": query, "num": n},
            timeout=10,
        )
        resp.raise_for_status()
        organic = resp.json().get("organic", [])
        urls = [r["link"] for r in organic if "link" in r]
        _log(f"[serper] '{query}' → {len(urls)} URL(s)")
        return urls
    except Exception as e:
        _log(f"[serper] Search failed for '{query}': {e}", True)
        return []


# ── Content extraction ────────────────────────────────────────────────────────

def _extract_date(result) -> str:
    """
    Try to pull a publish date from a crawl4ai CrawlResult.
    Checks result.metadata first, then scans the markdown for ISO/long-form dates.
    Returns a date string like '2024-03-15' or 'unknown'.
    """
    import re as _re

    # crawl4ai may populate result.metadata with a published_date or date key
    meta = getattr(result, "metadata", None) or {}
    for key in ("published_date", "date", "publish_date", "datePublished", "article:published_time"):
        val = meta.get(key)
        if val:
            # Trim to date portion if it's a full ISO datetime
            return str(val)[:10]

    # Fallback: scan the first 3000 chars of markdown for a date pattern
    text = (result.markdown or "")[:3000]

    # ISO: 2024-03-15
    m = _re.search(r'\b(20\d{2}[-/]\d{2}[-/]\d{2})\b', text)
    if m:
        return m.group(1).replace("/", "-")

    # Long-form: March 15, 2024 / 15 March 2024
    months = r'(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)'
    m = _re.search(rf'\b({months}\s+\d{{1,2}},?\s+20\d{{2}})\b', text, _re.IGNORECASE)
    if m:
        return m.group(1)
    m = _re.search(rf'\b(\d{{1,2}}\s+{months}\s+20\d{{2}})\b', text, _re.IGNORECASE)
    if m:
        return m.group(1)

    return "unknown"


async def _crawl_async(urls: list[str], max_scrape: int, log_fn) -> list[dict]:
    """Crawl up to max_scrape non-junk URLs and return evidence dicts."""
    from crawl4ai import AsyncWebCrawler
    _log = _make_logger(log_fn)
    results = []
    async with AsyncWebCrawler() as crawler:
        for url in urls:
            if len(results) >= max_scrape:
                break
            if _is_junk_domain(url):
                _log(f"[crawl] Skipping junk domain: {url}")
                continue
            try:
                result = await asyncio.wait_for(crawler.arun(url=url), timeout=10)
                if result.success and result.markdown:
                    content = " ".join(result.markdown.split())[:2000]
                    timestamp = _extract_date(result)
                    results.append({"url": url, "content": content, "timestamp": timestamp})
                    _log(f"[crawl] Scraped: {url} (date: {timestamp})")
                else:
                    _log(f"[crawl] No content from: {url}")
            except asyncio.TimeoutError:
                _log(f"[crawl] Timed out (20s), skipping: {url}", True)
            except Exception as e:
                _log(f"[crawl] Failed {url}: {e}", True)
    _log(f"[crawl] Completed — {len(results)} page(s) scraped")
    return results


def crawl_urls(urls: list[str], max_scrape: int = 4, log_fn=None) -> list[dict]:
    """Sync wrapper around _crawl_async. Safe to call from ThreadPoolExecutor threads."""
    return asyncio.run(_crawl_async(urls, max_scrape, log_fn))


# ── Combined entry points ─────────────────────────────────────────────────────

def web_search(query: str, n: int = 10, max_scrape: int = 4, log_fn=None) -> list[dict]:
    """
    DDG-backed search + crawl.
    Propagates DDGRateLimitError — callers should catch it and abort.
    """
    urls = ddg_search(query, n=n, log_fn=log_fn)  # raises DDGRateLimitError on rate limit
    return crawl_urls(urls, max_scrape=max_scrape, log_fn=log_fn)


def serper_web_search(query: str, n: int = 10, max_scrape: int = 4, log_fn=None) -> list[dict]:
    """
    Serper.dev-backed search + crawl. Drop-in replacement for web_search.
    To switch the agentic pipeline: change `from scraper import web_search`
    to `from scraper import serper_web_search as web_search` in pipeline.py.
    """
    urls = serper_search(query, n=n, log_fn=log_fn)
    return crawl_urls(urls, max_scrape=max_scrape, log_fn=log_fn)
