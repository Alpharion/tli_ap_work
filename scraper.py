"""
scraper.py — URL collection and content extraction pipeline

Provides web_search(query, n, max_scrape, log_fn) → list[{"url": str, "content": str}]
Used by verify.py to replace the Playwright + Bing scraping approach.

URL collection : DuckDuckGo (ddgs) — primary
                 Serper.dev        — written but NOT wired up; swap into web_search if needed
Content extract: Crawl4AI async crawler
"""

import asyncio
import os
import threading


# Limit concurrent DDG calls across all ThreadPoolExecutor workers to avoid rate limits
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


# ── URL collection ────────────────────────────────────────────────────────────

def ddg_search(query: str, n: int = 10, log_fn=None) -> list[str]:
    """Return up to n direct destination URLs from DuckDuckGo. Returns [] on rate limit or error."""
    from ddgs import DDGS
    _log = _make_logger(log_fn)
    with _DDG_SEMAPHORE:
        try:
            with DDGS() as ddgs:
                results = list(ddgs.text(query, max_results=n))
            urls = [r["href"] for r in results if "href" in r]
            _log(f"[ddg] '{query}' → {len(urls)} URL(s)")
            return urls
        except Exception as e:
            _log(f"[ddg] Search failed for '{query}': {e}", True)
            return []


def serper_search(query: str, n: int = 10, log_fn=None) -> list[str]:
    """
    Return up to n direct URLs from Serper.dev Google Search.
    NOT called by web_search — to switch from DDG, replace ddg_search with this in web_search.
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
                result = await crawler.arun(url=url)
                if result.success and result.markdown:
                    content = " ".join(result.markdown.split())[:2000]
                    results.append({"url": url, "content": content})
                    _log(f"[crawl] Scraped: {url}")
                else:
                    _log(f"[crawl] No content from: {url}")
            except Exception as e:
                _log(f"[crawl] Failed {url}: {e}", True)
    _log(f"[crawl] Completed — {len(results)} page(s) scraped")
    return results


def crawl_urls(urls: list[str], max_scrape: int = 4, log_fn=None) -> list[dict]:
    """Sync wrapper around _crawl_async. Safe to call from ThreadPoolExecutor threads."""
    return asyncio.run(_crawl_async(urls, max_scrape, log_fn))


# ── Combined entry point ──────────────────────────────────────────────────────

def web_search(query: str, n: int = 10, max_scrape: int = 4, log_fn=None) -> list[dict]:
    """
    Collect up to n URLs via DuckDuckGo, crawl up to max_scrape non-junk pages.
    Returns [{"url": str, "content": str}, ...] — identical shape to old playwright_search.
    To switch to Serper.dev: replace ddg_search(...) with serper_search(...) below.
    """
    urls = ddg_search(query, n=n, log_fn=log_fn)
    return crawl_urls(urls, max_scrape=max_scrape, log_fn=log_fn)
