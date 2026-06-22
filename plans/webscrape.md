# Plan: Replace Playwright with DDGS + Crawl4AI scraping pipeline

## Context
`verify.py` currently uses Playwright + playwright-stealth to headlessly drive a Bing browser
session for URL collection, then manually scrapes each result page. This has two problems:
1. **Bing redirect URLs** — Bing's anchor tags point to `bing.com/ck/a?...` redirects that
   behave unreliably in headless mode, causing irrelevant pages to be scraped.
2. **Speed** — browser launch per row, bot-detection sleeps, and manual page navigation.

The new approach:
- **`duckduckgo_search` (ddgs)** replaces Playwright for URL collection — already installed,
  returns actual destination URLs directly (no redirect chain), free with no API key.
- **Crawl4AI** replaces the manual Playwright page-scraping loop — handles JS rendering and
  returns clean markdown content per URL.
- A **Serper.dev function is written but not wired up** — ready to swap in if DDG rate limits
  become a problem in production.
- All scraping logic moves to a **new file `scraper.py`**, keeping `verify.py` focused on
  verification logic only.

---

## New File: `scraper.py`

### Structure

```
scraper.py
├── _JUNK_DOMAINS               # moved from verify.py
├── _is_junk_domain(url)        # moved from verify.py
├── ddg_search(query, n, log_fn)       # PRIMARY — called by web_search
├── serper_search(query, n, log_fn)    # WRITTEN but NOT called — future swap-in
├── crawl_urls(urls, max_scrape, log_fn)
└── web_search(query, n, max_scrape, log_fn)   # called by verify.py
```

### `ddg_search` (primary URL source)

```python
def ddg_search(query: str, n: int = 10, log_fn=None) -> list[str]:
    """Return up to n direct URLs from DuckDuckGo. Falls back to [] on rate limit."""
    from duckduckgo_search import DDGS
    _log = _make_logger(log_fn)
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(query, max_results=n))
        urls = [r["href"] for r in results if "href" in r]
        _log(f"[ddg] '{query}' → {len(urls)} URL(s)")
        return urls
    except Exception as e:
        _log(f"[ddg] Search failed for '{query}': {e}", True)
        return []
```

Rate limit behaviour: if DDG blocks the request, the exception is caught and an empty list
is returned. `verify_supplier_row` degrades gracefully — `_ai_judge_all` receives
`"No search results returned."` and the LLM falls back to training knowledge.

### `serper_search` (written, not called)

```python
def serper_search(query: str, n: int = 10, log_fn=None) -> list[str]:
    """
    Return up to n direct URLs from Serper.dev Google Search.
    NOT called by web_search — swap in manually if DDG rate limits become a problem.
    Requires SERPER_API_KEY in .env.
    """
    import requests
    _log = _make_logger(log_fn)
    try:
        resp = requests.post(
            "https://google.serper.dev/search",
            headers={"X-API-KEY": os.getenv("SERPER_API_KEY", ""), "Content-Type": "application/json"},
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
```

To switch from DDG to Serper: change the one line in `web_search` from
`urls = ddg_search(...)` to `urls = serper_search(...)`.

### `crawl_urls`

```python
async def _crawl_async(urls: list[str], max_scrape: int, log_fn) -> list[dict]:
    """Async inner — crawl up to max_scrape non-junk URLs, return evidence dicts."""
    from crawl4ai import AsyncWebCrawler
    results = []
    async with AsyncWebCrawler() as crawler:
        for url in urls:
            if len(results) >= max_scrape:
                break
            if _is_junk_domain(url):
                continue
            try:
                result = await crawler.arun(url=url)
                if result.success and result.markdown:
                    content = " ".join(result.markdown.split())[:2000]
                    results.append({"url": url, "content": content})
            except Exception as e:
                _log(f"[crawl] Failed {url}: {e}", True)
    return results

def crawl_urls(urls: list[str], max_scrape: int = 4, log_fn=None) -> list[dict]:
    """Sync wrapper — safe to call from ThreadPoolExecutor threads via asyncio.run()."""
    return asyncio.run(_crawl_async(urls, max_scrape, log_fn))
```

### `web_search` (entry point called by verify.py)

```python
def web_search(query: str, n: int = 10, max_scrape: int = 4, log_fn=None) -> list[dict]:
    """Collect URLs via DDG, crawl up to max_scrape of them, return evidence dicts."""
    urls = ddg_search(query, n=n, log_fn=log_fn)
    return crawl_urls(urls, max_scrape=max_scrape, log_fn=log_fn)
```

Return shape: `[{"url": str, "content": str}, ...]` — identical to old `playwright_search`,
so `_build_evidence` and `_unique_urls` in `verify.py` require zero changes.

---

## Changes to `verify.py`

### Remove entirely
- `from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError`
- `from playwright_stealth import Stealth`
- `_JUNK_DOMAINS` set and `_is_junk_domain()` (lines ~36–48) — moved to `scraper.py`
- `playwright_search()` function (lines ~100–181) — replaced by `web_search` from `scraper.py`
- `import time` and `import random` — sleeps no longer needed

### Add at top of verify.py
```python
from scraper import web_search
```

### Update `verify_supplier_row` — replace the browser block

```python
# REMOVE this entire block:
try:
    with Stealth().use_sync(sync_playwright()) as p:
        browser = p.chromium.launch(headless=True)
        ...
        r1 = playwright_search(..., page, n=10, log_fn=log_fn, max_scrape=4)
        time.sleep(random.uniform(1, 3))
        r2 = playwright_search(...)
        time.sleep(random.uniform(1, 3))
        r3 = playwright_search(...)
        all_results = r1 + r2 + r3
        verdict = _ai_judge_all(...)
        browser.close()
except Exception as e:
    _log(f"[playwright] Browser error ...")

# REPLACE with:
r1 = web_search(f'"{company_name}" (supplier OR manufacturer)', n=10, max_scrape=4, log_fn=log_fn)
r2 = web_search(f'"{company_name}" "{supplies_to}" (supply OR partnership OR collaboration OR venture)', n=10, max_scrape=4, log_fn=log_fn)
r3 = web_search(f'"{company_name}" "{component_query}" (produce OR manufacture OR supply)', n=10, max_scrape=4, log_fn=log_fn)
all_results = r1 + r2 + r3
verdict = _ai_judge_all(_build_evidence(all_results), company_name, supplies_to, component_query, provider, log_fn=log_fn)
```

The `try/except` around the browser block is removed — `web_search` already handles its own
exceptions internally and returns `[]` on failure, so the outer try is not needed.

---

## `requirements.txt` changes

Add:
```
crawl4ai
```

Remove:
```
playwright>=1.40.0
playwright-stealth>=2.0.0
```

Note: `duckduckgo_search` is already present. `requests` is already present (used by serper_search).

---

## `.env` change (optional, for future Serper.dev swap)
```
SERPER_API_KEY=...   # leave blank for now — only needed if switching to serper_search
```

---

## Reused from existing code
| Existing item | Action |
|---|---|
| `_make_logger` (verify.py ~line 52) | Stays in `verify.py`; also used inside `scraper.py` — import it or duplicate the one-liner |
| `_build_evidence` (verify.py ~line 61) | Stays in `verify.py`, unchanged |
| `_unique_urls` (verify.py ~line 82) | Stays in `verify.py`, unchanged |
| `_ai_judge_all` (verify.py ~line 186) | Stays in `verify.py`, unchanged |
| `duckduckgo_search` (already in requirements.txt) | Used directly in `scraper.py` |

---

## Verification
1. `pip install crawl4ai`
2. Run `python -c "from scraper import web_search; import json; print(json.dumps(web_search('TSMC Taiwan supplier', n=5, max_scrape=2), indent=2))"` — confirm evidence dicts with real URLs and content
3. Upload a small test Excel (5–10 rows) to `/dev/verify` and check:
   - Logs show `[ddg]` and `[crawl]` prefixes (no `[playwright]` lines)
   - URLs in output Excel are direct company/news pages, not Bing redirect URLs
   - Excel label and notes columns populate as before
