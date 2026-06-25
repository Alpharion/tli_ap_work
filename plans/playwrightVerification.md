/# Plan: Replace DDG with Playwright + Bing in verify.py

## Context

The current `verify_supplier_row()` in `verify.py` uses `ddg_search_and_scrape()` (imported from `app.py`) for web evidence. DDG rate-limits aggressively when fired in rapid succession (3 queries × N supplier rows), returning empty results every time. The fix is to replace the DDG calls with a Playwright-based headless browser that searches Bing and scrapes the top article pages, giving real content and real URLs.

---

## What changes

**Only `verify.py` is modified** — one new helper function replaces the DDG import. `app.py` is untouched.

---

## New helper: `playwright_search(query, n=4) -> list[dict]`

Lives in `verify.py`, replaces the `ddg_search_and_scrape` calls.

**Returns:** list of `{"url": str, "content": str}` dicts — one per scraped article.

**Flow:**
1. Launch Playwright Chromium in headless mode (`sync_playwright`)
2. Open a new page with a realistic user-agent string
3. Navigate to `https://www.bing.com/search?q={encoded_query}`
4. Wait for search results to render (`#b_results` selector)
5. Extract top `n` result URLs from `<h2><a>` links inside `#b_results`, skipping Bing internal links (ads, `bing.com/`, `microsoft.com/`)
6. For each URL, open a new page, navigate, wait for `body`, extract plain text via `page.inner_text("body")` (Playwright handles JS-rendered pages natively), truncate to 2000 chars
7. Close browser
8. Return list of `{"url": url, "content": content}` dicts

**Error handling:** Any per-page failure is caught and skipped. If the whole search fails, return `[]` — callers fall back to "No evidence found".

---

## Changes to `verify_supplier_row()`

Replace the three `ddg_search_and_scrape` calls and the DDG import with `playwright_search`. The evidence string passed to `_ai_judge` is rebuilt from the results list:

```python
results = playwright_search(query, n=4)
evidence = "\n---\n".join(
    f"SOURCE: {r['url']}\nCONTENT: {r['content']}" for r in results
)
```

URLs are taken directly from `results` (real, visited URLs) rather than relying on the LLM to echo them back. After `_ai_judge` returns `answer=True`, attach all `r["url"]` values from that step's results to `all_urls`.

---

## Prompt update in `_ai_judge`

Remove the `"urls": [...]` field from the JSON schema since URLs now come from the search results directly. Simplify to:

```json
{
  "answer": true or false,
  "notes": "one sentence explanation"
}
```

---

## Dependency

Add to `requirements.txt`:
```
playwright
```

After install, user must run once:
```bash
playwright install chromium
```

---

## Files modified

- **`verify.py`** — add `playwright_search()`, update `verify_supplier_row()`, update `_ai_judge()` prompt
- **`requirements.txt`** — add `playwright`

---

## Verification

1. Install: `pip install playwright && playwright install chromium`
2. Start server, navigate to `/dev/verify`
3. Upload an agentic export Excel (depth=1 recommended for quick test)
4. Watch SSE log — should see per-row progress with Bing searches firing
5. Download annotated Excel — confirm URLs column contains real article URLs, and Yes/No fields are populated
