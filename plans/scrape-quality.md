# Plan: Improve Scrape Quality with Selector Fallback (verify.py)

## Context
`playwright_search()` currently grabs `page.inner_text("body")`, which includes navbars,
footers, sidebars, and cookie banners. The 2000-char cap then takes whatever appears first
in the DOM — usually the header/nav, not the article. This means the LLM is judging
suppliers based on low-quality content. The fix is to try progressively less specific
semantic selectors before falling back to `body`.

---

## Change — `playwright_search()` in `verify.py`

**Current code (inside the scrape loop):**
```python
text = page.inner_text("body")
text = " ".join(text.split())[:2000]
```

**Replacement:**
```python
CONTENT_SELECTORS = ["article", "main", "[role='main']", "#content", "#main-content", "body"]
text = ""
for selector in CONTENT_SELECTORS:
    el = page.query_selector(selector)
    if el:
        text = el.inner_text()
        break
text = " ".join(text.split())[:2000]
```

Selectors are tried in order — most specific first, `body` as the final safety net so
the function never returns empty content.

---

## File to modify

- `verify.py` — inside the article scrape loop in `playwright_search()`, the two lines
  that extract `text` from the page (~line 98–99 currently)

---

## Verification

1. Run `/dev/verify` on a small batch and check the SSE log — scraped content in
   `[playwright] Scraped: <url>` entries should reflect article body text, not nav/footer noise
2. Spot-check the "Verification Notes" column in the output Excel — notes should cite
   specific evidence (company names, product lines) rather than generic boilerplate
