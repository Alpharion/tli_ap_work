# Plan: Parallelise Verification + Reuse Browser (verify.py)

## Context
The verification pipeline in `verify.py` processes 146 supplier rows sequentially, and
spins up a new Playwright browser for every one of the 438 search queries. This results
in an estimated 1–2 hour runtime. Two optimisations are planned:
1. **Thread pool** — verify multiple rows concurrently
2. **Browser reuse** — one browser instance shared across all searches in a run

---

## Optimisation 1 — Thread Pool (biggest gain)

**Where:** `annotate_workbook()` in `verify.py` (~line 282)

Currently it loops over rows sequentially:
```python
for row_idx in range(2, ws.max_row + 1):
    ...
    result = verify_supplier_row(...)
```

**Change:** collect all rows first, then submit them to a `ThreadPoolExecutor`. Each
worker calls `verify_supplier_row()` for one row independently.

```python
from concurrent.futures import ThreadPoolExecutor, as_completed

rows_to_process = []
for row_idx in range(2, ws.max_row + 1):
    company = ws.cell(row=row_idx, column=col_company).value
    if company:
        rows_to_process.append((row_idx, company, supplies_to_val, components_val))

with ThreadPoolExecutor(max_workers=5) as executor:
    futures = {
        executor.submit(verify_supplier_row, company, supplies_to, components, provider, log_fn): row_idx
        for row_idx, company, supplies_to, components in rows_to_process
    }
    for future in as_completed(futures):
        row_idx = futures[future]
        result = future.result()
        # write result cells back to ws at row_idx
```

**Worker count:** 5 is a safe default — enough parallelism without hammering Bing or
hitting Gemini rate limits. Can be tuned up to 8 if rate limits allow.

**Thread safety note:** `openpyxl` worksheet writes are NOT thread-safe. The cell writes
must happen on the main thread after each future completes, not inside the worker.
`as_completed()` already handles this — results come back to the main thread one at a time.

---

## Optimisation 2 — Browser Reuse (medium gain)

**Where:** `playwright_search()` (~line 43) and `verify_supplier_row()` (~line 147)

Currently `playwright_search()` opens and closes a full browser on every call:
```python
with sync_playwright() as p:
    browser = p.chromium.launch(...)
    ...
    browser.close()
```

**Change:** accept an optional `page` parameter. The browser and context are created once
per worker thread in `verify_supplier_row()` and reused across its 3 search calls.

```python
# playwright_search signature change:
def playwright_search(query: str, page, n: int = 4, log_fn=None) -> list[dict]:
    # no browser lifecycle here — just use the passed page
    page.goto(search_url, ...)
    ...
```

```python
# verify_supplier_row creates browser once and passes page to all 3 searches:
def verify_supplier_row(company_name, supplies_to, components, provider, log_fn=None):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(user_agent=..., locale="en-US")
        page = context.new_page()

        results1 = playwright_search(query1, page, n=3, log_fn=log_fn)
        results2 = playwright_search(query2, page, n=4, log_fn=log_fn)
        results3 = playwright_search(query3, page, n=4, log_fn=log_fn)

        browser.close()
    # rest of LLM judging unchanged
```

Each thread gets its own browser — Playwright is not thread-safe so browsers must not
be shared across threads, only reused within a single thread's 3 queries.

---

## Files to modify

- `verify.py` only:
  - `playwright_search()` — remove browser lifecycle, accept `page` param
  - `verify_supplier_row()` — own the browser lifecycle, pass `page` to searches
  - `annotate_workbook()` — introduce `ThreadPoolExecutor`, collect rows first, write cells on main thread

---

## Expected improvement

| | Before | After |
|---|---|---|
| Browser launches | 438 | ~146 (1 per row) |
| Concurrent rows | 1 | 5 |
| Estimated runtime | ~2 hours | ~20–25 minutes |

---

## Verification

1. Run `/dev/verify` with a small test workbook (5–10 rows)
2. Confirm SSE log shows multiple rows being processed concurrently (progress messages
   interleaved, not strictly sequential)
3. Confirm output Excel has all 5 verification columns populated correctly
4. Check no cell overwrites or data corruption (thread safety of writes)
