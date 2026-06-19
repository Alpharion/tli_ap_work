# Plan: Training Data Confirmation Loop (verify.py)

## Context
Currently, when the LLM answers "Yes" based on training knowledge (no web evidence found),
the label column shows "Training Data". The goal is to add a confirmation loop: whenever a
field is "Yes" + "Training Data", run a targeted web search specifically for that field,
scrape up to 3 non-junk pages, re-ask the LLM for that one field only, and if confirmed
from web evidence — upgrade the label to "Web Evidence" and replace the URLs.

---

## Flow

### Phase 1 (unchanged): Initial 3-query sweep
3 Bing searches (n=10 each) → combined evidence → 1 LLM call → initial verdict + labels

### Phase 2 (new): Per-field confirmation loop
```
For each field in [company_exists, supply_ties, correct_component]:
  if verdict == "Yes" AND note contains "based on training knowledge":
      → run targeted Bing search (n=10, stop after 3 non-junk pages scraped)
      → build evidence from those 3 pages
      → ask LLM for this 1 field only
      if LLM confirms "Yes":
          → update note (remove training prefix — now web-backed)
          → update label to "Web Evidence"
          → replace URLs with the 3 confirmation page URLs
      else:
          → keep original "Yes (Training Data)" answer unchanged
```

---

## Code Changes (verify.py only)

### 1. Modify `playwright_search` — add `max_scrape` parameter
Current signature: `playwright_search(query, page, n=4, log_fn=None)`
New signature: `playwright_search(query, page, n=4, log_fn=None, max_scrape=None)`

Inside the scraping loop, track a counter and break once `max_scrape` non-junk pages
are successfully scraped. `None` (default) preserves current behaviour — no cap.

```python
scraped_count = 0
for href in hrefs:
    if max_scrape and scraped_count >= max_scrape:
        break
    # ... existing nav + junk check + content extraction ...
    results.append({"url": page.url, "content": text})
    scraped_count += 1
```

### 2. New `_ai_judge_single` function
Add after `_ai_judge_all`. Asks the LLM ONE specific field question using ONLY the new
web evidence scraped in Phase 2.

```python
def _ai_judge_single(field: str, evidence: str, company_name: str,
                     supplies_to: str, components: str, provider: str, log_fn=None) -> dict:
    # Returns {"answer": bool, "note": str}
```

Prompt per field (evidence-only, no training knowledge fallback):
- `company_exists`    → "Does the evidence confirm {company_name} exists as a real company?"
- `supply_ties`       → "Does the evidence confirm a supply/collaboration between {company_name} and {supplies_to}?"
- `correct_component` → "Does the evidence confirm {company_name} manufactures or produces {components}?"

Returns JSON: `{"answer": true/false, "note": "one sentence citing the web evidence"}`
The note is purely web-evidence-driven — no "training knowledge" prefix.
If `answer` is false, the caller keeps the original training data answer unchanged.

### 3. New `_confirm_with_web` function
Add after `_ai_judge_single`. Orchestrates one confirmation attempt for a single field.

```python
def _confirm_with_web(field, training_note, company_name, supplies_to,
                      components, provider, page, log_fn=None):
    # Returns {"confirmed": bool, "note": str, "urls": [str]} or None on failure
```

**Search query = training data note, stripped of its prefix.**

The training note looks like:
`"Company Exists is based on training knowledge — Umicore is a Belgian materials company specialising in battery cathode materials"`

Strip everything up to and including `" — "` to get the raw claim:
`"Umicore is a Belgian materials company specialising in battery cathode materials"`

This becomes the Bing search query directly — far more targeted than a generic template.

```python
query = training_note.split(" — ", 1)[-1] if " — " in training_note else training_note
```

Calls `playwright_search(query, page, n=10, max_scrape=3, log_fn=log_fn)`.
Builds evidence string from those ≤3 non-junk pages, calls `_ai_judge_single`.

### 4. Modify `verify_supplier_row`
After the initial `verdict = _ai_judge_all(...)` block and BEFORE `browser.close()`,
add the confirmation loop (browser must stay open since `_confirm_with_web` reuses `page`):

```python
FIELD_MAP = [
    ("company_exists",    "notes_exists"),
    ("supply_ties",       "notes_supply"),
    ("correct_component", "notes_component"),
]
confirmation_urls = {}   # field -> [url, ...] for confirmed fields only

for field, notes_key in FIELD_MAP:
    is_yes      = verdict[field]
    is_training = "based on training knowledge" in verdict[notes_key]
    if is_yes and is_training:
        conf = _confirm_with_web(field, verdict[notes_key], company_name, supplies_to,
                                  component_query, provider, page, log_fn)
        if conf and conf["confirmed"]:
            verdict[notes_key] = conf["note"]        # web-backed note, no training prefix
            confirmation_urls[field] = conf["urls"]  # replace URLs for this field
            _log(f"[confirm] {field} upgraded to Web Evidence")
        else:
            _log(f"[confirm] {field} could not be confirmed via web — keeping Training Data")

browser.close()   # ← moved to after confirmation loop
```

When building `unique_urls` for the return dict, confirmation URLs are appended first for
confirmed fields, then remaining Phase 1 URLs (deduped).

---

## Reused Code
| Function | How reused |
|---|---|
| `playwright_search` | Called in `_confirm_with_web` with `max_scrape=3` |
| `_is_junk_domain` | Already called inside `playwright_search` — no change needed |
| Same `page` object | Passed from `verify_supplier_row` into confirmation loop |
| `_ai_judge_all` | Phase 1 unchanged |

---

## Cost & Time Analysis

### Cost (Gemini Flash 2.0: ~$0.10/1M input, ~$0.40/1M output tokens)
Confirmation prompts are smaller than the initial call (~500 input tokens vs ~2000).

| Scenario | LLM calls/row | Est. cost/row | 146 rows total |
|---|---|---|---|
| Current | 1 | ~$0.0003 | ~$0.04 |
| New — best case (0 fields need confirmation) | 1 | ~$0.0003 | ~$0.04 |
| New — typical (~1.5 fields need confirmation) | 2.5 | ~$0.0006 | ~$0.09 |
| New — worst case (all 3 fields need confirmation) | 4 | ~$0.0009 | ~$0.13 |

### Time (with 5 workers)
Each confirmation adds ~30s (1 Bing search capped at 3 scrapes + 1 LLM call).

| Scenario | Time/row | 146 rows (5 workers) |
|---|---|---|
| Current | ~3 min | ~90 min |
| New — best case | ~3 min | ~90 min |
| New — typical (~1.5 confirmations/row) | ~3.8 min | ~110 min |
| New — worst case (3 confirmations/row) | ~4.5 min | ~135 min |

`company_exists` is the field most likely to be confirmed quickly (well-known companies),
so typical case skews towards 0–1 confirmation per row rather than 3.

---

## Verification
1. Run `/dev/verify` on a small test file (5–10 rows with known companies)
2. In the frontend log, look for `[confirm]` lines showing which fields triggered confirmation
3. In output Excel:
   - Confirmed fields show "Web Evidence" label + new URLs
   - Unconfirmed fields keep "Training Data" label + original answer intact
   - Fields answered from web evidence in Phase 1 are untouched (no `[confirm]` lines logged)
