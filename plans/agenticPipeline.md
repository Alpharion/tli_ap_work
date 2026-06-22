# Plan: Agentic Pipeline (`run_agentic_pipeline`)

## Overview

Instead of a single provider handling all AI calls, the agentic pipeline splits the work
by geography at every level — Gemini handles non-China companies, DeepSeek handles China
companies. Both run in parallel per step, then results are deduplicated before moving on.
This exploits DeepSeek's stronger Chinese-language training data while keeping Gemini for
everything else, without breaking any downstream functionality (export, verification, frontend).

---

## Key Design Decisions

| Decision | Rationale |
|---|---|
| DDG for all evidence gathering | One `web_search` call shared by both LLM prompts — halves DDG calls vs. calling `get_evidence` per provider |
| `DDGRateLimitError` kills pipeline immediately | Prevents sending empty evidence to LLMs and incurring cost for useless calls |
| Market share block skipped | Empty string values set on each OEM entry — too costly to run for testing |
| Gemini wins on dedup | Gemini has stronger general knowledge; its version of a company entry is preferred |
| Parallel calls per step | `ThreadPoolExecutor(max_workers=2)` for Gemini + DeepSeek at each OEM/tier step |
| `serper_web_search` as swap-in | One import alias change in `pipeline.py` switches evidence source from DDG to Serper.dev |

---

## Pipeline Structure

### Step 1 — Product Identification
- Provider: **Gemini only** (strongest general knowledge, Google Search grounding available)
- Evidence: `web_search(f"{product_input} product specification manufacturer supply chain")`
- If `DDGRateLimitError` → push error event, return immediately
- Output: `product_info` dict (same schema as `run_pipeline`)

### Step 2 — OEM Discovery
Two LLM calls fire **in parallel** using `ThreadPoolExecutor(max_workers=2)`:

| Thread | Provider | Prompt constraint |
|---|---|---|
| A | Gemini | "only include companies headquartered or primarily operating **outside China** (excluding Hong Kong, Taiwan)" |
| B | DeepSeek | "only include companies headquartered or primarily operating **in China** (including Hong Kong, Taiwan)" |

Evidence: one shared `web_search` call before spawning threads — both prompts receive the same raw text.

After both return → `_merge_dedup(gemini_results, deepseek_results)`:
- Normalise each name: `.lower().strip()`, strip common suffixes (Inc, Ltd, LLC, GmbH, Corp, Group, Holdings, Co)
- Build a set from Gemini names
- Append DeepSeek entries whose normalised name is NOT already in that set
- Result: unified OEM list, Gemini's version wins on any collision

### Step 3 — Market Share
**Skipped entirely.** Each OEM entry gets:
```python
oem_entry["market_share_pct"]    = ""
oem_entry["market_share_source"] = ""
oem_entry["market_rank"]         = 99
```
The `oem_list.sort(key=lambda x: x.get("market_rank", 99))` line is removed for this pipeline.

### Step 4 — Tier Discovery (repeated for each depth level)
For each parent company in `previous_parents`, a shared `web_search` call fetches evidence,
then two LLM calls fire in parallel:

| Thread | Provider | Prompt constraint |
|---|---|---|
| A | Gemini | "focus on suppliers headquartered or operating **outside China**" |
| B | DeepSeek | "focus on **Chinese suppliers** of {parent}" |

`_merge_dedup` applied to each parent's results before they are added to `tier_suppliers`.

Deduplication across the full tier (same company supplying multiple parents) uses the existing
key: `f"{company_name.lower()}|{oem_root}"` — unchanged from `run_pipeline`.

`time.sleep(1)` is already baked into `ddg_search` in `scraper.py`, so no extra sleep needed here.

### Step 5 — Executive Summary
- Provider: **Gemini only**
- Same prompt as `run_pipeline`

---

## Deduplication Helper

```python
def _merge_dedup(gemini_results: list, deepseek_results: list) -> list:
    _STRIP_SUFFIXES = {
        "inc", "ltd", "llc", "gmbh", "corp", "group", "holdings", "co", "corporation", "limited"
    }

    def normalise(name: str) -> str:
        parts = name.lower().strip().split()
        parts = [p.rstrip(".,") for p in parts if p.rstrip(".,") not in _STRIP_SUFFIXES]
        return " ".join(parts)

    seen = {normalise(e.get("company_name", "")) for e in gemini_results}
    merged = list(gemini_results)
    for entry in deepseek_results:
        if normalise(entry.get("company_name", "")) not in seen:
            merged.append(entry)
    return merged
```

---

## Rate Limit Handling

`ddg_search` in `scraper.py` raises `DDGRateLimitError` on any rate-limit response.
`web_search` propagates it. In `run_agentic_pipeline`:

```python
from scraper import web_search, DDGRateLimitError

try:
    evidence_results = web_search(query)
except DDGRateLimitError as e:
    push("error", message=f"🚫 DDG rate limit hit — aborting pipeline to avoid empty LLM calls. ({e})")
    return  # no LLM calls happen after this point
```

This check wraps every `web_search` call in the pipeline (product ID, OEM, each tier per parent).

---

## Switching Evidence Source to Serper.dev

One line change in `pipeline.py`:
```python
# Current (DDG)
from scraper import web_search, DDGRateLimitError

# Switch to Serper.dev
from scraper import serper_web_search as web_search, DDGRateLimitError
```
`serper_web_search` has the same signature and return shape. `DDGRateLimitError` will simply
never be raised when using Serper (Serper errors are logged and return `[]`).

---

## Output Compatibility

`supply_chain` dict structure is **identical** to `run_pipeline`:
```python
{
  "product": {...},
  "oems":    [...],
  "tiers":   {"tier_1": [...], "tier_2": [...]},
  "summary": "...",
  "provider": "agentic"   # only change — was "gemini" / "deepseek" / etc.
}
```

| Downstream | Impact |
|---|---|
| `build_bulk_workbook` (excel_export.py) | Reads `oems` and `tiers` — no change |
| `verify.py` | Reads `company_name`, `supplies_to`, `components_supplied` per row — no change |
| Frontend SSE events | Same event types and payloads — no change |
| `/api/map` route | Unchanged — still calls `run_pipeline`. New route `/api/map_agentic` calls `run_agentic_pipeline` |

---

## New Route in `app.py`

```python
@app.route("/api/map_agentic")
def map_agentic():
    product = request.args.get("product", "").strip()
    depth   = max(0, min(int(request.args.get("depth", 2)), 3))
    # provider param is ignored — agentic mode always uses Gemini + DeepSeek
    ...
    threading.Thread(target=lambda: run_agentic_pipeline(product, depth, queue, [], {}), daemon=True).start()
    ...
```

The frontend can call `/api/map_agentic` the same way it calls `/api/map` — same SSE event
stream, same rendering logic.

---

## Files Changed

| File | Change |
|---|---|
| `pipeline.py` | Add `run_agentic_pipeline`, `_merge_dedup` |
| `app.py` | Add `/api/map_agentic` route |
| `scraper.py` | Already updated — `DDGRateLimitError`, `time.sleep(1)`, `serper_web_search` |
