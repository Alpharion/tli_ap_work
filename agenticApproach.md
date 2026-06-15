# Plan: `run_agentic_pipeline` — multi-model agentic supply-chain mapping

## Context

The existing `run_pipeline` (`app.py:421-807`) uses **one** LLM provider for every step:
product identification, OEM discovery, and each tier of supplier discovery. Testing
revealed that the three providers have complementary strengths:

- **Anthropic** — best at *inferring* supply ties between companies and the components involved.
- **Gemini** — best at pulling *source hints / URLs* that back a supply tie (it has Google Search grounding built in).
- **DeepSeek** — best at surfacing *China-based* companies.

The goal is a new, separate function `run_agentic_pipeline` that orchestrates these three
models as a "team of specialists" per tier, so each model does what it is best at:
Anthropic finds suppliers (explicitly **not** China-based), DeepSeek adds a few China-based
companies for geographic diversity, and Gemini enriches each tie with a source URL. A
validation layer then merges duplicate companies that the models named differently (e.g.
"TSMC" vs "Taiwan Semiconductor"), using Gemini for the fuzzy semantic match rather than
brittle string comparison.

This is **additive** — `run_pipeline` stays untouched. We work on the `agentic` branch
(already checked out).

## Design decisions (confirmed with user)

| Decision | Choice |
|---|---|
| Gemini source-hint granularity | **One Gemini call per parent** (enriches that parent's whole supplier batch) |
| DeepSeek China-company scope | **OEM discovery + every tier** |
| Trigger | **Dev endpoint only** — new SSE route, no provider-dropdown change yet |
| Dedup/validation | **Gemini, per tier** (one call); same helper also dedups the OEM list |

## Role assignment

| Stage | Anthropic | DeepSeek | Gemini |
|---|---|---|---|
| Product ID | ✅ identify (DDG evidence) | — | — |
| OEM discovery | ✅ non-China OEMs | ✅ China OEMs | dedup/merge OEM list |
| Each tier, per parent | ✅ non-China suppliers + components | ✅ a few China suppliers | enrich batch w/ source URLs |
| Each tier, after all parents | — | — | dedup/merge whole tier |
| Executive summary | ✅ | — | — |

## Event-contract constraint (must preserve)

The frontend (`templates/index.html:1126-1189`) and the export routes depend on exact event
shapes. `run_agentic_pipeline` MUST emit the same events as `run_pipeline`:
`product_identified {product}`, `oem_discovered {oems}`, `tier_complete {tier, suppliers}`,
`complete {supply_chain}`, plus `status {message}`. Supplier dicts must keep keys:
`company_name, country, supplies_to, oem_root, components_supplied, confidence, source_hint, lat, lng`.
OEM dicts: `company_name, country, role, market_share, confidence, notes, lat, lng`.
Set `supply_chain["provider"] = "agentic"`.

## New code

All new code lives in `app.py`. Three small new helpers + the main function + one dev route.

### Helper 1 — `gemini_enrich_suppliers(parent, product_name, suppliers, api_key) -> list`
One Gemini grounding call (reuse the `google.genai` + `GoogleSearch` pattern from
`gemini_search_and_answer`, `app.py:286-319`). Input: the combined supplier list for a
single parent. Ask Gemini to return, for each supplier, a `source_hint` (short citation)
and a backing URL, leaving the rest of the record intact. Parse with `safe_parse_json`
(`app.py:194`). On any failure, return the input list unchanged (graceful degradation).

### Helper 2 — `gemini_dedupe_companies(records, api_key) -> list`
One Gemini call. Send the list of `company_name` (+ `country`) values and ask Gemini to
cluster duplicates / pick a canonical name. **Gemini only decides which names are the same**;
**Python does the field merge deterministically**: for each cluster keep one record, union
`components_supplied`, keep the highest `confidence` (high>medium>low), and keep any non-empty
`source_hint`/url. On failure, fall back to the existing exact-string dedup so the pipeline
never breaks. Used for both the OEM list and each tier's combined list.

### Helper 3 (optional, for readability) — prompt builders
Small functions returning the Anthropic "exclude China" supplier prompt and the DeepSeek
"China-only" supplier prompt, so the main loop stays readable. These reuse the existing
tier/OEM prompt text from `run_pipeline` with one added rule:
- Anthropic prompt adds: *"Do NOT include China-based companies — a separate specialist covers those."*
- DeepSeek prompt adds: *"List ONLY China-based companies (mainland China / Hong Kong)."*

### Main — `def run_agentic_pipeline(product_input, depth, queue, oem_context, collected_tiers)`
Mirrors `run_pipeline`'s structure and its `push()` helper, but with the role split:

1. **Product ID** — `call_ai(id_prompt, "anthropic")` with `get_evidence(query, "anthropic")`
   (DDG). Geocode via `get_coords`. `push("product_identified", ...)`.
2. **OEM discovery** —
   - Anthropic: existing `oem_prompt` + "exclude China" rule → non-China OEMs.
   - DeepSeek: China-only OEM prompt → China OEMs.
   - Combine → `gemini_dedupe_companies` → geocode → ensure primary OEM present
     (reuse logic at `app.py:599-615`) → `push("oem_discovered", ...)`.
3. **Tier loop** `for tier_num in range(1, depth+1)` — reuse the 4-per-parent cap
   (`app.py:644-653`) and `previous_parents`/`next_parents` threading:
   - For each parent: Anthropic suppliers (exclude China) + DeepSeek China suppliers →
     combine → `gemini_enrich_suppliers(parent, ...)` (one call) → extend tier list;
     register survivors as `next_parents`.
   - After all parents: `gemini_dedupe_companies` on the whole tier → geocode survivors →
     update `collected_tiers` → `push("tier_complete", tier=tier_num, suppliers=...)`.
   - Break early if a tier yields nothing (reuse `app.py:782-785`).
4. **Summary** — `call_ai(summary_prompt, "anthropic")`. Set `supply_chain["provider"]="agentic"`.
   `push("complete", supply_chain=...)`.

### Dev route — `GET /api/map_agentic`
Copy the structure of `/api/map` (`app.py:1560-1584+`): read `product` + `depth`
(no `provider` arg), spawn the background thread calling `run_agentic_pipeline`, stream the
queue as SSE via the existing `sse()` helper. Not wired to the dropdown — tested by hitting
the URL directly.

## Reuse (do not rewrite)
- `call_ai` (`app.py:324`) — already supports anthropic/gemini/deepseek.
- `get_evidence` / `ddg_search_and_scrape` / `gemini_search_and_answer` (`app.py:286-374`).
- `safe_parse_json` (`app.py:194`), `get_coords` (`app.py:156`), `sse` (`app.py:415`).
- 4-per-parent cap + parent threading + early-break logic inside `run_pipeline`.

## Graceful degradation (match existing defensive style)
- Missing `DEEPSEEK_API_KEY` → skip the China step, emit a `status` note, continue.
- Missing `GEMINI_API_KEY` → skip enrichment; dedup falls back to exact-string matching.
- Anthropic is the only hard dependency (it drives ID, ties, summary).

## Cost note (per tier)
`parents × (1 Anthropic + 1 DeepSeek + 1 Gemini-enrich) + 1 Gemini-dedup`, plus 1 dedup at
the OEM stage. Roughly 3× the LLM calls of `run_pipeline` per parent — expected, given the
specialist split.

## Verification
1. Confirm `.env` has `ANTHROPIC_API_KEY`, `GEMINI_API_KEY`, `DEEPSEEK_API_KEY`.
2. **Ask before running the app** (per project rule). Once approved: start `python app.py`.
3. Hit `http://127.0.0.1:5000/api/map_agentic?product=Intel%20Core%20i9-14900K&depth=2`
   and watch the SSE stream (browser/curl). Verify: `product_identified` → `oem_discovered`
   (mix of non-China + China OEMs, no duplicates) → `tier_complete` per tier (suppliers have
   `source_hint`/URLs, China companies present, no duplicate names) → `complete`.
4. Sanity-check the `complete` payload feeds the existing PDF/DOCX export routes unchanged.
