# Plan: Pharmaceutical Mode & Regulatory Research

## What This Adds

When a user checks "💊 This is a pharmaceutical product" before running the pipeline,
the backend automatically researches FDA, EMA, TGA, HSA, PMDA, MHRA, and NMPA
approval/registration status for each discovered OEM. Results stream live and appear
as a colour-coded table in the frontend below the OEM cards.

---

## Files to Change

| File | Change | Risk |
|---|---|---|
| `pipeline.py` | Add `is_pharma=False` to both `run_pipeline` and `run_agentic_pipeline`; add `REGULATORY_BODIES` constant and `research_regulatory_status()` function | Low — default `False` means all existing callers are unaffected |
| `app.py` | `/api/map`, `/api/map_all`, and `/api/map_agentic` each read `is_pharma` from query params and pass it through | Low — one new `request.args.get` per route, no existing logic touched |
| `report_export.py` | Add a "Regulatory Approval Status" section to the DOCX export handler | Low — section only renders when `regulatory` list is non-empty in the POST body |
| `templates/index.html` | Pharmaceutical checkbox, regulatory CSS, `regSection` HTML block, `renderRegulatory()` JS, update `startMapping()` to pass `is_pharma` param, update `reset()` to clear the regulatory section | Medium — new HTML/CSS/JS added; all existing event handlers and functions are untouched |

## Files NOT Touched

- `verify.py`
- `insights.py`
- `excel_export.py`
- `ai.py`
- `scraper.py`
- `coords.py`

---

## Backend: `pipeline.py`

### New additions at module level

**`REGULATORY_BODIES`** — a list of 7 dicts, one per regulatory body:
`FDA (USA)`, `EMA (EU)`, `TGA (Australia)`, `HSA (Singapore)`, `PMDA (Japan)`,
`MHRA (UK)`, `NMPA (China)`.
Each entry carries `id`, `full` name, and `country`.

**`research_regulatory_status(oem_name, product_name, provider)`** — new function:
- Runs a DDG/Gemini search for `"{oem_name} {product_name} FDA EMA TGA HSA PMDA MHRA NMPA regulatory approval"`
- Passes the combined evidence + a structured prompt to `call_ai` (single call)
- Prompt asks the LLM to classify status per body: `approved | registered | pending | not found | unknown`
- Returns a list of 7 dicts: `{body, full_name, country, status, details, confidence}`
- On parse failure, returns all 7 with `status="unknown"` and `confidence="low"` — never raises

### Changes to `run_pipeline`

- Add `is_pharma: bool = False` to the function signature (default keeps all current callers working)
- After the `push("oem_discovered", ...)` and before tier discovery, add a guarded block:
  - If `is_pharma` is `True` and `oem_list` is non-empty, iterate over OEMs and call `research_regulatory_status` for each
  - Push a `regulatory_found` SSE event with the results so the frontend renders them live
  - Attach `regulatory` list to `supply_chain` so it's included in the `complete` event payload

### Changes to `run_agentic_pipeline`

- Add `is_pharma: bool = False` to the function signature (default keeps all current callers working)
- Same guarded block as above inserted at the same relative position (after OEM discovery, before tier loop)
- Uses **Gemini only** for all regulatory LLM calls — rationale: FDA/EMA/TGA/HSA/PMDA/MHRA are Western bodies where Gemini's English-language coverage is stronger; splitting per-body for NMPA alone adds complexity that isn't worth it

---

## Backend: `app.py`

Three routes each get one new line to read `is_pharma` from the query string and pass it to the pipeline function:

- `GET /api/map` → passes to `run_pipeline`
- `GET /api/map_all` → passes to `run_pipeline` (each iteration)
- `GET /api/map_agentic` → passes to `run_agentic_pipeline`

No other logic in these routes changes.

---

## Backend: `report_export.py`

The existing `/api/export/docx` handler builds a Word document section by section.
A new **section 5 — Regulatory Approval Status** is appended after the tier tables section.

- Only renders when `data.get("regulatory", [])` is non-empty
- Per OEM: company name heading, country sub-heading, 7-row table (one row per regulatory body)
- Columns: Body | Region | Status | Details | Confidence
- Status cell text is colour-coded: green = approved/registered, amber = pending, red = not found, grey = unknown
- Non-pharma DOCX exports (where `regulatory` is absent or `[]`) are completely unaffected

---

## Frontend: `templates/index.html`

### New CSS

Appended to the existing `<style>` block — does not modify any existing rules:
- `.reg-section` — card container (same visual pattern as `.oem-section`)
- `.reg-label` — section heading label
- `.reg-oem-block`, `.reg-oem-name`, `.reg-oem-country` — per-OEM layout
- `.reg-table`, `.reg-table th`, `.reg-table td` — the regulatory body table
- `.reg-status` — inline badge for status values
- `.rs-approved`, `.rs-registered` — green
- `.rs-pending` — amber
- `.rs-not-found` — red
- `.rs-unknown` — grey

### New HTML

Inserted after `#oemSection`, before `#downloadBar` — nothing around it changes:
```
<div id="regSection">   ← hidden by default
  per-OEM regulatory table rendered here by JS
</div>
```

### New pharmaceutical checkbox

Inserted inside the search panel below the existing search row.
Reads as: "💊 This is a pharmaceutical product — search regulatory approvals (FDA, EMA, TGA, HSA, PMDA, MHRA, NMPA)".
Hidden until needed, no impact on existing layout.

### Updated `startMapping()`

Reads `isPharmaCheck.checked` and appends `&is_pharma=true` to the SSE URL when checked.
Both `▶ MAP CHAIN` and `⚡ MAP AGENTIC` buttons call `startMapping()` with different
endpoints, so the checkbox applies to both with no additional wiring needed.

### New JS functions / constants

All new — no existing functions are modified:
- `REG_BODIES_ORDER` — ordered array of body IDs for consistent table row order
- `REG_BODY_REGION` — lookup map: body ID → region string
- `statusClass(status)` — maps status string to CSS class
- `renderRegulatory(regulatory)` — renders the `#regSection` from the `regulatory_found` SSE event

### Updated `reset()`

Adds two lines to the existing `reset()` function to clear `#regSection` visibility and empty `#regContent` on each new search. No other lines in `reset()` change.

### Updated SSE `onmessage` handler

Adds one new `else if` branch to handle `msg.type === 'regulatory_found'` — calls `renderRegulatory(msg.regulatory)`.
All existing branches (`status`, `product_identified`, `oem_discovered`, `tier_complete`, `complete`, `stream_end`, etc.) are untouched.

---

## What Stays the Same / Won't Break

| Concern | Why it's safe |
|---|---|
| All existing SSE events | Only a new `regulatory_found` event type is added; existing handlers are untouched |
| Non-pharma pipeline runs | `is_pharma` defaults to `False`; the regulatory block is completely skipped |
| Bulk export (`/dev/bulk_export`, `/dev/multi_single_export`) | These call `run_pipeline` without `is_pharma` — the default `False` means zero change in behaviour |
| Agentic pipeline (existing calls) | Same default `False` pattern |
| DOCX export for non-pharma products | Section 5 only renders when `regulatory` is non-empty — non-pharma exports produce identical output |
| Verify pipeline | `verify.py` not touched |
| Insights pipeline | `insights.py` not touched |
| Excel export | `excel_export.py` not touched |
| Map UI layout | New HTML is inserted in a new block; existing product card, OEM section, download bar, map, and list tab are untouched |

---

## SSE Event Flow (pharma run)

```
status: "🤖 Provider: ..."
status: "🔍 Identifying product..."
product_identified
status: "🏢 Searching for OEMs..."
oem_discovered
status: "💊 Researching regulatory approvals..."
status: "  🔍 Checking regulatory status for OEM A..."
status: "  🔍 Checking regulatory status for OEM B..."
regulatory_found          ← NEW event, rendered in #regSection
status: "✅ Regulatory research complete..."
status: "🏭 Researching Tier-1 suppliers..."
tier_complete
...
complete
stream_end
```
