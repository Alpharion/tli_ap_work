# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Pointers to take note of through this project:
1. **NO UNSOLICITED REFACTORING:** NEVER refactor code without asking. Always come up with a brief explaination for the code that you want to change and explain what you are going to do
2. **NO RUNNING THE APP AUTONOMOUSLY:** NEVER run the app by yourself, even for testing purposes. Always ask before running the app, as running the app by yourself can incur unwanted costs.
3. **MENTOR MODE ACTIVE:** You are to work through this project acting as a mentor, I am trying to learn on the job, so kindly make your explainations clear, as if you are a senior developer explaining to a junior developer.
4. **VERSION CONTROL:** For any code refactoring, ensure that you are on the git branch agentic

## Running the App

```bash
# Activate venv first (Windows)
.venv\Scripts\activate

# Start Flask dev server
python app.py
# App runs at http://127.0.0.1:5000
```

## Environment Setup

Create a `.env` file in the project root with API keys for whichever providers you want to use:

```
ANTHROPIC_API_KEY=...
GEMINI_API_KEY=...
OPENAI_API_KEY=...
DEEPSEEK_API_KEY=...
```

Install dependencies:
```bash
pip install -r requirements.txt
```

## Bulk Export / Dev Endpoints

- `GET /dev/bulk_export` — runs `PRODUCTS_TO_TEST` list through the pipeline and returns a single `.xlsx` workbook
- `GET /dev/multi_single_export` — same products but streams SSE and auto-downloads one `.xlsx` per product

To change which products are batch-processed and at what depth, edit the constants near the top of `app.py`:

```python
PRODUCTS_TO_TEST = [...]   # list of product strings
DEPTH    = 2               # tier depth per product (1–3)
PROVIDER = "gemini"        # default provider for bulk runs
```

## Architecture

Single-file Flask backend (`app.py`) + single-template frontend (`templates/index.html`).

**Backend pipeline (`run_pipeline`):**
1. Identify the product via AI — returns structured JSON (name, category, OEM, coords)
2. Discover all OEM manufacturers of that product
3. For each OEM, recursively find Tier-1 through Tier-N suppliers (controlled by `depth`)
4. Generate an executive summary
5. All steps stream results to the caller via a shared `queue` list; the HTTP route converts these to Server-Sent Events (SSE)

**Search backends:**
- Gemini: uses built-in Google Search grounding via `google-genai` SDK — no separate search API key
- Anthropic / OpenAI / DeepSeek: fall back to DuckDuckGo (`ddgs` / `duckduckgo_search`) with HTML scraping; gracefully degrades to model training knowledge if DuckDuckGo is blocked

**AI call abstraction (`call_ai`):** single function dispatching to Anthropic, OpenAI, Gemini, or DeepSeek based on a `provider` string. All prompts request raw JSON output; `safe_parse_json` handles malformed responses with four recovery strategies.

**Country geocoding:** static `COUNTRY_COORDS` dict maps lowercase country names to `(lat, lng)` tuples. `get_coords()` does the lookup.

**Excel export (`build_bulk_workbook`):** uses `openpyxl` to build a workbook with sheets: INDEX, ALL_OEMs, ALL_TIER_1…N, and one detail sheet per product. Helper functions `write_index_sheet`, `write_oem_sheet`, `write_tier_sheet`, `write_product_sheet` handle each sheet type.

**Frontend (`templates/index.html`):** self-contained — all JS and CSS are inline. Leaflet.js renders the world map; `leaflet.heat` adds heatmap overlays. SSE events from `/api/map` or `/api/map_all` drive live rendering. `mergeSupplyChains()` (in JS) merges results from multiple AI providers into a unified view.

**Routes:**
- `GET /` — serves `index.html`
- `GET /api/map` — single-provider SSE stream
- `GET /api/map_all` — runs all configured providers in series, streams combined SSE
- `GET /api/providers` — returns list of providers with `configured` flag
- `POST /api/export/pdf` and `POST /api/export/docx` — download report endpoints (called from frontend)

## Key Constraints

- `available_providers()` currently returns only Anthropic. To re-enable other providers, uncomment the relevant entries in that function.
- Tier depth is capped at 3 in the `/api/map` route handler.
- Each OEM in a tier gets at most 4 upstream suppliers fed into the next tier (enforced by a cap loop in `run_pipeline`).
- Excel sheet names are capped at 31 characters (Excel limit); `safe_sheet_name()` enforces this.
