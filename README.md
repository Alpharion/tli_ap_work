# Supply Chain Mapper

An AI-powered supply chain intelligence platform. Given any product, it discovers OEM manufacturers and their upstream suppliers tier by tier, verifies each supplier relationship against live web evidence, and generates an analytical insights report from the verified data.

---

## Table of Contents

1. [Architecture Overview](#architecture-overview)
2. [Running the App](#running-the-app)
3. [Pipeline 1 — Agentic Supply Chain Discovery](#pipeline-1--agentic-supply-chain-discovery)
4. [Pipeline 2 — Supplier Verification](#pipeline-2--supplier-verification)
5. [Pipeline 3 — Insights Report Generation](#pipeline-3--insights-report-generation)
6. [Export](#export)
7. [Routes Reference](#routes-reference)

---

## Architecture Overview

```
User Input (product name)
        │
        ▼
┌─────────────────────────┐
│  Discovery Pipeline     │  pipeline.py → run_agentic_pipeline()
│  Gemini + DeepSeek      │
│  DDG web search + RAG   │
└───────────┬─────────────┘
            │  Exports to .xlsx
            ▼
┌─────────────────────────┐
│  Verification Pipeline  │  verify.py → annotate_workbook()
│  3 web searches / row   │
│  Combined → LLM judge   │
└───────────┬─────────────┘
            │  Annotated .xlsx
            ▼
┌─────────────────────────┐
│  Insights Layer         │  insights.py
│  Pandas + NetworkX      │
│  Claude narrative       │
└───────────┬─────────────┘
            │
            ▼
       DOCX Report
```

**Backend:** Single Flask app (`app.py`) with three Flask Blueprints — `verify_bp`, `report_bp`, `insights_bp`.  
**Frontend:** Each pipeline has its own standalone page (dark terminal UI). The main map UI is at `/`.  
**Streaming:** All pipelines push results as Server-Sent Events (SSE) so the frontend updates live.

---

## Running the App

```bash
# 1. Activate virtual environment (Windows)
.venv\Scripts\activate

# 2. Install dependencies
pip install -r requirements.txt

# 3. Create .env with your API keys
ANTHROPIC_API_KEY=...
GEMINI_API_KEY=...
DEEPSEEK_API_KEY=...
SERPER_API_KEY=...   # optional — swap-in for DDG search

# 4. Start the server
python app.py
# → http://127.0.0.1:5000
```

---

## Pipeline 1 — Agentic Supply Chain Discovery

**Route:** `GET /api/map_agentic`  
**Frontend:** Main map UI at `/` → "⚡ MAP AGENTIC" button  
**Source:** `pipeline.py → run_agentic_pipeline()`

### Why two models?

The agentic pipeline splits every LLM call by geography:

| Model | Geography | Rationale |
|---|---|---|
| **Gemini 2.5 Flash** | Non-China companies (excl. HK, Taiwan) | Stronger general-knowledge + Google Search grounding |
| **DeepSeek Chat** | China companies (incl. HK, Taiwan) | Trained on Chinese-language sources; stronger coverage of Chinese manufacturers |

Both calls fire in parallel via `ThreadPoolExecutor(max_workers=2)` at every step. Results are merged and deduplicated before moving on — on a name collision, Gemini's entry wins.

### Web Search as RAG

Before any LLM call, the pipeline fetches live web evidence using DuckDuckGo (`ddgs`) + `Crawl4AI`. This is a form of **Retrieval-Augmented Generation (RAG)**:

```
Query → DDG (up to 10 URLs) → Crawl4AI (scrapes up to 4 pages, 2000 chars each)
      → Evidence block → injected into LLM prompt as context
```

The LLM is explicitly instructed to use the scraped evidence as its primary source and fall back to training knowledge only if the evidence is absent or inconclusive. This grounds the output in current web data rather than potentially stale training knowledge.

Evidence is fetched **once per step** and shared between the Gemini and DeepSeek calls, halving the number of web searches.

**Rate limit handling:** If DDG rate-limits the session, a `DDGRateLimitError` is raised and the entire pipeline aborts immediately — preventing the cost of firing LLM calls on empty evidence.

### Step-by-step pipeline

#### Step 1 — Product Identification
```
web_search("{product} product specification manufacturer supply chain")
    → scraped evidence
    → Gemini only (single call, max 800 tokens)
    → Returns: product_name, category, industry, key_components, oem_manufacturer, oem_country
```

#### Step 2 — OEM Discovery
```
web_search("{product_name} manufacturers OEM brands companies who make produce")
    → scraped evidence (shared by both calls)
    → PARALLEL:
        Gemini  → "only non-China companies" + evidence → JSON array of OEMs
        DeepSeek → "only China companies" + evidence    → JSON array of OEMs
    → _merge_dedup(gemini_oems, deepseek_oems)
    → Market share fields set to "" (skipped for cost in agentic mode)
```

#### Step 3 — Tier Discovery (repeated per depth level, default depth=2)

For each parent company in the current tier:
```
web_search("{parent} direct Tier-N suppliers manufacturers components parts")
    → scraped evidence
    → PARALLEL:
        Gemini  → "non-China suppliers of {parent}" → JSON array
        DeepSeek → "Chinese suppliers of {parent}"  → JSON array
    → _merge_dedup()
    → Dedup across entire tier (same company supplying multiple parents → one entry per oem_root)
```

At Tier-2+, each parent is capped at 4 upstream suppliers to prevent exponential growth.

#### Step 4 — Executive Summary
```
Gemini only (single call, max 400 tokens)
    → 150–200 word plain-text summary of the full supply chain
```

### Deduplication

Company names are normalised before comparison — legal suffixes (Inc, Ltd, LLC, GmbH, Corp, Group, Holdings, Co, Corporation, Limited, Technologies, Technology, Electronics) are stripped, names lowercased:

```python
"Samsung Electronics Co. Ltd." → "samsung"
"CATL"                         → "catl"
```

If the same normalised name appears in both Gemini and DeepSeek results, Gemini's version of the entry is kept.

### Output

The pipeline emits SSE events of types: `status`, `product_identified`, `oem_discovered`, `tier_complete`, `complete`, `error`. The `complete` event carries the full `supply_chain` dict which can be exported to Excel.

---

## Pipeline 2 — Supplier Verification

**Route:** `POST /dev/verify/run` + `GET /dev/verify/stream/<id>`  
**Frontend:** `GET /dev/verify`  
**Source:** `verify.py → annotate_workbook() → verify_supplier_row()`

Takes an `.xlsx` file exported from the discovery pipeline and adds five columns to each tier sheet: `Verification Notes`, `Company Exists`, `Supply Ties Exist`, `Correct Component Supplied`, `URL`.

### How each row is verified

Each supplier row answers three questions through a web search + LLM pipeline:

**Q1 — Does this company exist?**  
`"{company_name}" supplier manufacturer`

**Q2 — Is there a supply relationship?**  
`"{company_name}" "{supplies_to}" supply partnership`

**Q3 — Does this company make the right component?**  
`"{company_name}" "{components_supplied}" manufacture supply`

All three DDG searches run, each scraping up to 4 non-junk pages via Crawl4AI (20-second timeout per URL). The results are concatenated into a single evidence block:

```
SOURCE: https://example.com/...
CONTENT: ... scraped text (up to 2000 chars per page) ...
---
SOURCE: https://...
CONTENT: ...
```

This combined evidence block (up to ~6000 chars) is sent to the LLM **in a single call** covering all three questions at once — minimising LLM cost.

### The LLM judge prompt

The LLM (configurable — Gemini or Anthropic) receives the combined evidence and answers all three questions in one structured JSON response:

```json
{
  "company_exists":     true/false,
  "source_exists":      "web_evidence" | "training_knowledge",
  "supply_ties":        true/false,
  "source_supply":      "web_evidence" | "training_knowledge",
  "correct_component":  true/false,
  "source_component":   "web_evidence" | "training_knowledge",
  "notes_exists":       "one sentence of reasoning",
  "notes_supply":       "one sentence of reasoning",
  "notes_component":    "one sentence of reasoning"
}
```

The prompt instructs the LLM to use web evidence as primary source and fall back to training knowledge only if the evidence is absent or inconclusive. Each answer is labelled with its source (`web_evidence` or `training_knowledge`) so the reader knows how confident to be.

### Fallback if LLM call fails

`_call_ai_json()` retries up to 3 times with exponential backoff (5s after attempt 1, 10s after attempt 2). If all 3 attempts fail, the row gets `False` for all three fields with notes set to `"LLM call failed"`.

If web searches return no results (DDG blocked or rate-limited), `_build_evidence([])` returns `"No search results returned."` — the LLM still fires and can answer from training knowledge, labelled as `"training_knowledge"` in the output.

### Parallelism and throughput

All rows across all tier sheets are submitted to a `ThreadPoolExecutor(max_workers=5)` simultaneously. Results are written back to the workbook as each future completes, so the sheet fills in out of order.

A snapshot of the current workbook state is saved to a temp file every 10 completed rows. If the SSE connection drops mid-run, the partial file can be downloaded at:
```
GET /dev/verify/snapshot/<stream_id>
```
The snapshot URL is shown as a persistent link above the log panel as soon as the upload succeeds.

### Output columns added

| Column | Values | Source label column |
|---|---|---|
| Company Exists | Yes / No | Company Exists Label (Web Evidence / Training Data) |
| Supply Ties Exist | Yes / No | Supply Ties Exists Label |
| Correct Component Supplied | Yes / No | Correct Component Supplied Label |
| Verification Notes | Combined reasoning sentences | — |
| URL | Comma-separated scraped URLs | — |

---

## Pipeline 3 — Insights Report Generation

**Route:** `POST /api/insights/generate` + `GET /api/insights/stream/<id>`  
**Frontend:** `GET /dev/insights`  
**Source:** `insights.py`

Accepts one or more verified `.xlsx` files and generates a downloadable DOCX report with quantitative supply chain analysis and AI-authored narrative.

### Row classification

Every row in the uploaded files is classified into one of four buckets:

| Company Exists | Supply Ties | Correct Component | Classification | Used in analysis? |
|---|---|---|---|---|
| Yes | Yes | Yes | **Fully verified** | ✅ Yes — ground truth |
| Yes | No | Yes | **Model inference** | ✅ Yes — included, flagged ⚠ |
| Yes | Yes | No | **Needs research** | ❌ No — listed separately |
| Anything else | — | — | **Excluded** | ❌ No |

- **Fully verified:** All three checks passed. Treated as ground truth in all calculations.
- **Model inference:** The company exists and makes the right component, but the direct supply tie to the customer was not web-confirmed — Anthropic inferred it. Included in analysis but marked ⚠ in all tables with a yellow cell fill.
- **Needs research:** Supply relationship confirmed but component unknown. Excluded from analysis metrics; surfaced in a dedicated "Requires Further Research" section at the end of the report so the analyst knows to follow up.
- **Excluded:** Neither company existence nor component match confirmed. Dropped entirely.

### Analysis stack

| Library | What it computes |
|---|---|
| **Pandas** | Country concentration (count + % of total), component frequency (split comma-separated values), tier/file distribution, cross-product supplier overlap |
| **NetworkX** | Directed graph (company_name → supplies_to). Out-degree → hub companies. Betweenness centrality → bottlenecks (nodes whose removal most disconnects the graph) |
| **Claude (claude-sonnet-4-6)** | One narrative call fed all pre-computed metrics as structured JSON |
| **matplotlib** | Bar chart of top-10 country concentration embedded as PNG in the DOCX |
| **python-docx** | Final DOCX assembly |

All quantitative figures are computed by Pandas/NetworkX before any LLM call — Claude receives the numbers as facts and cannot fabricate them. This eliminates hallucination risk on statistics.

### Report sections

1. **Executive Summary** — overall picture, verified vs inferred counts, top country, biggest risk
2. **Geopolitical Concentration** — country breakdown with bar chart and table; strategic risk analysis
3. **Key Components** — most frequent components, supply risk implications
4. **Hub Companies** — top suppliers by out-degree (most customers); ⚠ flag on inferred entries
5. **Potential Bottlenecks** — top nodes by betweenness centrality; explanation of cascade risk
6. **Cross-Product Supplier Overlap** *(only if multiple files uploaded)* — suppliers shared across products
7. **Data Summary** — supplier counts per source file
8. **Requires Further Research** *(only if any)* — table of needs-research entries with recommended next steps

Each section contains 2–3 paragraphs of analytical prose with specific company names, countries, and numeric citations drawn from the metrics JSON.

---

## Export

Verified supply chain data can be exported to Excel via:
- `GET /dev/bulk_export` — runs `PRODUCTS_TO_TEST` list through the pipeline, returns one `.xlsx` workbook
- `POST /api/export/docx` — generates a Word document for a single supply chain (called from the map frontend)
- `POST /api/export/pdf` — generates a PDF (called from the map frontend)
- `POST /api/export/excel` — generates an Excel workbook (called from the map frontend)

---

## Routes Reference

| Method | Route | Description |
|---|---|---|
| GET | `/` | Main map UI |
| GET | `/api/map` | Single-provider SSE discovery pipeline |
| GET | `/api/map_agentic` | Agentic (Gemini + DeepSeek) SSE discovery pipeline |
| GET | `/api/map_all` | All configured providers in series |
| GET | `/api/providers` | List providers + configured status |
| POST | `/api/export/excel` | Download Excel for a supply chain |
| POST | `/api/export/docx` | Download Word doc for a supply chain |
| POST | `/api/export/pdf` | Download PDF for a supply chain |
| GET | `/dev/verify` | Verification upload page |
| POST | `/dev/verify/run` | Upload Excel, start verification, return stream_id |
| GET | `/dev/verify/stream/<id>` | SSE stream for verification progress |
| GET | `/dev/verify/snapshot/<id>` | Download partial results mid-run |
| GET | `/dev/verify/download/<id>` | Download completed verified Excel |
| GET | `/dev/insights` | Insights report upload page |
| POST | `/api/insights/generate` | Upload verified files, start analysis, return stream_id |
| GET | `/api/insights/stream/<id>` | SSE stream for insights generation progress |
| GET | `/api/insights/download/<id>` | Download completed DOCX report |
| GET | `/dev/bulk_export` | Bulk Excel export for PRODUCTS_TO_TEST list |
| GET | `/dev/multi_single_export` | Stream per-product Excel downloads |

---

## Plans

Detailed implementation plans for each feature are in the [`/plans`](./plans/) directory:

- [`plans/agenticPipeline.md`](./plans/agenticPipeline.md) — Agentic pipeline design: geography split, dedup logic, rate limit handling, Serper swap
- [`plans/insight_layer.md`](./plans/insight_layer.md) — Insights layer design: tech stack evaluation, row classification tiers, DOCX structure
