# Plan: Insights Report Layer

## Context
After the verification pipeline produces annotated Excel files, we need a new layer that
accepts one or more of those verified files, runs quantitative supply chain analysis, and
generates a combined Word document report with AI-authored narrative insights. This closes
the loop from raw supply chain discovery → verification → actionable intelligence.

---

## User Decisions
| Question | Answer |
|---|---|
| Report scope | One combined report across all uploaded files |
| Output format | Word / DOCX (downloadable) |
| Row classification | Two tiers — see below |
| Frontend | Standalone route `/dev/insights`, same style as `/dev/verify` |

### Row classification (three tiers)

| Company Exists | Supply Ties | Correct Component | Classification | Treatment |
|---|---|---|---|---|
| Yes | Yes | Yes | **Fully verified** | Included in analysis as ground truth, no flag |
| Yes | No / Unknown | Yes | **Model inference** | Included in analysis, flagged ⚠ in report tables |
| Yes | Yes | No | **Needs research** | Excluded from analysis, listed in a separate "Requires Further Research" section in the report |
| Anything else | — | — | **Excluded** | Not used at all |

- **Fully verified**: all three checks passed — treat as ground truth
- **Model inference**: company exists and makes the right component, but the direct supply tie wasn't web-confirmed. Anthropic inferred the relationship. Include in analysis but flag so the reader knows it's softer data. The `_inferred` flag is carried through metrics JSON and DOCX tables (⚠ + yellow fill).
- **Needs research**: supply tie is confirmed but the component being supplied couldn't be verified. The relationship is interesting but incomplete — we know A supplies to B, but not what. Excluded from the main metrics (e.g. won't count toward component frequency or hub scores) but surfaced as a separate table at the end of the DOCX so the analyst knows to follow up.
- The metrics JSON includes `"needs_research": [{"company": ..., "supplies_to": ..., "source_file": ...}, ...]` so Claude can mention the count in the executive summary.

---

## Tech Stack Evaluation

### Options considered and verdicts

| Category | Options | Verdict |
|---|---|---|
| Graph analysis | NetworkX, igraph, graph-tool | **NetworkX** — pip install, Pythonic API, has exactly the algorithms needed. igraph/graph-tool are faster but need C builds; overkill at ~500 nodes |
| Automated insight libs | ydata-profiling, sweetviz, lida, PandasAI | **None** — all are generic stats profilers or LLM wrappers. None understand supply chain semantics (what "supplies_to" means topologically). Custom Pandas + NetworkX → Claude is strictly better |
| Agent frameworks | LlamaIndex CSV agent, LangChain pandas agent | **Skip** — designed for open-ended Q&A, issue 5–15 LLM calls per report, can hallucinate column names. One direct Claude call with pre-computed JSON is cheaper, faster, and deterministic |
| Local LLMs (free) | Ollama + Llama 3.1 8B / Mistral 7B | **Skip for now** — 8GB VRAM minimum, significant quality gap vs Claude Sonnet for structured analytical narrative. Revisit only if API cost becomes a hard constraint |
| Charts in DOCX | matplotlib, plotly | **matplotlib** — saves to `BytesIO` buffer, `doc.add_picture(buf)` accepts it directly. Plotly is interactive HTML and needs an extra `kaleido` export step |

### Chosen stack

| Library | Purpose | Already installed? |
|---|---|---|
| **Pandas** | Row filtering, country aggregation, component frequency, tier distribution | No — add to requirements.txt |
| **NetworkX** | Directed graph — hub detection, bottleneck detection (betweenness centrality), choke points | No — add to requirements.txt |
| **matplotlib** | Bar chart of country concentration embedded in DOCX | No — add to requirements.txt |
| **Claude (claude-sonnet-4-6)** | One narrative generation call, fed structured metrics JSON — no hallucination risk because all numbers are pre-computed | Yes — via `ai.py` |
| **python-docx** | DOCX output | Yes — in requirements.txt, pattern in `report_export.py` |
| **openpyxl** | Read the uploaded verified Excel files | Yes |

**Why this combination works:** Pandas and NetworkX compute exact numbers from the real data.
Claude receives those numbers as structured JSON and writes the narrative — it cannot fabricate
a statistic because the statistic is handed to it. One LLM call total, regardless of how many
rows or files.

---

## Files to Change

| File | Change |
|---|---|
| `insights.py` | **New** — Flask Blueprint with `/dev/insights` page, analysis pipeline, SSE stream, DOCX download |
| `app.py` | Register `insights_bp`, add `from insights import insights_bp` |
| `requirements.txt` | Add `networkx`, `pandas`, `matplotlib` |

---

## Backend: `insights.py`

### Blueprint setup
```python
insights_bp = Blueprint("insights", __name__)
_insights_store = {}   # stream_id → {"path": str, "filename": str}
_insights_lock  = threading.Lock()
```

### Step 1 — Data ingestion
- Accept `multipart/form-data` with multiple `.xlsx` files via `POST /api/insights/generate`
- Read each file with `openpyxl` (already installed)
- Dynamically detect column names from header row (case-insensitive match)
- Classify each row:
  ```python
  exists    = row["Company Exists"]    == "Yes"
  supply    = row["Supply Ties"]       == "Yes"
  component = row["Correct Component"] == "Yes"

  if exists and supply and component:
      row["_inferred"]        = False
      row["_needs_research"]  = False
      analysis_rows.append(row)       # fully verified — ground truth
  elif exists and component:
      row["_inferred"]        = True
      row["_needs_research"]  = False
      analysis_rows.append(row)       # model inference — included, flagged
  elif exists and supply:
      row["_inferred"]        = False
      row["_needs_research"]  = True
      research_rows.append(row)       # needs research — excluded from analysis
  else:
      pass                            # excluded entirely
  ```
- Tag each row with its source filename (for cross-product overlap detection)
- Combine all files into a single list of row dicts (both tiers included)

### Step 2 — Pandas analysis
Build a DataFrame from the filtered rows, then compute:
- **Country concentration**: `df.groupby("country")["company_name"].count()` → top countries by supplier count, sorted descending, with % of total
- **Component frequency**: `df["components_supplied"].value_counts()` → most common components across the supply chain
- **Tier distribution**: count suppliers per tier level
- **Cross-file overlap**: suppliers appearing in more than one uploaded product file (if multiple files uploaded)

### Step 3 — NetworkX analysis
Build a directed graph where edges run `company_name → supplies_to`:
```python
G = nx.DiGraph()
for row in filtered_rows:
    G.add_edge(row["company_name"], row["supplies_to"])
```
Compute:
- **Hub companies** (high out-degree): nodes that supply the most other companies
  → `sorted(G.out_degree(), key=lambda x: x[1], reverse=True)[:10]`
- **Bottlenecks** (high betweenness centrality): nodes whose removal would most disconnect the graph
  → `nx.betweenness_centrality(G)` top 10
- **Choke points** (high in-degree): nodes that many companies depend on upstream

### Step 4 — Claude narrative call
Package all computed metrics into structured JSON:
```python
metrics = {
    "total_verified_suppliers": ...,      # fully verified count
    "total_inferred_suppliers": ...,      # model inference count
    "products_analysed": [...],
    "country_concentration": [{"country": "China", "count": 42, "pct": 38.5}, ...],
    "top_components": [{"component": ..., "count": ...}, ...],
    "hub_companies": [{"company": ..., "supplies_to_count": ..., "inferred": bool}, ...],
    "bottleneck_companies": [{"company": ..., "centrality_score": ..., "inferred": bool}, ...],
    "tier_distribution": {"tier_1": 30, "tier_2": 65, ...},
    "cross_product_suppliers": [{"company": ..., "products": [...], "inferred": bool}, ...],
    "needs_research": [{"company": ..., "supplies_to": ..., "source_file": ...}, ...]
}
```
One call to `call_ai` from `ai.py` using `provider="anthropic"`.

Prompt instructs Claude to:
1. Write an executive summary (3–5 sentences)
2. Analyse geopolitical market concentration with specific country %s
3. Identify key components and their supply risk
4. Call out hub companies to watch and why
5. Flag potential bottlenecks and single points of failure
6. Note cross-product supplier interdependencies (if >1 file uploaded)

**Grounding rule in prompt:** Claude is told to cite the specific numbers from the metrics
JSON — every claim must reference a figure from the data.

Claude returns a JSON object with keys:
`executive_summary`, `geo_concentration`, `key_components`, `hub_companies`,
`bottlenecks`, `cross_product` (empty string if only 1 file)

### Step 5 — DOCX generation
Follow the pattern in `report_export.py` (`/api/export/docx` handler) using `python-docx`.

Document sections:
1. **Title**: "Supply Chain Insights Report"
2. **Report Metadata**: date generated, files analysed, total fully-verified suppliers, total model-inference suppliers
3. **Executive Summary** (Claude narrative — must mention inferred count vs verified count)
4. **Geopolitical Concentration** (Claude narrative + Pandas table: country → count → % + matplotlib bar chart embedded as PNG)
5. **Key Components** (Claude narrative + component frequency table)
6. **Hub Companies** (Claude narrative + top 10 hub table with out-degree counts; ⚠ flag on inferred rows)
7. **Potential Bottlenecks** (Claude narrative + betweenness centrality table; ⚠ flag on inferred rows)
8. **Data Summary** (raw counts per tier, per product file, verified vs inferred breakdown)
9. **Requires Further Research** (table of entries where supply tie was confirmed but component was not — company, supplies to, source file)

**Flagging in tables:** any row where `_inferred=True` gets a ⚠ symbol and light yellow fill. The "Requires Further Research" section uses a distinct orange header to visually separate it from the main analysis.

**Chart generation pattern (matplotlib, safe in Flask threads):**
```python
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

fig, ax = plt.subplots(figsize=(6, 3))
ax.barh(countries[:10], counts[:10], color="#1A237E")
ax.set_xlabel("Supplier Count")
ax.set_title("Geographical Concentration (Top 10)")
chart_buf = io.BytesIO()
fig.savefig(chart_buf, format="png", bbox_inches="tight", dpi=150)
plt.close(fig)
chart_buf.seek(0)
doc.add_picture(chart_buf, width=Inches(5.5))
```

Store final DOCX in a temp file registered in `_insights_store[stream_id]`.
SSE progress events pushed to a queue (same pattern as `verify_bp`).

### Routes
```python
POST /api/insights/generate          # multipart upload, starts background thread, returns stream_id
GET  /api/insights/stream/<stream_id>  # SSE progress stream (with keepalive heartbeat)
GET  /api/insights/download/<stream_id>  # returns DOCX, cleans up temp file
```

---

## Frontend: `/dev/insights` (standalone page)

Served by `GET /dev/insights` inside `insights_bp` — same approach as `/dev/verify` in `verify.py`.
Self-contained HTML string returned as a `Response(html, mimetype="text/html")`.
Same dark terminal visual style as the verify page.

Page elements:
- Heading: "Supply Chain **Insights**" (keyword highlighted in accent colour)
- Description: "Upload one or more verified Excel files. Fully-verified rows (all 3 checks passed) are treated as ground truth. Rows where only Company Exists + Correct Component are Yes are included but flagged as model inference."
- Multi-file input: `<input type="file" accept=".xlsx" multiple>`
- "Generate Report" button
- Log panel `<div id="log">` — streams SSE status messages
- Timer (same pattern as verify page)

`generateInsights()` JS flow:
1. Validate at least one file selected
2. `POST /api/insights/generate` with `FormData` (all selected files appended under key `files`)
3. Receive `stream_id` in JSON response
4. Open `EventSource` to `/api/insights/stream/<stream_id>`
5. Append status messages to log panel
6. On `stream_end` event: auto-trigger DOCX download via programmatic `<a>` click to `/api/insights/download/<stream_id>`
7. Re-enable button, stop timer

---

## requirements.txt additions
```
networkx>=3.0
pandas>=2.0.0
matplotlib>=3.7.0
```

---

## Column name mapping
`insights.py` reads header rows dynamically and matches case-insensitively:

| Logical field | Expected column name in verified Excel |
|---|---|
| Company name | `Company Name` |
| Country | `Country` |
| Supplies to | `Supplies To` |
| Component | `Components Supplied` |
| Company exists verdict | `Company Exists` |
| Correct component verdict | `Correct Component Supplied` |
| Tier | derived from sheet name (e.g. `ALL_TIER_1` → tier_1) |
| Source file / product | tracked as a tag added during ingestion |

---

## Verification (how to test end-to-end)
1. Run the verification pipeline on a small test file to produce a verified `.xlsx`
2. Upload that file to the new Insights section in the UI
3. SSE log should show: "Reading files…", "Running analysis…", "Generating narrative…", "Building document…"
4. DOCX downloads automatically on completion
5. Open DOCX — confirm: narrative cites real company names and country percentages from the file, all tables populated, bar chart renders, no placeholder text
6. Upload 2 different verified files simultaneously — confirm cross-product supplier section appears in the report
