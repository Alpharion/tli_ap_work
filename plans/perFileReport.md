# Plan: Per-File Verified Supply Data Reports + Final Report Reorder

## Overview

Two features in one change to `insights.py`:
1. For each uploaded Excel, generate a standalone Word doc (`<product> verified supply data report <timestamp>.docx`) and stream it to the browser as soon as it is ready — before the combined insights report is built.
2. Reorder the per-product sub-sections in the final combined report so the top OEM and its network diagram appear first.

---

## Feature 1 — Per-File Word Docs

### What each doc contains

| Section | Source |
|---|---|
| Title: `{product_name} — Verified Supply Data Report` | Product sheet, cell A1 |
| Generated timestamp + verified/inferred counts | Computed from rows |
| Product overview (Category, Industry, OEM Manufacturer, OEM Country, Key Components, Description) | Product sheet label→value rows 3–8 |
| AI Summary | First provider summary cell in product sheet |
| Tier 1 Suppliers table | analysis_rows where tier == 1, for this file |
| Tier 2 Suppliers table | analysis_rows where tier == 2, for this file |
| Requires Further Research note | research_rows for this file (if any) |

**Table columns (both tier tables):**
Company | Country | Supplies To | Components Supplied | Status (Verified / Model Inference)

### Reading the product sheet

New helper `_read_product_sheet(ws)` reads:
- Cell (1,1) → product name
- Rows 3–8 col A/B → overview field dict (label: value)
- Scans for first row starting with "Provider:" → reads the next row col A as AI summary

`_ingest_files` is extended to also read the product sheet (sheets whose title is NOT "tier" / "oem" / "index") per file, returning a 5th value:
```
product_info_by_file: {filename: {name, category, industry, oem_manufacturer, oem_country, key_components, description, ai_summary}}
```

### New builder function

```python
def _build_per_file_docx(product_info, file_rows, research_rows_for_file) -> bytes
```

Builds a clean Word doc from product_info + the rows for that one file.
Uses the same `_ct`, `_hdr_row`, `_set_bg` helpers as `_build_docx`.

### Timing and download flow

In `_run_insights`, after all files are ingested, loop over each file:

```
for each file:
    build per-file docx
    store under key  f"{stream_id}_{file_index}"
    push SSE event: {"type": "download_file", "key": ..., "filename": ...}
    sleep 350ms          ← gives Firefox time to register the download
then build final combined report
push stream_end (triggers final report download as before)
```

### New backend route

```
GET /api/insights/download_file/<key>
```

Serves the per-file docx from a separate store `_insights_file_store`, then deletes the temp file.

### Frontend change

In the SSE `onmessage` handler, handle the new event type:
```js
} else if (msg.type === 'download_file') {
    const a = document.createElement('a');
    a.href = '/api/insights/download_file/' + msg.key;
    a.download = msg.filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    append('📄 Downloaded: ' + msg.filename, 'done');
}
```

---

## Feature 2 — Final Report Product Section Reorder

Current order inside each product sub-section (`_add_product_section`):
1. Hub Companies table
2. Potential Bottlenecks table
3. Key Components table
4. Tier 1 Suppliers to [OEM] table
5. Tier 2 Suppliers table
6. Regulatory (pharma)
7. Geographic Distribution (3 choropleths)
8. Supply Chain Flow (Sankey)
9. Supply Network node-link diagram

**New order:**
1. **Identified Top OEM** — heading naming the OEM
2. **Supply Network diagram** (node-link) — moved to top so the reader sees the full tree first
3. Tier 1 Suppliers to [OEM] table
4. Tier 2 Suppliers table
5. Hub Companies table
6. Potential Bottlenecks table
7. Key Components table
8. Regulatory (pharma)
9. Geographic Distribution (3 choropleths)
10. Supply Chain Flow (Sankey)

---

## Files Changed

| File | Change |
|---|---|
| `insights.py` | All of the above: new helper, extended `_ingest_files`, new builder, route, frontend JS, `_add_product_section` reorder |

---

## Filename Format

```
<product_name>_verified_supply_data_report_<YYYYMMDD_HHMMSS>.docx
```

Spaces in product name replaced with underscores. Timestamp to the second guarantees uniqueness across a session.
