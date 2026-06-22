"""
insights.py — Supply chain insights report generator.

Reads one or more verified Excel files, classifies rows into three tiers,
runs Pandas + NetworkX analysis, generates a Claude narrative, and produces
a downloadable DOCX report.

Row classification:
  Company Exists + Supply Ties + Correct Component = Yes  → fully verified (ground truth)
  Company Exists + Correct Component = Yes, Supply Ties No → model inference (included, flagged ⚠)
  Company Exists + Supply Ties = Yes, Correct Component No → needs research (excluded, listed separately)
  Anything else                                            → excluded entirely

Register in app.py via:
    from insights import insights_bp
    app.register_blueprint(insights_bp)
"""

import io
import json
import os
import tempfile
import threading
import uuid
import datetime
from collections import Counter

import openpyxl
from flask import Blueprint, Response, jsonify, request, send_file

from ai import call_ai, safe_parse_json

insights_bp = Blueprint("insights", __name__)

_insights_queues      = {}
_insights_queues_lock = threading.Lock()
_insights_store       = {}
_insights_store_lock  = threading.Lock()


# ── SSE helpers ───────────────────────────────────────────────────────────────

def _push(stream_id, event):
    with _insights_queues_lock:
        if stream_id in _insights_queues:
            _insights_queues[stream_id].append(event)

def _status(stream_id, msg):
    _push(stream_id, {"type": "status", "message": msg})

def _log(stream_id, msg, is_error=False):
    _push(stream_id, {"type": "log", "message": msg, "is_error": is_error})


# ── Column detection ──────────────────────────────────────────────────────────

_COL_ALIASES = {
    "company_name":      ["company name", "company"],
    "country":           ["country"],
    "supplies_to":       ["supplies to", "supplies_to"],
    "components":        ["components supplied", "components_supplied"],
    "company_exists":    ["company exists"],
    "supply_ties":       ["supply ties exist", "supply ties"],
    "correct_component": ["correct component supplied", "correct component"],
    "product":           ["product"],
}

def _detect_cols(ws):
    """Return {logical_name: 1-based column index} for a worksheet."""
    header = {
        str(cell.value).strip().lower(): cell.column
        for cell in ws[1] if cell.value
    }
    result = {}
    for logical, aliases in _COL_ALIASES.items():
        for alias in aliases:
            if alias in header:
                result[logical] = header[alias]
                break
    return result


# ── Data ingestion ────────────────────────────────────────────────────────────

def _ingest_files(files_data, stream_id):
    """
    Classify rows from all uploaded files into analysis_rows and research_rows.

    files_data: list of {"filename": str, "bytes": bytes}
    Returns (analysis_rows, research_rows).
    """
    analysis_rows = []
    research_rows = []

    for item in files_data:
        filename = item["filename"]
        _status(stream_id, f"📂 Reading {filename}…")

        try:
            wb = openpyxl.load_workbook(io.BytesIO(item["bytes"]), data_only=True)
        except Exception as e:
            _log(stream_id, f"[error] Could not open {filename}: {e}", True)
            continue

        tier_sheets = [ws for ws in wb.worksheets if "tier" in ws.title.lower()]
        if not tier_sheets:
            _log(stream_id, f"[warn] No tier sheets found in {filename} — skipping", True)
            continue

        for ws in tier_sheets:
            cols = _detect_cols(ws)
            if not cols.get("company_name"):
                _log(stream_id, f"[warn] Sheet '{ws.title}' has no 'Company Name' column — skipping", True)
                continue

            for row in ws.iter_rows(min_row=2, values_only=True):
                def get(key):
                    col = cols.get(key)
                    if col is None:
                        return ""
                    val = row[col - 1]
                    return str(val).strip() if val is not None else ""

                company = get("company_name")
                if not company or company.lower() in ("none", "nan", ""):
                    continue

                exists    = get("company_exists").lower()    == "yes"
                supply    = get("supply_ties").lower()       == "yes"
                component = get("correct_component").lower() == "yes"

                entry = {
                    "company_name": company,
                    "country":      get("country") or "Unknown",
                    "supplies_to":  get("supplies_to"),
                    "components":   get("components"),
                    "product":      get("product"),
                    "source_file":  filename,
                }

                if exists and supply and component:
                    entry["_inferred"]       = False
                    entry["_needs_research"] = False
                    analysis_rows.append(entry)
                elif exists and component:
                    entry["_inferred"]       = True
                    entry["_needs_research"] = False
                    analysis_rows.append(entry)
                elif exists and supply:
                    entry["_inferred"]       = False
                    entry["_needs_research"] = True
                    research_rows.append(entry)
                # else: excluded entirely

    return analysis_rows, research_rows


# ── Pandas analysis ───────────────────────────────────────────────────────────

def _run_pandas(analysis_rows, stream_id):
    import pandas as pd
    _status(stream_id, "📊 Running Pandas analysis…")

    df = pd.DataFrame(analysis_rows)
    total = len(df)

    # Country concentration
    country_counts = (
        df.groupby("country")["company_name"]
        .count()
        .sort_values(ascending=False)
    )
    country_concentration = [
        {"country": c, "count": int(n), "pct": round(n / total * 100, 1)}
        for c, n in country_counts.items()
        if c and c != "Unknown"
    ][:20]

    # Component frequency — split comma-separated values
    all_components = []
    for comp_str in df["components"].dropna():
        all_components.extend(
            c.strip() for c in comp_str.split(",")
            if c.strip() and c.strip().lower() not in ("none", "nan", "")
        )
    comp_counter = Counter(all_components)
    top_components = [
        {"component": c, "count": n}
        for c, n in comp_counter.most_common(15)
    ]

    # Tier distribution
    tier_dist = df.groupby("source_file")["company_name"].count().to_dict()

    # Cross-product overlap (only meaningful if multiple files)
    cross_product = []
    unique_files = df["source_file"].nunique()
    if unique_files > 1:
        for company, grp in df.groupby("company_name"):
            files = grp["source_file"].unique().tolist()
            if len(files) > 1:
                cross_product.append({
                    "company":  company,
                    "products": files,
                    "inferred": bool(grp["_inferred"].any()),
                })
        cross_product.sort(key=lambda x: len(x["products"]), reverse=True)

    return {
        "country_concentration":  country_concentration,
        "top_components":         top_components,
        "tier_distribution":      tier_dist,
        "cross_product_suppliers": cross_product[:10],
    }


# ── NetworkX analysis ─────────────────────────────────────────────────────────

def _run_networkx(analysis_rows, stream_id):
    import networkx as nx
    _status(stream_id, "🕸️ Running graph analysis…")

    G = nx.DiGraph()
    node_inferred = {}

    for row in analysis_rows:
        src = row["company_name"]
        dst = row["supplies_to"]
        if src and dst:
            G.add_edge(src, dst)
            node_inferred[src] = node_inferred.get(src, False) or row["_inferred"]

    # Hub companies — high out-degree (supplies to many customers)
    out_degree = sorted(G.out_degree(), key=lambda x: x[1], reverse=True)
    hub_companies = [
        {
            "company":          n,
            "supplies_to_count": d,
            "inferred":         node_inferred.get(n, False),
        }
        for n, d in out_degree[:10] if d > 0
    ]

    # Bottlenecks — high betweenness centrality
    if G.number_of_nodes() > 1:
        bc = nx.betweenness_centrality(G)
        bottleneck_companies = [
            {
                "company":          n,
                "centrality_score": round(s, 4),
                "inferred":         node_inferred.get(n, False),
            }
            for n, s in sorted(bc.items(), key=lambda x: x[1], reverse=True)[:10]
            if s > 0
        ]
    else:
        bottleneck_companies = []

    return {
        "hub_companies":       hub_companies,
        "bottleneck_companies": bottleneck_companies,
    }


# ── Claude narrative ──────────────────────────────────────────────────────────

def _generate_narrative(metrics, stream_id):
    _status(stream_id, "🤖 Generating narrative with Claude…")

    prompt = f"""You are a senior supply chain intelligence analyst writing a formal briefing document.
Below is a structured JSON object containing quantitative metrics computed from one or more
verified supplier datasets. Write substantive, detailed analysis for each section.

Rules:
- Every claim MUST cite a specific number, percentage, or company name from the metrics JSON.
  Do not introduce facts or companies not present in the data.
- Each section should be 2-3 full paragraphs of analytical prose.
- Within each section, cite concrete examples from the data (e.g. specific company names,
  specific countries, specific components, specific centrality scores).
- Where relevant, explain the strategic implication of what the numbers mean — not just
  what they are, but why they matter for supply chain resilience.
- Use plain prose paragraphs separated by \\n\\n. Do not use bullet points or headers inside
  the text values.

METRICS:
{json.dumps(metrics, indent=2)}

Return ONLY valid JSON (no markdown, no code fences) with these exact keys:
{{
  "executive_summary": "2-3 paragraphs: overall picture of the supply chain dataset — total verified vs inferred suppliers, which products were analysed, the dominant geography, the most critical component category, and the single most important risk finding. Name specific companies and countries.",
  "geo_concentration": "2-3 paragraphs: deep analysis of geographic concentration. Which countries dominate and by what percentage? What does that imply for geopolitical risk (sanctions, trade restrictions, natural disasters)? Cite the top 3-5 countries by name and count. Discuss what a disruption in the top country would mean for the overall supply chain.",
  "key_components": "2-3 paragraphs: which components appear most frequently across the supply chain and what does that tell us? Are there components supplied by only a small number of companies (concentration risk)? Cite the top components by name and frequency. Discuss substitutability and criticality.",
  "hub_companies": "2-3 paragraphs: identify the most connected suppliers by out-degree. Name the top hub companies and their customer counts. Explain why a hub company going offline would cascade across the supply chain. Note which of these are model-inferred relationships vs fully verified.",
  "bottlenecks": "2-3 paragraphs: analyse the betweenness centrality results. Name the top bottleneck companies and their centrality scores. Explain in plain terms what betweenness centrality means — a company with high betweenness sits on many of the shortest paths through the supply network, so its failure disconnects large portions of the chain. Recommend mitigation strategies (dual sourcing, safety stock, etc).",
  "cross_product": "2-3 paragraphs on suppliers shared across multiple products, naming specific companies and the products they serve. Explain the dual risk: shared suppliers increase efficiency but create correlated failure — a disruption hits multiple product lines at once. Write empty string if only 1 product or file was analysed.",
  "needs_research": "2 paragraphs: summarise the entries that require further research (Company Exists + Supply Ties confirmed but component unverified). Name specific companies if present. Recommend concrete next steps — targeted web searches, direct outreach, or procurement team verification. Write empty string if the list is empty."
}}"""

    raw = call_ai(prompt, "anthropic", max_tokens=4000)
    parsed = safe_parse_json(raw)
    if not isinstance(parsed, dict) or "executive_summary" not in parsed:
        return {
            "executive_summary": raw,
            "geo_concentration": "", "key_components": "", "hub_companies": "",
            "bottlenecks": "", "cross_product": "", "needs_research": "",
        }
    return parsed


# ── DOCX builder ──────────────────────────────────────────────────────────────

def _build_docx(metrics, narrative, analysis_rows, research_rows):
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    def _set_bg(cell, hex6):
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd  = OxmlElement("w:shd")
        shd.set(qn("w:val"),   "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"),  hex6)
        tcPr.append(shd)

    def _ct(cell, text, bold=False, size=10, hex_color=None):
        para = cell.paragraphs[0]
        run  = para.add_run(str(text) if text is not None else "—")
        run.bold      = bold
        run.font.size = Pt(size)
        if hex_color:
            b = bytes.fromhex(hex_color)
            run.font.color.rgb = RGBColor(b[0], b[1], b[2])

    def _hdr_row(tbl, headers, bg_hex):
        r = tbl.rows[0].cells
        for i, h in enumerate(headers):
            _set_bg(r[i], bg_hex)
            _ct(r[i], h, bold=True, hex_color="FFFFFF")

    def _flag_row(cells, entry, name_key="company"):
        name  = entry.get(name_key, "")
        label = f"⚠ {name}" if entry.get("inferred") else name
        _ct(cells[0], label)
        if entry.get("inferred"):
            _ct(cells[-1], "Model inference")
            _set_bg(cells[-1], "FFFDE7")

    doc = Document()
    for sec in doc.sections:
        sec.top_margin    = Inches(1)
        sec.bottom_margin = Inches(1)
        sec.left_margin   = Inches(1)
        sec.right_margin  = Inches(1)

    # ── Title ────────────────────────────────────────────────────────────────
    t = doc.add_heading("Supply Chain Insights Report", 0)
    t.runs[0].font.color.rgb = RGBColor(0x1A, 0x23, 0x7E)

    date_para = doc.add_paragraph(f"Generated: {datetime.date.today().strftime('%B %d, %Y')}")
    date_para.runs[0].font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    doc.add_paragraph()

    for line in [
        f"Files analysed: {', '.join(metrics['products_analysed'])}",
        f"Fully verified suppliers: {metrics['total_verified_suppliers']}",
        f"Model inference suppliers: {metrics['total_inferred_suppliers']}",
        f"Entries requiring further research: {len(research_rows)}",
    ]:
        doc.add_paragraph(line).runs[0].font.size = Pt(10)

    # ── Executive Summary ────────────────────────────────────────────────────
    doc.add_heading("1. Executive Summary", level=1)
    doc.add_paragraph(narrative.get("executive_summary", ""))

    # ── Geopolitical Concentration ───────────────────────────────────────────
    doc.add_heading("2. Geopolitical Concentration", level=1)
    doc.add_paragraph(narrative.get("geo_concentration", ""))

    if metrics["country_concentration"]:
        top = metrics["country_concentration"][:10]
        labels = [c["country"] for c in reversed(top)]
        values = [c["count"]   for c in reversed(top)]
        fig, ax = plt.subplots(figsize=(6, max(3, len(labels) * 0.45)))
        ax.barh(labels, values, color="#1A237E")
        ax.set_xlabel("Supplier Count", fontsize=8)
        ax.set_title("Geographical Concentration (Top 10)", fontsize=9)
        ax.tick_params(labelsize=7)
        chart_buf = io.BytesIO()
        fig.savefig(chart_buf, format="png", bbox_inches="tight", dpi=150)
        plt.close(fig)
        chart_buf.seek(0)
        doc.add_picture(chart_buf, width=Inches(5.5))

    tbl = doc.add_table(rows=1, cols=3)
    tbl.style = "Table Grid"
    _hdr_row(tbl, ["Country", "Supplier Count", "% of Total"], "1A237E")
    for entry in metrics["country_concentration"]:
        row = tbl.add_row().cells
        _ct(row[0], entry["country"])
        _ct(row[1], entry["count"])
        _ct(row[2], f"{entry['pct']}%")
    doc.add_paragraph()

    # ── Key Components ───────────────────────────────────────────────────────
    doc.add_heading("3. Key Components", level=1)
    doc.add_paragraph(narrative.get("key_components", ""))
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = "Table Grid"
    _hdr_row(tbl, ["Component", "Frequency"], "1B5E20")
    for entry in metrics["top_components"]:
        row = tbl.add_row().cells
        _ct(row[0], entry["component"])
        _ct(row[1], entry["count"])
    doc.add_paragraph()

    # ── Hub Companies ────────────────────────────────────────────────────────
    doc.add_heading("4. Hub Companies", level=1)
    doc.add_paragraph(narrative.get("hub_companies", ""))
    tbl = doc.add_table(rows=1, cols=3)
    tbl.style = "Table Grid"
    _hdr_row(tbl, ["Company", "Customers Supplied", "Note"], "1565C0")
    for entry in metrics["hub_companies"]:
        row = tbl.add_row().cells
        _flag_row(row, entry)
        _ct(row[1], entry["supplies_to_count"])
    doc.add_paragraph()

    # ── Bottlenecks ──────────────────────────────────────────────────────────
    doc.add_heading("5. Potential Bottlenecks", level=1)
    doc.add_paragraph(narrative.get("bottlenecks", ""))
    tbl = doc.add_table(rows=1, cols=3)
    tbl.style = "Table Grid"
    _hdr_row(tbl, ["Company", "Centrality Score", "Note"], "6A1B9A")
    for entry in metrics["bottleneck_companies"]:
        row = tbl.add_row().cells
        _flag_row(row, entry)
        _ct(row[1], entry["centrality_score"])
    doc.add_paragraph()

    # ── Cross-product overlap ────────────────────────────────────────────────
    next_sec = 6
    if metrics["cross_product_suppliers"]:
        doc.add_heading(f"{next_sec}. Cross-Product Supplier Overlap", level=1)
        doc.add_paragraph(narrative.get("cross_product", ""))
        tbl = doc.add_table(rows=1, cols=3)
        tbl.style = "Table Grid"
        _hdr_row(tbl, ["Company", "Files / Products", "Note"], "E65100")
        for entry in metrics["cross_product_suppliers"]:
            row = tbl.add_row().cells
            _flag_row(row, entry)
            _ct(row[1], ", ".join(entry["products"]))
        doc.add_paragraph()
        next_sec += 1

    # ── Data Summary ─────────────────────────────────────────────────────────
    doc.add_heading(f"{next_sec}. Data Summary", level=1)
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = "Table Grid"
    _hdr_row(tbl, ["Source File", "Supplier Count"], "37474F")
    for source, count in metrics["tier_distribution"].items():
        row = tbl.add_row().cells
        _ct(row[0], source)
        _ct(row[1], count)
    doc.add_paragraph()
    next_sec += 1

    # ── Requires Further Research ────────────────────────────────────────────
    if research_rows:
        doc.add_heading(f"{next_sec}. Requires Further Research", level=1)
        narr = narrative.get("needs_research", "")
        if narr:
            doc.add_paragraph(narr)
        doc.add_paragraph(
            "The following entries have a confirmed company existence and supply tie, "
            "but the component supplied could not be verified. Follow-up investigation is recommended."
        )
        tbl = doc.add_table(rows=1, cols=4)
        tbl.style = "Table Grid"
        _hdr_row(tbl, ["Company", "Supplies To", "Country", "Source File"], "BF360C")
        for entry in research_rows:
            row = tbl.add_row().cells
            _ct(row[0], entry["company_name"])
            _ct(row[1], entry["supplies_to"])
            _ct(row[2], entry["country"])
            _ct(row[3], entry["source_file"])

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ── Background worker ─────────────────────────────────────────────────────────

def _run_insights(stream_id, files_data):
    try:
        analysis_rows, research_rows = _ingest_files(files_data, stream_id)

        if not analysis_rows:
            _push(stream_id, {
                "type": "error",
                "message": "No qualifying rows found. Ensure uploaded files have been through the verification pipeline and have at least one row where Company Exists = Yes and Correct Component Supplied = Yes.",
            })
            return

        _status(stream_id, f"✅ {len(analysis_rows)} analysis rows loaded ({sum(1 for r in analysis_rows if not r['_inferred'])} verified, {sum(1 for r in analysis_rows if r['_inferred'])} inferred). {len(research_rows)} flagged for further research.")

        pandas_metrics = _run_pandas(analysis_rows, stream_id)
        graph_metrics  = _run_networkx(analysis_rows, stream_id)

        products = sorted({r["product"] for r in analysis_rows if r["product"]})
        if not products:
            products = sorted({r["source_file"] for r in analysis_rows})

        metrics = {
            "total_verified_suppliers": sum(1 for r in analysis_rows if not r["_inferred"]),
            "total_inferred_suppliers": sum(1 for r in analysis_rows if r["_inferred"]),
            "products_analysed":        products,
            **pandas_metrics,
            **graph_metrics,
            "needs_research": [
                {
                    "company":     r["company_name"],
                    "supplies_to": r["supplies_to"],
                    "source_file": r["source_file"],
                }
                for r in research_rows
            ],
        }

        narrative = _generate_narrative(metrics, stream_id)

        _status(stream_id, "📄 Building DOCX report…")
        docx_bytes = _build_docx(metrics, narrative, analysis_rows, research_rows)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(docx_bytes)
            tmp_path = tmp.name

        date_str  = datetime.date.today().strftime("%Y%m%d")
        filename  = f"supply_chain_insights_{date_str}.docx"
        with _insights_store_lock:
            _insights_store[stream_id] = {"path": tmp_path, "filename": filename}

        _status(stream_id, "✅ Report ready — downloading…")

    except Exception as e:
        _push(stream_id, {"type": "error", "message": str(e)})
    finally:
        _push(stream_id, {"type": "stream_end"})


# ── Routes ────────────────────────────────────────────────────────────────────

@insights_bp.route("/dev/insights")
def insights_page():
    html = """<!DOCTYPE html>
<html>
<head>
  <title>Supply Chain Insights</title>
  <style>
    body { font-family: monospace; padding: 2rem; background: #111; color: #eee; }
    h2 span { color: #e8ff47; }
    p.desc { color: #aaa; font-size: .88rem; max-width: 700px; line-height: 1.7; margin-bottom: 1.2rem; }
    p.desc strong { color: #eee; }
    input[type=file] { margin: 1rem 0; display: block; color: #eee; }
    button { background: #333; color: #eee; border: 1px solid #555; padding: .5rem 1.5rem;
             cursor: pointer; font-family: monospace; font-size: .9rem; }
    button:hover { background: #444; }
    button:disabled { opacity: .4; cursor: not-allowed; }
    #log-container { height: 420px; overflow-y: auto; background: #1a1a1a; border: 1px solid #333;
                     padding: .75rem; margin-top: 1.5rem; border-radius: 4px; }
    #log div { margin: 2px 0; font-size: .85rem; line-height: 1.6; }
    .done     { color: #4caf50; }
    .error    { color: #f44336; }
    .log-line { color: #888; font-size: .78rem; }
    #timer { display: none; font-size: .85rem; color: #e8ff47; margin-top: .75rem; letter-spacing: .05em; }
    #timer span { font-weight: bold; }
  </style>
</head>
<body>
  <h2>Supply Chain <span>Insights</span></h2>
  <p class="desc">
    Upload one or more verified Excel files (.xlsx) exported from the verification pipeline.<br><br>
    <strong>All 3 Yes</strong> — treated as ground truth, used as-is.<br>
    <strong>Company Exists + Correct Component = Yes, Supply Ties unconfirmed</strong> — included but flagged ⚠ as model inference.<br>
    <strong>Company Exists + Supply Ties = Yes, Correct Component unconfirmed</strong> — excluded from analysis, surfaced in a "Requires Further Research" section.<br>
    All other rows are excluded.
  </p>

  <input type="file" id="fileInput" accept=".xlsx" multiple>
  <button id="genBtn" onclick="generateInsights()">Generate Report</button>

  <div id="timer">⏱ Elapsed: <span id="timerVal">0.0s</span></div>
  <div id="log-container"><div id="log"></div></div>

  <script>
    const log          = document.getElementById('log');
    const logContainer = document.getElementById('log-container');

    function append(msg, cls) {
      const d = document.createElement('div');
      if (cls) d.className = cls;
      d.textContent = msg;
      log.appendChild(d);
      logContainer.scrollTop = logContainer.scrollHeight;
    }

    let timerInterval = null, timerStart = null;

    function startTimer() {
      timerStart = Date.now();
      document.getElementById('timer').style.display = 'block';
      document.getElementById('timerVal').textContent = '0.0s';
      if (timerInterval) clearInterval(timerInterval);
      timerInterval = setInterval(() => {
        document.getElementById('timerVal').textContent =
          ((Date.now() - timerStart) / 1000).toFixed(1) + 's';
      }, 100);
    }

    function stopTimer() {
      if (timerInterval) { clearInterval(timerInterval); timerInterval = null; }
      if (!timerStart) return;
      document.getElementById('timerVal').textContent =
        ((Date.now() - timerStart) / 1000).toFixed(1) + 's ✓';
    }

    async function generateInsights() {
      const files = document.getElementById('fileInput').files;
      if (!files.length) { append('⚠️ Please select at least one .xlsx file.'); return; }

      document.getElementById('genBtn').disabled = true;
      log.innerHTML = '';
      append(`Uploading ${files.length} file(s)…`);
      startTimer();

      const fd = new FormData();
      for (const f of files) fd.append('files', f);

      let streamId;
      try {
        const res  = await fetch('/api/insights/generate', { method: 'POST', body: fd });
        const data = await res.json();
        if (data.error) {
          append('Error: ' + data.error, 'error');
          stopTimer();
          document.getElementById('genBtn').disabled = false;
          return;
        }
        streamId = data.stream_id;
      } catch (e) {
        append('Upload failed: ' + e, 'error');
        stopTimer();
        document.getElementById('genBtn').disabled = false;
        return;
      }

      append('Files received — running analysis…');
      const evtSource = new EventSource(`/api/insights/stream/${streamId}`);

      evtSource.onmessage = (e) => {
        const msg = JSON.parse(e.data);
        if (msg.type === 'status') {
          append(msg.message);
        } else if (msg.type === 'log') {
          append(msg.message, msg.is_error ? 'error' : 'log-line');
        } else if (msg.type === 'error') {
          stopTimer();
          append('Error: ' + msg.message, 'error');
          evtSource.close();
          document.getElementById('genBtn').disabled = false;
        } else if (msg.type === 'stream_end') {
          stopTimer();
          evtSource.close();
          document.getElementById('genBtn').disabled = false;
          const a = document.createElement('a');
          a.href     = `/api/insights/download/${streamId}`;
          a.download = '';
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          append('✓ Report downloaded.', 'done');
        }
      };

      evtSource.onerror = () => {
        stopTimer();
        append('SSE connection lost.', 'error');
        evtSource.close();
        document.getElementById('genBtn').disabled = false;
      };
    }
  </script>
</body>
</html>"""
    return Response(html, mimetype="text/html")


@insights_bp.route("/api/insights/generate", methods=["POST"])
def insights_generate():
    uploaded = request.files.getlist("files")
    if not uploaded:
        return jsonify({"error": "No files uploaded"}), 400

    files_data = []
    for f in uploaded:
        if not f.filename.lower().endswith(".xlsx"):
            return jsonify({"error": f"{f.filename} is not an .xlsx file"}), 400
        files_data.append({"filename": f.filename, "bytes": f.read()})

    stream_id = uuid.uuid4().hex
    with _insights_queues_lock:
        _insights_queues[stream_id] = []

    threading.Thread(
        target=_run_insights,
        args=(stream_id, files_data),
        daemon=True,
    ).start()

    return jsonify({"stream_id": stream_id})


@insights_bp.route("/api/insights/stream/<stream_id>")
def insights_stream(stream_id):
    import time as _time

    def generate():
        while True:
            with _insights_queues_lock:
                queue  = _insights_queues.get(stream_id, [])
                events, _insights_queues[stream_id] = queue[:], []

            for event in events:
                yield f"data: {json.dumps(event)}\n\n"
                if event.get("type") == "stream_end":
                    with _insights_queues_lock:
                        _insights_queues.pop(stream_id, None)
                    return

            if not events:
                yield ": keepalive\n\n"
            _time.sleep(0.5)

    return Response(generate(), mimetype="text/event-stream", headers={
        "Cache-Control":    "no-cache",
        "X-Accel-Buffering": "no",
    })


@insights_bp.route("/api/insights/download/<stream_id>")
def insights_download(stream_id):
    with _insights_store_lock:
        entry = _insights_store.pop(stream_id, None)

    if not entry:
        return jsonify({"error": "Report not found or already downloaded"}), 404

    path     = entry["path"]
    filename = entry["filename"]

    response = send_file(
        path,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    @response.call_on_close
    def cleanup():
        try:
            os.remove(path)
        except OSError:
            pass

    return response
