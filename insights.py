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
import re
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

_OEM_COL_ALIASES = {
    "company_name":      ["company name", "company"],
    "product":           ["product"],
    "regulatory_body":   ["regulatory body"],
    "region":            ["region"],
    "status":            ["status"],
    "details":           ["details"],
    "confidence":        ["confidence"],
}

def _detect_cols(ws, aliases=None):
    """Return {logical_name: 1-based column index} for a worksheet."""
    if aliases is None:
        aliases = _COL_ALIASES
    header = {
        str(cell.value).strip().lower(): cell.column
        for cell in ws[1] if cell.value
    }
    result = {}
    for logical, alias_list in aliases.items():
        for alias in alias_list:
            if alias in header:
                result[logical] = header[alias]
                break
    return result


# ── Data ingestion ────────────────────────────────────────────────────────────

def _ingest_files(files_data, stream_id):
    """
    Classify rows from all uploaded files into analysis_rows and research_rows.
    Also reads OEM sheets for regulatory data (pharmaceutical products only).

    files_data: list of {"filename": str, "bytes": bytes}
    Returns (analysis_rows, research_rows, regulatory_by_file).
      regulatory_by_file: {filename: [{"oem_name", "regulatory_body", "region",
                                       "status", "details", "confidence"}, ...]}
      Non-empty only when the uploaded file was generated for a pharma product
      (indicated by the presence of "Regulatory Body" column in the OEM sheet).
    """
    analysis_rows      = []
    research_rows      = []
    regulatory_by_file = {}   # filename → list of regulatory dicts

    for item in files_data:
        filename = item["filename"]
        _status(stream_id, f"📂 Reading {filename}…")

        try:
            wb = openpyxl.load_workbook(io.BytesIO(item["bytes"]), data_only=True)
        except Exception as e:
            _log(stream_id, f"[error] Could not open {filename}: {e}", True)
            continue

        # ── Tier sheets ───────────────────────────────────────────────────────
        tier_sheets = [ws for ws in wb.worksheets if "tier" in ws.title.lower()]
        if not tier_sheets:
            _log(stream_id, f"[warn] No tier sheets found in {filename} — skipping", True)
            continue

        for ws in tier_sheets:
            cols = _detect_cols(ws)
            if not cols.get("company_name"):
                _log(stream_id, f"[warn] Sheet '{ws.title}' has no 'Company Name' column — skipping", True)
                continue

            tier_match = re.search(r'tier[_\s]*(\d+)', ws.title.lower())
            tier_num   = int(tier_match.group(1)) if tier_match else 1

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
                    "tier":         tier_num,
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

        # ── OEM sheets — read regulatory data for pharma products ─────────────
        oem_sheets = [ws for ws in wb.worksheets if "oem" in ws.title.lower()]
        reg_rows   = []
        for ws in oem_sheets:
            cols = _detect_cols(ws, _OEM_COL_ALIASES)
            if not cols.get("regulatory_body"):
                continue  # non-pharma OEM sheet — no regulatory columns
            for row in ws.iter_rows(min_row=2, values_only=True):
                def oget(key):
                    col = cols.get(key)
                    if col is None:
                        return ""
                    val = row[col - 1]
                    return str(val).strip() if val is not None else ""
                oem_name = oget("company_name")
                reg_body = oget("regulatory_body")
                if not oem_name or not reg_body:
                    continue
                reg_rows.append({
                    "oem_name":        oem_name,
                    "regulatory_body": reg_body,
                    "region":          oget("region"),
                    "status":          oget("status"),
                    "details":         oget("details"),
                    "confidence":      oget("confidence"),
                })
        if reg_rows:
            regulatory_by_file[filename] = reg_rows
            _log(stream_id, f"[pharma] {filename}: {len(reg_rows)} regulatory row(s) found")

    return analysis_rows, research_rows, regulatory_by_file


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

    # Circular relationships — mutual edges (A supplies B and B supplies A)
    circular_pairs = []
    seen_pairs = set()
    for a, b in G.edges():
        if G.has_edge(b, a):
            key = frozenset([a, b])
            if key not in seen_pairs:
                seen_pairs.add(key)
                circular_pairs.append({
                    "company_a": a,
                    "company_b": b,
                    "inferred":  node_inferred.get(a, False) or node_inferred.get(b, False),
                })

    return {
        "hub_companies":        hub_companies,
        "bottleneck_companies": bottleneck_companies,
        "circular_pairs":       circular_pairs,
    }


# ── Per-product helpers ───────────────────────────────────────────────────────

def _group_by_source(analysis_rows):
    """Group rows by source_file. Returns list of (display_name, rows) sorted by source_file."""
    groups = {}
    for row in analysis_rows:
        groups.setdefault(row["source_file"], []).append(row)
    result = []
    for source_file, rows in sorted(groups.items()):
        names = [r["product"] for r in rows if r.get("product")]
        display = names[0] if names else source_file
        result.append((display, rows))
    return result


def _most_important_oem(rows):
    """Return the supplies_to value with the highest in-degree across these rows."""
    counter = Counter(r["supplies_to"] for r in rows if r.get("supplies_to"))
    return counter.most_common(1)[0][0] if counter else None


def _analyse_product(rows):
    """Compute per-product hub companies, bottlenecks, and top components."""
    import networkx as nx
    G = nx.DiGraph()
    node_inferred = {}
    for row in rows:
        src, dst = row["company_name"], row["supplies_to"]
        if src and dst:
            G.add_edge(src, dst)
            node_inferred[src] = node_inferred.get(src, False) or row["_inferred"]

    hub_companies = [
        {"company": n, "supplies_to_count": d, "inferred": node_inferred.get(n, False)}
        for n, d in sorted(G.out_degree(), key=lambda x: x[1], reverse=True)[:5]
        if d > 0
    ]

    if G.number_of_nodes() > 1:
        bc = nx.betweenness_centrality(G)
        bottlenecks = [
            {"company": n, "centrality_score": round(s, 4), "inferred": node_inferred.get(n, False)}
            for n, s in sorted(bc.items(), key=lambda x: x[1], reverse=True)[:5]
            if s > 0
        ]
    else:
        bottlenecks = []

    all_components = []
    for row in rows:
        comp_str = row.get("components", "")
        if comp_str:
            all_components.extend(c.strip() for c in comp_str.split(",") if c.strip())
    top_components = [{"component": c, "count": n} for c, n in Counter(all_components).most_common(10)]

    return {"hub_companies": hub_companies, "bottlenecks": bottlenecks, "top_components": top_components}


def _chart_choropleth(rows, title="Supplier Geographic Concentration"):
    """Choropleth PNG of supplier countries. Pass all rows for combined, or product rows for per-product."""
    try:
        import plotly.express as px
        import pandas as pd
        country_counts = Counter(
            r["country"] for r in rows
            if r.get("country") and r["country"] not in ("Unknown", "")
        )
        if not country_counts:
            return None
        df = pd.DataFrame(list(country_counts.items()), columns=["country", "count"])
        fig = px.choropleth(
            df, locations="country", locationmode="country names",
            color="count", color_continuous_scale="Blues",
            title=title,
        )
        fig.update_layout(margin={"l": 0, "r": 0, "t": 40, "b": 0}, height=400)
        return fig.to_image(format="png", width=800, height=400)
    except Exception:
        return None


def _chart_sankey(rows):
    """Sankey diagram PNG for one product: Tier-N → … → Tier-1 → OEM."""
    try:
        import plotly.graph_objects as go

        # Collect unique node names preserving tier order
        node_set, node_index = [], {}
        def _node(name):
            if name not in node_index:
                node_index[name] = len(node_set)
                node_set.append(name)
            return node_index[name]

        # Assign colours by tier (source company's tier)
        tier_colours = {1: "rgba(21,101,192,0.7)", 2: "rgba(230,81,0,0.7)", 3: "rgba(106,27,154,0.7)"}
        link_src, link_tgt, link_val, link_col = [], [], [], []

        for row in rows:
            src, dst = row.get("company_name", ""), row.get("supplies_to", "")
            if src and dst:
                link_src.append(_node(src))
                link_tgt.append(_node(dst))
                link_val.append(1)
                link_col.append(tier_colours.get(row.get("tier", 1), "rgba(100,100,100,0.4)"))

        if not link_src:
            return None

        # Node colours: OEM = green, else by most common tier of outgoing edges
        node_tier = {}
        for row in rows:
            src = row.get("company_name", "")
            if src:
                node_tier.setdefault(src, row.get("tier", 1))
        node_colours = []
        oem_names = {r.get("supplies_to", "") for r in rows if r.get("supplies_to")} - \
                    {r.get("company_name", "") for r in rows if r.get("company_name")}
        for name in node_set:
            if name in oem_names:
                node_colours.append("rgba(27,94,32,0.9)")
            else:
                node_colours.append(tier_colours.get(node_tier.get(name, 1), "rgba(100,100,100,0.7)"))

        fig = go.Figure(go.Sankey(
            node=dict(label=node_set, color=node_colours, pad=12, thickness=16),
            link=dict(source=link_src, target=link_tgt, value=link_val, color=link_col),
        ))
        fig.update_layout(
            title_text="Supply Chain Flow (Tier-N → OEM)",
            font_size=10, height=500, margin={"l": 20, "r": 20, "t": 50, "b": 20},
        )
        return fig.to_image(format="png", width=900, height=500)
    except Exception:
        return None


def _chart_nodemap(rows):
    """Node-link diagram PNG centred on the most important OEM for one product.

    Node tier is computed relative to the OEM, not from the global sheet tier:
      - OEM itself          → green
      - Directly supplies OEM (relative Tier 1) → blue
      - Supplies a Tier-1 node (relative Tier 2) → orange
    """
    try:
        import networkx as nx
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib.patches import Patch

        oem = _most_important_oem(rows)
        if not oem:
            return None

        # Build subgraph: OEM + relative Tier-1 + relative Tier-2
        rel_tier1 = {r["company_name"] for r in rows if r.get("supplies_to") == oem}
        subgraph_rows = [
            r for r in rows
            if r.get("supplies_to") == oem or r.get("supplies_to") in rel_tier1
        ]
        if not subgraph_rows:
            return None

        G = nx.DiGraph()
        for row in subgraph_rows:
            G.add_edge(row["company_name"], row["supplies_to"])

        # Assign relative tier to each node
        def _rel_tier(node):
            if node == oem:
                return 0          # OEM
            if node in rel_tier1:
                return 1          # directly feeds OEM
            return 2              # feeds a Tier-1 node

        colour_map = {0: "#1B5E20", 1: "#1565C0", 2: "#E65100"}
        node_list   = list(G.nodes())
        n_nodes     = len(node_list)

        # Scale sizes and font down for large graphs
        if n_nodes <= 12:
            node_size, font_size, k_spread = 900, 8, 2.0
        elif n_nodes <= 25:
            node_size, font_size, k_spread = 550, 7, 1.8
        else:
            node_size, font_size, k_spread = 320, 6, 1.5

        node_colours = [colour_map.get(_rel_tier(n), "#888888") for n in node_list]
        sizes        = [node_size * 1.6 if n == oem else node_size for n in node_list]

        pos = nx.spring_layout(G, seed=42, k=k_spread)

        fig, ax = plt.subplots(figsize=(10, 7))
        ax.set_facecolor("#F8F8F8")

        nx.draw_networkx_nodes(G, pos, ax=ax, nodelist=node_list,
                               node_color=node_colours, node_size=sizes, alpha=0.9)
        nx.draw_networkx_edges(G, pos, ax=ax,
                               edge_color="#999999", arrows=True, arrowsize=14,
                               width=1.0, connectionstyle="arc3,rad=0.05")
        nx.draw_networkx_labels(G, pos, ax=ax,
                                font_size=font_size, font_color="black", font_weight="bold",
                                bbox=dict(facecolor="white", alpha=0.65,
                                          edgecolor="none", boxstyle="round,pad=0.2"))

        ax.set_title(f"Supply Network — Top OEM: {oem}", fontsize=10, pad=10)
        ax.axis("off")

        legend_elements = [
            Patch(facecolor="#1B5E20", label="OEM"),
            Patch(facecolor="#1565C0", label="Tier 1 (direct supplier to OEM)"),
            Patch(facecolor="#E65100", label="Tier 2 (supplier to Tier 1)"),
        ]
        ax.legend(handles=legend_elements, loc="lower left", fontsize=8,
                  framealpha=0.8, edgecolor="#cccccc")

        buf = io.BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight", dpi=150)
        plt.close(fig)
        buf.seek(0)
        return buf.read()
    except Exception:
        return None


def _add_picture_safe(doc, png_bytes, width_inches=5.5):
    """Embed a PNG into the DOCX, silently skip if bytes is None or on error."""
    if not png_bytes:
        return
    try:
        from docx.shared import Inches
        doc.add_picture(io.BytesIO(png_bytes), width=Inches(width_inches))
    except Exception:
        pass


def _add_product_section(doc, helpers, display_name, rows, sub_num, regulatory_rows=None):
    """Render one product sub-section: metrics tables + optional regulatory + Sankey + node-link."""
    _ct       = helpers["_ct"]
    _hdr_row  = helpers["_hdr_row"]
    _flag_row = helpers["_flag_row"]
    _set_bg   = helpers["_set_bg"]

    doc.add_heading(f"{sub_num}. {display_name}", level=2)

    metrics = _analyse_product(rows)

    # Hub companies
    if metrics["hub_companies"]:
        doc.add_heading("Hub Companies", level=3)
        tbl = doc.add_table(rows=1, cols=3)
        tbl.style = "Table Grid"
        _hdr_row(tbl, ["Company", "Customers Supplied", "Note"], "1565C0")
        for entry in metrics["hub_companies"]:
            row = tbl.add_row().cells
            _flag_row(row, entry)
            _ct(row[1], entry["supplies_to_count"])
        doc.add_paragraph()

    # Bottlenecks
    if metrics["bottlenecks"]:
        doc.add_heading("Potential Bottlenecks", level=3)
        tbl = doc.add_table(rows=1, cols=3)
        tbl.style = "Table Grid"
        _hdr_row(tbl, ["Company", "Centrality Score", "Note"], "6A1B9A")
        for entry in metrics["bottlenecks"]:
            row = tbl.add_row().cells
            _flag_row(row, entry)
            _ct(row[1], entry["centrality_score"])
        doc.add_paragraph()

    # Key components
    if metrics["top_components"]:
        doc.add_heading("Key Components", level=3)
        tbl = doc.add_table(rows=1, cols=2)
        tbl.style = "Table Grid"
        _hdr_row(tbl, ["Component", "Frequency"], "1B5E20")
        for entry in metrics["top_components"]:
            row = tbl.add_row().cells
            _ct(row[0], entry["component"])
            _ct(row[1], entry["count"])
        doc.add_paragraph()

    # Regulatory status (pharma products only)
    if regulatory_rows:
        from docx.shared import RGBColor
        doc.add_heading("Regulatory Status (OEM Manufacturers)", level=3)

        _STATUS_COLOURS = {
            "approved":   ("C8E6C9", "1B5E20"),   # green bg, dark text
            "registered": ("C8E6C9", "1B5E20"),
            "pending":    ("FFF9C4", "F57F17"),   # yellow
            "not found":  ("FFCDD2", "B71C1C"),   # red
        }

        tbl = doc.add_table(rows=1, cols=5)
        tbl.style = "Table Grid"
        _hdr_row(tbl, ["OEM Manufacturer", "Regulatory Body", "Region", "Status", "Confidence"], "4A148C")

        for entry in regulatory_rows:
            row = tbl.add_row().cells
            _ct(row[0], entry["oem_name"])
            _ct(row[1], entry["regulatory_body"])
            _ct(row[2], entry["region"])
            status_lower = entry["status"].lower() if entry["status"] else ""
            colours = _STATUS_COLOURS.get(status_lower, ("F5F5F5", "000000"))
            _ct(row[3], entry["status"])
            _set_bg(row[3], colours[0])
            _ct(row[4], entry["confidence"])
        doc.add_paragraph()

    # Per-product choropleth
    choropleth_png = _chart_choropleth(rows, title=f"Geographic Distribution — {display_name}")
    if choropleth_png:
        doc.add_heading("Geographic Distribution", level=3)
        _add_picture_safe(doc, choropleth_png, 5.5)
        doc.add_paragraph()

    # Sankey
    sankey_png = _chart_sankey(rows)
    if sankey_png:
        doc.add_heading("Supply Chain Flow", level=3)
        _add_picture_safe(doc, sankey_png, 5.5)
        doc.add_paragraph()

    # Node-link
    oem = _most_important_oem(rows)
    nodemap_png = _chart_nodemap(rows)
    if nodemap_png:
        doc.add_heading(f"Supply Network — Top OEM: {oem}", level=3)
        _add_picture_safe(doc, nodemap_png, 5.5)
        doc.add_paragraph()


# ── Claude narrative ──────────────────────────────────────────────────────────

def _generate_narrative(metrics, stream_id, regulatory_by_file=None):
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
  "circular_relationships": "2-3 paragraphs: analyse any circular supply relationships detected (Company A supplies Company B AND Company B supplies Company A). Name the specific pairs. Explain why circular relationships are anomalous in a supply chain — they may indicate data quality issues, legitimate cross-supply agreements, or joint venture structures that require scrutiny. Discuss the operational and financial risk these create (circular dependency, price-fixing exposure, single-point-of-failure amplification). Write empty string if no circular pairs were detected.",
  "cross_product": "2-3 paragraphs on suppliers shared across multiple products, naming specific companies and the products they serve. Explain the dual risk: shared suppliers increase efficiency but create correlated failure — a disruption hits multiple product lines at once. Write empty string if only 1 product or file was analysed.",
  "needs_research": "2 paragraphs: summarise the entries that require further research (Company Exists + Supply Ties confirmed but component unverified). Name specific companies if present. Recommend concrete next steps — targeted web searches, direct outreach, or procurement team verification. Write empty string if the list is empty."
}}"""

    # Collect pharma regulatory rows (used by third call below)
    all_reg_rows = []
    if regulatory_by_file:
        for rows in regulatory_by_file.values():
            all_reg_rows.extend(rows)

    metrics_json = json.dumps(metrics, indent=2)

    prompt_a = f"""You are a senior supply chain intelligence analyst writing a formal briefing document.
Below is a structured JSON object containing quantitative metrics from one or more verified supplier datasets.
Write substantive, detailed analytical prose for each section.

Rules:
- Every claim MUST cite a specific number, percentage, or company name from the metrics JSON.
- Each section should be 2-3 full paragraphs of analytical prose separated by \\n\\n.
- Do not use bullet points or headers inside the text values.

METRICS:
{metrics_json}

Return ONLY valid JSON (no markdown, no code fences) with these exact keys:
{{
  "executive_summary": "2-3 paragraphs: overall picture — total verified vs inferred suppliers, products analysed, dominant geography, most critical component category, and the single most important risk finding. Name specific companies and countries.",
  "geo_concentration": "2-3 paragraphs: which countries dominate and by what percentage? Geopolitical risk implications (sanctions, trade restrictions). Cite top 3-5 countries by name and count. What would a disruption in the top country mean?",
  "key_components": "2-3 paragraphs: which components appear most frequently? Are there components supplied by very few companies? Cite top components by name and frequency. Discuss substitutability and criticality.",
  "hub_companies": "2-3 paragraphs: identify the most connected suppliers by out-degree. Name top hub companies and their customer counts. Explain cascade risk. Note which relationships are inferred vs verified."
}}"""

    prompt_b = f"""You are a senior supply chain intelligence analyst writing a formal briefing document.
Below is a structured JSON object containing quantitative metrics from one or more verified supplier datasets.
Write substantive, detailed analytical prose for each section.

Rules:
- Every claim MUST cite a specific number, percentage, or company name from the metrics JSON.
- Each section should be 2-3 full paragraphs of analytical prose separated by \\n\\n.
- Do not use bullet points or headers inside the text values.

METRICS:
{metrics_json}

Return ONLY valid JSON (no markdown, no code fences) with these exact keys:
{{
  "bottlenecks": "2-3 paragraphs: analyse betweenness centrality results. Name top bottleneck companies and their scores. Explain in plain terms what betweenness centrality means. Recommend mitigation strategies.",
  "circular_relationships": "2-3 paragraphs: analyse any circular supply relationships detected. Name specific pairs. Explain why circularity is anomalous and the risks it creates. Write empty string if none detected.",
  "cross_product": "2-3 paragraphs: suppliers shared across multiple products — name specific companies and products. Explain the dual risk of shared suppliers (efficiency vs correlated failure). Write empty string if only 1 product analysed.",
  "needs_research": "2 paragraphs: summarise entries requiring further research. Name specific companies. Recommend concrete next steps. Write empty string if none."
}}"""

    _status(stream_id, "🤖 Generating narrative part 1/2…")
    raw_a  = call_ai(prompt_a, "anthropic", max_tokens=6000)
    part_a = safe_parse_json(raw_a)

    _status(stream_id, "🤖 Generating narrative part 2/2…")
    raw_b  = call_ai(prompt_b, "anthropic", max_tokens=6000)
    part_b = safe_parse_json(raw_b)

    # Merge — fall back to empty string for any section that failed to parse
    _EMPTY = {"executive_summary": "", "geo_concentration": "", "key_components": "",
              "hub_companies": "", "bottlenecks": "", "circular_relationships": "",
              "cross_product": "", "needs_research": "", "regulatory_analysis": ""}

    result = {**_EMPTY}
    if isinstance(part_a, dict):
        result.update({k: v for k, v in part_a.items() if k in _EMPTY})
    if isinstance(part_b, dict):
        result.update({k: v for k, v in part_b.items() if k in _EMPTY})

    # If part_a failed entirely, surface the raw text in executive_summary so nothing is silently lost
    if not isinstance(part_a, dict) or "executive_summary" not in part_a:
        result["executive_summary"] = raw_a if not isinstance(part_a, dict) else result["executive_summary"]

    # Pharma regulatory section — third call if needed
    if all_reg_rows:
        _status(stream_id, "🤖 Generating regulatory narrative…")
        reg_json_str = json.dumps(all_reg_rows, indent=2)
        prompt_reg = f"""You are a senior supply chain intelligence analyst.
Analyse the regulatory approval landscape for the OEM manufacturers in this pharmaceutical supply chain.

REGULATORY DATA:
{reg_json_str[:3000]}

Return ONLY valid JSON (no markdown):
{{
  "regulatory_analysis": "2-3 paragraphs: which regulatory bodies are represented (FDA, EMA, PMDA, etc.)? Which OEMs are fully approved vs pending vs not found? What does the approval pattern imply for market access, compliance risk, and supply chain resilience? Cite specific OEM names, regulatory bodies, and status values."
}}"""
        raw_reg  = call_ai(prompt_reg, "anthropic", max_tokens=2000)
        part_reg = safe_parse_json(raw_reg)
        if isinstance(part_reg, dict) and "regulatory_analysis" in part_reg:
            result["regulatory_analysis"] = part_reg["regulatory_analysis"]

    return result


# ── DOCX builder ──────────────────────────────────────────────────────────────

def _build_docx(metrics, narrative, analysis_rows, research_rows, regulatory_by_file=None):  # noqa: C901
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

    def _add_narrative(text: str):
        """Split on double-newlines and add each chunk as a separate DOCX paragraph."""
        if not text:
            return
        for chunk in str(text).split("\n\n"):
            chunk = chunk.strip()
            if chunk:
                doc.add_paragraph(chunk)

    # Pass closures to per-product section renderer
    _docx_helpers = {"_ct": _ct, "_hdr_row": _hdr_row, "_flag_row": _flag_row, "_set_bg": _set_bg}

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
    _add_narrative(narrative.get("executive_summary", ""))

    # ── Geographical Concentration ───────────────────────────────────────────
    doc.add_heading("2. Geographical Concentration", level=1)
    _add_narrative(narrative.get("geo_concentration", ""))

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

    choropleth_png = _chart_choropleth(analysis_rows, title="Supplier Geographic Concentration (All Products)")
    _add_picture_safe(doc, choropleth_png, 5.5)

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
    _add_narrative(narrative.get("key_components", ""))
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
    _add_narrative(narrative.get("hub_companies", ""))
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
    _add_narrative(narrative.get("bottlenecks", ""))
    tbl = doc.add_table(rows=1, cols=3)
    tbl.style = "Table Grid"
    _hdr_row(tbl, ["Company", "Centrality Score", "Note"], "6A1B9A")
    for entry in metrics["bottleneck_companies"]:
        row = tbl.add_row().cells
        _flag_row(row, entry)
        _ct(row[1], entry["centrality_score"])
    doc.add_paragraph()

    # ── Regulatory Analysis (pharma only) ────────────────────────────────────
    next_sec = 6
    all_reg_rows = []
    if regulatory_by_file:
        for rows in regulatory_by_file.values():
            all_reg_rows.extend(rows)

    if all_reg_rows:
        doc.add_heading(f"{next_sec}. Regulatory Landscape", level=1)
        reg_narrative = narrative.get("regulatory_analysis", "")
        if reg_narrative:
            _add_narrative(reg_narrative)

        tbl = doc.add_table(rows=1, cols=5)
        tbl.style = "Table Grid"
        _hdr_row(tbl, ["OEM Manufacturer", "Regulatory Body", "Region", "Status", "Confidence"], "4A148C")

        _REG_STATUS_COLOURS = {
            "approved":   "C8E6C9",
            "registered": "C8E6C9",
            "pending":    "FFF9C4",
            "not found":  "FFCDD2",
        }
        for entry in all_reg_rows:
            row = tbl.add_row().cells
            _ct(row[0], entry["oem_name"])
            _ct(row[1], entry["regulatory_body"])
            _ct(row[2], entry["region"])
            _ct(row[3], entry["status"])
            bg = _REG_STATUS_COLOURS.get((entry["status"] or "").lower(), "F5F5F5")
            _set_bg(row[3], bg)
            _ct(row[4], entry["confidence"])
        doc.add_paragraph()
        next_sec += 1

    # ── Circular Relationships ───────────────────────────────────────────────
    if metrics.get("circular_pairs"):
        doc.add_heading(f"{next_sec}. Circular Supply Relationships", level=1)
        narr_circ = narrative.get("circular_relationships", "")
        if narr_circ:
            _add_narrative(narr_circ)
        doc.add_paragraph(
            "The following company pairs have a mutual supply relationship "
            "(A supplies B and B supplies A). These warrant further investigation."
        )
        tbl = doc.add_table(rows=1, cols=3)
        tbl.style = "Table Grid"
        _hdr_row(tbl, ["Company A", "Company B", "Note"], "B71C1C")
        for entry in metrics["circular_pairs"]:
            row = tbl.add_row().cells
            label_a = f"⚠ {entry['company_a']}" if entry.get("inferred") else entry["company_a"]
            label_b = f"⚠ {entry['company_b']}" if entry.get("inferred") else entry["company_b"]
            _ct(row[0], label_a)
            _ct(row[1], label_b)
            note = "Model inference" if entry.get("inferred") else "Circular dependency"
            _ct(row[2], note)
            if entry.get("inferred"):
                _set_bg(row[2], "FFFDE7")
            else:
                _set_bg(row[2], "FFEBEE")
        doc.add_paragraph()
        next_sec += 1

    # ── Cross-product overlap ────────────────────────────────────────────────
    if metrics["cross_product_suppliers"]:
        doc.add_heading(f"{next_sec}. Cross-Product Supplier Overlap", level=1)
        _add_narrative(narrative.get("cross_product", ""))
        tbl = doc.add_table(rows=1, cols=3)
        tbl.style = "Table Grid"
        _hdr_row(tbl, ["Company", "Files / Products", "Note"], "E65100")
        for entry in metrics["cross_product_suppliers"]:
            row = tbl.add_row().cells
            _flag_row(row, entry)
            _ct(row[1], ", ".join(entry["products"]))
        doc.add_paragraph()
        next_sec += 1

    # ── Product Breakdown ────────────────────────────────────────────────────
    product_groups = _group_by_source(analysis_rows)
    doc.add_heading(f"{next_sec}. Product Breakdown", level=1)
    doc.add_paragraph(
        "The following sub-sections provide per-product analysis including hub companies, "
        "potential bottlenecks, key components, and supply chain visualisations."
    )
    reg = regulatory_by_file or {}
    for sub_i, (display_name, prod_rows) in enumerate(product_groups, 1):
        # Regulatory rows keyed by source file — all rows in a group share the same source file
        source_file   = prod_rows[0]["source_file"] if prod_rows else None
        reg_rows_prod = reg.get(source_file, [])
        _add_product_section(doc, _docx_helpers, display_name, prod_rows,
                             f"{next_sec}.{sub_i}", regulatory_rows=reg_rows_prod)
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
            _add_narrative(narr)
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
        analysis_rows, research_rows, regulatory_by_file = _ingest_files(files_data, stream_id)

        if not analysis_rows:
            _push(stream_id, {
                "type": "error",
                "message": "No qualifying rows found. Ensure uploaded files have been through the verification pipeline and have at least one row where Company Exists = Yes and Correct Component Supplied = Yes.",
            })
            return

        total_reg = sum(len(v) for v in regulatory_by_file.values())
        pharma_note = f" 💊 {total_reg} regulatory row(s) found." if total_reg else ""
        _status(stream_id, f"✅ {len(analysis_rows)} analysis rows loaded ({sum(1 for r in analysis_rows if not r['_inferred'])} verified, {sum(1 for r in analysis_rows if r['_inferred'])} inferred). {len(research_rows)} flagged for further research.{pharma_note}")

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

        narrative = _generate_narrative(metrics, stream_id, regulatory_by_file=regulatory_by_file)

        _status(stream_id, "🗺️ Generating choropleth map…")
        _status(stream_id, "📄 Building DOCX report (per-product charts may take ~10s)…")
        docx_bytes = _build_docx(metrics, narrative, analysis_rows, research_rows,
                                 regulatory_by_file=regulatory_by_file)

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
  <link href="https://fonts.googleapis.com/css2?family=Syne:wght@800&display=swap" rel="stylesheet">
  <style>
    .logo{font-family:'Syne',sans-serif;font-weight:800;font-size:1.4rem;letter-spacing:-.02em}
    .logo span{color:#e8ff47}
    body { font-family: monospace; padding: 2rem; background: #111; color: #eee; }
    h2 span { color: #e8ff47; }
    p.desc { color: #aaa; font-size: .88rem; max-width: 700px; line-height: 1.7; margin-bottom: 1.2rem; }
    p.desc strong { color: #eee; }
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
    #file-list { margin: .75rem 0; display: flex; flex-direction: column; gap: .35rem; min-height: 1.5rem; }
    .file-row { display: flex; align-items: center; gap: .75rem; padding: .4rem .65rem; background: #1a1a1a; border: 1px solid #333; border-radius: 3px; }
    .file-name { flex: 1; font-size: .83rem; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .file-size { color: #555; font-size: .75rem; flex-shrink: 0; }
    .remove-btn { background: none; border: none; color: #555; cursor: pointer; padding: 0 2px; font-size: .85rem; line-height: 1; flex-shrink: 0; font-family: monospace; }
    .remove-btn:hover:not(:disabled) { color: #f44336; }
    #addBtn { background: #1a2a1a; color: #4caf50; border: 1px dashed #2e5c2e; padding: .35rem 1rem; font-size: .82rem; font-family: monospace; cursor: pointer; }
    #addBtn:hover:not(:disabled) { background: #213321; }
  </style>
</head>
<body>
  <div class="logo" style="margin-bottom:1.5rem">
    <img src="/static/images/nusW-tliap_transparent_bg.png" width="320" height="80" style="vertical-align:middle"> Supplier<span>Map</span>
  </div>
  <h2>Supply Chain <span>Insights</span></h2>
  <p class="desc">
    Upload one or more verified Excel files (.xlsx) exported from the verification pipeline.<br><br>
    <strong>All 3 Yes</strong> — treated as ground truth, used as-is.<br>
    <strong>Company Exists + Correct Component = Yes, Supply Ties unconfirmed</strong> — included but flagged ⚠ as model inference.<br>
    <strong>Company Exists + Supply Ties = Yes, Correct Component unconfirmed</strong> — excluded from analysis, surfaced in a "Requires Further Research" section.<br>
    All other rows are excluded.
  </p>

  <div id="file-list"><div style="color:#555;font-size:.82rem;padding:.3rem 0">No files added.</div></div>
  <input type="file" id="hiddenInput" accept=".xlsx" multiple style="display:none" onchange="addFiles(this.files)">
  <button id="addBtn" onclick="document.getElementById('hiddenInput').click()">+ Add Files</button>

  <div style="margin-top:1rem">
    <button id="genBtn" onclick="generateInsights()">▶ Generate Report</button>
  </div>

  <div id="timer">⏱ Elapsed: <span id="timerVal">0.0s</span></div>
  <div id="log-container"><div id="log"></div></div>

  <script>
    const log          = document.getElementById('log');
    const logContainer = document.getElementById('log-container');

    const append = (msg, cls) => {
      const d = document.createElement('div');
      if (cls) d.className = cls;
      d.textContent = msg;
      log.appendChild(d);
      logContainer.scrollTop = logContainer.scrollHeight;
    };

    let fileQueue     = [];
    let timerInterval = null, timerStart = null;

    // ── file queue management ────────────────────────────────────────────────
    function addFiles(newFiles) {
      for (const f of newFiles) {
        if (!fileQueue.some(q => q.name === f.name && q.size === f.size))
          fileQueue.push(f);
      }
      renderFileList();
      document.getElementById('hiddenInput').value = '';
    }

    function removeFile(idx) {
      fileQueue.splice(idx, 1);
      renderFileList();
    }

    function renderFileList() {
      const list = document.getElementById('file-list');
      if (!fileQueue.length) {
        list.innerHTML = '<div style="color:#555;font-size:.82rem;padding:.3rem 0">No files added.</div>';
        return;
      }
      list.innerHTML = fileQueue.map((f, i) => `
        <div class="file-row" id="file-row-${i}">
          <button class="remove-btn" onclick="removeFile(${i})">✕</button>
          <span class="file-name">${f.name}</span>
          <span class="file-size">${(f.size / 1024).toFixed(1)} KB</span>
        </div>`).join('');
    }

    function setRunning(running) {
      document.getElementById('genBtn').disabled  = running;
      document.getElementById('addBtn').disabled  = running;
      document.querySelectorAll('.remove-btn').forEach(b => b.disabled = running);
    }

    // ── timer ────────────────────────────────────────────────────────────────
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

    // ── generate ─────────────────────────────────────────────────────────────
    async function generateInsights() {
      if (!fileQueue.length) { append('⚠️ Add at least one .xlsx file first.'); return; }

      log.innerHTML = '';
      setRunning(true);
      append('Uploading ' + fileQueue.length + ' file(s)…');
      startTimer();

      const fd = new FormData();
      for (const f of fileQueue) fd.append('files', f);

      let streamId;
      try {
        const res  = await fetch('/api/insights/generate', { method: 'POST', body: fd });
        const data = await res.json();
        if (data.error) {
          append('Error: ' + data.error, 'error');
          stopTimer(); setRunning(false); return;
        }
        streamId = data.stream_id;
      } catch(e) {
        append('Upload failed: ' + e, 'error');
        stopTimer(); setRunning(false); return;
      }

      append('Files received — running analysis…');
      const es = new EventSource('/api/insights/stream/' + streamId);

      es.onmessage = (e) => {
        const msg = JSON.parse(e.data);
        if (msg.type === 'status') {
          append(msg.message);
        } else if (msg.type === 'log') {
          append(msg.message, msg.is_error ? 'error' : 'log-line');
        } else if (msg.type === 'error') {
          stopTimer(); setRunning(false);
          append('Error: ' + msg.message, 'error');
          es.close();
        } else if (msg.type === 'stream_end') {
          stopTimer(); setRunning(false);
          es.close();
          const a = document.createElement('a');
          a.href = '/api/insights/download/' + streamId;
          a.download = '';
          document.body.appendChild(a); a.click(); document.body.removeChild(a);
          append('✓ Report downloaded.', 'done');
        }
      };

      es.onerror = () => {
        stopTimer(); setRunning(false);
        append('SSE connection lost.', 'error');
        es.close();
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
