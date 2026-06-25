"""
pipeline.py — Supply chain discovery pipeline.

Exports:
  run_pipeline(product_input, depth, provider, queue, oem_context, collected_tiers)
  run_agentic_pipeline(product_input, depth, queue, oem_context, collected_tiers)
"""

import datetime
import json
import re
from concurrent.futures import ThreadPoolExecutor

from ai import call_ai, safe_parse_json, get_evidence, _source_url_cache
from coords import get_coords
from scraper import web_search, DDGRateLimitError


# ── Deduplication helpers (used by run_agentic_pipeline) ──────────────────────

_STRIP_SUFFIXES = {
    "inc", "ltd", "llc", "gmbh", "corp", "group", "holdings", "co",
    "corporation", "limited", "technologies", "technology", "electronics",
}

def _normalise_name(name: str) -> str:
    parts = name.lower().strip().split()
    return " ".join(p.rstrip(".,") for p in parts if p.rstrip(".,") not in _STRIP_SUFFIXES)

def _merge_dedup(gemini_results: list, deepseek_results: list) -> list:
    """Union of two supplier lists. On name collision, Gemini's entry wins."""
    seen   = {_normalise_name(e.get("company_name", "")) for e in gemini_results if e.get("company_name")}
    merged = list(gemini_results)
    for entry in deepseek_results:
        if _normalise_name(entry.get("company_name", "")) not in seen:
            merged.append(entry)
    return merged


def run_pipeline(product_input: str, depth: int, provider: str, queue: list, oem_context: list, collected_tiers: dict):
    def push(t, **kw): queue.append({"type": t, **kw})

    search_label = "Google Search" if provider == "gemini" else "DuckDuckGo"
    push("status", message=f"🤖 Provider: {provider} | 🔍 Search: {search_label}")
    push("status", message=f"🔍 Identifying product: {product_input}")

    id_evidence = get_evidence(
        f"{product_input} product specification manufacturer supply chain", provider)

    evidence_note = (
        f"Web research findings:\n{id_evidence}" if id_evidence and id_evidence.strip()
        else "No live web search available — use your training knowledge."
    )
    search_status = (
        "✅ Web evidence gathered" if id_evidence and id_evidence.strip()
        else "⚠️ Search unavailable — using AI training knowledge"
    )
    push("status", message=search_status)

    id_prompt = f"""You are a supply chain research assistant.
Product/Part input: "{product_input}"

{evidence_note}

Return ONLY valid JSON (no markdown) with this exact schema:
{{
  "product_name": "...",
  "product_category": "...",
  "industry": "...",
  "description": "...",
  "key_components": ["...", "..."],
  "oem_manufacturer": "... or unknown",
  "oem_country": "country name or unknown"
}}"""

    try:
        raw = call_ai(id_prompt, provider, max_tokens=800)
        parsed_product = safe_parse_json(raw)
        product_info = parsed_product if isinstance(parsed_product, dict) else None
        if not product_info: raise ValueError("No valid JSON returned")
    except Exception as e:
        push("status", message=f"⚠️ Product ID parse error: {e}")
        product_info = {
            "product_name": product_input, "product_category": "Unknown",
            "industry": "Unknown", "description": "Could not auto-identify.",
            "key_components": [], "oem_manufacturer": "Unknown", "oem_country": "Unknown"
        }

    coords = get_coords(product_info.get("oem_country", ""))
    if coords: product_info["lat"], product_info["lng"] = coords

    push("product_identified", product=product_info)
    push("status", message=f"✅ Identified: {product_info.get('product_name')}")

    # ── OEM Discovery ─────────────────────────────────────────────────────────
    push("status", message="🏢 Searching for all OEM manufacturers of this product…")

    oem_evidence = get_evidence(
        f"{product_info.get('product_name')} manufacturers OEM brands companies who make produce", provider)
    oem_note = (
        "Web research findings:\n" + oem_evidence if oem_evidence and oem_evidence.strip()
        else "No live web search — use your training knowledge."
    )

    existing_note = ""
    if oem_context:
        names = ", ".join(oem_context)
        existing_note = f"""
    EXISTING COMPANIES ALREADY IN LIST: {names}
    - If you know any of these companies by a different name, use EXACTLY the name shown above
    - Do not add duplicates — only add companies genuinely not in this list
    """

    oem_prompt = f"""You are a supply chain research assistant.
Product: {product_info.get('product_name')} ({product_info.get('industry')})

{oem_note}

Find ALL known OEM manufacturers / brands that produce or sell this product or equivalent products.
Include the primary OEM already identified ({product_info.get('oem_manufacturer')}) plus any others.

ELIGIBILITY CRITERIA — a company must meet ALL of the following to be included:
- Currently active and operating as of 2024
- Has been actively manufacturing or producing {product_info.get('product_name')} within the last 5 years ({datetime.datetime.now().year-5} - {datetime.datetime.now().year})
- Has demonstrable, documented manufacturing activity — not just distribution, reselling, or licensing
- Exclude any company that has ceased production, been acquired and shut down, or exited the market before {datetime.datetime.now().year-5}

CRITICAL COMPANY NAME RULES:
- Use the shortest globally recognised name only (e.g. "Samsung" not "Samsung Electronics Co. Ltd.")
- No legal suffixes: drop Inc, Ltd, LLC, GmbH, Co., Corp, Group, Holdings
- Use English names only (e.g. "Panasonic" not "Panasonic Corporation")
- Be consistent: if a company is known by an acronym (TSMC, BASF, ABB), use the acronym
- If the name of the company is followed by (), check the contents within the bracket, and if it is another name for the said company, drop the brackets and its contents

{existing_note}

CRITICAL: Return ONLY a raw JSON array starting with [ and ending with ].
Do NOT return a single object. Do NOT wrap in markdown. Do NOT add explanation text.
Each OEM must be a SEPARATE element in the array.

Example of correct format:
[
  {{"company_name": "Company A", "country": "Japan", "role": "OEM manufacturer", "market_share": "high", "confidence": "high", "notes": "..."}},
  {{"company_name": "Company B", "country": "China", "role": "Contract manufacturer", "market_share": "medium", "confidence": "high", "notes": "..."}}
]

Schema for each element:
{{
  "company_name": "exact company name",
  "country": "country name",
  "role": "OEM manufacturer | Brand owner | Contract manufacturer | Licensor",
  "market_share": "high|medium|low|unknown",
  "confidence": "high|medium|low",
  "notes": "brief note"
}}

Confidence: high=publicly documented, medium=limited docs, low=inferred.
List 2-8 significant OEMs as SEPARATE array elements."""

    try:
        raw_oem = call_ai(oem_prompt, provider, max_tokens=1000)
        print(f"\n[OEM RAW OUTPUT]:\n{raw_oem[:800]}\n")
        parsed = safe_parse_json(raw_oem)
        print(f"[OEM PARSED TYPE]: {type(parsed)}, value: {str(parsed)[:200]}")
        if isinstance(parsed, dict):
            for v in parsed.values():
                if isinstance(v, list):
                    parsed = v
                    break
        oem_list = parsed if isinstance(parsed, list) else []
        if len(oem_list) <= 1:
            objects = re.findall(r'\{[^{}]+\}', raw_oem, re.DOTALL)
            if len(objects) > len(oem_list):
                recovered = []
                for obj in objects:
                    try:
                        recovered.append(json.loads(obj))
                    except Exception:
                        try:
                            recovered.append(json.loads(re.sub(r',\s*([\]}])', r'', obj)))
                        except Exception:
                            pass
                if len(recovered) > len(oem_list):
                    oem_list = recovered
                    push("status", message=f"  🔧 Recovered {len(oem_list)} OEMs from raw text")
        push("status", message=f"  📋 Parsed {len(oem_list)} OEM(s)")
    except Exception as e:
        push("status", message=f"⚠️ OEM discovery parse error: {e}")
        oem_list = []

    for oem in oem_list:
        name = oem.get("company_name", "")
        if name and name not in oem_context:
            oem_context.append(name)
        c = get_coords(oem.get("country", ""))
        if c: oem["lat"], oem["lng"] = c

    # ── Market Share Research ─────────────────────────────────────────────────
    push("status", message="📊 Researching market share from industry reports…")
    pname = product_info.get("product_name", "")
    iname = product_info.get("industry", "")
    cname = product_info.get("product_category", "")
    oem_names_str = ", ".join([o.get("company_name", "") for o in oem_list if o.get("company_name", "")])

    ms_queries = [
        f"{pname} {cname} market share IDC Gartner Statista 2024 2025",
        f"{iname} {pname} manufacturer market share percentage revenue report",
        f"{oem_names_str} {pname} market share ranking",
    ]
    ms_parts, ms_urls = [], []
    for q in ms_queries:
        ev = get_evidence(q, provider)
        if ev and ev.strip():
            ms_parts.append(ev)
            ms_urls.extend(_source_url_cache.get(q, []))
            push("status", message=f"  🌐 Market data: {q[:60]}…")

    ms_note = (
        "Market research findings:\n" + "\n---\n".join(ms_parts)
        if ms_parts else
        "No live market data — estimate from training knowledge."
    )

    ms_prompt = f"""You are a market intelligence analyst.
Product: {pname} ({iname} / {cname})
Known OEMs: {oem_names_str}

{ms_note}

Based on evidence above, estimate each OEM's market share.
Use actual numbers where found (e.g. "~32% global market").

Return ONLY a valid JSON array (no markdown). Each element:
{{
  "company_name": "exact name matching OEM list above",
  "market_share": "high|medium|low",
  "market_share_pct": "e.g. ~35% or 10-15% or <5% or unknown",
  "market_share_source": "e.g. IDC Q3 2024, Gartner 2023, estimated",
  "market_rank": 1
}}
Rank 1=largest. Only include companies from the known OEMs list."""

    try:
        raw_ms  = call_ai(ms_prompt, provider, max_tokens=800)
        ms_data = safe_parse_json(raw_ms)
        if isinstance(ms_data, list):
            ms_lookup = {m.get("company_name", "").strip().lower(): m for m in ms_data}
            enriched  = 0
            for oem_entry in oem_list:
                key   = oem_entry.get("company_name", "").strip().lower()
                match = ms_lookup.get(key)
                if not match:
                    for k, v in ms_lookup.items():
                        if key in k or k in key:
                            match = v; break
                if match:
                    oem_entry["market_share"]        = match.get("market_share", oem_entry.get("market_share", "unknown"))
                    oem_entry["market_share_pct"]    = match.get("market_share_pct", "unknown")
                    oem_entry["market_share_source"] = match.get("market_share_source", "unknown")
                    oem_entry["market_rank"]         = match.get("market_rank", 99)
                    enriched += 1
                else:
                    oem_entry.setdefault("market_share_pct", "unknown")
                    oem_entry.setdefault("market_share_source", "AI estimate")
                    oem_entry.setdefault("market_rank", 99)
            oem_list.sort(key=lambda x: x.get("market_rank", 99))
            push("status", message=f"  ✅ Market share enriched for {enriched}/{len(oem_list)} OEMs")
        else:
            push("status", message="  ⚠️ Could not parse market share data")
            for oem in oem_list:
                oem.setdefault("market_share_pct", "unknown")
                oem.setdefault("market_share_source", "AI estimate")
    except Exception as e:
        push("status", message=f"  ⚠️ Market share error: {e}")
        for oem_entry in oem_list:
            oem_entry.setdefault("market_share_pct", "unknown")
            oem_entry.setdefault("market_share_source", "AI estimate")

    primary_oem = product_info.get("oem_manufacturer", "")
    if primary_oem and primary_oem.lower() not in ("unknown", ""):
        existing_names = [o.get("company_name", "").lower() for o in oem_list]
        if not any(primary_oem.lower() in n or n in primary_oem.lower() for n in existing_names):
            primary_coords = get_coords(product_info.get("oem_country", ""))
            primary_entry  = {
                "company_name": primary_oem,
                "country":      product_info.get("oem_country", ""),
                "role":         "OEM manufacturer",
                "market_share": "high",
                "confidence":   "high",
                "notes":        "Primary identified manufacturer",
            }
            if primary_coords:
                primary_entry["lat"], primary_entry["lng"] = primary_coords
            oem_list.insert(0, primary_entry)

    push("oem_discovered", oems=oem_list)
    push("status", message=f"✅ Found {len(oem_list)} OEM manufacturer(s)")

    supply_chain = {"product": product_info, "oems": oem_list, "tiers": {}}

    oem_names = [
        {"name": o.get("company_name", ""), "confidence": o.get("confidence", "")}
        for o in oem_list if o.get("company_name", "").lower() not in ("unknown", "")
    ]
    if not oem_names:
        oem_names = [{"name": product_info.get("oem_manufacturer", product_input), "confidence": "high"}]

    previous_parents = [
        {"name": n.get("name", ""), "oem_root": n.get("name", ""),
         "confidence": n.get("confidence", ""), "parent": n.get("name", "")}
        for n in oem_names
    ]

    def tier_note(tier_num):
        key   = f"tier_{tier_num}"
        names = collected_tiers.get(key, [])
        return (
            f"The following companies are already in the tier {tier_num} list: {', '.join(names)}. "
            "For each of the company names, check if it is an alternate name for any of the names "
            "in the given list. If you find the same company, use EXACTLY the same name as in the list."
            if names else ""
        )

    for tier_num in range(1, depth + 1):
        push("status", message=f"🏭 Researching Tier-{tier_num} suppliers…")
        tier_suppliers = []
        next_parents   = []

        if tier_num > 1:
            lookup_dict = {}
            filtered    = []
            for entry in previous_parents:
                parent = entry.get("parent", "")
                count  = lookup_dict.get(parent, 0)
                if count < 4:
                    lookup_dict[parent] = count + 1
                    filtered.append(entry)
            previous_parents = filtered

        for parent_info in previous_parents:
            parent   = parent_info["name"]
            oem_root = parent_info["oem_root"]
            if not parent or parent.lower() == "unknown":
                continue

            query = (
                f"{parent} direct Tier 1 suppliers manufacturers components parts"
                if tier_num == 1 else
                f"{parent} suppliers raw materials components manufacturers supply chain"
            )

            evidence      = get_evidence(query, provider)
            evidence_note = (
                f"Web research findings:\n{evidence}" if evidence and evidence.strip()
                else "No live web search — use your training knowledge."
            )
            status_icon = "🌐" if evidence and evidence.strip() else "🧠"
            push("status", message=f"  {status_icon} [{oem_root}] → {parent}: searching suppliers…")

            tier_prompt = f"""You are a supply chain research assistant.
Product: {product_info.get('product_name')} ({product_info.get('industry')})
OEM root: {oem_root}
Finding Tier-{tier_num} suppliers — companies that DIRECTLY supply components that are part of the manufacturing chain for {product_info.get('product_name')} to: {parent}

{evidence_note}
STRICT INCLUSION RULES — a supplier must meet ALL of the following:
- Supplies {parent} with something physically used in making "{product_info.get('product_name')}"
  or one of its listed key components
- If {parent} makes multiple product lines, only include suppliers relevant to the
  "{product_info.get('product_name')}" product line specifically
- Currently active and operating as of {datetime.datetime.now().year}
- Direct supply relationship within the last 5 years — not distribution, reselling, or licensing

STRICT EXCLUSION RULES — do NOT include:
- Suppliers that provide {parent} with general industrial inputs unrelated to
  "{product_info.get('product_name')}" (e.g. office supplies, IT services, generic packaging)
- Suppliers for {parent}'s other unrelated product lines
- Logistics, shipping, or freight companies
- Financial, legal, or consulting service providers
- Any company that does not contribute a physical material, chemical, or manufactured
  component that ends up in "{product_info.get('product_name')}"
- Recycling companies, waste processors, or end-of-life battery collectors
- Any company whose primary relationship with {parent} is as a customer, not a supplier
- Any company that manufactures the same end product as {parent} — i.e. another producer
  of "{product_info.get('product_name')}". Even if two manufacturers have cross-supply
  agreements, do not include them — they are market competitors, not supply chain inputs
- Any company that appears in the OEM manufacturer list for "{product_info.get('product_name')}" —
  OEM manufacturers of the same product are never valid Tier-{tier_num} material suppliers
  to each other
- Internal subsidiaries or parent companies of {parent} — only include independent
  third-party suppliers

CRITICAL COMPANY NAME RULES:
- Use the shortest globally recognised name only (e.g. "Samsung" not "Samsung Electronics Co. Ltd.")
- No legal suffixes: drop Inc, Ltd, LLC, GmbH, Co., Corp, Group, Holdings
- Use English names only (e.g. "Panasonic" not "Panasonic Corporation")
- Be consistent: if a company is known by an acronym (TSMC, BASF, ABB), use the acronym
- If the name of the company is followed by (), check the contents within the bracket, and if it is another name for the said company, drop the brackets and its contents
{tier_note(tier_num)}

Return ONLY a valid JSON array (no markdown). Each element:
{{
  "company_name": "exact company name",
  "country": "country name",
  "supplies_to": "{parent}",
  "oem_root": "{oem_root}",
  "components_supplied": ["component1", "component2"],
  "confidence": "high|medium|low",
  "source_hint": "brief source note"
}}

List 3-5 distinct real direct suppliers to {parent} specifically. Do NOT mix in suppliers of other companies."""

            try:
                raw   = call_ai(tier_prompt, provider, max_tokens=1200)
                found = safe_parse_json(raw)
                if isinstance(found, list):
                    for s in found:
                        s["oem_root"]    = oem_root
                        s["supplies_to"] = parent
                    tier_suppliers.extend(found)
                    for s in found:
                        sname = s.get("company_name", "").strip()
                        if sname and sname.lower() != "unknown":
                            next_parents.append({
                                "name": sname, "oem_root": oem_root,
                                "confidence": s.get("confidence"),
                                "parent": s.get("supplies_to"),
                            })
            except Exception as e:
                push("status", message=f"  ⚠️ Parse error tier {tier_num} [{parent}]: {e}")

        previous_parents = next_parents

        seen, unique = set(), []
        for s in tier_suppliers:
            key = f"{s.get('company_name','').strip().lower()}|{s.get('oem_root','')}"
            if key and key not in seen:
                seen.add(key)
                c = get_coords(s.get("country", ""))
                if c: s["lat"], s["lng"] = c
                unique.append(s)

        for s in unique:
            name = s.get("company_name", "")
            collected_tiers.setdefault(f"tier_{tier_num}", [])
            if name and name not in collected_tiers[f"tier_{tier_num}"]:
                collected_tiers[f"tier_{tier_num}"].append(name)

        supply_chain["tiers"][f"tier_{tier_num}"] = unique
        push("tier_complete", tier=tier_num, suppliers=unique)
        push("status", message=f"✅ Tier-{tier_num}: {len(unique)} suppliers found")

        if not [s.get("company_name", "") for s in unique]:
            push("status", message=f"⚠️ No suppliers at Tier-{tier_num}, stopping.")
            break

    # ── Executive summary ─────────────────────────────────────────────────────
    push("status", message="📊 Generating executive summary…")
    summary_prompt = f"""You are a supply chain analyst.
Supply chain data:
{json.dumps(supply_chain, indent=2)[:6000]}

Write a concise executive summary (150-200 words) covering:
- What the product is and its industry
- Key Tier-1 suppliers and their roles
- Geographic concentration or risks
- Overall supply chain complexity
Return plain text only, no markdown."""

    try:
        supply_chain["summary"] = call_ai(summary_prompt, provider, max_tokens=400).strip()
    except Exception as e:
        supply_chain["summary"] = f"Summary unavailable: {e}"

    supply_chain["provider"] = provider
    push("complete", supply_chain=supply_chain)


def verify_manufacturer(company_name: str, product_name: str, provider: str) -> dict:
    """Three-layer manufacturer verification: DDG → keyword filter → LLM."""
    try:
        from ddgs import DDGS
        with DDGS() as ddg:
            results = list(ddg.text(
                f"{company_name} {product_name} manufacturer supplier",
                max_results=5
            ))
    except Exception:
        return {"verified": False, "confidence": "low",
                "reason": "DDG unavailable", "source_urls": [], "layer": 1}

    if not results:
        return {"verified": False, "confidence": "low",
                "reason": "No web presence found", "source_urls": [], "layer": 1}

    urls     = [r.get("href", "") for r in results if r.get("href")]
    snippets = [r.get("body", "") + " " + r.get("title", "") for r in results]
    combined = " ".join(snippets).lower()

    MANUFACTURE_KEYWORDS = [
        "manufactures", "manufacturer", "produces", "production",
        "supplier", "supplies", "fabricates", "assembles", "factory", "plant", "facility"
    ]
    EXCLUDE_KEYWORDS = [
        "distributor", "reseller", "retailer", "discontinued",
        "shut down", "acquired and closed", "bankrupt", "no longer produces"
    ]

    company_hit  = company_name.lower() in combined
    product_hit  = product_name.lower() in combined
    mfg_hit      = any(k in combined for k in MANUFACTURE_KEYWORDS)
    exclude_hit  = any(k in combined for k in EXCLUDE_KEYWORDS)

    if not company_hit or exclude_hit:
        return {
            "verified":    False,
            "confidence":  "low",
            "reason":      "excluded by keyword filter" if exclude_hit else "company not found in search results",
            "source_urls": urls,
            "layer":       2,
        }

    if company_hit and product_hit and mfg_hit:
        return {
            "verified":    True,
            "confidence":  "high",
            "reason":      "company, product and manufacturing activity all confirmed in search results",
            "source_urls": urls,
            "layer":       2,
        }

    evidence = "\n---\n".join(snippets[:3])
    prompt   = f"""Based ONLY on the following web search snippets, determine if "{company_name}" manufactures or supplies "{product_name}".
Do not use your training knowledge — only use what the snippets say.

Snippets:
{evidence}

Reply ONLY with a JSON object:
{{
  "verified": true/false,
  "confidence": "high|medium|low",
  "reason": "one sentence citing specific snippet evidence",
  "manufactures_product": true/false
}}"""

    try:
        raw    = call_ai(prompt, provider, max_tokens=150)
        result = safe_parse_json(raw)
        result["source_urls"] = urls
        result["layer"]       = 3
        return result
    except Exception:
        score = sum([company_hit, product_hit, mfg_hit])
        return {
            "verified":    score >= 2,
            "confidence":  "medium" if score >= 2 else "low",
            "reason":      "LLM verification failed — keyword fallback used",
            "source_urls": urls,
            "layer":       2,
        }


def run_agentic_pipeline(product_input: str, depth: int, queue: list, oem_context: list, collected_tiers: dict):
    """
    Dual-model agentic pipeline.
    - Gemini   → non-China companies at every level
    - DeepSeek → China companies at every level
    Results are merged and deduplicated (Gemini wins on collision) before
    each tier level is processed, keeping the total LLM call count down.
    Evidence is gathered via DDG once per step and shared by both models.
    Aborts immediately if DDG rate-limits the session.
    """
    def push(t, **kw): queue.append({"type": t, **kw})

    def _format_evidence(results: list) -> str:
        if not results:
            return ""
        return "\n\n---\n\n".join(f"Source: {r['url']}\n{r['content']}" for r in results)

    def _parse_list(raw: str) -> list:
        """Parse a JSON array from raw LLM output with fallback object recovery."""
        parsed = safe_parse_json(raw)
        if isinstance(parsed, dict):
            for v in parsed.values():
                if isinstance(v, list):
                    parsed = v
                    break
        result = parsed if isinstance(parsed, list) else []
        if len(result) <= 1:
            objects = re.findall(r'\{[^{}]+\}', raw, re.DOTALL)
            if len(objects) > len(result):
                recovered = []
                for obj in objects:
                    try:
                        recovered.append(json.loads(obj))
                    except Exception:
                        try:
                            recovered.append(json.loads(re.sub(r',\s*([\]}])', r'', obj)))
                        except Exception:
                            pass
                if len(recovered) > len(result):
                    result = recovered
        return result

    def _run_parallel(gemini_prompt: str, deepseek_prompt: str, max_tokens: int, label: str):
        """Fire both LLM calls in parallel; return (gemini_list, deepseek_list)."""
        g_found, d_found = [], []
        with ThreadPoolExecutor(max_workers=2) as ex:
            fut_g = ex.submit(call_ai, gemini_prompt,   "gemini",   max_tokens)
            fut_d = ex.submit(call_ai, deepseek_prompt, "deepseek", max_tokens)
            try:
                g_found = _parse_list(fut_g.result())
                push("status", message=f"  🌍 Gemini {label}: {len(g_found)} result(s)")
            except Exception as e:
                push("status", message=f"  ⚠️ Gemini {label} failed: {e}")
            try:
                d_found = _parse_list(fut_d.result())
                push("status", message=f"  🇨🇳 DeepSeek {label}: {len(d_found)} result(s)")
            except Exception as e:
                push("status", message=f"  ⚠️ DeepSeek {label} failed: {e}")
        return g_found, d_found

    push("status", message="🤖 Agentic Mode | 🌍 Gemini (non-China) + 🇨🇳 DeepSeek (China) | 🔍 DDG Search")
    push("status", message=f"🔍 Identifying product: {product_input}")

    # ── Step 1: Product Identification (Gemini + DDG) ─────────────────────────
    try:
        id_results = web_search(
            f"{product_input} product specification manufacturer supply chain",
            n=10, max_scrape=4
        )
    except DDGRateLimitError as e:
        push("error", message=f"🚫 DDG rate limit — aborting before product ID. ({e})")
        return

    id_ev_text    = _format_evidence(id_results)
    id_ev_note    = f"Web research findings:\n{id_ev_text}" if id_ev_text else "No live web search — use your training knowledge."
    push("status", message="✅ Web evidence gathered" if id_ev_text else "⚠️ No web evidence — using training knowledge")

    id_prompt = f"""You are a supply chain research assistant.
Product/Part input: "{product_input}"

{id_ev_note}

Return ONLY valid JSON (no markdown) with this exact schema:
{{
  "product_name": "...",
  "product_category": "...",
  "industry": "...",
  "description": "...",
  "key_components": ["...", "..."],
  "oem_manufacturer": "... or unknown",
  "oem_country": "country name or unknown"
}}"""

    try:
        raw          = call_ai(id_prompt, "gemini", max_tokens=800)
        parsed_prod  = safe_parse_json(raw)
        product_info = parsed_prod if isinstance(parsed_prod, dict) else None
        if not product_info: raise ValueError("No valid JSON returned")
    except Exception as e:
        push("status", message=f"⚠️ Product ID parse error: {e}")
        product_info = {
            "product_name": product_input, "product_category": "Unknown",
            "industry": "Unknown", "description": "Could not auto-identify.",
            "key_components": [], "oem_manufacturer": "Unknown", "oem_country": "Unknown",
        }

    coords = get_coords(product_info.get("oem_country", ""))
    if coords: product_info["lat"], product_info["lng"] = coords
    push("product_identified", product=product_info)
    push("status", message=f"✅ Identified: {product_info.get('product_name')}")

    # ── Step 2: OEM Discovery (parallel Gemini non-China + DeepSeek China) ────
    push("status", message="🏢 Discovering OEMs — Gemini (non-China) + DeepSeek (China) in parallel…")

    try:
        oem_ev = web_search(
            f"{product_info.get('product_name')} manufacturers OEM brands companies who make produce",
            n=10, max_scrape=4
        )
    except DDGRateLimitError as e:
        push("error", message=f"🚫 DDG rate limit — aborting before OEM discovery. ({e})")
        return

    oem_ev_text = _format_evidence(oem_ev)
    oem_ev_note = f"Web research findings:\n{oem_ev_text}" if oem_ev_text else "No live web search — use your training knowledge."

    existing_note = ""
    if oem_context:
        existing_note = (
            f"\nEXISTING COMPANIES ALREADY IN LIST: {', '.join(oem_context)}\n"
            "- If you know any of these companies by a different name, use EXACTLY the name shown above\n"
            "- Do not add duplicates — only add companies genuinely not in this list\n"
        )

    _OEM_SCHEMA = """{
  "company_name": "exact company name",
  "country": "country name",
  "role": "OEM manufacturer | Brand owner | Contract manufacturer | Licensor",
  "market_share": "high|medium|low|unknown",
  "confidence": "high|medium|low",
  "notes": "brief note"
}"""

    def _oem_prompt(geo: str) -> str:
        return f"""You are a supply chain research assistant.
Product: {product_info.get('product_name')} ({product_info.get('industry')})

{oem_ev_note}

Find ALL known OEM manufacturers / brands that produce or sell this product or equivalent products.
Include the primary OEM already identified ({product_info.get('oem_manufacturer')}) plus any others.

ELIGIBILITY CRITERIA — a company must meet ALL of the following:
- Currently active and operating as of 2024
- Actively manufacturing {product_info.get('product_name')} within the last 5 years ({datetime.datetime.now().year - 5}–{datetime.datetime.now().year})
- Demonstrable manufacturing activity — not just distribution, reselling, or licensing
- Exclude any company that has ceased production or exited the market before {datetime.datetime.now().year - 5}
- GEOGRAPHY: {geo}

CRITICAL COMPANY NAME RULES:
- Use the shortest globally recognised name only (e.g. "Samsung" not "Samsung Electronics Co. Ltd.")
- No legal suffixes: drop Inc, Ltd, LLC, GmbH, Co., Corp, Group, Holdings
- Use English names only
- Be consistent: if a company is known by an acronym (TSMC, BASF, ABB), use the acronym
- If a company name is followed by (), check the bracket contents — if it is another name for the same company, drop the brackets
{existing_note}
CRITICAL: Return ONLY a raw JSON array starting with [ and ending with ].
Do NOT return a single object. Do NOT wrap in markdown. Do NOT add explanation text.
Schema for each element:
{_OEM_SCHEMA}

List 2–8 significant OEMs as SEPARATE array elements."""

    gemini_oems, deepseek_oems = _run_parallel(
        _oem_prompt("Only include companies headquartered or primarily operating OUTSIDE China (do NOT include mainland China, Hong Kong, or Taiwan)"),
        _oem_prompt("Only include companies headquartered or primarily operating IN China (mainland China, Hong Kong, and Taiwan only)"),
        1000, "OEMs"
    )

    oem_list = _merge_dedup(gemini_oems, deepseek_oems)
    push("status", message=f"  📋 Merged: {len(oem_list)} OEM(s) after dedup")

    for oem in oem_list:
        oem.setdefault("market_share_pct",    "")
        oem.setdefault("market_share_source", "")
        oem.setdefault("market_rank",         99)
        name = oem.get("company_name", "")
        if name and name not in oem_context:
            oem_context.append(name)
        c = get_coords(oem.get("country", ""))
        if c: oem["lat"], oem["lng"] = c

    primary_oem = product_info.get("oem_manufacturer", "")
    if primary_oem and primary_oem.lower() not in ("unknown", ""):
        existing_names = [o.get("company_name", "").lower() for o in oem_list]
        if not any(primary_oem.lower() in n or n in primary_oem.lower() for n in existing_names):
            pc = get_coords(product_info.get("oem_country", ""))
            pe = {
                "company_name": primary_oem, "country": product_info.get("oem_country", ""),
                "role": "OEM manufacturer", "market_share": "high", "confidence": "high",
                "notes": "Primary identified manufacturer",
                "market_share_pct": "", "market_share_source": "", "market_rank": 99,
            }
            if pc: pe["lat"], pe["lng"] = pc
            oem_list.insert(0, pe)

    push("oem_discovered", oems=oem_list)
    push("status", message=f"✅ Found {len(oem_list)} OEM manufacturer(s)")

    supply_chain = {"product": product_info, "oems": oem_list, "tiers": {}}

    oem_names = [
        {"name": o.get("company_name", ""), "confidence": o.get("confidence", "")}
        for o in oem_list if o.get("company_name", "").lower() not in ("unknown", "")
    ]
    if not oem_names:
        oem_names = [{"name": product_info.get("oem_manufacturer", product_input), "confidence": "high"}]

    previous_parents = [
        {"name": n["name"], "oem_root": n["name"], "confidence": n["confidence"], "parent": n["name"]}
        for n in oem_names
    ]

    def tier_note(tier_num: int) -> str:
        names = collected_tiers.get(f"tier_{tier_num}", [])
        return (
            f"The following companies are already in the tier {tier_num} list: {', '.join(names)}. "
            "If you know any of these companies by a different name, use EXACTLY the name shown above."
            if names else ""
        )

    _TIER_SCHEMA = """{
  "company_name": "exact company name",
  "country": "country name",
  "supplies_to": "PARENT_PLACEHOLDER",
  "oem_root": "OEM_ROOT_PLACEHOLDER",
  "components_supplied": ["component1", "component2"],
  "confidence": "high|medium|low",
  "source_hint": "brief source note"
}"""

    # ── Step 3: Tier Discovery ────────────────────────────────────────────────
    for tier_num in range(1, depth + 1):
        push("status", message=f"🏭 Tier-{tier_num}: Gemini (non-China) + DeepSeek (China) in parallel…")
        tier_suppliers = []
        next_parents   = []

        if tier_num > 1:
            lookup_dict, filtered = {}, []
            for entry in previous_parents:
                parent = entry.get("parent", "")
                count  = lookup_dict.get(parent, 0)
                if count < 4:
                    lookup_dict[parent] = count + 1
                    filtered.append(entry)
            previous_parents = filtered

        for parent_info in previous_parents:
            parent   = parent_info["name"]
            oem_root = parent_info["oem_root"]
            if not parent or parent.lower() == "unknown":
                continue

            query = (
                f"{parent} direct Tier 1 suppliers manufacturers components parts"
                if tier_num == 1 else
                f"{parent} suppliers raw materials components manufacturers supply chain"
            )

            try:
                ev_results = web_search(query, n=10, max_scrape=4)
            except DDGRateLimitError as e:
                push("error", message=f"🚫 DDG rate limit at Tier-{tier_num} [{parent}] — aborting. ({e})")
                return

            ev_text       = _format_evidence(ev_results)
            evidence_note = f"Web research findings:\n{ev_text}" if ev_text else "No live web search — use your training knowledge."
            push("status", message=f"  {'🌐' if ev_text else '🧠'} [{oem_root}] → {parent}: searching suppliers…")

            tier_schema = _TIER_SCHEMA.replace("PARENT_PLACEHOLDER", parent).replace("OEM_ROOT_PLACEHOLDER", oem_root)

            def _tier_prompt(geo: str) -> str:
                return f"""You are a supply chain research assistant.
Product: {product_info.get('product_name')} ({product_info.get('industry')})
OEM root: {oem_root}
Finding Tier-{tier_num} suppliers — companies that DIRECTLY supply components that are part of the manufacturing chain for {product_info.get('product_name')} to: {parent}

{evidence_note}
STRICT INCLUSION RULES — a supplier must meet ALL of the following:
- Supplies {parent} with something physically used in making "{product_info.get('product_name')}" or one of its key components
- Relevant specifically to the "{product_info.get('product_name')}" product line of {parent}
- Currently active as of {datetime.datetime.now().year}
- Direct supply relationship within the last 5 years — not distribution, reselling, or licensing
- GEOGRAPHY: {geo}

STRICT EXCLUSION RULES — do NOT include:
- General industrial inputs unrelated to "{product_info.get('product_name')}" (e.g. office supplies, IT services, packaging)
- Suppliers for {parent}'s other unrelated product lines
- Logistics, shipping, freight, financial, legal, or consulting providers
- Any company that does not contribute a physical material, chemical, or manufactured component to "{product_info.get('product_name')}"
- Recycling companies, waste processors, or end-of-life collectors
- Any company whose primary relationship with {parent} is as a customer, not a supplier
- Competitors that also manufacture "{product_info.get('product_name')}" — these are market rivals, not supply inputs
- OEM manufacturers of the same product — they are never valid tier suppliers to each other
- Internal subsidiaries or parent companies of {parent}

CRITICAL COMPANY NAME RULES:
- Use the shortest globally recognised name only
- No legal suffixes: drop Inc, Ltd, LLC, GmbH, Co., Corp, Group, Holdings
- Use English names only; use acronyms where standard (TSMC, BASF, ABB)
- Drop bracketed alternate names
{tier_note(tier_num)}
Return ONLY a valid JSON array (no markdown). Each element:
{tier_schema}

List 3–5 distinct real direct suppliers to {parent} specifically. Do NOT mix in suppliers of other companies."""

            g_found, d_found = _run_parallel(
                _tier_prompt("Focus exclusively on suppliers headquartered or operating OUTSIDE China (do NOT include mainland China, Hong Kong, or Taiwan)"),
                _tier_prompt("Focus exclusively on Chinese suppliers (mainland China, Hong Kong, and Taiwan only)"),
                1200, f"Tier-{tier_num} [{parent}]"
            )

            merged = _merge_dedup(g_found, d_found)
            for s in merged:
                s["oem_root"]    = oem_root
                s["supplies_to"] = parent
            tier_suppliers.extend(merged)

            for s in merged:
                sname = s.get("company_name", "").strip()
                if sname and sname.lower() != "unknown":
                    next_parents.append({
                        "name": sname, "oem_root": oem_root,
                        "confidence": s.get("confidence"), "parent": parent,
                    })

        previous_parents = next_parents

        seen, unique = set(), []
        for s in tier_suppliers:
            key = f"{s.get('company_name', '').strip().lower()}|{s.get('oem_root', '')}"
            if key and key not in seen:
                seen.add(key)
                c = get_coords(s.get("country", ""))
                if c: s["lat"], s["lng"] = c
                unique.append(s)

        for s in unique:
            name = s.get("company_name", "")
            collected_tiers.setdefault(f"tier_{tier_num}", [])
            if name and name not in collected_tiers[f"tier_{tier_num}"]:
                collected_tiers[f"tier_{tier_num}"].append(name)

        supply_chain["tiers"][f"tier_{tier_num}"] = unique
        push("tier_complete", tier=tier_num, suppliers=unique)
        push("status", message=f"✅ Tier-{tier_num}: {len(unique)} suppliers found")

        if not any(s.get("company_name") for s in unique):
            push("status", message=f"⚠️ No suppliers at Tier-{tier_num}, stopping.")
            break

    # ── Step 4: Executive Summary (Gemini only) ───────────────────────────────
    push("status", message="📊 Generating executive summary (Gemini)…")
    summary_prompt = f"""You are a supply chain analyst.
Supply chain data:
{json.dumps(supply_chain, indent=2)[:6000]}

Write a concise executive summary (150–200 words) covering:
- What the product is and its industry
- Key Tier-1 suppliers and their roles
- Geographic concentration or risks
- Overall supply chain complexity
Return plain text only, no markdown."""

    try:
        supply_chain["summary"] = call_ai(summary_prompt, "gemini", max_tokens=400).strip()
    except Exception as e:
        supply_chain["summary"] = f"Summary unavailable: {e}"

    supply_chain["provider"] = "agentic"
    push("complete", supply_chain=supply_chain)
