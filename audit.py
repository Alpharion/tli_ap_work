"""
audit.py — Two-stage adversarial auditor for supplier verification.

The auditor receives Stage 1's verdicts and the same web-evidence block,
then challenges those verdicts. Import into verify.py via:

    from audit import ai_audit_row, write_audit_headers, write_audit_row, AUDIT_COL_COUNT
"""

import time

from openpyxl.styles import Alignment, Font, PatternFill

from ai import call_ai, safe_parse_json

# ── Fills ─────────────────────────────────────────────────────────────────────

AUDIT_HDR_FILL  = PatternFill("solid", fgColor="4A148C")  # deep purple header
AUDIT_CONF_FILL = PatternFill("solid", fgColor="C8E6C9")  # green  — Confirmed
AUDIT_DISP_FILL = PatternFill("solid", fgColor="FFCDD2")  # red    — Disputed
AUDIT_UNCF_FILL = PatternFill("solid", fgColor="FFF9C4")  # yellow — Unconfirmed
AUDIT_HDR_FONT  = Font(name="Calibri", bold=True, color="FFFFFF", size=10)

AUDIT_COL_COUNT = 6  # columns written per row by write_audit_row()

_VALID_VERDICTS = {"Confirmed", "Disputed", "Unconfirmed"}
_WRAP = Alignment(wrap_text=True, vertical="top")


# ── Internal helpers (re-implemented here to avoid circular imports) ──────────

def _make_logger(log_fn):
    def _log(msg, is_error=False):
        print(msg)
        if log_fn:
            log_fn(msg, is_error)
    return _log


def _call_ai_json(prompt: str, provider: str, max_tokens: int,
                  required_key: str, log) -> dict | None:
    for attempt in range(1, 4):
        try:
            raw    = call_ai(prompt, provider, max_tokens=max_tokens)
            parsed = safe_parse_json(raw)
            if isinstance(parsed, dict) and required_key in parsed:
                return parsed
            log(f"[llm] Attempt {attempt}: response missing '{required_key}' — retrying", True)
        except Exception as e:
            log(f"[llm] Attempt {attempt} failed: {e}", True)
        if attempt < 3:
            time.sleep(5 * attempt)
    log("[llm] All 3 attempts failed", True)
    return None


# ── Fill helper ───────────────────────────────────────────────────────────────

def audit_fill(val: str) -> PatternFill:
    """Map an audit verdict string to its cell fill colour."""
    return {"Confirmed": AUDIT_CONF_FILL, "Disputed": AUDIT_DISP_FILL}.get(val, AUDIT_UNCF_FILL)


# ── Column metadata ───────────────────────────────────────────────────────────

def audit_col_headers(audit_provider: str) -> list[tuple[str, int]]:
    """Return [(header_name, column_width), ...] for the 6 audit columns."""
    label = audit_provider.capitalize()
    return [
        (f"Audit: Company Exists",           16),
        (f"Audit: Supply Ties Exist",         18),
        (f"Audit: Correct Component",         22),
        (f"Audit Notes [{label}]",            55),
        (f"Audit Verdict [{label}]",          18),
        (f"Audit Verdict Reason [{label}]",   45),
    ]


# ── Sheet helpers ─────────────────────────────────────────────────────────────

def write_audit_headers(ws, start_col: int, audit_provider: str) -> None:
    """Write the purple header row for the 6 audit columns starting at start_col."""
    for i, (hdr, width) in enumerate(audit_col_headers(audit_provider)):
        cell = ws.cell(row=1, column=start_col + i, value=hdr)
        cell.fill = AUDIT_HDR_FILL
        cell.font = AUDIT_HDR_FONT
        ws.column_dimensions[cell.column_letter].width = width


def write_audit_row(ws, row_idx: int, start_col: int, result: dict) -> None:
    """Write the 6 audit result cells for a single data row."""
    # Per-check verdict cells (columns 0–2)
    for k, key in enumerate(["audit_company_exists", "audit_supply_ties", "audit_correct_component"]):
        val = result.get(key, "")
        c   = ws.cell(row=row_idx, column=start_col + k, value=val)
        c.alignment = _WRAP
        if val:
            c.fill = audit_fill(val)

    # Combined notes (column 3)
    parts = []
    if result.get("audit_notes_exists"):
        parts.append(f"[Exists] {result['audit_notes_exists']}")
    if result.get("audit_notes_supply"):
        parts.append(f"[Supply] {result['audit_notes_supply']}")
    if result.get("audit_notes_component"):
        parts.append(f"[Component] {result['audit_notes_component']}")
    c = ws.cell(row=row_idx, column=start_col + 3, value=" | ".join(parts))
    c.alignment = _WRAP

    # Overall verdict (column 4) — Partial maps to yellow (Unconfirmed colour)
    verdict_val = result.get("audit_verdict", "")
    c = ws.cell(row=row_idx, column=start_col + 4, value=verdict_val)
    c.alignment = _WRAP
    if verdict_val:
        c.fill = audit_fill(verdict_val if verdict_val != "Partial" else "Unconfirmed")

    # Verdict reason (column 5)
    c = ws.cell(row=row_idx, column=start_col + 5, value=result.get("audit_verdict_reason", ""))
    c.alignment = _WRAP


# ── Core audit function ───────────────────────────────────────────────────────

def ai_audit_row(
    evidence: str,
    company_name: str,
    supplies_to: str,
    components: str,
    stage1: dict,
    audit_provider: str,
    log_fn=None,
) -> dict:
    """
    Adversarial second-pass: challenge Stage 1 verdicts using the same evidence block.

    stage1 is the raw verdict dict from _ai_judge_all() — keys:
        company_exists (bool), supply_ties (bool), correct_component (bool),
        notes_exists, notes_supply, notes_component

    Returns a dict of audit_* fields ready to be merged into the
    verify_supplier_row() result via result.update(audit).
    """
    _log = _make_logger(log_fn)

    stage1_summary = (
        f"- Company Exists: {stage1['company_exists']} — \"{stage1['notes_exists']}\"\n"
        f"- Supply Ties Exist: {stage1['supply_ties']} — \"{stage1['notes_supply']}\"\n"
        f"- Correct Component Supplied: {stage1['correct_component']} — \"{stage1['notes_component']}\""
    )

    prompt = f"""You are a skeptical supply chain auditor. A previous AI model assessed a supplier and you must challenge its conclusions.

SUPPLIER UNDER REVIEW:
- Company: {company_name}
- Supplies to: {supplies_to}
- Components: {components}

STAGE 1 VERDICT:
{stage1_summary}

EVIDENCE (same evidence used by Stage 1):
{evidence[:5000]}

INSTRUCTIONS: Audit each of Stage 1's three verdicts. Be adversarial — look for weaknesses, gaps, or overreach.

For each verdict use exactly one of these three values:
- "Confirmed": the evidence clearly and directly supports Stage 1's claim
- "Disputed": the evidence contradicts Stage 1's claim, or Stage 1 made an unjustified inference
- "Unconfirmed": evidence is weak, absent, or ambiguous — Stage 1 may be right but it cannot be verified from this evidence

Return ONLY valid JSON (no markdown):
{{
  "audit_company_exists":    "Confirmed" or "Disputed" or "Unconfirmed",
  "audit_supply_ties":       "Confirmed" or "Disputed" or "Unconfirmed",
  "audit_correct_component": "Confirmed" or "Disputed" or "Unconfirmed",
  "audit_notes_exists":      "one sentence explaining your audit decision",
  "audit_notes_supply":      "one sentence explaining your audit decision",
  "audit_notes_component":   "one sentence explaining your audit decision"
}}"""

    _log(f"[audit] Sending to {audit_provider} for '{company_name}'")
    parsed = _call_ai_json(prompt, audit_provider, 500, "audit_company_exists", _log)

    def _safe(val):
        return val if val in _VALID_VERDICTS else "Unconfirmed"

    if parsed:
        ace = _safe(parsed.get("audit_company_exists"))
        ast = _safe(parsed.get("audit_supply_ties"))
        acc = _safe(parsed.get("audit_correct_component"))
        vals = {ace, ast, acc}

        if "Disputed" in vals:
            verdict = "Disputed"
        elif vals == {"Confirmed"}:
            verdict = "Confirmed"
        elif vals == {"Unconfirmed"}:
            verdict = "Unconfirmed"
        else:
            verdict = "Partial"   # mixed: some Confirmed, some Unconfirmed

        checks = [("Company Exists", ace), ("Supply Ties Exist", ast), ("Correct Component", acc)]
        if verdict == "Confirmed":
            verdict_reason = "All 3 checks confirmed by evidence."
        elif verdict == "Disputed":
            disputed = [name for name, val in checks if val == "Disputed"]
            verdict_reason = f"Disputed on: {', '.join(disputed)}."
        elif verdict == "Unconfirmed":
            verdict_reason = "Evidence insufficient to confirm any of the 3 checks."
        else:
            verdict_reason = "; ".join(f"{name}: {val}" for name, val in checks) + "."

        _log(f"[audit] Result for '{company_name}' — exists={ace}, supply={ast}, component={acc} → {verdict}")
        return {
            "audit_company_exists":    ace,
            "audit_supply_ties":       ast,
            "audit_correct_component": acc,
            "audit_notes_exists":      str(parsed.get("audit_notes_exists",   "")),
            "audit_notes_supply":      str(parsed.get("audit_notes_supply",   "")),
            "audit_notes_component":   str(parsed.get("audit_notes_component","")),
            "audit_verdict":           verdict,
            "audit_verdict_reason":    verdict_reason,
        }

    _log("[audit] All attempts failed — marking Unconfirmed", True)
    return {
        "audit_company_exists":    "Unconfirmed",
        "audit_supply_ties":       "Unconfirmed",
        "audit_correct_component": "Unconfirmed",
        "audit_notes_exists":      "Audit LLM call failed",
        "audit_notes_supply":      "Audit LLM call failed",
        "audit_notes_component":   "Audit LLM call failed",
        "audit_verdict":           "Unconfirmed",
        "audit_verdict_reason":    "Audit LLM call failed.",
    }
