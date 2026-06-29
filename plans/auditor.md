# Plan: Two-Stage Auditor Verification

## Context
The current verification pipeline uses a single LLM (user's choice: Gemini or Anthropic) to
judge each supplier row. This creates a single point of failure — if the primary model is
wrong or over-confident, there is no check on its output. The user wants a second model (the
one NOT chosen) to act as a skeptical auditor: it reviews Stage 1's verdicts against the same
web evidence and actively tries to debunk them. This is opt-in via a checkbox. Auditor results
appear as additional columns in the output Excel.

---

## Key Design Decisions

| Decision | Choice | Reason |
|---|---|---|
| Auditor evidence source | Reuse Stage 1's evidence block | No extra DDG/scraping cost; auditor evaluates whether the existing evidence is sufficient |
| Auditor model | Whichever of Gemini/Anthropic was NOT chosen by user | Simple swap; both have API keys configured |
| Auditor stance | Adversarial — challenge Stage 1 unless evidence clearly supports it | A confirming auditor adds no value |
| Per-check result values | "Confirmed" / "Disputed" / "Unconfirmed" | Confirmed = agrees + evidence; Disputed = contradicts; Unconfirmed = evidence insufficient |
| Overall verdict | "Confirmed" / "Disputed" / "Partial" | Disputed if any check Disputed; Partial if mix; Confirmed if all Confirmed |
| Concurrency | Auditor call added after Stage 1 per row, inside same ThreadPoolExecutor worker | No extra threads needed; adds ~1 LLM call per row wall-clock time |

---

## Files to Change

| File | Change |
|---|---|
| `verify.py` | Add `_ai_audit_row()`, update `verify_supplier_row()` + `annotate_workbook()` + upload route + HTML |

No other files need changing — Excel column writing is done inline in `annotate_workbook()`.

---

## Backend Changes (`verify.py`)

### 1. New function `_ai_audit_row()`

Add after `_ai_judge_all()` (around line 130):

```python
def _ai_audit_row(
    evidence: str,
    company_name: str,
    supplies_to: str,
    components: str,
    stage1: dict,           # result dict from _ai_judge_all
    audit_provider: str,
    log_fn=None,
) -> dict:
```

**Prompt design** — adversarial, not neutral:
- Tells the auditor it is reviewing a previous model's claims
- Shows Stage 1's verdicts and reasoning verbatim
- Shows the same evidence block
- Instructs: "Mark as Confirmed ONLY if evidence clearly and directly supports the verdict.
  Mark as Disputed if evidence contradicts. Mark as Unconfirmed if evidence is weak or absent."

**Expected JSON response:**
```json
{
  "audit_company_exists":   "Confirmed" | "Disputed" | "Unconfirmed",
  "audit_supply_ties":      "Confirmed" | "Disputed" | "Unconfirmed",
  "audit_correct_component":"Confirmed" | "Disputed" | "Unconfirmed",
  "audit_notes_exists":     "reasoning string",
  "audit_notes_supply":     "reasoning string",
  "audit_notes_component":  "reasoning string"
}
```

Uses `_call_ai_json()` with `required_key="audit_company_exists"` and same 3-retry backoff.

**Audit verdict** (computed in Python after the call, not by the LLM):
- Any check = "Disputed" → overall = "Disputed"
- All checks = "Confirmed" → overall = "Confirmed"
- Otherwise → "Partial"

---

### 2. Update `verify_supplier_row()`

Add two parameters: `use_auditor: bool = False`, `audit_provider: str = ""`

After the existing Stage 1 call completes, if `use_auditor`:
```python
audit = _ai_audit_row(evidence, company_name, supplies_to, components, stage1, audit_provider, log_fn)
result.update(audit)  # merges audit_* fields into the return dict
```

Return dict gains: `audit_company_exists`, `audit_supply_ties`, `audit_correct_component`,
`audit_notes_exists`, `audit_notes_supply`, `audit_notes_component`, `audit_verdict`.
Returns empty strings for all if `use_auditor=False`.

---

### 3. Update `annotate_workbook()`

Add parameters `use_auditor: bool`, `audit_provider: str`.

Pass through to `verify_supplier_row()` in the ThreadPoolExecutor submit call.

**When writing results to each row**, after the existing 11 verification columns, if `use_auditor`:
- Write 5 auditor columns (with colour-coding):
  - `Audit: Company Exists` — green=Confirmed, red=Disputed, yellow=Unconfirmed
  - `Audit: Supply Ties Exist` — same colours
  - `Audit: Correct Component` — same colours
  - `Audit Notes` — combined auditor reasoning (same format as Verification Notes: `[Exists] ... | [Supply] ... | [Component] ...`)
  - `Audit Verdict` — green=Confirmed, red=Disputed, yellow=Partial

Header row: include `Audit Provider: {audit_provider_name}` in the column header or as a merged header above.

---

### 4. Update upload route (`verify_run` or equivalent)

Read new form fields:
```python
use_auditor   = request.form.get('use_auditor', 'false') == 'true'
audit_provider = 'anthropic' if provider == 'gemini' else 'gemini'
```

Pass `use_auditor` and `audit_provider` to `annotate_workbook()`.

---

### 5. Update HTML in `verify.py`

Add checkbox below the provider selector:
```html
<label class="audit-check">
  <input type="checkbox" id="useAuditor">
  Include Auditor Review
  <span id="auditorLabel" style="color:#aaa;font-size:.8rem">(Gemini will cross-check)</span>
</label>
```

JS: update the auditor label when the provider dropdown changes:
```javascript
document.getElementById('provider').addEventListener('change', function() {
  const auditor = this.value === 'gemini' ? 'Anthropic' : 'Gemini';
  document.getElementById('auditorLabel').textContent = `(${auditor} will cross-check)`;
});
```

In `startVerification()` / `processFile()`, include in the FormData:
```javascript
fd.append('use_auditor', document.getElementById('useAuditor').checked ? 'true' : 'false');
```

---

## Excel Output (5 new auditor columns, appended only when auditor enabled)

| Column | Values | Colour |
|---|---|---|
| `Audit: Company Exists` | Confirmed / Disputed / Unconfirmed | Green / Red / Yellow |
| `Audit: Supply Ties Exist` | Confirmed / Disputed / Unconfirmed | Green / Red / Yellow |
| `Audit: Correct Component` | Confirmed / Disputed / Unconfirmed | Green / Red / Yellow |
| `Audit Notes` | `[Exists] … \| [Supply] … \| [Component] …` | White |
| `Audit Verdict` | Confirmed / Disputed / Partial | Green / Red / Yellow |

Header fill: a distinct purple (e.g. `#4A148C`) to visually separate from Stage 1 headers.

---

## Performance Notes

- Auditor adds **1 extra LLM call per row** (no extra web searches)
- Gemini is fast (~2–5 s); Anthropic is ~3–8 s
- With the existing 5-worker ThreadPoolExecutor and rows processed in parallel, the extra call
  happens sequentially within each worker's row — total wall-clock time roughly doubles per
  row, but parallelism across rows is unchanged
- Auditor is opt-in, so the default path is unchanged

---

## Verification (how to test)

1. Run the full pipeline on a product to produce a `.xlsx`
2. Open Verify Suppliers page, upload the file, select a provider, **check "Include Auditor Review"**
3. Confirm SSE log shows two-phase messages per row: e.g. `[Gemini] ✓ TSMC` then `[Anthropic Audit] ✓ TSMC`
4. Download the annotated Excel; confirm 5 auditor columns appear after the existing verification columns
5. Find a row where Stage 1 said "Yes" — check if auditor says "Confirmed" or "Disputed"
6. Run without the checkbox checked — confirm output has no auditor columns and runs at normal speed
