"""
verify.py — Supplier verification layer

Reads an exported agentic Excel file, runs three web-search + LLM checks
per supplier row, and returns an annotated Excel with five new columns:
  Verification Notes | Company Exists | Supply Ties | Correct Component Supplied | URLs

Registered in app.py via:
    from verify import verify_bp
    app.register_blueprint(verify_bp)
"""

import io
import json
import os
import tempfile
import threading
import uuid
from urllib.parse import quote_plus

import openpyxl
from flask import Blueprint, Response, jsonify, request, send_file
from openpyxl.styles import Font, PatternFill
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ── Import shared helpers from app.py ─────────────────────────────────────────
# These are imported at call-time (inside functions) to avoid circular imports
# since app.py imports nothing from verify.py at module level.

verify_bp = Blueprint("verify", __name__)

# Per-run queues: stream_id -> list of SSE event dicts
_verify_queues: dict = {}
_verify_queues_lock = threading.Lock()


# ── Selenium + Edge web search ───────────────────────────────────────────────

def playwright_search(query: str, n: int = 4, log_fn=None) -> list[dict]:
    """
    Search Bing headlessly via Playwright + Chromium and scrape the top n article pages.
    Returns a list of {"url": str, "content": str} dicts.
    Falls back to [] on any failure so callers degrade gracefully.
    log_fn(msg, is_error): optional callback to push log lines to the SSE queue.
    """
    def _log(msg, is_error=False):
        print(msg)
        if log_fn:
            log_fn(msg, is_error)

    results = []
    search_url = f"https://www.bing.com/search?q={quote_plus(query)}"
    SKIP_DOMAINS = ("bing.com", "microsoft.com", "msn.com", "go.microsoft.com")

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                ),
                locale="en-US",
            )
            page = context.new_page()

            # ── Step 1: fetch Bing search results page ────────────────────────
            try:
                page.goto(search_url, timeout=20000, wait_until="domcontentloaded")
                page.wait_for_selector("#b_results", timeout=8000)
            except PlaywrightTimeoutError as e:
                _log(f"[playwright] Bing results page failed for '{query}': {e}", is_error=True)
                browser.close()
                return []

            anchors = page.query_selector_all("#b_results h2 a, .b_algo h2 a")
            _log(f"[playwright] Raw anchors found: {len(anchors)}")

            urls = []
            for a in anchors:
                href = a.get_attribute("href") or ""
                if href.startswith("http") and not any(d in href for d in SKIP_DOMAINS):
                    urls.append(href)
                if len(urls) >= n:
                    break

            _log(f"[playwright] URLs after filter: {len(urls)}")

            # ── Step 2: scrape each article page ─────────────────────────────
            for url in urls:
                try:
                    page.goto(url, timeout=15000, wait_until="domcontentloaded")
                    text = page.inner_text("body")
                    text = " ".join(text.split())[:2000]
                    results.append({"url": url, "content": text})
                    _log(f"[playwright] Scraped: {url}")
                except Exception as e:
                    _log(f"[playwright] Failed to scrape {url}: {e}", is_error=True)

            browser.close()

    except Exception as e:
        _log(f"[playwright] Unexpected error for '{query}': {e}", is_error=True)

    _log(f"[playwright] Completed webcrawl for '{query}' — {len(results)} page(s) scraped")
    return results


# ── LLM verification helper ───────────────────────────────────────────────────

def _ai_judge(evidence: str, question: str, provider: str) -> dict:
    """
    Send evidence text + a yes/no question to the LLM.
    Returns {"answer": True|False, "notes": str, "urls": [str]}.
    """
    from app import call_ai, safe_parse_json

    prompt = f"""You are a supply chain fact-checker. Based ONLY on the web search evidence below,
answer the question. Do NOT use your training knowledge — only what the evidence says.

EVIDENCE:
{evidence[:4000]}

QUESTION: {question}

Return ONLY valid JSON (no markdown):
{{
  "answer": true or false,
  "notes": "one sentence explanation citing specific evidence, or 'No evidence found'"
}}

If the evidence does not clearly confirm, set answer to false."""

    try:
        raw = call_ai(prompt, provider, max_tokens=300)
        parsed = safe_parse_json(raw)
        if isinstance(parsed, dict) and "answer" in parsed:
            return {
                "answer": bool(parsed.get("answer", False)),
                "notes": str(parsed.get("notes", "")),
            }
    except Exception:
        pass
    return {"answer": False, "notes": "LLM call failed"}


def verify_supplier_row(company_name: str, supplies_to: str, components: str, provider: str, log_fn=None) -> dict:
    """
    Run three web-search + LLM checks for one supplier row.

    Returns:
        {
            "company_exists":     "Yes" | "No",
            "supply_ties":        "Yes" | "No",
            "correct_component":  "Yes" | "No",
            "notes":              str,
            "urls":               [str],
        }
    """
    all_notes = []
    all_urls: list[str] = []

    def _build_evidence(results: list[dict]) -> str:
        if not results:
            return "No search results returned."
        return "\n---\n".join(
            f"SOURCE: {r['url']}\nCONTENT: {r['content']}" for r in results
        )

    # ── Step 1: Company Exists ────────────────────────────────────────────────
    results1 = playwright_search(f'"{company_name}" manufacturer supplier', n=3, log_fn=log_fn)
    result1 = _ai_judge(
        _build_evidence(results1),
        f'Does "{company_name}" exist as a real company that manufactures or supplies industrial/commercial products?',
        provider,
    )
    company_exists = "Yes" if result1["answer"] else "No"
    if result1["notes"]:
        all_notes.append(f"[Exists] {result1['notes']}")
    if result1["answer"]:
        all_urls.extend(r["url"] for r in results1)

    # ── Step 2: Supply Ties ───────────────────────────────────────────────────
    results2 = playwright_search(f'"{company_name}" "{supplies_to}" supplier', n=4, log_fn=log_fn)
    result2 = _ai_judge(
        _build_evidence(results2),
        f'Does the evidence confirm a direct supply relationship where "{company_name}" supplies "{supplies_to}"?',
        provider,
    )
    supply_ties = "Yes" if result2["answer"] else "No"
    if result2["notes"]:
        all_notes.append(f"[Supply ties] {result2['notes']}")
    if result2["answer"]:
        all_urls.extend(r["url"] for r in results2)

    # ── Step 3: Correct Component Supplied ────────────────────────────────────
    component_query = components[:120] if components else "components"
    results3 = playwright_search(f'"{company_name}" manufactures "{component_query}"', n=4, log_fn=log_fn)
    result3 = _ai_judge(
        _build_evidence(results3),
        f'Does the evidence confirm that "{company_name}" manufactures or produces "{component_query}"?',
        provider,
    )
    correct_component = "Yes" if result3["answer"] else "No"
    if result3["notes"]:
        all_notes.append(f"[Component] {result3['notes']}")
    if result3["answer"]:
        all_urls.extend(r["url"] for r in results3)

    # Deduplicate URLs while preserving order
    seen_urls: set[str] = set()
    unique_urls = []
    for u in all_urls:
        if u not in seen_urls:
            seen_urls.add(u)
            unique_urls.append(u)

    return {
        "company_exists": company_exists,
        "supply_ties": supply_ties,
        "correct_component": correct_component,
        "notes": " | ".join(all_notes),
        "urls": unique_urls,
    }


# ── Workbook annotation ───────────────────────────────────────────────────────

VERIFY_HEADERS = [
    "Verification Notes",
    "Company Exists",
    "Supply Ties",
    "Correct Component Supplied",
    "URLs",
]

_HDR_FILL = PatternFill("solid", fgColor="1A237E")  # dark navy
_HDR_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
_YES_FILL = PatternFill("solid", fgColor="C8E6C9")   # light green
_NO_FILL  = PatternFill("solid", fgColor="FFCDD2")   # light red


def _find_col(ws, header_name: str) -> int | None:
    """Return 1-based column index for a header in row 1, or None."""
    for cell in ws[1]:
        if cell.value and str(cell.value).strip().lower() == header_name.lower():
            return cell.column
    return None


def annotate_workbook(wb: openpyxl.Workbook, stream_id: str, provider: str) -> openpyxl.Workbook:
    """
    Iterate every tier sheet in the workbook, run verification per supplier row,
    and append the 5 verification columns. Pushes SSE progress events to the queue.
    """
    def push(event: dict):
        with _verify_queues_lock:
            if stream_id in _verify_queues:
                _verify_queues[stream_id].append(event)

    # Identify tier sheets (written by write_tier_sheet — names contain "tier")
    tier_sheets = [ws for ws in wb.worksheets if "tier" in ws.title.lower()]

    if not tier_sheets:
        push({"type": "status", "message": "⚠️ No tier sheets found in this workbook."})
        return wb

    # Count total supplier rows for progress reporting
    total_rows = 0
    for ws in tier_sheets:
        col_company = _find_col(ws, "Company Name")
        if col_company:
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[col_company - 1]:
                    total_rows += 1

    push({"type": "start", "total": total_rows})
    push({"type": "timer_start"})
    processed = 0

    def log_fn(msg, is_error=False):
        push({"type": "log", "message": msg, "is_error": is_error})

    for ws in tier_sheets:
        col_company   = _find_col(ws, "Company Name")
        col_supplies  = _find_col(ws, "Supplies To")
        col_component = _find_col(ws, "Components Supplied")

        if not col_company:
            push({"type": "status", "message": f"⚠️ Skipping sheet '{ws.title}' — no 'Company Name' column"})
            continue

        # Find where to start appending (first empty column after last header)
        last_col = ws.max_column
        verify_start_col = last_col + 1

        # Write verification column headers
        for i, hdr in enumerate(VERIFY_HEADERS):
            cell = ws.cell(row=1, column=verify_start_col + i, value=hdr)
            cell.fill = _HDR_FILL
            cell.font = _HDR_FONT
            ws.column_dimensions[cell.column_letter].width = 22 if i < 4 else 50

        # Process each data row
        for row_idx in range(2, ws.max_row + 1):
            company_name = ws.cell(row=row_idx, column=col_company).value
            if not company_name:
                continue

            supplies_to = ws.cell(row=row_idx, column=col_supplies).value if col_supplies else ""
            components  = ws.cell(row=row_idx, column=col_component).value if col_component else ""

            processed += 1
            push({
                "type": "progress",
                "row": processed,
                "total": total_rows,
                "message": f"[{processed}/{total_rows}] Verifying: {company_name}",
            })

            result = verify_supplier_row(
                company_name=str(company_name),
                supplies_to=str(supplies_to or ""),
                components=str(components or ""),
                provider=provider,
                log_fn=log_fn,
            )

            # Write the 5 columns
            values = [
                result["notes"],
                result["company_exists"],
                result["supply_ties"],
                result["correct_component"],
                ", ".join(result["urls"]),
            ]
            for i, val in enumerate(values):
                cell = ws.cell(row=row_idx, column=verify_start_col + i, value=val)
                # Colour Yes/No cells
                if val in ("Yes", "No"):
                    cell.fill = _YES_FILL if val == "Yes" else _NO_FILL

    push({"type": "status", "message": f"✅ Verification complete — {processed} row(s) processed."})
    return wb


# ── Background worker ─────────────────────────────────────────────────────────

def _run_verification(stream_id: str, file_bytes: bytes, filename: str, provider: str):
    """Background thread: load workbook, annotate, save, signal file_ready."""
    from app import _export_files, _export_files_lock

    def push(event: dict):
        with _verify_queues_lock:
            if stream_id in _verify_queues:
                _verify_queues[stream_id].append(event)

    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        wb = annotate_workbook(wb, stream_id, provider)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        # Save to temp file for download
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(buf.read())
            tmp_path = tmp.name

        out_filename = filename.replace(".xlsx", "_verified.xlsx")
        file_id = uuid.uuid4().hex
        with _export_files_lock:
            _export_files[file_id] = {"path": tmp_path, "filename": out_filename}

        push({"type": "timer_stop"})
        push({"type": "file_ready", "file_id": file_id, "filename": out_filename})

    except Exception as e:
        push({"type": "error", "message": str(e)})
    finally:
        push({"type": "done"})


# ── Routes ────────────────────────────────────────────────────────────────────

@verify_bp.route("/dev/verify")
def verify_page():
    """Upload page — black terminal UI matching the rest of the dev tools."""
    html = """<!DOCTYPE html>
<html>
<head>
  <title>Supplier Verification</title>
  <style>
    body { font-family: monospace; padding: 2rem; background: #111; color: #eee; }
    h2 span { color: #e8ff47; }
    input[type=file] { margin: 1rem 0; display: block; color: #eee; }
    button { background: #333; color: #eee; border: 1px solid #555; padding: .5rem 1.5rem; cursor: pointer; font-family: monospace; }
    button:hover { background: #444; }
    #log div { margin: 2px 0; font-size: .85rem; line-height: 1.6; }
    .done    { color: #4caf50; }
    .error   { color: #f44336; }
    .log-line { color: #888; font-size: .78rem; }
    .log-err  { color: #e57373; font-size: .78rem; }
    select { background: #222; color: #eee; border: 1px solid #555; padding: .3rem; font-family: monospace; margin-bottom: 1rem; }
    label  { font-size: .9rem; color: #aaa; }
    #timer { display: none; font-size: .85rem; color: #e8ff47; margin-top: .75rem; letter-spacing: .05em; }
    #timer span { font-weight: bold; }
  </style>
</head>
<body>
  <h2>Supplier <span>Verification</span></h2>
  <p>Upload an Excel file exported from the agentic pipeline. Each supplier row will be
     verified via web search across 3 dimensions.</p>

  <label>LLM Provider for judgement:</label><br>
  <select id="provider">
    <option value="gemini">Gemini (recommended — cheapest)</option>
    <option value="anthropic">Anthropic</option>
  </select>

  <input type="file" id="fileInput" accept=".xlsx">
  <button onclick="startVerification()">Upload &amp; Verify</button>

  <div id="timer">⏱ Elapsed: <span id="timerVal">0.0s</span></div>
  <div id="log" style="margin-top:1.5rem"></div>

  <script>
    const log = document.getElementById('log');
    const append = (msg, cls) => {
      const d = document.createElement('div');
      if (cls) d.className = cls;
      d.textContent = msg;
      log.appendChild(d);
      d.scrollIntoView();
    };

    let timerInterval = null;
    let timerStart = null;

    function startTimer() {
      timerStart = Date.now();
      const disp = document.getElementById('timer');
      const val  = document.getElementById('timerVal');
      disp.style.display = 'block';
      val.textContent = '0.0s';
      if (timerInterval) clearInterval(timerInterval);
      timerInterval = setInterval(() => {
        val.textContent = ((Date.now() - timerStart) / 1000).toFixed(1) + 's';
      }, 100);
    }

    function stopTimer() {
      if (timerInterval) { clearInterval(timerInterval); timerInterval = null; }
      if (!timerStart) return;
      const elapsed = ((Date.now() - timerStart) / 1000).toFixed(1);
      document.getElementById('timerVal').textContent = elapsed + 's ✓';
    }

    async function startVerification() {
      const file = document.getElementById('fileInput').files[0];
      if (!file) { append('⚠️ Please select an Excel file first.'); return; }

      const provider = document.getElementById('provider').value;
      append(`Uploading ${file.name}…`);

      const fd = new FormData();
      fd.append('file', file);
      fd.append('provider', provider);

      let streamId;
      try {
        const res = await fetch('/dev/verify/run', { method: 'POST', body: fd });
        const data = await res.json();
        if (data.error) { append('Error: ' + data.error, 'error'); return; }
        streamId = data.stream_id;
      } catch (e) {
        append('Upload failed: ' + e, 'error'); return;
      }

      append('Upload complete — starting verification…');
      const evtSource = new EventSource(`/dev/verify/stream/${streamId}`);

      evtSource.onmessage = (e) => {
        const msg = JSON.parse(e.data);
        if (msg.type === 'timer_start') {
          startTimer();
        } else if (msg.type === 'timer_stop') {
          stopTimer();
        } else if (msg.type === 'start') {
          append(`Processing ${msg.total} supplier row(s)…`);
        } else if (msg.type === 'progress') {
          append(msg.message);
        } else if (msg.type === 'status') {
          append(msg.message);
        } else if (msg.type === 'log') {
          append(msg.message, msg.is_error ? 'log-err' : 'log-line');
        } else if (msg.type === 'file_ready') {
          append(`✓ Done — downloading ${msg.filename}`, 'done');
          const a = document.createElement('a');
          a.href = `/dev/verify/download/${msg.file_id}`;
          a.download = msg.filename;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          evtSource.close();
        } else if (msg.type === 'error') {
          stopTimer();
          append('Error: ' + msg.message, 'error');
          evtSource.close();
        } else if (msg.type === 'done') {
          evtSource.close();
        }
      };

      evtSource.onerror = () => {
        stopTimer();
        append('Connection lost.', 'error');
        evtSource.close();
      };
    }
  </script>
</body>
</html>"""
    return Response(html, mimetype="text/html")


@verify_bp.route("/dev/verify/run", methods=["POST"])
def verify_run():
    """Accept Excel upload, start background verification, return stream_id."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        return jsonify({"error": "Only .xlsx files are supported"}), 400

    provider = request.form.get("provider", "gemini")
    file_bytes = f.read()
    filename = f.filename
    stream_id = uuid.uuid4().hex

    with _verify_queues_lock:
        _verify_queues[stream_id] = []

    thread = threading.Thread(
        target=_run_verification,
        args=(stream_id, file_bytes, filename, provider),
        daemon=True,
    )
    thread.start()

    return jsonify({"stream_id": stream_id})


@verify_bp.route("/dev/verify/stream/<stream_id>")
def verify_stream(stream_id):
    """SSE stream for a running verification job."""
    import time as _time

    def generate():
        deadline = _time.time() + 3600  # 1-hour max
        while _time.time() < deadline:
            with _verify_queues_lock:
                queue = _verify_queues.get(stream_id, [])
                events, _verify_queues[stream_id] = queue[:], []

            for event in events:
                yield f"data: {json.dumps(event)}\n\n"
                if event.get("type") == "done":
                    with _verify_queues_lock:
                        _verify_queues.pop(stream_id, None)
                    return

            _time.sleep(0.5)

    return Response(generate(), mimetype="text/event-stream", headers={
        "Cache-Control": "no-cache",
        "X-Accel-Buffering": "no",
    })


@verify_bp.route("/dev/verify/download/<file_id>")
def verify_download(file_id):
    """Download the annotated Excel file."""
    from app import _export_files, _export_files_lock

    with _export_files_lock:
        entry = _export_files.pop(file_id, None)

    if not entry:
        return jsonify({"error": "File not found or already downloaded"}), 404

    path     = entry["path"]
    filename = entry["filename"]

    response = send_file(
        path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )

    @response.call_on_close
    def cleanup():
        try:
            os.remove(path)
        except OSError:
            pass

    return response
