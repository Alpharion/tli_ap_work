"""
Supplier Chain Mapper — Multi-provider AI backend
"""

import io
import json
import os
import re
import tempfile
import threading
import time
import uuid

from dotenv import load_dotenv
load_dotenv(override=True)

from flask import Flask, Response, jsonify, render_template, request, send_file
from flask_cors import CORS

from ai import available_providers
from excel_export import build_bulk_workbook
from pipeline import run_pipeline
from report_export import report_bp
from verify import verify_bp

app = Flask(__name__)
CORS(app)
app.register_blueprint(verify_bp)
app.register_blueprint(report_bp)

# ── Configuration ─────────────────────────────────────────────────────────────

PRODUCTS_TO_TEST = [
    "Lithium ion battery 18650 cylindrical"
]
DEPTH    = 2        # supply chain depth per product (1–3)
PROVIDER = "gemini" # anthropic | openai | gemini | deepseek

# ── Export file store (shared with verify.py via at-call-time import) ─────────

_export_files      = {}
_export_files_lock = threading.Lock()

# ── Helpers ───────────────────────────────────────────────────────────────────

def sse(data: dict) -> str:
    return f"data: {json.dumps(data)}\n\n"

def safe_filename(product_input: str, index: int) -> str:
    clean = re.sub(r'[\\/*?:"<>|]', "", product_input).strip()[:50]
    return f"{index:02d}_{clean}.xlsx"

def safe_sheet_name(product_input: str, index: int) -> str:
    clean = re.sub(r'[\\/*?\[\]:]', "", product_input).strip()[:26]
    return f"{index}_{clean}"


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/test")
def test():
    return {"test": "ok"}

@app.route("/health")
def health():
    return jsonify({"status": "ok", "providers": available_providers()})

@app.route("/api/providers")
def get_providers():
    return jsonify(available_providers())


@app.route("/api/map")
def map_suppliers():
    product  = request.args.get("product", "").strip()
    depth    = max(0, min(int(request.args.get("depth", 2)), 3))
    provider = request.args.get("provider", "gemini").strip().lower()

    if not product:
        return jsonify({"error": "product required"}), 400
    if provider not in ("anthropic", "openai", "gemini", "deepseek"):
        return jsonify({"error": f"Unknown provider: {provider}"}), 400

    queue, done = [], threading.Event()

    def run():
        try: run_pipeline(product, depth, provider, queue, [], {})
        except Exception as e: queue.append({"type": "error", "message": str(e)})
        finally: done.set()

    threading.Thread(target=run, daemon=True).start()

    def generate():
        sent = 0
        while not done.is_set() or sent < len(queue):
            while sent < len(queue):
                yield sse(queue[sent]); sent += 1
            time.sleep(0.1)
        yield sse({"type": "stream_end"})

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/api/map_all")
def map_all():
    collected_oems  = []
    collected_tiers = {}
    providers = available_providers()
    product   = request.args.get("product", "").strip()
    depth     = max(0, min(int(request.args.get("depth", 2)), 3))

    if not product:
        return jsonify({"error": "product required"}), 400

    queue, done = [], threading.Event()

    def run():
        try:
            queue.append({"type": "status", "message": "⚡Beginning series of pipeline runs."})
            for provider in providers:
                if not provider.get("configured"):
                    queue.append({"type": "status", "message": f"✅ Skipping {provider.get('name')} due to lack of API key"})
                else:
                    provider_id = provider.get("id")
                    queue.append({"type": "provider_start", "provider": provider["id"], "name": provider["name"],
                                  "message": f"Starting Pipeline with {provider['name']}"})
                    run_pipeline(product, depth, provider_id, queue, collected_oems, collected_tiers)
                    queue.append({"type": "provider_done", "provider": provider["id"], "name": provider["name"],
                                  "message": f"{provider['name']} pipeline complete."})
            queue.append({"type": "status", "message": "✅ Completed series of pipeline executions."})
        except Exception as e:
            queue.append({"type": "error", "message": str(e)})
        finally:
            done.set()

    threading.Thread(target=run, daemon=True).start()

    def generate():
        sent = 0
        while not done.is_set() or sent < len(queue):
            while sent < len(queue):
                yield sse(queue[sent]); sent += 1
            time.sleep(0.1)
        yield sse({"type": "stream_end"})

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


# ── Bulk dev export routes ────────────────────────────────────────────────────

@app.route("/dev/bulk_export")
def bulk_export():
    collected_oems  = []
    collected_tiers = {}
    providers = available_providers()
    results   = []

    for i, product_input in enumerate(PRODUCTS_TO_TEST, 1):
        print(f"\n[bulk_export] ── Product {i}/{len(PRODUCTS_TO_TEST)}: {product_input}")
        queue = []
        try:
            for provider in providers:
                print(f"\n[bulk_export] - Running pipeline for {provider.get('name')}")
                run_pipeline(product_input, DEPTH, provider.get("id"), queue, collected_oems, collected_tiers)
        except Exception as e:
            print(f"[bulk_export] Pipeline error for '{product_input}': {e}")
            results.append({
                "product_input": product_input,
                "supply_chain":  {"product": {"product_name": product_input},
                                  "oems": [], "tiers": {}, "summary": f"Error: {e}"},
                "sheet_prefix":  safe_sheet_name(product_input, i),
            })
            continue

        supply_chain = None
        for event in queue:
            if event.get("type") == "complete":
                incoming = event.get("supply_chain", {})
                if supply_chain is None:
                    supply_chain = incoming
                else:
                    existing_oem_names = {o.get("company_name", "").lower() for o in supply_chain.get("oems", [])}
                    for oem in incoming.get("oems", []):
                        if oem.get("company_name", "").lower() not in existing_oem_names:
                            supply_chain["oems"].append(oem)
                    for tier_key, suppliers in incoming.get("tiers", {}).items():
                        if tier_key not in supply_chain["tiers"]:
                            supply_chain["tiers"][tier_key] = suppliers
                        else:
                            existing_names = {s.get("company_name", "").lower() for s in supply_chain["tiers"][tier_key]}
                            for s in suppliers:
                                if s.get("company_name", "").lower() not in existing_names:
                                    supply_chain["tiers"][tier_key].append(s)

        if not supply_chain:
            supply_chain = {"product": {"product_name": product_input}, "oems": [], "tiers": {}}

        results.append({
            "product_input": product_input,
            "supply_chain":  supply_chain,
            "sheet_prefix":  safe_sheet_name(product_input, i),
        })
        tier_counts = {k: len(v) for k, v in supply_chain.get("tiers", {}).items()}
        print(f"[bulk_export] Done: oems={len(supply_chain.get('oems', []))} tiers={tier_counts}")

    print(f"\n[bulk_export] Building workbook for {len(results)} products...")
    wb_bytes = build_bulk_workbook(results, "All Providers", DEPTH)

    import datetime
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    filename  = f"bulk_export_{timestamp}.xlsx"
    print(f"[bulk_export] Sending {len(wb_bytes)} bytes as {filename}")

    return Response(
        wb_bytes,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


@app.route("/dev/multi_single_export")
def multi_single_export():
    accept = request.headers.get("Accept", "")
    if "text/html" in accept:
        return Response("""<!DOCTYPE html>
<html>
<head>
  <title>Bulk Export</title>
  <style>
    body { font-family: monospace; padding: 2rem; background: #111; color: #eee; }
    .done { color: #4caf50; }
    .error { color: #f44336; }
  </style>
</head>
<body>
  <h2>Bulk Export</h2>
  <div id="log"></div>
  <script>
    const log = document.getElementById('log');
    const append = (msg, cls) => {
      const d = document.createElement('div');
      d.className = cls || '';
      d.textContent = msg;
      log.appendChild(d);
    };
    const evtSource = new EventSource('/dev/multi_single_export/stream');
    evtSource.onmessage = (e) => {
      const msg = JSON.parse(e.data);
      if (msg.type === 'start') {
        append(`Starting export for ${msg.total} products...`);
      } else if (msg.type === 'file_ready') {
        append(`[${msg.index}/${msg.total}] ✓ ${msg.filename}`, 'done');
        const a = document.createElement('a');
        a.href = `/dev/multi_single_export/download/${msg.file_id}`;
        a.download = msg.filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
      } else if (msg.type === 'error') {
        append(`[error] ${msg.product}: ${msg.error}`, 'error');
      } else if (msg.type === 'done') {
        append(`All ${msg.total} exports complete.`, 'done');
        evtSource.close();
      }
    };
    evtSource.onerror = () => { append('SSE connection lost.', 'error'); evtSource.close(); };
  </script>
</body>
</html>""", mimetype="text/html")
    return Response(status=400)


@app.route("/dev/multi_single_export/stream")
def multi_export_stream():
    def generate():
        collected_oems  = []
        collected_tiers = {}
        providers = available_providers()
        total     = len(PRODUCTS_TO_TEST)

        yield f"data: {json.dumps({'type': 'start', 'total': total})}\n\n"

        for i, product_input in enumerate(PRODUCTS_TO_TEST, 1):
            print(f"\n[bulk_export] ── Product {i}/{total}: {product_input}")
            queue = []
            try:
                for provider in providers:
                    if not provider.get("configured"):
                        continue
                    run_pipeline(product_input, DEPTH, provider.get("id"), queue, collected_oems, collected_tiers)
            except Exception as e:
                print(f"[bulk_export] Pipeline error for '{product_input}': {e}")
                yield f"data: {json.dumps({'type': 'error', 'product': product_input, 'error': str(e)})}\n\n"
                continue

            provider_results = {}
            for event in queue:
                if event.get("type") == "complete":
                    sc   = event.get("supply_chain", {})
                    prov = sc.get("provider", "unknown")
                    if sc.get("oems") or sc.get("tiers"):
                        provider_results[prov] = sc

            supply_chain = {"product": {}, "oems": [], "tiers": {}}
            for sc in provider_results.values():
                if not supply_chain["product"]:
                    supply_chain["product"] = sc.get("product", {})
                for oem in sc.get("oems", []):
                    existing = [o.get("company_name", "").lower() for o in supply_chain["oems"]]
                    if oem.get("company_name", "").lower() not in existing:
                        supply_chain["oems"].append(oem)
                for tier_key, suppliers in sc.get("tiers", {}).items():
                    supply_chain["tiers"].setdefault(tier_key, [])
                    existing = [s.get("company_name", "").lower() for s in supply_chain["tiers"][tier_key]]
                    for s in suppliers:
                        if s.get("company_name", "").lower() not in existing:
                            supply_chain["tiers"][tier_key].append(s)

            if not supply_chain["product"]:
                supply_chain["product"] = {"product_name": product_input}

            result   = [{"product_input": product_input, "supply_chain": supply_chain,
                          "provider_results": provider_results,
                          "sheet_prefix": safe_sheet_name(product_input, i)}]
            wb_bytes = build_bulk_workbook(result, "All Providers", DEPTH)

            file_id  = str(uuid.uuid4())
            filename = safe_filename(product_input, i)
            tmp_path = os.path.join(tempfile.gettempdir(), f"{file_id}.xlsx")
            with open(tmp_path, "wb") as f:
                f.write(wb_bytes)

            with _export_files_lock:
                _export_files[file_id] = {"path": tmp_path, "filename": filename}

            tier_counts = {k: len(v) for k, v in supply_chain.get("tiers", {}).items()}
            print(f"[bulk_export] Ready: {filename} | oems={len(supply_chain.get('oems', []))} tiers={tier_counts}")
            yield f"data: {json.dumps({'type': 'file_ready', 'file_id': file_id, 'filename': filename, 'index': i, 'total': total})}\n\n"

        yield f"data: {json.dumps({'type': 'done', 'total': total})}\n\n"

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/dev/multi_single_export/download/<file_id>")
def multi_export_download(file_id):
    with _export_files_lock:
        entry = _export_files.pop(file_id, None)

    if not entry:
        return jsonify({"error": "File not found or already downloaded"}), 404

    response = send_file(
        entry["path"],
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=entry["filename"]
    )

    @response.call_on_close
    def cleanup():
        try:
            os.remove(entry["path"])
        except OSError:
            pass

    return response


if __name__ == "__main__":
    print("🚀 Supplier Chain Mapper — multi-provider")
    for p in available_providers():
        status = "✅ ready" if p["configured"] else "❌ no key"
        print(f"   {p['name']:25s} {status:12s} search: {p['search']}")
    print("   Running on http://localhost:5000")
    app.run(debug=False, port=5000, threaded=True)
