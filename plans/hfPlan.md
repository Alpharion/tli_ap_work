# Plan: Hugging Face Local Provider

## Context

The current pipeline (`run_pipeline`) calls cloud LLMs (Anthropic, Gemini, DeepSeek) via API. The goal is to explore running the same pipeline fully locally using two Hugging Face models:
- **Nezpic/supply-chain-gpt2-model** — GPT-2 fine-tuned on supply chain text, used to generate supply-chain-relevant prose from a prompt prefix
- **Qwen/Qwen2.5-1.5B-Instruct** — small instruction-tuned model, used to convert that prose into structured JSON matching the existing schema

---

## How the Two-Model Pipeline Works

The existing `call_ai(prompt, provider)` function takes a prompt asking for JSON and returns a string. The HF approach splits this into two stages:

```
Stage 1 — Domain knowledge extraction
  Input:  a prefix derived from the prompt (e.g. "Tier-1 suppliers of TSMC for 7nm lithography equipment include")
  Model:  Nezpic/supply-chain-gpt2-model (GPT-2 causal LM, text completion)
  Output: free-form supply chain prose

Stage 2 — JSON structuring
  Input:  the prose from Stage 1 + the original JSON schema instruction
  Model:  Qwen/Qwen2.5-1.5B-Instruct (instruction-tuned, chat template)
  Output: JSON string matching the existing schema
```

`safe_parse_json` (app.py:181) already handles malformed JSON with 4 recovery strategies — it will catch Qwen's imperfect output.

---

## Implementation Approach

### New file: `hf_provider.py`

Following the file-separation pattern, all HF logic lives here. `app.py` gets only 3 minimal additions.

```python
# hf_provider.py
from transformers import pipeline
import torch

_sc_pipe   = None   # Nezpic GPT-2 pipeline
_qwen_pipe = None   # Qwen2.5-1.5B-Instruct pipeline

def _load_models():
    global _sc_pipe, _qwen_pipe
    if _sc_pipe is None:
        _sc_pipe = pipeline("text-generation",
                            model="Nezpic/supply-chain-gpt2-model",
                            device=0 if torch.cuda.is_available() else -1)
    if _qwen_pipe is None:
        _qwen_pipe = pipeline("text-generation",
                              model="Qwen/Qwen2.5-1.5B-Instruct",
                              device=0 if torch.cuda.is_available() else -1,
                              torch_dtype=torch.float16 if torch.cuda.is_available() else torch.float32)

def call_hf(prompt: str, max_new_tokens: int = 512) -> str:
    _load_models()

    # Stage 1: use first 120 chars as GPT-2 prefix for supply chain priming
    prefix = prompt[:120].strip()
    sc_result = _sc_pipe(prefix, max_new_tokens=200, do_sample=False)[0]["generated_text"]

    # Stage 2: ask Qwen to structure GPT-2 prose as JSON per original schema
    messages = [
        {"role": "system", "content": "You are a supply chain data extractor. Return only valid JSON, no markdown."},
        {"role": "user",   "content": f"Original request:\n{prompt}\n\nAdditional context:\n{sc_result}\n\nReturn JSON only."}
    ]
    qwen_result = _qwen_pipe(messages, max_new_tokens=max_new_tokens)[0]["generated_text"]
    if isinstance(qwen_result, list):
        last = qwen_result[-1]
        return last.get("content", "") if isinstance(last, dict) else str(last)
    return str(qwen_result)
```

### Changes to `app.py` (3 additions only)

```python
# 1. Top of file, after existing imports
from hf_provider import call_hf

# 2. Inside call_ai(), new elif branch after deepseek
elif provider == "huggingface":
    return call_hf(prompt, max_tokens)

# 3. Inside available_providers()
{"id": "huggingface", "name": "HuggingFace (Local)", "model": "supply-chain-gpt2 + Qwen2.5-1.5B",
 "env": None, "configured": True, "search": "None (offline)"},
```

### `requirements.txt` additions

```
transformers>=4.40.0
torch>=2.0.0
accelerate>=0.27.0
```

---

## Benefits

| Benefit | Detail |
|---|---|
| **Zero API cost** | No keys, no per-token billing — unlimited queries |
| **Fully offline** | After one-time model download (~1.5 GB total), no internet needed |
| **Privacy** | No product/supplier data leaves the machine |
| **Domain-specific priming** | Nezpic GPT-2 was fine-tuned on supply chain corpora — biases completions toward real supplier names and relationships |
| **No rate limits** | No throttling from provider APIs |

---

## Drawbacks

| Drawback | Detail |
|---|---|
| **Quality gap** | GPT-2 (117M–774M params) and Qwen 1.5B are far smaller than Claude Sonnet / Gemini 2.5 Flash — output will be less accurate and more likely to hallucinate |
| **No real-time search** | Gemini uses live Google Search grounding; HF models only have training-data knowledge (cutoff ~2023 for Qwen) |
| **JSON reliability** | Qwen 1.5B produces malformed JSON more often than larger models; `safe_parse_json` mitigates but some calls may still return `None` |
| **Two-stage latency** | Running two models sequentially per call (×N OEMs ×N tiers) makes the total pipeline much slower than a single API round-trip |
| **No streaming** | Local inference is blocking; long silent gaps between SSE progress events |
| **Two-stage quality loss** | GPT-2 prose → Qwen JSON is a lossy conversion; information may not survive structuring cleanly |

---

## Hardware Limitations

| Scenario | Memory needed | Speed per `call_ai` | Verdict |
|---|---|---|---|
| **CPU only** | ~4 GB RAM (fp32) | 3–8 min | Technically works; too slow for multi-OEM runs |
| **GPU 4 GB VRAM** | fits in fp16 | ~15–30 sec | Usable for depth=0–1, single OEM |
| **GPU 8 GB VRAM** | comfortable | ~5–15 sec | Acceptable for depth=1–2 |
| **GPU 16 GB+ VRAM** | both models loaded simultaneously | ~3–8 sec | Good; comparable to slow API |

> **Critical note:** A full depth-2 run (5 OEMs × 4 T1 × 4 T2 ≈ 85 `call_ai` calls) would take **4–11 hours on CPU** or **7–120 minutes on a 4–8 GB GPU**. This makes it unsuitable for interactive use at depth > 1 without a strong dedicated GPU.

Models are lazy-loaded on first call — server startup time is not affected.

---

## Files to Create / Modify

| File | Change |
|---|---|
| `hf_provider.py` | New file — all HF model loading and two-stage inference logic |
| `app.py` | 3 targeted additions: import, `call_ai` branch, `available_providers` entry |
| `requirements.txt` | Add `transformers`, `torch`, `accelerate` |

---

## Verification

1. `pip install transformers torch accelerate`
2. Start app — confirm "HuggingFace (Local)" appears in the provider dropdown
3. Run a depth=0 (OEM only) search — confirm it returns OEM JSON without crashing
4. Check SSE log — pipeline should progress through product ID and OEM steps
5. Confirm output renders on map and in Excel export (even if lower quality than cloud providers)
