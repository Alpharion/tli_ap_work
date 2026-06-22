"""
ai.py — AI provider abstraction and web search helpers.

Exports:
  call_ai(prompt, provider, max_tokens) -> str
  safe_parse_json(text) -> list | dict | None
  get_evidence(query, provider) -> str
  available_providers() -> list[dict]
  _source_url_cache: dict  (populated by ddg_search_and_scrape)
"""

import json
import os
import re
import time

from dotenv import load_dotenv
load_dotenv()


# ── JSON parsing helpers ───────────────────────────────────────────────────────

def clean_json(text: str) -> str:
    """Strip markdown fences and fix common AI JSON formatting issues."""
    text = re.sub(r"```[a-z]*", "", text).strip("` \n")
    text = re.sub(r",\s*([\]}])", r"\1", text)
    text = re.sub(r"//[^\n]*", "", text)
    text = re.sub(r"/\*.*?\*/", "", text, flags=re.DOTALL)
    return text.strip()


def extract_balanced(text: str, open_char: str, close_char: str) -> str | None:
    """Extract the largest balanced bracket block from text."""
    best = None
    for start in range(len(text)):
        if text[start] != open_char:
            continue
        depth = 0
        for end in range(start, len(text)):
            if text[end] == open_char:
                depth += 1
            elif text[end] == close_char:
                depth -= 1
                if depth == 0:
                    candidate = text[start:end+1]
                    if best is None or len(candidate) > len(best):
                        best = candidate
                    break
    return best


def safe_parse_json(text: str) -> list | dict | None:
    """Try multiple strategies to parse potentially malformed JSON from AI."""
    cleaned = clean_json(text)

    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass

    arr_block = extract_balanced(cleaned, "[", "]")
    if arr_block:
        try:
            return json.loads(arr_block)
        except json.JSONDecodeError:
            fixed = re.sub(r",\s*([\]}])", r"", arr_block)
            try:
                return json.loads(fixed)
            except json.JSONDecodeError:
                pass

    obj_block = extract_balanced(cleaned, "{", "}")
    if obj_block:
        try:
            return json.loads(obj_block)
        except json.JSONDecodeError:
            fixed = re.sub(r",\s*([\]}])", r"", obj_block)
            try:
                return json.loads(fixed)
            except json.JSONDecodeError:
                pass

    try:
        open_brackets = cleaned.count("[") - cleaned.count("]")
        open_braces   = cleaned.count("{") - cleaned.count("}")
        fixed = cleaned + ("}" * max(0, open_braces)) + ("]" * max(0, open_brackets))
        fixed = re.sub(r",\s*([\]}])", r"", fixed)
        return json.loads(fixed)
    except json.JSONDecodeError:
        pass

    return None


# ── Search backends ────────────────────────────────────────────────────────────

_source_url_cache: dict = {}


def ddg_search_and_scrape(query: str, n: int = 4) -> str:
    """
    Try DuckDuckGo search + scraping. Falls back gracefully if blocked.
    Returns an evidence string or empty string if search is unavailable.
    """
    import requests
    from bs4 import BeautifulSoup

    results = []
    for attempt in range(2):
        try:
            time.sleep(2 * (attempt + 1))
            try:
                from ddgs import DDGS
            except ImportError:
                from duckduckgo_search import DDGS
            with DDGS() as ddgs:
                results = list(ddgs.text(query, max_results=n))
            if results:
                break
        except Exception as e:
            print(f"DDG attempt {attempt+1} failed: {e}")

    if not results:
        _source_url_cache[query] = []
        print(f"Search unavailable for: {query} — AI will use training knowledge")
        return ""

    urls, parts = [], []
    for r in results[:3]:
        url     = r.get("href", "")
        snippet = r.get("body", "")
        page    = ""
        if url:
            urls.append(url)
            try:
                hdrs = {"User-Agent": "Mozilla/5.0 (compatible; SupplierMapper/1.0)"}
                resp = requests.get(url, headers=hdrs, timeout=8)
                resp.raise_for_status()
                soup = BeautifulSoup(resp.text, "html.parser")
                for t in soup(["script", "style", "nav", "footer", "header"]):
                    t.decompose()
                page = soup.get_text(separator=" ", strip=True)[:1500]
            except Exception:
                pass
        parts.append(f"SOURCE: {url}\nSNIPPET: {snippet}\nCONTENT: {page}")
    _source_url_cache[query] = urls
    return "\n---\n".join(parts)


def gemini_search_and_answer(query: str, api_key: str) -> str:
    """Use Gemini with Google Search grounding. Falls back to base knowledge on error."""
    try:
        from google import genai
        from google.genai import types
        client = genai.Client(api_key=api_key)
        resp = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=f"Search the web and provide detailed factual information about: {query}\nInclude company names, countries, and supply chain relationships.",
            config=types.GenerateContentConfig(
                tools=[types.Tool(google_search=types.GoogleSearch())]
            )
        )
        text_parts = [
            part.text for part in resp.candidates[0].content.parts
            if hasattr(part, "text") and part.text
        ]
        return "\n".join(text_parts) if text_parts else "No results."
    except Exception as e:
        print(f"Gemini search grounding error: {e}, falling back to base knowledge")
        try:
            from google import genai
            client = genai.Client(api_key=api_key)
            resp = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=f"Based on your knowledge, provide information about: {query}\nInclude company names, countries, and supply chain relationships."
            )
            return resp.text
        except Exception as e2:
            return f"Search failed: {e2}"


# ── AI call abstraction ────────────────────────────────────────────────────────

def call_ai(prompt: str, provider: str, max_tokens: int = 1000) -> str:
    if provider == "anthropic":
        import anthropic
        key = os.getenv("ANTHROPIC_API_KEY")
        if not key: raise ValueError("ANTHROPIC_API_KEY not set in .env")
        c = anthropic.Anthropic(api_key=key)
        r = c.messages.create(model="claude-sonnet-4-6", max_tokens=max_tokens,
                               messages=[{"role": "user", "content": prompt}])
        return r.content[0].text

    elif provider == "openai":
        from openai import OpenAI
        key = os.getenv("OPENAI_API_KEY")
        if not key: raise ValueError("OPENAI_API_KEY not set in .env")
        c = OpenAI(api_key=key)
        r = c.chat.completions.create(model="gpt-4.1-mini", max_tokens=max_tokens,
                                       messages=[{"role": "user", "content": prompt}])
        return r.choices[0].message.content

    elif provider == "gemini":
        from google import genai
        key = os.getenv("GEMINI_API_KEY")
        if not key: raise ValueError("GEMINI_API_KEY not set in .env")
        client = genai.Client(api_key=key)
        r = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        return r.text

    elif provider == "deepseek":
        from openai import OpenAI
        key = os.getenv("DEEPSEEK_API_KEY")
        if not key: raise ValueError("DEEPSEEK_API_KEY not set in .env")
        client = OpenAI(api_key=key, base_url="https://api.deepseek.com")
        resp = client.chat.completions.create(
            model="deepseek-chat",
            max_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt}]
        )
        return resp.choices[0].message.content

    else:
        raise ValueError(f"Unknown provider: {provider}")


def get_evidence(query: str, provider: str) -> str:
    """Get web evidence using the best method for each provider."""
    if provider == "gemini":
        return gemini_search_and_answer(query, os.getenv("GEMINI_API_KEY", ""))
    else:
        return ddg_search_and_scrape(query)


def build_prompt_with_evidence(base_prompt: str, evidence: str) -> str:
    if evidence and evidence.strip():
        return base_prompt.replace("{EVIDENCE_BLOCK}",
            f"Use this web evidence to inform your answer:\n{evidence}")
    else:
        return base_prompt.replace("{EVIDENCE_BLOCK}",
            "No live web search is available. Use your training knowledge "
            "to provide the most accurate answer you can. Be explicit about "
            "what you know with high vs low confidence.")


# ── Provider registry ──────────────────────────────────────────────────────────

def available_providers() -> list[dict]:
    return [
        {"id": "anthropic", "name": "Anthropic Claude", "model": "claude-sonnet-4-6",
         "env": "ANTHROPIC_API_KEY", "configured": bool(os.getenv("ANTHROPIC_API_KEY")),
         "search": "DuckDuckGo"},
        {"id": "gemini", "name": "Google Gemini", "model": "gemini-2.5-flash",
         "env": "GEMINI_API_KEY", "configured": bool(os.getenv("GEMINI_API_KEY")),
         "search": "Google Search (built-in)"},
        {"id": "deepseek", "name": "DeepSeek V3", "model": "deepseek-chat",
         "env": "DEEPSEEK_API_KEY", "configured": bool(os.getenv("DEEPSEEK_API_KEY")),
         "search": "DuckDuckGo"},
    ]
