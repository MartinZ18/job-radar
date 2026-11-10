"""Turn queued interview posts (data/interview_posts.json) into structured
items — questions, live-coding exercises, requested projects, behavioural
questions — with a Groq LLM, and merge them into data/interview_items.json.

Each item is keyed by its normalized text; seeing the same item in several
posts increments `veces`, which is what "most requested" means downstream.

Needs GROQ_API_KEY in .env.

Usage:
    python scripts/extract_interviews.py            # process the queue
    python scripts/extract_interviews.py --limit 8  # only the first N posts (cheap test)
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
import unicodedata
from datetime import date

import requests
from dotenv import load_dotenv

from common import DATA_DIR, ROOT
from fetch_interviews import POSTS_PATH, SEEN_PATH, load_config, load_json, save_json

ITEMS_PATH = DATA_DIR / "interview_items.json"
GROQ_URL = "https://api.groq.com/openai/v1/chat/completions"

SYSTEM_PROMPT = """You extract interview-preparation material from forum posts and comments about software job interviews.

From the posts you receive, list every concrete, reusable item of one of these types:
- "pregunta": a technical question that was asked (or that the author says is commonly asked).
- "livecoding": a live coding / pair programming / take-home exercise that was requested (what had to be built, constraints, duration if stated).
- "proyecto": a project or portfolio piece the author says companies ask for or value.
- "comportamental": a behavioural / non-technical question.

Rules:
- Only include items that are concrete enough to practise. Skip opinions, rants, tool ads and generic advice ("study data structures").
- The item must be something an INTERVIEWER or COMPANY asked or required of a candidate. Never extract the question the author is asking the forum ("is it normal that...", "how often do they...") — that is not interview material.
- "proyecto" only when the author reports that a company asked for, tested with, or explicitly valued that kind of project. The author's own side project is not an item.
- Prefer items that name what was evaluated. If the post only says "they asked me hard questions", extract nothing.
- Write `texto` in Spanish (neutral, professional), as a single clear sentence or question. Keep technical terms as they are (REST, JWT, SQL, RAG...).
- `detalle`: optional, at most two sentences of useful context in Spanish (company type, seniority, what they were checking). Empty string if none.
- `seccion` must be exactly one of: {secciones}.
- `idioma` describes the language context of the interview, exactly one of:
  "es" = the post is written in Spanish, or the interview/company is Spanish-speaking (Argentina, LATAM, Spain);
  "en-bajo" = English-speaking context but the author says basic English was enough, or no English was tested;
  "en" = English-language interview context (fluent English expected or not mentioned).
- `post` is the 1-based index of the post the item comes from.
- If a post has nothing extractable, produce no items for it.

Answer with a JSON object: {{"items": [{{"post": 1, "tipo": "...", "seccion": "...", "idioma": "...", "texto": "...", "detalle": "..."}}, ...]}}"""

IDIOMAS = ("es", "en-bajo", "en")


def normalize(text: str) -> str:
    t = unicodedata.normalize("NFKD", text.lower())
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = re.sub(r"[^a-z0-9 ]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def groq_call(api_key: str, model: str, system: str, user: str) -> tuple[dict, dict]:
    payload = {
        "model": model,
        "temperature": 0,
        "response_format": {"type": "json_object"},
        "messages": [{"role": "system", "content": system}, {"role": "user", "content": user}],
    }
    for attempt in range(4):
        resp = requests.post(GROQ_URL, json=payload, headers={"Authorization": f"Bearer {api_key}"}, timeout=90)
        if resp.status_code == 429:  # free tier: respect retry-after, then try again
            wait = float(resp.headers.get("retry-after", 10 * (attempt + 1)))
            print(f"  rate limited, waiting {wait:.0f}s", file=sys.stderr)
            time.sleep(wait)
            continue
        resp.raise_for_status()
        body = resp.json()
        return json.loads(body["choices"][0]["message"]["content"]), body.get("usage", {})
    raise RuntimeError("groq: still rate limited after retries")


def build_user_message(posts: list[dict]) -> str:
    parts = []
    for i, p in enumerate(posts, start=1):
        parts.append(f"### Post {i}\nSource: {p['source']}\nTitle: {p['title']}\n{p['body'] or '(no body)'}")
    return "\n\n".join(parts)


def merge(items_db: dict, extracted: list[dict], posts: list[dict], tipos: list[str], secciones: list[str]) -> int:
    added = 0
    today = date.today().isoformat()
    for it in extracted:
        texto = (it.get("texto") or "").strip()
        if len(texto) < 12:
            continue
        idx = it.get("post")
        post = posts[idx - 1] if isinstance(idx, int) and 1 <= idx <= len(posts) else None
        tipo = it.get("tipo") if it.get("tipo") in tipos else "pregunta"
        seccion = it.get("seccion") if it.get("seccion") in secciones else "Otro"
        idioma = it.get("idioma") if it.get("idioma") in IDIOMAS else "en"
        key = normalize(texto)
        entry = items_db.get(key)
        if entry is None:
            entry = {"tipo": tipo, "seccion": seccion, "idioma": idioma, "texto": texto,
                     "detalle": (it.get("detalle") or "").strip(), "veces": 0, "fuentes": [], "primera_vez": today}
            items_db[key] = entry
            added += 1
        elif idioma == "es" and entry.get("idioma") != "es":  # a Spanish sighting upgrades the item
            entry["idioma"] = "es"
        if post and post["url"] not in [f["url"] for f in entry["fuentes"]]:
            entry["fuentes"].append({"url": post["url"], "titulo": post["title"], "origen": post["source"]})
            entry["veces"] += 1
        if not entry["detalle"] and it.get("detalle"):
            entry["detalle"] = it["detalle"].strip()
    return added


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--limit", type=int, help="process only the first N queued posts")
    args = parser.parse_args()

    load_dotenv(ROOT / ".env")
    api_key = os.environ.get("GROQ_API_KEY", "").strip()
    if not api_key:
        print("GROQ_API_KEY missing in .env", file=sys.stderr)
        return 1

    cfg = load_config()
    llm = cfg["llm"]
    tipos, secciones = cfg["tipos"], cfg["secciones"]
    system = SYSTEM_PROMPT.format(secciones=", ".join(f'"{s}"' for s in secciones))

    queue: list[dict] = load_json(POSTS_PATH, [])
    if not queue:
        print("no queued posts")
        return 0
    budget = args.limit or int(llm.get("max_posts_per_run", 40))
    priority = llm.get("source_priority", [])

    def rank(post: dict) -> int:
        for i, name in enumerate(priority):
            if post["source"].startswith(name):
                return i
        return len(priority)

    queue.sort(key=rank)
    batch_size = int(llm.get("posts_per_call", 4))
    todo, rest = queue[:budget], queue[budget:]

    items_db: dict = load_json(ITEMS_PATH, {})
    seen: dict = load_json(SEEN_PATH, {})
    total_added = total_items = tokens_in = tokens_out = 0
    processed: list[dict] = []

    for start in range(0, len(todo), batch_size):
        batch = todo[start:start + batch_size]
        try:
            result, usage = groq_call(api_key, llm["model"], system, build_user_message(batch))
        except Exception as exc:
            print(f"batch {start // batch_size + 1}: ERROR {exc} — stopping, queue kept", file=sys.stderr)
            break
        extracted = result.get("items", []) if isinstance(result, dict) else []
        added = merge(items_db, extracted, batch, tipos, secciones)
        total_items += len(extracted)
        total_added += added
        processed.extend(batch)
        tokens_in += usage.get("prompt_tokens", 0)
        tokens_out += usage.get("completion_tokens", 0)
        print(f"batch {start // batch_size + 1}: {len(batch)} posts -> {len(extracted)} items ({added} new; "
              f"{usage.get('prompt_tokens', 0)} in / {usage.get('completion_tokens', 0)} out tokens)")
        time.sleep(float(llm.get("delay_seconds", 3)))

    save_json(ITEMS_PATH, items_db)
    for p in processed:
        seen[p["url"]] = date.today().isoformat()
    save_json(SEEN_PATH, seen)
    remaining = [p for p in queue if p["url"] not in seen]
    save_json(POSTS_PATH, remaining)
    print(f"{len(processed)} post(s) processed, {total_items} item(s) extracted, {total_added} new; "
          f"{len(items_db)} total in items db; {len(remaining)} post(s) left in queue; "
          f"tokens {tokens_in} in / {tokens_out} out")
    return 0


if __name__ == "__main__":
    sys.exit(main())
