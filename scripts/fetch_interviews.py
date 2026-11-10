"""Collect forum/community posts about tech interviews (what gets asked, what
kind of live coding or take-home, which projects) and queue the unseen ones in
data/interview_posts.json for extract_interviews.py.

Usage:
    python scripts/fetch_interviews.py            # queue new posts
    python scripts/fetch_interviews.py --dry-run  # only print
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path

import feedparser
import requests
import yaml

from common import DATA_DIR, ROOT, clean_text

CONFIG_PATH = ROOT / "config" / "interviews.yaml"
POSTS_PATH = DATA_DIR / "interview_posts.json"      # queue for the extractor
SEEN_PATH = DATA_DIR / "interview_seen.json"        # url -> date first fetched


@dataclass
class Post:
    url: str
    title: str
    body: str
    source: str
    published: str  # ISO date or ""

    def to_dict(self) -> dict:
        return asdict(self)


def load_config() -> dict:
    with CONFIG_PATH.open(encoding="utf-8") as f:
        return yaml.safe_load(f)


def load_json(path: Path, default):
    if not path.exists():
        return default
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def save_json(path: Path, data) -> None:
    DATA_DIR.mkdir(exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


class Http:
    def __init__(self, cfg: dict):
        self.session = requests.Session()
        self.session.headers["User-Agent"] = cfg["user_agent"]
        self.default_delay = float(cfg.get("delay_seconds", 4))
        self.delay = self.default_delay
        self.timeout = float(cfg.get("timeout_seconds", 25))
        self._last = 0.0

    def get(self, url: str, params: dict | None = None) -> requests.Response:
        wait = self.delay - (time.monotonic() - self._last)
        if wait > 0:
            time.sleep(wait)
        resp = self.session.get(url, params=params, timeout=self.timeout)
        self._last = time.monotonic()
        resp.raise_for_status()
        return resp


def _iso(ts: int | float | None) -> str:
    return datetime.fromtimestamp(ts, tz=timezone.utc).date().isoformat() if ts else ""


# ---------- sources ----------

def fetch_hn(src: dict, http: Http, body_chars: int) -> list[Post]:
    posts = []
    for q in src["queries"]:
        data = http.get("https://hn.algolia.com/api/v1/search",
                        {"query": q, "tags": "story", "hitsPerPage": src.get("hits_per_query", 20)}).json()
        for h in data.get("hits", []):
            url = f"https://news.ycombinator.com/item?id={h['objectID']}"
            body = clean_text(h.get("story_text") or "", body_chars)
            posts.append(Post(url, h.get("title", ""), body, src["name"], _iso(h.get("created_at_i"))))
    for q in src.get("comment_queries", []):
        data = http.get("https://hn.algolia.com/api/v1/search",
                        {"query": q, "tags": "comment", "hitsPerPage": src.get("comments_per_query", 30)}).json()
        for h in data.get("hits", []):
            body = clean_text(h.get("comment_text") or "", body_chars)
            if len(body) < 200:  # one-liners carry nothing extractable
                continue
            url = f"https://news.ycombinator.com/item?id={h['objectID']}"
            title = f"Comentario en: {h.get('story_title') or 'HN'}"
            posts.append(Post(url, title, body, f"{src['name']} (comentarios)", _iso(h.get("created_at_i"))))
    return posts


def fetch_devto(src: dict, http: Http, body_chars: int) -> list[Post]:
    posts = []
    for tag in src["tags"]:
        data = http.get("https://dev.to/api/articles", {"tag": tag, "per_page": src.get("per_tag", 30), "top": 90}).json()
        for a in data:
            body = clean_text(a.get("description") or "", body_chars)
            posts.append(Post(a["url"], a["title"], body, src["name"], (a.get("published_at") or "")[:10]))
    return posts


def fetch_discourse(src: dict, http: Http, body_chars: int) -> list[Post]:
    posts = []
    base = src["base_url"].rstrip("/")
    for q in src["queries"]:
        data = http.get(f"{base}/search.json", {"q": q}).json()
        blurbs = {p["topic_id"]: p.get("blurb", "") for p in data.get("posts", [])}
        for t in data.get("topics", []):
            url = f"{base}/t/{t['slug']}/{t['id']}"
            body = clean_text(blurbs.get(t["id"], ""), body_chars)
            posts.append(Post(url, t["title"], body, src["name"], (t.get("created_at") or "")[:10]))
    return posts


def fetch_reddit_rss(src: dict, http: Http, body_chars: int) -> list[Post]:
    posts = []
    for sub, queries in src["subreddits"].items():
        for q in queries:
            try:
                resp = http.get(f"https://www.reddit.com/r/{sub}/search.rss",
                                {"q": q, "restrict_sr": 1, "sort": "new", "t": "year", "limit": 25})
            except requests.HTTPError as exc:  # 429/403: reddit throttles anonymous feeds
                print(f"  [Reddit r/{sub}] skipped: {exc.response.status_code}", file=sys.stderr)
                continue
            feed = feedparser.parse(resp.text)
            for e in feed.entries:
                body = clean_text(e.get("summary", ""), body_chars)
                published = ""
                if e.get("published_parsed"):
                    published = time.strftime("%Y-%m-%d", e.published_parsed)
                posts.append(Post(e.link, e.title, body, f"Reddit r/{sub}", published))
    return posts


FETCHERS = {"hn": fetch_hn, "devto": fetch_devto, "discourse": fetch_discourse, "reddit_rss": fetch_reddit_rss}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    cfg = load_config()
    http = Http(cfg["http"])
    body_chars = int(cfg["llm"].get("body_chars", 1500))
    prefilter = [re.compile(p, re.IGNORECASE) for p in cfg.get("prefilter", [])]
    seen: dict = load_json(SEEN_PATH, {})
    queued: dict[str, Post] = {p["url"]: Post(**p) for p in load_json(POSTS_PATH, [])}

    for src in cfg["sources"]:
        if not src.get("enabled", True):
            continue
        fetcher = FETCHERS.get(src["type"])
        if fetcher is None:
            print(f"[{src['name']}] unknown type '{src['type']}', skipped", file=sys.stderr)
            continue
        http.delay = float(src.get("delay_seconds", http.default_delay))
        try:
            found = fetcher(src, http, body_chars)
        except Exception as exc:  # a broken source must not stop the others
            print(f"[{src['name']}] ERROR: {exc}", file=sys.stderr)
            continue
        fresh = [p for p in found if p.url not in seen and p.url not in queued]
        relevant = [p for p in fresh if all(rx.search(f"{p.title} {p.body}") for rx in prefilter)]
        print(f"[{src['name']}] {len(found)} fetched, {len(fresh)} unseen, {len(relevant)} look like interview posts")
        for p in relevant:
            queued[p.url] = p

    for p in list(queued.values())[:15]:
        print(f"  + [{p.source}] {p.title[:90]}")
    if args.dry_run:
        print(f"dry run: {len(queued)} post(s) would be queued")
        return 0
    save_json(POSTS_PATH, [p.to_dict() for p in queued.values()])
    print(f"{len(queued)} post(s) queued in data/interview_posts.json")
    return 0


if __name__ == "__main__":
    sys.exit(main())
