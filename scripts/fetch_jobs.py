"""Fetch job postings from every enabled source, normalize them, keep the
relevant ones that were not seen before, and leave them in data/pending_jobs.json
for update_excel.py.

Usage:
    python scripts/fetch_jobs.py            # write pending_jobs.json
    python scripts/fetch_jobs.py --dry-run  # only print what would be queued
    python scripts/fetch_jobs.py --all      # ignore the relevance filter (debug)
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from datetime import date, datetime, timezone
from email.utils import parsedate_to_datetime
from urllib.parse import urljoin

import feedparser
import requests
from bs4 import BeautifulSoup

from common import (
    INTERNSHIP,
    NOT_SPECIFIED,
    NOT_SPECIFIED_F,
    ROOT,
    Job,
    RelevanceFilter,
    clean_text,
    infer_compensacion,
    infer_modalidad,
    infer_nivel,
    infer_stack,
    load_config,
    load_health,
    load_seen,
    record_source_run,
    save_health,
    save_pending,
    tracker_keys,
)


class Http:
    """Tiny polite client: identifiable User-Agent and a delay between calls."""

    def __init__(self, cfg: dict):
        self.session = requests.Session()
        self.session.headers["User-Agent"] = cfg["user_agent"]
        self.delay = float(cfg.get("delay_seconds", 2))
        self.timeout = float(cfg.get("timeout_seconds", 20))
        self._last_call = 0.0

    def get(self, url: str) -> str:
        wait = self.delay - (time.monotonic() - self._last_call)
        if wait > 0:
            time.sleep(wait)
        resp = self.session.get(url, timeout=self.timeout)
        self._last_call = time.monotonic()
        resp.raise_for_status()
        return resp.text

    def get_json(self, url: str, params: dict | None = None):
        wait = self.delay - (time.monotonic() - self._last_call)
        if wait > 0:
            time.sleep(wait)
        resp = self.session.get(url, params=params, timeout=self.timeout)
        self._last_call = time.monotonic()
        resp.raise_for_status()
        return resp.json()


def _fmt_date(dt: datetime | None) -> str:
    return dt.strftime("%d/%m/%Y") if dt else NOT_SPECIFIED


# ---------- source: generic RSS (n8n Community Jobs is a Discourse feed) ----------

def fetch_rss(source: dict, http: Http, vocabulary: list[str]) -> list[Job]:
    """One or more feeds (`url` or `urls`). Options: `company_in_title` splits
    "Company: Title" (We Work Remotely); `modalidad` fixes the modality. The
    feed's <country> or <region>, when present, becomes the location."""
    jobs: list[Job] = []
    links: set[str] = set()
    for url in source.get("urls") or [source["url"]]:
        feed = feedparser.parse(http.get(url))
        for entry in feed.entries:
            title = clean_text(entry.get("title", ""), 200)
            body = clean_text(entry.get("summary", "") or entry.get("description", ""))
            text = f"{title} {body}"
            published = None
            if entry.get("published"):
                try:
                    published = parsedate_to_datetime(entry["published"])
                except (TypeError, ValueError):
                    published = None
            author = entry.get("author") or entry.get("dc_creator")
            empresa = f"{NOT_SPECIFIED_F} (publica: {author})" if author else NOT_SPECIFIED_F
            if source.get("company_in_title") and ": " in title:
                empresa, title = (part.strip() for part in title.split(": ", 1))
            link = entry.get("link", "")
            if link in links:
                continue
            links.add(link)
            jobs.append(
                Job(
                    titulo=title,
                    empresa=empresa,
                    fuente=source["name"],
                    link=link,
                    modalidad=source.get("modalidad") or infer_modalidad(text),
                    ubicacion=clean_text(entry.get("country") or entry.get("region") or "", 150) or NOT_SPECIFIED,
                    nivel=infer_nivel(text),
                    stack=infer_stack(text, vocabulary),
                    compensacion=infer_compensacion(text),
                    fecha_publicacion=_fmt_date(published),
                    resumen=body,
                )
            )
    return jobs


# ---------- source: generic JSON API (Himalayas, Jobicy, Remotive, Working Nomads) ----------
# Declared entirely in sources.yaml: `url`, optional `params`, `queries` +
# `query_param`, paging by `page_param` or by `cursor_field` + `cursor_param`
# (up to `pages` pages per query), `items` (dotted path to the list; omit when
# the response is the list), and `fields` mapping Job fields to item keys.

def _pick(data, path: str):
    for part in path.split("."):
        if not isinstance(data, dict):
            return None
        data = data.get(part)
    return data


def _as_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, (list, tuple)):
        return ", ".join(str(v) for v in value if v not in (None, ""))
    return str(value)


def _parse_date(value) -> datetime | None:
    """Epoch seconds or milliseconds, ISO 8601 (with or without zone) or RFC 822."""
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)) or (isinstance(value, str) and value.strip().isdigit()):
        ts = float(value)
        return datetime.fromtimestamp(ts / 1000 if ts > 1e11 else ts, tz=timezone.utc)
    text = str(value).strip()
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        pass
    try:
        return parsedate_to_datetime(text)
    except (TypeError, ValueError):
        return None


def _level(value: str, text: str) -> str:
    """Map an API's own level ("Senior", ["Entry-level"], "Any") to the tracker's."""
    v = value.lower()
    if re.search(r"senior|\blead\b|staff|principal|expert|director", v):
        return "Senior"
    if re.search(r"\bmid\b|mid-level|intermediate", v):
        return "Semi-junior / Mid"
    if re.search(r"junior|entry|intern|trainee|graduate", v):
        return "Junior"
    return infer_nivel(text)


def _json_job(item: dict, source: dict, vocabulary: list[str]) -> Job | None:
    fields = source["fields"]

    def field(name: str) -> str:
        return _as_text(_pick(item, fields[name])) if fields.get(name) else ""

    titulo = clean_text(field("titulo"), 200)
    if not titulo:
        return None
    resumen = clean_text(field("resumen"))
    text = f"{titulo} {resumen}"
    return Job(
        titulo=titulo,
        empresa=clean_text(field("empresa"), 100) or NOT_SPECIFIED_F,
        fuente=source["name"],
        link=field("link").strip(),
        modalidad=source.get("modalidad") or infer_modalidad(text),
        ubicacion=clean_text(field("ubicacion"), 150) or NOT_SPECIFIED,
        nivel=_level(field("nivel"), text),
        stack=infer_stack(text, vocabulary),
        compensacion=infer_compensacion(text),
        fecha_publicacion=_fmt_date(_parse_date(_pick(item, fields["fecha"]) if fields.get("fecha") else None)),
        resumen=resumen,
    )


def fetch_json(source: dict, http: Http, vocabulary: list[str]) -> list[Job]:
    jobs: list[Job] = []
    links: set[str] = set()
    pages = max(1, int(source.get("pages", 1)))
    for query in source.get("queries") or [None]:
        params = dict(source.get("params") or {})
        if query is not None:
            params[source["query_param"]] = query
        cursor = None
        for page in range(1, pages + 1):
            if source.get("page_param"):
                params[source["page_param"]] = page
            if cursor:
                params[source["cursor_param"]] = cursor
            data = http.get_json(source["url"], dict(params))
            items = _pick(data, source["items"]) if source.get("items") else data
            if not isinstance(items, list) or not items:
                break  # an empty page ends this query
            for item in items:
                job = _json_job(item, source, vocabulary) if isinstance(item, dict) else None
                if job and job.link not in links:
                    links.add(job.link)
                    jobs.append(job)
            if source.get("cursor_field"):
                cursor = _pick(data, source["cursor_field"])
                if not cursor:
                    break
            elif not source.get("page_param"):
                break  # no paging declared: one request per query
    return jobs


# ---------- source: python.org.ar/trabajo (HTML list) ----------

def fetch_pyar(source: dict, http: Http, vocabulary: list[str]) -> list[Job]:
    """List page markup (verified 2026-09):
        div.mp-10 > h4 > a[href=/trabajo/<slug>/] + span.job-date "dd/mm/yyyy, City"
        div.mp-10 > p (excerpt)
        div.mp-10 > div.text-muted > span.label (tags)
    """
    soup = BeautifulSoup(http.get(source["url"]), "lxml")
    jobs: list[Job] = []
    for box in soup.select("div.mp-10"):
        link_tag = box.select_one("h4 a[href]")
        if not link_tag:
            continue
        title = clean_text(link_tag.get_text(), 200)
        date_tag = box.select_one("span.job-date")
        published, city = NOT_SPECIFIED, NOT_SPECIFIED
        if date_tag:
            parts = [p.strip() for p in date_tag.get_text().split(",", 1)]
            if parts and parts[0]:
                published = parts[0]
            if len(parts) > 1 and parts[1]:
                city = parts[1]
        excerpt_tag = box.find("p", recursive=False)
        body = clean_text(excerpt_tag.get_text()) if excerpt_tag else ""
        tags = [clean_text(t.get_text(), 40) for t in box.select("div.text-muted span.label")]
        text = f"{title} {body} {' '.join(tags)}"
        jobs.append(
            Job(
                titulo=title,
                empresa=NOT_SPECIFIED_F,
                fuente=source["name"],
                link=urljoin(source["url"], link_tag["href"]),
                modalidad=infer_modalidad(text),
                ubicacion=city,
                nivel=infer_nivel(text),
                stack=", ".join(tags) if tags else infer_stack(text, vocabulary),
                compensacion=infer_compensacion(text),
                fecha_publicacion=published,
                resumen=body,
            )
        )
    return jobs


# ---------- source: Get on Board (public JSON API v0) ----------

GOB_SENIORITY = {1: "Sin experiencia", 2: "Junior", 3: "Semi-junior / Mid", 4: "Senior", 5: "Experto"}
GOB_MODALITY = {1: "Full time", 2: "Part time", 3: "Freelance", 4: INTERNSHIP}  # /api/v0/modalities
GOB_REMOTE = {"no_remote": "Presencial", "temporarily_remote": "Remoto (temporal)", "fully_remote": "Remoto",
              "remote_local": "Remoto (mismo país)", "hybrid": "Híbrido"}


def fetch_getonbrd(source: dict, http: Http, vocabulary: list[str]) -> list[Job]:
    jobs: list[Job] = []
    for query in source["queries"]:
        data = http.get_json("https://www.getonbrd.com/api/v0/search/jobs",
                             {"query": query, "per_page": source.get("per_query", 20), "expand": '["company"]'})
        for item in data.get("data", []):
            a = item["attributes"]
            company = (((a.get("company") or {}).get("data") or {}).get("attributes") or {}).get("name") or NOT_SPECIFIED_F
            text = f"{a.get('title', '')} {clean_text(a.get('description') or '')} {clean_text(a.get('desirable') or '')}"
            seniority = GOB_SENIORITY.get(((a.get("seniority") or {}).get("data") or {}).get("id"), NOT_SPECIFIED)
            modality = GOB_MODALITY.get(((a.get("modality") or {}).get("data") or {}).get("id"))
            remote = GOB_REMOTE.get(a.get("remote_modality") or "", NOT_SPECIFIED)
            countries = ", ".join(a.get("countries") or []) or NOT_SPECIFIED
            salary = NOT_SPECIFIED
            if a.get("min_salary") or a.get("max_salary"):
                salary = f"USD {a.get('min_salary') or '?'}-{a.get('max_salary') or '?'}/mes"
            published = _fmt_date(datetime.fromtimestamp(a["published_at"])) if a.get("published_at") else NOT_SPECIFIED
            jobs.append(
                Job(
                    titulo=clean_text(a.get("title", ""), 200),
                    empresa=company,
                    fuente=source["name"],
                    link=(item.get("links") or {}).get("public_url") or f"https://www.getonbrd.com/jobs/{item['id']}",
                    modalidad=f"{remote} · {modality}" if modality else remote,
                    ubicacion=countries,
                    nivel=seniority,
                    stack=infer_stack(text, vocabulary),
                    compensacion=salary,
                    fecha_publicacion=published,
                    resumen=clean_text(a.get("description") or "", 600),
                )
            )
    return jobs


# ---------- source: Remote OK (public JSON API; English-heavy, filtered by tag) ----------

def fetch_remoteok(source: dict, http: Http, vocabulary: list[str]) -> list[Job]:
    jobs: list[Job] = []
    for tag in source["tags"]:
        data = http.get_json("https://remoteok.com/api", {"tags": tag})
        for item in data:
            if not isinstance(item, dict) or "position" not in item:  # first element is a legal notice
                continue
            text = f"{item.get('position', '')} {clean_text(item.get('description') or '')} {' '.join(item.get('tags') or [])}"
            salary = NOT_SPECIFIED
            if item.get("salary_min") or item.get("salary_max"):
                salary = f"USD {item.get('salary_min') or '?'}-{item.get('salary_max') or '?'}/año"
            jobs.append(
                Job(
                    titulo=clean_text(item.get("position", ""), 200),
                    empresa=clean_text(item.get("company") or NOT_SPECIFIED_F, 100),
                    fuente=source["name"],
                    link=item.get("url", ""),
                    modalidad="Remoto",
                    ubicacion=item.get("location") or "Sin restricción",
                    nivel=infer_nivel(text),
                    stack=", ".join((item.get("tags") or [])[:8]) or infer_stack(text, vocabulary),
                    compensacion=salary,
                    fecha_publicacion=(item.get("date") or "")[:10].replace("-", "/") and
                                      "/".join(reversed((item.get("date") or "")[:10].split("-"))) or NOT_SPECIFIED,
                    resumen=clean_text(item.get("description") or "", 600),
                )
            )
    return jobs


# ---------- source: Becas IT (Google Sites page, every posting on one URL) ----------

BECAS_COMPANY = "Multinacional de IT (Becas IT)"
_BECAS_TAIL = re.compile(r"\s+[·\-–]\s+Imp\..*$|\s+Activa$")


def fetch_becasit(source: dict, http: Http, vocabulary: list[str]) -> list[Job]:
    """Page markup (verified 2026-11): one <section> per posting, headed by an
    h1/h2 with an id, "✔️ Pasante X · Imp. Multinac. de IT Activa". The id becomes
    the link anchor, because dedup keys on the link and the page URL is shared."""
    soup = BeautifulSoup(http.get(source["url"]), "lxml")
    jobs: list[Job] = []
    for heading in soup.select("h1[id], h2[id]"):
        raw = clean_text(heading.get_text(" "), 200)
        if not re.search(r"\bpasante\b", raw, re.IGNORECASE) or re.search(r"\bcerrada\b", raw, re.IGNORECASE):
            continue
        section = heading.find_parent("section")
        body = clean_text(section.get_text(" ")) if section else ""
        jobs.append(
            Job(
                titulo=_BECAS_TAIL.sub("", re.sub(r"^\W+", "", raw)),
                empresa=BECAS_COMPANY,
                fuente=source["name"],
                link=f"{source['url']}#{heading['id']}",
                modalidad=infer_modalidad(body),
                nivel=INTERNSHIP,
                stack=infer_stack(body, vocabulary),
                resumen=body,
            )
        )
    return jobs


# ---------- source: hand-picked postings (data/curated_jobs.json) ----------

def fetch_curated(source: dict, http: Http, vocabulary: list[str]) -> list[Job]:
    """Postings found by hand on sites the radar cannot read. A JSON list of Job
    fields; they skip the relevance filter but still go through dedup."""
    path = ROOT / source["path"]
    if not path.exists():
        return []
    with path.open(encoding="utf-8") as f:
        return [Job(**{"fuente": source["name"], **item}) for item in json.load(f)]


FETCHERS = {
    "rss": fetch_rss,
    "json": fetch_json,
    "pyar_html": fetch_pyar,
    "getonbrd": fetch_getonbrd,
    "remoteok": fetch_remoteok,
    "becasit_html": fetch_becasit,
    "curated": fetch_curated,
}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--dry-run", action="store_true", help="print, do not write pending_jobs.json")
    parser.add_argument("--all", action="store_true", help="skip the relevance filter")
    args = parser.parse_args()

    cfg = load_config()
    http = Http(cfg["http"])
    relevance = RelevanceFilter(cfg["filter"])
    vocabulary = cfg.get("stack_vocabulary", [])
    seen = set(load_seen()) | tracker_keys()

    health = load_health()
    today = date.today().isoformat()

    pending: list[Job] = []
    for source in cfg["sources"]:
        if not source.get("enabled", True):
            continue
        # A hand-picked list is empty by design most of the time: not a broken source.
        tracked = source["type"] != "curated"
        fetcher = FETCHERS.get(source["type"])
        if fetcher is None:
            print(f"[{source['name']}] unknown type '{source['type']}', skipped", file=sys.stderr)
            if tracked:
                record_source_run(health, source["name"], 0, f"tipo desconocido '{source['type']}'", today)
            continue
        try:
            found = fetcher(source, http, vocabulary)
        except Exception as exc:  # one broken source must not stop the others
            print(f"[{source['name']}] ERROR: {exc}", file=sys.stderr)
            if tracked:
                record_source_run(health, source["name"], 0, str(exc) or type(exc).__name__, today)
            continue
        if tracked:
            record_source_run(health, source["name"], len(found), "", today)
        new = [j for j in found if not j.is_seen(seen)]
        if args.all or source["type"] == "curated":
            kept = new
        else:
            located = source.get("location_filter", False)
            kept = [j for j in new if relevance.is_relevant(j) and (not located or relevance.accepts_location(j))]
        print(f"[{source['name']}] {len(found)} fetched, {len(new)} unseen, {len(kept)} relevant")
        for job in kept:  # a board can return the same posting under two queries/tags
            if not job.is_seen(seen):
                if relevance.is_internship(job):
                    job.nivel = INTERNSHIP
                pending.append(job)
                seen.update({job.key, job.alt_key})

    for job in pending:
        print(f"  + {job.titulo}  ({job.nivel}, {job.modalidad})  {job.link}")

    if args.dry_run:
        print(f"dry run: {len(pending)} job(s) would be queued")
        return 0
    save_health(health)
    save_pending(pending)
    print(f"{len(pending)} job(s) queued in data/pending_jobs.json")
    return 0


if __name__ == "__main__":
    sys.exit(main())
