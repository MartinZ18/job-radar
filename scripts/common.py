"""Shared pieces for job-radar: paths, config, the Job model, relevance filter
and the seen-jobs store used for deduplication."""

from __future__ import annotations

import hashlib
import html
import json
import re
from dataclasses import asdict, dataclass
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = ROOT / "config" / "sources.yaml"
DATA_DIR = ROOT / "data"
SEEN_PATH = DATA_DIR / "seen_jobs.json"
PENDING_PATH = DATA_DIR / "pending_jobs.json"
LAST_ADDED_PATH = DATA_DIR / "last_added.json"  # rows appended by the latest update_excel run
TRACKER_PATH = ROOT / "tracker_vacantes_ia_backend.xlsx"
HEALTH_PATH = DATA_DIR / "source_health.json"  # per-source run streaks, read by notify.py

NOT_SPECIFIED = "No especificado"
NOT_SPECIFIED_F = "No especificada"
INTERNSHIP = "Pasantía"  # value of the "Nivel" column for internships


@dataclass
class Job:
    """One normalized job posting. Field names mirror the tracker columns."""

    titulo: str
    empresa: str
    fuente: str
    link: str
    modalidad: str = NOT_SPECIFIED
    ubicacion: str = NOT_SPECIFIED
    nivel: str = NOT_SPECIFIED
    stack: str = NOT_SPECIFIED
    compensacion: str = NOT_SPECIFIED
    fecha_publicacion: str = NOT_SPECIFIED  # dd/mm/yyyy
    resumen: str = ""  # short plain-text excerpt, only used for "Notas"
    estado_publicacion: str = "Abierta"
    estado_postulacion: str = "No postulado"
    notas: str = ""  # replaces the generated "Notas" text (hand-curated rows)

    @property
    def key(self) -> str:
        """Primary dedup key: the link is the most stable identity a posting
        has (topic id, slug). Falls back to `alt_key` when there is no link."""
        basis = self.link.strip().lower()
        return hashlib.sha1(basis.encode("utf-8")).hexdigest() if basis else self.alt_key

    @property
    def alt_key(self) -> str:
        """Secondary dedup key (title + company + source), for rows loaded by
        hand with a generic link such as the listing page."""
        basis = f"{self.titulo}|{self.empresa}|{self.fuente}".lower()
        return hashlib.sha1(basis.encode("utf-8")).hexdigest()

    def is_seen(self, seen: set[str]) -> bool:
        return self.key in seen or self.alt_key in seen

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "Job":
        return cls(**d)


def load_config() -> dict:
    with CONFIG_PATH.open(encoding="utf-8") as f:
        return yaml.safe_load(f)


# ---------- relevance filter ----------

def _compile(patterns: list[str]) -> list[re.Pattern]:
    return [re.compile(p, re.IGNORECASE) for p in patterns]


class RelevanceFilter:
    """Keep a posting when its TITLE carries a tech signal, nothing in
    title+body is excluded, the structured level is not senior, and it can be
    done from home: remote, at home, or not stated as on-site.

    Internships follow their own rule: their titles rarely name a technology and
    their bodies often mention senior mentors, so only the title can exclude them
    and the software signal may come from the body."""

    def __init__(self, cfg: dict):
        self.tech = _compile(cfg.get("tech", []))
        self.exclude = _compile(cfg.get("exclude", []))
        self.exclude_levels = {s.lower() for s in cfg.get("exclude_levels", [])}
        internship = cfg.get("internship", {})
        self.internship_title = _compile(internship.get("title", []))
        self.internship_software = _compile(internship.get("software", []))
        self.internship_exclude = _compile(internship.get("exclude_title", []))
        self.home = _compile(cfg.get("home", []))

    def is_internship(self, job: Job) -> bool:
        return INTERNSHIP.lower() in job.modalidad.lower() or any(p.search(job.titulo) for p in self.internship_title)

    def _internship_fits(self, job: Job) -> bool:
        if any(p.search(job.titulo) for p in self.internship_exclude):
            return False
        text = f"{job.titulo}\n{job.resumen}"
        if not any(p.search(text) for p in self.internship_software):
            return False
        return self._reachable(job, text)

    def _reachable(self, job: Job, text: str) -> bool:
        """Remote, at home, or silent about its modality: that can be asked, a
        missed posting cannot be recovered."""
        if re.search(r"remot", job.modalidad, re.IGNORECASE):
            return True
        if any(p.search(f"{job.ubicacion}\n{text}") for p in self.home):
            return True
        return not re.search(r"presencial|h[ií]brido", job.modalidad, re.IGNORECASE)

    def is_relevant(self, job: Job) -> bool:
        if job.nivel.lower() in self.exclude_levels:
            return False
        if self.is_internship(job):
            return self._internship_fits(job)
        text = f"{job.titulo}\n{job.resumen}"
        if any(p.search(text) for p in self.exclude):
            return False
        return any(p.search(job.titulo) for p in self.tech) and self._reachable(job, text)


def fix_mojibake(text: str) -> str:
    """Some APIs (Remote OK) double-encode UTF-8: 'DiseÃ±ador' -> 'Diseñador'."""
    if "Ã" not in text and "â" not in text:
        return text
    try:
        return text.encode("latin-1").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return text


# ---------- inference helpers (shared by every source) ----------

def infer_nivel(text: str) -> str:
    t = text.lower()
    if re.search(r"semi[- ]?senior|\bssr\b|semi[- ]?junior", t):
        return "Semi-junior / Mid"
    if re.search(r"\bjunior\b|\bjr\.?\b|trainee|pasant[ií]a|practicante", t):
        return "Junior"
    return NOT_SPECIFIED


def infer_modalidad(text: str) -> str:
    t = text.lower()
    if re.search(r"h[ií]brid", t):
        return "Híbrido"
    if re.search(r"\bremot[oa]\b|\bremote\b|home ?office", t):
        return "Remoto"
    if re.search(r"presencial|on[- ]?site", t):
        return "Presencial"
    return NOT_SPECIFIED


_MONEY = re.compile(
    r"(?:USD|US\$|U\$S|ARS|\$|€)\s?\d[\d.,]*(?:\s?(?:k|K))?(?:\s?(?:-|a|to)\s?(?:USD|\$)?\s?\d[\d.,]*(?:k|K)?)?"
    r"(?:\s?(?:/|por|per)\s?(?:mes|hora|hour|month|h))?",
)


def infer_compensacion(text: str) -> str:
    m = _MONEY.search(text)
    return m.group(0).strip() if m else NOT_SPECIFIED


def infer_stack(text: str, vocabulary: list[str]) -> str:
    found = []
    for term in vocabulary:
        if re.search(rf"(?<![\w-]){re.escape(term)}(?![\w-])", text, re.IGNORECASE):
            found.append(term)
    return ", ".join(found) if found else NOT_SPECIFIED


def clean_text(html_or_text: str, limit: int = 600) -> str:
    """Strip tags/whitespace; keep a bounded excerpt for filtering and notes."""
    text = html.unescape(re.sub(r"<[^>]+>", " ", fix_mojibake(html_or_text)))
    text = re.sub(r"\s+", " ", text).strip()
    return text[:limit]


# ---------- seen / pending stores ----------

def tracker_keys() -> set[str]:
    """Dedup keys of the rows already in the tracker, so anything loaded by hand
    (or before this tool existed) is never appended twice."""
    if not TRACKER_PATH.exists():
        return set()
    import openpyxl  # local import: fetch_jobs.py should not need it otherwise

    wb = openpyxl.load_workbook(TRACKER_PATH, read_only=True)
    if "Vacantes" not in wb.sheetnames:
        return set()
    keys = set()
    for titulo, empresa, fuente, link in wb["Vacantes"].iter_rows(min_row=2, min_col=2, max_col=5, values_only=True):
        if titulo:
            job = Job(str(titulo), str(empresa or ""), str(fuente or ""), str(link or ""))
            keys.update({job.key, job.alt_key})
    wb.close()
    return keys


def load_seen() -> dict[str, dict]:
    if not SEEN_PATH.exists():
        return {}
    with SEEN_PATH.open(encoding="utf-8") as f:
        return json.load(f) or {}


def save_seen(seen: dict[str, dict]) -> None:
    DATA_DIR.mkdir(exist_ok=True)
    with SEEN_PATH.open("w", encoding="utf-8") as f:
        json.dump(seen, f, ensure_ascii=False, indent=2, sort_keys=True)


def load_pending() -> list[Job]:
    if not PENDING_PATH.exists():
        return []
    with PENDING_PATH.open(encoding="utf-8") as f:
        return [Job.from_dict(d) for d in json.load(f)]


def save_pending(jobs: list[Job]) -> None:
    DATA_DIR.mkdir(exist_ok=True)
    with PENDING_PATH.open("w", encoding="utf-8") as f:
        json.dump([j.to_dict() for j in jobs], f, ensure_ascii=False, indent=2)


def load_last_added() -> list[Job]:
    if not LAST_ADDED_PATH.exists():
        return []
    with LAST_ADDED_PATH.open(encoding="utf-8") as f:
        return [Job.from_dict(d) for d in json.load(f)]


def save_last_added(jobs: list[Job]) -> None:
    DATA_DIR.mkdir(exist_ok=True)
    with LAST_ADDED_PATH.open("w", encoding="utf-8") as f:
        json.dump([j.to_dict() for j in jobs], f, ensure_ascii=False, indent=2)


# ---------- source health ----------
# A source "fails" a run when it raises or brings 0 postings (fetched, before
# dedup and filtering): a changed page or a dead API, not a quiet week.

def load_health() -> dict[str, dict]:
    if not HEALTH_PATH.exists():
        return {}
    with HEALTH_PATH.open(encoding="utf-8") as f:
        return json.load(f) or {}


def save_health(health: dict[str, dict]) -> None:
    HEALTH_PATH.parent.mkdir(parents=True, exist_ok=True)
    with HEALTH_PATH.open("w", encoding="utf-8") as f:
        json.dump(health, f, ensure_ascii=False, indent=2, sort_keys=True)


def record_source_run(health: dict[str, dict], name: str, fetched: int, error: str, today: str) -> None:
    entry = health.setdefault(name, {"fallas": 0, "ultima_ok": "", "error": "", "avisado": False,
                                     "recuperada": False})
    if fetched > 0 and not error:
        if entry["avisado"]:
            entry["recuperada"] = True  # notify.py reports it once, then clears it
        entry.update(fallas=0, ultima_ok=today, error="", avisado=False)
    else:
        entry["fallas"] += 1
        entry["error"] = (error or "0 avisos (¿cambió la página?)")[:150]
