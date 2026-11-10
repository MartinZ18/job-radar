"""Regenerate the interview-radar notes inside the Obsidian vault from
data/interview_items.json.

Four generated notes live in <vault>/<folder>/ and are rewritten on every run:
questions by section, live coding, most requested projects, behavioural.
The projects note also gets what the tracker's postings ask for (technologies,
frequent combinations, what they say about projects); no LLM call is involved.
Nothing else in the vault is touched, except one link line added to the index
note the first time.

Usage:
    python scripts/update_obsidian.py
"""

from __future__ import annotations

import re
import sys
import zipfile
from collections import Counter, defaultdict
from datetime import date, datetime
from itertools import combinations
from pathlib import Path

import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

from common import ROOT, TRACKER_PATH
from common import load_config as load_sources_config
from extract_interviews import ITEMS_PATH
from fetch_interviews import load_config, load_json

NOTES = {  # tipo -> (file name, title, tags)
    "pregunta": ("Preguntas técnicas (radar).md", "Preguntas técnicas — radar", "[entrevistas, radar, preguntas]"),
    "livecoding": ("Live coding (radar).md", "Live coding y pruebas técnicas — radar", "[entrevistas, radar, livecoding]"),
    "proyecto": ("Proyectos más pedidos (radar).md", "Proyectos más pedidos — radar", "[entrevistas, radar, proyectos]"),
    "comportamental": ("Comportamental (radar).md", "Preguntas de comportamiento — radar", "[entrevistas, radar, comportamental]"),
}
INDEX_LINK = "- [[Radar/Preguntas técnicas (radar)]] · [[Radar/Live coding (radar)]] · [[Radar/Proyectos más pedidos (radar)]] · [[Radar/Comportamental (radar)]] — recopilado automáticamente de foros por job-radar"


IDIOMA_TAG = {"es": "🇪🇸 español", "en-bajo": "🇬🇧 inglés básico", "en": "🇬🇧 inglés"}


def render_item(it: dict, destacado_min: int) -> str:
    star = " ⭐" if it["veces"] >= destacado_min else ""
    idioma = IDIOMA_TAG.get(it.get("idioma", "en"), "")
    lines = [f"- **{it['texto']}**{star} — visto {it['veces']}× · {idioma}"]
    if it.get("detalle"):
        lines.append(f"    - {it['detalle']}")
    for f in it["fuentes"][:3]:
        lines.append(f"    - [{f['origen']}: {f['titulo'][:70]}]({f['url']})")
    return "\n".join(lines)


# ---------- what the tracker's postings ask for ----------

TOP_TECH = 15
TOP_COMBOS = 10
TOP_QUOTES = 10
# Portfolio signals only: "proyecto del sector salud" or "project-based work" talk
# about the contract, not about the candidate's own projects.
PROJECT_WORDS = re.compile(
    r"github|portfolio|portafolio|repositori"
    r"|proyectos? (propios?|personales?|reales?)|\b(personal|side|own|real) projects?",
    re.IGNORECASE)


def read_tracker(path: Path) -> list[dict]:
    """Postings of the "Vacantes" sheet (rows with a title; the footer note has none)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb["Vacantes"]
        out = []
        for r in ws.iter_rows(min_row=2, max_col=13, values_only=True):
            r = [("" if v is None else str(v)) for v in r] + [""] * (13 - len(r))
            if r[1].strip():
                out.append({"fecha": r[0], "titulo": r[1], "link": r[4], "stack": r[8], "notas": r[12]})
        return out
    finally:
        wb.close()


def _parse_date(text: str) -> date | None:
    try:
        return datetime.strptime(text.strip(), "%d/%m/%Y").date()
    except ValueError:
        return None


def _techs(job: dict, vocabulary: list[str]) -> list[str]:
    """Vocabulary terms named in the title or the Stack column, in vocabulary order.
    Notes are left out: free text gives false hits ("make sure" is not Make)."""
    text = f"{job['titulo']} | {job['stack']}"
    return [t for t in vocabulary if re.search(rf"(?<![\w-]){re.escape(t)}(?![\w-])", text, re.IGNORECASE)]


def _md_link(title: str, url: str) -> str:
    title = title.replace("[", "(").replace("]", ")")
    if not url:
        return title
    target = f"<{url}>" if re.search(r"[\s()]", url) else url
    return f"[{title}]({target})"


def render_tracker_section(jobs: list[dict] | None, vocabulary: list[str], error: str = "") -> list[str]:
    out = ["## Lo que piden las vacantes", ""]
    if jobs is None:
        return out + [f"_No se pudo leer el tracker ({error}). Cerralo en Excel y volvé a correr "
                      "`scripts/update_obsidian.py`._", ""]
    if not jobs:
        return out + ["_El tracker todavía no tiene vacantes._", ""]

    dates = sorted(d for d in (_parse_date(j["fecha"]) for j in jobs) if d)
    span = f" ({dates[0]:%d/%m/%Y} – {dates[-1]:%d/%m/%Y})" if dates else ""
    out += [f"Calculado sobre {len(jobs)} vacantes del tracker{span}. Cada vacante cuenta una vez por "
            "tecnología, según su título y la columna Stack.", ""]

    per_job = [_techs(j, vocabulary) for j in jobs]
    order = {t: i for i, t in enumerate(vocabulary)}
    counts = Counter(t for techs in per_job for t in techs)
    out += ["### Tecnologías", "", "| Tecnología | Vacantes | % |", "|---|---|---|"]
    for t, c in sorted(counts.items(), key=lambda x: (-x[1], order[x[0]]))[:TOP_TECH]:
        out.append(f"| {t} | {c} | {round(100 * c / len(jobs))}% |")
    out.append("")

    pairs = Counter(p for techs in per_job for p in combinations(techs, 2))
    frequent = sorted(((p, c) for p, c in pairs.items() if c >= 2),
                      key=lambda x: (-x[1], order[x[0][0]], order[x[0][1]]))[:TOP_COMBOS]
    out += ["### Combinaciones frecuentes", "",
            "Tecnologías que aparecen juntas en la misma vacante: un proyecto que las combine cubre varios avisos.", ""]
    out += [f"- {a} + {b} — {c} vacantes" for (a, b), c in frequent] or ["_Ninguna combinación se repite todavía._"]
    out.append("")

    quotes = []
    for j in jobs:
        for sentence in re.split(r"(?<=[.!?…])\s+", j["notas"]):
            if PROJECT_WORDS.search(sentence):
                quotes.append((_parse_date(j["fecha"]) or date.min, j, sentence.strip()[:200]))
    quotes.sort(key=lambda q: q[0], reverse=True)
    out += ["### Qué dicen sobre proyectos", ""]
    out += [f"- {_md_link(j['titulo'], j['link'])} — «{s}»" for _, j, s in quotes[:TOP_QUOTES]] \
        or ["_Ninguna vacante menciona proyectos, portfolio o GitHub en sus notas._"]
    out.append("")
    return out


def tracker_section() -> list[str]:
    vocabulary = load_sources_config().get("stack_vocabulary", [])
    try:
        jobs = read_tracker(TRACKER_PATH)
    except (OSError, KeyError, zipfile.BadZipFile, InvalidFileException) as e:
        print(f"  cannot read tracker: {e}", file=sys.stderr)
        return render_tracker_section(None, vocabulary, type(e).__name__)
    return render_tracker_section(jobs, vocabulary)


def render_note(tipo: str, items: list[dict], secciones: list[str], destacado_min: int,
                extra: list[str] | None = None) -> str:
    name, title, tags = NOTES[tipo]
    by_section: dict[str, list[dict]] = defaultdict(list)
    for it in items:
        by_section[it["seccion"]].append(it)
    for lst in by_section.values():
        lst.sort(key=lambda x: (-x["veces"], x["texto"].lower()))

    out = [
        "---",
        f"tags: {tags}",
        f"generado: {date.today().isoformat()}",
        "---",
        "",
        f"# {title}",
        "",
        "⬅ [[00 Índice]]",
        "",
        "> [!info] Nota generada",
        f"> La regenera `job-radar` (`scripts/update_obsidian.py`) — no la edites a mano, lo que agregues se pierde. "
        f"{len(items)} ítems; ⭐ = visto en {destacado_min} o más posts. Cada ítem enlaza a sus fuentes.",
        "",
    ]
    out += extra or []
    if tipo == "proyecto":
        ranked = sorted(items, key=lambda x: (-x["veces"], x["texto"].lower()))
        out += ["## Ranking", "", "Proyectos que se mencionan en foros sobre entrevistas.", ""]
        out += [render_item(it, destacado_min) for it in ranked]
        out.append("")
    else:
        ordered = [s for s in secciones if s in by_section] + [s for s in by_section if s not in secciones]
        for sec in ordered:
            out += [f"## {sec}", ""]
            out += [render_item(it, destacado_min) for it in by_section[sec]]
            out.append("")
    if not items:
        out += ["_Todavía no hay ítems de este tipo. Corré `run_interviews.bat`._", ""]
    return "\n".join(out)


def ensure_index_link(index_path: Path) -> bool:
    if not index_path.exists():
        return False
    text = index_path.read_text(encoding="utf-8")
    if "Radar/Preguntas técnicas (radar)" in text:
        return False
    marker = "## Notas\n"
    if marker in text:
        head, tail = text.split(marker, 1)
        # append after the existing bullet list of that section
        lines = tail.split("\n")
        i = 0
        while i < len(lines) and (lines[i].startswith("- ") or lines[i].strip() == ""):
            i += 1
        lines.insert(max(i - 1, 0), INDEX_LINK)
        text = head + marker + "\n".join(lines)
    else:
        text = text.rstrip("\n") + "\n\n## Radar\n\n" + INDEX_LINK + "\n"
    index_path.write_text(text, encoding="utf-8")
    return True


def main() -> int:
    cfg = load_config()
    ob = cfg["obsidian"]
    vault = ROOT / ob["vault_dir"]
    if not vault.exists():
        print(f"vault not found: {vault}", file=sys.stderr)
        return 1
    folder = vault / ob.get("folder", "Radar")
    folder.mkdir(exist_ok=True)
    destacado_min = int(ob.get("min_veces_destacado", 2))

    all_items = list(load_json(ITEMS_PATH, {}).values())
    idiomas = set(ob.get("idiomas") or [])  # empty = show everything
    items = [it for it in all_items if not idiomas or it.get("idioma", "en") in idiomas]
    if idiomas:
        print(f"  idiomas {sorted(idiomas)}: {len(items)} of {len(all_items)} item(s) shown")
    for tipo, (name, _, _) in NOTES.items():
        subset = [it for it in items if it["tipo"] == tipo]
        extra = tracker_section() if tipo == "proyecto" else None
        (folder / name).write_text(render_note(tipo, subset, cfg["secciones"], destacado_min, extra),
                                   encoding="utf-8")
        print(f"  {name}: {len(subset)} item(s)")
    if ensure_index_link(vault / ob.get("index_note", "00 Índice.md")):
        print("  link added to the index note")
    print(f"{len(items)} item(s) written to {folder.name}/")
    return 0


if __name__ == "__main__":
    sys.exit(main())
