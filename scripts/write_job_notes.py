"""Write one short Obsidian note per job appended in the latest run
(data/last_added.json) into <vault>/Vacantes/<empresa>-<yyyy-mm-dd>.md.

Runs between update_excel.py and notify.py: notify deletes last_added.json.
It only creates files: an existing note is never overwritten (a numbered name
is used instead), and a job that already has a note is skipped, so running it
twice is harmless.

Usage:
    python scripts/write_job_notes.py                        # write notes for the latest run
    python scripts/write_job_notes.py --dry-run              # show the file names only
    python scripts/write_job_notes.py --backfill 2026-11-01  # notes for tracker rows registered since that date
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date, datetime
from pathlib import Path

import openpyxl
import yaml

import common
from common import NOT_SPECIFIED, ROOT, Job, load_last_added

INTERVIEWS_CONFIG = ROOT / "config" / "interviews.yaml"


def _vault_dir() -> Path:
    with INTERVIEWS_CONFIG.open(encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    return ROOT / cfg["obsidian"]["vault_dir"]


NOTES_DIR = _vault_dir() / "Vacantes"

UNSAFE_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}
KEY_LINE = re.compile(r'^clave: "([0-9a-f]+)"$', re.MULTILINE)
MAX_NAME = 60


def safe_name(job: Job) -> str:
    """Company name usable as a Windows file name; the title when the company is unknown."""
    empresa = job.empresa.strip()
    base = job.titulo if not empresa or empresa.startswith(NOT_SPECIFIED[:-1]) else empresa
    base = re.sub(r"\s+", " ", UNSAFE_CHARS.sub(" ", base)).strip(" .")[:MAX_NAME].strip(" .")
    if not base:
        base = "Vacante"
    if base.upper() in RESERVED:
        base += "_"
    return base


def existing_keys(folder: Path) -> set[str]:
    keys = set()
    for path in folder.glob("*.md"):
        try:
            match = KEY_LINE.search(path.read_text(encoding="utf-8"))
        except (OSError, UnicodeDecodeError):
            continue
        if match:
            keys.add(match.group(1))
    return keys


def q(value: str) -> str:
    """Double-quoted YAML scalar (JSON strings are valid YAML)."""
    return json.dumps(value or "", ensure_ascii=False)


def render(job: Job, registrada: str) -> str:
    front = {
        "titulo": job.titulo, "empresa": job.empresa, "fuente": job.fuente, "link": job.link,
        "modalidad": job.modalidad, "ubicacion": job.ubicacion, "nivel": job.nivel, "stack": job.stack,
        "compensacion": job.compensacion, "publicada": job.fecha_publicacion, "registrada": registrada,
        "estado": job.estado_postulacion, "clave": job.key,
    }
    lines = ["---", "tags: [vacante]"] + [f"{k}: {q(v)}" for k, v in front.items()] + ["---", ""]
    lines.append(f"# {job.titulo}")
    lines.append("")
    aviso = f" · [Ver aviso](<{job.link}>)" if job.link else ""
    lines.append(f"**{job.empresa}** · {job.fuente} · {job.modalidad} · {job.nivel}{aviso}")
    if job.stack and job.stack != NOT_SPECIFIED:
        lines.append(f"Stack: {job.stack}")
    if job.resumen:
        lines += ["", "> " + re.sub(r"\s+", " ", job.resumen).strip()]
    lines += [
        "",
        "## Preparación",
        "",
        "- [ ] Leer el aviso completo y anotar requisitos que no estén cubiertos en las notas del vault",
        "- [ ] Repasar [[Backend Fundamentos]], [[IA y LLM]] o [[Automatización n8n]] según el stack",
        "- [ ] Postular y actualizar el estado en el tracker",
        "",
        "## Apuntes",
        "",
        "",
        "---",
        "Volver al [[00 Índice]]",
        "",
    ]
    return "\n".join(lines)


def write_note(folder: Path, job: Job, registrada: str) -> Path:
    """Create the note under the first free name; never touches an existing file."""
    stem = f"{safe_name(job)}-{registrada}"
    content = render(job, registrada)
    n = 1
    while True:
        path = folder / (f"{stem}.md" if n == 1 else f"{stem}-{n}.md")
        try:
            with path.open("x", encoding="utf-8", newline="\n") as f:
                f.write(content)
            return path
        except FileExistsError:
            n += 1


def _iso_date(text: str) -> date:
    try:
        return date.fromisoformat(text)
    except ValueError:
        raise argparse.ArgumentTypeError(f"expected a date like 2026-11-01, got '{text}'")


def _row_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        return datetime.strptime(str(value or "").strip(), "%d/%m/%Y").date()
    except ValueError:
        return None


def jobs_from_tracker(since: date) -> list[tuple[Job, str]]:
    """Rows of the "Vacantes" sheet registered on or after `since`, with the
    date each one was registered. Rows without a title (the footer note) or
    without a readable date are skipped."""
    wb = openpyxl.load_workbook(common.TRACKER_PATH, read_only=True, data_only=True)
    try:
        out = []
        for r in wb["Vacantes"].iter_rows(min_row=2, max_col=13, values_only=True):
            r = list(r) + [None] * (13 - len(r))
            cell = [("" if v is None else str(v).strip()) for v in r]
            registered = _row_date(r[0])
            if not cell[1] or registered is None or registered < since:
                continue
            job = Job(
                titulo=cell[1], empresa=cell[2] or common.NOT_SPECIFIED_F, fuente=cell[3], link=cell[4],
                modalidad=cell[5] or NOT_SPECIFIED, ubicacion=cell[6] or NOT_SPECIFIED,
                nivel=cell[7] or NOT_SPECIFIED, stack=cell[8] or NOT_SPECIFIED,
                compensacion=cell[9] or NOT_SPECIFIED, estado_publicacion=cell[10] or "Abierta",
                estado_postulacion=cell[11] or "No postulado", resumen=cell[12],
            )
            out.append((job, registered.isoformat()))
        return out
    finally:
        wb.close()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--backfill", type=_iso_date, metavar="YYYY-MM-DD",
                        help="write notes for tracker rows registered since this date that have none")
    args = parser.parse_args()

    if args.backfill:
        try:
            pairs = jobs_from_tracker(args.backfill)
        except (OSError, KeyError) as e:
            print(f"cannot read tracker: {e}", file=sys.stderr)
            return 1
    else:
        today = date.today().isoformat()
        pairs = [(job, today) for job in load_last_added()]
    if not pairs:
        print(f"no tracker rows since {args.backfill}" if args.backfill
              else "no new jobs in the last run, no notes written")
        return 0

    folder = NOTES_DIR
    try:
        folder.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        print(f"cannot create {folder}: {e}", file=sys.stderr)
        return 1

    done = existing_keys(folder)
    written = failed = 0
    for job, registrada in pairs:
        if job.key in done:
            continue
        if args.dry_run:
            print(f"  would write {safe_name(job)}-{registrada}.md")
            continue
        try:
            path = write_note(folder, job, registrada)
        except OSError as e:
            failed += 1
            print(f"  cannot write note for '{job.titulo}': {e}", file=sys.stderr)
            continue
        done.add(job.key)
        written += 1
        print(f"  note: {path.name}")
    print(f"{written} note(s) written in {folder.name}/" + (f", {failed} failed" if failed else ""))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
