"""Append queued jobs (data/pending_jobs.json) to the "Vacantes" sheet of the
tracker, copying the cell style of the existing rows, then mark them as seen.

The "Resumen" sheet is never touched: its COUNTIF formulas recalculate on open.
The script only appends; it never edits or deletes existing rows.

Usage:
    python scripts/update_excel.py            # append pending jobs
    python scripts/update_excel.py --dry-run  # show what would be appended
    python scripts/update_excel.py --manual   # type one job by hand (e.g. from a LinkedIn alert)
"""

from __future__ import annotations

import argparse
import sys
from copy import copy
from datetime import date

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from common import (
    NOT_SPECIFIED,
    NOT_SPECIFIED_F,
    PENDING_PATH,
    TRACKER_PATH,
    Job,
    load_pending,
    load_seen,
    save_last_added,
    save_pending,
    save_seen,
    tracker_keys,
)

SHEET = "Vacantes"
STYLE_TEMPLATE_ROW = 2  # first data row; its formatting is copied to new rows
TITLE_COL = 2  # column B, used to find the last real data row
NUM_COLS = 13


def last_data_row(ws: Worksheet) -> int:
    """ws.max_row counts formatted-but-empty rows; walk back to the last title."""
    for row in range(ws.max_row, 1, -1):
        if ws.cell(row=row, column=TITLE_COL).value not in (None, ""):
            return row
    return 1


def copy_row_style(ws: Worksheet, src_row: int, dst_row: int) -> None:
    for col in range(1, NUM_COLS + 1):
        src = ws.cell(row=src_row, column=col)
        dst = ws.cell(row=dst_row, column=col)
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.fill = copy(src.fill)
        dst.number_format = src.number_format
    if ws.row_dimensions[src_row].height:
        ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height


def detach_footer_note(ws: Worksheet, after_row: int) -> dict | None:
    """The tracker keeps a merged help note (A:M) right under the data. Take it
    out so rows can be appended, and return what is needed to put it back."""
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row > after_row and rng.min_col == 1 and rng.max_col >= NUM_COLS:
            cell = ws.cell(row=rng.min_row, column=1)
            note = {"value": cell.value, "font": copy(cell.font), "alignment": copy(cell.alignment),
                    "fill": copy(cell.fill), "height": ws.row_dimensions[rng.min_row].height}
            ws.unmerge_cells(str(rng))
            for col in range(1, NUM_COLS + 1):
                ws.cell(row=rng.min_row, column=col).value = None
            return note
    return None


def attach_footer_note(ws: Worksheet, row: int, note: dict) -> None:
    cell = ws.cell(row=row, column=1, value=note["value"])
    cell.font, cell.alignment, cell.fill = note["font"], note["alignment"], note["fill"]
    if note["height"]:
        ws.row_dimensions[row].height = note["height"]
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=NUM_COLS)


FORMULA_PREFIXES = ("=", "+", "-", "@", chr(9), chr(13))


def write_cell(ws: Worksheet, row: int, col: int, value: str):
    """Write external data as plain text: a value starting with =, +, -, @, tab or CR
    is forced to a string cell (and flagged quotePrefix) so Excel never evaluates it."""
    cell = ws.cell(row=row, column=col, value=value)
    if isinstance(value, str) and value.startswith(FORMULA_PREFIXES):
        cell.data_type = "s"
    return cell


def job_to_row(job: Job) -> list[str]:
    notas = job.notas or f"Publicada {job.fecha_publicacion}. Cargada automáticamente por job-radar."
    if job.resumen and not job.notas:
        notas += f" {job.resumen[:180].rstrip()}…"
    return [
        date.today().strftime("%d/%m/%Y"),  # Fecha registrada (stored as text, like the existing rows)
        job.titulo,
        job.empresa,
        job.fuente,
        job.link,
        job.modalidad,
        job.ubicacion,
        job.nivel,
        job.stack,
        job.compensacion,
        job.estado_publicacion,
        job.estado_postulacion,
        notas,
    ]


def append_jobs(jobs: list[Job], dry_run: bool = False) -> int:
    if not jobs:
        print("nothing to append")
        return 0
    if not TRACKER_PATH.exists():
        print(f"tracker not found: {TRACKER_PATH}", file=sys.stderr)
        return 1

    wb = openpyxl.load_workbook(TRACKER_PATH)
    if SHEET not in wb.sheetnames:
        print(f"sheet '{SHEET}' not found in tracker", file=sys.stderr)
        return 1
    ws = wb[SHEET]
    row = last_data_row(ws)
    note = None if dry_run else detach_footer_note(ws, row)

    for job in jobs:
        row += 1
        print(f"  row {row}: {job.titulo}  [{job.fuente}]")
        if dry_run:
            continue
        for col, value in enumerate(job_to_row(job), start=1):
            write_cell(ws, row, col, value)
        copy_row_style(ws, STYLE_TEMPLATE_ROW, row)
        for col, value in enumerate(job_to_row(job), start=1):
            if isinstance(value, str) and value.startswith(FORMULA_PREFIXES):
                ws.cell(row=row, column=col).quotePrefix = True

    if dry_run:
        print(f"dry run: {len(jobs)} row(s) would be appended")
        return 0
    if note:
        attach_footer_note(ws, row + 1, note)

    try:
        wb.save(TRACKER_PATH)
    except PermissionError:
        print("cannot save: close the tracker in Excel/LibreOffice and run again", file=sys.stderr)
        return 1

    seen = load_seen()
    today = date.today().isoformat()
    for job in jobs:
        seen[job.key] = {"titulo": job.titulo, "fuente": job.fuente, "link": job.link, "registrada": today}
    save_seen(seen)
    save_last_added(jobs)  # consumed by notify.py
    print(f"{len(jobs)} row(s) appended to '{SHEET}'")
    return 0


def prompt_manual_job() -> Job:
    def ask(label: str, default: str = "") -> str:
        raw = input(f"{label}{f' [{default}]' if default else ''}: ").strip()
        return raw or default

    return Job(
        titulo=ask("Título"),
        empresa=ask("Empresa", NOT_SPECIFIED_F),
        fuente=ask("Fuente", "LinkedIn (búsqueda guardada + alerta)"),
        link=ask("Link"),
        modalidad=ask("Modalidad", NOT_SPECIFIED),
        ubicacion=ask("Ubicación", NOT_SPECIFIED),
        nivel=ask("Nivel", "Junior"),
        stack=ask("Stack", NOT_SPECIFIED),
        compensacion=ask("Compensación", NOT_SPECIFIED),
        fecha_publicacion=ask("Fecha de publicación (dd/mm/yyyy)", date.today().strftime("%d/%m/%Y")),
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--manual", action="store_true", help="enter one job interactively")
    args = parser.parse_args()

    if args.manual:
        job = prompt_manual_job()
        if not job.titulo:
            print("title is required", file=sys.stderr)
            return 1
        if job.is_seen(set(load_seen()) | tracker_keys()):
            print("already in the tracker (same link or same title/company/source)", file=sys.stderr)
            return 1
        return append_jobs([job], dry_run=args.dry_run)

    jobs = load_pending()
    code = append_jobs(jobs, dry_run=args.dry_run)
    if code == 0 and jobs and not args.dry_run:
        save_pending([])
        PENDING_PATH.unlink(missing_ok=True)
    return code


if __name__ == "__main__":
    sys.exit(main())
