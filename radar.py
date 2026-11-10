"""job-radar: one command for everything.

    python radar.py setup [--schedule]   create what is missing (tracker, .env, vault); never overwrites
    python radar.py run [--log]          fetch -> Excel -> notes -> Telegram (what Task Scheduler runs)
    python radar.py weekly [--log]       weekly Telegram digest + interview radar
    python radar.py interviews [--log]   interview radar only (forums -> Groq -> Obsidian)
    python radar.py status               tracker, last run, sources, configuration, notes

--log appends the output to data/radar.log, which rotates to radar.log.1 past 1 MB.
"""

from __future__ import annotations

import argparse
import contextlib
import importlib
import os
import shutil
import subprocess
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "scripts"))

import openpyxl  # noqa: E402
from dotenv import dotenv_values  # noqa: E402

import common  # noqa: E402
import create_tracker  # noqa: E402
import write_job_notes  # noqa: E402

ENV_PATH = ROOT / ".env"
ENV_EXAMPLE_PATH = ROOT / ".env.example"
LOG_PATH = common.DATA_DIR / "radar.log"
VAULT_DIR = write_job_notes.NOTES_DIR.parent
MAX_LOG_BYTES = 1_000_000

# (script, arguments, stop the chain if it fails) — same order and rules as the old .bat files
INTERVIEWS = [("fetch_interviews", [], False), ("extract_interviews", [], False), ("update_obsidian", [], False)]
CHAINS = {
    "run": [("fetch_jobs", [], True), ("update_excel", [], True), ("write_job_notes", [], False), ("notify", [], False)],
    "weekly": [("notify", ["--weekly"], False)] + INTERVIEWS,
    "interviews": INTERVIEWS,
}

INDEX_NOTE = """---
tags: [entrevistas, indice]
---

# Índice — preparación de entrevistas

Vault de Obsidian que mantiene job-radar. Abrí la carpeta `obsidian_vault` como vault.

## Notas

- `Vacantes/` — una nota por vacante nueva; son tuyas, editalas libremente
- Agregá acá tus propias notas de estudio; job-radar no las toca
"""


# ---------- running the scripts ----------

def _step(name: str, args: list[str]) -> int:
    module = importlib.import_module(name)
    saved, sys.argv = sys.argv, [name, *args]
    try:
        code = module.main()
    except SystemExit as exc:  # e.g. notify.py without .env: report it, let the chain decide
        if isinstance(exc.code, str):
            print(exc.code, file=sys.stderr)
            code = 1
        else:
            code = exc.code
    except Exception as exc:  # a crashing step must not take the log or the rest of the chain with it
        print(f"{name}: {type(exc).__name__}: {exc}", file=sys.stderr)
        code = 1
    finally:
        sys.argv = saved
    return code or 0


def _chain(steps) -> int:
    for name, args, required in steps:
        code = _step(name, args)
        if code:
            print(f"{name} FAILED (exit {code})", file=sys.stderr)
            if required:
                return 1
    return 0


@contextlib.contextmanager
def _logging(command: str, enabled: bool):
    if not enabled:
        yield
        return
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    if LOG_PATH.exists() and LOG_PATH.stat().st_size > MAX_LOG_BYTES:
        LOG_PATH.replace(LOG_PATH.with_name(LOG_PATH.name + ".1"))
    with LOG_PATH.open("a", encoding="utf-8") as log, \
            contextlib.redirect_stdout(log), contextlib.redirect_stderr(log):
        print(f"==== {datetime.now():%Y-%m-%d %H:%M} {command} ====")
        yield


# ---------- status ----------

def _env_present(key: str, env: dict) -> bool:
    return bool((os.environ.get(key) or env.get(key) or "").strip())


def _last_run() -> str:
    """Last `run` header in the log (old .bat headers have no command name)."""
    if not LOG_PATH.exists():
        return ""
    last = ""
    with LOG_PATH.open(encoding="utf-8", errors="replace") as f:
        for line in f:
            if line.startswith("==== "):
                text = line.strip().strip("=").strip()
                if not text.endswith(("weekly", "interviews")):
                    last = text.removesuffix(" run")
    return last


def status() -> int:
    print("job-radar — estado\n")

    tracker = common.TRACKER_PATH
    if not tracker.exists():
        print("Tracker: no existe — corré  python radar.py setup")
    else:
        try:
            wb = openpyxl.load_workbook(tracker, read_only=True, data_only=True)
            try:
                rows = [r for r in wb["Vacantes"].iter_rows(min_row=2, max_col=13, values_only=True) if r[1]]
            finally:
                wb.close()
            estados = Counter(str(r[11] or "Sin estado") for r in rows)
            detail = " · ".join(f"{e} {n}" for e, n in estados.most_common())
            print(f"Tracker: {tracker.name} — {len(rows)} vacantes ({detail})")
        except (OSError, KeyError) as exc:
            print(f"Tracker: no se pudo leer ({type(exc).__name__}); ¿está abierto en Excel?")

    print(f"Última corrida: {_last_run() or 'todavía no hay registro'}")

    cfg = common.load_config()
    health = common.load_health()
    print(f"\nFuentes ({len(cfg['sources'])}):")
    for source in cfg["sources"]:
        name, entry = source["name"], health.get(source["name"])
        if not source.get("enabled", True):
            print(f"  · {name} — desactivada")
        elif source["type"] == "curated":
            print(f"  · {name} — carga manual, no se controla")
        elif entry is None:
            print(f"  · {name} — todavía no corrió")
        elif entry["fallas"] == 0:
            print(f"  ✓ {name} — anduvo {entry['ultima_ok']}")
        else:
            told = " (ya avisado por Telegram)" if entry.get("avisado") else ""
            print(f"  ✗ {name} — {entry['fallas']} corridas sin avisos; último error: {entry['error']}{told}")

    env = dotenv_values(ENV_PATH) if ENV_PATH.exists() else {}
    telegram = _env_present("TELEGRAM_BOT_TOKEN", env) and _env_present("TELEGRAM_CHAT_ID", env)
    groq = _env_present("GROQ_API_KEY", env)
    print("\nConfiguración: "
          + ("Telegram ✓" if telegram else "Telegram ✗ (faltan TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID en .env)")
          + " · " + ("Groq ✓" if groq else "Groq ✗ (falta GROQ_API_KEY en .env; el radar de entrevistas no corre)"))

    notes = write_job_notes.NOTES_DIR
    count = len(list(notes.glob("*.md"))) if notes.exists() else 0
    print(f"Notas por vacante: {count} notas en {notes.name}/")

    try:
        import extract_interviews
        import fetch_interviews
        queue = fetch_interviews.load_json(fetch_interviews.POSTS_PATH, [])
        items = fetch_interviews.load_json(extract_interviews.ITEMS_PATH, {})
        print(f"Radar de entrevistas: {len(items)} ítems guardados, {len(queue)} posts en cola")
    except (OSError, ValueError) as exc:
        print(f"Radar de entrevistas: no se pudo leer su estado ({type(exc).__name__})")
    return 0


# ---------- setup ----------

def _schedule() -> None:
    if sys.platform.startswith("win"):
        tasks = [
            ("job-radar", ROOT / "run_radar.bat", ["/SC", "WEEKLY", "/D", "MON,WED,FRI", "/ST", "09:00"]),
            ("job-radar-weekly", ROOT / "run_weekly.bat", ["/SC", "WEEKLY", "/D", "FRI", "/ST", "18:00"]),
        ]
        for name, script, when in tasks:
            result = subprocess.run(["schtasks", "/Create", "/F", "/TN", name, "/TR", f'"{script}"', *when],
                                    capture_output=True, text=True)
            mark = "✓" if result.returncode == 0 else "✗"
            print(f"{mark} Tarea programada {name}: {(result.stdout or result.stderr).strip()}")
        return
    print("Agregá estas líneas con  crontab -e :")
    print(f"0 9 * * 1,3,5  cd {ROOT} && python3 radar.py run --log")
    print(f"0 18 * * 5     cd {ROOT} && python3 radar.py weekly --log")


def setup(schedule: bool) -> int:
    tracker = create_tracker.TRACKER_PATH
    if tracker.exists():
        print(f"✓ Tracker: {tracker.name} ya existe, no se toca")
    else:
        tracker.parent.mkdir(parents=True, exist_ok=True)
        create_tracker.build().save(tracker)
        print(f"✓ Tracker: {tracker.name} creado")

    if ENV_PATH.exists():
        print("✓ .env: ya existe, no se toca")
    elif ENV_EXAMPLE_PATH.exists():
        shutil.copyfile(ENV_EXAMPLE_PATH, ENV_PATH)
        print("✓ .env: creado a partir de .env.example")
    env = dotenv_values(ENV_PATH) if ENV_PATH.exists() else {}
    missing = [k for k in ("TELEGRAM_BOT_TOKEN", "TELEGRAM_CHAT_ID", "GROQ_API_KEY") if not _env_present(k, env)]
    if missing:
        print(f"  Completá en .env: {', '.join(missing)} (el README explica cómo conseguirlos)")

    index = VAULT_DIR / "00 Índice.md"
    VAULT_DIR.mkdir(parents=True, exist_ok=True)
    if index.exists():
        print("✓ Vault: ya existe, no se toca")
    else:
        index.write_text(INDEX_NOTE, encoding="utf-8")
        print(f"✓ Vault: creado en {VAULT_DIR.name}/ (abrí obsidian_vault como vault en Obsidian)")

    if "TELEGRAM_BOT_TOKEN" not in missing and "TELEGRAM_CHAT_ID" not in missing:
        ok = _step("notify", ["--test"]) == 0
        print("✓ Telegram: mensaje de prueba enviado" if ok else "✗ Telegram: no se pudo enviar el mensaje de prueba")

    if schedule:
        _schedule()
    else:
        print("\nPara que corra solo:  python radar.py setup --schedule")
    print()
    return status()


# ---------- entry point ----------

def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="radar.py", description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)
    setup_parser = sub.add_parser("setup", help="create what is missing; never overwrites")
    setup_parser.add_argument("--schedule", action="store_true",
                              help="also create the scheduled tasks (Windows) or print the cron lines")
    for name, text in (("run", "fetch -> Excel -> notes -> Telegram"),
                       ("weekly", "weekly digest + interview radar"),
                       ("interviews", "interview radar only")):
        chain_parser = sub.add_parser(name, help=text)
        chain_parser.add_argument("--log", action="store_true", help="append output to data/radar.log")
    sub.add_parser("status", help="tracker, last run, sources, configuration")
    args = parser.parse_args(argv)

    if args.command == "setup":
        return setup(args.schedule)
    if args.command == "status":
        return status()
    with _logging(args.command, args.log):
        return _chain(CHAINS[args.command])


if __name__ == "__main__":
    sys.exit(main())
