"""Send a Telegram message with the jobs appended by the latest run
(data/last_added.json), or a weekly digest of everything registered in the
last 7 days. Silent when there is nothing to report.

Needs TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID in .env (see README).

Usage:
    python scripts/notify.py            # jobs from the latest run, then clears last_added.json
    python scripts/notify.py --weekly   # digest of the last 7 days (from seen_jobs.json)
    python scripts/notify.py --test     # send a hello message to check the credentials
"""

from __future__ import annotations

import argparse
import html
import os
import sys
from datetime import date, timedelta

import requests
from dotenv import load_dotenv

import common
from common import INTERNSHIP, LAST_ADDED_PATH, ROOT, load_health, load_last_added, load_seen, save_health

TELEGRAM_API = "https://api.telegram.org/bot{token}/sendMessage"
MAX_MESSAGE = 4000  # Telegram caps messages at 4096 chars


def credentials() -> tuple[str, str]:
    load_dotenv(ROOT / ".env")
    token = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
    chat_id = os.environ.get("TELEGRAM_CHAT_ID", "").strip()
    if not token or not chat_id:
        raise SystemExit("TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID missing in .env (see README)")
    return token, chat_id


def send(text: str) -> None:
    token, chat_id = credentials()
    resp = requests.post(
        TELEGRAM_API.format(token=token),
        json={"chat_id": chat_id, "text": text, "parse_mode": "HTML", "disable_web_page_preview": True},
        timeout=20,
    )
    if resp.status_code != 200:
        raise SystemExit(f"telegram error {resp.status_code}: {resp.text[:200]}")


def _line(titulo: str, fuente: str, link: str, extra: str = "") -> str:
    t, f = html.escape(titulo), html.escape(fuente)
    body = f'• <a href="{html.escape(link, quote=True)}">{t}</a> — {f}' if link else f"• {t} — {f}"
    return body + (f" ({html.escape(extra)})" if extra else "")


def _fit(lines: list[str]) -> str:
    """Join whole lines up to the limit: cutting inside an <a> tag makes Telegram
    reject the entire message."""
    kept, size = [], 0
    for i, line in enumerate(lines):
        if size + len(line) + 1 > MAX_MESSAGE - 40:
            kept.append(f"… y {len(lines) - i} más en el tracker")
            break
        kept.append(line)
        size += len(line) + 1
    return "\n".join(kept)


def _job_line(j) -> str:
    extra = ", ".join(x for x in (j.nivel, j.modalidad) if x and x != "No especificado")
    return _line(j.titulo, j.fuente, j.link, extra)


def format_run(jobs) -> str:
    internships = [j for j in jobs if j.nivel == INTERNSHIP]
    others = [j for j in jobs if j.nivel != INTERNSHIP]
    lines = [f"<b>job-radar</b> — {len(jobs)} vacante(s) nueva(s) en el tracker"]
    if internships:
        lines.append("<b>Pasantías</b>")
        lines += [_job_line(j) for j in internships]
        if others:
            lines.append("<b>Otras</b>")
    lines += [_job_line(j) for j in others]
    return _fit(lines)


def format_weekly(seen: dict) -> str | None:
    since = (date.today() - timedelta(days=7)).isoformat()
    recent = [v for v in seen.values() if v.get("registrada", "") >= since]
    if not recent:
        return None
    recent.sort(key=lambda v: v.get("registrada", ""), reverse=True)
    lines = [f"<b>job-radar</b> — resumen semanal: {len(recent)} vacante(s) registrada(s) desde {since}"]
    for v in recent:
        lines.append(_line(v.get("titulo", "?"), v.get("fuente", "?"), v.get("link", ""), v.get("registrada", "")))
    return _fit(lines)


def format_health(health: dict, cfg: dict) -> tuple[str | None, list[str], list[str]]:
    """Message about sources that brought nothing for `health.max_fallas` runs in
    a row (once per outage) and about those that came back. Only enabled,
    non-curated sources of the current config count."""
    max_fallas = int(cfg.get("health", {}).get("max_fallas", 3))
    active = [s["name"] for s in cfg["sources"] if s.get("enabled", True) and s["type"] != "curated"]
    broken = [n for n in active if n in health and health[n]["fallas"] >= max_fallas and not health[n]["avisado"]]
    back = [n for n in active if n in health and health[n].get("recuperada")]
    if not broken and not back:
        return None, [], []
    lines = []
    if broken:
        lines.append("⚠️ <b>job-radar</b> — fuentes sin traer avisos")
        for n in broken:
            e = health[n]
            since = f" Última vez que anduvo: {e['ultima_ok']}." if e["ultima_ok"] else " Nunca anduvo."
            lines.append(f"• <b>{html.escape(n)}</b>: {e['fallas']} corridas seguidas sin traer avisos. "
                         f"Último error: {html.escape(e['error'])}.{since}")
    for n in back:
        lines.append(f"✅ <b>{html.escape(n)}</b> volvió a traer avisos.")
    return "\n".join(lines), broken, back


def report_health() -> None:
    """Send the source-health message, if any; mark it sent only after Telegram accepts it."""
    health = load_health()
    text, broken, back = format_health(health, common.load_config())
    if text is None:
        return
    send(text)
    for n in broken:
        health[n]["avisado"] = True
    for n in back:
        health[n]["recuperada"] = False
    save_health(health)
    print(f"health: {len(broken)} broken source(s) reported, {len(back)} recovered")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--weekly", action="store_true", help="digest of the last 7 days")
    parser.add_argument("--test", action="store_true", help="send a test message")
    args = parser.parse_args()

    if args.test:
        send("<b>job-radar</b> — credenciales OK, listo para avisar.")
        print("test message sent")
        return 0

    if args.weekly:
        text = format_weekly(load_seen())
        if text is None:
            print("nothing registered in the last 7 days, no message sent")
            return 0
        send(text)
        print("weekly digest sent")
        return 0

    jobs = load_last_added()
    if jobs:
        send(format_run(jobs))
        LAST_ADDED_PATH.unlink(missing_ok=True)
        print(f"notified {len(jobs)} job(s)")
    else:
        print("no new jobs in the last run, no message sent")
    report_health()
    return 0


if __name__ == "__main__":
    sys.exit(main())
