"""E2E: run the whole radar chain (fetch -> Excel -> notes -> Telegram) several
times while sources break and come back. Checks that a source that brings
nothing for `health.max_fallas` runs in a row is reported once, that its
recovery is reported once, and that job messages behave as before."""

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import requests

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

import common  # noqa: E402
import fetch_jobs  # noqa: E402
import notify  # noqa: E402
import update_excel  # noqa: E402
import write_job_notes  # noqa: E402
from tests.test_internships_e2e import BECAS_URL, build_tracker, gob_item  # noqa: E402

BECAS_OK = """<html><body>
<section><h2 id="h.py_l">✔️ Pasante Python Backend Developer · Imp. Multinac. de IT Activa</h2>
<p>Modalidad remota. Python.</p></section></body></html>"""
BECAS_BROKEN = "<html><body><p>Sitio en mantenimiento</p></body></html>"  # markup changed: 0 postings, no error


class SourceHealthE2E(unittest.TestCase):
    def setUp(self):
        self.scratch = Path(tempfile.mkdtemp(prefix="radar_health_"))
        self.addCleanup(shutil.rmtree, self.scratch, ignore_errors=True)
        self.tracker = self.scratch / "tracker.xlsx"
        build_tracker(self.tracker)
        curated = self.scratch / "curated.json"
        curated.write_text("[]", encoding="utf-8")  # nothing picked by hand: not a broken source

        cfg = common.load_config()
        cfg["http"]["delay_seconds"] = 0
        cfg["health"] = {"max_fallas": 3}
        cfg["sources"] = [
            {"name": "Becas IT", "type": "becasit_html", "url": BECAS_URL, "enabled": True},
            {"name": "Get on Board", "type": "getonbrd", "queries": ["python junior"], "per_query": 20,
             "enabled": True},
            {"name": "Selección manual", "type": "curated", "path": str(curated), "enabled": True},
            {"name": "Vieja", "type": "rss", "url": "https://example.com/rss", "enabled": False},
        ]
        self.cfg = cfg

        self.becas = BECAS_OK
        self.gob_up = True
        self.gob_items = [gob_item("junior-1", "Junior Python Developer")]
        self.telegram_up = True
        self.sent = []

        pending = self.scratch / "pending.json"
        last = self.scratch / "last.json"
        patches = [
            mock.patch.object(common, "SEEN_PATH", self.scratch / "seen.json"),
            mock.patch.object(common, "PENDING_PATH", pending),
            mock.patch.object(common, "LAST_ADDED_PATH", last),
            mock.patch.object(common, "HEALTH_PATH", self.scratch / "health.json"),
            mock.patch.object(common, "TRACKER_PATH", self.tracker),
            mock.patch.object(common, "load_config", return_value=cfg),
            mock.patch.object(update_excel, "PENDING_PATH", pending),
            mock.patch.object(update_excel, "TRACKER_PATH", self.tracker),
            mock.patch.object(notify, "LAST_ADDED_PATH", last),
            mock.patch.object(write_job_notes, "NOTES_DIR", self.scratch / "Vacantes"),
            mock.patch.object(fetch_jobs, "load_config", return_value=cfg),
            mock.patch.object(fetch_jobs.Http, "get", lambda _, url: self.becas),
            mock.patch.object(fetch_jobs.Http, "get_json", self.fake_get_json),
            mock.patch.object(notify, "send", self.fake_send),
        ]
        for patch in patches:
            patch.start()
            self.addCleanup(patch.stop)

    def fake_get_json(self, url, params=None):
        if not self.gob_up:
            raise requests.ConnectionError("getonbrd.com unreachable")
        return {"data": self.gob_items}

    def fake_send(self, text):
        if not self.telegram_up:
            raise SystemExit("telegram error 502: Bad Gateway")
        self.sent.append(text)

    def run_chain(self):
        """Same order and failure handling as run_radar.bat."""
        for module in (fetch_jobs, update_excel, write_job_notes, notify):
            with mock.patch.object(sys, "argv", [module.__name__]):
                try:
                    module.main()
                except SystemExit:
                    if module is not notify:
                        raise

    def new_messages(self, before):
        return self.sent[before:]

    def test_broken_source_is_reported_once_and_its_recovery_too(self):
        self.run_chain()  # everything works: one job message (Becas + GoB postings)
        self.assertEqual(len(self.sent), 1)
        self.assertNotIn("sin traer avisos", self.sent[0])

        self.becas = BECAS_BROKEN
        self.gob_up = False  # Get on Board fails only twice: below the limit
        for _ in range(2):
            n = len(self.sent)
            self.run_chain()
            self.assertEqual(self.new_messages(n), [])  # no jobs, no alert yet: silent as before
        self.gob_up = True
        self.gob_items.append(gob_item("junior-2", "Junior Backend Developer"))

        n = len(self.sent)
        self.run_chain()  # third run in a row with nothing from Becas IT
        msgs = self.new_messages(n)
        self.assertEqual(len(msgs), 2)  # the new job, then the alert
        self.assertIn("Junior Backend Developer", msgs[0])
        alert = msgs[1]
        self.assertIn("Becas IT", alert)
        self.assertIn("3 corridas seguidas", alert)
        self.assertIn("0 avisos", alert)
        self.assertNotIn("Get on Board", alert)  # back before reaching the limit
        self.assertNotIn("Selección manual", alert)  # empty by design
        self.assertNotIn("Vieja", alert)  # disabled

        n = len(self.sent)
        self.run_chain()  # still broken: no repeated alert
        self.assertEqual(self.new_messages(n), [])

        self.becas = BECAS_OK
        n = len(self.sent)
        self.run_chain()  # Becas IT is back (its posting was already seen: no job message)
        msgs = self.new_messages(n)
        self.assertEqual(len(msgs), 1)
        self.assertIn("Becas IT", msgs[0])
        self.assertIn("volvió a traer avisos", msgs[0])

        n = len(self.sent)
        self.run_chain()
        self.assertEqual(self.new_messages(n), [])

    def test_alert_survives_a_telegram_failure_and_errors_are_shown(self):
        self.gob_up = False
        for _ in range(2):
            self.run_chain()
        self.telegram_up = False
        self.assertEqual(len(self.sent), 1)  # first run: the Becas IT posting
        self.run_chain()  # third failure, but Telegram is down: nothing sent
        self.assertEqual(len(self.sent), 1)
        self.telegram_up = True
        n = len(self.sent)
        self.run_chain()  # fourth run: the pending alert goes out now, once
        msgs = self.new_messages(n)
        self.assertEqual(len(msgs), 1)
        self.assertIn("Get on Board", msgs[0])
        self.assertIn("4 corridas seguidas", msgs[0])
        self.assertIn("getonbrd.com unreachable", msgs[0])
        self.assertNotIn("Becas IT", msgs[0])

    def test_dry_run_leaves_health_untouched(self):
        self.gob_up = False
        with mock.patch.object(sys, "argv", ["fetch_jobs", "--dry-run"]):
            fetch_jobs.main()
        self.assertFalse((self.scratch / "health.json").exists())

    def test_health_file_records_last_success(self):
        self.run_chain()
        self.gob_up = False
        self.run_chain()
        health = json.loads((self.scratch / "health.json").read_text(encoding="utf-8"))
        self.assertEqual(health["Get on Board"]["fallas"], 1)
        self.assertIn("unreachable", health["Get on Board"]["error"])
        self.assertTrue(health["Get on Board"]["ultima_ok"])
        self.assertEqual(health["Becas IT"]["fallas"], 0)
        self.assertNotIn("Selección manual", health)


if __name__ == "__main__":
    unittest.main()
