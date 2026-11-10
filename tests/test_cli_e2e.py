"""E2E for radar.py, the single entry point.

`run` must behave exactly like the old run_radar.bat chain, so every test of
the internships chain is run again through `radar.py run --log`. On top of
that: a tracker open in Excel is reported by Telegram and nothing is lost,
`status` summarizes everything without printing secrets, `setup` creates what
is missing and never overwrites, and the log rotates."""

import contextlib
import hashlib
import io
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from collections import Counter
from pathlib import Path
from unittest import mock

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT))

import common  # noqa: E402
import create_tracker  # noqa: E402
import radar  # noqa: E402
import update_excel  # noqa: E402
import write_job_notes  # noqa: E402
import tests.test_internships_e2e as chain  # noqa: E402

SECRET_KEYS = ("TELEGRAM_BOT_TOKEN", "TELEGRAM_CHAT_ID", "GROQ_API_KEY")


def capture(argv):
    out = io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(out):
        code = radar.main(argv)
    return code, out.getvalue()


class RunThroughCli(chain.InternshipsE2E):
    """The whole internships chain, driven by `radar.py run --log`."""

    def setUp(self):
        super().setUp()
        self.log = self.scratch / "radar.log"
        self.env = self.scratch / ".env"
        for patch in (
            mock.patch.object(radar, "LOG_PATH", self.log),
            mock.patch.object(radar, "ENV_PATH", self.env),
            mock.patch.dict(os.environ, {k: "" for k in SECRET_KEYS}),
        ):
            patch.start()
            self.addCleanup(patch.stop)

    def run_chain(self):
        self.assertEqual(radar.main(["run", "--log"]), 0)

    def test_log_has_every_step(self):
        self.run_chain()
        log = self.log.read_text(encoding="utf-8")
        self.assertRegex(log, r"==== \d{4}-\d{2}-\d{2} \d{2}:\d{2} run ====")
        self.assertIn("[Becas IT]", log)           # fetch
        self.assertIn("row(s) appended", log)       # Excel
        self.assertIn("note(s) written", log)       # notes
        self.assertIn("notified", log)              # Telegram

    def test_open_excel_is_reported_and_nothing_is_lost(self):
        before = self.rows()
        with mock.patch.object(update_excel.openpyxl.Workbook, "save", side_effect=PermissionError("locked")):
            self.assertEqual(radar.main(["run"]), 1)
        self.assertEqual(self.rows(), before)
        self.assertEqual(len(self.sent), 1)
        self.assertIn("abierto", self.sent[0])
        self.assertIn("Cerralo", self.sent[0])
        self.assertFalse(list(self.notes.glob("*.md")) if self.notes.exists() else [])

        self.run_chain()  # closed again: the same postings go in now
        self.assertEqual(len(self.rows()), 14)
        self.assertEqual(len(self.sent), 2)
        self.assertIn("Pasante Python Backend Developer", self.sent[1])

    def test_status(self):
        self.env.write_text("TELEGRAM_BOT_TOKEN=secret-token-123\nTELEGRAM_CHAT_ID=987654\nGROQ_API_KEY=\n",
                            encoding="utf-8")
        self.run_chain()
        code, out = capture(["status"])
        self.assertEqual(code, 0)
        rows = self.rows()
        self.assertIn(f"{len(rows)} vacantes", out)
        for estado, n in Counter(r[11] for r in rows.values()).items():
            self.assertIn(f"{estado} {n}", out)
        self.assertRegex(out, r"Última corrida: \d{4}-\d{2}-\d{2} \d{2}:\d{2}")
        self.assertRegex(out, r"✓ Becas IT")
        self.assertIn("Selección manual", out)
        self.assertIn("Telegram ✓", out)
        self.assertIn("Groq ✗", out)
        self.assertIn(f"{len(list(self.notes.glob('*.md')))} notas", out)
        for secret in ("secret-token-123", "987654"):
            self.assertNotIn(secret, out)

    def test_status_reports_a_broken_source(self):
        self.run_chain()
        health = common.load_health()
        health["Becas IT"].update(fallas=3, error="0 avisos (¿cambió la página?)")
        common.save_health(health)
        code, out = capture(["status"])
        self.assertRegex(out, r"✗ Becas IT — 3 corridas sin avisos")

    def test_log_rotates(self):
        self.log.write_text("x" * (radar.MAX_LOG_BYTES + 10), encoding="utf-8")
        self.run_chain()
        self.assertEqual((self.log.parent / "radar.log.1").stat().st_size, radar.MAX_LOG_BYTES + 10)
        self.assertLess(self.log.stat().st_size, radar.MAX_LOG_BYTES)
        self.assertIn("[Becas IT]", self.log.read_text(encoding="utf-8"))


class SetupCli(unittest.TestCase):
    def setUp(self):
        self.scratch = Path(tempfile.mkdtemp(prefix="radar_setup_"))
        self.addCleanup(shutil.rmtree, self.scratch, ignore_errors=True)
        self.tracker = self.scratch / "tracker.xlsx"
        self.env = self.scratch / ".env"
        self.example = self.scratch / ".env.example"
        self.example.write_text("TELEGRAM_BOT_TOKEN=\nTELEGRAM_CHAT_ID=\nGROQ_API_KEY=\n", encoding="utf-8")
        self.vault = self.scratch / "vault"
        self.calls = []
        for patch in (
            mock.patch.object(common, "TRACKER_PATH", self.tracker),
            mock.patch.object(common, "HEALTH_PATH", self.scratch / "health.json"),
            mock.patch.object(create_tracker, "TRACKER_PATH", self.tracker),
            mock.patch.object(radar, "ENV_PATH", self.env),
            mock.patch.object(radar, "ENV_EXAMPLE_PATH", self.example),
            mock.patch.object(radar, "VAULT_DIR", self.vault),
            mock.patch.object(radar, "LOG_PATH", self.scratch / "radar.log"),
            mock.patch.object(write_job_notes, "NOTES_DIR", self.vault / "Vacantes"),
            mock.patch.object(radar.subprocess, "run", self.fake_run),
            mock.patch.dict(os.environ, {k: "" for k in SECRET_KEYS}),
        ):
            patch.start()
            self.addCleanup(patch.stop)

    def fake_run(self, cmd, **kwargs):
        self.calls.append(cmd)
        return subprocess.CompletedProcess(cmd, 0, "", "")

    def test_setup_creates_what_is_missing_and_never_overwrites(self):
        code, out = capture(["setup"])
        self.assertEqual(code, 0)
        self.assertEqual(openpyxl.load_workbook(self.tracker).sheetnames, ["Vacantes", "Resumen"])
        self.assertEqual(self.env.read_text(encoding="utf-8"), self.example.read_text(encoding="utf-8"))
        self.assertTrue((self.vault / "00 Índice.md").exists())
        self.assertIn("TELEGRAM_BOT_TOKEN", out)  # tells what is still missing
        self.assertEqual(self.calls, [])          # no scheduled tasks unless asked

        self.env.write_text("TELEGRAM_BOT_TOKEN=mine\n", encoding="utf-8")
        (self.vault / "00 Índice.md").write_text("mi índice", encoding="utf-8")
        update_excel.append_jobs([common.Job(titulo="Mi vacante", empresa="A", fuente="x", link="https://x.io/1")])
        digest = hashlib.sha256(self.tracker.read_bytes()).hexdigest()

        code, out = capture(["setup"])
        self.assertEqual(code, 0)
        self.assertEqual(self.env.read_text(encoding="utf-8"), "TELEGRAM_BOT_TOKEN=mine\n")
        self.assertEqual((self.vault / "00 Índice.md").read_text(encoding="utf-8"), "mi índice")
        self.assertEqual(hashlib.sha256(self.tracker.read_bytes()).hexdigest(), digest)
        self.assertNotIn("mine", out)

    def test_setup_schedule_on_windows(self):
        with mock.patch.object(radar.sys, "platform", "win32"):
            code, _ = capture(["setup", "--schedule"])
        self.assertEqual(code, 0)
        self.assertEqual(len(self.calls), 2)
        radar_task, weekly_task = (" ".join(map(str, c)) for c in self.calls)
        self.assertIn("/TN job-radar ", radar_task + " ")
        self.assertIn("MON,WED,FRI", radar_task)
        self.assertIn(str(radar.ROOT / "run_radar.bat"), radar_task)
        self.assertIn("/TN job-radar-weekly", weekly_task)
        self.assertIn(str(radar.ROOT / "run_weekly.bat"), weekly_task)

    def test_setup_schedule_elsewhere_prints_cron(self):
        with mock.patch.object(radar.sys, "platform", "linux"):
            code, out = capture(["setup", "--schedule"])
        self.assertEqual(code, 0)
        self.assertEqual(self.calls, [])
        self.assertIn("0 9 * * 1,3,5", out)
        self.assertIn("radar.py run --log", out)


if __name__ == "__main__":
    unittest.main()
