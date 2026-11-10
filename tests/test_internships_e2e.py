"""E2E: run fetch -> Excel -> Telegram against a scratch tracker with faked HTTP
responses, using the real config/sources.yaml filter. Checks that internships
get through, that what does not suit the profile stays out, and that the rest of
the radar behaves as before."""

import json
import re
import shutil
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

import common  # noqa: E402
import fetch_jobs  # noqa: E402
import notify  # noqa: E402
import update_excel  # noqa: E402
import write_job_notes  # noqa: E402

BECAS_URL = "https://becasit.itechcareer.com/pasant%C3%ADas"
FOOTER = "Nota de ayuda"
RESUMEN_FORMULA = "=COUNTIF(Vacantes!$L$2:$L$200,A4)"

BECAS_HTML = """<html><body>
<section><h1 id="h.top">Pasantías IT</h1></section>
<section><h2 id="h.py_l">✔️ Pasante Python Backend Developer · Imp. Multinac. de IT Activa</h2>
<p>Empresa con sede en Buenos Aires. Modalidad remota. Python y POO.</p></section>
<section><h2 id="h.ia_l">✔️ Pasante Analista IA · Imp. Multinac. de IT Activa</h2>
<p>IA generativa, n8n.</p></section>
<section><h2 id="h.hd_l">✔️ Pasante Soporte Help Desk · Imp. Multinac. de IT Activa</h2>
<p>Soporte a usuarios de sistemas.</p></section>
<section><h1 id="h.dba_l">Pasante DBA SQL</h1><p>SQL Server, bases de datos.</p></section>
<section><h2 id="h.old_l">✔️ Pasante Java Developer · Imp. Multinac. Cerrada</h2><p>Java.</p></section>
<section><h1 id="h.cv_l">Envianos tu CV a</h1><p>pasantias@example.com</p></section>
</body></html>"""


def gob_item(job_id, title, *, modality=1, seniority=2, remote="fully_remote", countries=None, body="Python y APIs"):
    return {
        "id": job_id,
        "links": {"public_url": f"https://www.getonbrd.com/jobs/{job_id}"},
        "attributes": {
            "title": title,
            "description": body,
            "modality": {"data": {"id": modality}},
            "seniority": {"data": {"id": seniority}},
            "remote_modality": remote,
            "countries": countries or [],
            "published_at": 1793000000,
            "company": {"data": {"attributes": {"name": "Acme"}}},
        },
    }


GOB_ITEMS = [
    gob_item("generic-title", "Analista de Datos", modality=4, seniority=1),
    gob_item("trainee-remote", "Trainee Python", body="Con mentoría de un senior manager del equipo mobile"),
    gob_item("chile-onsite", "Practicante Agentes IA", modality=4, remote="no_remote", countries=["Chile"]),
    gob_item("chile-hybrid", "Pasante Desarrollo Backend", modality=4, remote="hybrid", countries=["Chile"]),
    gob_item("ba-hybrid", "Pasante Python Palermo", modality=4, remote="hybrid", countries=["Argentina"]),
    gob_item("ba-onsite", "Trainee Backend CABA", remote="no_remote", countries=["Argentina"]),
    gob_item("mendoza-onsite", "Pasante Python", modality=4, remote="no_remote", countries=["Argentina"],
             body="Oficinas en Godoy Cruz, Mendoza. Python y APIs"),
    gob_item("mendoza-hybrid", "Trainee Datos Mendoza", remote="hybrid", countries=["Argentina"]),
    gob_item("remote-argentina", "Pasante Backend Remoto", modality=4, countries=["Argentina"]),
    gob_item("sales", "Practicante Área Comercial", modality=4, body="Empresa de software"),
    gob_item("not-software", "Practicante de Cocina", modality=4, body="Preparación de platos"),
    gob_item("senior-level", "Intern Lead Python", modality=4, seniority=4),
    gob_item("junior-ok", "Junior Python Developer"),
    gob_item("unclear-city", "Analista Python Junior", remote="", countries=["Córdoba capital"]),
    gob_item("junior-chile", "Junior Backend Santiago", remote="hybrid", countries=["Chile"]),
    gob_item("junior-ba", "Junior Java Developer", remote="no_remote", countries=["Argentina"]),
    gob_item("junior-mendoza", "Junior Full-Stack Mendoza", remote="hybrid", countries=["Argentina"]),
    gob_item("junior-body-excluded", "Backend Developer Junior", body="Reportarás a un senior"),
    gob_item("freelance", "Python Developer", modality=3),
]

CURATED = [{
    "titulo": "Pasante de Desarrollo Backend e IA",
    "empresa": "Retail Ejemplo",
    "link": "https://example.com/retail",
    "modalidad": "Híbrido",
    "nivel": "Pasantía",
    "notas": "Verificado a mano.",
}, {
    "titulo": "Propuesta directa",
    "empresa": "Startup Ejemplo",
    "link": "https://example.com/startup",
    "estado_publicacion": "Desconocido",
    "estado_postulacion": "Postulado",
    "notas": "Mail enviado.",
}]


def build_tracker(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Vacantes"
    ws.append(["Fecha", "Titulo", "Empresa", "Fuente", "Link", "Modalidad", "Ubicacion",
               "Nivel", "Stack", "Compensacion", "Estado pub", "Estado post", "Notas"])
    ws.append(["01/11/2026", "Pasante Analista IA", "Multinacional de IT (Becas IT)", "Becas IT", BECAS_URL,
               "", "", "", "", "", "Abierta", "Postulado", "Ya enviada"])
    ws.cell(row=4, column=1, value=FOOTER)
    ws.merge_cells("A4:M4")
    resumen = wb.create_sheet("Resumen")
    resumen["A4"] = "Postulado"
    resumen["B4"] = RESUMEN_FORMULA
    wb.save(path)


def fake_get(self, url):
    return BECAS_HTML if "becasit" in url else "<rss></rss>"


def fake_get_json(self, url, params=None):
    return {"data": GOB_ITEMS} if "getonbrd" in url else []


class InternshipsE2E(unittest.TestCase):
    def setUp(self):
        self.scratch = Path(tempfile.mkdtemp(prefix="radar_e2e_"))
        self.addCleanup(shutil.rmtree, self.scratch, ignore_errors=True)
        self.tracker = self.scratch / "tracker.xlsx"
        build_tracker(self.tracker)
        curated = self.scratch / "curated.json"
        curated.write_text(json.dumps(CURATED, ensure_ascii=False), encoding="utf-8")

        cfg = common.load_config()
        for source in cfg["sources"]:
            if source["type"] == "curated":
                source["path"] = str(curated)
        cfg["http"]["delay_seconds"] = 0
        cfg["filter"]["home"] = ["mendoza"]  # the fixtures assume this home, whatever config/ says

        pending = self.scratch / "pending.json"
        self.notes = self.scratch / "vault" / "Vacantes"
        self.sent = []
        patches = [
            mock.patch.object(common, "DATA_DIR", self.scratch),
            mock.patch.object(common, "SEEN_PATH", self.scratch / "seen.json"),
            mock.patch.object(common, "PENDING_PATH", pending),
            mock.patch.object(common, "LAST_ADDED_PATH", self.scratch / "last.json"),
            mock.patch.object(common, "HEALTH_PATH", self.scratch / "health.json"),
            mock.patch.object(common, "TRACKER_PATH", self.tracker),
            mock.patch.object(update_excel, "PENDING_PATH", pending),
            mock.patch.object(update_excel, "TRACKER_PATH", self.tracker),
            mock.patch.object(notify, "LAST_ADDED_PATH", self.scratch / "last.json"),
            mock.patch.object(fetch_jobs, "load_config", return_value=cfg),
            mock.patch.object(fetch_jobs.Http, "get", fake_get),
            mock.patch.object(fetch_jobs.Http, "get_json", fake_get_json),
            mock.patch.object(notify, "send", self.sent.append),
            mock.patch.object(write_job_notes, "NOTES_DIR", self.notes),
        ]
        for patch in patches:
            patch.start()
            self.addCleanup(patch.stop)

    def run_chain(self):
        for module in (fetch_jobs, update_excel, write_job_notes, notify):
            with mock.patch.object(sys, "argv", [module.__name__]):
                self.assertEqual(module.main(), 0)

    def rows(self):
        ws = openpyxl.load_workbook(self.tracker)["Vacantes"]
        return {r[1]: r for r in ws.iter_rows(min_row=2, values_only=True) if r[1]}

    def test_full_chain(self):
        self.run_chain()
        rows = self.rows()

        self.assertEqual(set(rows), {
            "Pasante Analista IA",  # already in the tracker: not added twice
            "Analista de Datos", "Trainee Python", "Junior Python Developer", "Python Developer",
            "Pasante Python", "Trainee Datos Mendoza", "Pasante Backend Remoto", "Junior Full-Stack Mendoza",
            "Analista Python Junior",
            "Pasante Python Backend Developer", "Pasante DBA SQL",
            "Pasante de Desarrollo Backend e IA", "Propuesta directa",
        })

        # Get on Board labels its internship modality correctly, and freelance too.
        self.assertIn("Pasantía", rows["Analista de Datos"][5])
        self.assertEqual(rows["Analista de Datos"][7], "Pasantía")
        self.assertIn("Freelance", rows["Python Developer"][5])
        self.assertEqual(rows["Python Developer"][7], "Junior")
        self.assertEqual(rows["Junior Python Developer"][7], "Junior")
        self.assertEqual(rows["Trainee Python"][7], "Pasantía")

        # Becas IT postings share one page: each row needs its own anchor.
        python_row = rows["Pasante Python Backend Developer"]
        self.assertEqual(python_row[4], f"{BECAS_URL}#h.py_l")
        self.assertEqual(rows["Pasante DBA SQL"][4], f"{BECAS_URL}#h.dba_l")
        self.assertEqual(python_row[3], "Becas IT")
        self.assertEqual(python_row[5], "Remoto")
        self.assertEqual(python_row[7], "Pasantía")

        # Curated rows keep what was recorded by hand.
        startup = rows["Propuesta directa"]
        self.assertEqual((startup[10], startup[11]), ("Desconocido", "Postulado"))
        self.assertEqual(startup[12], "Mail enviado.")
        retail = rows["Pasante de Desarrollo Backend e IA"]
        self.assertEqual((retail[3], retail[10], retail[11]), ("Selección manual", "Abierta", "No postulado"))

        # The workbook around the data stays intact.
        wb = openpyxl.load_workbook(self.tracker)
        self.assertEqual(wb["Resumen"]["B4"].value, RESUMEN_FORMULA)
        ws = wb["Vacantes"]
        self.assertEqual(ws.cell(row=len(rows) + 2, column=1).value, FOOTER)
        self.assertEqual(rows["Pasante Analista IA"][12], "Ya enviada")

        # Telegram lists internships first.
        self.assertEqual(len(self.sent), 1)
        message = self.sent[0]
        self.assertLess(message.index("Pasante Python Backend Developer"), message.index("Junior Python Developer"))
        self.assertIn("Pasantías", message)

    def test_second_run_adds_nothing_and_stays_silent(self):
        self.run_chain()
        before = self.rows()
        self.run_chain()
        self.assertEqual(self.rows(), before)
        self.assertEqual(len(self.sent), 1)

    def note_links(self):
        links = {}
        for path in self.notes.glob("*.md"):
            match = re.search(r'^link: "(.*)"$', path.read_text(encoding="utf-8"), re.MULTILINE)
            links[path.name] = match.group(1) if match else None
        return links

    def test_job_notes(self):
        today = date.today().isoformat()
        # A note written by hand earlier today under the name the radar would pick.
        self.notes.mkdir(parents=True)
        hand_note = self.notes / f"Acme-{today}.md"
        hand_note.write_text("mis apuntes de la entrevista", encoding="utf-8")

        self.run_chain()
        rows = self.rows()
        links = self.note_links()

        # One note per appended row (the row already in the tracker gets none),
        # and the hand-written note is left alone.
        self.assertEqual(hand_note.read_text(encoding="utf-8"), "mis apuntes de la entrevista")
        appended = [r for t, r in rows.items() if t != "Pasante Analista IA"]
        self.assertEqual(sorted(v for v in links.values() if v), sorted(r[4] for r in appended))

        # Several Acme postings the same day get numbered names instead of overwriting.
        acme = [n for n in links if n.startswith("Acme-")]
        self.assertIn(f"Acme-{today}-2.md", acme)
        self.assertEqual(len(acme), sum(1 for r in appended if r[2] == "Acme") + 1)

        for name in links:
            self.assertNotRegex(name, r'[<>:"/\\|?*]')

        note = next(p for p in self.notes.glob("*.md")
                    if f"{BECAS_URL}#h.py_l" in p.read_text(encoding="utf-8")).read_text(encoding="utf-8")
        self.assertTrue(note.startswith("---\n"))
        self.assertIn('titulo: "Pasante Python Backend Developer"', note)
        self.assertIn('nivel: "Pasantía"', note)
        self.assertIn(f'registrada: "{today}"', note)
        self.assertIn("[[00 Índice]]", note)

        # Telegram still goes out after the notes step.
        self.assertEqual(len(self.sent), 1)

    def test_job_notes_rerun_and_unsafe_names(self):
        jobs = [
            common.Job(titulo='Dev: "Python" / IA', empresa="CON", fuente="x", link="https://example.com/a"),
            common.Job(titulo="Backend", empresa="Foo: Bar/Baz?. ", fuente="x", link="https://example.com/b"),
            common.Job(titulo="Data", empresa=common.NOT_SPECIFIED_F, fuente="x", link="https://example.com/c"),
        ]
        common.save_last_added(jobs)
        for _ in range(2):  # Telegram failed: last_added.json survives and the step runs again
            with mock.patch.object(sys, "argv", ["write_job_notes"]):
                self.assertEqual(write_job_notes.main(), 0)

        links = self.note_links()
        self.assertEqual(sorted(links.values()),
                         ["https://example.com/a", "https://example.com/b", "https://example.com/c"])
        for name in links:
            self.assertNotRegex(name, r'[<>:"/\\|?*]')
            self.assertNotEqual(name.split("-")[0].strip().upper(), "CON")
            self.assertFalse(name.startswith("No especificad"))
        name_a = next(n for n, link in links.items() if link.endswith("/a"))
        self.assertIn('titulo: "Dev: \\"Python\\" / IA"', (self.notes / name_a).read_text(encoding="utf-8"))

    def test_job_notes_backfill_from_tracker(self):
        today = date.today()
        self.run_chain()
        appended = {r[4] for t, r in self.rows().items() if t != "Pasante Analista IA"}
        # Lose some notes, as happened to the rows added before the notes step existed.
        notes = sorted(self.notes.glob("*.md"))
        for path in notes[:3]:
            path.unlink()
        kept = {p.name for p in notes[3:]}

        with mock.patch.object(sys, "argv", ["write_job_notes", "--backfill", today.isoformat()]):
            self.assertEqual(write_job_notes.main(), 0)
        links = self.note_links()
        # Every row registered since that date has exactly one note again; the
        # older row (01/11/2026) is left out, and surviving notes are untouched.
        self.assertEqual(sorted(links.values()), sorted(appended))
        self.assertTrue(kept <= set(links))

        # Going further back picks the older row up, dated when it was registered.
        with mock.patch.object(sys, "argv", ["write_job_notes", "--backfill", "2026-11-01"]):
            self.assertEqual(write_job_notes.main(), 0)
        links = self.note_links()
        self.assertEqual(len(links), len(appended) + 1)
        older = next(n for n, link in links.items() if link == BECAS_URL)
        self.assertTrue(older.endswith("-2026-11-01.md"))
        text = (self.notes / older).read_text(encoding="utf-8")
        self.assertIn('registrada: "2026-11-01"', text)
        self.assertIn('estado: "Postulado"', text)
        self.assertIn("Ya enviada", text)

        # A bad date is refused without writing anything.
        with mock.patch.object(sys, "argv", ["write_job_notes", "--backfill", "11/2026"]), \
                self.assertRaises(SystemExit) as exit_:
            write_job_notes.main()
        self.assertEqual(exit_.exception.code, 2)
        self.assertEqual(len(self.note_links()), len(appended) + 1)

    def test_long_telegram_message_keeps_valid_html(self):
        jobs = [common.Job(titulo=f"Pasante Python {i}", empresa="Acme", fuente="Get on Board",
                           link=f"https://example.com/{'x' * 80}/{i}", nivel="Pasantía") for i in range(80)]
        message = notify.format_run(jobs)
        self.assertLessEqual(len(message), notify.MAX_MESSAGE)
        self.assertEqual(message.count("<a "), message.count("</a>"))
        self.assertRegex(message, r"y \d+ más")


if __name__ == "__main__":
    unittest.main()
