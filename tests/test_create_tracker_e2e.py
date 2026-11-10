"""E2E: create an empty tracker with scripts/create_tracker.py, then run the real
Excel step on it. Checks the structure the other scripts rely on, that an
existing tracker is never overwritten, and that dedup works against it."""

import hashlib
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

import common  # noqa: E402
import create_tracker  # noqa: E402
import update_excel  # noqa: E402

HEADER = ["Fecha registrada", "Título", "Empresa", "Fuente", "Link", "Modalidad", "Ubicación", "Nivel",
          "Stack", "Compensación (aprox.)", "Estado publicación", "Estado postulación", "Notas"]


class CreateTrackerE2E(unittest.TestCase):
    def setUp(self):
        self.scratch = Path(tempfile.mkdtemp(prefix="radar_create_"))
        self.addCleanup(shutil.rmtree, self.scratch, ignore_errors=True)
        self.tracker = self.scratch / "sub" / "tracker.xlsx"
        for patch in (
            mock.patch.object(create_tracker, "TRACKER_PATH", self.tracker),
            mock.patch.object(common, "TRACKER_PATH", self.tracker),
            mock.patch.object(common, "SEEN_PATH", self.scratch / "seen.json"),
            mock.patch.object(common, "LAST_ADDED_PATH", self.scratch / "last.json"),
            mock.patch.object(update_excel, "TRACKER_PATH", self.tracker),
        ):
            patch.start()
            self.addCleanup(patch.stop)

    def create(self):
        with mock.patch.object(sys, "argv", ["create_tracker"]):
            return create_tracker.main()

    def test_structure_and_first_rows(self):
        self.assertEqual(self.create(), 0)
        wb = openpyxl.load_workbook(self.tracker)
        self.assertEqual(wb.sheetnames, ["Vacantes", "Resumen"])
        ws = wb["Vacantes"]
        self.assertEqual([c.value for c in ws[1]], HEADER)
        self.assertEqual(ws.freeze_panes, "A2")
        lists = {str(dv.sqref): dv.formula1 for dv in ws.data_validations.dataValidation}
        self.assertEqual(lists["K2:K200"], '"Abierta,Cerrada,Programa recurrente,Desconocido"')
        self.assertEqual(lists["L2:L200"], '"No postulado,Postulado,En proceso,Entrevista,Rechazado,Oferta"')
        resumen = wb["Resumen"]
        self.assertEqual(resumen["A4"].value, "No postulado")
        self.assertEqual(resumen["B4"].value, "=COUNTIF(Vacantes!$L$2:$L$200,A4)")
        self.assertEqual(resumen["B10"].value, "=COUNTA(Vacantes!$B$2:$B$200)")
        self.assertEqual(common.tracker_keys(), set())

        jobs = [common.Job(titulo="Backend Jr", empresa="Acme", fuente="Get on Board", link="https://x.io/1"),
                common.Job(titulo="=Pasante", empresa="Beta", fuente="Manual", link="https://x.io/2")]
        self.assertEqual(update_excel.append_jobs(jobs), 0)
        ws = openpyxl.load_workbook(self.tracker)["Vacantes"]
        self.assertEqual([ws.cell(row=r, column=2).value for r in (2, 3, 4)], ["Backend Jr", "=Pasante", None])
        self.assertEqual(ws.cell(row=3, column=2).data_type, "s")  # still never a formula
        self.assertEqual(ws.cell(row=2, column=12).value, "No postulado")
        self.assertTrue(ws.cell(row=3, column=13).alignment.wrap_text)  # row style carried over
        # What is in the tracker now counts as seen.
        self.assertTrue(all(j.is_seen(common.tracker_keys()) for j in jobs))
        self.assertEqual(openpyxl.load_workbook(self.tracker)["Resumen"]["B4"].value,
                         "=COUNTIF(Vacantes!$L$2:$L$200,A4)")

    def test_never_overwrites_an_existing_tracker(self):
        self.assertEqual(self.create(), 0)
        update_excel.append_jobs([common.Job(titulo="Mi vacante", empresa="A", fuente="x", link="https://x.io/9")])
        digest = hashlib.sha256(self.tracker.read_bytes()).hexdigest()
        self.assertEqual(self.create(), 1)
        self.assertEqual(hashlib.sha256(self.tracker.read_bytes()).hexdigest(), digest)


if __name__ == "__main__":
    unittest.main()
