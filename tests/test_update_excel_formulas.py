"""E2E: run the real append_jobs path against a scratch copy of the tracker with
hostile vacancy data and check with openpyxl that nothing is stored as a formula."""

import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

import update_excel  # noqa: E402
from common import Job  # noqa: E402

EVIL = '=HYPERLINK("http://evil","x")'


def build_fixture(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Vacantes"
    ws.append(["Fecha", "Titulo", "Empresa", "Fuente", "Link", "Modalidad", "Ubicacion",
               "Nivel", "Stack", "Compensacion", "Estado pub", "Estado post", "Notas"])
    ws.append(["01/01/2026", "Existente", "Acme", "x", "http://a", "", "", "", "", "", "", "", ""])
    wb.save(path)


class FormulaInjectionTest(unittest.TestCase):
    def run_append(self, jobs):
        scratch = Path(tempfile.mkdtemp(prefix="radar_xlsx_", dir=tempfile.gettempdir()))
        self.scratch = scratch
        tracker = scratch / "tracker.xlsx"
        build_fixture(tracker)
        with mock.patch.object(update_excel, "TRACKER_PATH", tracker), \
                mock.patch.object(update_excel, "load_seen", return_value={}), \
                mock.patch.object(update_excel, "save_seen"), \
                mock.patch.object(update_excel, "save_last_added"):
            self.assertEqual(update_excel.append_jobs(jobs), 0)
        return openpyxl.load_workbook(tracker)["Vacantes"]

    def test_malicious_values_are_text_not_formulas(self):
        payloads = [EVIL, "+cmd|' /C calc'!A0", "-2+3", "@SUM(1+1)", "\t=1+1", "\r=1+1"]
        jobs = [Job(titulo=p, empresa=p, fuente="t", link="http://ok/" + str(i),
                    stack=p, compensacion=p, resumen=p) for i, p in enumerate(payloads)]
        ws = self.run_append(jobs)
        for r, p in enumerate(payloads, start=3):
            for col in (2, 3, 9, 10):
                cell = ws.cell(row=r, column=col)
                self.assertEqual(cell.data_type, "s", f"row {r} col {col} is {cell.data_type}")
                self.assertEqual(cell.value, p)  # text kept intact, no visible apostrophe
        self.assertEqual(ws.cell(row=3, column=2).value, EVIL)

    def test_normal_values_unchanged(self):
        ws = self.run_append([Job(titulo="Backend Jr", empresa="Acme", fuente="t",
                                  link="https://x.io/job/1")])
        self.assertEqual(ws.cell(row=3, column=2).value, "Backend Jr")
        self.assertEqual(ws.cell(row=3, column=5).value, "https://x.io/job/1")
        self.assertEqual(ws.cell(row=3, column=2).data_type, "s")


if __name__ == "__main__":
    unittest.main()
