"""E2E: regenerate the Radar notes from a scratch tracker and scratch interview
items. Checks that "Proyectos más pedidos" gains what the tracker's postings ask
for (normalized technologies, frequent combinations, what they say about
projects) while the forum ranking and the other notes behave as before."""

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

import fetch_interviews  # noqa: E402
import update_obsidian  # noqa: E402

HEADER = ["Fecha", "Titulo", "Empresa", "Fuente", "Link", "Modalidad", "Ubicacion",
          "Nivel", "Stack", "Compensacion", "Estado pub", "Estado post", "Notas"]


def row(fecha, titulo, stack, notas="", link=None):
    return [fecha, titulo, "Acme", "x", link or f"https://example.com/{titulo.replace(' ', '-')}",
            "Remoto", "", "Junior", stack, "", "Abierta", "No postulado", notas]


ROWS = [
    row("01/09/2026", "Automatizador IA Junior", "n8n, Make, ChatGPT/Claude, APIs, webhooks",
        "Part-time. Valoran proyectos reales por sobre certificados."),
    row("05/09/2026", "Python Backend Jr", "python, FastAPI, Docker, dev, digital nomad"),
    row("10/09/2026", "Pasante Python", "Python, APIs",
        "Make sure you love testing. Enviar link a GitHub con un proyecto propio."),
    row("15/09/2026", "n8n Developer", "n8n, APIs, design, customer support"),
    row("20/09/2026", "Data Trainee", "No especificado", "Contar con portfolio es un plus."),
    row("21/09/2026", "Full Stack", "Node",
        "Se suma a un proyecto del sector salud. Open to project-based work."),
]

ITEMS = {
    "a": {"tipo": "proyecto", "texto": "Chatbot con RAG", "seccion": "IA y LLM", "veces": 3, "idioma": "es",
          "fuentes": [{"origen": "n8n", "titulo": "Post", "url": "https://example.com/p"}]},
    "b": {"tipo": "pregunta", "texto": "¿Qué es REST?", "seccion": "Backend y APIs", "veces": 1, "idioma": "es",
          "fuentes": []},
}


def build_tracker(path: Path, rows) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Vacantes"
    ws.append(HEADER)
    for r in rows:
        ws.append(r)
    footer = len(rows) + 2
    ws.cell(row=footer, column=1, value="Nota de ayuda: no editar esta fila")
    ws.merge_cells(start_row=footer, start_column=1, end_row=footer, end_column=13)
    wb.save(path)


class ProjectsNoteE2E(unittest.TestCase):
    def setUp(self):
        self.scratch = Path(tempfile.mkdtemp(prefix="radar_projects_"))
        self.addCleanup(shutil.rmtree, self.scratch, ignore_errors=True)
        self.vault = self.scratch / "vault"
        self.vault.mkdir()
        (self.vault / "00 Índice.md").write_text("# Índice\n\n## Notas\n\n- [[Glosario]]\n", encoding="utf-8")
        self.tracker = self.scratch / "tracker.xlsx"
        items = self.scratch / "items.json"
        items.write_text(json.dumps(ITEMS, ensure_ascii=False), encoding="utf-8")

        cfg = fetch_interviews.load_config()
        cfg["obsidian"]["vault_dir"] = str(self.vault)
        for patch in (
            mock.patch.object(update_obsidian, "load_config", return_value=cfg),
            mock.patch.object(update_obsidian, "ITEMS_PATH", items),
            mock.patch.object(update_obsidian, "TRACKER_PATH", self.tracker),
        ):
            patch.start()
            self.addCleanup(patch.stop)

    def run_main(self):
        with mock.patch.object(sys, "argv", ["update_obsidian"]):
            self.assertEqual(update_obsidian.main(), 0)
        return (self.vault / "Radar" / "Proyectos más pedidos (radar).md").read_text(encoding="utf-8")

    def section(self, note, title):
        start = note.index(f"### {title}")
        end = note.find("\n#", start + 1)
        return note[start:end if end != -1 else None]

    def test_tracker_section(self):
        build_tracker(self.tracker, ROWS)
        note = self.run_main()

        self.assertIn("## Lo que piden las vacantes", note)
        self.assertIn("6 vacantes del tracker (01/09/2026 – 21/09/2026)", note)

        tech = self.section(note, "Tecnologías")
        # Case duplicates merge; Remote OK noise and "No especificado" never count.
        self.assertIn("| Python | 2 | 33% |", tech)
        self.assertIn("| APIs | 3 | 50% |", tech)
        self.assertIn("| n8n | 2 | 33% |", tech)  # counted once per posting, title or stack
        for noise in ("dev", "design", "customer support", "digital nomad", "No especificado"):
            self.assertNotIn(f"| {noise} |", tech)
        # Notes are not scanned for technologies: "Make sure" is not Make.
        self.assertIn("| Make | 1 | 17% |", tech)
        # Ranked by count.
        self.assertLess(tech.index("| APIs |"), tech.index("| Python |"))
        self.assertLess(tech.index("| Python |"), tech.index("| Make |"))

        combos = self.section(note, "Combinaciones frecuentes")
        self.assertIn("n8n + APIs — 2 vacantes", combos)
        self.assertNotIn("FastAPI + Docker", combos)  # seen once only

        quotes = self.section(note, "Qué dicen sobre proyectos")
        self.assertIn("Valoran proyectos reales por sobre certificados.", quotes)
        self.assertIn("Enviar link a GitHub con un proyecto propio.", quotes)
        self.assertNotIn("Make sure", quotes)
        # Contract wording is not a portfolio signal.
        self.assertNotIn("sector salud", quotes)
        self.assertNotIn("project-based", quotes)
        self.assertIn("Contar con portfolio es un plus.", quotes)
        self.assertLess(quotes.index("portfolio"), quotes.index("GitHub"))  # newest first
        self.assertIn("[Pasante Python](https://example.com/Pasante-Python)", quotes)

        # The forum ranking is still there, and the footer row was not a posting.
        self.assertIn("## Ranking", note)
        self.assertIn("**Chatbot con RAG**", note)
        self.assertNotIn("Nota de ayuda", note)

        # Other notes and the index link behave as before.
        radar = self.vault / "Radar"
        self.assertIn("¿Qué es REST?", (radar / "Preguntas técnicas (radar).md").read_text(encoding="utf-8"))
        self.assertTrue((radar / "Live coding (radar).md").exists())
        self.assertTrue((radar / "Comportamental (radar).md").exists())
        self.assertIn("Radar/Preguntas técnicas (radar)", (self.vault / "00 Índice.md").read_text(encoding="utf-8"))

    def test_missing_tracker_still_writes_notes(self):
        note = self.run_main()  # no tracker built
        self.assertIn("## Lo que piden las vacantes", note)
        self.assertIn("No se pudo leer el tracker", note)
        self.assertIn("**Chatbot con RAG**", note)


if __name__ == "__main__":
    unittest.main()
