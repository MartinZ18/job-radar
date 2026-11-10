"""E2E: fetch -> Excel with the generic JSON sources (Himalayas, Jobicy,
Remotive, Working Nomads) and the multi-feed RSS source (We Work Remotely),
with faked HTTP shaped like each real API. Checks pagination, field mapping,
dates, levels, the per-source location filter, and source health."""

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

import common  # noqa: E402
import fetch_jobs  # noqa: E402
import update_excel  # noqa: E402
from tests.test_internships_e2e import build_tracker  # noqa: E402

HIM = "https://himalayas.app/jobs/api/search"
JCY = "https://jobicy.com/api/v2/remote-jobs"
RMT = "https://remotive.com/api/remote-jobs"
WN = "https://www.workingnomads.com/api/exposed_jobs/"
WWR_BACK = "https://weworkremotely.com/categories/remote-back-end-programming-jobs.rss"
WWR_FULL = "https://weworkremotely.com/categories/remote-full-stack-programming-jobs.rss"


def him(slug, title, restrictions, seniority=("Entry-level",), pub=1793800000):
    return {"title": title, "companyName": "Acme", "guid": f"https://himalayas.app/companies/acme/jobs/{slug}",
            "locationRestrictions": list(restrictions), "seniority": list(seniority), "pubDate": pub,
            "description": "<p>Python, APIs and Docker</p>", "categories": ["Backend"]}


HIM_PAGES = {
    ("python", 1): [him("a", "Python Backend Developer", []),
                    him("b", "Python Developer", ["Brazil"]),
                    him("c", "Junior Python Engineer", ["Argentina", "Chile"]),
                    him("d", "Backend Engineer", [], seniority=["Senior"])],
    ("python", 2): [him("a", "Python Backend Developer", []),  # repeated across pages
                    him("e", "API Developer", ["United States", "LATAM"])],
}


def rss(feed, items):
    body = "".join(
        f"<item><title>{t}</title><link>https://weworkremotely.com/remote-jobs/{feed}-{i}</link>"
        f"{'<region>' + r + '</region>' if r else ''}{'<country>' + c + '</country>' if c else ''}"
        f"<pubDate>Thu, 05 Nov 2026 17:12:27 +0000</pubDate><description>Python APIs</description></item>"
        for i, (t, r, c) in enumerate(items))
    return f'<?xml version="1.0"?><rss version="2.0"><channel><title>x</title>{body}</channel></rss>'


FEEDS = {
    WWR_BACK: rss("back", [("Vertican: Python Backend Developer", "Anywhere in the World",
                    "Bolivia, Brazil, Chile, Colombia, Peru, and Uruguay"),
                   ("Omega: Backend Engineer (Python)", "Anywhere in the World", ""),
                   ("Backend Developer without company", "", "")]),
    WWR_FULL: rss("full", [("Sigma: Full Stack Developer", "Latin America Only", "")]),
}

SOURCES = [
    {"name": "Himalayas", "type": "json", "url": HIM, "queries": ["python", "n8n"], "query_param": "q",
     "page_param": "page", "pages": 3, "items": "jobs", "modalidad": "Remoto", "location_filter": True,
     "fields": {"titulo": "title", "empresa": "companyName", "link": "guid", "resumen": "description",
                "fecha": "pubDate", "ubicacion": "locationRestrictions", "nivel": "seniority"}},
    {"name": "Jobicy", "type": "json", "url": JCY, "params": {"count": 50, "tag": "python"}, "items": "jobs",
     "modalidad": "Remoto", "location_filter": True,
     "fields": {"titulo": "jobTitle", "empresa": "companyName", "link": "url", "resumen": "jobExcerpt",
                "fecha": "pubDate", "ubicacion": "jobGeo", "nivel": "jobLevel"}},
    {"name": "Remotive", "type": "json", "url": RMT, "params": {"category": "software-dev", "limit": 100},
     "items": "jobs", "modalidad": "Remoto", "location_filter": True,
     "fields": {"titulo": "title", "empresa": "company_name", "link": "url", "resumen": "description",
                "fecha": "publication_date", "ubicacion": "candidate_required_location"}},
    {"name": "Working Nomads", "type": "json", "url": WN, "modalidad": "Remoto",
     "fields": {"titulo": "title", "empresa": "company_name", "link": "url", "resumen": "description",
                "fecha": "pub_date", "ubicacion": "location"}},
    {"name": "We Work Remotely", "type": "rss", "urls": [WWR_BACK, WWR_FULL], "company_in_title": True,
     "modalidad": "Remoto", "location_filter": True},
]


class MoreSourcesE2E(unittest.TestCase):
    def setUp(self):
        self.scratch = Path(tempfile.mkdtemp(prefix="radar_sources_"))
        self.addCleanup(shutil.rmtree, self.scratch, ignore_errors=True)
        self.tracker = self.scratch / "tracker.xlsx"
        build_tracker(self.tracker)
        cfg = common.load_config()
        cfg["http"]["delay_seconds"] = 0
        cfg["sources"] = json.loads(json.dumps(SOURCES))
        self.calls = []
        pending = self.scratch / "pending.json"
        for patch in (
            mock.patch.object(common, "SEEN_PATH", self.scratch / "seen.json"),
            mock.patch.object(common, "PENDING_PATH", pending),
            mock.patch.object(common, "LAST_ADDED_PATH", self.scratch / "last.json"),
            mock.patch.object(common, "HEALTH_PATH", self.scratch / "health.json"),
            mock.patch.object(common, "TRACKER_PATH", self.tracker),
            mock.patch.object(update_excel, "PENDING_PATH", pending),
            mock.patch.object(update_excel, "TRACKER_PATH", self.tracker),
            mock.patch.object(fetch_jobs, "load_config", return_value=cfg),
            mock.patch.object(fetch_jobs.Http, "get", lambda _, url: FEEDS[url]),
            mock.patch.object(fetch_jobs.Http, "get_json", self.fake_get_json),
        ):
            patch.start()
            self.addCleanup(patch.stop)

    def fake_get_json(self, url, params=None):
        params = dict(params or {})
        self.calls.append((url, params))
        if url == HIM:
            return {"jobs": HIM_PAGES.get((params["q"], params["page"]), []), "totalCount": 5}
        if url == JCY:
            return {"jobs": [
                {"jobTitle": "Python Developer", "companyName": "Beta", "url": "https://jobicy.com/jobs/1",
                 "jobGeo": "LATAM", "jobLevel": "Any", "jobExcerpt": "APIs", "pubDate": "2026-11-05T10:00:00+00:00"},
                {"jobTitle": "Backend Developer", "companyName": "Beta", "url": "https://jobicy.com/jobs/2",
                 "jobGeo": "Thailand", "jobLevel": "Any", "jobExcerpt": "APIs", "pubDate": "2026-11-05T10:00:00+00:00"},
            ]}
        if url == RMT:
            return {"0-legal-notice": "Link back to Remotive.", "jobs": [
                {"title": "Backend Developer (Python)", "company_name": "Gamma", "url": "https://remotive.com/j/1",
                 "candidate_required_location": "Worldwide", "publication_date": "2026-11-07T01:10:55",
                 "description": "<p>Python APIs</p>"},
                {"title": "Python Developer", "company_name": "Gamma", "url": "https://remotive.com/j/2",
                 "candidate_required_location": "USA Only", "publication_date": "2026-11-07T01:10:55",
                 "description": "<p>Python</p>"},
            ]}
        if url == WN:  # a bare list, no wrapper object
            return [{"title": "Python Developer", "company_name": "Delta", "url": "https://wn.example/1",
                     "location": "Time zone: CET (+/- 3 hours)", "pub_date": "2026-11-09T08:52:38-05:00",
                     "description": "<p>Python</p>"}]
        raise AssertionError(f"unexpected url {url}")

    def run_chain(self):
        for module in (fetch_jobs, update_excel):
            with mock.patch.object(sys, "argv", [module.__name__]):
                self.assertEqual(module.main(), 0)

    def rows(self):
        ws = openpyxl.load_workbook(self.tracker)["Vacantes"]
        return {(r[3], r[1]): r for r in ws.iter_rows(min_row=3, values_only=True) if r[1]}

    def test_new_sources(self):
        self.run_chain()
        rows = self.rows()
        self.assertEqual(set(rows), {
            ("Himalayas", "Python Backend Developer"),  # no restriction
            ("Himalayas", "Junior Python Engineer"),    # Argentina allowed
            ("Himalayas", "API Developer"),             # LATAM allowed
            ("Jobicy", "Python Developer"),
            ("Remotive", "Backend Developer (Python)"),
            ("Working Nomads", "Python Developer"),     # time zone text: no location filter on this source
            ("We Work Remotely", "Backend Engineer (Python)"),
            ("We Work Remotely", "Full Stack Developer"),
            ("We Work Remotely", "Backend Developer without company"),
        })

        # Pagination: a second page is read, an empty page stops that query.
        him_calls = [(p["q"], p["page"]) for u, p in self.calls if u == HIM]
        self.assertEqual(him_calls, [("python", 1), ("python", 2), ("python", 3), ("n8n", 1)])
        self.assertEqual(sum(1 for u, _ in self.calls if u == RMT), 1)

        # Field mapping, dates in every shape, levels.
        a = rows[("Himalayas", "Python Backend Developer")]
        self.assertEqual((a[2], a[4], a[5]), ("Acme", "https://himalayas.app/companies/acme/jobs/a", "Remoto"))
        self.assertIn("Publicada", a[12])
        c = rows[("Himalayas", "Junior Python Engineer")]
        self.assertEqual((c[6], c[7]), ("Argentina, Chile", "Junior"))
        self.assertIn("05/11/2026", rows[("Jobicy", "Python Developer")][12])
        self.assertIn("07/11/2026", rows[("Remotive", "Backend Developer (Python)")][12])
        self.assertIn("09/11/2026", rows[("Working Nomads", "Python Developer")][12])

        # "Company: Title" is split; a title without the separator is left whole.
        omega = rows[("We Work Remotely", "Backend Engineer (Python)")]
        self.assertEqual((omega[2], omega[6]), ("Omega", "Anywhere in the World"))
        self.assertEqual(rows[("We Work Remotely", "Backend Developer without company")][2], "No especificada")

        health = json.loads((self.scratch / "health.json").read_text(encoding="utf-8"))
        self.assertEqual({n: e["fallas"] for n, e in health.items()},
                         {"Himalayas": 0, "Jobicy": 0, "Remotive": 0, "Working Nomads": 0, "We Work Remotely": 0})

    def test_title_noise_is_filtered(self):
        noise = ["Cryptography Specialist for AI lab", "Freelance Mechanical CFD Engineer - AI Trainer",
                 "Remote Industrial Automation Consultant", "Booster AI Mentor", "Software Engineer Interviewer"]
        keep = ["AI Engineer", "Desarrollador IA Junior", "AI Automation Expert (n8n / Make / Zapier / AI Agents)",
                "Ingeniero de IA trainee", "Plataforma de e-commerce con IA", "Desenvolvedor IA & Automações remoto",
                "Junior Python Developer"]
        pages = {("python", 1): [him(f"t{i}", t, []) for i, t in enumerate(noise + keep)]}
        # A body that mentions mentoring must not sink a good posting: title-only exclusions.
        pages[("python", 1)][-1]["description"] = "<p>Python APIs. Vas a tener mentoría y un mentor asignado.</p>"
        with mock.patch.dict(HIM_PAGES, pages, clear=True):
            self.run_chain()
        titles = {t for (src, t) in self.rows() if src == "Himalayas"}
        self.assertEqual(titles, set(keep))

    def test_senior_in_the_body(self):
        """Cases from the 2026-11 audit of 764 real postings."""
        def body(slug, title, text):
            return {**him(slug, title, [], seniority=[]), "description": f"<p>{text}</p>"}
        pages = {("python", 1): [
            body("s1", "Junior Software Engineer", "Maintain clean code under the guidance of senior engineers."),
            body("s2", "Backend Engineer", "Backend Engineer (Junior/Middle/Senior). Python and APIs."),
            body("s3", "Backend Developer", "About the role: as a Senior Backend Engineer you'll design APIs."),
            body("s4", "Software Engineer Position - React and Rest", "We are seeking a highly skilled Senior Software Engineer."),
        ]}
        with mock.patch.dict(HIM_PAGES, pages, clear=True):
            self.run_chain()
        titles = {t for (src, t) in self.rows() if src == "Himalayas"}
        self.assertEqual(titles, {"Junior Software Engineer", "Backend Engineer"})

    def test_second_run_adds_nothing(self):
        self.run_chain()
        before = self.rows()
        self.run_chain()
        self.assertEqual(self.rows(), before)


if __name__ == "__main__":
    unittest.main()
