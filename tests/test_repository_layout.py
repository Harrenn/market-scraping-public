from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class RepositoryLayoutTests(unittest.TestCase):
    def test_dealer_scrapers_use_working_output_folder(self) -> None:
        scripts = list((ROOT / "dealer_scrapers").glob("*/scrape_*_database.py"))
        self.assertEqual(len(scripts), 12)
        for script in scripts:
            source = script.read_text(encoding="utf-8")
            self.assertIn("working/outputs/dealer-databases", source)

    def test_generated_workbooks_are_not_in_source_tree(self) -> None:
        source_workbooks = [
            path
            for path in ROOT.rglob("*.xlsx")
            if "working" not in path.parts and "examples" not in path.parts
        ]
        self.assertEqual(source_workbooks, [])


if __name__ == "__main__":
    unittest.main()
