#!/usr/bin/env python3
from __future__ import annotations

import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from full_scrape.motortrade.scrape_motortrade_analytics import MotortradeAnalyticsScraper


def main() -> int:
    output_path = Path(__file__).resolve().parents[2] / "working/outputs/dealer-databases/motortrade_database.xlsx"
    scraper = MotortradeAnalyticsScraper(output_path, output_style="database")
    scraper.run()
    print(f"[OK] Wrote {output_path.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
