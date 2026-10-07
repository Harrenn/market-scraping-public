# Philippine Motorcycle Market Scraping

Source-only tooling for collecting public information from Philippine motorcycle dealer websites. The project focuses on dealer reach, catalogs and specifications, public pricing/financing, promotions, services, forms, jobs, parts, and source provenance.

It supports three workflows:

1. Update dealer branch counts in a working workbook.
2. Generate website-reconnaissance Markdown reports.
3. Build SQL-friendly, 22-sheet dealer database workbooks using a shared schema.

## Repository layout

```text
scripts/                    Branch updater, site reports, and workbook audit tools
dealer_scrapers/<dealer>/   One source-specific database scraper per dealer
full_scrape/motortrade/     Shared workbook schema plus Motortrade extraction
docs/                       Preferences, architecture, schema, and operating guides
examples/                   Small synthetic input examples
tests/                      Offline parser and workbook behavior tests
working/data/               Local dealer input workbook (ignored)
working/outputs/            Local generated outputs (ignored)
.venv/                      Local Python environment; excluded from Git
```

The repository root is the only code root. This public distribution starts with fresh history and contains no working datasets or captured reports. Inputs in `working/data/` and outputs in `working/outputs/` stay local and are ignored by Git, along with Python caches and virtual environments.

## Setup

Python 3.11 or 3.12 is recommended.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

## Branch-count workflow

Create your own local workbook at `working/data/Dealers.xlsx` with dealer names in column A, website URLs in column B, and branch counts in column C. The synthetic CSV at `examples/dealers.example.csv` illustrates the layout; its example.com entry is a placeholder, not a supported dealer to scrape. No working workbook is bundled.

```bash
python scripts/scrape_dealer_branches.py --dry-run
python scripts/scrape_dealer_branches.py
python scripts/scrape_dealer_branches.py --only "Motortrade"
```

Successful non-dry runs write the count to column C of the workbook.

## Reconnaissance reports

```bash
python scripts/generate_dealer_reports.py --help
```

Reports default to `working/outputs/reports/`. They are generated evidence, not maintained documentation.

## Dealer database workbooks

Each script under `dealer_scrapers/<dealer>/` creates a workbook in `working/outputs/dealer-databases/`. For example:

```bash
python dealer_scrapers/motortrade/scrape_motortrade_database.py
python dealer_scrapers/kservico/scrape_kservico_database.py
python dealer_scrapers/ropali/scrape_ropali_database.py
```

All dealer scripts use the shared schema documented in [docs/data-dictionary.md](docs/data-dictionary.md). A critical rule is to preserve only relationships directly supported by public evidence: a dealer-wide catalog must not be expanded into invented branch/model availability rows.

The current branch sheets can be compiled into one import-ready file with the exact columns `dealer_id`, `branch_name`, `full_address`, `latitude`, and `longitude`. The generated file is `working/outputs/dealer-databases/branch_united.xlsx`; it contains only source-backed branch rows and does not add placeholder rows for dealers with incomplete public coverage.

## Workbook audit

Check the priority motorcycle identity and pricing fields across the current dealer workbooks with:

```bash
python scripts/audit_motorcycle_priority_fields.py
```

## Documentation

- [Project preferences](docs/project-preferences.md)
- [Architecture](docs/architecture.md)
- [Data dictionary](docs/data-dictionary.md)
- [Adding a dealer](docs/adding-a-dealer.md)
- [Dealer status](docs/dealer-status.md)
- [Responsible use](docs/responsible-use.md)

## Tests

```bash
python -m unittest discover -s tests -v
```

The default suite is offline and fixture-driven. Live-site runs are intentionally separate because dealer websites change and may rate-limit or fail temporarily.

## Publication privacy

Keep generated spreadsheets, reports, website-recheck logs/state, browser state and credentials local. Captured URLs can contain third-party API keys, session tokens or nonces; redact them before sharing output. Ignore rules do not sanitize report contents or remove already tracked files. This repository publishes reusable source and synthetic fixtures, not the original project datasets.
