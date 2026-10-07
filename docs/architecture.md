# Architecture

## Workflows

### Branch counts

`scripts/scrape_dealer_branches.py` reads dealer names and websites from an Excel workbook, dispatches to dealer-specific handlers, and writes branch counts to column C. Most handlers use HTTP; a small number use Playwright when rendering is required.

### Website reconnaissance

`scripts/generate_dealer_reports.py` inventories public pages, endpoints, forms, maps, and visible data surfaces. The reports help design or update database scrapers and are stored under `working/outputs/reports`.

`scripts/audit_motorcycle_priority_fields.py` checks current dealer workbooks for identity, price, and data-quality problems before combined exports are accepted.

### Dealer databases

Each module under `dealer_scrapers/<dealer>/` owns that dealer's extraction logic and writes the shared 22-sheet workbook contract. The operational entry points that are not tied to one dealer live under `scripts/`.

`full_scrape/motortrade/scrape_motortrade_analytics.py` is the canonical schema and `TableStore` implementation used by the dealer scrapers, as well as Motortrade's detailed extraction engine. `dealer_scrapers/motortrade/scrape_motortrade_database.py` is the command-line entry point that invokes it. Other dealer scrapers reuse the shared schema and low-level helpers from that module, but keep their website discovery and extraction rules dealer-specific.

Keep project inputs and generated snapshots under `working/`. Inputs and generated outputs stay local and are excluded from this public source distribution. `.venv/`, Python caches, editor state, and operating-system metadata stay local and are ignored by Git.

## Evidence model

Every row should retain its public source URL or crawl context. Dealer-wide products and branches are separate entities unless a source explicitly proves an exact relationship. Unknown values remain blank and are explained through relationship probes or data-quality issues rather than guessed.

## Branch extraction

Branch discovery is deliberately dealer-specific. Static cards and LocalBusiness structured data are parsed directly; public locator APIs are preferred when available; rendered widget network traffic is inspected when the initial HTML is only a shell. Aggregate branch counts and city lists are coverage evidence, not exact branch records.

`working/outputs/dealer-databases/branch_united.xlsx` is a derived export of the dealer workbooks, not a separate scrape source. It keeps only `dealer_id`, `branch_name`, `full_address`, `latitude`, and `longitude`, and never creates placeholder dealers or branch/model relationships.
