# Local working files

Create `data/Dealers.xlsx` locally for branch-count updates. Dealer names go in column A, websites in column B and counts in column C. See the synthetic CSV under examples/ for the layout.

Scrapers create datasets under `outputs/dealer-databases/`; report tools use `outputs/reports/`. Keep website-recheck state, logs and snapshots under `outputs/` as well. These directories and all Excel workbooks are ignored by Git.

No original input, scraped dataset or captured report is included. Review and redact output before sharing it; captured URLs may include third-party credentials.

The optional combined branch export uses exactly `dealer_id`, `branch_name`, `full_address`, `latitude`, and `longitude`, populated only from source-backed branch rows.
