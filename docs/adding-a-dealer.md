# Adding a dealer

1. Read `docs/project-preferences.md`, the latest dealer run log, and the latest website-recheck report before probing or running the scraper.
2. Review `robots.txt`, sitemaps, public APIs, rendered pages, and relevant terms before implementation.
3. Record which sources support branches, products, pricing, financing, promotions, forms, jobs, parts, and relationships. A branch page or heading is not evidence that the static HTML contains branch rows: inspect its rendered DOM and network calls before concluding the source is empty.
4. Add `dealer_scrapers/<slug>/scrape_<slug>_database.py` using the shared schemas and `TableStore`.
5. Preserve source URLs and crawl timestamps on generated records.
6. Leave unsupported fields blank. Add a `data_quality_issue` rather than inventing a value.
7. Never infer exact branch/model availability from separate dealer-wide branch and product lists.
8. Default output to `working/outputs/dealer-databases/<slug>_database.xlsx`.
9. Add fixture-based parsing and relationship tests under `tests/`.
10. Run the offline suite, then an explicit live smoke run with polite request pacing.
11. Update `docs/dealer-status.md` with the supported public surfaces and known gaps.

## Branch discovery checklist

Use dealer-specific evidence in this order rather than assuming one locator pattern applies everywhere:

1. Check structured data, static branch cards, contact pages, and sitemaps.
2. Inspect public REST or JSON endpoints referenced by the page.
3. Render locator pages when static HTML contains only a widget shell, then inspect their public network responses.
4. Preserve exact names, addresses, coordinates, contacts, and hours only when the source exposes them.
5. If a page exposes only aggregate counts or city names, document that coverage in `data_quality_issue`; do not manufacture branch entities.

When refreshing the cross-dealer branch file, include only rows from each workbook's `branch` sheet and retain the five-column contract documented in `working/README.md`.

Browser inspection may be used to understand public rendering, but do not submit forms, bypass authentication, capture customer information, or preserve session credentials.
