# Motorcycle priority field audit

## Purpose

This audit answers whether blanks and suspicious values in `dealer_id`, `brand_name`, `model_name`, `srp_amount`, `downpayment_amount`, and `monthly_amount` are public-source limitations or scraper defects. It covers all 12 dealer workbooks and the combined `motorcycle_details.xlsx`.

Audit context: July 30, 2026, Asia/Manila. Only public, unauthenticated pages and APIs were inspected. Forms were not submitted.

## Method

1. Profile every dealer `motorcycle` sheet for blanks, numeric zero, duplicate IDs, equal SRP/monthly values, and suspicious model/title mismatches.
2. Trace each pattern through its dealer-specific scraper.
3. Recheck representative and exceptional rows against the live public HTML, JSON, AJAX, or listing payload.
4. Fix only source-backed parser/catalog defects and add regression tests.
5. Regenerate affected workbooks, rerun coverage checks, and rebuild the combined output.

Static requests and BeautifulSoup were used when data was present in HTML/API responses. Dynamic or protected behavior was treated according to each dealer's existing scraper; no private portal was accessed.

## Live source findings

| Dealer | Public evidence checked | Conclusion for priority blanks |
| --- | --- | --- |
| CDO 2 Cycles | Exact Wix model page, including MIO AEROX. | Cash amount is visible; no DP/monthly labels are present. |
| Dueksam | `POST /includes/motorcycles.php` with `type=fetchDT`. | Rows expose `entity_id`, `model`, `price`, optional `sp_price`, and image fields, but no DP/monthly fields. |
| Guanzon | Financed Brusky 125 and SRP-only CUV e: price cards. | Standard cards expose SRP/DP/monthly terms; CUV e: and EM1 e: omit DP/monthly markup. |
| KServico | WooCommerce model page and application CTA. | Product price is public; installment flow does not disclose model DP/monthly amounts on the page. |
| Motor Ace | All 15 rows previously carrying zero SRP. | Every page visibly published a positive labelled SRP; the zeros were Store API sentinels. |
| Motorcentral | Main detail blocks, SegundaMano pages, Monarch zero-DP models, and Bajaj irregular formatting. | Main labels are authoritative. Eleven zero DPs are explicitly displayed. Two second-hand products expose no direct brand taxonomy/label. |
| Motortrade | Detail titles, `.pcd-pricing`, specification blocks, and suspicious equal prices. | Missing fields correspond to absent elements. Nine pages visibly label monthly equal to SRP. Several pages have title/spec-model conflicts. |
| MotoXpress | `/v1/motorcycles` records plus rendered ADV160. | API zero monthly/DP values are sentinels; rendered pages omit those values rather than advertising zero. |
| Premio | WooCommerce detail page. | Cash price is visible; no model-specific DP/monthly amount is shown. |
| Premiumbikes | WooCommerce detail page. | Suggested unit price is visible; no model-specific DP/monthly amount is shown. |
| Ropali | Every row lacking downpayment plus a priced MIO AEROX page. | Pages explicitly say the displayed price is only a downpayment. Seven products show no numeric price. |
| Wheeltek | All active regular-bike and big-bike listing pages and detail fallback. | Listing cards directly expose model, `data-brand`, and optional SRP. Thirty active cards/details have no numeric SRP; no DP/monthly offers are published. |

## Defects found and fixed

### Motor Ace zero SRP

WooCommerce returned numeric zero for 15 products even though detail pages published positive SRPs. The scraper now accepts Store API price only when positive and otherwise parses the exact `SRP:` label. All 43 current rows now have positive SRP.

### Motorcentral price contamination

Three separate failure modes were reproduced:

- `P211, 388` was not parsed as one number, allowing SRP extraction to spill into downpayment.
- A title dash mismatch could make extraction start in a global search/related-product block.
- A later descriptive `PRICE:` section could overwrite the main cash price with a 12-month amount.

The parser now normalizes comma spacing, scopes to the matching main product container, stops at adjacent price labels, and keeps the first authoritative value per field.

### Motortrade model identity

Some pages have no primary heading in the expected container, so an unscoped related-card title was previously used. Page metadata/title now precedes unscoped card fallbacks. `model_name` uses that page identity; conflicting visible `spec_model` is retained separately and logged.

Motortrade also publishes occasional SEO-decorated titles and a Bajaj manufacturer label containing a marketing tagline. The source-specific parser now removes known pipe-delimited SEO text from the canonical model and trims dash-delimited brand taglines while preserving the unmodified source evidence in `source_page`.

### Motortrade promotional heading contamination

A July 31, 2026 follow-up found seven Kymco rows whose `model_name` was a visible `REBATE` heading rather than the motorcycle name. The affected URLs were legitimate model detail pages, and each page exposed the correct identity in its primary `h1` and visible `Model` specification. The same selector-order problem also allowed broader SEO descriptors such as `Price Philippines` and `Scooter PH` into ten other Motortrade names.

The extractor now evaluates primary `h1` headings before nested `h2` content, rejects rebate headings as model identities, and falls back to the visible specification model when necessary. The cross-dealer audit now fails on promotional and common SEO-title patterns in `model_name`. After regeneration, all 170 Motortrade names pass this check and all 16 Motortrade Kymco rows contain actual model names.

### MotoXpress zero sentinels

Zero API price fields were previously emitted as real amounts. Price fields now require a positive value; zero becomes blank while the raw API record remains represented through provenance/spec text.

### Wheeltek catalog scope and brands

The previous scraper mixed active `/motorcycles/` rows, stale `/vehicles/` aliases, cars, and watercraft. It also missed brands that were available on listing cards. The scraper now uses active regular-bike/big-bike cards as scope and reads their `data-brand`, title, and SRP. Current output is 111 active motorcycles with complete brand/model coverage.

## Remaining source limitations

- Motorcentral's two blank brands are deliberate. The public SegundaMano records identify PCX160 ABS Roadsync and Click v4 but do not attach a manufacturer taxonomy or scoped manufacturer label.
- Motorcentral's 11 zero downpayments are deliberate because the pages visibly publish `P 0`.
- Motortrade's nine equal SRP/monthly values are source-labelled anomalies. Correcting them without another public source would invent data.
- Ropali has no SRP or monthly facts in scope. Its public product price is explicitly a downpayment.
- Wheeltek's 30 blank SRPs are blank on both the active listing card and exact detail fallback.
- All-null DP/monthly columns at CDO 2 Cycles, Dueksam, KServico, Motor Ace, Premio, Premiumbikes, and Wheeltek reflect non-disclosure. A loan/application link is not a numeric financing fact.

## Reproducibility

Run parser tests:

```bash
.venv/bin/python -m unittest \
  tests.test_scrape_motortrade_analytics \
  tests.test_scrape_motoxpress_database \
  tests.test_standalone_dealer_scrapers
```

Run the workbook audit:

```bash
.venv/bin/python scripts/audit_motorcycle_priority_fields.py
```

Regenerate the corrected dealer workbooks:

```bash
.venv/bin/python dealer_scrapers/motorace/scrape_motorace_database.py
.venv/bin/python dealer_scrapers/motorcentralgroup/scrape_motorcentralgroup_database.py
.venv/bin/python dealer_scrapers/motortrade/scrape_motortrade_database.py
.venv/bin/python dealer_scrapers/motoxpress/scrape_motoxpress_database.py
.venv/bin/python dealer_scrapers/wheeltek/scrape_wheeltek_database.py
```

The exact live pages and run-level caveats remain in each workbook's `source_page` and `data_quality_issue` sheets.
