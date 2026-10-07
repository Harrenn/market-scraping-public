# Dealer status

The current database workflow includes:

- CDO 2 Cycles
- Dueksam
- Guanzon Motorcycle Group
- KServico
- Motor Ace
- Motorcentral Group
- Motortrade
- MotoXpress
- Premio
- Premiumbikes
- Ropali
- Wheeltek

Coverage differs by the public surface each dealer exposes. WordPress/WooCommerce dealers commonly provide product taxonomies and detail tables; MotoXpress relies on public Next.js/API data; Ropali combines WooCommerce products with embedded map data; Motortrade provides the broadest reference implementation.

## Verified branch sources

The July 14, 2026 branch re-audit verified these previously empty branch sheets:

| Dealer | Exact rows | Public evidence | Limitation |
| --- | ---: | --- | --- |
| CDO 2 Cycles | 15 | Rendered Wix Google Map initialization payload | Static HTML exposes only one LocalBusiness record; the scraper must inspect the map iframe message to obtain all exact markers. |
| Motor Ace | 181 | Elfsight Store Locator public boot response discovered through rendered network inspection | Address subfields may require later normalization; full addresses and coordinates are retained. |
| Premio | 199 | Static Elementor branch cards on `/store/` | Cards expose names, addresses, and phones but not coordinates. |
| Premiumbikes | 219 | Public `branch-locator/v1/branches` REST response | Exact branch contacts and coordinates are available. |

The combined `branch_united.xlsx` is rebuilt from all 12 dealer workbooks after live branch verification. It includes exact branch records only; rendered locator payloads are used when static HTML is incomplete.

Known limitations include changing page layouts, intermittent origin failures, missing sitemaps, incomplete financing detail, and dealer-wide catalogs without exact branch inventory. Generated workbook `relationship_probe` and `data_quality_issue` sheets are the authoritative run-level record of these gaps.

## Priority motorcycle field audit

The July 30, 2026 audit rechecked `brand_name`, `model_name`, `srp_amount`, `downpayment_amount`, and `monthly_amount` against public pages/APIs for all 12 dealers.

- Motor Ace was regenerated after 15 WooCommerce zero SRPs were proven false; all 43 models now use positive Store API prices or the detail page's labelled SRP.
- Motorcentral was regenerated with price extraction scoped to the main product block. Irregular values such as `P211, 388` no longer spill into adjacent fields, and descriptive 12-month amounts cannot overwrite cash price.
- Motortrade was regenerated with primary-heading model identity, metadata/specification fallbacks, and source-specific cleanup for rebate headings, SEO title fragments, and brand taglines. The July 31 follow-up corrected seven Kymco rebate-name rows and ten additional SEO-decorated model names. Conflicting visible specification-model values and nine source-published `MONTHLY = SRP` cases are retained as explicit `data_quality_issue` rows.
- MotoXpress API zero price sentinels are now blank. The public API and rendered pages do not present them as zero-value offers.
- Wheeltek is now based on 111 active regular-bike/big-bike listing cards. It has complete listing-backed brands, excludes cars/watercraft and stale duplicate aliases, and leaves the 30 source-blank SRPs empty.
- Remaining all-null downpayment/monthly columns for other dealers are expected public non-disclosure, not failed joins or inferred missing data. See `data-dictionary.md` for exact per-dealer rules.

Run `python scripts/audit_motorcycle_priority_fields.py` from the repository root to reproduce the workbook-level coverage and zero-sentinel checks.
