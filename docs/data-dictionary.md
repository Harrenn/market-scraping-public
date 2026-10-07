# Dealer database data dictionary

Every database-style workbook contains the same 22 sheets. The public source is authoritative: missing values stay blank, source-labelled zero offers remain zero, and one commercial field is never inferred from another.

## Priority motorcycle fields

| Column | Meaning | Population rule |
| --- | --- | --- |
| `dealer_id` | Stable configured dealer key used across every sheet. | Assigned by the dealer scraper, not scraped from page text. |
| `brand_name` | Public manufacturer/brand evidence for the model. | Taken from a source API object, taxonomy, URL/listing attribute, or scoped page label. It stays blank when none of those sources identifies the brand reliably. |
| `model_name` | Canonical public model/variant name for analytics. | Taken from the source model field or exact detail/listing title. Raw conflicting specification labels remain in `spec_model` and are not allowed to replace the page identity. |
| `srp_amount` | Public cash price, SRP, MSRP, or explicitly equivalent suggested price in PHP. | Numeric only. It is never calculated from downpayment or monthly values. |
| `downpayment_amount` | Public model-specific initial/downpayment amount in PHP. | Numeric only. A blank means no amount was published. A literal zero is retained only when the page explicitly publishes a zero-downpayment offer. |
| `monthly_amount` | Public model-specific periodic amount labelled monthly in PHP. | Numeric only. Loan term is not assumed. Daily prices are not stored here. |

Blank and zero are different. Blank means the public source did not disclose a trustworthy value for that row at crawl time. Zero is a real value only when the source visibly labels zero; otherwise API/database zero sentinels are normalized to blank.

## Priority field coverage

Coverage below is from the verified July 30, 2026 Asia/Manila workbooks.

| Dealer | Models | Brand | Model | SRP | Downpayment | Monthly |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| CDO 2 Cycles | 68 | 68 | 68 | 68 | 0 | 0 |
| Dueksam | 134 | 134 | 134 | 134 | 0 | 0 |
| Guanzon | 86 | 86 | 86 | 86 | 84 | 84 |
| KServico | 48 | 48 | 48 | 48 | 0 | 0 |
| Motor Ace | 43 | 43 | 43 | 43 | 0 | 0 |
| Motorcentral Group | 111 | 109 | 111 | 111 | 111 | 111 |
| Motortrade | 170 | 170 | 170 | 170 | 129 | 138 |
| MotoXpress | 58 | 58 | 58 | 58 | 56 | 49 |
| Premio | 61 | 61 | 61 | 61 | 0 | 0 |
| Premiumbikes | 143 | 143 | 143 | 143 | 0 | 0 |
| Ropali | 35 | 35 | 35 | 0 | 28 | 0 |
| Wheeltek | 111 | 111 | 111 | 81 | 0 | 0 |

## Dealer-specific priority field provenance

| Dealer | Brand and model source | Price source | Why blanks remain |
| --- | --- | --- | --- |
| CDO 2 Cycles | Brand-specific Wix URL family and exact detail-page title. | Currency amount on the exact Wix model page. | The model template publishes a cash price and financing CTA, but no model-specific downpayment or monthly amount. |
| Dueksam | Public `includes/motorcycles.php` AJAX rows: request brand plus `model`. | `sp_price` when populated, otherwise `price`. | The AJAX payload exposes price but no downpayment/monthly keys; the static page is only a shell. |
| Guanzon | WordPress product title plus public product brand/category terms. | Labelled `.sp-product-price` card: SRP, DP, and first displayed monthly term. | CUV e: and EM1 e: publish SRP only; their cards have no DP or monthly rows. |
| KServico | WooCommerce product title and category/brand evidence. | Product JSON-LD offer or scoped WooCommerce price element. | Product pages link to installment/application flows but do not publish model-specific downpayment or monthly amounts. |
| Motor Ace | WooCommerce product title/category. | Positive Store API price; when the API sends zero, the labelled detail-page `SRP:` value is used. | Detail pages publish SRP and a loan CTA, not model-specific downpayment/monthly values. |
| Motorcentral Group | Product title and product-brand/category evidence. | Main product container labels `Price`, `Estimated Down Payment`, and `Estimated Monthly`. | Two SegundaMano rows, PCX160 ABS Roadsync and Click v4, have no direct public brand taxonomy/label, so brand stays blank. Eleven Monarch pages visibly publish `P 0` downpayment; those zeros are real source values. |
| Motortrade | Brand from visible `Make`; canonical model from the primary page `h1`, with metadata/title and visible `Model` specification as fallbacks. Rebate headings, source marketing taglines, and common SEO descriptors are rejected. Conflicting visible `Model` specs remain in `spec_model`. | Scoped `.pcd-pricing` elements for SRP, downpayment, and monthly. | Blank DP/monthly values mean the corresponding element is absent. Nine pages visibly publish `MONTHLY` equal to `SRP`; values are preserved and flagged as source anomalies, not corrected. |
| MotoXpress | Public API `brand.name` and `model_name`. | API `suggested_retail_price`, `all_in_downpayment`, and `monthly_amortization`. | API zero price sentinels are blanked because the rendered page omits those offers. Two Barako rows lack DP; nine models lack monthly. |
| Premio | WooCommerce product title and brand category. | Scoped WooCommerce product price. | Pages publish branch-variable cash prices only; no model-specific DP/monthly values are shown. |
| Premiumbikes | WooCommerce product title and brand/category evidence. | Scoped WooCommerce product price. | Pages publish suggested unit price and inquiry/loan links, but no model-specific DP/monthly amounts. |
| Ropali | WooCommerce product title with brand parsed from the title/content. | WooCommerce current/sale price, explicitly reclassified as downpayment because every product page states that the displayed price is only a downpayment. | SRP and monthly are never inferred. Seven public products have no numeric WooCommerce price, so downpayment is blank too. |
| Wheeltek | Active regular-bike/big-bike listing title and `data-brand` attribute. | Active listing-card SRP, with exact detail-page `Prices starts at` as fallback. | Thirty active listing/detail pages publish no numeric SRP. Apply links do not expose DP/monthly amounts. Stale route aliases, cars, and watercraft are excluded. |

## Provenance

- `dealer`: dealer identity and primary domain.
- `crawl_run`: run timestamps, source count, and notes.
- `source_page`: source URLs, page metadata, extracted text/HTML, and taxonomy references.
- `data_quality_issue`: row-level source conflicts, parse warnings, and explicitly retained anomalies.

## Motorcycles and commercial terms

- `motorcycle`: one row per supported dealer/model record with brand, model, vehicle type, selected commercial fields, and source identifiers.
- `motorcycle_spec`: long-form specification label/value pairs. Raw labels are retained even if they conflict with the canonical page title.
- `pricing`: one or more source/context price observations. Use `motorcycle_id` plus `source_context`; do not join only by dealer.
- `financing`: model-specific or generic financing facts. `motorcycle_id` is nullable for generic application flows.

## Branches and relationships

- `branch`: public branch name, address, contacts, coordinates, and source.
- `branch_product_coverage`: explicitly supported broad branch/category coverage.
- `branch_model_availability`: exact branch/model evidence only. Never create a dealer-wide Cartesian join.

## Promotions and forms

- `promo`: public promotion records and validity text.
- `motorcycle_promo_link`: evidence-backed promo/model links.
- `form_definition`: public form identity, action, method, and purpose.
- `form_field`: visible public field schema. Submitted customer data is never collected.

## Other public surfaces

- `aftersales_service`, `job_listing`, `parts_catalog`, `preowned_motorcycle`, and `external_cta` capture their named public website content.
- `website_filter_reference` records public categories, tags, brands, or UI filters. It is metadata, not a fact-table join.
- `relationship_probe` records attempts to prove a relationship and their result.

Primary keys are stable within a workbook run. Foreign-key-like IDs should resolve to their parent sheet, but the workbook remains an exchange format rather than an enforcing database engine.
