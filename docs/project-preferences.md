# Project Preferences

## Purpose And Precedence

This file records the project conventions. Read it before dealer scraping, workbook generation, audits, combined exports, schema changes, or documentation changes.

Apply instructions in this order:

1. The user's explicit request in the current chat.
2. This preference file.
3. The project architecture, data dictionary, dealer status, and run logs.
4. Older generated reports or historical outputs.

Update this file when project conventions change. Do not convert a one-off request into a permanent rule unless it is clearly intended to recur.

## Working Style

- Act as the scraper pilot, not a command runner.
- Read project documentation, the dealer scraper and tests, the latest dealer run log, and the latest website-recheck evidence before each run.
- Check what changed on the live website before running the scraper.
- Create a baseline run log when none exists, then log every passed, blocked, failed, or no-op run.
- Stage output, compare it with the last known-good workbook, investigate anomalies, fix source-specific bugs, add regression tests, and only then promote it.
- A zero exit code does not prove correct data.
- Preserve the last known-good workbook when a website is blocked, offline, challenged, or only partially available.

## Dealer-Specific Treatment

- Treat every dealer website as a separate implementation.
- Never assume a parser, selector, endpoint, taxonomy, rendering model, locator, or price rule from one dealer applies to another.
- Reuse common workbook schemas and low-level utilities only when useful; keep source discovery and extraction logic dealer-specific.
- Use static requests and BeautifulSoup when public HTML or APIs expose the data.
- Use Playwright or browser/network inspection for rendered shells, dynamic listings, locators, calculators, forms, or protection layers.
- Do not submit forms, authenticate, bypass controls, or scrape private portals.

## Extraction Scope

- Extract as much useful public information as the website exposes. Preserve raw and normalized fields when helpful.
- Do not prune columns based on assumed business preference; the user will prune later.
- Include motorcycles, complete visible specifications, pricing, financing, branches, promos, aftersales/services, forms, jobs, parts, preowned inventory, source pages, public endpoints, and evidence-backed relationships when available.
- Exclude the licensed-plate registry/table from all dealer outputs.
- Record public links to private portals as CTAs only.

## Analytics And Schema Preferences

- Prefer clear, SQL-friendly business tables over source-native taxonomy structures.
- Put reliable business attributes directly on tables such as `motorcycle` when that is the least confusing analytics interface.
- `website_filter_reference` may preserve source/UI metadata, but it is not motorcycle membership evidence and should not be required for normal business queries.
- Keep generic financing rows nullable; never spread them to every model.
- Keep unsupported values blank and explain them in `data_quality_issue` or documentation.
- Never invent relationships. A shared `dealer_id` does not prove branch inventory, promo applicability, financing applicability, or filter membership.
- Populate `branch_model_availability` only when a public source exposes an exact branch and exact model relationship.
- General methodology and schema documentation belongs in `docs/`; dealer-specific run evidence belongs in run logs, `dealer-status.md`, workbook provenance, and `data_quality_issue`.

## Identity And Data-Quality Rules

- Canonical model identity should prefer an exact source API field or primary page heading, then a scoped visible `Model` specification, then clean metadata.
- Reject promo/rebate headings, SEO taglines, category pages, pagination URLs, navigation/footer text, and related-card titles as current model identities.
- Verify suspicious blanks, zeros, duplicates, classifications, prices, and all-one-value distributions against the live source.
- A blank means the public source did not expose a trustworthy value.
- A numeric zero is valid only when explicitly published; source/API zero sentinels become blank.
- Preserve source-labelled anomalies rather than silently correcting them, and flag them clearly.
- Compare refreshes by stable keys, not row position.

## Pricing Preferences

- Store prices as numeric values without peso signs or currency text.
- Keep `srp_amount`, `downpayment_amount`, and `monthly_amount` semantically separate.
- Never infer SRP from downpayment or monthly values.
- Never classify a product-page amount as SRP when the source says it is only a downpayment.
- Do not parse arbitrary nearby numbers or model numbers as prices.
- Preserve labelled source values even when unusual, but document and test verified anomalies.

## Branch Preferences

- Extract exact public branch records whenever available, including name, full address, coordinates, and other exposed fields.
- Branch parsing is dealer-specific. Inspect static cards, JSON-LD, public APIs, rendered maps, iframe messages, and widget/network payloads as appropriate.
- Aggregate counts, city lists, brands, products, and categories are coverage only, not exact branches or model availability.
- For Dueksam, include only locations whose store identity indicates motorcycle dealing through `MC` or `3S`.

## Dealer Workbooks

- Keep one database workbook per dealer with one sheet per table.
- Use the shared 22-sheet business schema unless an explicit change is requested.
- Keep one standalone scraper per dealer under `dealer_scrapers/<slug>/`.
- Default generated workbook path:
  `working/outputs/dealer-databases/<slug>_database.xlsx`.
- Test keys, foreign-key coverage, nulls, duplicates, field semantics, row-count drift, and representative live samples before accepting output.

## Combined Motorcycle Workbook

The simplified `motorcycle_details.xlsx` contract is:

```text
dealer_id
brand_name
model_name
srp_amount
downpayment_amount
monthly_amount
```

Keep these as the exact column names and order unless the user explicitly changes them.

Preserve these worksheet names and order:

1. `motortrade`
2. `motoxpress`
3. `motorace`
4. `cdo2cycles`
5. `premio`
6. `motorcentralgroup`
7. `kservico`
8. `premiumbikes`
9. `guanzon`
10. `dueksam`
11. `wheeltek`
12. `ropali`

Worksheet names do not always equal `dealer_id`. In particular:

- Sheet `guanzon` uses `dealer_id = guanzon_motorcycle_group`.
- The other current worksheet names use the same value as `dealer_id`.

When repopulating this workbook:

- Refresh it from the current audited dealer workbooks.
- Preserve sheet names, sheet order, column names, column order, tables, and formatting.
- Add current models and remove stale or invalid records based on source evidence.
- Reconcile every source row count with the combined output.

## Combined Branch Workbook

`branch_united.xlsx` contains exactly:

```text
dealer_id
branch_name
full_address
latitude
longitude
```

Compile it from every dealer workbook's `branch` sheet. Do not add placeholder rows for dealers without exact public branch records.

## Documentation And Delivery

- Explain how each important field is sourced and why expected blanks remain.
- Document what changed, what failed, how it was fixed, tests performed, and unresolved risks.
- Update reusable skills and project documentation when a discovered bug changes the repeatable workflow.
- Preserve existing workbook structure and formatting during targeted edits.
- Verify final copies by content checks and checksums when delivering the same workbook to another location.

