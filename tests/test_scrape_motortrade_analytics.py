from __future__ import annotations

import unittest

from bs4 import BeautifulSoup

from full_scrape.motortrade.scrape_motortrade_analytics import (
    DATABASE_SHEETS,
    SECTION_SHEETS,
    build_database_rows,
    build_section_rows,
    extract_form_fields,
    extract_motorcycle_display_name,
    extract_motorcycle_links,
    extract_pricing_from_soup,
    find_branch_model_evidence,
    infer_scoped_vehicle_type,
    is_motorcycle_detail_url,
    model_names_compatible,
    normalize_brand,
    normalize_url,
    normalize_coverage_values,
    parse_currency_amount,
    split_numeric_unit,
)


class ScrapeMotortradeAnalyticsTests(unittest.TestCase):
    def test_normalize_brand_removes_source_marketing_tagline(self) -> None:
        self.assertEqual(
            normalize_brand("BAJAJ – The World's No.1 Three-wheeler"),
            "BAJAJ",
        )

    def test_display_name_removes_motortrade_seo_title_decorations(self) -> None:
        prefixed = BeautifulSoup(
            '<meta property="og:title" content="Motortrade | Philippine\'s Best Motorcycle Dealer | SUZUKI GSX R150">',
            "lxml",
        )
        suffixed = BeautifulSoup(
            '<meta property="og:title" content="All-New Honda Giorno+ Scooter | Price &amp; Features">',
            "lxml",
        )
        self.assertEqual(extract_motorcycle_display_name(prefixed), "SUZUKI GSX R150")
        self.assertEqual(extract_motorcycle_display_name(suffixed), "All-New Honda Giorno+ Scooter")

    def test_normalize_url(self) -> None:
        self.assertEqual(
            normalize_url("https://Motortrade.com.ph/motorcycles/honda-adv-160//?x=1&&"),
            "https://motortrade.com.ph/motorcycles/honda-adv-160?x=1",
        )

    def test_parse_currency_amount(self) -> None:
        self.assertEqual(parse_currency_amount("SRP: ₱125,900"), 125900.0)
        self.assertEqual(parse_currency_amount("MONTHLY:  ₱5,800"), 5800.0)
        self.assertEqual(parse_currency_amount("No price"), "")

    def test_split_numeric_unit(self) -> None:
        self.assertEqual(split_numeric_unit("157cc"), (157.0, "cc"))
        self.assertEqual(split_numeric_unit("1,890 x 686 x 1,116 mm"), (1890.0, "x 686 x 1,116 mm"))
        self.assertEqual(split_numeric_unit("PGM-FI"), ("", ""))

    def test_extract_pricing_from_soup(self) -> None:
        soup = BeautifulSoup(
            """
            <p class="pcd-pricing">
              <span class="pcd-price"><b>SRP:</b> ₱166,900</span>
              <span class="pcd-downpayment"><b>Downpayment:</b> ₱15,000</span>
              <span class="pcd-price-msrp here">MONTHLY: ₱7,800</span>
            </p>
            <p class="psnote">Indicative pricing only.</p>
            """,
            "lxml",
        )
        pricing = extract_pricing_from_soup(soup)
        self.assertEqual(pricing["srp_amount"], 166900.0)
        self.assertEqual(pricing["downpayment_amount"], 15000.0)
        self.assertEqual(pricing["monthly_amount"], 7800.0)
        self.assertIn("Indicative", pricing["note_text"])

    def test_display_name_ignores_related_card_titles(self) -> None:
        soup = BeautifulSoup(
            """
            <main>
              <section class="vehicle-content"><h1>Honda Airblade 160</h1></section>
              <section class="related">
                <div class="item"><h2 class="Vehicle-Title">HONDA ADV160 (ABS TYPE)</h2></div>
              </section>
            </main>
            """,
            "lxml",
        )
        self.assertEqual(extract_motorcycle_display_name(soup, "Wrong - Motortrade"), "Honda Airblade 160")

    def test_display_name_prefers_page_metadata_over_unscoped_vehicle_cards(self) -> None:
        soup = BeautifulSoup(
            """
            <html>
              <head><meta property="og:title" content="SUZUKI RAIDER PRO - Motortrade"></head>
              <body>
                <div class="Vehicle-Title">SUZUKI BURGMAN 15</div>
                <div class="Vehicle-Title">SUZUKI DR160</div>
              </body>
            </html>
            """,
            "lxml",
        )

        self.assertEqual(extract_motorcycle_display_name(soup, "Wrong - Motortrade"), "SUZUKI RAIDER PRO")

    def test_display_name_prefers_primary_h1_over_rebate_heading(self) -> None:
        soup = BeautifulSoup(
            """
            <html>
              <head><meta property="og:title" content="KYMCO AK550 - Motortrade"></head>
              <body>
                <main>
                  <h1>KYMCO AK550</h1>
                  <section class="vehicle-content"><h2>REBATE : 50,000</h2></section>
                </main>
              </body>
            </html>
            """,
            "lxml",
        )

        self.assertEqual(extract_motorcycle_display_name(soup), "KYMCO AK550")

    def test_display_name_rejects_rebate_h2_when_metadata_has_model(self) -> None:
        soup = BeautifulSoup(
            """
            <html>
              <head><meta property="og:title" content="KYMCO Like 150i ABS - Motortrade"></head>
              <body><section class="vehicle-content"><h2>REBATE : 5,000</h2></section></body>
            </html>
            """,
            "lxml",
        )

        self.assertEqual(extract_motorcycle_display_name(soup), "KYMCO Like 150i ABS")

    def test_model_name_compatibility_detects_conflicting_visible_spec(self) -> None:
        self.assertTrue(model_names_compatible("HONDA ADV160 (ABS TYPE)", "HONDA ADV160"))
        self.assertFalse(model_names_compatible("SUZUKI RAIDER PRO", "BURGMAN STREET EX"))

    def test_scoped_vehicle_type_ignores_nav_links(self) -> None:
        soup = BeautifulSoup(
            """
            <header><a href="/motorcycle-type/big-bike/">Big Bike</a></header>
            <main>
              <article>
                <nav class="breadcrumb"><a href="/motorcycle-type/regular-bike/">Regular Bike</a></nav>
                <h1>Honda Airblade 160</h1>
              </article>
            </main>
            """,
            "lxml",
        )
        self.assertEqual(
            infer_scoped_vehicle_type(soup, "https://motortrade.com.ph/motorcycles/honda-airblade-160-motorcycle-latest/"),
            "regular-bike",
        )

    def test_scoped_vehicle_type_does_not_use_header_only_taxonomy(self) -> None:
        soup = BeautifulSoup(
            """
            <header><a href="/motorcycle-type/big-bike/">Big Bike</a></header>
            <main><article><h1>Honda Airblade 160</h1></article></main>
            """,
            "lxml",
        )
        self.assertEqual(
            infer_scoped_vehicle_type(soup, "https://motortrade.com.ph/motorcycles/honda-airblade-160-motorcycle-latest/"),
            "",
        )

    def test_motorcycle_detail_url_excludes_archive_pagination(self) -> None:
        self.assertTrue(is_motorcycle_detail_url("https://motortrade.com.ph/motorcycles/honda-airblade-160-motorcycle-latest/"))
        self.assertFalse(is_motorcycle_detail_url("https://motortrade.com.ph/motorcycles/"))
        self.assertFalse(is_motorcycle_detail_url("https://motortrade.com.ph/motorcycles/page/2/?mot-type=regular-bike"))

    def test_extract_motorcycle_links_excludes_pagination(self) -> None:
        soup = BeautifulSoup(
            """
            <a href="/motorcycles/honda-airblade-160-motorcycle-latest/">Airblade</a>
            <a href="/motorcycles/page/2/?mot-type=regular-bike">2</a>
            """,
            "lxml",
        )
        self.assertEqual(
            extract_motorcycle_links(soup, "https://motortrade.com.ph/motorcycles/"),
            ["https://motortrade.com.ph/motorcycles/honda-airblade-160-motorcycle-latest"],
        )

    def test_normalize_coverage_values(self) -> None:
        self.assertEqual(normalize_coverage_values("33"), ["33"])
        self.assertEqual(normalize_coverage_values("Honda, Yamaha"), ["honda", "yamaha"])
        self.assertEqual(normalize_coverage_values(["Big Bike", "Regular Bike"]), ["big bike", "regular bike"])
        self.assertEqual(normalize_coverage_values(""), [])

    def test_branch_model_evidence_requires_branch_and_model_keys(self) -> None:
        self.assertFalse(find_branch_model_evidence([{"id": "5834", "title": "Branch", "categories": "33"}]))
        self.assertTrue(find_branch_model_evidence([{"id": "5834", "model_id": "abc123"}]))

    def test_build_section_rows_for_human_workbook(self) -> None:
        tables = {
            "motorcycle": [
                {
                    "motorcycle_id": "motorcycle-airblade",
                    "source_page_id": "sp-1",
                    "display_name": "HONDA AIRBLADE 160",
                    "brand_name": "HONDA",
                    "model_name": "Airblade 160",
                    "vehicle_type": "regular-bike",
                    "spec_displacement_cc": "157cc",
                    "detail_url": "https://motortrade.com.ph/motorcycles/airblade",
                }
            ],
            "source_page": [{"source_page_id": "sp-1", "url": "https://motortrade.com.ph/motorcycles/airblade"}],
            "pricing": [
                {
                    "pricing_id": "pricing-1",
                    "motorcycle_id": "motorcycle-airblade",
                    "source_context": "motorcycle_detail_page",
                    "srp_amount": 125900,
                    "downpayment_amount": 12000,
                    "monthly_amount": 5800,
                }
            ],
            "motorcycle_spec": [
                {
                    "motorcycle_id": "motorcycle-airblade",
                    "spec_group": "Specifications",
                    "spec_label": "Displacement (cc)",
                    "spec_value_raw": "157cc",
                }
            ],
            "branch": [{"branch_id": "branch-1", "branch_name": "Sample Branch"}],
            "branch_product_coverage": [
                {
                    "branch_id": "branch-1",
                    "coverage_type": "category",
                    "coverage_value_raw": "33",
                    "coverage_value_normalized": "33",
                    "notes": "Not exact model availability.",
                }
            ],
            "branch_model_availability": [],
            "relationship_probe": [
                {
                    "probe_id": "probe-1",
                    "relationship_name": "branch_to_motorcycle_model_availability",
                    "status": "coverage_only",
                    "evidence_found": False,
                    "rows_created": 0,
                    "notes": "Only coverage-level fields were present.",
                }
            ],
            "promo": [],
            "motorcycle_promo_link": [],
            "financing": [],
            "form_definition": [
                {
                    "form_id": "form-loan-1",
                    "form_name": "Loan Application",
                    "form_category": "loan",
                    "form_action": "https://forms.motortrade.com.ph/Loan/tabsix",
                }
            ],
            "form_field": [
                {
                    "form_id": "form-loan-1",
                    "field_order": 1,
                    "field_name": "first_name",
                    "field_label": "First Name",
                    "field_type": "text",
                }
            ],
            "service_offering": [],
            "job_listing": [],
            "parts_catalog": [],
            "preowned_motorcycle": [],
            "external_cta": [],
            "taxonomy_term": [],
            "filter_option": [],
            "data_quality_issue": [],
        }

        section_rows = build_section_rows(tables)

        self.assertIn("motorcycle_catalog", SECTION_SHEETS)
        self.assertIn("branch_coverage_notes", SECTION_SHEETS)
        self.assertEqual(section_rows["motorcycle_catalog"][0]["srp_amount"], 125900)
        self.assertEqual(section_rows["motorcycle_catalog"][0]["source_page_url"], "https://motortrade.com.ph/motorcycles/airblade")
        self.assertEqual(section_rows["motorcycle_specs"][0]["display_name"], "HONDA AIRBLADE 160")
        self.assertEqual(section_rows["forms"][0]["form_name"], "Loan Application")
        self.assertEqual(section_rows["branch_coverage_notes"][0]["note_type"], "coverage_only")
        self.assertEqual(section_rows["branch_coverage_notes"][1]["status"], "coverage_only")
        self.assertNotIn("branch_model_availability", SECTION_SHEETS)

    def test_build_database_rows_for_analyst_workbook(self) -> None:
        tables = {
            "dealer": [{"dealer_id": "motortrade", "dealer_name": "Motortrade"}],
            "crawl_run": [{"run_id": "run-1", "dealer_id": "motortrade"}],
            "source_page": [{"source_page_id": "sp-1", "url": "https://motortrade.com.ph/motorcycles/airblade"}],
            "motorcycle": [
                {
                    "motorcycle_id": "motorcycle-airblade",
                    "dealer_id": "motortrade",
                    "source_page_id": "sp-1",
                    "display_name": "HONDA AIRBLADE 160",
                    "brand_name": "HONDA",
                    "model_name": "Airblade 160",
                    "vehicle_type": "regular-bike",
                    "detail_url": "https://motortrade.com.ph/motorcycles/airblade",
                }
            ],
            "pricing": [
                {
                    "pricing_id": "pricing-1",
                    "motorcycle_id": "motorcycle-airblade",
                    "dealer_id": "motortrade",
                    "source_context": "motorcycle_detail_page",
                    "srp_amount": 125900,
                }
            ],
            "motorcycle_spec": [
                {
                    "motorcycle_id": "motorcycle-airblade",
                    "spec_label": "Displacement (cc)",
                    "spec_value_raw": "157cc",
                }
            ],
            "promo": [{"promo_id": "promo-1", "dealer_id": "motortrade", "title": "Blue Tag"}],
            "motorcycle_promo_link": [
                {
                    "motorcycle_id": "motorcycle-airblade",
                    "promo_id": "promo-1",
                    "link_type": "vehicle_promo_query",
                }
            ],
            "form_definition": [{"form_id": "form-1", "dealer_id": "motortrade", "form_name": "Loan"}],
            "form_field": [{"form_id": "form-1", "field_order": 1, "field_name": "first_name"}],
            "filter_option": [{"filter_name": "mot-type", "option_value": "regular-bike", "option_label": "Regular Bike"}],
            "branch": [],
            "branch_product_coverage": [],
            "branch_model_availability": [],
            "relationship_probe": [],
            "financing": [],
            "service_offering": [],
            "job_listing": [],
            "parts_catalog": [],
            "preowned_motorcycle": [],
            "external_cta": [],
            "taxonomy_term": [],
            "data_quality_issue": [],
        }

        database_rows = build_database_rows(tables)

        self.assertIn("motorcycle", DATABASE_SHEETS)
        self.assertNotIn("taxonomy_term", DATABASE_SHEETS)
        self.assertNotIn("taxonomy_assignment", DATABASE_SHEETS)
        self.assertEqual(database_rows["motorcycle"][0]["vehicle_type"], "regular-bike")
        self.assertEqual(database_rows["pricing"][0]["display_name"], "HONDA AIRBLADE 160")
        self.assertEqual(database_rows["motorcycle_promo_link"][0]["promo_title"], "Blue Tag")
        self.assertEqual(database_rows["form_field"][0]["dealer_id"], "motortrade")
        self.assertEqual(database_rows["website_filter_reference"][0]["dealer_id"], "motortrade")
        self.assertIn("Do not join", database_rows["website_filter_reference"][0]["notes"])

    def test_extract_form_fields(self) -> None:
        soup = BeautifulSoup(
            """
            <form action="/Loan/tabsix" method="POST">
              <div class="form-group">
                <label for="first_name">First Name</label>
                <input id="first_name" name="first_name" required placeholder="Juan" />
              </div>
              <div class="form-group">
                <label for="loan_type">Loan Type</label>
                <select id="loan_type" name="loan_type">
                  <option value="">Select</option>
                  <option value="new">New</option>
                  <option value="used">Used</option>
                </select>
              </div>
              <div class="form-group">
                <label><input type="checkbox" name="agreement" value="yes" /> I agree</label>
              </div>
            </form>
            """,
            "lxml",
        )
        form = soup.find("form")
        fields = extract_form_fields(form, "https://forms.motortrade.com.ph/loan")
        self.assertEqual(fields[0]["field_name"], "first_name")
        self.assertEqual(fields[0]["field_label"], "First Name")
        self.assertTrue(fields[0]["is_required"])
        self.assertEqual(len(fields[1]["options"]), 3)
        self.assertEqual(fields[2]["field_type"], "checkbox")
        self.assertEqual(fields[2]["options"][0]["option_label"], "I agree")


if __name__ == "__main__":
    unittest.main()
