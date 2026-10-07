from __future__ import annotations

import importlib.util
import unittest
from pathlib import Path

from bs4 import BeautifulSoup


SCRIPT_PATH = (
    Path(__file__).resolve().parents[1]
    / "dealer_scrapers"
    / "guanzon"
    / "scrape_guanzon_database.py"
)
SPEC = importlib.util.spec_from_file_location("scrape_guanzon_database", SCRIPT_PATH)
guanzon = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
SPEC.loader.exec_module(guanzon)


class GuanzonDatabaseScraperTests(unittest.TestCase):
    def test_extract_guanzon_pricing_from_price_card(self) -> None:
        soup = BeautifulSoup(
            """
            <div class="sp-product-price">
              <p>Get yours for only</p><p>₱141/day</p>
              <p>SRP: ₱77,000</p><p>DP: ₱3,000</p>
              <p>36 Months: ₱4,217</p><p>24 Months: ₱5,300</p><p>18 Months: ₱6,363</p>
              <p>*Prices shown are for reference only and may vary by branch.</p>
            </div>
            """,
            "lxml",
        )

        pricing = guanzon.extract_guanzon_pricing(soup)

        self.assertEqual(pricing["srp_amount"], 77000.0)
        self.assertEqual(pricing["downpayment_amount"], 3000.0)
        self.assertEqual(pricing["monthly_36_amount"], 4217.0)
        self.assertEqual(pricing["daily_amount"], 141.0)

    def test_extract_guanzon_specs_from_custom_blocks(self) -> None:
        soup = BeautifulSoup(
            """
            <div class="sp-specification">
              <div class="sp-spicifications-container"><p>Engine Type:</p><p class="sp-specs-value">SOHC</p></div>
              <div class="sp-spicifications-container"><p>Displacement:</p><p class="sp-specs-value">125 cc</p></div>
              <div class="sp-spicifications-container"><p>L x W x H:</p><p class="sp-specs-value">1,890 x 850 x 1,150 (mm)</p></div>
            </div>
            """,
            "lxml",
        )

        specs = guanzon.extract_guanzon_specs(soup)
        flattened = guanzon.flatten_specs(specs)

        self.assertEqual(specs[0], {"group": "Specifications", "label": "Engine Type", "value": "SOHC"})
        self.assertEqual(flattened["spec_displacement_cc"], "125 cc")
        self.assertEqual(flattened["spec_overall_dimensions"], "1,890 x 850 x 1,150 (mm)")

    def test_parse_guanzon_branch_page(self) -> None:
        soup = BeautifulSoup(
            """
            <div class="br-content">
              <div><div class="br-slug-text"><p>Honda 3S Shop</p></div></div>
              <div class="col br-sub-category-container">
                <div class="br-item-total"><p>1</p></div>
                <div class="br-post-meta-text"><p>LGK Guanzon Inc.</p></div>
                <div class="br-post-text"><p>Aguilar</p></div>
                <div class="br-post-address-text"><p>Bldg. RN &amp; D, Romulo Hwy, Brgy. Baybay, Aguilar, Pangasinan</p></div>
                <div class="br-post-cn-text"><p>0917-706-8297 / (075) 636-1715</p></div>
                <div class="br-post-email-text"><p>lgk_aguilar-multi@guanzongroup.com.ph</p></div>
                <div class="br-post-fb-text"><a href="https://facebook.example/branch">Message Us Now</a></div>
                <div class="br-map-text"><a data-store="15.893627,120.237039" data-store-id="9944">View on Map</a></div>
              </div>
            </div>
            """,
            "lxml",
        )

        branches = guanzon.parse_guanzon_branch_page(
            soup,
            "https://www.guanzongroup.com.ph/branches/motorcycle-branches/",
            "motorcycle_branch",
        )

        self.assertEqual(len(branches), 1)
        self.assertEqual(branches[0]["locator_id"], "9944")
        self.assertEqual(branches[0]["branch_name"], "LGK Guanzon Inc. - Aguilar")
        self.assertEqual(branches[0]["state_or_province"], "Pangasinan")
        self.assertEqual(branches[0]["brand_raw"], "Honda")
        self.assertEqual(branches[0]["product_raw"], "motorcycle")

    def test_vehicle_type_uses_displacement_for_big_bikes(self) -> None:
        specs = [{"group": "Specifications", "label": "Displacement", "value": "650 cc"}]

        vehicle_type, source = guanzon.infer_vehicle_type({"motorcycle", "honda"}, ["Sports Bikes"], specs, "CBR650R")

        self.assertEqual(vehicle_type, "big-bike")
        self.assertEqual(source, "spec_displacement_cc")


if __name__ == "__main__":
    unittest.main()
