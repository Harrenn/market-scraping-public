from __future__ import annotations

import unittest

from bs4 import BeautifulSoup

from dealer_scrapers.kservico.scrape_kservico_database import (
    extract_product_price,
    flatten_specs,
    infer_vehicle_type,
    parse_specs_from_html,
    parse_store_page,
)


class KservicoDatabaseScraperTests(unittest.TestCase):
    def test_parse_specs_from_inline_motorcycle_text(self) -> None:
        specs = parse_specs_from_html(
            """
            <p>Engine Type 4-stroke, SOHC</p>
            <p>Displacement 157 cc</p>
            <p>Starting System Electric</p>
            <p>Brake type (Front) Hydraulic Disk</p>
            <p>Brake Type (Rear) Drum</p>
            """
        )
        flattened = flatten_specs(specs)

        self.assertEqual(specs[0]["label"], "Engine Type")
        self.assertEqual(specs[0]["value"], "4-stroke, SOHC")
        self.assertEqual(flattened["spec_engine_type"], "4-stroke, SOHC")
        self.assertEqual(flattened["spec_displacement_cc"], "157 cc")
        self.assertIn("Hydraulic Disk", flattened["spec_brake_system"])

    def test_parse_specs_from_alternating_lines(self) -> None:
        specs = parse_specs_from_html(
            """
            <p>Battery</p>
            <p>48v 20ah Lead-Acid</p>
            <p>Motor</p>
            <p>500w</p>
            <p>Range</p>
            <p>35-40kms</p>
            """
        )

        self.assertEqual(specs[0], {"group": "Specifications", "label": "Battery", "value": "48v 20ah Lead-Acid"})
        self.assertEqual(specs[1], {"group": "Specifications", "label": "Motor", "value": "500w"})

    def test_extract_product_price_from_jsonld(self) -> None:
        soup = BeautifulSoup(
            """
            <script type="application/ld+json">
            {"@context":"https://schema.org/","@graph":[
              {"@type":"Product","name":"Honda ADV160","offers":[{"@type":"Offer","price":"167500.00","priceCurrency":"PHP"}]}
            ]}
            </script>
            """,
            "lxml",
        )

        price = extract_product_price(soup)

        self.assertEqual(price["srp_amount"], 167500.0)

    def test_infer_vehicle_type_from_category_slugs(self) -> None:
        self.assertEqual(infer_vehicle_type({"motorcycles", "honda"}), "regular-bike")
        self.assertEqual(infer_vehicle_type({"motorcycles", "big-bikes", "kawasaki"}), "big-bike")
        self.assertEqual(infer_vehicle_type({"e-bike", "hatasu"}), "electric")
        self.assertEqual(infer_vehicle_type({"bajaj-re", "bajaj"}), "tri-wheeler")

    def test_parse_store_page(self) -> None:
        soup = BeautifulSoup(
            """
            <main>
              <h1>Kservico - Narvacan</h1>
              <p>Address</p><p>Villoria, San Jose Narvacan Ilocos Sur 2704</p><p>Get Directions</p>
              <p>Contact</p><p>Mobile:</p><p>+63 9989651387</p><p>E-mail:</p><p>branch@example.com</p>
              <p>Store Opening Hours</p><p>Monday:</p><p>8:00 AM - 5:00 PM</p>
              <p>About Us</p><p>Motorcycle, Gadgets, Appliances, and Furniture Retail Shop At Narvacan</p>
              <p>360 Street View</p>
            </main>
            """,
            "lxml",
        )

        branch = parse_store_page(soup, "Kservico - Narvacan")

        self.assertEqual(branch["postal_code"], "2704")
        self.assertEqual(branch["phone"], "+63 9989651387")
        self.assertIn("Motorcycle", branch["product_raw"])


if __name__ == "__main__":
    unittest.main()
