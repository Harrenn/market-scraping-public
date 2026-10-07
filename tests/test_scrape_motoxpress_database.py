from __future__ import annotations

import unittest

from dealer_scrapers.motoxpress.scrape_motoxpress_database import (
    flatten_specs,
    infer_vehicle_type,
    motorcycle_specs,
    parse_amount,
    parse_published_price,
    parse_google_maps_coordinates,
    rich_text_to_text,
)


class MotoXpressDatabaseScraperTests(unittest.TestCase):
    def test_parse_amount_handles_api_numeric_strings(self) -> None:
        self.assertEqual(parse_amount("23000"), 23000.0)
        self.assertEqual(parse_amount("23,000"), 23000.0)
        self.assertEqual(parse_amount("₱139,900"), 139900.0)
        self.assertEqual(parse_amount(0), 0.0)
        self.assertEqual(parse_amount(""), "")

    def test_parse_published_price_treats_api_zero_as_missing(self) -> None:
        self.assertEqual(parse_published_price(0), "")
        self.assertEqual(parse_published_price("0"), "")
        self.assertEqual(parse_published_price("25,100"), 25100.0)

    def test_motorcycle_specs_and_flattened_columns(self) -> None:
        motorcycle = {
            "fuel_type": "Gasoline",
            "transmission": "Automatic",
            "engine_type": "4-Stroke, SOHC",
            "displacement": "156.9 cc",
            "front_brakes_types": "Disc",
            "rear_brakes_types": "Drum",
            "images": [{"color": "Matte Red"}, {"color": "Pearl White"}],
        }

        specs = motorcycle_specs(motorcycle)
        flattened = flatten_specs(specs, "Honda", "ADV 160", motorcycle)

        self.assertIn({"group": "Color Options", "label": "Color", "value": "Matte Red"}, specs)
        self.assertEqual(flattened["spec_make"], "Honda")
        self.assertEqual(flattened["spec_model"], "ADV 160")
        self.assertEqual(flattened["spec_transmission"], "Automatic")
        self.assertEqual(flattened["spec_displacement_cc"], "156.9 cc")
        self.assertIn("Front Brakes Type: Disc", flattened["spec_brake_system"])

    def test_infer_vehicle_type_from_fuel_and_displacement(self) -> None:
        self.assertEqual(infer_vehicle_type({"fuel_type": "Electric", "displacement": ""}), "electric")
        self.assertEqual(infer_vehicle_type({"fuel_type": "Gasoline", "displacement": "650 cc"}), "big-bike")
        self.assertEqual(infer_vehicle_type({"fuel_type": "Gasoline", "displacement": "125 cc"}), "regular-bike")

    def test_parse_google_maps_coordinates(self) -> None:
        self.assertEqual(
            parse_google_maps_coordinates("https://maps.google.com/maps?q=branch/@13.8718511,121.1992169,17z"),
            ("13.8718511", "121.1992169"),
        )
        self.assertEqual(
            parse_google_maps_coordinates("https://maps.google.com/?q=x!8m2!3d14.5804852!4d121.1744551"),
            ("14.5804852", "121.1744551"),
        )

    def test_rich_text_to_text(self) -> None:
        value = [{"children": [{"text": "Apply online"}, {"text": " with MotoXpress."}]}]

        self.assertEqual(rich_text_to_text(value), "Apply online with MotoXpress.")


if __name__ == "__main__":
    unittest.main()
