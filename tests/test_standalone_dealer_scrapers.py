from __future__ import annotations

import importlib.util
import sys
import unittest
from pathlib import Path

from bs4 import BeautifulSoup


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = {
    "motorace": REPO_ROOT / "dealer_scrapers" / "motorace" / "scrape_motorace_database.py",
    "motorcentralgroup": REPO_ROOT / "dealer_scrapers" / "motorcentralgroup" / "scrape_motorcentralgroup_database.py",
    "premiumbikes": REPO_ROOT / "dealer_scrapers" / "premiumbikes" / "scrape_premiumbikes_database.py",
    "premio": REPO_ROOT / "dealer_scrapers" / "premio" / "scrape_premio_database.py",
    "wheeltek": REPO_ROOT / "dealer_scrapers" / "wheeltek" / "scrape_wheeltek_database.py",
    "cdo2cycles": REPO_ROOT / "dealer_scrapers" / "cdo2cycles" / "scrape_cdo2cycles_database.py",
    "dueksam": REPO_ROOT / "dealer_scrapers" / "dueksam" / "scrape_dueksam_database.py",
    "ropali": REPO_ROOT / "dealer_scrapers" / "ropali" / "scrape_ropali_database.py",
}
CLASS_NAMES = {
    "motorace": "MotorAceDatabaseScraper",
    "motorcentralgroup": "MotorcentralGroupDatabaseScraper",
    "premiumbikes": "PremiumbikesDatabaseScraper",
    "premio": "PremioDatabaseScraper",
    "wheeltek": "WheeltekDatabaseScraper",
    "cdo2cycles": "Cdo2CyclesDatabaseScraper",
    "dueksam": "DueksamDatabaseScraper",
    "ropali": "RopaliDatabaseScraper",
}


def load_script(name: str):
    spec = importlib.util.spec_from_file_location(f"test_{name}_scraper", SCRIPTS[name])
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


class StandaloneDealerScraperTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.modules = {name: load_script(name) for name in SCRIPTS}

    def test_each_batch_dealer_script_is_standalone_and_dealer_named(self) -> None:
        for name, script in SCRIPTS.items():
            source = script.read_text()
            self.assertNotIn("Database.common_dealer_scraper", source)
            self.assertNotIn("class GenericDealerScraper", source)
            self.assertIn(f"class {CLASS_NAMES[name]}", source)

    def test_motorace_parses_woocommerce_spec_table(self) -> None:
        soup = BeautifulSoup("<table><tr><th>Displacement</th><td>155 cm3</td></tr><tr><th>Engine Type</th><td>SOHC</td></tr></table>", "lxml")

        specs = self.modules["motorace"].extract_specs_from_soup(soup)

        self.assertIn({"group": "Specifications", "label": "Displacement", "value": "155 cm3"}, specs)
        self.assertIn({"group": "Specifications", "label": "Engine Type", "value": "SOHC"}, specs)

    def test_motorace_classification_uses_product_evidence_not_related_page_noise(self) -> None:
        module = self.modules["motorace"]
        vehicle_type, _source = module.infer_vehicle_type("Yamaha Sniper 155 /motorcycles/sniper-155 Displacement 155 cm3")

        self.assertEqual(vehicle_type, "regular-bike")
        self.assertNotEqual(vehicle_type, "electric")

    def test_motorace_zero_store_price_falls_back_to_labelled_detail_srp(self) -> None:
        module = self.modules["motorace"]
        scraper = module.MotorAceDatabaseScraper(module.CONFIG, Path("/tmp/motorace-test.xlsx"))

        amount = scraper.product_price(
            {"prices": {"price": "0", "regular_price": "0", "currency_minor_unit": 2}},
            "Sniper 155 SRP: \u20b1148,250.00 Related products SRP: \u20b152,080.00",
        )

        self.assertEqual(amount, 148250.0)

    def test_motorace_parses_public_elfsight_branch_payload(self) -> None:
        module = self.modules["motorace"]
        app_id = "c2a99b6e-28aa-4d5a-a503-b0f9f179ac9e"
        payload = {
            "data": {
                "widgets": {
                    app_id: {
                        "data": {
                            "settings": {
                                "locations": [
                                    {
                                        "id": "branch-1",
                                        "name": "OSLOB - MAPI",
                                        "address": "GCMR+53, Oslob, Cebu, Philippines",
                                        "phone": "+63 917 525 6798",
                                        "email": "bc.oslob@motorace.ph",
                                        "website": "https://motorace.ph/",
                                        "customCoordinates": {"lat": 9.5329567, "lng": 123.4402034},
                                        "dayMondayOpen": True,
                                        "dayMondayHours": [{"timeRange": ["08:30", "17:30"]}],
                                    }
                                ]
                            }
                        }
                    }
                }
            }
        }

        branches = module.parse_motorace_elfsight_branches(payload, app_id)

        self.assertEqual(module.extract_motorace_elfsight_app_id(f'<div class="elfsight-app-{app_id}"></div>'), app_id)
        self.assertEqual(len(branches), 1)
        self.assertEqual(branches[0]["branch_name"], "OSLOB - MAPI")
        self.assertEqual(branches[0]["latitude"], 9.5329567)
        self.assertIn("Monday", branches[0]["days_str"])

    def test_motorcentral_parses_branch_contact_detail_text(self) -> None:
        module = self.modules["motorcentralgroup"]
        scraper = module.MotorcentralGroupDatabaseScraper(module.CONFIG, Path("/tmp/motorcentral-test.xlsx"))

        scraper.parse_branch_html(
            "https://motorcentralgroup.com/branches/",
            """
            <main>
              <p>Bajaj</p>
              <p>BAJAJ RE Tanay Branch</p><p>Contact Details:</p>
              <p>Exit Rd., Brgy. Plaza Aldea, Tanay, Rizal</p>
              <p>tanay@example.com</p><p>09171234567</p>
            </main>
            """,
        )

        self.assertEqual(scraper.tables.row_count("branch"), 1)
        self.assertEqual(scraper.tables.rows["branch"][0]["branch_name"], "BAJAJ RE Tanay Branch")

    def test_motorcentral_parses_sequential_detail_specs_without_merging_labels(self) -> None:
        module = self.modules["motorcentralgroup"]
        soup = BeautifulSoup(
            """
            <main>
              <h2>Specifications</h2>
              <p>Engine Type</p><p>4-Stroke, 1-cylinder, SOHC</p>
              <p>Displacement</p><p>249 CC</p>
              <p>Starting System</p><p>Electric</p>
              <p>Ignition System</p><p>Electronic Ignition</p>
              <p>Related Motorcycle</p><p>CRF150L</p>
            </main>
            """,
            "lxml",
        )

        specs = module.extract_specs_from_soup(soup)

        self.assertIn({"group": "Specifications", "label": "Engine Type", "value": "4-Stroke, 1-cylinder, SOHC"}, specs)
        self.assertIn({"group": "Specifications", "label": "Displacement", "value": "249 CC"}, specs)
        self.assertNotIn({"group": "Specifications", "label": "Engine Type", "value": "4-Stroke, 1-cylinder, SOHC Displacement 249 CC"}, specs)

    def test_motorcentral_parses_labelled_prices_and_rejects_fake_store_price(self) -> None:
        module = self.modules["motorcentralgroup"]
        soup = BeautifulSoup(
            """
            <main>
              <h1 class="product_title">Raider PRO 150</h1>
              <div>Price</div><div>P148,250</div>
              <div>Estimated Down Payment</div><div>P11,400</div>
              <div>Estimated Monthly</div><div>P7,135</div>
              <section>Specifications</section>
              <div>Contact 6619640522</div>
            </main>
            """,
            "lxml",
        )

        pricing = module.extract_motorcentral_pricing_from_soup(soup, "Raider PRO 150")

        self.assertEqual(pricing["srp_amount"], 148250.0)
        self.assertEqual(pricing["downpayment_amount"], 11400.0)
        self.assertEqual(pricing["monthly_amount"], 7135.0)
        self.assertEqual(module.parse_motorcentral_labelled_amount("PHP 6,619,640,522.00"), ("", ""))
        self.assertEqual(module.parse_motorcentral_labelled_amount("150"), ("", ""))

    def test_motorcentral_price_parser_handles_spaced_commas_without_field_spill(self) -> None:
        module = self.modules["motorcentralgroup"]
        soup = BeautifulSoup(
            """
            <main>
              <h1>Bajaj RE</h1>
              <div>Price</div><div>P211, 388</div>
              <div>Estimated Down Payment</div><div>P9,999</div>
              <div>Estimated Monthly</div><div>P9,141</div>
            </main>
            """,
            "lxml",
        )

        pricing = module.extract_motorcentral_pricing_from_soup(soup, "Bajaj RE")

        self.assertEqual(pricing["srp_amount"], 211388.0)
        self.assertEqual(pricing["downpayment_amount"], 9999.0)
        self.assertEqual(pricing["monthly_amount"], 9141.0)

    def test_motorcentral_price_parser_scopes_to_main_product(self) -> None:
        module = self.modules["motorcentralgroup"]
        soup = BeautifulSoup(
            """
            <header><div>Price 59,495 Down Payment 2,500 Monthly 2,762.20</div></header>
            <main>
              <section>
                <h1>PCX160 ABS (Roadsync) \u2013 (7206701)</h1>
                <div>Price</div><div>P135,491.00</div>
                <div>Estimated Down Payment</div><div>P10,000</div>
                <div>Estimated Monthly</div><div>P6,383.92</div>
                <div>PRICE:</div><div>12 MONTHS - 13,794.86</div>
              </section>
              <section>Related Motorcycle Price 87,683 Down Payment 3,500 Monthly 4,101.96</section>
            </main>
            """,
            "lxml",
        )

        pricing = module.extract_motorcentral_pricing_from_soup(
            soup,
            "PCX160 ABS (Roadsync) - (7206701)",
        )

        self.assertEqual(pricing["srp_amount"], 135491.0)
        self.assertEqual(pricing["downpayment_amount"], 10000.0)
        self.assertEqual(pricing["monthly_amount"], 6383.92)

    def test_premiumbikes_parses_product_specs_and_rejects_blog_as_exact_model(self) -> None:
        module = self.modules["premiumbikes"]
        soup = BeautifulSoup("<table><tr><th>Transmission Type</th><td>Automatic</td></tr><tr><th>Fuel System</th><td>PGM-FI</td></tr></table>", "lxml")
        scraper = module.PremiumbikesDatabaseScraper(module.CONFIG, Path("/tmp/premium-test.xlsx"))

        specs = module.extract_specs_from_soup(soup)

        self.assertIn({"group": "Specifications", "label": "Transmission Type", "value": "Automatic"}, specs)
        self.assertFalse(scraper.is_exact_motorcycle_page("https://premiumbikes.ph/brunch-but-make-it-a-ride/", "Kawasaki Brunch", "brand blog article"))

    def test_premiumbikes_parses_public_branch_api_rows(self) -> None:
        module = self.modules["premiumbikes"]
        branches = module.parse_premiumbikes_branch_payload(
            [
                {
                    "id": "1",
                    "branch_name": "Premiumbikes Tabaco, Albay",
                    "address": "P-4 Bonbon Tabaco, Albay Philippines 4511",
                    "email": "tabaco@example.com",
                    "globe": "9778169207",
                    "smart": "9988404882",
                    "landline": "N/A",
                    "latitude": "13.362810",
                    "longitude": "123.727800",
                    "cluster": "BICOL",
                    "region": "BICOL",
                    "location": "SOUTH",
                }
            ]
        )

        self.assertEqual(len(branches), 1)
        self.assertEqual(branches[0]["city"], "Tabaco")
        self.assertEqual(branches[0]["state_or_province"], "Albay")
        self.assertEqual(branches[0]["postal_code"], "4511")
        self.assertEqual(branches[0]["phone"], "9778169207 / 9988404882")

    def test_premio_parses_compact_spec_cells(self) -> None:
        soup = BeautifulSoup(
            """
            <table><tr>
              <td>Engine Type 4-stroke, SOHC</td>
              <td>Displacement 330 cc</td>
              <td>Brake Type (Front) Hydraulic Disk with ABS</td>
            </tr></table>
            """,
            "lxml",
        )

        specs = self.modules["premio"].extract_specs_from_soup(soup)

        self.assertIn({"group": "Specifications", "label": "Engine Type", "value": "4-stroke, SOHC"}, specs)
        self.assertIn({"group": "Specifications", "label": "Brake Type (Front)", "value": "Hydraulic Disk with ABS"}, specs)

    def test_premio_parses_store_directory_cards(self) -> None:
        module = self.modules["premio"]
        branches = module.extract_premio_branches(
            """
            <h4>Cavite/Batangas</h4>
            <ul class="elementor-icon-list-items">
              <li><i class="fas fa-store"></i><span>Premio Yamaha Bauan</span></li>
              <li><i class="fas fa-map-marker-alt"></i><span>Manghinao Proper, Bauan, Batangas</span></li>
              <li><i class="fas fa-phone-alt"></i><span>09178191681</span></li>
            </ul>
            """
        )

        self.assertEqual(len(branches), 1)
        self.assertEqual(branches[0]["branch_name"], "Premio Yamaha Bauan")
        self.assertEqual(branches[0]["store_group"], "Cavite/Batangas")

    def test_wheeltek_does_not_treat_products_landing_page_as_exact_model(self) -> None:
        module = self.modules["wheeltek"]
        scraper = module.WheeltekDatabaseScraper(module.CONFIG, Path("/tmp/wheeltek-test.xlsx"))

        self.assertFalse(scraper.is_exact_motorcycle_page("https://wheeltek.com.ph/products/", "Products", "Small Bikes Big Bikes Cars Watercrafts"))
        self.assertTrue(
            scraper.is_wheeltek_motorcycle_record(
                "vehicles",
                "https://wheeltek.com.ph/vehicles/raider-r150-fi/",
                "Back to Regular Bikes Raider R150 FI Specifications Engine Type Displacement motorcycle",
            )
        )
        self.assertFalse(
            scraper.is_wheeltek_motorcycle_record(
                "vehicles",
                "https://wheeltek.com.ph/vehicles/soluto-1-4-ex-at/",
                "Kia Soluto Sedan Gasoline engine displacement 1368 cc",
            )
        )

    def test_wheeltek_listing_cards_supply_brand_price_and_active_scope(self) -> None:
        module = self.modules["wheeltek"]
        soup = BeautifulSoup(
            """
            <div class="product-wrapper" data-brand="cf moto">
              <div class="product-details">
                <h2 class="title">450CL-C</h2>
                <p class="price"><strong>SRP:</strong> P 287,900</p>
                <a class="btn-learn" href="/vehicles/450cl-c/">Learn More</a>
              </div>
            </div>
            """,
            "lxml",
        )

        cards = module.parse_wheeltek_listing_cards(
            soup,
            "https://wheeltek.com.ph/products/big-bikes/",
            "big-bike",
        )

        self.assertEqual(len(cards), 1)
        self.assertEqual(cards[0]["brand_name"], "CFMOTO")
        self.assertEqual(cards[0]["srp_amount"], 287900.0)
        self.assertEqual(cards[0]["vehicle_type"], "big-bike")

        scraper = module.WheeltekDatabaseScraper(module.CONFIG, Path("/tmp/wheeltek-test.xlsx"))
        scraper.wheeltek_listing_by_url = {cards[0]["detail_url"]: cards[0]}
        self.assertTrue(scraper.is_wheeltek_motorcycle_record("vehicles", cards[0]["detail_url"], ""))
        self.assertFalse(
            scraper.is_wheeltek_motorcycle_record(
                "watercraft",
                "https://wheeltek.com.ph/vehicles/ultra-310x/",
                "Ultra 310X",
            )
        )

    def test_cdo2cycles_parses_wix_sequential_specs(self) -> None:
        soup = BeautifulSoup(
            """
            <main>
              <h2>SPECIFICATION</h2>
              <p>ENGINE</p>
              <p>Engine</p><p>Cylinder</p><p>Displacement</p><p>Bore & Stroke</p>
              <p>Compression Ratio</p><p>Maximum Horse Power</p><p>Maximum Torque</p>
              <p>Starting System</p><p>Lubrication System</p><p>Engine Oil Capacity</p>
              <p>Clutch Type</p><p>Transmission Type</p><p>Fuel Tank Capacity</p><p>Fuel Supply</p>
              <p>Liquid cooled</p><p>Single cylinder</p><p>155cm3</p><p>58.0 x 58.7mm</p>
              <p>11.6</p><p>11.3kW</p><p>14.2N-m</p><p>Electric starter</p>
              <p>Wet sump</p><p>1.10L</p><p>Dry, centrifugal, shoe</p>
              <p>V-belt automatic</p><p>5.5L</p><p>Fuel injection</p>
              <p>CHASSIS</p>
              <p>Frame Type</p><p>Front Suspension</p><p>Rear Suspension</p><p>Front Tire</p>
              <p>Rear Tire</p><p>Front Brake</p><p>Rear Brake</p>
              <p>Backbone</p><p>Telescopic fork</p><p>Unit swing</p><p>110/80-14 Tubeless</p>
              <p>140/70-14 Tubeless</p><p>Hydraulic single disc brake</p><p>Hydraulic single disc brake</p>
            </main>
            """,
            "lxml",
        )

        specs = self.modules["cdo2cycles"].extract_specs_from_soup(soup)

        self.assertIn({"group": "Engine", "label": "Displacement", "value": "155cm3"}, specs)
        self.assertIn({"group": "Engine", "label": "Transmission Type", "value": "V-belt automatic"}, specs)
        self.assertIn({"group": "Chassis", "label": "Front Tire", "value": "110/80-14 Tubeless"}, specs)
        self.assertIn({"group": "Chassis", "label": "Rear Brake", "value": "Hydraulic single disc brake"}, specs)

    def test_cdo2cycles_does_not_fabricate_specs_from_non_cdo_blocks(self) -> None:
        soup = BeautifulSoup(
            """
            <main>
              <div>Engine Type Visit our showroom for more details</div>
              <div>Displacement Ask branch staff for current availability</div>
            </main>
            """,
            "lxml",
        )

        specs = self.modules["cdo2cycles"].extract_specs_from_soup(soup)

        self.assertEqual(specs, [])

    def test_cdo2cycles_extracts_exact_local_business_branch(self) -> None:
        module = self.modules["cdo2cycles"]
        record = module.extract_cdo_local_business(
            """
            <script type="application/ld+json">
            {"@type":"LocalBusiness","name":"CDO 2 Cycles","telephone":"09967896541","address":{"@type":"PostalAddress","streetAddress":"Western Kolambog Lapasan Highway","addressLocality":"Cagayan de Oro","addressRegion":"10","postalCode":"9000"}}
            </script>
            <script>var x={"businessLocationCoordinates":{"latitude":8.482015,"longitude":124.666704}};</script>
            """
        )

        self.assertEqual(record["name"], "CDO 2 Cycles")
        self.assertEqual(record["full_address"], "Western Kolambog Lapasan Highway, Cagayan de Oro, 10, 9000")
        self.assertEqual(record["latitude"], 8.482015)

    def test_cdo2cycles_parses_rendered_wix_map_locations(self) -> None:
        module = self.modules["cdo2cycles"]
        branches = module.parse_cdo_map_message(
            {
                "type": "SET_INITIAL_LOCATIONS",
                "data": """{"locations":[{"address":"Ground Floor National Highway, Corner Bajas Road, Cagayan de Oro, 9000 Misamis Oriental, Philippines","latitude":8.482015,"longitude":124.666704,"title":"CDO 2 CYCLES - Yamaha Revzone"},{"address":"Apokon Rd, Tagum, Davao del Norte, Philippines","latitude":7.4372702,"longitude":125.8218138,"title":"CDO 2 CYCLES - TAGUM BRANCH"}]}""",
            }
        )

        self.assertEqual(len(branches), 2)
        self.assertEqual(branches[0]["branch_name"], "CDO 2 CYCLES - Yamaha Revzone")
        self.assertEqual(branches[0]["city"], "Cagayan de Oro")
        self.assertEqual(branches[0]["state_or_province"], "Misamis Oriental")
        self.assertEqual(branches[0]["postal_code"], "9000")
        self.assertEqual(branches[1]["latitude"], 7.4372702)

    def test_cdo2cycles_uses_scoped_model_title_and_price(self) -> None:
        module = self.modules["cdo2cycles"]
        soup = BeautifulSoup(
            """
            <html><head><title>RAIDER J CROSSOVER | CDO 2 Cycles</title></head>
            <body><div>CDO 2 Cycles</div><h2>RAIDER J CROSSOVER</h2><div>₱ 70,900</div><div>P 80</div></body></html>
            """,
            "lxml",
        )

        self.assertEqual(module.extract_cdo_model_title(soup, "https://www.cdo2cycles.com/suzuki-motorcycles/raider-j-crossover"), "RAIDER J CROSSOVER")
        self.assertEqual(module.extract_cdo_page_price(soup), 70900.0)

    def test_dueksam_ajax_rows_create_motorcycle_without_fake_specs(self) -> None:
        module = self.modules["dueksam"]
        scraper = module.DueksamDatabaseScraper(module.CONFIG, Path("/tmp/dueksam-test.xlsx"))

        motorcycle_id = scraper.add_motorcycle(
            "source-1",
            "https://dueksam.com.ph/motorcycles?entity_id=1847",
            "https://dueksam.com.ph/motorcycles",
            "The All-New CUV e",
            "Honda",
            '{"model":"The All-New CUV e","price":"199900.0000"}',
            price=199900.0,
        )

        self.assertEqual(motorcycle_id, "motorcycle-dueksam-motorcycles-entity-id-1847")
        self.assertEqual(scraper.tables.row_count("motorcycle"), 1)
        self.assertEqual(scraper.tables.row_count("motorcycle_spec"), 0)

    def test_dueksam_parses_browser_rendered_detail_specs(self) -> None:
        module = self.modules["dueksam"]

        specs = module.parse_dueksam_specs_from_text(
            """
            DETAILS

            ENGINE

            Engine type\t4-Stroke, 2-Valve, SOHC
            Displacement\t125cc
            Max Torque
            11.0 Nm @ 5,000 rpm
            Brake Type (Front)\tHydraulic Disc
            CONTACT INFORMATION
            """
        )

        self.assertIn({"group": "Engine", "label": "Engine type", "value": "4-Stroke, 2-Valve, SOHC"}, specs)
        self.assertIn({"group": "Engine", "label": "Displacement", "value": "125cc"}, specs)
        self.assertIn({"group": "Engine", "label": "Max Torque", "value": "11.0 Nm @ 5,000 rpm"}, specs)

    def test_dueksam_branch_filter_keeps_only_mc_or_3s_store_signals(self) -> None:
        module = self.modules["dueksam"]

        self.assertTrue(module.dueksam_branch_has_motorcycle_store_signal("BANTAYAN MC"))
        self.assertTrue(module.dueksam_branch_has_motorcycle_store_signal("DANAO HONDA 3S"))
        self.assertTrue(module.dueksam_branch_has_motorcycle_store_signal("MC-MULTI"))
        self.assertFalse(module.dueksam_branch_has_motorcycle_store_signal("LIGAO", "Purok 1 McKinley Street Calzada Ligao City"))
        self.assertFalse(module.dueksam_branch_has_motorcycle_store_signal("Passi Honda Wingshop"))

    def test_ropali_parses_woocommerce_minor_unit_price(self) -> None:
        module = self.modules["ropali"]

        amount = module.store_api_price_amount({"price": "1060000", "currency_minor_unit": 2})

        self.assertEqual(amount, 10600.0)

    def test_ropali_extracts_mapbox_branch_features(self) -> None:
        module = self.modules["ropali"]

        features = module.extract_mapbox_store_features(
            """
            document.addEventListener('DOMContentLoaded',function(){
              const stores={'type':'FeatureCollection','features':[{'type':'Feature','geometry':{'type':'Point','coordinates':[121.0603103,14.5787072]},'properties':{'address':'Ropali Head Office','city':'Pasig','postalCode':'1105','state':'MM'}}]};
              function isValidCoordinate(coord){return true;}
            });
            """
        )

        self.assertEqual(len(features), 1)
        self.assertEqual(features[0]["properties"]["address"], "Ropali Head Office")

    def test_ropali_extracts_product_detail_specs(self) -> None:
        module = self.modules["ropali"]
        soup = BeautifulSoup(
            """
            <table>
              <tr><th>Engine Type</th><td>Liquid cooled, 4-stroke</td></tr>
              <tr><th>Engine Displacement</th><td>155 cm3</td></tr>
            </table>
            """,
            "lxml",
        )

        specs = module.extract_specs_from_soup(soup)

        self.assertIn({"group": "Additional information", "label": "Engine Type", "value": "Liquid cooled, 4-stroke"}, specs)
        self.assertIn({"group": "Additional information", "label": "Engine Displacement", "value": "155 cm3"}, specs)


if __name__ == "__main__":
    unittest.main()
