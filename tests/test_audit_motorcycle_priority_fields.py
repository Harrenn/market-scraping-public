from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from scripts.audit_motorcycle_priority_fields import audit_workbook


class AuditMotorcyclePriorityFieldsTests(unittest.TestCase):
    def test_promotional_heading_cannot_pass_as_model_name(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            workbook_path = Path(temporary_directory) / "dealer.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "motorcycle"
            sheet.append(
                [
                    "motorcycle_id",
                    "dealer_id",
                    "brand_name",
                    "model_name",
                    "srp_amount",
                    "downpayment_amount",
                    "monthly_amount",
                ]
            )
            sheet.append(
                [
                    "motorcycle-kymco-ak550",
                    "motortrade",
                    "KYMCO",
                    "REBATE : 50,000",
                    519000,
                    104000,
                    15200,
                ]
            )
            workbook.save(workbook_path)

            result = audit_workbook(workbook_path, "motortrade")

        self.assertEqual(result["invalid_model_names"], ["REBATE : 50,000"])
        self.assertTrue(
            any("promotional or SEO text" in error for error in result["errors"])
        )


if __name__ == "__main__":
    unittest.main()
