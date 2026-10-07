#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from openpyxl import load_workbook


DEALER_FILES = {
    "cdo2cycles": "cdo2cycles_database.xlsx",
    "dueksam": "dueksam_database.xlsx",
    "guanzon_motorcycle_group": "guanzon_database.xlsx",
    "kservico": "kservico_database.xlsx",
    "motorace": "motorace_database.xlsx",
    "motorcentralgroup": "motorcentralgroup_database.xlsx",
    "motortrade": "motortrade_database.xlsx",
    "motoxpress": "motoxpress_database.xlsx",
    "premio": "premio_database.xlsx",
    "premiumbikes": "premiumbikes_database.xlsx",
    "ropali": "ropali_database.xlsx",
    "wheeltek": "wheeltek_database.xlsx",
}
PRIORITY_FIELDS = [
    "brand_name",
    "model_name",
    "srp_amount",
    "downpayment_amount",
    "monthly_amount",
]
PRICE_FIELDS = ["srp_amount", "downpayment_amount", "monthly_amount"]
INVALID_MODEL_NAME_PATTERN = re.compile(
    r"^\s*(?:rebate|promo(?:tion)?|discount|save)\b"
    r"|\b(?:price|prices|spec|specs|feature|features)\s*&\s*"
    r"(?:price|prices|spec|specs|feature|features|srp)\b"
    r"|\bprice\s+philippines\b"
    r"|\b(?:bike|scooter|maxi-scooter)\s+(?:ph|philippines)\b",
    flags=re.IGNORECASE,
)


def is_blank(value: object) -> bool:
    return value is None or value == ""


def audit_workbook(path: Path, expected_dealer_id: str) -> dict[str, object]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    if "motorcycle" not in workbook.sheetnames:
        return {"dealer_id": expected_dealer_id, "path": str(path), "errors": ["missing motorcycle sheet"]}

    sheet = workbook["motorcycle"]
    rows = sheet.iter_rows(values_only=True)
    headers = next(rows)
    columns = {str(header): index for index, header in enumerate(headers)}
    required = ["motorcycle_id", "dealer_id", *PRIORITY_FIELDS]
    missing_columns = [field for field in required if field not in columns]
    if missing_columns:
        return {
            "dealer_id": expected_dealer_id,
            "path": str(path),
            "errors": [f"missing columns: {', '.join(missing_columns)}"],
        }

    records = [row for row in rows if any(not is_blank(value) for value in row)]
    errors: list[str] = []
    motorcycle_ids = [str(row[columns["motorcycle_id"]] or "") for row in records]
    if len(motorcycle_ids) != len(set(motorcycle_ids)):
        errors.append("duplicate motorcycle_id values")
    if any(row[columns["dealer_id"]] != expected_dealer_id for row in records):
        errors.append("dealer_id mismatch")
    if any(is_blank(row[columns["model_name"]]) for row in records):
        errors.append("blank model_name")
    invalid_model_names = [
        str(row[columns["model_name"]])
        for row in records
        if INVALID_MODEL_NAME_PATTERN.search(str(row[columns["model_name"]] or ""))
    ]
    if invalid_model_names:
        errors.append(
            "promotional or SEO text found in model_name: "
            + ", ".join(sorted(set(invalid_model_names)))
        )

    zero_counts = {
        field: sum(
            isinstance(row[columns[field]], (int, float)) and row[columns[field]] == 0
            for row in records
        )
        for field in PRICE_FIELDS
    }
    if zero_counts["srp_amount"]:
        errors.append("zero SRP values must be source-verified or normalized to blank")
    if zero_counts["monthly_amount"]:
        errors.append("zero monthly values must be source-verified or normalized to blank")
    if zero_counts["downpayment_amount"] and expected_dealer_id != "motorcentralgroup":
        errors.append("zero downpayment values are only source-verified for Motorcentral")

    populated = {
        field: sum(not is_blank(row[columns[field]]) for row in records)
        for field in PRIORITY_FIELDS
    }
    return {
        "dealer_id": expected_dealer_id,
        "path": str(path),
        "rows": len(records),
        "populated": populated,
        "zero_counts": zero_counts,
        "invalid_model_names": invalid_model_names,
        "errors": errors,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Audit the six priority motorcycle output fields.")
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=Path("working/outputs/dealer-databases"),
        help="Directory containing the 12 dealer database workbooks.",
    )
    parser.add_argument("--json", action="store_true", help="Print JSON instead of the Markdown summary.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    results = []
    for dealer_id, filename in DEALER_FILES.items():
        path = args.input_dir / filename
        if not path.exists():
            results.append({"dealer_id": dealer_id, "path": str(path), "errors": ["missing workbook"]})
            continue
        results.append(audit_workbook(path, dealer_id))

    if args.json:
        print(json.dumps(results, indent=2))
    else:
        print("| Dealer | Rows | Brand | Model | SRP | Downpayment | Monthly | Zeros (SRP/DP/Monthly) |")
        print("| --- | ---: | ---: | ---: | ---: | ---: | ---: | --- |")
        for result in results:
            populated = result.get("populated", {})
            zeros = result.get("zero_counts", {})
            print(
                f"| {result['dealer_id']} | {result.get('rows', 0)} "
                f"| {populated.get('brand_name', 0)} | {populated.get('model_name', 0)} "
                f"| {populated.get('srp_amount', 0)} | {populated.get('downpayment_amount', 0)} "
                f"| {populated.get('monthly_amount', 0)} "
                f"| {zeros.get('srp_amount', 0)}/{zeros.get('downpayment_amount', 0)}/{zeros.get('monthly_amount', 0)} |"
            )
        errors = [
            f"{result['dealer_id']}: {error}"
            for result in results
            for error in result.get("errors", [])
        ]
        if errors:
            print("\nErrors:")
            for error in errors:
                print(f"- {error}")

    return 1 if any(result.get("errors") for result in results) else 0


if __name__ == "__main__":
    raise SystemExit(main())
