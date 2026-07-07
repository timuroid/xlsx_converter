from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

from .rules import DEFAULT_RULES_PATH, load_rules
from .transformer import transform_workbook


def main() -> int:
    parser = argparse.ArgumentParser(
        prog="sap_xlsx_transformer",
        description="Add calculation formulas to SAP XLSX export.",
    )
    parser.add_argument("input", type=Path, help="Path to source SAP XLSX file.")
    parser.add_argument("output", type=Path, help="Path for transformed XLSX file.")
    parser.add_argument(
        "--rules",
        type=Path,
        default=DEFAULT_RULES_PATH,
        help="Path to TOML rules config.",
    )
    parser.add_argument(
        "--calc-date",
        type=date.fromisoformat,
        default=None,
        help="Calculation date in YYYY-MM-DD format. Defaults to today.",
    )
    args = parser.parse_args()

    rules = load_rules(args.rules)
    summary = transform_workbook(args.input, args.output, rules, args.calc_date)

    print(f"Output: {summary.output_path}")
    print(f"Sheet: {summary.worksheet_name}")
    print(f"Data rows: {summary.max_row - rules.sheet.first_data_row + 1}")
    print(f"DR rows processed: {summary.processed_dr_rows}")
    print(f"Base formula rows: {summary.base_formula_rows}")

    return 0
