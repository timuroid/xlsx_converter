from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any
import tomllib


DEFAULT_RULES_PATH = Path(__file__).resolve().parents[1] / "rules.toml"


@dataclass(frozen=True)
class SheetRules:
    header_row: int
    first_data_row: int
    sap_last_column: str
    calculation_date_cell: str


@dataclass(frozen=True)
class FilterRules:
    document_type_column: str
    process_document_type: str


@dataclass(frozen=True)
class InputColumns:
    payment_date_column: str
    clearing_date_column: str
    amount_column: str


@dataclass(frozen=True)
class LogicRules:
    fill_base_calculations_for_dr: bool
    clear_calculation_cells_for_non_dr: bool
    keep_source_filters: bool
    unhide_rows_when_clearing_filters: bool


@dataclass(frozen=True)
class TransformerRules:
    sheet: SheetRules
    filters: FilterRules
    columns: dict[str, dict[str, str]]
    inputs: InputColumns
    penalty_rate: float
    logic: LogicRules

    @property
    def calculation_columns(self) -> list[str]:
        return list(self.columns.keys())


def load_rules(path: str | Path | None = None) -> TransformerRules:
    rules_path = Path(path) if path else DEFAULT_RULES_PATH
    with rules_path.open("rb") as file:
        raw: dict[str, Any] = tomllib.load(file)

    return TransformerRules(
        sheet=SheetRules(**raw["sheet"]),
        filters=FilterRules(**raw["filters"]),
        columns=raw["columns"],
        inputs=InputColumns(**raw["inputs"]),
        penalty_rate=float(raw["values"]["penalty_rate"]),
        logic=LogicRules(**raw["logic"]),
    )
