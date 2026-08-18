"""VLOOKUP-style merge operations with an auditable result."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
from typing import Any, Iterable

import pandas as pd

from .quality import is_missing


@dataclass(frozen=True)
class MatchOptions:
    trim_spaces: bool = True
    case_sensitive: bool = False
    numeric_equivalence: bool = True
    duplicate_strategy: str = "first"

    def __post_init__(self) -> None:
        if self.duplicate_strategy not in {"first", "expand"}:
            raise ValueError("duplicate_strategy must be 'first' or 'expand'.")


@dataclass
class LookupSpec:
    file_name: str
    sheet_name: str
    dataframe: pd.DataFrame
    key_column: str
    return_columns: list[str]
    prefix: str = ""


@dataclass(frozen=True)
class LookupAudit:
    source: str
    sheet: str
    lookup_key: str
    lookup_rows: int
    distinct_keys: int
    duplicate_key_rows: int
    matched_base_rows: int
    unmatched_base_rows: int
    missing_base_keys: int
    rows_after_lookup: int
    returned_columns: int


@dataclass
class MergeResult:
    dataframe: pd.DataFrame
    audits: list[LookupAudit] = field(default_factory=list)
    base_rows: int = 0
    base_key: str = ""

    @property
    def matched_cells_summary(self) -> tuple[int, int]:
        return (
            sum(audit.matched_base_rows for audit in self.audits),
            sum(audit.unmatched_base_rows for audit in self.audits),
        )

    def audit_frame(self) -> pd.DataFrame:
        return pd.DataFrame(
            [
                {
                    "Lookup source": audit.source,
                    "Worksheet": audit.sheet,
                    "Lookup key": audit.lookup_key,
                    "Lookup rows": audit.lookup_rows,
                    "Distinct keys": audit.distinct_keys,
                    "Duplicate-key rows": audit.duplicate_key_rows,
                    "Matched base rows": audit.matched_base_rows,
                    "Unmatched base rows": audit.unmatched_base_rows,
                    "Missing base keys": audit.missing_base_keys,
                    "Rows after lookup": audit.rows_after_lookup,
                    "Columns returned": audit.returned_columns,
                }
                for audit in self.audits
            ]
        )


def default_prefix(file_name: str) -> str:
    stem = Path(file_name).stem.strip()
    cleaned = re.sub(r"[^\w -]+", "", stem, flags=re.UNICODE)
    return re.sub(r"\s+", "_", cleaned).strip("_") or "lookup"


def _canonical_decimal(value: Any) -> str | None:
    try:
        decimal = Decimal(str(value))
    except (InvalidOperation, ValueError):
        return None
    if not decimal.is_finite():
        return None
    normalized = decimal.normalize()
    return format(normalized, "f")


def canonicalize_key(value: Any, options: MatchOptions) -> str | None:
    """Create a stable comparison key without collapsing zero-padded IDs."""

    if is_missing(value):
        return None
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return f"date:{pd.Timestamp(value).isoformat()}"
    if isinstance(value, bool):
        return f"bool:{str(value).lower()}"

    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        numeric = _canonical_decimal(value)
        if numeric is not None:
            return f"number:{numeric}"

    text = str(value)
    if options.trim_spaces:
        text = text.strip()
    if not options.case_sensitive:
        text = text.casefold()
    if not text:
        return None

    if options.numeric_equivalence:
        candidate = text
        signless = candidate[1:] if candidate[:1] in {"+", "-"} else candidate
        integer_part = signless.split(".", 1)[0]
        has_ambiguous_leading_zero = len(integer_part) > 1 and integer_part.startswith("0")
        if not has_ambiguous_leading_zero:
            numeric = _canonical_decimal(candidate)
            if numeric is not None:
                return f"number:{numeric}"
    return f"text:{text}"


def canonicalize_series(series: pd.Series, options: MatchOptions) -> pd.Series:
    return series.map(lambda value: canonicalize_key(value, options), na_action=None)


def _unique_column_name(base: str, existing: Iterable[str]) -> str:
    existing_names = set(existing)
    if base not in existing_names:
        return base
    suffix = 2
    while f"{base}__{suffix}" in existing_names:
        suffix += 1
    return f"{base}__{suffix}"


def _validate_spec(spec: LookupSpec) -> None:
    if spec.key_column not in spec.dataframe.columns:
        raise ValueError(f"Lookup key {spec.key_column!r} is missing from {spec.file_name!r}.")
    missing = [column for column in spec.return_columns if column not in spec.dataframe.columns]
    if missing:
        raise ValueError(f"Return columns missing from {spec.file_name!r}: {', '.join(missing)}")


def merge_lookups(
    base_dataframe: pd.DataFrame,
    base_key: str,
    lookups: list[LookupSpec],
    options: MatchOptions | None = None,
    base_return_columns: list[str] | None = None,
) -> MergeResult:
    """Left-join each lookup to the base table using VLOOKUP-like semantics."""

    options = options or MatchOptions()
    if base_key not in base_dataframe.columns:
        raise ValueError(f"Base lookup key {base_key!r} does not exist.")

    if base_return_columns is None:
        base_columns = list(base_dataframe.columns)
    else:
        missing_base_columns = [
            column for column in base_return_columns if column not in base_dataframe.columns
        ]
        if missing_base_columns:
            raise ValueError(
                "Base return columns do not exist: " + ", ".join(missing_base_columns)
            )
        # The lookup key is always retained, even when it is not selected separately.
        base_columns = [
            base_key,
            *[column for column in base_return_columns if column != base_key],
        ]

    result = base_dataframe[base_columns].copy()
    result["__lookup_base_row"] = range(len(result))
    canonical_base = canonicalize_series(result[base_key], options)
    result["__lookup_key"] = canonical_base
    base_rows = len(result)
    missing_base_keys = int(canonical_base.isna().sum())
    audits: list[LookupAudit] = []

    for spec in lookups:
        _validate_spec(spec)
        prefix = spec.prefix.strip() or default_prefix(spec.file_name)
        lookup = spec.dataframe[[spec.key_column, *spec.return_columns]].copy()
        lookup["__lookup_key"] = canonicalize_series(lookup[spec.key_column], options)
        lookup = lookup[lookup["__lookup_key"].notna()].copy()

        duplicate_mask = lookup["__lookup_key"].duplicated(keep=False)
        duplicate_key_rows = int(duplicate_mask.sum())
        distinct_keys = int(lookup["__lookup_key"].nunique())
        lookup_keys = set(lookup["__lookup_key"].tolist())
        matched_mask = canonical_base.notna() & canonical_base.isin(lookup_keys)
        matched_base_rows = int(matched_mask.sum())
        unmatched_base_rows = int((canonical_base.notna() & ~matched_mask).sum())

        if options.duplicate_strategy == "first":
            lookup = lookup.drop_duplicates("__lookup_key", keep="first")

        rename_map: dict[str, str] = {}
        for column in spec.return_columns:
            candidate = f"{prefix}.{column}"
            output_name = _unique_column_name(candidate, [*result.columns, *rename_map.values()])
            rename_map[column] = output_name

        lookup = lookup[["__lookup_key", *spec.return_columns]].rename(columns=rename_map)
        result = result.merge(lookup, on="__lookup_key", how="left", sort=False, copy=False)

        status_base = f"{prefix}.Match status"
        status_column = _unique_column_name(status_base, result.columns)
        result[status_column] = "Not matched"
        result.loc[result["__lookup_key"].isna(), status_column] = "Missing key"
        matched_output = result["__lookup_key"].isin(lookup_keys)
        result.loc[matched_output, status_column] = "Matched"

        audits.append(
            LookupAudit(
                source=spec.file_name,
                sheet=spec.sheet_name,
                lookup_key=spec.key_column,
                lookup_rows=len(spec.dataframe),
                distinct_keys=distinct_keys,
                duplicate_key_rows=duplicate_key_rows,
                matched_base_rows=matched_base_rows,
                unmatched_base_rows=unmatched_base_rows,
                missing_base_keys=missing_base_keys,
                rows_after_lookup=len(result),
                returned_columns=len(spec.return_columns),
            )
        )

    result = result.sort_values("__lookup_base_row", kind="stable")
    result = result.drop(columns=["__lookup_base_row", "__lookup_key"]).reset_index(drop=True)
    return MergeResult(result, audits=audits, base_rows=base_rows, base_key=base_key)
