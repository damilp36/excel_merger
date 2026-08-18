"""Data-quality profiling and preview highlighting."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, time
from numbers import Number
from typing import Any

import pandas as pd

from .excel_io import LoadedSheet


MISSING_STYLE = "background-color: #fee2e2; color: #991b1b"
INCONSISTENT_STYLE = "background-color: #fef3c7; color: #92400e"
DUPLICATE_KEY_STYLE = "box-shadow: inset 3px 0 #8b5cf6"


@dataclass(frozen=True)
class ColumnQuality:
    column: str
    inferred_type: str
    non_empty: int
    missing: int
    missing_percent: float
    distinct: int
    inconsistent: int


@dataclass
class QualityReport:
    file_name: str
    sheet_name: str
    row_count: int
    column_count: int
    missing_cells: int
    duplicate_rows: int
    blank_rows_removed: int
    blank_columns_removed: int
    header_changes: list[str] = field(default_factory=list)
    columns: list[ColumnQuality] = field(default_factory=list)
    inconsistent_cells: set[tuple[int, str]] = field(default_factory=set)

    @property
    def mixed_type_columns(self) -> int:
        return sum(column.inconsistent > 0 for column in self.columns)

    @property
    def warnings(self) -> list[str]:
        messages: list[str] = []
        if self.missing_cells:
            messages.append(f"{self.missing_cells:,} missing cells")
        if self.mixed_type_columns:
            messages.append(f"{self.mixed_type_columns:,} mixed-type columns")
        if self.duplicate_rows:
            messages.append(f"{self.duplicate_rows:,} duplicate rows")
        if self.header_changes:
            messages.append(f"{len(self.header_changes):,} header fixes")
        if self.blank_rows_removed:
            messages.append(f"{self.blank_rows_removed:,} trailing blank rows removed")
        if self.blank_columns_removed:
            messages.append(f"{self.blank_columns_removed:,} blank columns removed")
        return messages

    def columns_frame(self) -> pd.DataFrame:
        return pd.DataFrame(
            [
                {
                    "Column": column.column,
                    "Likely type": column.inferred_type,
                    "Filled": column.non_empty,
                    "Missing": column.missing,
                    "Missing %": column.missing_percent,
                    "Unique values": column.distinct,
                    "Inconsistent cells": column.inconsistent,
                }
                for column in self.columns
            ]
        )


def is_missing(value: Any) -> bool:
    if value is None:
        return True
    try:
        result = pd.isna(value)
    except (TypeError, ValueError):
        return False
    return bool(result) if not hasattr(result, "__len__") else False


def infer_value_type(value: Any) -> str:
    if is_missing(value):
        return "missing"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (datetime, date, time, pd.Timestamp)):
        return "date/time"
    if isinstance(value, Number):
        return "number"
    if isinstance(value, str):
        return "text"
    return "other"


def analyze_quality(loaded: LoadedSheet) -> QualityReport:
    dataframe = loaded.dataframe
    columns: list[ColumnQuality] = []
    inconsistent_cells: set[tuple[int, str]] = set()

    for column_name in dataframe.columns:
        series = dataframe[column_name]
        types = series.map(infer_value_type)
        populated_types = types[types != "missing"]
        if populated_types.empty:
            dominant = "empty"
            inconsistent_mask = pd.Series(False, index=series.index)
        else:
            counts = populated_types.value_counts()
            dominant = str(counts.index[0])
            inconsistent_mask = (types != "missing") & (types != dominant)
            for row_index in series.index[inconsistent_mask]:
                inconsistent_cells.add((int(row_index), str(column_name)))

        missing = int(series.isna().sum())
        columns.append(
            ColumnQuality(
                column=str(column_name),
                inferred_type=dominant,
                non_empty=int(series.notna().sum()),
                missing=missing,
                missing_percent=(missing / len(series)) if len(series) else 0.0,
                distinct=int(series.nunique(dropna=True)),
                inconsistent=int(inconsistent_mask.sum()),
            )
        )

    header_changes = [
        f"Column {change.position}: {change.original} → {change.replacement} ({change.reason})"
        for change in loaded.header_changes
    ]
    return QualityReport(
        file_name=loaded.file_name,
        sheet_name=loaded.sheet_name,
        row_count=len(dataframe),
        column_count=len(dataframe.columns),
        missing_cells=int(dataframe.isna().sum().sum()),
        duplicate_rows=int(dataframe.duplicated().sum()),
        blank_rows_removed=loaded.blank_rows_removed,
        blank_columns_removed=loaded.blank_columns_removed,
        header_changes=header_changes,
        columns=columns,
        inconsistent_cells=inconsistent_cells,
    )


def style_preview(
    dataframe: pd.DataFrame,
    report: QualityReport,
    key_column: str | None = None,
    max_rows: int = 200,
) -> pd.io.formats.style.Styler:
    """Return a compact Styler with visible quality warnings."""

    preview = dataframe.head(max_rows)
    styles = pd.DataFrame("", index=preview.index, columns=preview.columns)

    for row_index in preview.index:
        for column in preview.columns:
            value = preview.at[row_index, column]
            if is_missing(value):
                styles.at[row_index, column] = MISSING_STYLE
            elif (int(row_index), str(column)) in report.inconsistent_cells:
                styles.at[row_index, column] = INCONSISTENT_STYLE

    if key_column and key_column in preview.columns:
        duplicates = preview[key_column].notna() & preview[key_column].duplicated(keep=False)
        for row_index in preview.index[duplicates]:
            existing = styles.at[row_index, key_column]
            styles.at[row_index, key_column] = f"{existing}; {DUPLICATE_KEY_STYLE}".strip("; ")

    return preview.style.apply(lambda _: styles, axis=None).format(na_rep="—")
