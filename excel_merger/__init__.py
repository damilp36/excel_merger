"""Core services for the Lookup Studio Streamlit application."""

from .excel_io import LoadedSheet, WorkbookReadError, list_sheet_names, stream_sheet
from .exporter import ExportLimitError, build_output_workbook, clean_output_filename
from .final_result import (
    FinalResultDetails,
    GradeScale,
    ProcessedFinalResult,
    build_final_result_workbook,
    prepare_final_result,
)
from .merge import LookupSpec, MatchOptions, MergeResult, merge_lookups
from .quality import QualityReport, analyze_quality

__all__ = [
    "ExportLimitError",
    "FinalResultDetails",
    "GradeScale",
    "LoadedSheet",
    "LookupSpec",
    "MatchOptions",
    "MergeResult",
    "ProcessedFinalResult",
    "QualityReport",
    "WorkbookReadError",
    "analyze_quality",
    "build_output_workbook",
    "build_final_result_workbook",
    "clean_output_filename",
    "list_sheet_names",
    "merge_lookups",
    "prepare_final_result",
    "stream_sheet",
]
