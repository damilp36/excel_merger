"""Final-result calculations and workbook generation.

The layout is based on a conventional academic mark sheet: course details,
grade distribution, pass statistics, remarks, and a detailed result table.
Uploaded reference values are never copied into the generated workbook.
"""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
import math
from typing import Any

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

from .quality import is_missing


NAVY = "17324D"
BLUE = "2563EB"
LIGHT_BLUE = "DBEAFE"
PALE_GREEN = "DCFCE7"
PALE_RED = "FEE2E2"
PALE_AMBER = "FEF3C7"
PALE_GRAY = "F8FAFC"
WHITE = "FFFFFF"
TEXT = "172033"
MUTED = "64748B"
BORDER_COLOR = "CBD5E1"


@dataclass(frozen=True)
class GradeScale:
    a_minimum: float = 70
    b_minimum: float = 60
    c_minimum: float = 50
    d_minimum: float = 45
    e_minimum: float = 40
    pass_mark: float = 40

    def validation_errors(self, maximum_total: float) -> list[str]:
        errors: list[str] = []
        thresholds = [
            self.a_minimum,
            self.b_minimum,
            self.c_minimum,
            self.d_minimum,
            self.e_minimum,
        ]
        if not all(
            thresholds[index] > thresholds[index + 1]
            for index in range(len(thresholds) - 1)
        ):
            errors.append("Grade minimums must descend in order from A through E.")
        if any(value < 0 or value > maximum_total for value in thresholds):
            errors.append(
                f"Every grade minimum must be between 0 and {maximum_total:g}."
            )
        if self.pass_mark < 0 or self.pass_mark > maximum_total:
            errors.append(f"The pass mark must be between 0 and {maximum_total:g}.")
        return errors

    def grade_for(self, total: float) -> tuple[str, int]:
        if total >= self.a_minimum:
            return "A", 5
        if total >= self.b_minimum:
            return "B", 4
        if total >= self.c_minimum:
            return "C", 3
        if total >= self.d_minimum:
            return "D", 2
        if total >= self.e_minimum:
            return "E", 1
        return "F", 0


@dataclass(frozen=True)
class FinalResultDetails:
    programme: str
    session: str
    semester: str
    course_code: str
    course_title: str
    credit_units: int
    course_status: str
    lecturers: str
    remarks: str
    identifier_heading: str
    maximum_ca_score: float
    maximum_exam_score: float
    grade_scale: GradeScale

    @property
    def maximum_total(self) -> float:
        return self.maximum_ca_score + self.maximum_exam_score

    def validation_errors(self) -> list[str]:
        errors: list[str] = []
        if self.maximum_ca_score <= 0:
            errors.append("Maximum CA score must be greater than zero.")
        if self.maximum_exam_score <= 0:
            errors.append("Maximum exam score must be greater than zero.")
        if self.maximum_total > 1_000:
            errors.append("The combined maximum score must not exceed 1,000.")
        if not self.identifier_heading.strip():
            errors.append("The identifier heading cannot be blank.")
        if self.maximum_total > 0:
            errors.extend(self.grade_scale.validation_errors(self.maximum_total))
        return errors


@dataclass
class ProcessedFinalResult:
    dataframe: pd.DataFrame
    grade_counts: dict[str, int]
    pass_count: int
    fail_count: int
    incomplete_count: int
    invalid_ca_count: int
    invalid_exam_count: int

    @property
    def complete_count(self) -> int:
        return self.pass_count + self.fail_count

    @property
    def total_count(self) -> int:
        return len(self.dataframe)

    @property
    def pass_rate(self) -> float:
        return self.pass_count / self.complete_count if self.complete_count else 0.0


def _clean_identifier(value: Any) -> str | None:
    if is_missing(value):
        return None
    text = str(value).strip()
    return text or None


def _numeric_scores(series: pd.Series, maximum: float) -> tuple[pd.Series, pd.Series]:
    numeric = pd.to_numeric(series, errors="coerce")
    source_has_value = series.map(lambda value: not is_missing(value))
    invalid = source_has_value & (numeric.isna() | (numeric < 0) | (numeric > maximum))
    clean = numeric.mask(invalid)
    return clean.astype("Float64"), invalid


def prepare_final_result(
    merged_dataframe: pd.DataFrame,
    identifier_column: str,
    ca_column: str,
    exam_column: str,
    details: FinalResultDetails,
) -> ProcessedFinalResult:
    """Validate scores and calculate a preview of the final result."""

    errors = details.validation_errors()
    if errors:
        raise ValueError(" ".join(errors))
    selected = [identifier_column, ca_column, exam_column]
    missing_columns = [column for column in selected if column not in merged_dataframe.columns]
    if missing_columns:
        raise ValueError("Selected columns do not exist: " + ", ".join(missing_columns))
    if len(set(selected)) != 3:
        raise ValueError("Identifier, CA score, and exam score must use different columns.")

    ca_scores, invalid_ca = _numeric_scores(
        merged_dataframe[ca_column], details.maximum_ca_score
    )
    exam_scores, invalid_exam = _numeric_scores(
        merged_dataframe[exam_column], details.maximum_exam_score
    )
    identifiers = merged_dataframe[identifier_column].map(_clean_identifier)
    complete = ca_scores.notna() & exam_scores.notna()
    totals = (ca_scores + exam_scores).where(complete)

    grades: list[str | None] = []
    grade_points: list[int | None] = []
    statuses: list[str] = []
    for is_complete, total in zip(complete.tolist(), totals.tolist(), strict=True):
        if not is_complete or is_missing(total):
            grades.append(None)
            grade_points.append(None)
            statuses.append("Incomplete")
            continue
        grade, point = details.grade_scale.grade_for(float(total))
        grades.append(grade)
        grade_points.append(point)
        statuses.append("Pass" if float(total) >= details.grade_scale.pass_mark else "Fail")

    result = pd.DataFrame(
        {
            "S/No.": range(1, len(merged_dataframe) + 1),
            details.identifier_heading.strip(): identifiers,
            "CA Score": ca_scores,
            "Exam Score": exam_scores,
            "Grand Total": totals,
            "Grade Letter": grades,
            "Grade Point": grade_points,
            "Student Status": statuses,
        }
    )
    grade_counts = {
        grade: int(result["Grade Letter"].eq(grade).sum()) for grade in "ABCDEF"
    }
    return ProcessedFinalResult(
        dataframe=result,
        grade_counts=grade_counts,
        pass_count=statuses.count("Pass"),
        fail_count=statuses.count("Fail"),
        incomplete_count=statuses.count("Incomplete"),
        invalid_ca_count=int(invalid_ca.sum()),
        invalid_exam_count=int(invalid_exam.sum()),
    )


def _safe_excel_value(value: Any) -> Any:
    if is_missing(value):
        return None
    if hasattr(value, "item") and not isinstance(value, (str, bytes)):
        try:
            value = value.item()
        except (TypeError, ValueError, AttributeError):
            pass
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return None
    return value


def _set_literal(cell: Any, value: Any) -> None:
    safe = _safe_excel_value(value)
    cell.value = safe
    if isinstance(safe, str) and safe.startswith("="):
        cell.data_type = "s"


def _fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)


def _apply_section_header(worksheet: Any, cell_range: str, title: str) -> None:
    worksheet.merge_cells(cell_range)
    cell = worksheet[cell_range.split(":", 1)[0]]
    cell.value = title
    cell.fill = _fill(NAVY)
    cell.font = Font(name="Aptos Display", size=12, bold=True, color=WHITE)
    cell.alignment = Alignment(horizontal="left", vertical="center")


def _build_settings_sheet(
    workbook: Workbook,
    details: FinalResultDetails,
) -> Any:
    worksheet = workbook.create_sheet("Settings")
    worksheet.sheet_view.showGridLines = False
    worksheet.merge_cells("A1:D1")
    worksheet["A1"] = "Final Result Settings"
    worksheet["A1"].fill = _fill(NAVY)
    worksheet["A1"].font = Font(name="Aptos Display", size=16, bold=True, color=WHITE)
    worksheet["A1"].alignment = Alignment(vertical="center")
    worksheet.row_dimensions[1].height = 30

    details_rows = [
        ("Programme", details.programme),
        ("Session", details.session),
        ("Semester", details.semester),
        ("Course Code", details.course_code),
        ("Course Title", details.course_title),
        ("Credit Units", details.credit_units),
        ("Course Status", details.course_status),
        ("Lecturer(s)", details.lecturers),
        ("Remarks", details.remarks),
    ]
    for row_number, (label, value) in enumerate(details_rows, start=3):
        worksheet.cell(row_number, 1, label)
        _set_literal(worksheet.cell(row_number, 2), value)

    worksheet.merge_cells("A13:D13")
    worksheet["A13"] = "Score Settings"
    worksheet["A13"].fill = _fill(BLUE)
    worksheet["A13"].font = Font(name="Aptos", bold=True, color=WHITE)
    score_rows = [
        ("Maximum CA Score", details.maximum_ca_score),
        ("Maximum Exam Score", details.maximum_exam_score),
        ("Pass Mark", details.grade_scale.pass_mark),
    ]
    for row_number, (label, value) in enumerate(score_rows, start=14):
        worksheet.cell(row_number, 1, label)
        worksheet.cell(row_number, 2, value)

    worksheet.merge_cells("A18:D18")
    worksheet["A18"] = "Grading Scale"
    worksheet["A18"].fill = _fill(BLUE)
    worksheet["A18"].font = Font(name="Aptos", bold=True, color=WHITE)
    worksheet.append([])
    for column, value in enumerate(["Grade", "Minimum Total", "Grade Point"], start=1):
        cell = worksheet.cell(19, column, value)
        cell.fill = _fill(LIGHT_BLUE)
        cell.font = Font(name="Aptos", bold=True, color=TEXT)
    grade_rows = [
        ("A", details.grade_scale.a_minimum, 5),
        ("B", details.grade_scale.b_minimum, 4),
        ("C", details.grade_scale.c_minimum, 3),
        ("D", details.grade_scale.d_minimum, 2),
        ("E", details.grade_scale.e_minimum, 1),
        ("F", 0, 0),
    ]
    for row_number, values in enumerate(grade_rows, start=20):
        for column_number, value in enumerate(values, start=1):
            worksheet.cell(row_number, column_number, value)

    thin = Side(style="thin", color=BORDER_COLOR)
    for row in worksheet.iter_rows(min_row=3, max_row=25, min_col=1, max_col=3):
        for cell in row:
            if cell.row not in {13, 18}:
                cell.font = Font(
                    name="Aptos",
                    size=10,
                    bold=cell.row == 19,
                    color=TEXT,
                )
            cell.border = Border(bottom=thin)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for row_number in list(range(3, 12)) + list(range(14, 17)):
        worksheet.cell(row_number, 1).font = Font(name="Aptos", bold=True, color=MUTED)
    worksheet.column_dimensions["A"].width = 23
    worksheet.column_dimensions["B"].width = 42
    worksheet.column_dimensions["C"].width = 16
    worksheet.column_dimensions["D"].width = 4
    worksheet.freeze_panes = "A3"
    return worksheet


def _build_final_sheet(
    workbook: Workbook,
    processed: ProcessedFinalResult,
    details: FinalResultDetails,
) -> Any:
    worksheet = workbook.create_sheet("Final Result", 0)
    worksheet.sheet_view.showGridLines = False
    worksheet.freeze_panes = "A40"

    worksheet.merge_cells("A1:H1")
    worksheet["A1"] = "MARK SHEET"
    worksheet["A1"].fill = _fill(NAVY)
    worksheet["A1"].font = Font(name="Aptos Display", size=18, bold=True, color=WHITE)
    worksheet["A1"].alignment = Alignment(horizontal="center", vertical="center")
    worksheet.row_dimensions[1].height = 34

    detail_links = [
        ("Programme", "=Settings!B3"),
        ("Session", "=Settings!B4"),
        ("Semester", "=Settings!B5"),
        ("Course Code", "=Settings!B6"),
        ("Course Title", "=Settings!B7"),
        ("Credit Unit(s)", "=Settings!B8"),
        ("Course Status", "=Settings!B9"),
        ("Lecturer(s)", "=Settings!B10"),
    ]
    for row_number, (label, formula) in enumerate(detail_links, start=3):
        worksheet.cell(row_number, 1, label)
        worksheet.merge_cells(start_row=row_number, start_column=2, end_row=row_number, end_column=8)
        worksheet.cell(row_number, 2, formula)
        worksheet.cell(row_number, 1).font = Font(name="Aptos", bold=True, color=MUTED)
        worksheet.cell(row_number, 2).font = Font(name="Aptos", color=TEXT)
        worksheet.cell(row_number, 2).alignment = Alignment(wrap_text=True)

    _apply_section_header(worksheet, "A12:H12", "Executive Summary")
    summary_headers = ["Grades", "A", "B", "C", "D", "E", "F", "Total Records"]
    for column_number, value in enumerate(summary_headers, start=1):
        cell = worksheet.cell(13, column_number, value)
        cell.fill = _fill(LIGHT_BLUE)
        cell.font = Font(name="Aptos", bold=True, color=TEXT)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    data_start = 40
    data_end = data_start + len(processed.dataframe) - 1
    data_range_end = max(data_end, data_start)
    worksheet["A14"] = "Count"
    worksheet["A15"] = "Percentage"
    for offset, grade in enumerate("ABCDEF", start=2):
        column_letter = chr(64 + offset)
        worksheet.cell(14, offset, f'=COUNTIF($F${data_start}:$F${data_range_end},"{grade}")')
        worksheet.cell(15, offset, f'=IF($H$14=0,0,{column_letter}14/$H$14)')
        worksheet.cell(15, offset).number_format = "0.0%"
    worksheet["H14"] = f"=COUNTA($B${data_start}:$B${data_range_end})"
    worksheet["H15"] = "=IF(H14=0,0,SUM(B14:G14)/H14)"
    worksheet["H15"].number_format = "0.0%"

    stats = [
        ("A17", "Number of passes", "B17", f'=COUNTIF($H${data_start}:$H${data_range_end},"Pass")'),
        ("C17", "Number of fails", "D17", f'=COUNTIF($H${data_start}:$H${data_range_end},"Fail")'),
        ("E17", "Incomplete records", "F17", f'=COUNTIF($H${data_start}:$H${data_range_end},"Incomplete")'),
        ("G17", "Pass rate", "H17", "=IF((B17+D17)=0,0,B17/(B17+D17))"),
    ]
    for label_cell, label, value_cell, formula in stats:
        worksheet[label_cell] = label
        worksheet[label_cell].font = Font(name="Aptos", bold=True, color=MUTED)
        worksheet[value_cell] = formula
    worksheet["H17"].number_format = "0.0%"

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Grade Distribution"
    chart.y_axis.title = "Number of records"
    chart.x_axis.title = "Grade"
    chart.height = 6.5
    chart.width = 14.5
    chart.legend = None
    chart.add_data(Reference(worksheet, min_col=2, max_col=7, min_row=14), from_rows=True)
    chart.set_categories(Reference(worksheet, min_col=2, max_col=7, min_row=13))
    worksheet.add_chart(chart, "B19")

    _apply_section_header(worksheet, "A33:H33", "Lecturer's Remarks")
    worksheet.merge_cells("A34:H34")
    worksheet["A34"] = "=Settings!B11"
    worksheet["A34"].alignment = Alignment(vertical="top", wrap_text=True)
    worksheet.row_dimensions[34].height = 38

    _apply_section_header(worksheet, "A36:H36", "Main Result")
    worksheet["A37"] = "Maximum CA score"
    worksheet["B37"] = "=Settings!B14"
    worksheet["D37"] = "Maximum exam score"
    worksheet["E37"] = "=Settings!B15"
    worksheet["G37"] = "Maximum total"
    worksheet["H37"] = "=Settings!B14+Settings!B15"
    for cell_ref in ["A37", "D37", "G37"]:
        worksheet[cell_ref].font = Font(name="Aptos", bold=True, color=MUTED)

    headers = list(processed.dataframe.columns)
    for column_number, value in enumerate(headers, start=1):
        cell = worksheet.cell(39, column_number, value)
        cell.fill = _fill(BLUE)
        cell.font = Font(name="Aptos", size=10, bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    worksheet.row_dimensions[39].height = 42

    thin = Side(style="thin", color=BORDER_COLOR)
    for zero_index, row in processed.dataframe.iterrows():
        row_number = data_start + zero_index
        worksheet.cell(row_number, 1, int(row["S/No."]))
        _set_literal(worksheet.cell(row_number, 2), row.iloc[1])
        worksheet.cell(row_number, 3, _safe_excel_value(row["CA Score"]))
        worksheet.cell(row_number, 4, _safe_excel_value(row["Exam Score"]))
        worksheet.cell(
            row_number,
            5,
            f'=IF(OR(C{row_number}="",D{row_number}=""),"",C{row_number}+D{row_number})',
        )
        worksheet.cell(
            row_number,
            6,
            (
                f'=IF(E{row_number}="","",IF(E{row_number}>=Settings!$B$20,"A",'
                f'IF(E{row_number}>=Settings!$B$21,"B",IF(E{row_number}>=Settings!$B$22,"C",'
                f'IF(E{row_number}>=Settings!$B$23,"D",IF(E{row_number}>=Settings!$B$24,"E","F"))))))'
            ),
        )
        worksheet.cell(
            row_number,
            7,
            (
                f'=IF(E{row_number}="","",IF(E{row_number}>=Settings!$B$20,5,'
                f'IF(E{row_number}>=Settings!$B$21,4,IF(E{row_number}>=Settings!$B$22,3,'
                f'IF(E{row_number}>=Settings!$B$23,2,IF(E{row_number}>=Settings!$B$24,1,0))))))'
            ),
        )
        worksheet.cell(
            row_number,
            8,
            (
                f'=IF(E{row_number}="","Incomplete",'
                f'IF(E{row_number}>=Settings!$B$16,"Pass","Fail"))'
            ),
        )
        for column_number in range(1, 9):
            cell = worksheet.cell(row_number, column_number)
            cell.border = Border(bottom=thin)
            cell.font = Font(name="Aptos", size=10, color=TEXT)
            cell.alignment = Alignment(
                horizontal="left" if column_number == 2 else "center",
                vertical="center",
            )
        if row_number % 2 == 0:
            for column_number in range(1, 9):
                worksheet.cell(row_number, column_number).fill = _fill(PALE_GRAY)

    if len(processed.dataframe):
        table = Table(displayName="FinalResultTable", ref=f"A39:H{data_end}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        worksheet.add_table(table)
        status_range = f"H{data_start}:H{data_end}"
        worksheet.conditional_formatting.add(
            status_range,
            FormulaRule(
                formula=[f'$H{data_start}="Pass"'],
                fill=_fill(PALE_GREEN),
                font=Font(color="166534"),
            ),
        )
        worksheet.conditional_formatting.add(
            status_range,
            FormulaRule(
                formula=[f'$H{data_start}="Fail"'],
                fill=_fill(PALE_RED),
                font=Font(color="991B1B"),
            ),
        )
        worksheet.conditional_formatting.add(
            status_range,
            FormulaRule(
                formula=[f'$H{data_start}="Incomplete"'],
                fill=_fill(PALE_AMBER),
                font=Font(color="92400E"),
            ),
        )

    widths = [9, 24, 15, 15, 16, 15, 14, 18]
    for column_number, width in enumerate(widths, start=1):
        worksheet.column_dimensions[chr(64 + column_number)].width = width
    worksheet.auto_filter.ref = f"A39:H{data_range_end}"
    worksheet.print_area = f"A1:H{data_range_end}"
    worksheet.print_title_rows = "39:39"
    worksheet.page_setup.orientation = "landscape"
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.page_margins.left = 0.25
    worksheet.page_margins.right = 0.25
    worksheet.page_margins.top = 0.5
    worksheet.page_margins.bottom = 0.5
    return worksheet


def build_final_result_workbook(
    processed: ProcessedFinalResult,
    details: FinalResultDetails,
) -> bytes:
    """Create an editable final-result workbook with live Excel formulas."""

    workbook = Workbook()
    workbook.remove(workbook.active)
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    _build_settings_sheet(workbook, details)
    _build_final_sheet(workbook, processed, details)
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def inspect_final_result_workbook(content: bytes) -> dict[str, Any]:
    """Return compact structural details used by tests and local verification."""

    workbook = load_workbook(BytesIO(content), data_only=False)
    try:
        result = workbook["Final Result"]
        settings = workbook["Settings"]
        return {
            "sheet_names": workbook.sheetnames,
            "title": result["A1"].value,
            "headers": [result.cell(39, column).value for column in range(1, 9)],
            "first_total_formula": result["E40"].value,
            "first_grade_formula": result["F40"].value,
            "settings_title": settings["A1"].value,
            "chart_count": len(result._charts),
            "table_count": len(result.tables),
            "freeze_panes": result.freeze_panes,
        }
    finally:
        workbook.close()
