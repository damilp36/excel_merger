from io import BytesIO

import pandas as pd
from openpyxl import load_workbook

from excel_merger.final_result import (
    FinalResultDetails,
    GradeScale,
    build_final_result_workbook,
    inspect_final_result_workbook,
    prepare_final_result,
)


def make_details() -> FinalResultDetails:
    return FinalResultDetails(
        programme="Example Programme",
        session="2025/2026",
        semester="First",
        course_code="ABC 101",
        course_title="Example Course",
        credit_units=3,
        course_status="Core",
        lecturers="Example Lecturer",
        remarks="Review incomplete records.",
        identifier_heading="Student ID",
        maximum_ca_score=30,
        maximum_exam_score=70,
        grade_scale=GradeScale(),
    )


def test_prepares_grades_statuses_and_incomplete_records() -> None:
    source = pd.DataFrame(
        {
            "Identifier": ["S-1", "S-2", "S-3", "S-4", "S-5"],
            "Assessment": [25, 18, 14, "not a score", 31],
            "Examination": [52, 34, 20, 30, 40],
        }
    )

    processed = prepare_final_result(
        source,
        "Identifier",
        "Assessment",
        "Examination",
        make_details(),
    )

    assert processed.dataframe["Grand Total"].tolist()[:3] == [77.0, 52.0, 34.0]
    assert processed.dataframe["Grade Letter"].tolist()[:3] == ["A", "C", "F"]
    assert processed.dataframe["Student Status"].tolist() == [
        "Pass",
        "Pass",
        "Fail",
        "Incomplete",
        "Incomplete",
    ]
    assert processed.grade_counts == {"A": 1, "B": 0, "C": 1, "D": 0, "E": 0, "F": 1}
    assert processed.invalid_ca_count == 2
    assert processed.invalid_exam_count == 0


def test_rejects_invalid_grade_boundaries() -> None:
    details = make_details()
    invalid_scale = GradeScale(
        a_minimum=60,
        b_minimum=70,
        c_minimum=50,
        d_minimum=45,
        e_minimum=40,
    )
    invalid_details = FinalResultDetails(
        **{**details.__dict__, "grade_scale": invalid_scale}
    )

    assert "descend in order" in " ".join(invalid_details.validation_errors())


def test_builds_formula_driven_final_result_workbook() -> None:
    source = pd.DataFrame(
        {
            "Identifier": ["S-1", "S-2"],
            "Assessment": [25, 18],
            "Examination": [52, 34],
        }
    )
    details = make_details()
    processed = prepare_final_result(
        source,
        "Identifier",
        "Assessment",
        "Examination",
        details,
    )
    content = build_final_result_workbook(processed, details)
    inspection = inspect_final_result_workbook(content)

    assert inspection["sheet_names"] == ["Final Result", "Settings"]
    assert inspection["title"] == "MARK SHEET"
    assert inspection["headers"] == list(processed.dataframe.columns)
    assert inspection["first_total_formula"].startswith("=IF(OR(C40")
    assert "Settings!$B$20" in inspection["first_grade_formula"]
    assert inspection["chart_count"] == 1
    assert inspection["table_count"] == 1
    assert inspection["freeze_panes"] == "A40"

    workbook = load_workbook(BytesIO(content), data_only=False)
    try:
        assert workbook["Settings"]["B6"].value == "ABC 101"
        assert workbook["Settings"]["B14"].value == 30
        assert workbook["Final Result"]["B3"].value == "=Settings!B3"
    finally:
        workbook.close()
