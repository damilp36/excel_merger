from streamlit.testing.v1 import AppTest

import pandas as pd

from excel_merger.merge import MergeResult


def test_initial_app_renders_and_file_count_controls_uploaders() -> None:
    app = AppTest.from_file("app.py", default_timeout=20).run()

    assert not app.exception
    assert len(app.number_input) == 1
    assert len(app.get("file_uploader")) == 3
    assert (
        app.get("file_uploader")[0].label
        == "Upload an existing result workbook"
    )
    assert app.info[0].value == "Upload at least two Excel files to begin."

    app.number_input[0].set_value(4).run()

    assert not app.exception
    assert len(app.get("file_uploader")) == 5


def test_clean_quality_success_alert_uses_a_valid_emoji() -> None:
    app = AppTest.from_string(
        """
import pandas as pd
from app import render_quality_summary
from excel_merger.excel_io import LoadedSheet
from excel_merger.quality import analyze_quality

loaded = LoadedSheet(
    file_name="clean.xlsx",
    sheet_name="Data",
    header_row=1,
    dataframe=pd.DataFrame({"ID": [1, 2], "Name": ["Ada", "Ben"]}),
    header_changes=[],
)
render_quality_summary(analyze_quality(loaded))
"""
    ).run()

    assert not app.exception
    assert app.success[0].value == "No obvious structural or cell-quality issues detected."


def test_final_result_page_renders_from_generated_merge() -> None:
    merged = pd.DataFrame(
        {
            "Student ID": ["S-1", "S-2"],
            "Academic Level": [300, 300],
            "CA": [20, 18],
            "Exam": [45, 36],
            "Exam.Match status": ["Matched", "Matched"],
        }
    )
    app = AppTest.from_file("app.py", default_timeout=20)
    app.session_state["active_page"] = "final_result"
    app.session_state["generated_lookup"] = {
        "fingerprint": "test-fingerprint",
        "result": MergeResult(merged, base_rows=2, base_key="Student ID"),
        "workbook": b"test",
    }
    app.run()

    assert not app.exception
    assert "Choose the identifier and score fields" in [item.value for item in app.subheader]
    assert len(app.selectbox) == 5
    assert app.selectbox[0].value == "Student ID"
    assert app.selectbox[1].value == "CA"
    assert app.selectbox[2].value == "Exam"
    assert app.selectbox[3].label == "Semester"
    assert app.selectbox[3].options == ["Harmattan", "Rain"]
    assert app.selectbox[4].label == "Course status"
    assert app.selectbox[4].options == ["Core", "Elective", "Required"]

    next(
        button
        for button in app.button
        if button.label == "Create final result workbook"
    ).click().run()

    assert not app.exception
    apply_button = next(
        button for button in app.button if button.label == "Apply score edits"
    )
    assert apply_button.disabled
    assert len(app.get("download_button")) == 1


def test_existing_final_result_can_be_opened_without_a_lookup() -> None:
    app = AppTest.from_string(
        """
import pandas as pd
from app import _open_imported_final_result, render_final_result_page
from excel_merger.final_result import (
    FinalResultDetails,
    GradeScale,
    build_final_result_workbook,
    import_final_result_workbook,
    prepare_final_result,
)

details = FinalResultDetails(
    programme="Example Programme",
    session="2025/2026",
    semester="Harmattan",
    course_code="ABC 101",
    course_title="Example Course",
    credit_units=3,
    course_status="Core",
    lecturers="Example Lecturer",
    remarks="",
    identifier_heading="Student ID",
    maximum_ca_score=30,
    maximum_exam_score=70,
    grade_scale=GradeScale(),
)
source = pd.DataFrame({
    "Student ID": ["S-1"],
    "CA": [20],
    "Exam": [50],
})
processed = prepare_final_result(source, "Student ID", "CA", "Exam", details)
content = build_final_result_workbook(processed, details)
imported = import_final_result_workbook(content)
_open_imported_final_result(imported, content, "previous_result.xlsx")
render_final_result_page()
"""
    ).run()

    assert not app.exception
    assert app.session_state["active_page"] == "final_result"
    assert isinstance(app.session_state["generated_lookup"]["result"], MergeResult)
    assert app.session_state["generated_lookup"]["imported_final_result"] is True
    assert app.selectbox[0].value == "Student ID"
    assert app.selectbox[1].value == "CA Score"
    assert app.selectbox[2].value == "Exam Score"
    assert next(
        widget for widget in app.number_input if widget.label == "Credit units"
    ).value == 3
    assert not any(
        "created with a default value" in warning.value
        for warning in app.warning
    )


def test_ordinary_result_workbook_can_be_opened_for_column_mapping() -> None:
    app = AppTest.from_string(
        """
import pandas as pd
from app import _open_existing_source_for_mapping, render_final_result_page
from excel_merger.excel_io import LoadedSheet

loaded = LoadedSheet(
    file_name="ordinary_result.xlsx",
    sheet_name="Results",
    header_row=2,
    dataframe=pd.DataFrame({
        "Student ID": ["S-1", "S-2"],
        "CA": [20, 18],
        "Exam": [50, 42],
    }),
    header_changes=[],
)
_open_existing_source_for_mapping(loaded, b"ordinary workbook")
render_final_result_page()
"""
    ).run()

    assert not app.exception
    assert app.session_state["active_page"] == "final_result"
    assert app.session_state["generated_lookup"]["imported_final_result"] is False
    assert app.session_state["generated_lookup"]["source_sheet"] == "Results"
    assert app.selectbox[0].value == "Student ID"
    assert app.selectbox[1].value == "CA"
    assert app.selectbox[2].value == "Exam"
    assert next(
        widget
        for widget in app.text_input
        if widget.label == "Final-result filename"
    ).value == "updated_ordinary_result.xlsx"
