from streamlit.testing.v1 import AppTest

import pandas as pd

from excel_merger.merge import MergeResult


def test_initial_app_renders_and_file_count_controls_uploaders() -> None:
    app = AppTest.from_file("app.py", default_timeout=20).run()

    assert not app.exception
    assert len(app.number_input) == 1
    assert len(app.get("file_uploader")) == 2
    assert app.info[0].value == "Upload at least two Excel files to begin."

    app.number_input[0].set_value(4).run()

    assert not app.exception
    assert len(app.get("file_uploader")) == 4


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
