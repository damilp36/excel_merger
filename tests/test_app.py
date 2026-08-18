from streamlit.testing.v1 import AppTest


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
